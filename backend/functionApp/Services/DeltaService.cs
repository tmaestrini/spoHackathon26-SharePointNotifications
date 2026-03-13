using Azure.Data.Tables;
using functionApp.Helpers;
using functionApp.Models;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using graph = Microsoft.Graph.Models;
using Microsoft.Graph.Sites.Item.Lists.Item.Items.DeltaWithToken;
using Microsoft.Kiota.Abstractions;
using Microsoft.SharePoint.Client;
using System.Linq;

namespace functionApp.Services;

public class DeltaService
{
    private readonly TableClient _tableClient;
    private readonly ILogger<DeltaService> _logger;
    private readonly HttpClient _httpClient;
    private readonly AppSettings _appSettings;

    public DeltaService(TableServiceClient tableServiceClient, AppSettings appSettings, ILogger<DeltaService> logger, HttpClient httpClient)
    {
        _logger = logger;
        _appSettings = appSettings;
        _tableClient = tableServiceClient.GetTableClient(appSettings.TableDeltas);
        _tableClient.CreateIfNotExists();
        _httpClient = httpClient;
        _logger.LogInformation("DeltaService initialized with table '{TableName}'.", appSettings.TableDeltas);
    }

    public async Task SaveAsync(DeltaEntity entity)
    {
        _logger.LogInformation("Saving delta entity {RowKey}.", entity.RowKey);
        await _tableClient.UpsertEntityAsync(entity, TableUpdateMode.Replace);
        _logger.LogInformation("Delta entity {RowKey} saved successfully.", entity.RowKey);
    }

    public async Task<DeltaEntity?> GetAsync(string subscriptionId)
    {
        _logger.LogInformation("Fetching delta entity {RowKey}.", subscriptionId);
        try
        {
            var response = await _tableClient.GetEntityAsync<DeltaEntity>("Deltas", subscriptionId);
            return response.Value;
        }
        catch (Azure.RequestFailedException ex) when (ex.Status == 404)
        {
            _logger.LogWarning("Delta entity {RowKey} not found.", subscriptionId);
            return null;
        }
    }

    public async Task<List<DeltaEntity>> GetByListAsync(string subscriptionId)
    {
        _logger.LogInformation("Fetching delta entities for subscription {subscriptionId}.", subscriptionId);
        var results = new List<DeltaEntity>();
        await foreach (var entity in _tableClient.QueryAsync<DeltaEntity>(
            e => e.PartitionKey == "Deltas" && e.RowKey == subscriptionId))
        {
            results.Add(entity);
        }
        _logger.LogInformation("Found {Count} delta entities for subscription {subscriptionId}.", results.Count, subscriptionId);
        return results;
    }

    public async Task<List<DeltaEntity>> GetExpiringAsync(DateTime before)
    {
        _logger.LogInformation("Fetching delta entities expiring before {before}.", before);
        var results = new List<DeltaEntity>();
        await foreach (var entity in _tableClient.QueryAsync<DeltaEntity>(
            e => e.PartitionKey == "Deltas" && e.Timestamp < before))
        {
            results.Add(entity);
        }
        _logger.LogInformation("Found {Count} expiring delta entities.", results.Count);
        return results;
    }

    public async Task<bool> DeleteAsync(string subscriptionId)
    {
        _logger.LogInformation("Deleting delta entity {subscriptionId}.", subscriptionId);
        try
        {
            await _tableClient.DeleteEntityAsync("Deltas", subscriptionId);
            _logger.LogInformation("Delta entity {subscriptionId} deleted successfully.", subscriptionId);
            return true;
        }
        catch (Azure.RequestFailedException ex) when (ex.Status == 404)
        {
            _logger.LogWarning("Delta entity {subscriptionId} not found for deletion.", subscriptionId);
            return false;
        }
    }

    /// <summary>
    /// Handles SharePoint library delta changes for the specified registration
    /// </summary>
    /// <param name="webhookNotification">The webhook notification that triggered this delta request</param>
    /// <returns>List of delta item changes classified by change type</returns>
    public async Task<List<DeltaItemChange>> GetDeltaForNotificationAsync(WebhookNotificationModel webhookNotification)
    {
        var deltaChanges = new List<DeltaItemChange>();
        var processedItems = new HashSet<string>();

        try
        {
            _logger.LogInformation($"Processing delta for webhook subscription: {webhookNotification.SubscriptionId}");

            var graphClient = ConnectionHelper.GraphClient(_appSettings, _logger);

            // Get the site information using the site URL from the webhook notification
            var site = await graphClient
                                .Sites[$"{_appSettings.SharePointTenantName}:{webhookNotification.SiteUrl}"]
                                .GetAsync();

            var listId = webhookNotification.Resource;
            var existingDelta = await GetAsync(webhookNotification.SubscriptionId);
            string? newDeltaLink = null;

            // If there is an existing delta token, use it to get changes since last checkpoint
            if (existingDelta != null && !string.IsNullOrEmpty(existingDelta.DeltaLink))
            {
                _logger.LogInformation($"Using existing delta link for subscription: {webhookNotification.SubscriptionId}");

                // Call stored deltaLink to get changes since checkpoint
                var request = new RequestInformation
                {
                    HttpMethod = Method.GET,
                    URI = new Uri(existingDelta.DeltaLink)
                };
                var deltaResponse = await graphClient.RequestAdapter.SendAsync<DeltaWithTokenGetResponse>(
                    request,
                    DeltaWithTokenGetResponse.CreateFromDiscriminatorValue
                );


                // Process all pages and collect changed items
                var pageIterator = PageIterator<graph.ListItem, Microsoft.Graph.Sites.Item.Lists.Item.Items.DeltaWithToken.DeltaWithTokenGetResponse>
                                    .CreatePageIterator(
                                        graphClient,
                                        deltaResponse,
                                        (item) =>
                                        {
                                            if (item?.Id != null && !processedItems.Contains(item.Id))
                                            {
                                                var changeType = ClassifyChangeType(item);
                                                deltaChanges.Add(new DeltaItemChange
                                                {
                                                    ItemId = item.Id,
                                                    ChangeType = changeType,
                                                    Item = changeType != DeltaChangeType.Deleted ? item : null
                                                });
                                                processedItems.Add(item.Id);
                                                _logger.LogDebug($"Item {item.Id} classified as {changeType}");
                                            }
                                            return true;
                                        });

                await pageIterator.IterateAsync();

                // Capture the new deltaLink from the final response
                newDeltaLink = ExtractDeltaLink(deltaResponse);
            }
            else
            {
                _logger.LogInformation($"No existing delta token found. Performing initial delta sync for subscription: {webhookNotification.SubscriptionId}");

                // Send initial delta request
                var deltaResponse = await graphClient
                                            .Sites[site.Id]
                                            .Lists[listId]
                                            .Items
                                            .Delta
                                            .GetAsDeltaGetResponseAsync((x) =>
                                            {
                                                x.QueryParameters.Top = 100;
                                            });

                // Follow pagination if present - process all pages until we get deltaLink
                var pageIterator = PageIterator<graph.ListItem, Microsoft.Graph.Sites.Item.Lists.Item.Items.Delta.DeltaGetResponse>
                                    .CreatePageIterator(
                                        graphClient,
                                        deltaResponse,
                                        (item) =>
                                        {
                                            if (item?.Id != null && !processedItems.Contains(item.Id))
                                            {
                                                var changeType = ClassifyChangeType(item);
                                                deltaChanges.Add(new DeltaItemChange
                                                {
                                                    ItemId = item.Id,
                                                    ChangeType = changeType,
                                                    Item = changeType != DeltaChangeType.Deleted ? item : null
                                                });
                                                processedItems.Add(item.Id);
                                                _logger.LogDebug($"Item {item.Id} classified as {changeType}");
                                            }
                                            return true;
                                        });

                await pageIterator.IterateAsync();

                // Capture the deltaLink from the final response
                newDeltaLink = ExtractDeltaLink(deltaResponse);
            }

            // Store the returned deltaLink as the checkpoint
            if (!string.IsNullOrEmpty(newDeltaLink))
            {
                await SaveDeltaCheckpointAsync(webhookNotification.SubscriptionId, newDeltaLink);
                _logger.LogInformation($"Updated delta checkpoint for subscription: {webhookNotification.SubscriptionId}");
            }
            else
            {
                _logger.LogWarning($"No delta link found in response for subscription: {webhookNotification.SubscriptionId}");
            }

            deltaChanges = await EnrichDeletedInformationAsync(deltaChanges, site.WebUrl);

            _logger.LogInformation($"Processed delta for webhook subscription: {webhookNotification.SubscriptionId}");
            _logger.LogInformation($"Found {deltaChanges.Count} delta changes: " +
                $"{deltaChanges.Count(c => c.ChangeType == DeltaChangeType.Created)} created, " +
                $"{deltaChanges.Count(c => c.ChangeType == DeltaChangeType.Updated)} updated, " +
                $"{deltaChanges.Count(c => c.ChangeType == DeltaChangeType.Deleted)} deleted");

            return deltaChanges;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, $"Error processing delta for subscription: {webhookNotification.SubscriptionId}");
            throw;
        }
    }

    private async Task<List<DeltaItemChange>> EnrichDeletedInformationAsync(List<DeltaItemChange> deltaChanges, string siteUrl)
    {
        if (!deltaChanges.Any(c => c.ChangeType == DeltaChangeType.Deleted))
        {
            return deltaChanges;
        }

        if (String.IsNullOrEmpty(siteUrl))
        {
            _logger.LogWarning("Site URL is null or empty. Cannot enrich deleted items with additional information.");
            return deltaChanges;
        }

        var deletedChanges = deltaChanges.Where(c => c.ChangeType == DeltaChangeType.Deleted).ToList();
        _logger.LogInformation("Enriching {count} deleted items with additional information from SharePoint.", deletedChanges.Count);

        try
        {
            // Try REST API approach first since you mentioned it works
            var enrichedItems = await TryEnrichWithRestApiAsync(siteUrl, deletedChanges);
            if (enrichedItems > 0)
            {
                _logger.LogInformation("Successfully enriched {count} deleted items using REST API.", enrichedItems);
                return deltaChanges;
            }

            // Fallback to CSOM approach with improved matching
            _logger.LogInformation("REST API approach didn't work, falling back to CSOM approach.");
            return await TryEnrichWithCsomAsync(siteUrl, deltaChanges, deletedChanges);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error enriching deleted items. Returning original changes without enrichment.");
            return deltaChanges;
        }
    }

    private async Task<int> TryEnrichWithRestApiAsync(string siteUrl, List<DeltaItemChange> deletedChanges)
    {
        try
        {
            var enrichedCount = 0;
            var accessToken = _appSettings.GetAppOnlyAccessToken(siteUrl, _logger);

            if (string.IsNullOrEmpty(accessToken))
            {
                _logger.LogWarning("Could not get access token for REST API call.");
                return 0;
            }

            // Use REST API to get recycle bin items
            var restUrl = $"{siteUrl.TrimEnd('/')}/_api/web/recyclebin?$top=1000";
            _logger.LogDebug("Calling REST API: {url}", restUrl);

            using var request = new HttpRequestMessage(HttpMethod.Get, restUrl);
            request.Headers.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", accessToken);
            request.Headers.Add("Accept", "application/json;odata=verbose");

            var response = await _httpClient.SendAsync(request);

            if (!response.IsSuccessStatusCode)
            {
                _logger.LogWarning("REST API call failed with status: {statusCode}", response.StatusCode);
                return 0;
            }

            var jsonContent = await response.Content.ReadAsStringAsync();
            _logger.LogDebug("REST API response received. Length: {length}", jsonContent.Length);

            // Parse the JSON response
            using var jsonDoc = System.Text.Json.JsonDocument.Parse(jsonContent);
            var results = jsonDoc.RootElement.GetProperty("d").GetProperty("results");

            _logger.LogInformation("Found {count} items in recycle bin via REST API.", results.GetArrayLength());

            foreach (var recycleBinItem in results.EnumerateArray())
            {
                var itemId = recycleBinItem.GetProperty("Id").GetInt32().ToString();
                var title = recycleBinItem.TryGetProperty("Title", out var titleProp) ? titleProp.GetString() : "";

                _logger.LogDebug("Recycle bin item - ID: {id}, Title: {title}", itemId, title);

                // Try different ID matching strategies
                var matchingChange = deletedChanges.FirstOrDefault(c => 
                    c.ItemId == itemId || 
                    c.ItemId.EndsWith($":{itemId}") ||
                    c.ItemId.Contains(itemId));

                if (matchingChange != null)
                {
                    _logger.LogDebug("Found match for item ID {itemId}", itemId);

                    if (matchingChange.Item == null)
                    {
                        matchingChange.Item = new Microsoft.Graph.Models.ListItem();
                    }

                    matchingChange.Item.Name = title;

                    if (matchingChange.Item.Deleted == null)
                    {
                        matchingChange.Item.Deleted = new Microsoft.Graph.Models.Deleted();
                    }

                    // Add deletion metadata
                    if (recycleBinItem.TryGetProperty("DeletedDate", out var deletedDate))
                    {
                        matchingChange.Item.Deleted.AdditionalData["DeletedDateTime"] = deletedDate.GetString();
                    }

                    if (recycleBinItem.TryGetProperty("DeletedByName", out var deletedBy))
                    {
                        matchingChange.Item.Deleted.AdditionalData["DeletedBy"] = deletedBy.GetString();
                    }

                    enrichedCount++;
                }
                else
                {
                    _logger.LogDebug("No matching change found for recycle bin item ID: {id}", itemId);
                }
            }

            return enrichedCount;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error using REST API to enrich deleted items.");
            return 0;
        }
    }

    private async Task<List<DeltaItemChange>> TryEnrichWithCsomAsync(string siteUrl, List<DeltaItemChange> deltaChanges, List<DeltaItemChange> deletedChanges)
    {
        _logger.LogInformation("Attempting to enrich deleted items using CSOM approach.");

        // Log the deleted changes we're trying to enrich
        foreach (var change in deletedChanges)
        {
            _logger.LogDebug("CSOM: Looking for deleted item with ID: {itemId}", change.ItemId);
        }

        try
        {
            var spoClient = ConnectionHelper.GetContext(_appSettings, siteUrl, _logger);

            // Use the Recycle Bin API to attempt to retrieve additional information about the deleted items
            spoClient.Load(spoClient.Web.RecycleBin);
            await spoClient.ExecuteQueryRetryAsync();
            RecycleBinItemCollection recycleBinItems = spoClient.Web.RecycleBin;

            _logger.LogInformation("Found {count} items in first stage recycle bin via CSOM.", recycleBinItems?.Count ?? 0);

            if (recycleBinItems == null || recycleBinItems.Count == 0)
            {
                _logger.LogInformation("No items found in the recycle bin. Checking the second stage recycle bin.");
                recycleBinItems = spoClient.Web.GetRecycleBinItems(null, 1000, false, RecycleBinOrderBy.DefaultOrderBy, RecycleBinItemState.SecondStageRecycleBin);
                spoClient.Load(recycleBinItems);
                await spoClient.ExecuteQueryRetryAsync();

                _logger.LogInformation("Found {count} items in second stage recycle bin via CSOM.", recycleBinItems?.Count ?? 0);

                if (recycleBinItems == null || recycleBinItems.Count == 0)
                {
                    _logger.LogInformation("No items found in the second stage recycle bin either. Cannot enrich deleted items.");
                    return deltaChanges;
                }
            }

            var enrichedCount = 0;

            // Log all recycle bin item IDs for debugging
            _logger.LogDebug("CSOM Recycle bin items found:");
            foreach (var rbItem in recycleBinItems)
            {
                _logger.LogDebug("CSOM Recycle bin item - ID: '{id}', Title: '{title}', ItemType: {itemType}", 
                    rbItem.Id, rbItem.Title, rbItem.ItemType);
            }

            // For deleted items, we may want to fetch additional information if needed
            foreach (var change in deletedChanges)
            {
                _logger.LogDebug("CSOM: Searching for change item ID: '{itemId}'", change.ItemId);

                // Try multiple matching strategies due to potential ID format differences
                var deletedItem = recycleBinItems.FirstOrDefault(r => 
                    r.Id.ToString() == change.ItemId ||
                    r.Id.ToString() == change.ItemId.Split(':').LastOrDefault() ||
                    change.ItemId.EndsWith($":{r.Id}") ||
                    change.ItemId.Contains(r.Id.ToString()) ||
                    r.Id.ToString().Contains(change.ItemId));

                if (deletedItem == null)
                {
                    _logger.LogDebug("CSOM: ✗ No matching recycle bin item found for change item ID: '{itemId}'", change.ItemId);
                    continue;
                }

                _logger.LogInformation("CSOM: ✓ Found matching recycle bin item for change item ID: '{itemId}' -> Recycle bin ID: '{rbId}'", 
                    change.ItemId, deletedItem.Id);

                // Update the change item with additional information from the recycle bin if available
                if (change.Item == null)
                {
                    change.Item = new Microsoft.Graph.Models.ListItem();
                }

                // Map relevant properties from the deleted item in the recycle bin to our ListItem model
                change.Item.Name = deletedItem.Title;

                if (change.Item.Deleted == null)
                {
                    change.Item.Deleted = new Microsoft.Graph.Models.Deleted();
                }

                change.Item.Deleted.AdditionalData["DeletedDateTime"] = deletedItem.DeletedDate;
                change.Item.Deleted.AdditionalData["DeletedBy"] = deletedItem.DeletedBy?.UserPrincipalName;

                enrichedCount++;
            }

            _logger.LogInformation("Successfully enriched {count} deleted items using CSOM approach.", enrichedCount);
            return deltaChanges;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error in CSOM enrichment approach");
            return deltaChanges;
        }
    }

    /// <summary>
    /// Classifies the change type of a SharePoint list item
    /// </summary>
    /// <param name="item">The list item to classify</param>
    /// <returns>The change type classification</returns>
    private DeltaChangeType ClassifyChangeType(graph.ListItem item)
    {
        // Check if item contains the deleted marker
        if (item.Deleted != null)
        {
            return DeltaChangeType.Deleted;
        }

        // For non-deleted items, we classify based on creation vs modification time
        // This is a simplified approach - in a real scenario, you might want to
        // track known items separately to distinguish between created and updated
        if (item.CreatedDateTime.HasValue && item.LastModifiedDateTime.HasValue)
        {
            var timeDifference = item.LastModifiedDateTime.Value - item.CreatedDateTime.Value;
            // If modified within a 5 minutes of creation, consider it newly created
            if (timeDifference.TotalSeconds < (60 * 5))
            {
                return DeltaChangeType.Created;
            }
        }

        // Default to Updated for existing items that appear without deleted marker
        return DeltaChangeType.Updated;
    }

    /// <summary>
    /// Extracts the delta link from the delta response
    /// </summary>
    /// <param name="deltaResponse">The delta response object</param>
    /// <returns>The delta link URL or null if not found</returns>
    private string? ExtractDeltaLink(object deltaResponse)
    {
        try
        {
            // Try to get the delta link from OdataDeltaLink property
            var deltaLinkProperty = deltaResponse.GetType().GetProperty("OdataDeltaLink");
            if (deltaLinkProperty != null)
            {
                var deltaLink = deltaLinkProperty.GetValue(deltaResponse) as string;
                if (!string.IsNullOrEmpty(deltaLink))
                {
                    return deltaLink;
                }
            }

            _logger.LogWarning("Could not extract delta link from response");
            return null;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error extracting delta link from response");
            return null;
        }
    }

    /// <summary>
    /// Saves the delta checkpoint for a subscription
    /// </summary>
    /// <param name="subscriptionId">The subscription ID</param>
    /// <param name="deltaLink">The delta link</param>
    private async Task SaveDeltaCheckpointAsync(string subscriptionId, string deltaLink)
    {
        var deltaEntity = new DeltaEntity
        {
            RowKey = subscriptionId,
            DeltaLink = deltaLink,
            Timestamp = DateTimeOffset.UtcNow
        };

        await SaveAsync(deltaEntity);
    }
}
