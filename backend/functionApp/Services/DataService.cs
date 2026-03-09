using Azure.Data.Tables;
using Azure.Identity;
using functionApp.Helpers;
using functionApp.Models;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using Microsoft.Graph.Applications.Delta;
using Microsoft.Graph.Models;
using PnP.Framework.Provisioning.Model.Drive;
using System.Text.Json;
using System.Web;

namespace functionApp.Services;

public class DataService
{
    private readonly ILogger<DataService> _logger;
    private readonly AppSettings _appSettings;
    private readonly HttpClient _httpClient;

    public DataService(AppSettings appSettings, ILogger<DataService> logger, HttpClient httpClient)
    {
        _logger = logger;
        _appSettings = appSettings;
        _logger.LogInformation("DataService initialized.");
    }

    /// <summary>
    /// Handles SharePoint library delta changes for the specified registration
    /// </summary>
    /// <param name="registration">The notification registration containing site and list information</param>
    /// <param name="webhookNotification">The webhook notification that triggered this delta request</param>
    /// <returns>List of changed items from the SharePoint delta endpoint</returns>
    public async Task<List<ListItem>> GetDeltaAsync(WebhookNotificationModel webhookNotification)
    {
        List<ListItem> changedItems = new List<ListItem>();

        try
        {
            _logger.LogInformation($"Processing delta for webhook subscription: {webhookNotification.SubscriptionId}");

            var graphClient = ConnectionHelper.GraphClient(_appSettings, _logger);

            // Get the site information using the site URL from the webhook notification
            var site = await graphClient
                                .Sites[$"{_appSettings.SharePointTenantName}:{webhookNotification.SiteUrl}"]
                                .GetAsync();

            var listId = webhookNotification.Resource;

            var delta = await graphClient
                            .Sites[site.Id]
                            .Lists[listId]
                            .Items
                            .Delta
                            .GetAsDeltaGetResponseAsync();

            var deltaResponse = await graphClient
                                        .Sites[site.Id]
                                        .Lists[listId]
                                        .Items
                                        .Delta
                                        .GetAsDeltaGetResponseAsync((x) =>
                                        {
                                            x.QueryParameters.Top = 25;
                                        });

            // Iterator to cycle through all the pages of the delta response and collect changed items
            var pageIterator = PageIterator<ListItem, Microsoft.Graph.Sites.Item.Lists.Item.Items.Delta.DeltaGetResponse>
                                .CreatePageIterator(
                                    graphClient,
                                    deltaResponse,
                                    (item) =>
                                    {
                                        Console.WriteLine(item.Id);
                                        changedItems.Add(item);
                                        return true;
                                    });

            await pageIterator.IterateAsync();

            _logger.LogInformation($"Processed delta for webhook subscription: {webhookNotification.SubscriptionId}");
            _logger.LogInformation($"Found {changedItems.Count} changed items.");

            return changedItems;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, $"Error processing delta for subscription: {webhookNotification.SubscriptionId}");
            throw;
        }
    }
}
