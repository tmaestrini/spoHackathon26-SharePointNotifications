using Azure.Data.Tables;
using functionApp.Models;
using Microsoft.Extensions.Logging;

namespace functionApp.Services;

public class WebhookSubscriptionService
{
    private readonly TableClient _tableClient;
    private readonly ILogger<WebhookSubscriptionService> _logger;

    public WebhookSubscriptionService(TableServiceClient tableServiceClient, AppSettings appSettings, ILogger<WebhookSubscriptionService> logger)
    {
        _logger = logger;
        _tableClient = tableServiceClient.GetTableClient(appSettings.TableWebhookSubscriptions);
        _tableClient.CreateIfNotExists();
        _logger.LogInformation("WebhookSubscriptionService initialized with table '{TableName}'.", appSettings.TableWebhookSubscriptions);
    }

    public async Task SaveAsync(WebhookSubscriptionEntity entity)
    {
        _logger.LogInformation("Saving webhook subscription {SubscriptionId} for list {ListId}.", entity.RowKey, entity.ListId);
        await _tableClient.UpsertEntityAsync(entity, TableUpdateMode.Replace);
        _logger.LogInformation("Webhook subscription {SubscriptionId} saved successfully.", entity.RowKey);
    }

    public async Task<WebhookSubscriptionEntity?> GetAsync(string subscriptionId)
    {
        _logger.LogInformation("Fetching webhook subscription {SubscriptionId}.", subscriptionId);
        try
        {
            var response = await _tableClient.GetEntityAsync<WebhookSubscriptionEntity>("Webhooks", subscriptionId);
            return response.Value;
        }
        catch (Azure.RequestFailedException ex) when (ex.Status == 404)
        {
            _logger.LogWarning("Webhook subscription {SubscriptionId} not found.", subscriptionId);
            return null;
        }
    }

    public async Task<List<WebhookSubscriptionEntity>> GetByListAsync(Guid listId)
    {
        _logger.LogInformation("Fetching webhook subscriptions for list {ListId}.", listId);
        var results = new List<WebhookSubscriptionEntity>();
        await foreach (var entity in _tableClient.QueryAsync<WebhookSubscriptionEntity>(
            e => e.PartitionKey == "Webhooks" && e.ListId == listId))
        {
            results.Add(entity);
        }
        _logger.LogInformation("Found {Count} webhook subscriptions for list {ListId}.", results.Count, listId);
        return results;
    }

    public async Task<List<WebhookSubscriptionEntity>> GetExpiringAsync(DateTime before)
    {
        _logger.LogInformation("Fetching webhook subscriptions expiring before {Before}.", before);
        var results = new List<WebhookSubscriptionEntity>();
        await foreach (var entity in _tableClient.QueryAsync<WebhookSubscriptionEntity>(
            e => e.PartitionKey == "Webhooks" && e.ExpirationDateTime < before))
        {
            results.Add(entity);
        }
        _logger.LogInformation("Found {Count} expiring webhook subscriptions.", results.Count);
        return results;
    }

    public async Task<bool> DeleteAsync(string subscriptionId)
    {
        _logger.LogInformation("Deleting webhook subscription {SubscriptionId}.", subscriptionId);
        try
        {
            await _tableClient.DeleteEntityAsync("Webhooks", subscriptionId);
            _logger.LogInformation("Webhook subscription {SubscriptionId} deleted successfully.", subscriptionId);
            return true;
        }
        catch (Azure.RequestFailedException ex) when (ex.Status == 404)
        {
            _logger.LogWarning("Webhook subscription {SubscriptionId} not found for deletion.", subscriptionId);
            return false;
        }
    }
}
