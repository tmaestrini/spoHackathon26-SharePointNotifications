using functionApp.Helpers;
using functionApp.Models;
using Microsoft.Extensions.Logging;
using Microsoft.SharePoint.Client;

namespace functionApp.Services;

public class WebhookService
{
    private readonly ILogger<WebhookService> _logger;
    private readonly AppSettings _appSettings;
    private readonly WebhookSubscriptionService _subscriptionService;

    public WebhookService(AppSettings appSettings, WebhookSubscriptionService subscriptionService, ILogger<WebhookService> logger)
    {
        _appSettings = appSettings;
        _subscriptionService = subscriptionService;
        _logger = logger;
    }

    /// <summary>
    /// Registers a SharePoint webhook subscription on the specified list/library.
    /// </summary>
    /// <param name="siteUrl">The full URL of the SharePoint site (e.g. https://tenant.sharepoint.com/sites/MySite).</param>
    /// <param name="listId">The GUID of the list or library to subscribe to.</param>
    /// <returns>True if the webhook was registered; false otherwise.</returns>
    public async Task<bool> RegisterWebhookAsync(string siteUrl, Guid listId)
    {
        if (string.IsNullOrEmpty(_appSettings.WebhookUrl))
        {
            _logger.LogError("WebhookUrl is not configured in application settings.");
            throw new InvalidOperationException("WebhookUrl is not configured.");
        }

        _logger.LogInformation("Registering webhook for site: {SiteUrl}, list: {ListId}, notificationUrl: {WebhookUrl}",
            siteUrl, listId, _appSettings.WebhookUrl);

        try
        {
            var ctx = _appSettings.GetContext(siteUrl, _logger);

            var list = ctx.Web.Lists.GetById(listId);
            ctx.Load(list);
            await ctx.ExecuteQueryAsync();

            // Register the webhook subscription (expirationInMonths: 4 months, close to SharePoint's 180-day max)
            var subscription = list.AddWebhookSubscription(_appSettings.WebhookUrl, 4);
            await ctx.ExecuteQueryAsync();

            if (subscription != null)
            {
                _logger.LogInformation("Webhook subscription created successfully. SubscriptionId: {SubscriptionId}, ExpirationDateTime: {Expiration}",
                    subscription.Id, subscription.ExpirationDateTime);

                // Persist the subscription to Azure Table Storage for expiry tracking and renewal
                var entity = new WebhookSubscriptionEntity
                {
                    PartitionKey = "Webhooks",
                    RowKey = subscription.Id,
                    ListId = listId,
                    SiteUrl = siteUrl,
                    NotificationUrl = subscription.NotificationUrl,
                    Resource = subscription.Resource,
                    ExpirationDateTime = subscription.ExpirationDateTime,
                    LastUpdated = DateTime.UtcNow
                };
                await _subscriptionService.SaveAsync(entity);

                return true;
            }

            _logger.LogWarning("Webhook subscription returned null for site: {SiteUrl}, list: {ListId}.", siteUrl, listId);
            return false;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Failed to register webhook for site: {SiteUrl}, list: {ListId}.", siteUrl, listId);
            throw;
        }
    }

    /// <summary>
    /// Removes the webhook subscription from the specified list that matches the configured WebhookUrl.
    /// </summary>
    /// <param name="siteUrl">The full URL of the SharePoint site.</param>
    /// <param name="listId">The GUID of the list or library.</param>
    /// <returns>True if a matching webhook subscription was found and removed; false if none was found.</returns>
    public async Task<bool> RemoveWebhookAsync(string siteUrl, Guid listId)
    {
        if (string.IsNullOrEmpty(_appSettings.WebhookUrl))
        {
            _logger.LogWarning("WebhookUrl is not configured. Cannot remove webhook.");
            return false;
        }

        _logger.LogInformation("Removing webhook for site: {SiteUrl}, list: {ListId}, notificationUrl: {WebhookUrl}",
            siteUrl, listId, _appSettings.WebhookUrl);

        try
        {
            var ctx = _appSettings.GetContext(siteUrl, _logger);

            var list = ctx.Web.Lists.GetById(listId);
            ctx.Load(list);
            await ctx.ExecuteQueryAsync();

            // Get all webhook subscriptions on the list
            var subscriptions = list.GetWebhookSubscriptions();
            await ctx.ExecuteQueryAsync();

            // Find the subscription that matches our notification URL
            var matchingSubscription = subscriptions
                .FirstOrDefault(s => string.Equals(s.NotificationUrl, _appSettings.WebhookUrl, StringComparison.OrdinalIgnoreCase));

            if (matchingSubscription == null)
            {
                _logger.LogWarning("No webhook subscription found matching URL {WebhookUrl} on list {ListId}.", _appSettings.WebhookUrl, listId);
                return false;
            }

            list.RemoveWebhookSubscription(Guid.Parse(matchingSubscription.Id));
            await ctx.ExecuteQueryAsync();

            // Remove the subscription from Azure Table Storage
            await _subscriptionService.DeleteAsync(matchingSubscription.Id);

            _logger.LogInformation("Webhook subscription {SubscriptionId} removed from list {ListId}.", matchingSubscription.Id, listId);
            return true;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Failed to remove webhook for site: {SiteUrl}, list: {ListId}.", siteUrl, listId);
            throw;
        }
    }

    /// <summary>
    /// Renews a webhook subscription on the specified list, extending its expiration by 6 months.
    /// Updates the subscription record in Azure Table Storage with the new expiration date.
    /// </summary>
    /// <param name="subscription">The webhook subscription entity from table storage.</param>
    /// <returns>The new expiration DateTime if renewal succeeded; null otherwise.</returns>
    public async Task<DateTime?> RenewWebhookAsync(WebhookSubscriptionEntity subscription)
    {
        _logger.LogInformation("Renewing webhook {SubscriptionId} for site: {SiteUrl}, list: {ListId}.",
            subscription.RowKey, subscription.SiteUrl, subscription.ListId);

        try
        {
            var ctx = _appSettings.GetContext(subscription.SiteUrl, _logger);

            var list = ctx.Web.Lists.GetById(subscription.ListId);
            ctx.Load(list);
            await ctx.ExecuteQueryAsync();

            // Get the existing webhook subscription from SharePoint
            var spSubscriptions = list.GetWebhookSubscriptions();
            await ctx.ExecuteQueryAsync();

            var existing = spSubscriptions.FirstOrDefault(s => s.Id == subscription.RowKey);
            if (existing == null)
            {
                _logger.LogWarning("Webhook subscription {SubscriptionId} not found on SharePoint list {ListId}. Removing stale record.",
                    subscription.RowKey, subscription.ListId);
                await _subscriptionService.DeleteAsync(subscription.RowKey);
                return null;
            }

            // Renew by updating the expiration date (6 months from now)
            var newExpiration = DateTime.UtcNow.AddMonths(6);
            existing.ExpirationDateTime = newExpiration;
            list.UpdateWebhookSubscription(existing);
            await ctx.ExecuteQueryAsync();

            // Update the record in table storage
            subscription.ExpirationDateTime = newExpiration;
            subscription.LastUpdated = DateTime.UtcNow;
            await _subscriptionService.SaveAsync(subscription);

            _logger.LogInformation("Webhook {SubscriptionId} renewed. New expiration: {Expiration}.",
                subscription.RowKey, newExpiration);
            return newExpiration;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Failed to renew webhook {SubscriptionId} for site: {SiteUrl}, list: {ListId}.",
                subscription.RowKey, subscription.SiteUrl, subscription.ListId);
            throw;
        }
    }
}
