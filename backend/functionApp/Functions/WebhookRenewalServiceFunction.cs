using functionApp.Services;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Extensions.Logging;

namespace functionApp.Functions;

public class WebhookRenewalServiceFunction
{
    private readonly ILogger<WebhookRenewalServiceFunction> _logger;
    private readonly WebhookService _webhookService;
    private readonly WebhookSubscriptionService _subscriptionService;

    public WebhookRenewalServiceFunction(
        ILogger<WebhookRenewalServiceFunction> logger,
        WebhookService webhookService,
        WebhookSubscriptionService subscriptionService)
    {
        _logger = logger;
        _webhookService = webhookService;
        _subscriptionService = subscriptionService;
    }

    /// <summary>
    /// Runs on the 1st of every month at 5:00 AM UTC.
    /// Finds all webhook subscriptions expiring within the next 45 days and renews them.
    /// </summary>
    [Function("MonthlyRenewWebhooks")]
    public async Task RunMonthlyRenewWebhooks([TimerTrigger("0 0 5 1 * *")] TimerInfo myTimer)
    {
        _logger.LogInformation("Monthly webhook renewal timer triggered at {ExecutionTime}.", DateTime.UtcNow);

        // Find webhooks expiring within the next 45 days
        var targetDate = DateTime.UtcNow.AddDays(45);
        var expiringSubscriptions = await _subscriptionService.GetExpiringAsync(targetDate);

        _logger.LogInformation("Found {Count} webhook subscriptions expiring before {TargetDate}.",
            expiringSubscriptions.Count, targetDate);

        var renewed = 0;
        var failed = 0;

        foreach (var subscription in expiringSubscriptions)
        {
            try
            {
                _logger.LogInformation("Renewing webhook {SubscriptionId} for site {SiteUrl}, list {ListId}. Current expiration: {Expiration}.",
                    subscription.RowKey, subscription.SiteUrl, subscription.ListId, subscription.ExpirationDateTime);

                var newExpiration = await _webhookService.RenewWebhookAsync(subscription);

                if (newExpiration.HasValue)
                {
                    _logger.LogInformation("Webhook {SubscriptionId} renewed. New expiration: {Expiration}.",
                        subscription.RowKey, newExpiration.Value);
                    renewed++;
                }
                else
                {
                    _logger.LogWarning("Webhook {SubscriptionId} was stale and has been removed.", subscription.RowKey);
                    failed++;
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Failed to renew webhook {SubscriptionId} for site {SiteUrl}.",
                    subscription.RowKey, subscription.SiteUrl);
                failed++;
            }
        }

        _logger.LogInformation("Webhook renewal completed. Renewed: {Renewed}, Failed: {Failed}, Total: {Total}.",
            renewed, failed, expiringSubscriptions.Count);
    }
}
