using Azure;
using Azure.Data.Tables;

namespace functionApp.Models;

public class WebhookSubscriptionEntity : ITableEntity
{
    public string PartitionKey { get; set; } = "Webhooks";
    public string RowKey { get; set; } = string.Empty; // Subscription Id
    public DateTimeOffset? Timestamp { get; set; }
    public ETag ETag { get; set; }

    public Guid ListId { get; set; }
    public string SiteUrl { get; set; } = string.Empty;
    public string NotificationUrl { get; set; } = string.Empty;
    public string Resource { get; set; } = string.Empty;
    public string? ChangeToken { get; set; }
    public DateTime ExpirationDateTime { get; set; }
    public DateTime LastUpdated { get; set; }
}
