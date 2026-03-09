using System.Text.Json.Serialization;

namespace functionApp.Models;

/// <summary>
/// Represents the wrapper object that contains an array of webhook notifications
/// </summary>
public class WebhookNotificationData
{
    [JsonPropertyName("value")]
    public WebhookNotificationModel[] Value { get; set; } = [];
}

/// <summary>
/// Represents a message that will be queued for processing
/// </summary>
public class NotificationQueueMessage
{
    public List<NotificationRegistration> Registrations { get; set; } = null!;
    public WebhookNotificationModel WebhookNotification { get; set; } = null!;
    public DateTime QueuedAt { get; set; }
}

/// <summary>
/// Helper class to parse resource information from webhook URLs
/// </summary>
public class ResourceInfo
{
    public Guid SiteId { get; set; }
    public Guid WebId { get; set; }
    public Guid ListId { get; set; }
    public int? ItemId { get; set; }
}