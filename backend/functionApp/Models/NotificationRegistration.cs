using System.Text.Json.Serialization;

namespace functionApp.Models;

public enum ChangeType
{
    CREATED,
    UPDATED,
    DELETED,
    ALL
}

public enum NotificationChannel
{
    TEAMS,
    EMAIL
}

public class NotificationRegistration
{
    public Guid Id { get; set; }
    public Guid UserId { get; set; }
    public ChangeType ChangeType { get; set; }
    public Guid SiteId { get; set; }
    public Guid WebId { get; set; }
    public Guid ListId { get; set; }
    public int? ItemId { get; set; }
    public NotificationChannel[] NotificationChannels { get; set; } = [];
    public string? Description { get; set; }
}
