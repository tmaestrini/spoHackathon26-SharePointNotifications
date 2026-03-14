using System.Text.Json.Serialization;

namespace functionApp.Models;

[JsonConverter(typeof(JsonStringEnumConverter))]
public enum ChangeType
{
    CREATED,
    UPDATED,
    DELETED,
    ALL
}

[JsonConverter(typeof(JsonStringEnumConverter))]
public enum NotificationChannel
{
    TEAMS,
    EMAIL
}

public class NotificationRegistration
{
    [JsonPropertyName("id")]
    public Guid Id { get; set; }

    [JsonPropertyName("userId")]
    public Guid UserId { get; set; }

    [JsonPropertyName("changeType")]
    public ChangeType ChangeType { get; set; }

    [JsonPropertyName("siteId")]
    public Guid SiteId { get; set; }

    [JsonPropertyName("webId")]
    public Guid WebId { get; set; }

    [JsonPropertyName("listId")]
    public Guid ListId { get; set; }

    [JsonPropertyName("siteUrl")]
    public string? SiteUrl { get; set; }

    [JsonPropertyName("itemId")]
    public int? ItemId { get; set; }

    [JsonPropertyName("notificationChannels")]
    public NotificationChannel[] NotificationChannels { get; set; } = [];

    [JsonPropertyName("description")]
    public string? Description { get; set; }
}
