using Azure.Data.Tables;
using Azure;
using System.Text.Json;

namespace functionApp.Models;

public class NotificationRegistrationEntity : ITableEntity
{
    public string PartitionKey { get; set; } = string.Empty; // UserId
    public string RowKey { get; set; } = string.Empty; // Registration Id
    public DateTimeOffset? Timestamp { get; set; }
    public ETag ETag { get; set; }

    public string ChangeType { get; set; } = string.Empty;
    public Guid SiteId { get; set; }
    public Guid WebId { get; set; }
    public Guid ListId { get; set; }
    public int? ItemId { get; set; }
    public string NotificationChannelsJson { get; set; } = "[]";
    public string? Description { get; set; }

    public static NotificationRegistrationEntity FromModel(NotificationRegistration model)
    {
        return new NotificationRegistrationEntity
        {
            PartitionKey = model.UserId.ToString(),
            RowKey = model.Id.ToString(),
            ChangeType = model.ChangeType.ToString(),
            SiteId = model.SiteId,
            WebId = model.WebId,
            ListId = model.ListId,
            ItemId = model.ItemId,
            NotificationChannelsJson = JsonSerializer.Serialize(model.NotificationChannels),
            Description = model.Description
        };
    }

    public NotificationRegistration ToModel()
    {
        return new NotificationRegistration
        {
            Id = Guid.Parse(RowKey),
            UserId = Guid.Parse(PartitionKey),
            ChangeType = Enum.Parse<ChangeType>(ChangeType),
            SiteId = SiteId,
            WebId = WebId,
            ListId = ListId,
            ItemId = ItemId,
            NotificationChannels = JsonSerializer.Deserialize<NotificationChannel[]>(NotificationChannelsJson) ?? [],
            Description = Description
        };
    }
}
