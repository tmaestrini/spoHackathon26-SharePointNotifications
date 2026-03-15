using Azure;
using Azure.Data.Tables;

namespace functionApp.Models;

public class DeltaEntity : ITableEntity
{
    public string PartitionKey { get; set; } = "Deltas";
    public string RowKey { get; set; } = string.Empty; // Subscription Id
    public DateTimeOffset? Timestamp { get; set; }
    public ETag ETag { get; set; }

    public string DeltaLink { get; set; } = string.Empty;
}
