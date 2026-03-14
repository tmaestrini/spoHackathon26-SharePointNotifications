namespace functionApp.Models;

public class DeltaItemChange
{
    public string ItemId { get; set; } = string.Empty;
    public DeltaChangeType ChangeType { get; set; }
    public Microsoft.Graph.Models.ListItem? Item { get; set; }

    /// <summary>Current version metadata and field values</summary>
    public ItemVersionInfo? CurrentVersionInfo { get; set; }

    /// <summary>Previous version metadata and field values</summary>
    public ItemVersionInfo? PreviousVersionInfo { get; set; }

    /// <summary>Field-level changes between current and previous version</summary>
    public List<FieldChange>? FieldChanges { get; set; }

    /// <summary>Current file content (for documents only)</summary>
    public byte[]? CurrentFileContent { get; set; }

    /// <summary>Current file name (for documents only)</summary>
    public string? CurrentFileName { get; set; }

    /// <summary>Full URL to the file (for documents only)</summary>
    public string? FileUrl { get; set; }

    /// <summary>Previous file version content (for documents only)</summary>
    public byte[]? PreviousFileContent { get; set; }

    /// <summary>Previous file version label (for documents only)</summary>
    public string? PreviousFileVersionLabel { get; set; }
}

public class ItemVersionInfo
{
    public string VersionLabel { get; set; } = string.Empty;
    public Dictionary<string, string> FieldValues { get; set; } = new();
}

public class FieldChange
{
    public string FieldTitle { get; set; } = string.Empty;
    public string? PreviousValue { get; set; }
    public string? NewValue { get; set; }
}