namespace functionApp.Models;

public class DeltaItemChange
{
    public string ItemId { get; set; } = string.Empty;
    public DeltaChangeType ChangeType { get; set; }
    public Microsoft.Graph.Models.ListItem? Item { get; set; }
}