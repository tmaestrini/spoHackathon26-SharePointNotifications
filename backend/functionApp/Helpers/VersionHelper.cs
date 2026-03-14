using functionApp.Models;
using Microsoft.Extensions.Logging;
using Microsoft.SharePoint.Client;

namespace functionApp.Helpers;

public static class VersionHelper
{
    /// <summary>
    /// Enriches non-deleted delta changes with current/previous version information and file content for documents.
    /// </summary>
    public static async Task EnrichVersionInformationAsync(
        List<DeltaItemChange> deltaChanges,
        AppSettings appSettings,
        string siteUrl,
        string listId,
        ILogger logger)
    {
        var nonDeletedChanges = deltaChanges
            .Where(c => c.ChangeType != DeltaChangeType.Deleted)
            .ToList();

        if (nonDeletedChanges.Count == 0)
        {
            logger.LogInformation("No non-deleted items to enrich with version information.");
            return;
        }

        logger.LogInformation("Enriching {Count} non-deleted items with version information.", nonDeletedChanges.Count);

        ClientContext? ctx = null;
        try
        {
            ctx = ConnectionHelper.GetContext(appSettings, siteUrl, logger);
            var web = ctx.Web;
            ctx.Load(web, w => w.Lists);
            var list = Guid.TryParse(listId, out _)
                ? web.Lists.GetById(Guid.Parse(listId))
                : web.Lists.GetByTitle(listId);
            ctx.Load(list);
            await ctx.ExecuteQueryRetryAsync();

            foreach (var change in nonDeletedChanges)
            {
                try
                {
                    if (!int.TryParse(change.ItemId, out var numericItemId))
                    {
                        logger.LogWarning("Could not parse item ID '{ItemId}' as integer. Skipping version enrichment.", change.ItemId);
                        continue;
                    }

                    var item = list.GetItemById(numericItemId);
                    ctx.Load(item);
                    ctx.Load(item,
                        x => x.Versions,
                        x => x.File,
                        x => x.File.VroomDriveID,
                        x => x.File.VroomItemID);
                    await ctx.ExecuteQueryRetryAsync();

                    // Build current version info from the item itself
                    change.CurrentVersionInfo = BuildCurrentVersionInfo(item, logger);

                    // Build previous version info from versions collection
                    if (item.Versions.Count > 1)
                    {
                        var previousVersion = item.Versions[1]; // index 0 = current, index 1 = previous
                        ctx.Load(previousVersion, v => v.Changes);
                        await ctx.ExecuteQueryRetryAsync();

                        change.PreviousVersionInfo = BuildPreviousVersionInfo(previousVersion, logger);

                        // Collect field-level changes between versions
                        if (previousVersion.Changes.Any())
                        {
                            change.FieldChanges = previousVersion.Changes
                                .Select(c => new FieldChange
                                {
                                    FieldTitle = c.FieldTitle,
                                    PreviousValue = c.PreviousValue,
                                    NewValue = c.NewValue
                                })
                                .ToList();

                            logger.LogInformation("Item {ItemId}: Found {Count} field changes between versions.",
                                change.ItemId, change.FieldChanges.Count);
                        }
                    }
                    else
                    {
                        logger.LogDebug("Item {ItemId}: Only one version exists, no previous version available.", change.ItemId);
                    }

                    // If it's a document (has a file), get current and previous file content
                    if (item.File != null && item.File.ServerObjectIsNull == false)
                    {
                        await EnrichFileVersionsAsync(ctx, item, change, logger);
                    }
                }
                catch (Exception ex)
                {
                    logger.LogWarning(ex, "Could not fetch version information for item {ItemId}. Skipping.", change.ItemId);
                }
            }
        }
        finally
        {
            ctx?.Dispose();
        }
    }

    private static ItemVersionInfo BuildCurrentVersionInfo(ListItem item, ILogger logger)
    {
        var versionInfo = new ItemVersionInfo
        {
            VersionLabel = item.Versions.Count > 0 ? item.Versions[0].VersionLabel : "1.0",
        };

        // Collect field values from the current item
        foreach (var fieldValue in item.FieldValues)
        {
            versionInfo.FieldValues[fieldValue.Key] = fieldValue.Value?.ToString() ?? string.Empty;
        }

        logger.LogDebug("Built current version info: version {VersionLabel} with {FieldCount} fields.",
            versionInfo.VersionLabel, versionInfo.FieldValues.Count);

        return versionInfo;
    }

    private static ItemVersionInfo BuildPreviousVersionInfo(ListItemVersion version, ILogger logger)
    {
        var versionInfo = new ItemVersionInfo
        {
            VersionLabel = version.VersionLabel,
        };

        // Collect field values from the previous version
        foreach (var fieldValue in version.FieldValues)
        {
            versionInfo.FieldValues[fieldValue.Key] = fieldValue.Value?.ToString() ?? string.Empty;
        }

        logger.LogDebug("Built previous version info: version {VersionLabel} with {FieldCount} fields.",
            versionInfo.VersionLabel, versionInfo.FieldValues.Count);

        return versionInfo;
    }

    private static async Task EnrichFileVersionsAsync(
        ClientContext ctx,
        ListItem item,
        DeltaItemChange change,
        ILogger logger)
    {
        try
        {
            var file = item.File;
            ctx.Load(file, f => f.Name, f => f.ServerRelativeUrl, f => f.Versions);
            await ctx.ExecuteQueryRetryAsync();

            logger.LogInformation("Item {ItemId} is a document: '{FileName}'. Fetching file versions.",
                change.ItemId, file.Name);

            // Get current file content
            var currentFileData = file.OpenBinaryStream();
            await ctx.ExecuteQueryRetryAsync();

            using (var ms = new MemoryStream())
            {
                currentFileData.Value.CopyTo(ms);
                change.CurrentFileContent = ms.ToArray();
                change.CurrentFileName = file.Name;
                logger.LogDebug("Retrieved current file content for '{FileName}': {Size} bytes.",
                    file.Name, change.CurrentFileContent.Length);
            }

            // Get previous file version content if available
            if (file.Versions.Count > 0)
            {
                var previousFileVersion = file.Versions[file.Versions.Count - 1]; // most recent previous version
                ctx.Load(previousFileVersion, v => v.VersionLabel, v => v.Url);
                await ctx.ExecuteQueryRetryAsync();

                var previousFileData = previousFileVersion.OpenBinaryStream();
                await ctx.ExecuteQueryRetryAsync();

                using (var ms = new MemoryStream())
                {
                    previousFileData.Value.CopyTo(ms);
                    change.PreviousFileContent = ms.ToArray();
                    change.PreviousFileVersionLabel = previousFileVersion.VersionLabel;
                    logger.LogDebug("Retrieved previous file version '{VersionLabel}' for '{FileName}': {Size} bytes.",
                        previousFileVersion.VersionLabel, file.Name, change.PreviousFileContent.Length);
                }
            }
            else
            {
                logger.LogDebug("No previous file versions available for '{FileName}'.", file.Name);
            }
        }
        catch (Exception ex)
        {
            logger.LogWarning(ex, "Could not fetch file versions for item {ItemId}. Skipping file enrichment.", change.ItemId);
        }
    }

    /// <summary>
    /// Enriches non-deleted document items with their full file URL.
    /// </summary>
    public static async Task EnrichFileUrlsAsync(
        List<DeltaItemChange> deltaChanges,
        AppSettings appSettings,
        string siteUrl,
        string listId,
        string tenantName,
        ILogger logger)
    {
        var documentChanges = deltaChanges
            .Where(c => c.ChangeType != DeltaChangeType.Deleted && c.CurrentFileName != null)
            .ToList();

        if (documentChanges.Count == 0)
        {
            logger.LogInformation("No document items to enrich with file URLs.");
            return;
        }

        logger.LogInformation("Enriching {Count} document items with file URLs.", documentChanges.Count);

        var tenantUrl = $"https://{tenantName}";
        ClientContext? ctx = null;
        try
        {
            ctx = ConnectionHelper.GetContext(appSettings, siteUrl, logger);
            var list = Guid.TryParse(listId, out _)
                ? ctx.Web.Lists.GetById(Guid.Parse(listId))
                : ctx.Web.Lists.GetByTitle(listId);

            foreach (var change in documentChanges)
            {
                try
                {
                    if (!int.TryParse(change.ItemId, out var numericItemId))
                    {
                        logger.LogWarning("Could not parse item ID '{ItemId}' as integer. Skipping file URL enrichment.", change.ItemId);
                        continue;
                    }

                    var item = list.GetItemById(numericItemId);
                    ctx.Load(item, i => i.File);
                    ctx.Load(item.File, f => f.ServerRelativeUrl);
                    await ctx.ExecuteQueryRetryAsync();

                    if (item.File != null && !string.IsNullOrEmpty(item.File.ServerRelativeUrl))
                    {
                        change.FileUrl = $"{tenantUrl}{item.File.ServerRelativeUrl}";
                        logger.LogDebug("Item {ItemId}: File URL set to {FileUrl}", change.ItemId, change.FileUrl);
                    }
                }
                catch (Exception ex)
                {
                    logger.LogWarning(ex, "Could not retrieve file URL for item {ItemId}. Skipping.", change.ItemId);
                }
            }
        }
        finally
        {
            ctx?.Dispose();
        }
    }
}
