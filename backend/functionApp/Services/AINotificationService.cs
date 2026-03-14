using functionApp.Models;
using GitHub.Copilot.SDK;
using Microsoft.Extensions.Logging;
using System.Text.Json;

namespace functionApp.Services;

public class AINotificationService
{
    private readonly ILogger<AINotificationService> _logger;
    private readonly AppSettings _appSettings;

    public AINotificationService(AppSettings appSettings, ILogger<AINotificationService> logger)
    {
        _logger = logger;
        _appSettings = appSettings;
        _logger.LogInformation("AINotificationService initialized.");
    }

    public async Task<string> ProcessNotificationAsync(List<DeltaItemChange> items, NotificationRegistration registration)
    {
        if (items.Count == 0)
        {
            _logger.LogInformation("No delta items to process for registration {RegistrationId}.", registration.Id);
            return "No changes detected.";
        }

        _logger.LogInformation("Processing {Count} delta items with AI for registration {RegistrationId}.", items.Count, registration.Id);

        try
        {
            var clientOptions = new CopilotClientOptions { Logger = _logger };
            if (!string.IsNullOrEmpty(_appSettings.GitHubToken))
            {
                clientOptions.GitHubToken = _appSettings.GitHubToken;
            }
            await using var client = new CopilotClient(clientOptions);

            await using var session = await client.CreateSessionAsync(new SessionConfig
            {
                Model = "gpt-5",
                SystemMessage = new SystemMessageConfig
                {
                    Mode = SystemMessageMode.Replace,
                    Content = "You are a SharePoint notification assistant. Your job is to summarize list or library changes into clear, concise, human-readable notification messages. Focus on what changed (created, updated, deleted), which items were affected, and any relevant field values. Keep the summary brief and actionable."
                },
                OnPermissionRequest = PermissionHandler.ApproveAll
            });

            var changeSummary = items.Select(i => new
            {
                i.ItemId,
                ChangeType = i.ChangeType.ToString(),
                Fields = i.Item?.Fields?.AdditionalData
            });

            var prompt = $"""
                The user has subscribed to notifications with the following description:
                "{registration.Description ?? "All changes"}"

                Change type filter: {registration.ChangeType}

                Here are the detected changes in the SharePoint list/library:
                {JsonSerializer.Serialize(changeSummary, new JsonSerializerOptions { WriteIndented = true })}

                Please provide a concise, human-readable notification summary of these changes.
                """;

            var reply = await session.SendAndWaitAsync(
                new MessageOptions { Prompt = prompt }//,
                //timeout: TimeSpan.FromSeconds(30)
                );

            var result = reply?.Data?.Content ?? "Changes were detected but could not be summarized.";
            _logger.LogInformation("AI processing complete for registration {RegistrationId}.", registration.Id);
            return result;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error processing notification with AI for registration {RegistrationId}. Falling back to basic summary.", registration.Id);
            return GenerateFallbackSummary(items);
        }
    }

    private static string GenerateFallbackSummary(List<DeltaItemChange> items)
    {
        var created = items.Count(i => i.ChangeType == DeltaChangeType.Created);
        var updated = items.Count(i => i.ChangeType == DeltaChangeType.Updated);
        var deleted = items.Count(i => i.ChangeType == DeltaChangeType.Deleted);

        var parts = new List<string>();
        if (created > 0) parts.Add($"{created} item(s) created");
        if (updated > 0) parts.Add($"{updated} item(s) updated");
        if (deleted > 0) parts.Add($"{deleted} item(s) deleted");

        return parts.Count > 0
            ? $"SharePoint changes detected: {string.Join(", ", parts)}."
            : "No changes detected.";
    }
}
