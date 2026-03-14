using functionApp.Models;
using Microsoft.Extensions.Logging;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace functionApp.Services;

public class FoundryAINotificationService
{
    private readonly ILogger<FoundryAINotificationService> _logger;
    private readonly AppSettings _appSettings;
    private readonly IHttpClientFactory _httpClientFactory;

    private static readonly JsonSerializerOptions _jsonOptions = new()
    {
        PropertyNameCaseInsensitive = true
    };

    public FoundryAINotificationService(
        AppSettings appSettings,
        ILogger<FoundryAINotificationService> logger,
        IHttpClientFactory httpClientFactory)
    {
        _logger = logger;
        _appSettings = appSettings;
        _httpClientFactory = httpClientFactory;
        _logger.LogInformation("FoundryAINotificationService initialized.");
    }

    public async Task<string> ProcessNotificationAsync(List<DeltaItemChange> items, NotificationRegistration registration)
    {
        if (items.Count == 0)
        {
            _logger.LogInformation("No delta items to process for registration {RegistrationId}.", registration.Id);
            return "No changes detected.";
        }

        _logger.LogInformation("Processing {Count} delta items with Azure AI Foundry for registration {RegistrationId}.",
            items.Count, registration.Id);

        try
        {
            var changeSummary = items.Select(i => new
            {
                i.ItemId,
                ChangeType = i.ChangeType.ToString(),
                Fields = i.Item?.Fields?.AdditionalData
            });

            var systemMessage = "You are a SharePoint notification assistant. Your job is to summarize list or library changes into clear, concise, human-readable notification messages. Focus on what changed (created, updated, deleted), which items were affected, and any relevant field values. Keep the summary brief and actionable.";

            var userMessage = $"""
                The user has subscribed to notifications with the following description:
                "{registration.Description ?? "All changes"}"

                Change type filter: {registration.ChangeType}

                Here are the detected changes in the SharePoint list/library:
                {JsonSerializer.Serialize(changeSummary, new JsonSerializerOptions { WriteIndented = true })}

                Please provide a concise, human-readable notification summary of these changes.
                """;

            //var requestUrl = $"{_appSettings.AzureFoundryApiUrl}/openai/deployments/{_appSettings.AzureFoundryModelName}/chat/completions?api-version=2024-08-01-preview";
            var requestUrl = _appSettings.AzureFoundryApiUrl;
            var requestBody = new
            {
                messages = new[]
                {
                    new { role = "system", content = systemMessage },
                    new { role = "user", content = userMessage }
                },
                max_tokens = 2000,
                temperature = 0.7
            };

            var httpClient = _httpClientFactory.CreateClient();
            using var request = new HttpRequestMessage(HttpMethod.Post, requestUrl);
            request.Content = new StringContent(
                JsonSerializer.Serialize(requestBody),
                Encoding.UTF8,
                "application/json");
            request.Headers.Add("api-key", _appSettings.AzureFoundryApiKey);

            _logger.LogInformation("Calling Azure AI Foundry at: {Url}", requestUrl);

            var response = await httpClient.SendAsync(request);
            var responseJson = await response.Content.ReadAsStringAsync();

            if (!response.IsSuccessStatusCode)
            {
                _logger.LogError("Azure AI Foundry returned {StatusCode}: {Body}", response.StatusCode, responseJson);
                response.EnsureSuccessStatusCode();
            }
            var data = JsonSerializer.Deserialize<ChatCompletionResponse>(responseJson, _jsonOptions);

            var result = data?.Choices?.FirstOrDefault()?.Message?.Content
                ?? "Changes were detected but could not be summarized.";

            _logger.LogInformation("Azure AI Foundry processing complete for registration {RegistrationId}.", registration.Id);
            return result;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error processing notification with Azure AI Foundry for registration {RegistrationId}. Falling back to basic summary.", registration.Id);
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

    // Response models for Azure OpenAI chat completions
    private class ChatCompletionResponse
    {
        [JsonPropertyName("choices")]
        public List<ChatChoice>? Choices { get; set; }
    }

    private class ChatChoice
    {
        [JsonPropertyName("message")]
        public ChatMessage? Message { get; set; }
    }

    private class ChatMessage
    {
        [JsonPropertyName("content")]
        public string? Content { get; set; }
    }
}
