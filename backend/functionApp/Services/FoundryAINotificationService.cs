using Azure;
using functionApp.Helpers;
using functionApp.Models;
using Microsoft.AspNetCore.DataProtection.KeyManagement;
using Microsoft.Extensions.Logging;
using Microsoft.Graph.Models;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using static System.Reflection.Metadata.BlobBuilder;

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

    public async Task<string> ProcessNotificationAsync(List<DeltaItemChange> items, NotificationRegistration registration, NotificationChannel channel)
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
                Fields = i.Item?.Fields?.AdditionalData,
                CurrentVersion = i.CurrentVersionInfo != null ? new
                {
                    i.CurrentVersionInfo.VersionLabel,
                    i.CurrentVersionInfo.FieldValues
                } : null,
                PreviousVersion = i.PreviousVersionInfo != null ? new
                {
                    i.PreviousVersionInfo.VersionLabel,
                    i.PreviousVersionInfo.FieldValues
                } : null,
                FieldChanges = i.FieldChanges?.Select(fc => new
                {
                    fc.FieldTitle,
                    fc.PreviousValue,
                    fc.NewValue
                }),
                IsDocument = i.CurrentFileName != null,
                FileName = i.CurrentFileName,
                FileUrl = i.FileUrl,
                CurrentFileContent = DocumentTextExtractor.ExtractTextContent(i.CurrentFileContent, i.CurrentFileName),
                PreviousFileContent = DocumentTextExtractor.ExtractTextContent(i.PreviousFileContent, i.CurrentFileName),
                PreviousFileVersionLabel = i.PreviousFileVersionLabel
            });
 
 
            var systemMessage = _appSettings.SystemPrompt;

            var outputFormat = channel == NotificationChannel.EMAIL
                ? _appSettings.EmailPrompt
                : $"""
                 Instructions - Use clean Markdown. Use this exact structure per item.

                 Here is a summary of the changes made:

                 **ChangeType**: Updated/Created/Deleted

                 **Document Name**: Updated Document Name or List Item ID <Add the URL of the file here as a link if it is a file>

                 **Version**: Updated from version X.X to Y.Y (or just current version)

                 **Last Modified By**: Display name of the user not the User principal name

                 **Modified Date**: Date

                 ## Content Changes

                 Describe what text was added, removed, or modified by comparing current and previous file content.

                 ## Metadata Changes

                 **FieldName**: old value → new value (if its a user field then Display name of the user not the User principal name)

                 For Created items, replace "Content Changes" with "Content" (summarize the content) and "Metadata Changes" with "Metadata" (list key values).

                 For Deleted items, omit Content/Metadata sections. Include Last Version, Deleted By, and Deleted On if available.

                 CRITICAL: Every single field label (ChangeType, Document Name, Version, Last Modified By, Modified Date, and any metadata field name) MUST be bold using **label** syntax. No exceptions.
                 """;

            var userMessage = $"""
                The user has subscribed to notifications with the following description:
                "{registration.Description ?? "All changes"}"
                
                Change type filter: {registration.ChangeType}
                
                Here are the detected changes in the SharePoint list/library. If it is a file, then compare the CurrentFileContent and PreviousFileContent to see the difference in changes:
                {JsonSerializer.Serialize(changeSummary, new JsonSerializerOptions { WriteIndented = true })}

                Below is the output format you must follow when generating the summary for these changes. Adhere to the structure and formatting rules exactly
                {outputFormat}
 
                Rules:
                - Replace ALL placeholders with actual data from the JSON — do NOT leave any placeholder text.
                - Include ALL fields that have values in the JSON data. Do NOT skip ChangeType, Document Name, Version, Modified By, or Date.
                - For Updated items with file content, ALWAYS include the Content Changes section comparing CurrentFileContent vs PreviousFileContent.
                - For Updated items with FieldChanges, ALWAYS include the Metadata Changes section listing each changed field.
                - Omit sections ONLY when the corresponding data is truly null/empty in the JSON.
                - Output ONLY the formatted summary — nothing else.
                """;
            _logger.LogInformation(systemMessage);
            _logger.LogInformation(outputFormat);
            _logger.LogInformation(userMessage);
            
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

            _logger.LogInformation(data?.Choices?.FirstOrDefault()?.Message?.Content);
            var result = data?.Choices?.FirstOrDefault()?.Message?.Content
                ?? "Changes were detected but could not be summarized.";
            // Strip (```html ... ```) that the AI may wrap around output
            result = StripExtraCharactersInEmailContent(result);
            _logger.LogInformation(result);
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

    private static string StripExtraCharactersInEmailContent(string text)
    {
        var trimmed = text.Trim();
        // Remove ```html or ```markdown at the start and ``` at the end
        if (trimmed.StartsWith("```"))
        {
            var firstNewline = trimmed.IndexOf('\n');
            if (firstNewline >= 0)
                trimmed = trimmed[(firstNewline + 1)..];
        }
        if (trimmed.EndsWith("```"))
        {
            trimmed = trimmed[..^3];
        }
        return trimmed.Trim();
    }

    /// <summary>
    /// Extracts text content from file bytes. Returns null for binary/non-text files or if content is empty.
    /// </summary>
    private static string? ExtractTextContent(byte[]? fileContent, string? fileName)
    {

        Stream abc = fileContent != null ? new MemoryStream(fileContent) : Stream.Null;
        if (fileContent == null || fileContent.Length == 0)
            return null;

        //Only extract text content for known text-based file types 
        var textExtensions = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            ".txt", ".docx", ".doc"
        };

        var extension = fileName != null ? Path.GetExtension(fileName) : null;
        if (extension != null && !textExtensions.Contains(extension))
            return $"[Binary file: {fileName}, size : {fileContent.Length}]";
        try
        {
            var text = Encoding.UTF8.GetString(fileContent);
            const int maxLength = 5000;
            if(text.Length > maxLength)
                text = text[..maxLength] + $"... [truncated, total length: {fileContent.Length}]";
            return text;
        }
        catch
        {
            return $"[Binary file : {fileName}, size : {fileContent.Length} bytes]";
        }
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
