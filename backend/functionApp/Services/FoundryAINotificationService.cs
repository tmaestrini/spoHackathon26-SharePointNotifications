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
                CurrentFileContent = DocumentTextExtractor.ExtractTextContent(i.CurrentFileContent, i.CurrentFileName),
                PreviousFileContent = DocumentTextExtractor.ExtractTextContent(i.PreviousFileContent, i.CurrentFileName),
                PreviousFileVersionLabel = i.PreviousFileVersionLabel
            });

            var systemMessage = "You are a SharePoint notification assistant. Your job is to summarize list or library changes into clear, concise, human-readable notification messages. Focus on what changed (created, updated, deleted), which items were affected, and any relevant field values. When version information and field-level changes are available, highlight what specifically changed between the current and previous version. For documents, compare the current and previous file content and describe what was added, removed, or modified. Mention the file name and version labels. Keep the summary brief and actionable.";

            var userMessage = $"""
                The user has subscribed to notifications with the following description:
                "{registration.Description ?? "All changes"}"

                Change type filter: {registration.ChangeType}

                Here are the detected changes in the SharePoint list/library. If it is a file, then compare the CurrentFileContent and PreviousFileContent to see the difference in changes:
                {JsonSerializer.Serialize(changeSummary, new JsonSerializerOptions { WriteIndented = true })}

                Please provide a concise, human-readable notification summary of these changes in the below similar format or in a much better way. 
                
                If its an update, The summary should be in below format and should only include the fields that were changed:
                
                'SharePoint Notification

                ChangeType : Updated
                Document Name: Updated Document Name or List Item ID
                Version: Updated from version to version if available, otherwise just mention the current version label
                Last Modified By: ABCDEF
                Modification Date: March 14, 2026

                Content Changes

                Added or Modified Text

                "Hi abc, I have completed the Document changes as well, This is the Text which is returned."

                Metadata Changes
                Field1: Previous Value -> New Value

                No other field-level changes were detected in the document metadata.'

                If it's a creation, summarize the changes in a similar concise format, mentioning the key details of the created, such as the item type (document, list item), name, and any relevant field values.

                'SharePoint Notification
                
                ChangeType : Created
                Document Name: Document Name or List Item ID
                Version: if available, mention the version label for created item
                Last Modified By: ABCDEF
                Modification Date: March 14, 2026
                
                Content : 
                Summarize the content of the document here. If it's a non-text file, just mention the file name and type.
                
                Metadata:
                Specify the key metadata values for the created item. For example:
                Field1: Value1'

                If its a deletion, the summary should be like below format:
                'SharePoint Notification

                ChangeType : Deleted
                Document Name: if its a document mention the document name, if its a list item then mention the item id.
                Version: if available, mention the last version label before deletion
                When it was deleted: if available, mention the deletion date and time
                Who deleted it: if available, mention the user who performed the deletion'

                """;

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
