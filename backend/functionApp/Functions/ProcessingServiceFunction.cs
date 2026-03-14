using functionApp.Models;
using functionApp.Services;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Extensions.Logging;
using System.Text.Json;
using System.Text.Json.Serialization;
using Azure.Storage.Queues;
using Azure.Data.Tables;
using functionApp.Helpers;

namespace functionApp.Functions;

public class ProcessingServiceFunction
{
    private readonly ILogger<ProcessingServiceFunction> _logger;
    private readonly NotificationRegistryService _registryService;
    private readonly QueueServiceClient _queueServiceClient;
    private readonly AppSettings _appSettings;
    private static readonly JsonSerializerOptions _jsonOptions = new()
    {
        PropertyNameCaseInsensitive = false,
        Converters = { new JsonStringEnumConverter(allowIntegerValues: false) }
    };

    public ProcessingServiceFunction(
        ILogger<ProcessingServiceFunction> logger, 
        NotificationRegistryService registryService,
        QueueServiceClient queueServiceClient,
        AppSettings appSettings)
    {
        _logger = logger;
        _registryService = registryService;
        _queueServiceClient = queueServiceClient;
        _appSettings = appSettings;
    }

    [Function("ProcessWebhookNotification")]
    public async Task<IActionResult> ProcessWebhookNotification(
        [HttpTrigger(AuthorizationLevel.Function, "get", "post", Route = null)] HttpRequest req)
    {
        _logger.LogInformation("Processing SharePoint webhook notification.");

        // SharePoint webhook validation requires echoing back the validationToken
        if (req.Query.TryGetValue("validationtoken", out var validationToken))
        {
            _logger.LogInformation("Validating SharePoint webhook subscription.");
            _logger.LogInformation("Webhook validation successful.");
            return new OkObjectResult(validationToken.ToString());
        }

        try
        {
            // Validate request has content
            if (!req.ContentLength.HasValue || req.ContentLength.Value == 0)
            {
                _logger.LogWarning("Received webhook request with no content");
                return new BadRequestObjectResult("No request content received.");
            }

            // Read and validate the complete request body
            using var reader = new StreamReader(req.Body);
            var requestBody = await reader.ReadToEndAsync();
            
            if (string.IsNullOrEmpty(requestBody))
            {
                _logger.LogWarning("Request body is empty despite Content-Length: {ContentLength}", req.ContentLength);
                return new BadRequestObjectResult("Empty request body received.");
            }

            if (requestBody.Length != req.ContentLength.Value)
            {
                _logger.LogWarning("Request body length mismatch. Expected: {Expected}, Actual: {Actual}", 
                    req.ContentLength.Value, requestBody.Length);
                return new BadRequestObjectResult("Incomplete request data received.");
            }

            _logger.LogInformation("Received webhook payload: {Body}", requestBody);

            // Deserialize the notification data
            var webhookData = JsonSerializer.Deserialize<WebhookNotificationData>(requestBody, _jsonOptions);

            if (webhookData?.Value == null || !webhookData.Value.Any())
            {
                return new BadRequestObjectResult("No notification data received.");
            }

            _logger.LogInformation("Processing {Count} webhook notifications.", webhookData.Value.Length);

            // Process each notification
            // SharePoint may batch multiple notifications together, so we need to handle each one
            foreach (var notification in webhookData.Value)
            {
                _logger.LogInformation("Processing webhook notifications {Count}.", notification);
                await ProcessSingleNotification(notification);
            }

            return new OkResult();
        }
        catch (JsonException ex)
        {
            _logger.LogError(ex, "Failed to deserialize webhook notification data.");
            return new BadRequestObjectResult("Invalid JSON payload.");
        }
        catch (BadHttpRequestException ex) when (ex.Message.Contains("Unexpected end of request content"))
        {
            _logger.LogWarning(ex, "Request content was truncated - possible network issue or malformed request from SharePoint. Content-Length: {ContentLength}", req.ContentLength);
            return new BadRequestObjectResult("Request content was incomplete - possible network connectivity issue.");
        }
        catch (IOException ex)
        {
            _logger.LogError(ex, "IO error while reading request body. Content-Length: {ContentLength}", req.ContentLength);
            return new BadRequestObjectResult("Error reading request data.");
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error processing webhook notification.");
            return new StatusCodeResult(StatusCodes.Status500InternalServerError);
        }
    }

    /// <summary>
    /// Processes a single webhook notification and correlates it with registered subscriptions.
    /// </summary>
    /// <remarks>This method logs the processing steps and handles exceptions that may occur during the
    /// notification processing. It retrieves site information and finds matching registrations before queuing a
    /// notification message.</remarks>
    /// <param name="notification">The webhook notification model containing details about the resource and subscription to process.</param>
    /// <returns>A task representing the asynchronous operation of processing the notification.</returns>
    private async Task ProcessSingleNotification(WebhookNotificationModel notification)
    {
        _logger.LogInformation("Processing notification for resource: {Resource}", notification.Resource);

        try
        {
            // Construct the full site URL
            var siteUrlBuilder = new UriBuilder($"https://{_appSettings.SharePointTenantName}")
            {
                Path = notification.SiteUrl
            };

            // Get site information to correlate with registrations
            var context = _appSettings.GetContext(siteUrlBuilder.ToString(), _logger);

            // Load Site and it's ID
            context.Load(context.Site, s => s.Id);
            await context.ExecuteQueryAsync();

            // Query table storage to find matching registrations
            var registrations = await FindMatchingRegistrations(
                context.Site.Id,
                notification.WebId,
                notification.Resource
                );
            
            _logger.LogInformation($"Found matching registration for subscription: {notification.SubscriptionId}");

            // Create queue message for matching registrations
            await QueueNotificationMessage(registrations, notification);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error processing single notification for resource: {Resource}", notification.Resource);
            throw;
        }
    }

    /// <summary>
    /// Retrieves notification registrations that match the specified site, web, and list identifiers.
    /// </summary>
    /// <remarks>This method queries registration data from table storage, filtering by the provided siteId,
    /// webId, and listId.</remarks>
    /// <param name="siteId">The unique identifier of the site used to filter notification registrations.</param>
    /// <param name="webId">The unique identifier of the web associated with the site.</param>
    /// <param name="listId">The unique identifier of the list for which to find matching notification registrations.</param>
    /// <returns>The result contains a list of notification registrations matching the provided identifiers.</returns>
    private async Task<List<NotificationRegistration>> FindMatchingRegistrations(Guid siteId, string webId, string listId)
    {
        //TODO : Webhook Notification doesnt have SiteID Instead we filter out with webID and listID. We can remove siteID from registration in future if not required.
        _logger.LogInformation($"Finding matching registrations for site \"{siteId}\", web \"{webId}\", list \"{listId}\".");

        var allRegistrations = new List<NotificationRegistration>();

        // Get all registrations from table storage
        // We filter registrations based on siteId, webId, and listId.
        var registrations = _registryService.GetTableClient()
                                    .QueryAsync<NotificationRegistrationEntity>()
                                    .Where(n => 
                                        //n.SiteId == siteId &&
                                        n.WebId.ToString().ToLowerInvariant() == webId.ToLowerInvariant() &&
                                        n.ListId.ToString().ToLowerInvariant() == listId.ToLowerInvariant()
                                    )
                                    .Select(r => r.ToModel());

        return await registrations.ToListAsync();
    }

    /// <summary>
    /// Queues a notification message for processing by serializing the provided registrations and notification model
    /// and sending it to the designated queue.
    /// </summary>
    /// <remarks>This method ensures that the target queue exists before sending the message. It logs
    /// information about the queuing process and propagates exceptions that occur during the operation.</remarks>
    /// <param name="registrations">A list of notification registrations specifying the recipients and their preferences for receiving
    /// notifications. Cannot be null.</param>
    /// <param name="notification">The webhook notification model coming from SharePoint and containing the details of the notification to be sent, 
    /// including the subscription ID and relevant data. Cannot be null.</param>
    /// <returns>A task that represents the asynchronous operation of queuing the notification message.</returns>
    private async Task QueueNotificationMessage(List<NotificationRegistration> registrations, WebhookNotificationModel notification)
    {
        try
        {
            var queueClient = _queueServiceClient.GetQueueClient(_appSettings.NotificationQueueName);
            await queueClient.CreateIfNotExistsAsync();

            var queueMessage = new NotificationQueueMessage
            {
                Registrations = registrations,
                WebhookNotification = notification,
                QueuedAt = DateTime.UtcNow
            };

            // Serialize the message to JSON and then encode it as a Base64 string for queue storage
            var messageJson = JsonSerializer.Serialize(queueMessage, _jsonOptions);
            byte[] bytes = System.Text.Encoding.UTF8.GetBytes(messageJson);
            string base64String = Convert.ToBase64String(bytes);
            await queueClient.SendMessageAsync(base64String);

            _logger.LogInformation($"Queued notification message for subscription {notification.SubscriptionId}");
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, $"Failed to queue notification message for subscription {notification.SubscriptionId}");
            throw;
        }
    }
}