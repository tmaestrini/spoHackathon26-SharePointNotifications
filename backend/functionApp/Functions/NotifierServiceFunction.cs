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

public class NotifierServiceFunction
{
    private readonly ILogger<NotifierServiceFunction> _logger;
    private readonly NotificationRegistryService _registryService;
    private readonly QueueServiceClient _queueServiceClient;
    private readonly AppSettings _appSettings;
    private readonly DataService _dataService;
    private readonly AINotificationService _aiNotificationService;
    private static readonly JsonSerializerOptions _jsonOptions = new()
    {
        PropertyNameCaseInsensitive = true,
        Converters = { new JsonStringEnumConverter() }
    };

    public NotifierServiceFunction(
        ILogger<NotifierServiceFunction> logger, 
        NotificationRegistryService registryService,
        QueueServiceClient queueServiceClient,
        AppSettings appSettings,
        DataService dataService,
        AINotificationService aiNotificationService)
    {
        _logger = logger;
        _registryService = registryService;
        _queueServiceClient = queueServiceClient;
        _appSettings = appSettings;
        _dataService = dataService;
        _aiNotificationService = aiNotificationService;
    }

    [Function(nameof(ProcessNotificationQueue))]
    public async Task ProcessNotificationQueue(
        [QueueTrigger("%NotificationQueueName%")] string queueMessage)
    {
        try
        {
            _logger.LogInformation("Processing notification queue message: {Message}", queueMessage);

            var notificationMessage = JsonSerializer.Deserialize<NotificationQueueMessage>(queueMessage, _jsonOptions);

            if (notificationMessage == null)
            {
                _logger.LogWarning("Failed to deserialize queue message: {Message}", queueMessage);
                return;
            }

            _logger.LogInformation("Processing notification for {RegistrationCount} registrations", 
                notificationMessage.Registrations?.Count ?? 0);

            // Retrieve the delta for the webhook subscription
            var delta = await _dataService.GetDeltaAsync(notificationMessage.WebhookNotification);

            // TODO: Here you can add logic to:
            // 1. Filter changes based on registration.ChangeType (CREATED, UPDATED, DELETED, ALL)
            // 2. Use the AINotificationService to generate AI-based summaries or insights about the change
            //    _aiNotificationService.ProcessNotification();
            // 3. Send notifications through the specified channels (TEAMS, EMAIL)

            _logger.LogInformation("Successfully processed notification queue message");
        }
        catch (JsonException ex)
        {
            _logger.LogError(ex, "Failed to deserialize queue message: {Message}", queueMessage);
            throw; // Re-throw to move message to poison queue
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error processing notification queue message");
            throw; // Re-throw to move message to poison queue
        }
    }
}