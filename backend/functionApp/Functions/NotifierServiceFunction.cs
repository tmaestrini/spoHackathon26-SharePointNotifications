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
using System.Runtime.InteropServices;
using Microsoft.Graph.Chats.Item.PermissionGrants.Item;

namespace functionApp.Functions;

public class NotifierServiceFunction
{
    private readonly ILogger<NotifierServiceFunction> _logger;
    private readonly NotificationRegistryService _registryService;
    private readonly QueueServiceClient _queueServiceClient;
    private readonly AppSettings _appSettings;
    private readonly DeltaService _deltaService;
    private readonly AINotificationService _aiNotificationService;
    private readonly FoundryAINotificationService _foundryAINotificationService;
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
        DeltaService deltaService,
        AINotificationService aiNotificationService,
        FoundryAINotificationService foundryAINotificationService)
    {
        _logger = logger;
        _registryService = registryService;
        _queueServiceClient = queueServiceClient;
        _appSettings = appSettings;
        _deltaService = deltaService;
        _aiNotificationService = aiNotificationService;
        _foundryAINotificationService = foundryAINotificationService;
    }

    [Function(nameof(ProcessNotificationQueue))]
    public async Task ProcessNotificationQueue(
        [QueueTrigger("notifications")] string queueMessage)
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

            if (notificationMessage.Registrations == null || notificationMessage.Registrations.Count == 0)
            {
                _logger.LogInformation("No registration found for webhook notification");
                return;
            }

            _logger.LogInformation("Processing notification for {RegistrationCount} registrations",
                notificationMessage.Registrations?.Count ?? 0);

            // Retrieve the delta for the webhook subscription
            var delta = await _deltaService.GetDeltaForNotificationAsync(notificationMessage.WebhookNotification);

            var created = delta.Where(d => d.ChangeType == DeltaChangeType.Created);
            var updated = delta.Where(d => d.ChangeType == DeltaChangeType.Updated);
            var deleted = delta.Where(d => d.ChangeType == DeltaChangeType.Deleted);

            foreach (var registrationGroup in notificationMessage.Registrations!.GroupBy(r => r.ChangeType))
            {
                var itemsToNotify = new List<DeltaItemChange>();

                if (registrationGroup.Key == ChangeType.CREATED || registrationGroup.Key == ChangeType.ALL)
                    itemsToNotify.AddRange(created);

                if (registrationGroup.Key == ChangeType.UPDATED || registrationGroup.Key == ChangeType.ALL)
                    itemsToNotify.AddRange(updated);

                if (registrationGroup.Key == ChangeType.DELETED || registrationGroup.Key == ChangeType.ALL)
                    itemsToNotify.AddRange(deleted);

                foreach (var registration in registrationGroup)
                {
                    var notificationText = await _foundryAINotificationService.ProcessNotificationAsync(itemsToNotify, registration);

                    await SendNotificationAsync(registration, notificationText);
                }
            }

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

    private async Task SendNotificationAsync(NotificationRegistration registration, string notificationText)
    {
        if (registration.NotificationChannels == null || registration.NotificationChannels.Count() == 0)
        {
            _logger.LogWarning("No notification channels configured for registration {RegistrationId}", registration.Id);
            return;
        }

        foreach (var channel in registration.NotificationChannels)
        {
            switch (channel)
            {
                case NotificationChannel.TEAMS:
                    await SendTeamsNotificationAsync(registration, notificationText);
                    break;
                case NotificationChannel.EMAIL:
                    await SendEmailNotificationAsync(registration, notificationText);
                    break;
                default:
                    _logger.LogWarning("Unsupported notification channel {Channel} for registration {RegistrationId}", channel, registration.Id);
                    break;
            }
        }
    }

    private async Task SendTeamsNotificationAsync(NotificationRegistration registration, string notificationText)
    {
        // TODO: send notification to Microsoft Teams using Graph API
    }

    private async Task SendEmailNotificationAsync(NotificationRegistration registration, string notificationText)
    {
        // TODO: send notification email using Microsoft Graph API
    }
}