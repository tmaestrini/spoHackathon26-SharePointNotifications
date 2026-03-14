using Azure.Data.Tables;
using Azure.Storage.Queues;
using functionApp.Helpers;
using functionApp.Models;
using functionApp.Services;
using Google.Protobuf;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using Microsoft.Graph.Chats.Item.PermissionGrants.Item;
using Microsoft.Graph.Models;
using System.Runtime.InteropServices;
using System.Text.Json;
using System.Text.Json.Serialization;
using models = functionApp.Models;

namespace functionApp.Functions;

public class NotifierServiceFunction
{
    private readonly ILogger<NotifierServiceFunction> _logger;
    private readonly NotificationRegistryService _registryService;
    private readonly QueueServiceClient _queueServiceClient;
    private readonly AppSettings _appSettings;
    private readonly DeltaService _deltaService;
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
        DeltaService deltaService,
        AINotificationService aiNotificationService)
    {
        _logger = logger;
        _registryService = registryService;
        _queueServiceClient = queueServiceClient;
        _appSettings = appSettings;
        _deltaService = deltaService;
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

            if (notificationMessage.Registrations == null || notificationMessage.Registrations.Count == 0)
            {
                _logger.LogInformation("No registration found for webhook notification");
                return;
            }

            _logger.LogInformation("Processing notification for {RegistrationCount} registrations",
                notificationMessage.Registrations?.Count ?? 0);

            // Retrieve the delta for the webhook subscription
            var delta = await _deltaService.GetDeltaForNotificationAsync(notificationMessage.WebhookNotification);

            if (delta.Count == 0)
            {
                _logger.LogInformation("No changes found in delta for webhook notification");
                return;
            }

            var created = delta.Where(d => d.ChangeType == DeltaChangeType.Created);
            var updated = delta.Where(d => d.ChangeType == DeltaChangeType.Updated);
            var deleted = delta.Where(d => d.ChangeType == DeltaChangeType.Deleted);

            foreach (var registrationGroup in notificationMessage.Registrations!.GroupBy(r => r.ChangeType))
            {
                var itemsToNotify = new List<DeltaItemChange>();

                if (registrationGroup.Key == models.ChangeType.CREATED || registrationGroup.Key == models.ChangeType.ALL)
                    itemsToNotify.AddRange(created);

                if (registrationGroup.Key == models.ChangeType.UPDATED || registrationGroup.Key == models.ChangeType.ALL)
                    itemsToNotify.AddRange(updated);

                if (registrationGroup.Key == models.ChangeType.DELETED || registrationGroup.Key == models.ChangeType.ALL)
                    itemsToNotify.AddRange(deleted);

                foreach (var registration in registrationGroup)
                {
                    var notificationText = await _aiNotificationService.ProcessNotificationAsync(itemsToNotify);

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
            await SendNotificationAsync(registration, notificationText, channel);
        }
    }

    private async Task SendNotificationAsync(NotificationRegistration registration, string notificationText, NotificationChannel notificationChannel)
    {
        _logger.LogInformation("Sending notification for registration {RegistrationId} on channel {Channel}", registration.Id, notificationChannel);

        try
        {
            if (string.IsNullOrEmpty(_appSettings.NotificationFlowUrl))
            {
                _logger.LogWarning("NotificationFlowUrl is not configured in app settings");
                return;
            }

            // Retrieve the user to get their email for the Teams notification
            var appGraphClient = ConnectionHelper.GraphClient(_appSettings, _logger);
            var userId = registration.UserId.ToString();
            var user = await appGraphClient.Users[userId].GetAsync();

            if (user == null)
            {
                _logger.LogWarning("User with ID {UserId} not found for sending {Channel} notification", registration.UserId, notificationChannel);
                return;
            }

            using var httpClient = new HttpClient();

            var requestPayload = new
            {
                userPrincipalName = user.UserPrincipalName,
                notificationText,
                notificationType = notificationChannel.ToString()
            };

            var jsonContent = JsonSerializer.Serialize(requestPayload, _jsonOptions);
            var content = new StringContent(jsonContent, System.Text.Encoding.UTF8, "application/json");

            var response = await httpClient.PostAsync(_appSettings.NotificationFlowUrl, content);

            if (response.IsSuccessStatusCode)
            {
                _logger.LogInformation("Notification sent successfully for registration {RegistrationId} on channel {Channel}", registration.Id, notificationChannel);
            }
            else
            {
                _logger.LogWarning("Notification failed with status code {StatusCode} for registration {RegistrationId} on channel {Channel}", 
                    response.StatusCode, registration.Id, notificationChannel);
            }
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Failed to send Teams notification for registration {RegistrationId}", registration.Id);
            throw;
        }
    }
}