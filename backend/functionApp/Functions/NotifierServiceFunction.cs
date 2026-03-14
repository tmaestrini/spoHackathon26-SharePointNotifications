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
        _logger.LogInformation("Sending Teams notification for registration {RegistrationId}", registration.Id);

        // TODO: Implement Teams notification logic using Microsoft Graph API to send messages to user.
    }

    private async Task SendEmailNotificationAsync(NotificationRegistration registration, string notificationText)
    {
        _logger.LogInformation("Sending email notification for registration {RegistrationId}", registration.Id);

        try
        {
            // Use service user Graph client for sending emails
            var appGraphClient = ConnectionHelper.GraphClient(_appSettings, _logger);

            // Use service user Graph client for sending emails
            var graphClient = ConnectionHelper.GraphClientForServiceUser(_appSettings, _logger);

            var userId = registration.UserId.ToString();

            // Get user by userId
            var user = await appGraphClient.Users[userId].GetAsync();

            if (user == null)
            {
                _logger.LogWarning("User with ID {UserId} not found for email notification", registration.UserId);
                return;
            }

            if (string.IsNullOrEmpty(user.Mail))
            {
                _logger.LogWarning("User with ID {UserId} does not have an email address configured", registration.UserId);
                return;
            }

            var message = new Message
            {
                Subject = "Something changed!",
                Body = new ItemBody
                {
                    ContentType = BodyType.Text,
                    Content = notificationText
                },
                ToRecipients = new List<Recipient>()
                {
                    new Recipient
                    {
                        EmailAddress = new EmailAddress
                        {
                            Address = user.Mail
                        }
                    }
                }
            };

            // Create the send mail request body
            var sendMailRequest = new Microsoft.Graph.Users.Item.SendMail.SendMailPostRequestBody
            {
                Message = message,
                SaveToSentItems = true
            };

            await graphClient.Users[_appSettings.NotificationServiceUserName]
                .SendMail
                .PostAsync(sendMailRequest);

            _logger.LogInformation("Email notification sent successfully to {EmailAddress} for registration {RegistrationId}", 
                user.Mail, registration.Id);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Failed to send email notification for registration {RegistrationId}", registration.Id);
            throw;
        }
    }
}