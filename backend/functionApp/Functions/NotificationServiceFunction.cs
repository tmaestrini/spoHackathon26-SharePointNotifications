using functionApp.Models;
using functionApp.Services;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Extensions.Logging;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace functionApp.Functions;

public class NotificationServiceFunction
{
    private readonly ILogger<NotificationServiceFunction> _logger;
    private readonly NotificationRegistryService _registryService;
    private readonly WebhookService _webhookService;
    private static readonly JsonSerializerOptions _jsonOptions = new()
    {
        PropertyNameCaseInsensitive = true,
        Converters = { new JsonStringEnumConverter() }
    };

    public NotificationServiceFunction(ILogger<NotificationServiceFunction> logger, NotificationRegistryService registryService, WebhookService webhookService)
    {
        _logger = logger;
        _registryService = registryService;
        _webhookService = webhookService;
    }

    [Function("CreateRegistration")]
    public async Task<IActionResult> CreateRegistration(
        [HttpTrigger(AuthorizationLevel.Function, "post", Route = "registrations")] HttpRequest req)
    {
        _logger.LogInformation("Creating notification registration.");

        var registration = await JsonSerializer.DeserializeAsync<NotificationRegistration>(
            req.Body, _jsonOptions);

        if (registration is null)
            return new BadRequestObjectResult("Invalid registration payload.");

        if (registration.UserId == Guid.Empty)
            return new BadRequestObjectResult("UserId is required.");

        if (registration.NotificationChannels.Length == 0)
            return new BadRequestObjectResult("At least one NotificationChannel is required.");

        var created = await _registryService.CreateAsync(registration);

        // Register webhook on the SharePoint list/library
        try
        {
            if (!string.IsNullOrEmpty(registration.SiteUrl))
            {
                await _webhookService.RegisterWebhookAsync(registration.SiteUrl, registration.ListId);
            }
            else
            {
                _logger.LogWarning("SiteUrl not provided in registration {Id}. Webhook was not registered.", created.Id);
            }
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Failed to register webhook for registration {Id}. The registration was saved but the webhook was not created.", created.Id);
            // Registration is saved; webhook failure is logged but doesn't block the response.
        }

        return new CreatedResult($"/api/registrations/{created.UserId}/{created.Id}", created);
    }

    [Function("GetRegistration")]
    public async Task<IActionResult> GetRegistration(
        [HttpTrigger(AuthorizationLevel.Function, "get", Route = "registrations/{userId}/{id}")] HttpRequest req,
        Guid userId, Guid id)
    {
        _logger.LogInformation("Getting registration {Id} for user {UserId}.", id, userId);

        var registration = await _registryService.GetAsync(userId, id);
        return registration is not null
            ? new OkObjectResult(registration)
            : new NotFoundResult();
    }

    [Function("GetUserRegistrations")]
    public async Task<IActionResult> GetUserRegistrations(
        [HttpTrigger(AuthorizationLevel.Function, "get", Route = "registrations/{userId}")] HttpRequest req,
        Guid userId)
    {
        _logger.LogInformation("Getting all registrations for user {UserId}.", userId);

        var registrations = await _registryService.GetByUserAsync(userId);
        return new OkObjectResult(registrations);
    }

    [Function("UpdateRegistration")]
    public async Task<IActionResult> UpdateRegistration(
        [HttpTrigger(AuthorizationLevel.Function, "put", Route = "registrations/{userId}/{id}")] HttpRequest req,
        Guid userId, Guid id)
    {
        _logger.LogInformation("Updating registration {Id} for user {UserId}.", id, userId);

        var registration = await JsonSerializer.DeserializeAsync<NotificationRegistration>(
            req.Body, _jsonOptions);

        if (registration is null)
            return new BadRequestObjectResult("Invalid registration payload.");

        registration.Id = id;
        registration.UserId = userId;

        var updated = await _registryService.UpdateAsync(registration);
        return updated is not null
            ? new OkObjectResult(updated)
            : new NotFoundResult();
    }

    [Function("DeleteRegistration")]
    public async Task<IActionResult> DeleteRegistration(
        [HttpTrigger(AuthorizationLevel.Function, "delete", Route = "registrations/{userId}/{id}")] HttpRequest req,
        Guid userId, Guid id)
    {
        _logger.LogInformation("Deleting registration {Id} for user {UserId}.", id, userId);

        // Fetch the registration first so we can remove the webhook if needed
        var registration = await _registryService.GetAsync(userId, id);
        if (registration is null)
            return new NotFoundResult();

        var deleted = await _registryService.DeleteAsync(userId, id);
        if (!deleted)
            return new NotFoundResult();

        // Check if other registrations still exist for the same list
        // Only remove the webhook if this was the last registration for that site+list
        try
        {
            if (!string.IsNullOrEmpty(registration.SiteUrl))
            {
                var remaining = await _registryService.GetByListAsync(registration.SiteId, registration.WebId, registration.ListId);
                if (remaining.Count == 0)
                {
                    _logger.LogInformation("No remaining registrations for list {ListId}. Removing webhook.", registration.ListId);
                    await _webhookService.RemoveWebhookAsync(registration.SiteUrl, registration.ListId);
                }
                else
                {
                    _logger.LogInformation("{Count} registrations still exist for list {ListId}. Keeping webhook.", remaining.Count, registration.ListId);
                }
            }
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Failed to remove webhook for registration {Id}. The registration was deleted but the webhook may still be active.", id);
        }

        return new NoContentResult();
    }
}