using functionApp.Models;
using functionApp.Services;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Extensions.Logging;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace functionApp.Functions;

public class API_Func
{
    private readonly ILogger<API_Func> _logger;
    private readonly NotificationRegistryService _registryService;
    private static readonly JsonSerializerOptions _jsonOptions = new()
    {
        PropertyNameCaseInsensitive = true,
        Converters = { new JsonStringEnumConverter() }
    };

    public API_Func(ILogger<API_Func> logger, NotificationRegistryService registryService)
    {
        _logger = logger;
        _registryService = registryService;
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

        var deleted = await _registryService.DeleteAsync(userId, id);
        return deleted ? new NoContentResult() : new NotFoundResult();
    }
}