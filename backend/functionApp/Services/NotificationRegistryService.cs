using Azure.Data.Tables;
using functionApp.Models;
using Microsoft.Extensions.Logging;

namespace functionApp.Services;

public class NotificationRegistryService
{
    private readonly TableClient _tableClient;
    private readonly ILogger<NotificationRegistryService> _logger;

    public NotificationRegistryService(TableServiceClient tableServiceClient, AppSettings env, ILogger<NotificationRegistryService> logger)
    {
        _logger = logger;
        _tableClient = tableServiceClient.GetTableClient(env.TableNotificationRegistrations);
        _tableClient.CreateIfNotExists();
        _logger.LogInformation("NotificationRegistryService initialized with table '{TableName}'.", env.TableNotificationRegistrations);
    }

    public async Task<NotificationRegistration> CreateAsync(NotificationRegistration registration)
    {
        if (registration.Id == Guid.Empty)
            registration.Id = Guid.NewGuid();

        _logger.LogInformation("Creating registration {Id} for user {UserId}.", registration.Id, registration.UserId);
        var entity = NotificationRegistrationEntity.FromModel(registration);
        await _tableClient.AddEntityAsync(entity);
        _logger.LogInformation("Registration {Id} created successfully.", registration.Id);
        return registration;
    }

    public async Task<NotificationRegistration?> GetAsync(Guid userId, Guid registrationId)
    {
        _logger.LogInformation("Fetching registration {Id} for user {UserId}.", registrationId, userId);
        try
        {
            var response = await _tableClient.GetEntityAsync<NotificationRegistrationEntity>(
                userId.ToString(), registrationId.ToString());
            return response.Value.ToModel();
        }
        catch (Azure.RequestFailedException ex) when (ex.Status == 404)
        {
            _logger.LogWarning("Registration {Id} not found for user {UserId}.", registrationId, userId);
            return null;
        }
    }

    public async Task<List<NotificationRegistration>> GetByUserAsync(Guid userId)
    {
        _logger.LogInformation("Fetching all registrations for user {UserId}.", userId);
        var results = new List<NotificationRegistration>();
        await foreach (var entity in _tableClient.QueryAsync<NotificationRegistrationEntity>(
            e => e.PartitionKey == userId.ToString()))
        {
            results.Add(entity.ToModel());
        }
        _logger.LogInformation("Found {Count} registrations for user {UserId}.", results.Count, userId);
        return results;
    }

    public async Task<NotificationRegistration?> UpdateAsync(NotificationRegistration registration)
    {
        _logger.LogInformation("Updating registration {Id} for user {UserId}.", registration.Id, registration.UserId);
        try
        {
            var entity = NotificationRegistrationEntity.FromModel(registration);
            await _tableClient.UpsertEntityAsync(entity, TableUpdateMode.Replace);
            _logger.LogInformation("Registration {Id} updated successfully.", registration.Id);
            return registration;
        }
        catch (Azure.RequestFailedException ex) when (ex.Status == 404)
        {
            _logger.LogWarning("Registration {Id} not found for update.", registration.Id);
            return null;
        }
    }

    public async Task<bool> DeleteAsync(Guid userId, Guid registrationId)
    {
        _logger.LogInformation("Deleting registration {Id} for user {UserId}.", registrationId, userId);
        try
        {
            await _tableClient.DeleteEntityAsync(userId.ToString(), registrationId.ToString());
            _logger.LogInformation("Registration {Id} deleted successfully.", registrationId);
            return true;
        }
        catch (Azure.RequestFailedException ex) when (ex.Status == 404)
        {
            _logger.LogWarning("Registration {Id} not found for deletion.", registrationId);
            return false;
        }
    }
}
