# SharePoint Notifications Backend

Azure Functions backend for managing SharePoint notification registrations. Users can subscribe to changes on specific SharePoint lists/libraries and receive notifications via Teams or Email. The backend automatically registers and manages SharePoint webhook subscriptions, including renewal.

## Tech Stack

- **.NET 10** with **Azure Functions v4** (isolated worker model)
- **Azure Table Storage** for persisting notification registrations and webhook subscriptions
- **Azure Queue Storage** for asynchronous notification processing
- **Azure Key Vault** for certificate-based authentication
- **Microsoft Graph SDK** for Microsoft 365 integration and list item delta queries
- **PnP Framework** for SharePoint connectivity and webhook management
- **Application Insights** for telemetry and monitoring

## Project Structure

```
functionApp/
├── Program.cs                              # Host setup & dependency injection
├── host.json                               # Azure Functions host configuration
├── local.settings.json                     # Local development settings
├── Functions/
│   ├── NotificationServiceFunction.cs      # CRUD API for notification registrations + webhook lifecycle
│   ├── ProcessingServiceFunction.cs        # SharePoint webhook receiver & notification dispatcher
│   ├── NotifierServiceFunction.cs          # Queue-triggered notification processor
│   └── WebhookRenewalServiceFunction.cs    # Timer-triggered webhook renewal
├── Helpers/
│   └── ConnectionHelper.cs                 # SharePoint & Graph authentication helpers
├── Models/
│   ├── AppSettings.cs                      # Environment-based configuration model
│   ├── NotificationRegistration.cs         # Core registration domain model
│   ├── NotificationRegistrationEntity.cs   # Azure Table Storage entity mapping for registrations
│   ├── ProcessingModels.cs                 # Queue message & webhook notification wrapper models
│   ├── WebhookNotificationModel.cs         # SharePoint webhook payload model
│   └── WebhookSubscriptionEntity.cs        # Azure Table Storage entity for webhook subscriptions
└── Services/
    ├── AINotificationService.cs            # AI-powered notification processing (placeholder)
    ├── DataService.cs                      # Microsoft Graph delta queries for list item changes
    ├── NotificationRegistryService.cs      # CRUD operations for notification registrations
    ├── WebhookService.cs                   # SharePoint webhook register, remove & renew operations
    └── WebhookSubscriptionService.cs       # CRUD operations for webhook subscription records
```

## Implemented Functionalities

### 1. Notification Registration CRUD API

Full REST API for managing notification registrations (`NotificationServiceFunction`):

| Method | Route | Function | Description |
|--------|-------|----------|-------------|
| `POST` | `/api/registrations` | `CreateRegistration` | Create a new notification registration and register the webhook on the SharePoint list |
| `GET` | `/api/registrations/{userId}` | `GetUserRegistrations` | Get all registrations for a user |
| `GET` | `/api/registrations/{userId}/{id}` | `GetRegistration` | Get a specific registration by ID |
| `PUT` | `/api/registrations/{userId}/{id}` | `UpdateRegistration` | Update an existing registration |
| `DELETE` | `/api/registrations/{userId}/{id}` | `DeleteRegistration` | Delete a registration and remove the webhook if no other registrations exist for the same list |

Request/response payloads use the `NotificationRegistration` model:

- **Id** – Unique registration identifier (auto-generated on create)
- **UserId** – The subscribing user's ID
- **ChangeType** – `CREATED`, `UPDATED`, `DELETED`, or `ALL`
- **SiteId / WebId / ListId** – SharePoint resource identifiers (GUIDs)
- **SiteUrl** – Full SharePoint site URL (e.g. `https://tenant.sharepoint.com/sites/MySite`)
- **ItemId** – Optional specific item to watch
- **NotificationChannels** – Array of channels: `TEAMS` and/or `EMAIL`
- **Description** – Optional description of the registration

Input validation: `UserId` must be non-empty and at least one `NotificationChannel` is required.

### 2. SharePoint Webhook Lifecycle Management

Webhooks are automatically managed through the `WebhookService`:

- **Registration** – When a notification registration is created, a SharePoint webhook is registered on the specified list/library via PnP Framework's `AddWebhookSubscription` (6-month expiration). The subscription is persisted to the `WebhookSubscriptions` Azure Table.
- **Removal** – When the last notification registration for a given list is deleted, the corresponding webhook subscription is removed from SharePoint and the table record is cleaned up.
- **Renewal** – A timer-triggered function (`MonthlyRenewWebhooks`) runs on the 1st of every month at 5:00 AM UTC. It queries for all subscriptions expiring within 45 days and renews them via `UpdateWebhookSubscription`, extending the expiration by 6 months. Stale subscriptions (no longer present on SharePoint) are automatically cleaned up.

### 3. SharePoint Webhook Notification Processing

The webhook processing pipeline (`ProcessingServiceFunction`):

1. **Validation** – Handles SharePoint's `validationtoken` handshake when webhooks are being registered
2. **Notification Receiving** – `POST /api/webhook/notification` receives webhook notifications from SharePoint (supports batched notifications)
3. **Registration Correlation** – Each notification is correlated with matching registrations in table storage by `SiteId`, `WebId`, and `ListId`
4. **Queue Dispatch** – Matched notifications are serialized and sent to the `notifications` Azure Queue (Base64-encoded)

### 4. Queue-Based Notification Processing

The `NotifierServiceFunction` processes messages from the notification queue:

- Deserializes the `NotificationQueueMessage` containing matched registrations and the webhook notification
- Retrieves list item deltas via Microsoft Graph (`DataService.GetDeltaAsync`)
- Placeholder for: filtering by `ChangeType`, AI-based summarization via `AINotificationService`, and sending notifications through Teams/Email channels

### 5. Azure Table Storage Persistence

Two tables are used:

- **`NotificationRegistrations`** – Stores user notification registrations
  - **PartitionKey** = `UserId`, **RowKey** = Registration `Id`
  - Complex fields (`NotificationChannels`) serialized as JSON
  - Auto-created on startup

- **`WebhookSubscriptions`** – Stores active SharePoint webhook subscriptions
  - **PartitionKey** = `"Webhooks"`, **RowKey** = SharePoint Subscription `Id`
  - Tracks `ListId`, `SiteUrl`, `NotificationUrl`, `Resource`, `ExpirationDateTime`, `ChangeToken`, and `LastUpdated`
  - Auto-created on startup

### 6. SharePoint & Microsoft Graph Authentication

The `ConnectionHelper` provides authenticated connections using certificate-based app-only authentication:

- **Certificate retrieval from Azure Key Vault** using `DefaultAzureCredential`
- **Certificate caching** with a 1-hour expiration window (thread-safe)
- **SharePoint CSOM context** via PnP Framework `AuthenticationManager`
- **App-only access token** generation for SharePoint REST calls
- **Microsoft Graph client** with `ClientCertificateCredential`
- Local development supports loading certificates from the current user certificate store by thumbprint

### 7. Microsoft Graph Delta Queries

The `DataService` retrieves list item changes via Microsoft Graph's delta endpoint:

- Resolves the site by tenant name and site URL
- Queries the list items delta endpoint with pagination support (`PageIterator`)
- Returns a list of changed `ListItem` objects for downstream processing

### 8. Configuration Management

`AppSettings` auto-populates from environment variables using reflection:

| Setting | Description |
|---------|-------------|
| `AADAppId` | Entra ID (Azure AD) application ID |
| `AADAppSecret` | Entra ID application secret |
| `TenantId` | Azure tenant ID |
| `VaultUri` | Azure Key Vault URI |
| `VaultCertName` | Certificate name in Key Vault |
| `CertThumbprint` | Certificate thumbprint (local development) |
| `AzureWebJobsStorage` | Azure Storage connection string |
| `TableNotificationRegistrations` | Table name for notification registrations |
| `TableWebhookSubscriptions` | Table name for webhook subscriptions |
| `NotificationQueueName` | Queue name for notification processing |
| `SharePointTenantName` | SharePoint tenant hostname (e.g. `contoso.sharepoint.com`) |
| `WebhookUrl` | Public URL of the `ProcessWebhookNotification` endpoint |

### 9. Dependency Injection & Host Setup

`Program.cs` wires up:

- Azure Functions web application host with Application Insights telemetry
- `AppSettings` bound from configuration
- `TableServiceClient` and `QueueServiceClient` for Azure Storage
- `NotificationRegistryService` – notification registration CRUD
- `WebhookSubscriptionService` – webhook subscription CRUD
- `WebhookService` – SharePoint webhook register/remove/renew
- `DataService` – Microsoft Graph delta queries
- `AINotificationService` – AI notification processing (placeholder)

## Architecture Flow

```
User creates registration
        │
        ▼
NotificationServiceFunction.CreateRegistration
        │
        ├─► Save to NotificationRegistrations table
        │
        └─► WebhookService.RegisterWebhookAsync
                ├─► SharePoint: list.AddWebhookSubscription()
                └─► Save to WebhookSubscriptions table
                            │
        ┌───────────────────┘
        ▼
SharePoint fires webhook notification
        │
        ▼
ProcessingServiceFunction.ProcessWebhookNotification
        │
        ├─► Validate (echo validationtoken)
        ├─► Correlate with registrations (SiteId + WebId + ListId)
        └─► Queue notification message
                │
                ▼
NotifierServiceFunction.ProcessNotificationQueue
        │
        ├─► DataService.GetDeltaAsync (Graph delta query)
        ├─► Filter by ChangeType (TODO)
        ├─► AI summarization (TODO)
        └─► Send via Teams/Email (TODO)

Monthly timer (1st of month, 5:00 AM UTC)
        │
        ▼
WebhookRenewalServiceFunction.MonthlyRenewWebhooks
        │
        ├─► Query expiring subscriptions (next 45 days)
        └─► WebhookService.RenewWebhookAsync
                ├─► SharePoint: list.UpdateWebhookSubscription()
                └─► Update WebhookSubscriptions table
```

## Local Development

### Prerequisites

- .NET 10 SDK
- Azure Functions Core Tools v4
- Azurite (local storage emulator) or an Azure Storage account

### Configuration

Copy `local.settings - template.json` to `local.settings.json` and fill in the required values:

```json
{
  "IsEncrypted": false,
  "Values": {
    "AzureWebJobsStorage": "UseDevelopmentStorage=true",
    "FUNCTIONS_WORKER_RUNTIME": "dotnet-isolated",
    "AADAppId": "ENTRA_APP_ID",
    "AADAppSecret": "ENTRA_APP_SECRET",
    "TenantId": "ENTRA_TENANT_ID",
    "VaultUri": "KEY_VAULT_URI",
    "VaultCertName": "KEY_VAULT_CERT_NAME",
    "TableNotificationRegistrations": "NotificationRegistrations",
    "TableWebhookSubscriptions": "WebhookSubscriptions",
    "NotificationQueueName": "notifications",
    "SharePointTenantName": "SHAREPOINT_TENANT_NAME",
    "WebhookUrl": "WEBHOOK_NOTIFICATION_URL"
  }
}
```

### Run

```bash
func start
```
