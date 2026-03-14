# SharePoint Notifications Backend

Azure Functions backend for managing SharePoint notification registrations. Users can subscribe to changes on specific SharePoint lists/libraries and receive AI-summarized notifications via Teams or Email. The backend automatically registers and manages SharePoint webhook subscriptions, processes list item changes via Microsoft Graph delta queries with persistent checkpointing, and generates human-readable notification summaries using the GitHub Copilot SDK.

## Tech Stack

- **.NET 10** with **Azure Functions v4** (isolated worker model)
- **Azure Table Storage** for persisting notification registrations, webhook subscriptions, and delta token checkpoints
- **Azure Queue Storage** for asynchronous notification processing
- **Azure Key Vault** for certificate-based authentication
- **Microsoft Graph SDK** for Microsoft 365 integration and list item delta queries
- **PnP Framework** for SharePoint CSOM connectivity and webhook management
- **GitHub Copilot SDK** for AI-powered notification summarization (GPT-5)
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
│   ├── NotifierServiceFunction.cs          # Queue-triggered notification processor (delta + AI + dispatch)
│   └── WebhookRenewalServiceFunction.cs    # Timer-triggered webhook renewal
├── Helpers/
│   └── ConnectionHelper.cs                 # SharePoint CSOM & Graph authentication helpers
├── Models/
│   ├── AppSettings.cs                      # Environment-based configuration model (reflection-populated)
│   ├── DeltaChangeType.cs                  # Enum: Created, Updated, Deleted
│   ├── DeltaEntity.cs                      # Azure Table Storage entity for delta token checkpoints
│   ├── DeltaItemChange.cs                  # Represents a single list item change with its classification
│   ├── NotificationRegistration.cs         # Core registration domain model + enums (ChangeType, NotificationChannel)
│   ├── NotificationRegistrationEntity.cs   # Azure Table Storage entity mapping for registrations
│   ├── ProcessingModels.cs                 # Queue message, webhook notification wrapper & resource info models
│   ├── WebhookNotificationModel.cs         # SharePoint webhook payload model (Newtonsoft.Json)
│   └── WebhookSubscriptionEntity.cs        # Azure Table Storage entity for webhook subscriptions
└── Services/
    ├── AINotificationService.cs            # AI-powered notification summarization via GitHub Copilot SDK
    ├── DeltaService.cs                     # Microsoft Graph delta queries with persistent token checkpointing
    ├── NotificationRegistryService.cs      # CRUD operations for notification registrations
    ├── WebhookService.cs                   # SharePoint webhook register, remove & renew operations
    └── WebhookSubscriptionService.cs       # CRUD operations for webhook subscription records
```

## Implemented Functionalities

### 1. Notification Registration CRUD API

Full REST API for managing notification registrations (`NotificationServiceFunction`). All endpoints use `AuthorizationLevel.Function` (require a function key).

| Method | Route | Function | Description |
|--------|-------|----------|-------------|
| `POST` | `/api/registrations` | `CreateRegistration` | Create a new notification registration and register the webhook on the SharePoint list |
| `GET` | `/api/registrations/{userId}` | `GetUserRegistrations` | Get all registrations for a user |
| `GET` | `/api/registrations/{userId}/{id}` | `GetRegistration` | Get a specific registration by ID |
| `PUT` | `/api/registrations/{userId}/{id}` | `UpdateRegistration` | Update an existing registration |
| `DELETE` | `/api/registrations/{userId}/{id}` | `DeleteRegistration` | Delete a registration and remove the webhook if no other registrations exist for the same list |

Request/response payloads use the `NotificationRegistration` model:

- **Id** (`Guid`) – Unique registration identifier (auto-generated on create)
- **UserId** (`Guid`) – The subscribing user's ID
- **ChangeType** (`enum`) – `CREATED`, `UPDATED`, `DELETED`, or `ALL`
- **SiteId / WebId / ListId** (`Guid`) – SharePoint resource identifiers
- **SiteUrl** (`string?`) – Full SharePoint site URL (e.g. `https://tenant.sharepoint.com/sites/MySite`)
- **ItemId** (`int?`) – Optional specific item to watch
- **NotificationChannels** (`NotificationChannel[]`) – Array of channels: `TEAMS` and/or `EMAIL`
- **Description** (`string?`) – Optional description of the registration

**Input validation:**
- `UserId` must not be `Guid.Empty`
- At least one `NotificationChannel` is required

**Webhook registration behavior on create:**
- If `SiteUrl` is provided, a SharePoint webhook is registered on the target list
- Webhook registration failure is logged but does not block the API response (the registration is still saved)
- If `SiteUrl` is missing, a warning is logged and the webhook is skipped

**Webhook cleanup behavior on delete:**
- The function queries remaining registrations for the same list (by `SiteId`, `WebId`, `ListId`)
- Only removes the webhook if this was the last registration for that list
- Webhook removal failure is logged but does not block the delete response

### 2. SharePoint Webhook Lifecycle Management

Webhooks are automatically managed through the `WebhookService`:

- **Registration** – When a notification registration is created, a SharePoint webhook is registered on the specified list/library via PnP Framework's `AddWebhookSubscription` (4-month expiration). The subscription is persisted to the `WebhookSubscriptions` Azure Table including `ListId`, `SiteUrl`, `NotificationUrl`, `Resource`, `ExpirationDateTime`, and `LastUpdated`.
- **Removal** – When the last notification registration for a given list is deleted, the corresponding webhook subscription is removed from SharePoint (matched by `NotificationUrl`) and the table record is cleaned up.
- **Renewal** – A timer-triggered function (`MonthlyRenewWebhooks`) runs on the 1st of every month at 5:00 AM UTC (`0 0 5 1 * *`). It queries all subscriptions expiring within 45 days and renews them via PnP Framework's `UpdateWebhookSubscription`, extending the expiration by 6 months from the current date. Stale subscriptions (no longer present on the SharePoint list) are automatically detected and removed from table storage.

### 3. SharePoint Webhook Notification Processing

The webhook processing pipeline (`ProcessingServiceFunction`):

1. **Validation** – Handles SharePoint's `validationtoken` query-string handshake when webhooks are being registered (returns the token as plain text with `200 OK`)
2. **Notification Receiving** – `POST /api/ProcessWebhookNotification` receives webhook notifications from SharePoint (supports batched notifications via the `value` array). Includes robust error handling for incomplete/truncated payloads and IO errors.
3. **Site Resolution** – For each notification, the function reconstructs the full site URL from `SharePointTenantName` + `notification.SiteUrl` and loads the SharePoint Site ID via CSOM
4. **Registration Correlation** – Each notification is correlated with matching registrations in table storage by `WebId` and `ListId` (the `Resource` field). Note: `SiteId` filtering is currently bypassed.
5. **Queue Dispatch** – Matched notifications are wrapped in a `NotificationQueueMessage` (including all matching registrations, the webhook notification, and a `QueuedAt` timestamp), serialized to JSON, Base64-encoded, and sent to the `notifications` Azure Queue

### 4. Queue-Based Notification Processing

The `NotifierServiceFunction` processes messages from the `notifications` queue:

1. **Deserialization** – Deserializes the `NotificationQueueMessage` containing matched registrations and the webhook notification
2. **Delta Retrieval** – Retrieves list item changes via Microsoft Graph delta queries (`DeltaService.GetDeltaForNotificationAsync`)
3. **Change Classification** – Delta items are classified into `Created`, `Updated`, and `Deleted` categories
4. **ChangeType Filtering** – Registrations are grouped by their `ChangeType` filter. Each group receives only the relevant subset of delta items:
   - `CREATED` registrations → only created items
   - `UPDATED` registrations → only updated items
   - `DELETED` registrations → only deleted items
   - `ALL` registrations → all items
5. **AI Summarization** – For each registration, the matched items are processed through `AINotificationService.ProcessNotificationAsync` to generate a human-readable notification summary
6. **Notification Dispatch** – The generated summary is dispatched through the configured channels:
   - `TEAMS` → Send via Microsoft Teams (TODO: Graph API integration)
   - `EMAIL` → Send via email (TODO: Graph API integration)

Failed messages are re-thrown to be moved to the poison queue for manual inspection.

### 5. AI-Powered Notification Summarization

The `AINotificationService` uses the **GitHub Copilot SDK** to generate human-readable notification summaries:

- **Session creation** – Creates a `CopilotClient` session using a GitHub token (`GitHubToken` setting) with the **GPT-5** model
- **System prompt** – Configures the AI as a "SharePoint notification assistant" focused on summarizing list/library changes into clear, concise, actionable messages
- **Prompt construction** – Sends the registration description, change type filter, and a JSON representation of the delta items (including item IDs, change types, and field values from `AdditionalData`)
- **Fallback** – If the AI call fails, a basic summary is generated listing the count of created, updated, and deleted items (e.g. "SharePoint changes detected: 3 item(s) created, 1 item(s) updated.")
- **Empty changes** – Returns "No changes detected." when no delta items are present

### 6. Azure Table Storage Persistence

Three tables are used (all auto-created on startup):

- **`NotificationRegistrations`** – Stores user notification registrations
  - **PartitionKey** = `UserId` (Guid as string), **RowKey** = Registration `Id` (Guid as string)
  - Complex fields (`NotificationChannels`) serialized as JSON string
  - `ChangeType` stored as string enum name

- **`WebhookSubscriptions`** – Stores active SharePoint webhook subscriptions
  - **PartitionKey** = `"Webhooks"`, **RowKey** = SharePoint Subscription `Id`
  - Tracks `ListId`, `SiteUrl`, `NotificationUrl`, `Resource`, `ExpirationDateTime`, `ChangeToken`, and `LastUpdated`

- **`Deltas`** – Stores Microsoft Graph delta token checkpoints per webhook subscription
  - **PartitionKey** = `"Deltas"`, **RowKey** = Webhook Subscription `Id`
  - Stores the `DeltaLink` URL used for incremental delta queries
  - Enables efficient change tracking across webhook notification cycles

### 7. SharePoint & Microsoft Graph Authentication

The `ConnectionHelper` (static class) provides authenticated connections using certificate-based app-only authentication:

- **Certificate retrieval from Azure Key Vault** using `DefaultAzureCredential` (with interactive credentials enabled)
- **Thread-safe certificate caching** with a 1-hour expiration window (uses `lock` for thread safety)
- **SharePoint CSOM context** via PnP Framework `AuthenticationManager` (`GetContext` extension method on `AppSettings`)
- **App-only access token** generation for SharePoint REST calls (`GetAppOnlyAccessToken` extension method)
- **Microsoft Graph client** with `ClientCertificateCredential` (cached as static singleton)
- **Local development** (DEBUG builds): loads certificates from the current user certificate store by thumbprint (`CertThumbprint` setting), bypassing Key Vault

### 8. Microsoft Graph Delta Queries

The `DeltaService` retrieves and tracks list item changes via Microsoft Graph's delta endpoint with persistent checkpointing:

- **Initial sync** – On first call for a subscription, performs a full delta query (`/sites/{siteId}/lists/{listId}/items/delta?$top=100`) with pagination support via `PageIterator`
- **Incremental sync** – On subsequent calls, uses the stored `DeltaLink` from the `Deltas` table to request only changes since the last checkpoint
- **Change classification** – Each item is classified as:
  - `Deleted` – if the item has the `Deleted` marker set
  - `Created` – if the `LastModifiedDateTime` is within 5 minutes of `CreatedDateTime`
  - `Updated` – all other non-deleted items (default)
- **Delta checkpoint persistence** – After processing, the new `DeltaLink` from the Graph response is saved to the `Deltas` Azure Table for the next notification cycle
- **Deduplication** – Tracks processed item IDs in a `HashSet` to prevent duplicate entries across pages
- **Site resolution** – Resolves the SharePoint site via Graph using the tenant name and site URL path from the webhook notification

### 9. Configuration Management

`AppSettings` auto-populates properties from environment variables using reflection. The constructor iterates over all system environment variables and maps matching property names to their values.

| Setting | Description |
|---------|-------------|
| `AADAppId` | Entra ID (Azure AD) application ID |
| `AADAppSecret` | Entra ID application secret |
| `TenantId` | Azure tenant ID |
| `VaultUri` | Azure Key Vault URI |
| `VaultCertName` | Certificate name in Key Vault |
| `CertThumbprint` | Certificate thumbprint (local development only) |
| `AzureWebJobsStorage` | Azure Storage connection string |
| `TableNotificationRegistrations` | Table name for notification registrations |
| `TableWebhookSubscriptions` | Table name for webhook subscriptions |
| `TableDeltas` | Table name for delta token checkpoints |
| `NotificationQueueName` | Queue name for notification processing |
| `SharePointTenantName` | SharePoint tenant hostname (e.g. `contoso.sharepoint.com`) |
| `WebhookUrl` | Public URL of the `ProcessWebhookNotification` endpoint |
| `GitHubToken` | GitHub personal access token for Copilot SDK authentication |

### 10. Dependency Injection & Host Setup

`Program.cs` wires up:

- Azure Functions web application host (`FunctionsApplication.CreateBuilder` + `ConfigureFunctionsWebApplication`)
- Application Insights telemetry (`AddApplicationInsightsTelemetryWorkerService` + `ConfigureFunctionsApplicationInsights`)
- `AppSettings` bound from the `"AppSettings"` configuration section (also self-populated from environment variables via constructor reflection)
- `TableServiceClient` and `QueueServiceClient` for Azure Storage (using `AzureWebJobsStorage` connection string, defaults to `UseDevelopmentStorage=true`)
- `NotificationRegistryService` – notification registration CRUD
- `WebhookSubscriptionService` – webhook subscription CRUD
- `WebhookService` – SharePoint webhook register/remove/renew
- `DeltaService` – Microsoft Graph delta queries with persistent checkpointing
- `AINotificationService` – AI notification summarization via GitHub Copilot SDK
- `HttpClient` – registered via `AddHttpClient()` for DeltaService

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
                ├─► SharePoint: list.AddWebhookSubscription() [4-month expiry]
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
        ├─► Resolve site via CSOM (SharePointTenantName + SiteUrl)
        ├─► Correlate with registrations (WebId + ListId)
        └─► Queue NotificationQueueMessage (Base64-encoded JSON)
                │
                ▼
NotifierServiceFunction.ProcessNotificationQueue
        │
        ├─► DeltaService.GetDeltaForNotificationAsync
        │       ├─► Load delta checkpoint from Deltas table
        │       ├─► Graph delta query (initial or incremental)
        │       ├─► Classify items: Created / Updated / Deleted
        │       └─► Save new delta checkpoint to Deltas table
        │
        ├─► Group registrations by ChangeType filter
        ├─► AINotificationService.ProcessNotificationAsync
        │       ├─► GitHub Copilot SDK session (GPT-5)
        │       └─► Fallback: basic count summary on failure
        │
        └─► SendNotificationAsync per channel
                ├─► TEAMS (TODO: Graph API)
                └─► EMAIL (TODO: Graph API)

Monthly timer (1st of month, 5:00 AM UTC)
        │
        ▼
WebhookRenewalServiceFunction.MonthlyRenewWebhooks
        │
        ├─► Query expiring subscriptions (next 45 days)
        └─► For each subscription:
                ├─► WebhookService.RenewWebhookAsync
                │       ├─► SharePoint: list.UpdateWebhookSubscription() [+6 months]
                │       └─► Update WebhookSubscriptions table
                └─► Remove stale subscriptions not found on SharePoint
```

## Pending Work (TODOs)

- **Teams notification delivery** – `SendTeamsNotificationAsync` in `NotifierServiceFunction` needs Graph API integration to send activity feed notifications or chat messages
- **Email notification delivery** – `SendEmailNotificationAsync` in `NotifierServiceFunction` needs Graph API integration to send email via `sendMail`

## Local Development

### Prerequisites

- .NET 10 SDK
- Azure Functions Core Tools v4
- Azurite (local storage emulator) or an Azure Storage account
- A certificate installed in the current user certificate store (for local SharePoint/Graph authentication)
- A GitHub personal access token (for Copilot SDK AI summarization)

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
    "CertThumbprint": "LOCAL_CERT_THUMBPRINT",
    "TableNotificationRegistrations": "NotificationRegistrations",
    "TableWebhookSubscriptions": "WebhookSubscriptions",
    "TableDeltas": "Deltas",
    "NotificationQueueName": "notifications",
    "SharePointTenantName": "SHAREPOINT_TENANT_NAME",
    "WebhookUrl": "WEBHOOK_NOTIFICATION_URL",
    "GitHubToken": "GITHUB_PERSONAL_ACCESS_TOKEN"
  }
}
```

### Run

```bash
func start
```
