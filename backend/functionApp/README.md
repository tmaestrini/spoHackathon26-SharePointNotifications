# SharePoint Notifications Backend

Azure Functions backend for managing SharePoint notification registrations. Users can subscribe to changes on specific SharePoint lists/libraries and receive AI-summarized notifications via Teams or Email. The backend automatically registers and manages SharePoint webhook subscriptions, processes list item changes via Microsoft Graph delta queries with persistent checkpointing, enriches changes with version history and file content diffs, and generates human-readable notification summaries using Azure AI Foundry.

## Tech Stack

- **.NET 10** with **Azure Functions v4** (isolated worker model)
- **Azure Table Storage** for persisting notification registrations, webhook subscriptions, and delta token checkpoints
- **Azure Queue Storage** for asynchronous notification processing
- **Azure Key Vault** for certificate-based authentication
- **Microsoft Graph SDK** for Microsoft 365 integration, list item delta queries, and user resolution
- **PnP Framework** for SharePoint CSOM connectivity, webhook management, version history, and recycle bin access
- **Azure AI Foundry** for AI-powered notification summarization (via Azure OpenAI chat completions API)
- **Power Automate Flow** for notification delivery (Teams and Email via HTTP trigger)
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
│   ├── ConnectionHelper.cs                 # SharePoint CSOM & Graph authentication helpers
│   ├── DocumentTextExtractor.cs            # Text extraction from PDF, DOCX, DOC, and plain text files
│   └── VersionHelper.cs                    # Enriches delta changes with version history & file content
├── Models/
│   ├── AppSettings.cs                      # Environment-based configuration model (reflection-populated)
│   ├── DeltaChangeType.cs                  # Enum: Created, Updated, Deleted
│   ├── DeltaEntity.cs                      # Azure Table Storage entity for delta token checkpoints
│   ├── DeltaItemChange.cs                  # Represents a single list item change with version & file info
│   ├── NotificationRegistration.cs         # Core registration domain model + enums (ChangeType, NotificationChannel)
│   ├── NotificationRegistrationEntity.cs   # Azure Table Storage entity mapping for registrations
│   ├── ProcessingModels.cs                 # Queue message, webhook notification wrapper & resource info models
│   ├── WebhookNotificationModel.cs         # SharePoint webhook payload model (Newtonsoft.Json)
│   └── WebhookSubscriptionEntity.cs        # Azure Table Storage entity for webhook subscriptions
└── Services/
    ├── FoundryAINotificationService.cs     # AI notification summarization via Azure AI Foundry
    ├── DeltaService.cs                     # Microsoft Graph delta queries with version enrichment & recycle bin
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
2. **Delta Retrieval** – Retrieves list item changes via Microsoft Graph delta queries (`DeltaService.GetDeltaForNotificationAsync`), including version enrichment and deleted item enrichment from the recycle bin
3. **Change Classification** – Delta items are classified into `Created`, `Updated`, and `Deleted` categories
4. **ChangeType Filtering** – Registrations are grouped by their `ChangeType` filter. Each group receives only the relevant subset of delta items:
   - `CREATED` registrations → only created items
   - `UPDATED` registrations → only updated items
   - `DELETED` registrations → only deleted items
   - `ALL` registrations → all items
5. **AI Summarization** – For each registration, the matched items are processed through `FoundryAINotificationService.ProcessNotificationAsync` to generate a human-readable, structured notification summary
6. **Notification Dispatch** – The generated summary is dispatched through the configured channels via a Power Automate Flow:
   - Resolves the user's `UserPrincipalName` via Microsoft Graph
   - Sends a POST request to the `NotificationFlowUrl` endpoint with `userPrincipalName`, `notificationText`, and `notificationType` (the channel name: `TEAMS` or `EMAIL`)
   - The Power Automate Flow handles the actual delivery to Microsoft Teams or Email

Failed messages are re-thrown to be moved to the poison queue for manual inspection.

### 5. AI-Powered Notification Summarization

The `FoundryAINotificationService` uses the Azure AI Foundry (Azure OpenAI) chat completions API to generate human-readable notification summaries:

- **API call** – Sends a `POST` request to `AzureFoundryApiUrl` with `api-key` header authentication (`AzureFoundryApiKey` setting)
- **System prompt** – Configures the AI as a "SharePoint notification assistant" that summarizes changes, highlights field-level differences between versions, and compares current/previous file content for documents
- **Rich prompt construction** – Sends the registration description, change type filter, and a JSON representation of delta items including:
  - Item IDs, change types, and field values from Graph `AdditionalData`
  - Current and previous version metadata (`VersionLabel`, `FieldValues`)
  - Field-level changes between versions (`FieldTitle`, `PreviousValue`, `NewValue`)
  - Extracted text content from current and previous document versions (via `DocumentTextExtractor`)
  - File name and version labels for documents
- **Structured output templates** – The prompt includes specific output format templates for each change type (Created, Updated, Deleted) ensuring consistent, actionable notification messages
- **Fallback** – If the AI call fails, a basic summary is generated listing the count of created, updated, and deleted items
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

### 8. Microsoft Graph Delta Queries & Change Enrichment

The `DeltaService` retrieves and tracks list item changes via Microsoft Graph's delta endpoint with persistent checkpointing, and enriches changes with version history and recycle bin metadata.

#### Delta Query Processing

- **Initial sync** – On first call for a subscription, performs a full delta query (`/sites/{siteId}/lists/{listId}/items/delta?$top=100`) with pagination support via `PageIterator`
- **Incremental sync** – On subsequent calls, uses the stored `DeltaLink` from the `Deltas` table to request only changes since the last checkpoint
- **Change classification** – Each item is classified as:
  - `Deleted` – if the item has the `Deleted` marker set
  - `Created` – if the `LastModifiedDateTime` is within 5 minutes of `CreatedDateTime`
  - `Updated` – all other non-deleted items (default)
- **Delta checkpoint persistence** – After processing, the new `DeltaLink` from the Graph response is saved to the `Deltas` Azure Table for the next notification cycle
- **Deduplication** – Tracks processed item IDs in a `HashSet` to prevent duplicate entries across pages
- **Site resolution** – Resolves the SharePoint site via Graph using the tenant name and site URL path from the webhook notification

#### Version Enrichment (`VersionHelper`)

After delta processing, non-deleted items are enriched with version history via SharePoint CSOM:

- **Current version info** – Retrieves the current version label and all field values from the list item
- **Previous version info** – If more than one version exists, retrieves the previous version's label and field values
- **Field-level changes** – Extracts field-level changes between the current and previous version (field title, previous value, new value) using `ListItemVersion.Changes`
- **Document file content** – For document library items (items with an associated file):
  - Downloads the current file binary content and records the file name
  - Downloads the previous file version content (if versions exist) and records its version label
  - File content is later extracted as text by `DocumentTextExtractor` during AI summarization

#### Deleted Item Enrichment

Deleted items (which Graph returns without metadata) are enriched from the SharePoint recycle bin:

- **REST API approach (primary)** – Queries the site's recycle bin via `/_api/web/recyclebin?$top=1000` using an app-only access token. Matches deleted items by ID and enriches with `Title`, `DeletedDate`, and `DeletedByName`.
- **CSOM approach (fallback)** – If the REST API fails, falls back to PnP Framework CSOM to query first-stage and second-stage recycle bins. Uses multiple ID matching strategies to handle format differences between Graph and SharePoint IDs.

#### Document Text Extraction (`DocumentTextExtractor`)

Extracts readable text content from file bytes for AI summarization:

- **Supported formats** – `.pdf` (native stream parsing with deflate/ASCII85 decompression and BT/ET text block extraction), `.docx` (OpenXML paragraph extraction), `.doc` (basic binary text extraction), `.txt`, `.csv`, `.json`, `.xml`, `.html`, `.htm`, `.md`
- **Truncation** – Content is capped at 5,000 characters with a truncation indicator
- **Unsupported files** – Returns a placeholder with file name and size for unrecognized file types

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
| `AzureFoundryApiUrl` | Azure AI Foundry (Azure OpenAI) chat completions API endpoint URL |
| `AzureFoundryApiKey` | API key for Azure AI Foundry authentication |
| `NotificationFlowUrl` | Power Automate Flow HTTP trigger URL for sending notifications |
| `NotificationServiceUserName` | Service account user name for notification delivery |
| `NotificationMailSubject` | Subject line for email notifications |

### 10. Dependency Injection & Host Setup

`Program.cs` wires up:

- Azure Functions web application host (`FunctionsApplication.CreateBuilder` + `ConfigureFunctionsWebApplication`)
- Application Insights telemetry (`AddApplicationInsightsTelemetryWorkerService` + `ConfigureFunctionsApplicationInsights`)
- `AppSettings` bound from the `"AppSettings"` configuration section (also self-populated from environment variables via constructor reflection)
- `TableServiceClient` and `QueueServiceClient` for Azure Storage (using `AzureWebJobsStorage` connection string, defaults to `UseDevelopmentStorage=true`)
- `NotificationRegistryService` – notification registration CRUD
- `WebhookSubscriptionService` – webhook subscription CRUD
- `WebhookService` – SharePoint webhook register/remove/renew
- `DeltaService` – Microsoft Graph delta queries with version enrichment
- `FoundryAINotificationService` – AI notification summarization via Azure AI Foundry
- `HttpClient` – registered via `AddHttpClient()` for DeltaService and FoundryAINotificationService

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
        │       ├─► VersionHelper: Enrich with version history & file content
        │       ├─► EnrichDeletedInformation: Enrich from recycle bin (REST + CSOM)
        │       └─► Save new delta checkpoint to Deltas table
        │
        ├─► Group registrations by ChangeType filter
        ├─► FoundryAINotificationService.ProcessNotificationAsync
        │       ├─► Azure AI Foundry chat completion (with version diffs & file content)
        │       └─► Fallback: basic count summary on failure
        │
        └─► SendNotificationAsync per channel
                ├─► Resolve user UPN via Microsoft Graph
                └─► POST to Power Automate Flow (NotificationFlowUrl)
                        ├─► TEAMS notification
                        └─► EMAIL notification

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

- **Deleted item matching** – The recycle bin enrichment for deleted items uses multiple ID matching strategies; a more reliable approach based on `DeletedDateTime` filtering is under consideration
- **SiteId filtering** – Registration correlation currently bypasses `SiteId` filtering, matching only on `WebId` and `ListId`

## Local Development

### Prerequisites

- .NET 10 SDK
- Azure Functions Core Tools v4
- Azurite (local storage emulator) or an Azure Storage account
- A certificate installed in the current user certificate store (for local SharePoint/Graph authentication)
- An Azure AI Foundry (Azure OpenAI) endpoint and API key
- A Power Automate Flow with an HTTP trigger (for notification delivery)

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
    "AzureFoundryApiUrl": "AZURE_FOUNDRY_API_URL",
    "AzureFoundryApiKey": "AZURE_FOUNDRY_API_KEY",
    "NotificationServiceUserName": "SERVICE_ACCOUNT_USER_NAME",
    "NotificationMailSubject": "New SharePoint Changes Detected!",
    "NotificationFlowUrl": "POWER_AUTOMATE_FLOW_HTTP_TRIGGER_URL"
  }
}
```

### Run

```bash
func start
```
