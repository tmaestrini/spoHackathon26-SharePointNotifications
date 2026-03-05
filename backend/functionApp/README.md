# SharePoint Notifications Backend

Azure Functions backend for managing SharePoint notification registrations. Users can subscribe to changes on specific SharePoint sites, webs, lists, or items and receive notifications via Teams or Email.

## Tech Stack

- **.NET 10** with **Azure Functions v4** (isolated worker model)
- **Azure Table Storage** for persisting notification registrations
- **Azure Key Vault** for certificate-based authentication
- **Microsoft Graph SDK** for Microsoft 365 integration
- **PnP Framework** for SharePoint connectivity
- **Application Insights** for telemetry and monitoring

## Project Structure

```
functionApp/
├── Program.cs                          # Host setup & dependency injection
├── host.json                           # Azure Functions host configuration
├── local.settings.json                 # Local development settings
├── Functions/
│   └── API_Func.cs                     # HTTP-triggered API endpoints
├── Helpers/
│   └── ConnectionHelper.cs             # SharePoint & Graph authentication helpers
├── Models/
│   ├── AppSettings.cs                  # Environment-based configuration model
│   ├── NotificationRegistration.cs     # Core registration domain model
│   ├── NotificationRegistrationEntity.cs # Azure Table Storage entity mapping
│   └── WebhookNotificationModel.cs     # SharePoint webhook payload model
└── Services/
    └── NotificationRegistryService.cs  # CRUD operations for registrations
```

## Implemented Functionalities

### 1. Notification Registration CRUD API

Full REST API for managing notification registrations, exposed as Azure Functions HTTP triggers:

| Method | Route | Function | Description |
|--------|-------|----------|-------------|
| `POST` | `/api/registrations` | `CreateRegistration` | Create a new notification registration |
| `GET` | `/api/registrations/{userId}` | `GetUserRegistrations` | Get all registrations for a user |
| `GET` | `/api/registrations/{userId}/{id}` | `GetRegistration` | Get a specific registration by ID |
| `PUT` | `/api/registrations/{userId}/{id}` | `UpdateRegistration` | Update an existing registration |
| `DELETE` | `/api/registrations/{userId}/{id}` | `DeleteRegistration` | Delete a registration |

Request/response payloads use the `NotificationRegistration` model with the following fields:

- **Id** – Unique registration identifier (auto-generated on create)
- **UserId** – The subscribing user's ID
- **ChangeType** – `CREATED`, `UPDATED`, `DELETED`, or `ALL`
- **SiteId / WebId / ListId** – SharePoint resource identifiers
- **ItemId** – Optional specific item to watch
- **NotificationChannels** – Array of channels: `TEAMS` and/or `EMAIL`
- **Description** – Optional description of the registration

Input validation is enforced: `UserId` must be non-empty and at least one `NotificationChannel` is required.

### 2. Azure Table Storage Persistence

Registrations are stored in Azure Table Storage via the `NotificationRegistryService`:

- **PartitionKey** = `UserId` — enables efficient per-user queries
- **RowKey** = Registration `Id`
- Table is auto-created on startup if it doesn't exist
- Entity mapping between domain model and table entity is handled by `NotificationRegistrationEntity` with JSON serialization for complex fields (`NotificationChannels`)

### 3. SharePoint & Microsoft Graph Authentication

The `ConnectionHelper` provides authenticated connections to SharePoint and Microsoft Graph using certificate-based app-only authentication:

- **Certificate retrieval from Azure Key Vault** using `DefaultAzureCredential`
- **Certificate caching** with a 1-hour expiration window (thread-safe)
- **SharePoint CSOM context** via PnP Framework `AuthenticationManager`
- **App-only access token** generation for SharePoint REST calls
- **Microsoft Graph client** with `ClientCertificateCredential`

### 4. SharePoint Webhook Payload Model

The `WebhookNotificationModel` defines the structure for incoming SharePoint webhook notifications, including:

- `SubscriptionId`, `ClientState`, `ExpirationDateTime`
- `Resource`, `TenantId`, `SiteUrl`, `WebId`

### 5. Configuration Management

- **`AppSettings`** model auto-populates from environment variables using reflection
- Supports the following configuration keys:
  - `AADAppId` / `AADAppSecret` – Entra ID (Azure AD) app credentials
  - `TenantId` – Azure tenant
  - `VaultUri` / `VaultCertName` – Key Vault certificate location
  - `AzureWebJobsStorage` – Storage connection string
  - `TableNotificationRegistrations` – Table name for registrations

### 6. Dependency Injection & Host Setup

`Program.cs` wires up:

- Azure Functions web application host
- Application Insights telemetry
- `AppSettings` bound from configuration
- `TableServiceClient` for Azure Table Storage
- `NotificationRegistryService` as a singleton

## Local Development

### Prerequisites

- .NET 10 SDK
- Azure Functions Core Tools v4
- Azurite (local storage emulator) or an Azure Storage account

### Configuration

Copy `local.settings.json` and fill in the required values:

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
    "TableNotificationRegistrations": "NotificationRegistrations"
  }
}
```

### Run

```bash
func start
```
