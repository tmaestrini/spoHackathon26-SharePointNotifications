using functionApp.Models;
using functionApp.Services;
using Azure.Data.Tables;
using Azure.Storage.Queues;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Azure.Functions.Worker.Builder;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Options;

var builder = FunctionsApplication.CreateBuilder(args);

builder.ConfigureFunctionsWebApplication();

builder.Services.AddApplicationInsightsTelemetryWorkerService().ConfigureFunctionsApplicationInsights();

builder.Services.Configure<AppSettings>(builder.Configuration.GetSection("AppSettings"));
builder.Services.AddSingleton(resolver => resolver.GetRequiredService<IOptions<AppSettings>>().Value);

// Register Table Storage and NotificationRegistryService
var storageConnectionString = builder.Configuration["AzureWebJobsStorage"] ?? "UseDevelopmentStorage=true";
builder.Services.AddSingleton(new TableServiceClient(storageConnectionString));
builder.Services.AddSingleton<NotificationRegistryService>();

// Register Queue Storage for notification processing
builder.Services.AddSingleton(new QueueServiceClient(storageConnectionString));

// Register HttpClient and SharePointService
builder.Services.AddHttpClient();
builder.Services.AddSingleton<DeltaService>();
builder.Services.AddSingleton<WebhookSubscriptionService>();
builder.Services.AddSingleton<WebhookService>();
builder.Services.AddSingleton<AINotificationService>();
builder.Services.AddSingleton<FoundryAINotificationService>();

builder.Build().Run();
