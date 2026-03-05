using Azure.Identity;
using Azure.Security.KeyVault.Certificates;
using functionApp.Models;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using Microsoft.SharePoint.Client;
using PnP.Framework;
using System;
using System.Collections.Generic;
using System.Security.Cryptography.X509Certificates;
using System.Text;

namespace functionApp.Helpers
{
    public static class ConnectionHelper
    {
        // Static fields to cache the certificate and clients
        private static X509Certificate2 _cachedCertificate;
        private static readonly object _certificateLock = new object();
        private static GraphServiceClient _cachedGraphClient;
        private static DateTime _certificateRetrievalTime = DateTime.MinValue;
        private static readonly TimeSpan _certificateExpirationWindow = TimeSpan.FromHours(1); // Re-fetch certificate after 1 hour

        public static ClientContext GetContext(this AppSettings env, string site, ILogger logging = null)
        {
            X509Certificate2 cert2 = GetCertificate(env, logging);
            AuthenticationManager authmanager = new(env.AADAppId, cert2, env.TenantId);
            return authmanager.GetContext(site);
        }

        public static string GetAppOnlyAccessToken(this AppSettings env, string siteUrl, ILogger logging)
        {
            X509Certificate2 cert2 = GetCertificate(env, logging);
            AuthenticationManager authmanager = new(env.AADAppId, cert2, env.TenantId);
            return authmanager.GetAccessToken(siteUrl);
        }

        private static X509Certificate2 GetCertificate(AppSettings env, ILogger logging)
        {
            // Check if we need to refresh the certificate (first request or cache expiring)
            bool needsRefresh = false;

            lock (_certificateLock)
            {
                // Check if certificate is null or approaching expiration time
                if (_cachedCertificate == null ||
                    (DateTime.UtcNow - _certificateRetrievalTime) > _certificateExpirationWindow)
                {
                    needsRefresh = true;
                }

                if (!needsRefresh)
                {
                    logging?.LogInformation("Using cached certificate");
                    return _cachedCertificate;
                }

                // Retrieve fresh certificate
                logging?.LogInformation("Certificate cache empty or expired, retrieving new certificate");
                _cachedCertificate = env.GetCertificateFromKeyVault(logging);
                _certificateRetrievalTime = DateTime.UtcNow;

                return _cachedCertificate;
            }
        }

        public static X509Certificate2 GetCertificateFromKeyVault(this AppSettings env, ILogger logging)
        {
            logging?.LogInformation("Starting certificate download from KeyVault: {vaultUri}, certificate: {certName}", env.VaultUri, env.VaultCertName);

            try
            {

                logging?.LogInformation("Creating CertificateClient with DefaultAzureCredential");
                var client = new CertificateClient(new Uri(env.VaultUri), new DefaultAzureCredential(includeInteractiveCredentials: true));

                logging?.LogInformation("Attempting to download certificate");
                var certResponse = client.DownloadCertificate(env.VaultCertName);

                if (certResponse == null)
                {
                    logging?.LogInformation("Certificate Retrieved is null");
                    throw new Exception("Certificate could not be retrieved from Key Vault.");
                }

                logging?.LogInformation("Certificate successfully downloaded. Subject: {subject}", certResponse.Value.Subject);
                return certResponse.Value;
            }
            catch (Exception ex)
            {
                logging?.LogError(ex, "Failed to download certificate from KeyVault. Error: {message}", ex.Message);

                // Log inner exception details if available
                if (ex.InnerException != null)
                {
                    logging?.LogError("Inner exception: {innerMessage}", ex.InnerException.Message);
                }

                throw; // Re-throw the exception to maintain the original behavior
            }
        }

        public static GraphServiceClient GraphClient(this AppSettings env, ILogger logging)
        {
            // Return cached client if available
            if (_cachedGraphClient != null)
            {
                logging?.LogDebug("Returning cached Graph client");
                return _cachedGraphClient;
            }

            X509Certificate2 cert2 = GetCertificate(env, logging);
            var authCodeCredential = new ClientCertificateCredential(env.TenantId, env.AADAppId, cert2);
            _cachedGraphClient = new GraphServiceClient(authCodeCredential);
            return _cachedGraphClient;
        }
    }
}
