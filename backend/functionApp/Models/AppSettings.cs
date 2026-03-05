using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;

namespace functionApp.Models
{
    public class AppSettings
    {
        protected IDictionary localEnvironmentVariables { get => System.Environment.GetEnvironmentVariables(); }
        public string? AADAppId { get; private set; }
        public string? AADAppSecret { get; set; }
        public string? TenantId { get; private set; }
        public string? VaultUri { get; set; }
        public string? VaultCertName { get; set; }
        public string? AzureWebJobsStorage { get; internal set; }
        public string? TableNotificationRegistrations { get; set; }

        private string this[string prop]
        {
            set
            {
                var property = typeof(AppSettings).GetProperty(prop);
                if (property != null)
                {
                    property.SetValue(this, value, null);
                }
            }
            get
            {
                var property = typeof(AppSettings).GetProperty(prop);
                return property != null ? property.GetValue(this, null) as string ?? string.Empty : string.Empty;
            }
        }
        public AppSettings()
        {
            var localprops = typeof(AppSettings).GetProperties();


            foreach (DictionaryEntry e in localEnvironmentVariables)
            {
                var keyString = e.Key?.ToString();
                var isProp = keyString != null && typeof(AppSettings).GetProperties().Select(p => p.Name).ToList().Contains(keyString);
                if (isProp && keyString != null) this[keyString] = e.Value as string ?? "";
            }
        }
    }
}
