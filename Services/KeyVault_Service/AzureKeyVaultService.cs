using Azure.Identity;
using Azure.Security.KeyVault.Secrets;
using OYA_Web_Application.Models.Services;
using System;

namespace OYA_Web_Application.Services.KeyVault_Service
{
	public class AzureKeyVaultService
	{
		private Uri keyVaultUri => new Uri(Environment.GetEnvironmentVariable("OYA_KEY_VAULT_URI") ?? throw new InvalidOperationException("Configure OYA_KEY_VAULT_URI."));
		private readonly DefaultAzureCredential credential = new DefaultAzureCredential();

		public string RetrieveSecret(string secretName)
		{
			string secretValue = "";
			try
			{
				// string secretName = "LDAP-Server";
				//string secretName = "SQL-ConnectionString";
				var client = new SecretClient(keyVaultUri, credential);
				return client.GetSecret(secretName)?.Value?.Value ?? "";
			}
			catch
			{
				return secretValue;
			}
		}
	}
}

