# Local configuration for this source import

The source was imported without production credentials, runtime data, build outputs or old Git history.

Copy appsettings.example.json to appsettings.json and fill its settings locally. The real file is ignored by Git.
Azure Key Vault uses OYA_KEY_VAULT_URI and DefaultAzureCredential. Configure managed identity or the standard Azure environment variables outside Git.
PubNub publish, subscribe and secret keys belong in private configuration.
Existing workflows are retained with a .disabled suffix and will not deploy automatically. Review their repository and runner settings before enabling them.
