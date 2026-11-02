# Publish-only pipeline

Every push to main and a manual Actions run builds a verified Windows publish package.
The build runs on GitHub-hosted Windows. A dedicated runner on the destination downloads only that run's artifact and verifies its SHA-256 and every published file.

Destination: azrwusprdqbr02, C:\Application\OYA_WEB_V8 and successive versions.
The folder number always follows the highest existing numbered release. Old releases remain.
The pipeline never changes IIS, stops applications, launches QuickBooks, runs a sync or updates SQL.

Use the folder shown in the Actions summary, or select the newest numbered folder on the destination server.
Before manually activating a release, copy/configure the private runtime settings and external report paths on the server. Private settings and credentials are excluded from GitHub artifacts.
Frontend Azure identity uses OYA_KEY_VAULT_URI plus DefaultAzureCredential; PubNub credentials also remain server-side.
Backend QB_SQL_Sync.exe.config contains empty runtime values and generated assembly bindings; configure it from the existing server settings before running it.

Only pushes to this repository's main branch and manual main-branch runs can use the publish runner. Pull requests never run on the destination.
The older .disabled workflows are historical references and remain disabled.
