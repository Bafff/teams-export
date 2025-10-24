# Teams Export

Command-line utility for exporting Microsoft Teams 1:1 and group chat messages using the Microsoft Graph API.

![Arkadium MS Teams Chats Archive Export Flow](docs/teams-export-flow.svg)

Additional background lives in the internal wiki: [Arkadium IT Knowledge Base](https://arkadium.atlassian.net/wiki/spaces/IT/overview).

## Setup

1. Ensure Python 3.10 or later is available.
   - Fastest path: [`uv`](https://docs.astral.sh/uv/) – `uv python install 3.11 && uv venv --python 3.11 && uv pip install -e .`
   - Alternative: provision Python ≥3.10 and use `python -m venv .venv && source .venv/bin/activate && pip install -e .`
2. Install the project in editable mode:

   ```bash
   pip install -e .
   ```

3. Create (or import) the Azure AD application **Arkadium MS Teams Chats Archive Export** with delegated permissions `Chat.Read` and `Chat.ReadBasic`. You can import `azure/app-manifest.json` during registration to pre-populate the correct scope list, internal note, and wiki/home page URLs so tenant admins see the documentation context.
   - After creation, grant admin consent once so end users do not see repeated prompts.
   - Record the generated `Application (client) ID` and, if applicable, your tenant ID.
4. Copy `config.sample.json` to `~/.teams-exporter/config.json` and update the placeholders:

   ```json
   {
     "client_id": "YOUR_CLIENT_ID",
     "authority": "https://login.microsoftonline.com/YOUR_TENANT_ID",
    "scopes": ["Chat.Read", "Chat.ReadBasic"],
     "token_cache_path": "~/.teams-exporter/token_cache.json"
   }
   ```

   You can also provide `TEAMS_EXPORT_CLIENT_ID` and related environment variables instead of a config file.

## Usage

```
teams-export --user "john.smith@company.com" --from 2025-10-23 --to 2025-10-23 --format json
```

- `--user` targets 1:1 chats by participant name or email.
- `--chat` targets group chats by display name.
- `--from` / `--to` accept `YYYY-MM-DD`, `today`, or `last week`.
- `--format` supports `json` (default) or `csv`.
- `--list` prints available chats with participants.
- `--all` exports every chat in the provided window.
- `--force-login` clears the cache and forces a new device code login.

Exports are saved under `./exports/` by default with filenames like `john_smith_2025-10-23.json`.

## Token Cache

MSAL token cache is stored at `~/.teams-exporter/token_cache.json`. The cache refreshes automatically; re-run with `--force-login` to regenerate the device flow.

## Limitations

- Requires delegated permissions for the signed-in user.
- Attachments are referenced in the output but not downloaded.
- Microsoft Graph API throttling is not yet handled with automatic retries.

## Security Notes

- The CLI never stores usernames or passwords; authentication uses Azure AD device code flow with delegated scopes.
- Refresh and access tokens are cached locally in the path you configure (`token_cache.json`). Rotate/clear the cache by deleting that file or running with `--force-login`.
- No application secrets or certificates are created for this public client; there are no service-principal credentials to rotate unless you deliberately add them later.

## Azure AD App Automation

Prefer commands over the Azure Portal? The scripts below use the templates under `azure/` to reproduce the same configuration.

```bash
# 1. Create the public client app with delegated chat scopes
az ad app create \
  --display-name "Arkadium MS Teams Chats Archive Export" \
  --sign-in-audience AzureADMyOrg \
  --is-fallback-public-client \
  --public-client-redirect-uris https://login.microsoftonline.com/common/oauth2/nativeclient \
  --required-resource-accesses @azure/required-resource-accesses.json

# Capture the returned identifiers
#   appId  -> client ID used in config.sample.json
#   id     -> application object ID for subsequent PATCH/PUT calls

# 2. Apply internal note + documentation links
az rest \
  --method PATCH \
  --uri "https://graph.microsoft.com/v1.0/applications/<application-object-id>" \
  --headers Content-Type=application/json \
  --body '{
    "notes": "This application can be used to retrieve your history of the conversations from MS Teams using Graph API and Python.",
    "info": { "marketingUrl": "https://arkadium.atlassian.net/wiki/spaces/IT/overview" },
    "web": { "homePageUrl": "https://arkadium.atlassian.net/wiki/spaces/IT/overview" }
  }'

# 3. Upload the consent screen logo (PNG or JPG)
az rest \
  --method PUT \
  --uri "https://graph.microsoft.com/v1.0/applications/<application-object-id>/logo" \
  --headers "Content-Type=image/png" \
  --body @/path/to/logo.png

# Optional: mirror the logo to the enterprise application
az rest \
  --method PUT \
  --uri "https://graph.microsoft.com/v1.0/servicePrincipals/<service-principal-object-id>/logo" \
  --headers "Content-Type=image/png" \
  --body @/path/to/logo.png

# 4. Grant tenant-wide consent once the Graph permissions look correct
az ad app permission admin-consent --id <appId>
```

- Replace the placeholder IDs with the values returned from the create command (`appId` for client ID and `id` for subsequent REST operations; service principal ID appears in `az ad sp list --filter "appId eq '<appId>'"`).
- The same `required-resource-accesses.json` is what the manifest references; use one or the other to keep scope definitions in sync.
