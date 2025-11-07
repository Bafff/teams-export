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

### Quick Start (Interactive Mode)

The simplest way to export a chat is to run without any arguments:

```bash
teams-export
```

This will:
1. Authenticate with Microsoft Graph
2. Show an interactive menu with your 20 most recent chats
3. Let you select the chat by number
4. Export today's messages in Jira-friendly format

### Export by User Email (1:1 chats)

```bash
teams-export --user "john.smith@company.com"
```

### Export by Chat Name (Group chats)

```bash
teams-export --chat "Project Alpha Team"
```

### Export with Date Range

```bash
# Specific dates
teams-export --user "john.smith@company.com" --from 2025-10-23 --to 2025-10-25

# Using keywords
teams-export --user "john.smith@company.com" --from "last week" --to "today"
```

### Export in Different Formats

```bash
# Jira-friendly markdown (default) - perfect for copying into Jira tickets
teams-export --user "john.smith@company.com" --format jira

# JSON for programmatic processing
teams-export --user "john.smith@company.com" --format json

# CSV for spreadsheet analysis
teams-export --user "john.smith@company.com" --format csv
```

### Other Options

- `--list` prints available chats with participants.
- `--all` exports every chat in the provided window (uses parallel processing for speed).
- `--force-login` clears the cache and forces a new device code login.
- `--refresh-cache` forces refresh of chat list (bypasses 5-minute cache).
- `--output-dir` specifies where to save exports (default: `./exports/`).

### Examples

```bash
# Interactive selection with custom date range
teams-export --from "2025-10-01" --to "2025-10-31"

# Export all chats from last week in parallel
teams-export --all --from "last week" --format jira

# List all available chats
teams-export --list

# Export specific user's chat for today
teams-export --user "jane.doe@company.com"
```

Exports are saved under `./exports/` by default with filenames like `john_smith_2025-10-23.txt` (for Jira format) or `john_smith_2025-10-23.json`.

## Caching

### Token Cache
MSAL token cache is stored at `~/.teams-exporter/token_cache.json`. The cache refreshes automatically; re-run with `--force-login` to regenerate the device flow.

### Chat List Cache
To speed up repeated operations, the chat list is cached locally for 5 minutes at `~/.teams-exporter/cache/chats_cache.json`.

**First run:** Loads all chats from API (~30-60 seconds for 1000+ chats)
**Subsequent runs (within 5 min):** Instant load from cache

Use `--refresh-cache` to force a refresh if you know new chats were created.

## Features

### Performance Optimizations
- **Chat list caching**: 5-minute local cache makes repeated runs instant
- **Parallel exports**: When using `--all`, exports multiple chats concurrently (up to 3 at once)
- **Automatic retry**: Handles API rate limiting (429) and server errors (5xx) with exponential backoff
- **Optimized pagination**: Fetches 100 messages per request instead of 50
- **Smart filtering**: Stops fetching when messages are outside the date range

### User Experience Improvements
- **Interactive chat selection**: Beautiful menu with chat names, types, and last activity
- **Multiple match handling**: If search finds multiple chats, shows menu instead of error
- **Jira-ready format**: New default format perfect for pasting into Jira tickets
  - Clean HTML conversion (removes tags, preserves formatting)
  - Quote blocks for easy reading
  - Attachment and reaction indicators
  - Proper timestamp formatting
- **Smart defaults**: Defaults to today's date if not specified
- **Progress tracking**: Shows real-time progress for multi-chat exports

## Limitations

- Requires delegated permissions for the signed-in user.
- Attachments are referenced in the output but not downloaded.
- Parallel exports limited to 3 concurrent requests to avoid API throttling.

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
