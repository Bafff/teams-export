# Teams Export

Command-line utility for exporting Microsoft Teams 1:1 and group chat messages using the Microsoft Graph API.

## Setup

1. Ensure Python 3.10 or later is available.
2. Install the project in editable mode:

   ```bash
   pip install -e .
   ```

3. Create an Azure AD application with delegated permissions `Chat.Read`, `Chat.ReadBasic`, and `Chat.ReadWrite`. Note the client ID.
4. Add a config file at `~/.teams-exporter/config.json`:

   ```json
   {
     "client_id": "YOUR_CLIENT_ID",
     "authority": "https://login.microsoftonline.com/common",
     "scopes": ["Chat.Read", "Chat.ReadBasic", "Chat.ReadWrite"],
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
