## Testing Guide for Teams Export

This document provides testing procedures for the `teams-export` CLI tool, including unit tests, integration tests, and manual verification steps.

---

## Unit Tests (Offline)

The project includes comprehensive unit tests that run without requiring Microsoft Graph credentials. These tests use mocked Graph API responses and fixtures to verify core functionality.

### Running Unit Tests

**Prerequisites:**
```bash
pip install -e ".[dev]"  # Install with dev dependencies (pytest, pytest-cov, pytest-mock)
```

**Run all tests:**
```bash
pytest tests/
```

**Run with coverage report:**
```bash
pytest --cov=src/teams_export --cov-report=term-missing --cov-report=html tests/
```

**Run specific test modules:**
```bash
# Test date parsing
pytest tests/test_dates.py -v

# Test configuration loading
pytest tests/test_config.py -v

# Test message export logic
pytest tests/test_exporter.py -v
```

**View coverage HTML report:**
```bash
open htmlcov/index.html  # macOS
xdg-open htmlcov/index.html  # Linux
start htmlcov/index.html  # Windows
```

### Coverage Targets

The test suite aims for **≥60% code coverage** across:
- `dates.py`: Date parsing, keyword resolution, range validation
- `config.py`: Configuration loading, environment overrides, defaults
- `exporter.py`: Message transformation, filtering, attachment extraction, chat selection
- `delta.py`: Delta state persistence and management

---

## Integration Tests (Requires Graph Credentials)

Integration tests verify the tool works with real Microsoft Graph API data. These require valid Azure AD credentials and access to a Teams account with chat history.

### Prerequisites

1. **Azure AD App Registration:**
   - Follow setup steps in [README.md](README.md#setup)
   - Ensure you have `Chat.Read` and `Chat.ReadBasic` permissions
   - Grant admin consent if required by your tenant

2. **Configuration:**
   ```bash
   # Create config file
   cp config.sample.json ~/.teams-exporter/config.json

   # Edit with your client ID and tenant ID
   nano ~/.teams-exporter/config.json
   ```

3. **Authentication:**
   ```bash
   # Test authentication
   teams-export --list
   ```
   - Follow device code flow prompts
   - Verify you can see your chats listed

### Manual Integration Tests

#### Test 1: Basic Export (Today's Messages)

**Objective:** Verify basic export workflow with default settings.

```bash
teams-export --user "your.colleague@company.com"
```

**Expected Results:**
- Authentication succeeds (device code flow or cached token)
- Chat list loads (from cache or API)
- Export creates `exports/colleague_name_YYYY-MM-DD.md`
- File contains messages from last 24 hours
- Markdown formatting is correct

---

#### Test 2: Export with Date Range

**Objective:** Test date range filtering and historical message retrieval.

```bash
teams-export --user "your.colleague@company.com" --from "2025-10-01" --to "2025-10-31"
```

**Expected Results:**
- Export file named `colleague_name_2025-10-01_2025-10-31.md`
- Only messages within October 2025 are included
- Messages sorted oldest → newest
- Timestamps match expected range

---

#### Test 3: Image Download and Embedding

**Objective:** Verify image attachments are downloaded and referenced correctly.

**Target Chat:** Any chat with image attachments

```bash
teams-export --user "user.with.images@company.com" --format jira
```

**Expected Results:**
- `_files/` directory created next to export file
- Images downloaded with correct extensions (`.png`, `.jpg`, etc.)
- Markdown contains `![filename](colleague_name_YYYY-MM-DD_files/image.png)` references
- Images viewable when Markdown is rendered

**Verification:**
```bash
ls -lh exports/colleague_name_YYYY-MM-DD_files/
# Should show downloaded images

cat exports/colleague_name_YYYY-MM-DD.md | grep '!\['
# Should show Markdown image references
```

---

#### Test 4: Download All Attachment Types (PDFs, Office Docs)

**Objective:** Verify all file types (not just images) are downloaded with `--download-all-attachments`.

**Target Chat:** "Security Pipelines" chat (or any chat with PDFs/Office documents)

**Test Case A: Specific date with PDF (Security Pipelines chat)**
```bash
teams-export --chat "Security Pipelines" --from "2025-10-13" --to "2025-10-13" --download-all-attachments
```

**Expected Results:**
- PDF file from 13/10 09:38 is downloaded
- File saved in `_files/` directory with `.pdf` extension
- Markdown/HTML/DOCX export references the local file path
- File is viewable (not corrupted)

**Test Case B: Mixed attachment types**
```bash
teams-export --chat "Project Alpha" --from "2025-01-01" --to "2025-12-31" --download-all-attachments --format docx
```

**Expected Results:**
- All attachment types downloaded: images, PDFs, `.docx`, `.xlsx`, etc.
- Files named based on original attachment names
- DOCX export embeds images and links to other files
- No download errors or failed attachments

**Verification:**
```bash
# Check file types downloaded
ls -lh exports/project_alpha_*_files/
file exports/project_alpha_*_files/*  # Verify MIME types

# Verify PDF is readable
open exports/project_alpha_*_files/*.pdf  # macOS
xdg-open exports/project_alpha_*_files/*.pdf  # Linux
```

---

#### Test 5: Incremental Export (Delta Sync)

**Objective:** Verify delta sync correctly fetches only new messages.

**Step 1: Baseline export**
```bash
teams-export --user "active.colleague@company.com" --incremental --from "2025-01-01" --to "2025-12-31"
```

**Expected Results:**
- First run fetches all messages in range
- Delta state saved in `~/.teams-exporter/delta/`
- Export completes successfully

**Step 2: Immediate re-export (should be fast)**
```bash
teams-export --user "active.colleague@company.com" --incremental --from "2025-01-01" --to "2025-12-31"
```

**Expected Results:**
- Export completes in <5 seconds (vs. minutes for full sync)
- Output message: "No new messages for chat..." or message count matches new messages only
- Delta state updated with new delta token

**Step 3: Verify delta state**
```bash
ls -lh ~/.teams-exporter/delta/
cat ~/.teams-exporter/delta/chat_*.json
```

**Expected Delta State Fields:**
- `chat_id`: Chat identifier
- `delta_token`: Non-null delta token from Graph API
- `delta_link`: Full delta URL
- `last_sync_timestamp`: Recent Unix timestamp
- `last_message_id`: ID of most recent message
- `message_count`: Number of messages processed

**Step 4: Clear delta state and verify full re-sync**
```bash
rm -rf ~/.teams-exporter/delta/
teams-export --user "active.colleague@company.com" --incremental --from "2025-01-01" --to "2025-12-31"
```

**Expected Results:**
- Export takes longer (full message fetch)
- All messages re-fetched
- New delta state created

---

#### Test 6: Concurrent Attachment Downloads (Performance)

**Objective:** Verify attachments download concurrently (not sequentially).

**Target Chat:** Chat with 10+ attachments

```bash
time teams-export --chat "Design Reviews" --from "2025-09-01" --to "2025-12-31" --download-all-attachments --sequential-fetch

# (Advanced) The experimental --parallel-fetch flag currently hangs for some chats; only flip
# it on when testing fixes for the “Faster Message Pagination” TODO item.
```

**Expected Results:**
- Progress output shows concurrent downloads (e.g., "[3/15] Downloaded: file3.png")
- Total time significantly less than sequential (roughly 3-5x faster for 15+ files)
- All files downloaded successfully
- No file corruption or naming conflicts

**Performance Benchmark:**
- **Sequential (legacy):** ~2 seconds per file × 15 files = 30 seconds
- **Concurrent (new):** ~2 seconds per batch × 3 batches = 6-10 seconds

---

#### Test 7: HTML Format with Embedded Images

**Objective:** Verify HTML export embeds images as base64 data URLs.

```bash
teams-export --user "designer@company.com" --format html --from "2025-10-01" --to "2025-10-31"
```

**Expected Results:**
- HTML file created with embedded CSS
- Images converted to base64 `data:` URLs
- "Copy to Clipboard" button present and functional
- HTML renders correctly in browser
- Pasting into Jira/Confluence preserves images

**Verification:**
```bash
# Check for base64 images
grep 'data:image' exports/designer_*.html

# Open in browser
open exports/designer_*.html  # macOS
```

---

#### Test 8: DOCX Export with Attachments

**Objective:** Verify Word document export with embedded images.

```bash
teams-export --chat "Marketing Campaign" --format docx --from "2025-10-01" --to "2025-10-31"
```

**Expected Results:**
- `.docx` file created
- Images embedded directly in document (not linked)
- Document opens in Microsoft Word or LibreOffice
- Chat title, participants, and date range in header
- Messages formatted with sender names in blue
- Timestamps and reactions present

**Verification:**
```bash
# Open in Word
open exports/marketing_campaign_*.docx  # macOS
libreoffice exports/marketing_campaign_*.docx  # Linux
```

---

#### Test 9: Parallel Export (--all flag)

**Objective:** Verify parallel export of multiple chats.

```bash
teams-export --all --from "last week" --format jira
```

**Expected Results:**
- Progress shows parallel execution: `"[1/25] Exported 42 messages from..."`
- Maximum 3 concurrent exports (prevents API throttling)
- All chats exported successfully
- No race conditions or file naming conflicts

**Verification:**
```bash
ls -lh exports/
# Should show multiple export files with recent timestamps
```

---

#### Test 10: Error Handling and Retry Logic

**Objective:** Verify graceful handling of API errors.

**Simulate Rate Limiting (if possible):**
- Export a large chat during peak usage
- Monitor console output for retry messages

**Expected Retry Behavior:**
- HTTP 429 (rate limit): Waits based on `Retry-After` header
- HTTP 5xx (server error): Exponential backoff (2s, 4s, 8s, 16s)
- Network errors: Automatic retry up to 4 attempts
- Final failure: Clear error message, non-zero exit code

**Example Output:**
```
Rate limited. Waiting 30s before retry 1/4...
Server error 503. Retrying in 2s...
```

---

## Security Pipelines Chat Test (Specific)

**Background:** The "Security Pipelines" chat contains a PDF file shared on 13/10 at 09:38. This is a key integration test for verifying PDF download functionality.

### Test Procedure

**Step 1: Verify chat exists**
```bash
teams-export --list | grep -i "security"
```

**Step 2: Export with PDF download**
```bash
teams-export --chat "Security Pipelines" --from "2025-10-13" --to "2025-10-13" --download-all-attachments --format jira
```

**Step 3: Verify PDF downloaded**
```bash
# Check files directory
ls -lh exports/security_pipelines_2025-10-13_files/

# Verify it's actually a PDF
file exports/security_pipelines_2025-10-13_files/*.pdf

# Expected output: "PDF document, version X.X"
```

**Step 4: Verify PDF is readable**
```bash
# Open PDF
open exports/security_pipelines_2025-10-13_files/*.pdf  # macOS
xdg-open exports/security_pipelines_2025-10-13_files/*.pdf  # Linux

# Or use pdfinfo to verify metadata
pdfinfo exports/security_pipelines_2025-10-13_files/*.pdf
```

**Step 5: Verify Markdown references PDF**
```bash
grep '\.pdf' exports/security_pipelines_2025-10-13.md

# Expected: Link or reference to the PDF file
```

### Expected Results

- PDF file downloaded successfully
- File size > 0 bytes
- MIME type is `application/pdf`
- PDF opens without corruption
- Markdown export contains reference to local PDF path
- Timestamp matches 09:38 UTC on 2025-10-13

### Troubleshooting

**PDF not downloaded:**
1. Verify `--download-all-attachments` flag is used
2. Check attachment URL in Graph API response is valid
3. Verify network connectivity and authentication

**PDF corrupted:**
1. Check Content-Type header from Graph API
2. Verify file extension logic in `_get_extension_from_mime()`
3. Try re-downloading with `--force-login` to refresh credentials

**PDF not in Markdown:**
1. Verify PDF is within date range filter
2. Check message timestamp vs. filter range
3. Ensure attachment is correctly parsed from Graph API response

---

## Automated Integration Tests (Future Work)

For CI/CD integration, consider adding:

1. **Fixture-based Graph Mocks:**
   - Record real Graph API responses as JSON fixtures
   - Replay during tests to avoid live API calls
   - Use `pytest-vcr` or similar for HTTP recording

2. **Test Chat Setup:**
   - Create a dedicated test chat in Teams
   - Pre-populate with known messages and attachments
   - Run automated exports and verify output

3. **Snapshot Testing:**
   - Compare exported files against known-good snapshots
   - Detect regressions in formatting or content

---

## Reporting Issues

When reporting bugs, include:
1. **Command used:** Full `teams-export` command with all flags
2. **Expected behavior:** What should happen
3. **Actual behavior:** What actually happened (include error messages)
4. **Environment:**
   - OS and version (macOS 14, Ubuntu 22.04, etc.)
   - Python version (`python --version`)
   - Package version (`pip show teams-export`)
5. **Logs:** Any error output from the CLI
6. **Reproducibility:** Steps to reproduce the issue

---

## Coverage Summary

**Current Test Coverage (Estimated):**
- `dates.py`: 95% (23 tests)
- `config.py`: 90% (18 tests)
- `exporter.py`: 75% (25 tests)
- `delta.py`: 80% (planned)
- `graph.py`: 60% (integration tests needed)
- `formatters.py`: 50% (manual testing primary)

**Target:** ≥60% overall coverage with comprehensive unit tests + manual integration verification.
