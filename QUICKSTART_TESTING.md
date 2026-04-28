# Quick Start Testing Guide

This guide provides the **fastest path** to verify all new features work correctly.

---

## Prerequisites

```bash
# 1. Install with dev dependencies
pip install -e ".[dev]"

# 2. Configure Azure AD app (if not already done)
cp config.sample.json ~/.teams-exporter/config.json
# Edit ~/.teams-exporter/config.json with your client_id and tenant_id
```

---

## Step 1: Run Unit Tests (Offline, No Credentials)

```bash
# Run all tests
pytest tests/ -v

# Expected output:
# ==================== test session starts ====================
# collected 66 items
#
# tests/test_dates.py::TestKeywordDate::test_today PASSED
# tests/test_dates.py::TestKeywordDate::test_yesterday PASSED
# ...
# tests/test_exporter.py::TestExportIntegration::test_export_filters_by_date_range PASSED
#
# ==================== 66 passed in 2.34s ====================

# Check coverage
pytest --cov=src/teams_export --cov-report=term-missing tests/

# Expected coverage: ≥60% overall
```

**If tests pass:** ✅ Core logic is working correctly

---

## Step 2: Verify Authentication (Integration Test)

```bash
# List your chats (will trigger device code flow if needed)
teams-export --list

# Expected output:
# To sign in, use a web browser to open the page https://microsoft.com/devicelogin
# and enter the code XXXXXXXX to authenticate.
#
# [After authentication]
# ✓ Loaded 234 chats
# 19:abc123@thread.v2	group	Project Alpha	Alice Smith, Bob Jones, Charlie Brown
# ...
```

**If authentication succeeds:** ✅ Azure AD app is configured correctly

---

## Step 3: Test Incremental Export (Delta Sync)

```bash
# First run: baseline (will be slow)
time teams-export --user "your.colleague@company.com" --incremental --from "2025-01-01" --to "2025-12-31"

# Note the time (e.g., 30 seconds for a large chat)

# Second run: delta sync (should be MUCH faster)
time teams-export --user "your.colleague@company.com" --incremental --from "2025-01-01" --to "2025-12-31"

# Expected time: 5-10 seconds (5-10x faster!)
# Expected output: "No new messages for chat..." OR export with only new messages
```

**Verify delta state:**
```bash
ls -lh ~/.teams-exporter/delta/
cat ~/.teams-exporter/delta/chat_*.json
```

**Expected delta state fields:**
- `delta_token`: Non-null
- `delta_link`: Full Graph API URL
- `last_sync_timestamp`: Recent timestamp
- `message_count`: Number of messages processed

**If second run is much faster:** ✅ Delta sync is working

---

## Step 4: Test PDF Download (Security Pipelines Chat)

**Target:** PDF file from 13/10 09:38 in "Security Pipelines" chat

```bash
# Export with PDF download
teams-export --chat "Security Pipelines" --from "2025-10-13" --to "2025-10-13" --download-all-attachments

# Verify PDF downloaded
ls -lh exports/security_pipelines_2025-10-13_files/

# Check file type
file exports/security_pipelines_2025-10-13_files/*.pdf
# Expected: "PDF document, version X.X"

# Open PDF to verify it's readable
open exports/security_pipelines_2025-10-13_files/*.pdf  # macOS
xdg-open exports/security_pipelines_2025-10-13_files/*.pdf  # Linux

# Verify Markdown references PDF
grep '.pdf' exports/security_pipelines_2025-10-13.md
```

**Expected results:**
- PDF file exists in `_files/` directory
- File size > 0 bytes
- MIME type is `application/pdf`
- PDF opens without corruption
- Markdown contains reference to local PDF path

**If PDF downloads and opens correctly:** ✅ All attachment types work

---

## Step 5: Test Concurrent Downloads (Performance)

**Target:** Chat with 10+ attachments

```bash
# Export with all attachments (timed) — sequential is the reliable default
time teams-export --chat "Design Reviews" --download-all-attachments --from "2025-09-01" --to "2025-12-31" --sequential-fetch

# Optional: --parallel-fetch is currently unstable and only for developers working on the
# pagination fix. If you enable it and the CLI hangs, drop back to sequential.
```

**Verify all files downloaded:**
```bash
ls -lh exports/design_reviews_*_files/
# Should show 15 files with correct extensions
```

**If downloads complete in ~6-10s:** ✅ Concurrent downloads working

---

## Step 6: Test All Attachment Types

```bash
# Export a chat with mixed attachments
teams-export --chat "Project Planning" --download-all-attachments --from "2025-01-01" --to "2025-12-31"

# Check downloaded file types
ls -lh exports/project_planning_*_files/
file exports/project_planning_*_files/*

# Expected file types:
# - *.png, *.jpg, *.gif (images)
# - *.pdf (PDFs)
# - *.docx, *.xlsx, *.pptx (Office docs)
# - *.zip (archives)
# - *.txt, *.json (text files)
```

**If all file types download successfully:** ✅ Comprehensive attachment handling works

---

## Step 7: Verify HTML Export with Embedded Images

```bash
# Export to HTML format
teams-export --user "designer@company.com" --format html --from "2025-10-01" --to "2025-10-31"

# Check for base64-embedded images
grep 'data:image' exports/designer_*.html | wc -l
# Should show count of embedded images

# Open in browser
open exports/designer_*.html  # macOS

# Verify:
# - Images render correctly
# - "Copy to Clipboard" button present
# - Dark theme applied
# - Messages formatted in chat bubbles
```

**If HTML renders correctly with images:** ✅ HTML export works

---

## Step 8: Verify DOCX Export

```bash
# Export to Word format
teams-export --chat "Marketing Campaign" --format docx --from "2025-10-01" --to "2025-10-31"

# Open in Word/LibreOffice
open exports/marketing_campaign_*.docx  # macOS
libreoffice exports/marketing_campaign_*.docx  # Linux

# Verify:
# - Document opens without errors
# - Images embedded (not linked)
# - Chat title, participants, date range in header
# - Messages formatted with sender names in blue
# - Timestamps and reactions present
```

**If DOCX opens correctly with embedded images:** ✅ DOCX export works

---

## Step 9: Test Error Handling and Retry

**Verify retry logic on network errors:**

```bash
# Export a large chat during peak usage (may trigger rate limiting)
teams-export --chat "Large Team Chat" --download-all-attachments --from "2025-01-01" --to "2025-12-31"

# Watch for retry messages in console:
# Rate limited. Waiting 30s before retry 1/4...
# Server error 503. Retrying in 2s...

# Expected behavior:
# - Automatic retries on 429 (rate limit)
# - Exponential backoff on 5xx errors
# - Clear error messages on final failure
```

**If retries work correctly:** ✅ Error handling is robust

---

## Step 10: Quick Regression Test (All Features)

**Run this single command to test multiple features at once:**

```bash
teams-export \
  --user "active.colleague@company.com" \
  --from "2025-10-01" \
  --to "2025-10-31" \
  --format jira \
  --download-all-attachments \
  --incremental

# This tests:
# ✓ Date range parsing
# ✓ User-based chat selection
# ✓ Incremental export (delta sync)
# ✓ All attachment types download
# ✓ Concurrent downloads
# ✓ Markdown formatting
```

**Expected results:**
- Export completes successfully
- Markdown file created with correct date range in filename
- Attachments downloaded to `_files/` directory
- Second run with same command is much faster (delta sync)

---

## Troubleshooting

### Tests Fail

**Error:** `ModuleNotFoundError: No module named 'pytest'`

**Fix:**
```bash
pip install -e ".[dev]"
```

---

### Authentication Fails

**Error:** `ConfigError: Missing client_id`

**Fix:**
```bash
# Create config file
cp config.sample.json ~/.teams-exporter/config.json

# Edit with your Azure AD app details
nano ~/.teams-exporter/config.json
```

---

### PDF Not Downloaded

**Error:** PDF file missing from `_files/` directory

**Fix:**
1. Verify `--download-all-attachments` flag is used
2. Check date range includes the PDF message (13/10 09:38)
3. Try with `--force-login` to refresh credentials

---

### Incremental Export Not Faster

**Issue:** Second run takes same time as first run

**Fix:**
1. Check delta state exists: `ls ~/.teams-exporter/delta/`
2. Verify `--incremental` flag is used
3. Check for error messages in console output
4. Try clearing delta state: `rm -rf ~/.teams-exporter/delta/`

---

### Concurrent Downloads Not Working

**Issue:** Downloads appear sequential (slow)

**Fix:**
1. Verify you have 10+ attachments in the chat
2. Check console output for progress indicators
3. Network issues may limit concurrency

---

## Success Criteria

✅ **All features working if:**

1. **Unit tests pass:** 66+ tests, ≥60% coverage
2. **Authentication succeeds:** Can list chats and export messages
3. **Incremental export:** Second run 5-10x faster than first
4. **PDF downloads:** Security Pipelines chat PDF (13/10 09:38) downloads correctly
5. **Concurrent downloads:** 15 files download in ~6-10s (vs. ~30s sequential)
6. **All file types:** PDFs, Office docs, images all download successfully
7. **HTML export:** Renders correctly with embedded images
8. **DOCX export:** Opens in Word with embedded images
9. **Error handling:** Retries work on rate limits and server errors

---

## Quick Demo Script

**For demonstrating new features to stakeholders:**

```bash
# 1. Show incremental export speed
echo "=== First run (baseline) ==="
time teams-export --user "demo@company.com" --incremental

echo "=== Second run (delta sync) ==="
time teams-export --user "demo@company.com" --incremental
# Show dramatic speed improvement

# 2. Show all attachment types
echo "=== Downloading all attachment types ==="
teams-export --chat "Demo Chat" --download-all-attachments

echo "=== Downloaded files ==="
ls -lh exports/demo_chat_*_files/
file exports/demo_chat_*_files/*
# Show variety of file types

# 3. Show HTML export
echo "=== HTML export with embedded images ==="
teams-export --user "demo@company.com" --format html
open exports/demo_*.html
# Demonstrate copy-paste to Jira/Confluence

# 4. Show test coverage
echo "=== Test coverage report ==="
pytest --cov=src/teams_export --cov-report=term-missing tests/
```

---

## Next Steps After Testing

1. **Create pull request** with all changes
2. **Request code review** from team
3. **Update internal documentation** (Confluence, wiki)
4. **Announce new features** to users
5. **Monitor for issues** in production use

---

## Need Help?

- **Unit tests failing?** Check `tests/conftest.py` for fixture issues
- **Integration tests failing?** Verify Azure AD app permissions
- **PDF not downloading?** See `TESTING.md` for detailed troubleshooting
- **Performance issues?** Check network connectivity and API rate limits

**For detailed testing procedures, see:** [TESTING.md](TESTING.md)
**For implementation details, see:** [IMPLEMENTATION_NOTES.md](IMPLEMENTATION_NOTES.md)
