# Integration Testing Guide for Teams Export

This document describes how to test the teams-export tool with real Microsoft Teams data, specifically using the "Security Pipelines" chat for comprehensive testing.

## Prerequisites

1. **Microsoft Teams Account**: You need access to a Microsoft Teams account with existing chats.
2. **Azure AD App Registration**: The app must be registered and configured as described in README.md.
3. **Graph API Permissions**: Ensure you have `Chat.Read` and `Chat.ReadBasic` permissions granted.
4. **Test Chat**: Access to the "Security Pipelines" chat or another chat with diverse content including:
   - Text messages
   - Images
   - PDF files
   - Other document attachments

## Test Scenarios

### 1. Basic Export with Images

Test that the tool correctly exports messages and downloads inline images:

```bash
# Export with default image downloads
teams-export --user "security.pipelines@company.com" --from "2024-10-13" --to "2024-10-13"

# Verify:
# - Messages from October 13 are exported
# - Image at 09:38 is downloaded to exports/*_files/ directory
# - Markdown file contains correct image references
```

### 2. Full Attachment Download

Test downloading all attachment types including the PDF mentioned at 13/10 09:38:

```bash
# Export with all attachments
teams-export --user "security.pipelines@company.com" \
  --from "2024-10-13" --to "2024-10-13" \
  --download-all-attachments

# Verify:
# - PDF file from 09:38 is downloaded
# - All images are downloaded
# - Other document types are downloaded with correct extensions
# - Attachment directory structure is created (chat_name_date_files/)
```

### 3. Incremental Sync Testing

Test the delta sync functionality:

```bash
# Initial full export
teams-export --user "security.pipelines@company.com" \
  --from "2024-10-01" --to "2024-10-31" \
  --incremental

# Verify initial state saved
ls ~/.teams-exporter/delta/

# Run again - should detect no changes
teams-export --user "security.pipelines@company.com" \
  --from "2024-10-01" --to "2024-10-31" \
  --incremental

# Output should indicate "No new messages since last sync"

# Send a new message in the chat, then run again
teams-export --user "security.pipelines@company.com" \
  --from "2024-10-01" --to "2024-10-31" \
  --incremental

# Verify only new messages are fetched and exported
```

### 4. Export Formats Testing

Test all supported export formats:

```bash
# Markdown/Jira format (default)
teams-export --user "security.pipelines@company.com" --format jira

# HTML with embedded images
teams-export --user "security.pipelines@company.com" --format html

# Word document with embedded images
teams-export --user "security.pipelines@company.com" --format docx

# JSON raw data
teams-export --user "security.pipelines@company.com" --format json

# CSV for analysis
teams-export --user "security.pipelines@company.com" --format csv
```

Verify each format:
- **Markdown**: Proper formatting, clickable attachment links
- **HTML**: Images embedded as base64, can be pasted into Jira/Confluence
- **DOCX**: Opens in Word, images embedded, formatting preserved
- **JSON**: Valid JSON structure with all message fields
- **CSV**: Opens in Excel, proper column separation

### 5. Performance Testing

Test parallel fetching and downloading:

```bash
# Export a large chat with many messages and attachments
teams-export --user "security.pipelines@company.com" \
  --from "2024-01-01" --to "2024-12-31" \
  --download-all-attachments

# Monitor:
# - Parallel message pagination (should see "Fetching messages with parallel pagination...")
# - Parallel attachment downloads (shows progress with concurrent workers)
# - Overall export time vs sequential processing
```

### 6. Error Handling

Test various error scenarios:

```bash
# Non-existent user
teams-export --user "nonexistent@company.com"
# Should show "No chat matches the provided identifiers"

# Invalid date range
teams-export --user "security.pipelines@company.com" \
  --from "2024-12-31" --to "2024-01-01"
# Should show "Start date must be before end date"

# Network interruption
# Start export, disconnect network mid-download
# Should retry with exponential backoff

# API rate limiting
teams-export --all --download-all-attachments
# Should handle 429 responses gracefully with retry
```

## Manual Verification Checklist

After running the tests, manually verify:

### Message Content
- [ ] All messages within date range are exported
- [ ] Message timestamps are correct
- [ ] Sender information is accurate
- [ ] Message formatting is preserved
- [ ] Reactions are displayed
- [ ] Mentions are included

### Attachments
- [ ] Images are downloaded and viewable
- [ ] PDF files open correctly
- [ ] Document filenames are sanitized properly
- [ ] No duplicate downloads
- [ ] Attachment URLs in export link to local files
- [ ] Failed downloads are logged but don't stop export

### Incremental Sync
- [ ] Delta state files created in ~/.teams-exporter/delta/
- [ ] Second run with no changes is fast (skips export)
- [ ] New messages trigger incremental update
- [ ] Delta tokens are properly saved and reused

### Performance
- [ ] Parallel pagination faster than sequential for large chats
- [ ] Multiple attachments download concurrently
- [ ] Progress indicators show accurate information
- [ ] Memory usage remains reasonable for large exports

## Automated Test Execution

Run the test suite with pytest:

```bash
# Install test dependencies
pip install -e ".[dev]"

# Run all tests
pytest

# Run with coverage report
pytest --cov=src/teams_export --cov-report=html

# Run specific test modules
pytest tests/test_exporter.py
pytest tests/test_delta.py

# Run tests matching a pattern
pytest -k "incremental"

# Verbose output
pytest -vv
```

## Integration Test Script

For automated integration testing, use this script:

```bash
#!/bin/bash
# integration_test.sh

set -e

echo "Teams Export Integration Test Suite"
echo "===================================="

# Configuration
TEST_USER="security.pipelines@company.com"
TEST_DATE="2024-10-13"
OUTPUT_DIR="test_exports"

# Clean previous test outputs
rm -rf $OUTPUT_DIR
mkdir -p $OUTPUT_DIR

echo "Test 1: Basic export with images..."
teams-export --user "$TEST_USER" \
  --from "$TEST_DATE" --to "$TEST_DATE" \
  --output-dir "$OUTPUT_DIR/test1" \
  --format jira

echo "Test 2: Export with all attachments..."
teams-export --user "$TEST_USER" \
  --from "$TEST_DATE" --to "$TEST_DATE" \
  --output-dir "$OUTPUT_DIR/test2" \
  --download-all-attachments

echo "Test 3: Incremental export (first run)..."
teams-export --user "$TEST_USER" \
  --from "$TEST_DATE" --to "$TEST_DATE" \
  --output-dir "$OUTPUT_DIR/test3" \
  --incremental

echo "Test 4: Incremental export (second run - no changes)..."
teams-export --user "$TEST_USER" \
  --from "$TEST_DATE" --to "$TEST_DATE" \
  --output-dir "$OUTPUT_DIR/test4" \
  --incremental

echo "Test 5: All export formats..."
for FORMAT in jira html docx json csv; do
  echo "  Testing format: $FORMAT"
  teams-export --user "$TEST_USER" \
    --from "$TEST_DATE" --to "$TEST_DATE" \
    --output-dir "$OUTPUT_DIR/test5_$FORMAT" \
    --format "$FORMAT" \
    --no-download-attachments
done

echo "Test 6: Parallel export of multiple chats..."
teams-export --all \
  --from "$TEST_DATE" --to "$TEST_DATE" \
  --output-dir "$OUTPUT_DIR/test6" \
  --no-download-attachments

echo "===================================="
echo "Integration tests completed!"
echo "Check $OUTPUT_DIR for results"

# Verify outputs exist
echo ""
echo "Verification:"
find $OUTPUT_DIR -type f -name "*.md" | head -5
find $OUTPUT_DIR -type f -name "*.pdf" | head -5
find $OUTPUT_DIR -type f -name "*.png" | head -5
find $OUTPUT_DIR -type f -name "*.jpg" | head -5
```

## Troubleshooting

### Common Issues

1. **Authentication fails**:
   - Run with `--force-login` to refresh token
   - Check Azure AD app permissions
   - Verify client ID in config

2. **No chats found**:
   - Verify user has access to Teams chats
   - Check Graph API permissions are granted
   - Try `--refresh-cache` to reload chat list

3. **Attachments not downloading**:
   - Check network connectivity
   - Verify attachment URLs are accessible
   - Check disk space for downloads
   - Review error messages for specific failures

4. **Incremental sync not working**:
   - Check ~/.teams-exporter/delta/ directory exists
   - Verify write permissions
   - Clear delta state with `rm -rf ~/.teams-exporter/delta/`
   - Check Graph API supports delta queries for your tenant

### Debug Mode

For detailed debugging information:

```bash
# Enable verbose logging (if implemented)
TEAMS_EXPORT_DEBUG=1 teams-export --user "test@company.com"

# Check API responses
# Monitor network traffic with proxy tools like Charles or Fiddler

# Inspect cache and state files
cat ~/.teams-exporter/cache/chats_cache.json | jq .
cat ~/.teams-exporter/delta/*.json | jq .
```

## Reporting Issues

When reporting issues with integration tests:

1. Include the exact command that failed
2. Provide relevant error messages
3. Specify the chat characteristics (number of messages, attachment types)
4. Include teams-export version (`pip show teams-export`)
5. Attach relevant log output
6. Describe expected vs actual behavior

## Security Considerations for Testing

- Never commit real chat data to version control
- Sanitize any screenshots or logs before sharing
- Use test accounts when possible
- Clear cache and delta state after testing sensitive data
- Review exported files before sharing to ensure no PII/sensitive data

## Continuous Integration

For CI/CD pipelines, consider:

1. Using mock Graph API responses for unit tests
2. Creating dedicated test chats with known content
3. Running integration tests in isolated environments
4. Automating cleanup of test data
5. Setting up test result notifications

---

This guide should be updated as new features are added or test scenarios are identified.