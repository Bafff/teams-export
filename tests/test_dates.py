"""Tests for date parsing functionality."""

from datetime import datetime, timedelta
import pytest
from freezegun import freeze_time

from teams_export.dates import resolve_range, DateParseError


class TestDateParsing:
    """Test date parsing and range resolution."""

    @freeze_time("2025-01-15 12:00:00")
    def test_today_keyword(self):
        """Test 'today' keyword parsing."""
        start, end = resolve_range("today", "today")
        assert start.date() == datetime(2025, 1, 15).date()
        assert end.date() == datetime(2025, 1, 15).date()
        assert start < end  # End should be end of day

    @freeze_time("2025-01-15 12:00:00")
    def test_yesterday_keyword(self):
        """Test 'yesterday' keyword parsing."""
        start, end = resolve_range("yesterday", "today")
        assert start.date() == datetime(2025, 1, 14).date()
        assert end.date() == datetime(2025, 1, 15).date()

    @freeze_time("2025-01-15 12:00:00")
    def test_last_week_keyword(self):
        """Test 'last week' keyword parsing."""
        start, end = resolve_range("last week", "today")
        assert start.date() == datetime(2025, 1, 8).date()
        assert end.date() == datetime(2025, 1, 15).date()

    @freeze_time("2025-01-15 12:00:00")
    def test_last_month_keyword(self):
        """Test 'last month' keyword parsing."""
        start, end = resolve_range("last month", "today")
        # Should be approximately 30 days ago
        expected_start = datetime(2025, 1, 15) - timedelta(days=30)
        assert abs((start.date() - expected_start.date()).days) <= 1

    def test_iso_date_format(self):
        """Test ISO date format parsing."""
        start, end = resolve_range("2025-01-01", "2025-01-31")
        assert start.date() == datetime(2025, 1, 1).date()
        assert end.date() == datetime(2025, 1, 31).date()

    def test_flexible_date_formats(self):
        """Test various date formats."""
        # These should all parse successfully
        formats = [
            ("Jan 1, 2025", "January 31, 2025"),
            ("1/1/2025", "1/31/2025"),
            ("2025-01-01", "2025-01-31"),
        ]
        for start_str, end_str in formats:
            start, end = resolve_range(start_str, end_str)
            assert start.month == 1
            assert end.month == 1
            assert start.year == 2025
            assert end.year == 2025

    def test_none_defaults(self):
        """Test default behavior with None values."""
        # Both None should default to last 24 hours
        start, end = resolve_range(None, None)
        delta = end - start
        assert delta.days <= 1

        # Only end None should use today
        start, end = resolve_range("2025-01-01", None)
        assert start.date() == datetime(2025, 1, 1).date()
        assert end.date() >= datetime.now().date()

    def test_invalid_date_raises_error(self):
        """Test that invalid dates raise DateParseError."""
        with pytest.raises(DateParseError):
            resolve_range("not-a-date", "today")

        with pytest.raises(DateParseError):
            resolve_range("2025-13-45", "today")  # Invalid date

    def test_start_after_end_raises_error(self):
        """Test that start date after end date raises error."""
        with pytest.raises(DateParseError, match="Start date must be before end date"):
            resolve_range("2025-12-31", "2025-01-01")

    def test_timezone_awareness(self):
        """Test that returned datetimes are timezone-aware."""
        start, end = resolve_range("2025-01-01", "2025-01-31")
        assert start.tzinfo is not None
        assert end.tzinfo is not None