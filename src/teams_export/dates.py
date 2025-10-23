from __future__ import annotations

import datetime as dt
from typing import Tuple

from dateutil import parser

UTC = dt.timezone.utc


class DateParseError(ValueError):
    """Raised when a provided date string cannot be interpreted."""


def _keyword_date(value: str, today: dt.date) -> Tuple[dt.date, bool]:
    lowered = value.lower()
    if lowered == "today":
        return today, False
    if lowered == "yesterday":
        return today - dt.timedelta(days=1), False
    if lowered == "last week":
        return today - dt.timedelta(days=7), True
    if lowered == "last month":
        return today - dt.timedelta(days=30), True
    raise DateParseError(f"Unsupported relative date: {value}")


def _parse_date(value: str) -> Tuple[dt.date, bool]:
    today = dt.datetime.now(UTC).date()
    try:
        return _keyword_date(value, today)
    except DateParseError:
        pass

    try:
        parsed = parser.isoparse(value).date()
    except (ValueError, TypeError):
        try:
            parsed = parser.parse(value).date()
        except (ValueError, TypeError) as exc:  # pragma: no cover - defensive
            raise DateParseError(f"Could not parse date value: {value}") from exc
    return parsed, False


def resolve_range(
    start_value: str | None,
    end_value: str | None,
) -> tuple[dt.datetime, dt.datetime]:
    """Convert CLI inputs into an inclusive UTC datetime window."""

    today = dt.datetime.now(UTC).date()
    if start_value:
        start_date, span_to_today = _parse_date(start_value)
    else:
        start_date, span_to_today = today, False

    if end_value:
        end_date, _ = _parse_date(end_value)
    elif span_to_today:
        end_date = today
    else:
        end_date = start_date

    if end_date < start_date:
        raise DateParseError("End date precedes start date.")

    start_dt = dt.datetime.combine(start_date, dt.time.min, tzinfo=UTC)
    # Graph filter is inclusive, so clamp to final second of the day.
    end_dt = dt.datetime.combine(end_date, dt.time.min, tzinfo=UTC) + dt.timedelta(days=1) - dt.timedelta(seconds=1)
    return start_dt, end_dt


def to_iso(dt_obj: dt.datetime) -> str:
    """Format datetime in RFC3339 UTC format."""

    utc_dt = dt_obj.astimezone(UTC)
    return utc_dt.isoformat().replace("+00:00", "Z")
