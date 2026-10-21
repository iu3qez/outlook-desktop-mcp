"""
Unit test for utils.dates.restrict_date — locale-safe date literals for
Outlook Restrict()/DASL filters.

Pure unit test: no Outlook, no pywin32, runs on any platform. The Win32
GetDateFormatEx / GetTimeFormatEx calls are replaced with fakes that emulate
a day-first (DD/MM/YYYY) regional setting, so the test checks the
literal-building logic without a Windows machine.

Regression for the locale bug where a hard-coded strftime('%m/%d/%Y %H:%M')
literal such as '09/07/2026 20:15' was parsed by Outlook as 9 Jul instead of
7 Sep on non-US locales, silently returning wrong or empty filter results.

Adapted from the test in upstream PR #35.

Run: .venv\\Scripts\\python tests\\date_literal_test.py
"""
import ctypes
import os
import sys
from datetime import datetime, timedelta, timezone

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))

from outlook_desktop_mcp.utils import dates  # noqa: E402

# Record every SYSTEMTIME handed to the fake Win32 formatters and render a
# deterministic day-first date instead of calling the real API.
_formatted = []


class _FakeKernel32:
    @staticmethod
    def GetDateFormatEx(locale, flags, st_ref, fmt, buf, size, calendar):
        st = st_ref._obj
        _formatted.append(datetime(st.wYear, st.wMonth, st.wDay, st.wHour, st.wMinute))
        buf.value = f"{st.wDay:02d}/{st.wMonth:02d}/{st.wYear:04d}"
        return len(buf.value) + 1

    @staticmethod
    def GetTimeFormatEx(locale, flags, st_ref, fmt, buf, size):
        st = st_ref._obj
        buf.value = f"{st.wHour:02d}:{st.wMinute:02d}"
        return len(buf.value) + 1


class _FakeWindll:
    kernel32 = _FakeKernel32()


ctypes.windll = _FakeWindll()


def test_jet_literal_passes_local_wall_clock_through_unchanged():
    """Jet filters compare in local time: the datetime must reach the locale
    formatter with its wall-clock fields untouched (no tz shift)."""
    _formatted.clear()
    out = dates.restrict_date(datetime(2026, 9, 7, 20, 15))
    assert out == "07/09/2026 20:15", out
    assert _formatted == [datetime(2026, 9, 7, 20, 15)], _formatted


def test_seconds_are_dropped():
    """Outlook ignores seconds and misbehaves if the literal carries them."""
    _formatted.clear()
    dates.restrict_date(datetime(2026, 9, 7, 20, 15, 42))
    assert _formatted == [datetime(2026, 9, 7, 20, 15)], _formatted


def test_no_month_name_is_spliced_into_the_literal():
    """The builder must add no (English) month name of its own."""
    english_abbr = [
        "Jan", "Feb", "Mar", "Apr", "May", "Jun",
        "Jul", "Aug", "Sep", "Oct", "Nov", "Dec",
    ]
    for month in range(1, 13):
        out = dates.restrict_date(datetime(2026, month, 15, 12, 0))
        for abbr in english_abbr:
            assert abbr not in out, (month, out)


def test_dasl_literal_is_converted_to_utc():
    """DASL filters compare in UTC. An aware datetime at +02:00 must reach the
    formatter shifted back to UTC (18:15), while the Jet path keeps 20:15."""
    _formatted.clear()
    aware = datetime(2026, 9, 7, 20, 15, tzinfo=timezone(timedelta(hours=2)))

    assert dates.restrict_date(aware, utc=True) == "07/09/2026 18:15"
    assert _formatted[-1] == datetime(2026, 9, 7, 18, 15)

    assert dates.restrict_date(aware.replace(tzinfo=None)) == "07/09/2026 20:15"


def test_dasl_literal_leaves_utc_input_unchanged():
    """A datetime already in UTC is not shifted again."""
    _formatted.clear()
    aware_utc = datetime(2026, 9, 7, 20, 15, tzinfo=timezone.utc)
    assert dates.restrict_date(aware_utc, utc=True) == "07/09/2026 20:15"
    assert _formatted[-1] == datetime(2026, 9, 7, 20, 15)


if __name__ == "__main__":
    test_jet_literal_passes_local_wall_clock_through_unchanged()
    test_seconds_are_dropped()
    test_no_month_name_is_spliced_into_the_literal()
    test_dasl_literal_is_converted_to_utc()
    test_dasl_literal_leaves_utc_input_unchanged()
    print("OK")
