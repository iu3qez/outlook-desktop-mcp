"""
Unit test for utils.formatting.format_todo_item / outlook_date — the common
view list_todo builds over tasks and flagged emails/contacts.

Pure unit test: no Outlook, no pywin32, runs on any platform. Outlook items
are replaced with plain objects carrying the same property names.

Run: .venv\\Scripts\\python tests\\todo_format_test.py
"""
import os
import sys
from datetime import datetime
from types import SimpleNamespace

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))

from outlook_desktop_mcp.utils.formatting import (  # noqa: E402
    format_todo_item,
    outlook_date,
)

NO_DATE = datetime(4501, 1, 1)


def _task(**kw):
    base = dict(
        Class=48, EntryID="T1", Subject="Write report", DueDate=NO_DATE,
        StartDate=NO_DATE, Complete=False, Status=1, Importance=2,
        Categories="",
    )
    base.update(kw)
    return SimpleNamespace(**base)


def _mail(**kw):
    base = dict(
        Class=43, EntryID="M1", Subject="Invoice", TaskDueDate=NO_DATE,
        TaskStartDate=NO_DATE, TaskCompletedDate=NO_DATE, Importance=1,
        Categories="Red", FlagRequest="Follow up", SenderName="Alice",
    )
    base.update(kw)
    return SimpleNamespace(**base)


def test_no_date_sentinel_is_none():
    assert outlook_date(NO_DATE) is None
    assert outlook_date(None) is None
    assert outlook_date(datetime(2026, 3, 1)) == "2026-03-01 00:00:00"


def test_task_entry_uses_task_dates_and_status():
    entry = format_todo_item(_task(DueDate=datetime(2026, 3, 1)))
    assert entry["item_type"] == "task"
    assert entry["due_date"] == "2026-03-01 00:00:00"
    assert entry["start_date"] is None
    assert entry["complete"] is False
    assert entry["status"] == "in_progress"
    assert entry["importance"] == "high"
    assert "flag_request" not in entry


def test_flagged_email_uses_flag_dates():
    entry = format_todo_item(_mail(TaskDueDate=datetime(2026, 3, 2)))
    assert entry["item_type"] == "email"
    assert entry["due_date"] == "2026-03-02 00:00:00"
    assert entry["complete"] is False
    assert entry["flag_request"] == "Follow up"
    assert entry["sender_name"] == "Alice"
    assert "status" not in entry


def test_flagged_email_completed_when_completion_date_set():
    entry = format_todo_item(_mail(TaskCompletedDate=datetime(2026, 3, 3)))
    assert entry["complete"] is True


def test_flagged_contact_has_no_sender():
    contact = _mail(Class=40)
    del contact.SenderName
    entry = format_todo_item(contact)
    assert entry["item_type"] == "contact"
    assert "sender_name" not in entry


if __name__ == "__main__":
    test_no_date_sentinel_is_none()
    test_task_entry_uses_task_dates_and_status()
    test_flagged_email_uses_flag_dates()
    test_flagged_email_completed_when_completion_date_set()
    test_flagged_contact_has_no_sender()
    print("OK")
