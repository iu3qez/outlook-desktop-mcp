"""Helpers for extracting and formatting Outlook item data."""
import re

from outlook_desktop_mcp.tools._folder_constants import (
    BUSY_STATUS_NAMES,
    MEETING_STATUS_NAMES,
    RESPONSE_NAMES,
    TASK_STATUS_NAMES,
    IMPORTANCE_NAMES,
)


def truncate(text: str, max_length: int = 2000) -> str:
    if len(text) <= max_length:
        return text
    return text[:max_length] + "\n... [truncated]"


def strip_html(html: str) -> str:
    text = re.sub(r"<[^>]+>", "", html)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def format_email_summary(item) -> dict:
    """Extract key fields from an Outlook MailItem into a dict."""
    return {
        "entry_id": item.EntryID,
        "subject": item.Subject or "(no subject)",
        "sender": getattr(item, "SenderEmailAddress", "unknown"),
        "sender_name": getattr(item, "SenderName", "unknown"),
        "received_time": str(item.ReceivedTime),
        "unread": bool(item.UnRead),
        "has_attachments": bool(item.Attachments.Count > 0),
        "attachment_count": item.Attachments.Count,
    }


def format_email_full(item, body_max_length: int = 5000) -> dict:
    """Extract full email details including body."""
    result = format_email_summary(item)
    result["to"] = item.To or ""
    result["cc"] = item.CC or ""
    result["body"] = truncate(item.Body or "", body_max_length)
    return result


# --- Calendar formatting ---


def format_event_summary(item) -> dict:
    """Extract key fields from an Outlook AppointmentItem."""
    return {
        "entry_id": item.EntryID,
        "subject": item.Subject or "(no subject)",
        "start": str(item.Start),
        "end": str(item.End),
        "duration": item.Duration,
        "location": item.Location or "",
        "organizer": item.Organizer or "",
        "is_recurring": bool(item.IsRecurring),
        "all_day": bool(item.AllDayEvent),
        "busy_status": BUSY_STATUS_NAMES.get(item.BusyStatus, "unknown"),
        "meeting_status": MEETING_STATUS_NAMES.get(item.MeetingStatus, "unknown"),
        "required_attendees": item.RequiredAttendees or "",
        "optional_attendees": item.OptionalAttendees or "",
    }


def format_event_full(item, body_max_length: int = 5000) -> dict:
    """Full event details including body."""
    result = format_event_summary(item)
    result["body"] = truncate(item.Body or "", body_max_length)
    result["reminder_set"] = bool(item.ReminderSet)
    result["reminder_minutes"] = (
        item.ReminderMinutesBeforeStart if item.ReminderSet else None
    )
    result["categories"] = item.Categories or ""
    result["response_status"] = RESPONSE_NAMES.get(item.ResponseStatus, "unknown")
    return result


# --- Task formatting ---


def outlook_date(value) -> str | None:
    """Stringify an Outlook date, mapping the "no date" sentinel (1/1/4501)
    to None. Checked by year because str() of a pywintypes datetime is
    ISO-like, not the M/D/YYYY form Outlook shows."""
    if value is None or getattr(value, "year", 4501) >= 4501:
        return None
    return str(value)


def format_task_summary(item) -> dict:
    """Extract key fields from an Outlook TaskItem."""
    return {
        "entry_id": item.EntryID,
        "subject": item.Subject or "(no subject)",
        "status": TASK_STATUS_NAMES.get(item.Status, "unknown"),
        "percent_complete": item.PercentComplete,
        "due_date": outlook_date(item.DueDate),
        "start_date": outlook_date(item.StartDate),
        "importance": IMPORTANCE_NAMES.get(item.Importance, "normal"),
        "complete": bool(item.Complete),
        "categories": item.Categories or "",
        "owner": item.Owner or "",
    }


def format_task_full(item, body_max_length: int = 5000) -> dict:
    """Full task details including body."""
    result = format_task_summary(item)
    result["body"] = truncate(item.Body or "", body_max_length)
    result["reminder_set"] = bool(item.ReminderSet)
    result["date_completed"] = (
        str(item.DateCompleted) if item.Complete else None
    )
    return result


# --- To-Do formatting ---

# olObjectClass values of the item kinds that show up in the To-Do List
_TODO_ITEM_TYPES = {48: "task", 43: "email", 40: "contact"}
_OL_CLASS_TASK = 48


def format_todo_item(item) -> dict:
    """Common view of a To-Do List entry: a TaskItem, or an email/contact
    flagged for follow-up (whose dates live in the Task* properties)."""
    item_class = item.Class
    is_task = item_class == _OL_CLASS_TASK
    result = {
        "entry_id": item.EntryID,
        "item_type": _TODO_ITEM_TYPES.get(item_class, "other"),
        "subject": item.Subject or "(no subject)",
        "due_date": outlook_date(item.DueDate if is_task else item.TaskDueDate),
        "start_date": outlook_date(item.StartDate if is_task else item.TaskStartDate),
        "importance": IMPORTANCE_NAMES.get(item.Importance, "normal"),
        "categories": item.Categories or "",
    }
    if is_task:
        result["complete"] = bool(item.Complete)
        result["status"] = TASK_STATUS_NAMES.get(item.Status, "unknown")
    else:
        result["complete"] = outlook_date(item.TaskCompletedDate) is not None
        result["flag_request"] = getattr(item, "FlagRequest", "") or ""
        if item_class == 43:
            result["sender_name"] = getattr(item, "SenderName", "") or ""
    return result
