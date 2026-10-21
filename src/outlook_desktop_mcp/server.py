"""
Outlook Desktop MCP Server
===========================
Exposes Microsoft Outlook Desktop (Classic) as an MCP server over stdio.
Uses COM automation — no Microsoft Graph, no Entra app registration.
Just run this on Windows with Outlook open and you have a full email MCP server.

Entry point: python -m outlook_desktop_mcp.server
"""
import sys
import json
import logging
import re

# When launched via `python -m outlook_desktop_mcp.server`, runpy loads this
# file as sys.modules["__main__"] but NOT as sys.modules["outlook_desktop_mcp.server"].
# Without this alias, tool modules that later do `from outlook_desktop_mcp
# import server` re-execute this file, creating a second `bridge = OutlookBridge()`
# distinct from the one main() actually started — every tool call then times
# out because no COM thread services the duplicate bridge's queue.
if __name__ == "__main__":
    sys.modules["outlook_desktop_mcp.server"] = sys.modules["__main__"]

from mcp.server.fastmcp import FastMCP

from outlook_desktop_mcp.com_bridge import OutlookBridge
from datetime import datetime, timedelta

import os

from outlook_desktop_mcp.tools._folder_constants import (
    FOLDER_NAME_TO_ENUM,
    OL_MAIL_ITEM,
    OL_APPOINTMENT_ITEM,
    OL_FOLDER_CALENDAR,
    OL_FOLDER_TASKS,
    OL_MEETING,
    OL_MEETING_CANCELED,
    OL_RESPONSE_TENTATIVE,
    OL_RESPONSE_ACCEPTED,
    OL_RESPONSE_DECLINED,
    OL_REQUIRED,
    OL_OPTIONAL,
    OL_TASK_ITEM,
    OL_TASK_COMPLETE,
    TASK_STATUS_NAMES,
    IMPORTANCE_NAMES,
)
from outlook_desktop_mcp.utils.formatting import (
    format_email_summary,
    format_email_full,
    format_event_summary,
    format_event_full,
    format_task_summary,
    format_task_full,
)
from outlook_desktop_mcp.utils.errors import format_com_error
from outlook_desktop_mcp.utils.dates import restrict_date

# --- Logging (all to stderr, stdout is reserved for MCP JSON-RPC) ---

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(name)s] %(levelname)s: %(message)s",
    stream=sys.stderr,
)
logger = logging.getLogger("outlook_desktop_mcp")


# --- Security helpers ---

def _safe_dasl(query: str) -> str:
    """Sanitize a string for use in a DASL LIKE filter value.
    Escapes SQL wildcards (% and _) so user input is treated as literals,
    then escapes quote characters required by DASL syntax.
    """
    query = query.replace("%", "[%]").replace("_", "[_]")
    return query.replace("'", "''").replace('"', '""')


def _raw_dasl(filter_str: str) -> str:
    """Normalize a caller-supplied DASL condition so it can be AND-ed with
    other conditions: strips an optional leading "@SQL=" and wraps it in
    parentheses. The condition is passed to Restrict verbatim — it is not
    sanitized, so the caller owns its syntax."""
    filter_str = filter_str.strip()
    if filter_str[:5].upper() == "@SQL=":
        filter_str = filter_str[5:].strip()
    return f"({filter_str})" if filter_str else ""


# Outlook item Class constants (olObjectClass — distinct from olItemType used in CreateItem)
_OL_CLASS_MAIL = 43
_OL_CLASS_APPOINTMENT = 26
_OL_CLASS_TASK = 48


def _check_item_class(item, expected_class: int, label: str) -> str | None:
    """Return an error string if item is the wrong type, else None."""
    if item.Class != expected_class:
        return f"Error: Entry ID does not refer to a {label}."
    return None


# --- MCP Server ---

mcp = FastMCP(
    "outlook-desktop-mcp",
    instructions=(
        "This MCP server gives you full access to Microsoft Outlook Desktop on "
        "Windows via COM automation. It can send emails, read inbox messages, "
        "search across folders, mark messages as read/unread, move messages "
        "between folders (including archive), reply to emails, and list the "
        "complete folder hierarchy.\n\n"
        "All operations use the locally authenticated Outlook profile — no "
        "Microsoft Graph API, no Entra app registration, no OAuth tokens needed. "
        "The user's existing Outlook session handles all authentication.\n\n"
        "PREREQUISITE: Outlook Desktop (Classic) must be running. The new/modern "
        "Outlook (olk.exe) is NOT supported — only the classic OUTLOOK.EXE.\n\n"
        "AVAILABLE TOOL CATEGORIES:\n"
        "- Email: send, list, read, search, reply, mark read/unread, move, attachments\n"
        "- Calendar: list events, create appointments/meetings, update, delete, "
        "respond to invites, search events\n"
        "- Tasks: create, list, complete, update, delete to-do items\n"
        "- Categories: list and set color categories on any item\n"
        "- Rules: list, create and enable/disable mail rules\n"
        "- Out of Office: check auto-reply status\n"
        "- Folders: list folder hierarchy with item counts"
    ),
)

bridge = OutlookBridge()


# --- Helper: resolve store by account name ---

def _resolve_store(namespace, account: str = ""):
    """Resolve an account name to an Outlook Store object.

    If account is empty, returns DefaultStore.
    Otherwise does a case-insensitive substring match on Store.DisplayName.
    """
    if not account:
        return namespace.DefaultStore

    account_lower = account.lower().strip()
    for i in range(namespace.Stores.Count):
        store = namespace.Stores.Item(i + 1)
        if account_lower in store.DisplayName.lower():
            return store

    return None


def _require_store(namespace, account: str = ""):
    """Resolve store, raising ValueError if not found."""
    store = _resolve_store(namespace, account)
    if store is None:
        raise ValueError(f"Account '{account}' not found. Use list_accounts to see available accounts.")
    return store


# --- Helper: resolve folder by name ---

def _walk_folders(parent, name_lower: str):
    """Recursively search subfolders of parent for a folder matching name_lower."""
    try:
        # Some Exchange/shared-mailbox folder types raise COMError on .Count;
        # treat them as empty to avoid crashing the traversal.
        count = parent.Folders.Count
    except Exception:
        return None
    for i in range(count):
        try:
            f = parent.Folders.Item(i + 1)
            if f.Name.lower() == name_lower:
                return f
            found = _walk_folders(f, name_lower)
            if found:
                return found
        except Exception:
            continue
    return None


def _resolve_folder(namespace, folder_name: str, store=None):
    """Resolve a folder name to an Outlook MAPIFolder object.

    Resolution order:
    1. Slash-delimited path (e.g. "Inbox/Receipts") — traverse segment by segment
    2. Built-in Outlook folder enum (inbox, sent, deleted, etc.)
    3. Root-level folder name match (fast path)
    4. Recursive depth-first search of entire folder tree (fallback)
    5. Search folders (virtual/search folders not in regular tree)
    """
    folder_name = folder_name.strip()
    store = store or namespace.DefaultStore

    # Slash-delimited path: traverse segment by segment
    if "/" in folder_name:
        parts = [p.strip() for p in folder_name.split("/")]
        current = _resolve_folder(namespace, parts[0], store)
        if current is None:
            return None
        for part in parts[1:]:
            part_lower = part.lower()
            found = None
            try:
                count = current.Folders.Count
            except Exception:
                return None
            for i in range(count):
                try:
                    f = current.Folders.Item(i + 1)
                    if f.Name.lower() == part_lower:
                        found = f
                        break
                except Exception:
                    continue
            if found is None:
                return None
            current = found
        return current

    folder_lower = folder_name.lower()

    # Built-in Outlook folders
    if folder_lower in FOLDER_NAME_TO_ENUM:
        return store.GetDefaultFolder(FOLDER_NAME_TO_ENUM[folder_lower])

    # Root-level search (fast path)
    root = store.GetRootFolder()
    for i in range(root.Folders.Count):
        try:
            f = root.Folders.Item(i + 1)
            if f.Name.lower() == folder_lower:
                return f
        except Exception:
            continue

    # Recursive fallback: search entire folder tree
    result = _walk_folders(root, folder_lower)
    if result:
        return result

    # Search folders fallback: virtual folders (e.g. "Flagged", "To-Do") are
    # not in the regular folder tree — access via Store.GetSearchFolders()
    try:
        search_folders = store.GetSearchFolders()
        for i in range(search_folders.Count):
            try:
                f = search_folders.Item(i + 1)
                if f.Name.lower() == folder_lower:
                    return f
            except Exception:
                continue
    except Exception:
        pass

    return None


# =====================================================================
# TOOL: list_accounts
# =====================================================================

@mcp.tool()
async def list_accounts() -> str:
    """List all Outlook accounts (stores) configured in the profile.

    Returns a JSON array of account objects with display_name, store_id,
    and is_default. Use the display_name (or a unique substring) as the
    'account' parameter in other tools to target a specific account.

    Returns:
        JSON array of account objects.
    """
    def _list(outlook, namespace):
        default_id = namespace.DefaultStore.StoreID
        results = []
        for i in range(namespace.Stores.Count):
            store = namespace.Stores.Item(i + 1)
            results.append({
                "display_name": store.DisplayName,
                "store_id": store.StoreID,
                "is_default": store.StoreID == default_id,
            })
        return json.dumps(results, indent=2, default=str)

    try:
        return await bridge.call(_list)
    except Exception as e:
        return f"Error listing accounts: {format_com_error(e)}"


# =====================================================================
# TOOL 1: send_email
# =====================================================================

@mcp.tool()
async def send_email(
    to: str,
    subject: str,
    body: str,
    cc: str = "",
    bcc: str = "",
    html_body: str = "",
    account: str = "",
) -> str:
    """Send an email using the user's Outlook account.

    Creates and sends an email immediately through the default Outlook profile.
    The email will appear in the user's Sent Items folder after sending.

    Args:
        to: One or more recipient email addresses, separated by semicolons.
            Example: "alice@example.com" or "alice@example.com; bob@example.com"
        subject: The email subject line.
        body: The plain-text body of the email. If html_body is also provided,
            both are set and Outlook will prefer the HTML version.
        cc: Optional. CC recipients, separated by semicolons.
        bcc: Optional. BCC recipients, separated by semicolons.
        html_body: Optional. HTML-formatted body. When provided, Outlook renders
            the email as HTML. The plain-text body serves as fallback.
        account: Optional. Account display name (or substring) to send from.
            Default: primary account. Use list_accounts to see available accounts.

    Returns:
        A confirmation message with subject and recipients, or an error.
    """
    def _send(outlook, namespace, to, subject, body, cc, bcc, html_body, account):
        store = _require_store(namespace, account)
        mail = outlook.CreateItem(OL_MAIL_ITEM)
        # Set the sending account
        for acc in outlook.Session.Accounts:
            if acc.DeliveryStore.StoreID == store.StoreID:
                mail._oleobj_.Invoke(*(64209, 0, 8, 0, acc))  # SendUsingAccount
                break
        mail.To = to
        mail.Subject = subject
        mail.Body = body
        if cc:
            mail.CC = cc
        if bcc:
            mail.BCC = bcc
        if html_body:
            mail.HTMLBody = html_body
        mail.Send()
        return f"Email sent: '{subject}' to {to}"

    try:
        return await bridge.call(_send, to, subject, body, cc, bcc, html_body, account)
    except Exception as e:
        return f"Error sending email: {format_com_error(e)}"


# =====================================================================
# TOOL 1B: save_draft
# =====================================================================

@mcp.tool()
async def save_draft(
    to: str,
    subject: str,
    body: str,
    cc: str = "",
    bcc: str = "",
    html_body: str = "",
    account: str = "",
) -> str:
    """Save an email as a draft in Outlook's Drafts folder without sending.

    Creates an email and stores it in the Drafts folder of the selected account
    so the user can review, edit, and send it later from Outlook. The message
    is NOT sent — use send_email for that. This is the recommended tool when a
    workflow requires human review before delivery.

    Args:
        to: One or more recipient email addresses, separated by semicolons.
            Example: "alice@example.com" or "alice@example.com; bob@example.com"
        subject: The email subject line.
        body: The plain-text body of the email. If html_body is also provided,
            both are set and Outlook will prefer the HTML version.
        cc: Optional. CC recipients, separated by semicolons.
        bcc: Optional. BCC recipients, separated by semicolons.
        html_body: Optional. HTML-formatted body. When provided, Outlook renders
            the email as HTML. The plain-text body serves as fallback.
        account: Optional. Account display name (or substring) to save under.
            Default: primary account. Use list_accounts to see available accounts.

    Returns:
        A confirmation message with subject and recipients, or an error.
    """
    def _save(outlook, namespace, to, subject, body, cc, bcc, html_body, account):
        store = _require_store(namespace, account)
        mail = outlook.CreateItem(OL_MAIL_ITEM)
        # Set the sending account so the draft is filed under the correct Drafts folder
        for acc in outlook.Session.Accounts:
            if acc.DeliveryStore.StoreID == store.StoreID:
                mail._oleobj_.Invoke(*(64209, 0, 8, 0, acc))  # SendUsingAccount
                break
        mail.To = to
        mail.Subject = subject
        mail.Body = body
        if cc:
            mail.CC = cc
        if bcc:
            mail.BCC = bcc
        if html_body:
            mail.HTMLBody = html_body
        mail.Save()
        return f"Draft saved: '{subject}' to {to}"

    try:
        return await bridge.call(_save, to, subject, body, cc, bcc, html_body, account)
    except Exception as e:
        return f"Error saving draft: {format_com_error(e)}"


# =====================================================================
# TOOL 1C: list_drafts
# =====================================================================

@mcp.tool()
async def list_drafts(
    limit: int = 50,
    offset: int = 0,
    account: str = "",
) -> str:
    """List drafts currently saved in Outlook's Drafts folder.

    Returns a JSON array of draft summaries sorted by last-modified time
    (newest first). Each summary includes entry_id, subject, to, cc, bcc,
    last_modified, has_attachments, and a short body_preview (first ~200
    chars of the plain-text body). Use entry_id with read_draft to load the
    full body or with send_draft to dispatch the message.

    Args:
        limit: Maximum number of drafts to return. Default 50, capped at 200.
        offset: Number of newest drafts to skip before returning results.
            Useful for pagination. Default 0.
        account: Optional. Account display name (or substring) to target.
            Default: primary account. Use list_accounts to see available accounts.

    Returns:
        JSON array of draft summary objects.
    """
    def _list(outlook, namespace, limit, offset, account):
        limit = min(max(1, limit), 200)
        offset = max(0, offset)
        store = _require_store(namespace, account)
        drafts = store.GetDefaultFolder(16)  # olFolderDrafts

        items = drafts.Items
        # Drafts use LastModificationTime rather than ReceivedTime
        try:
            items.Sort("[LastModificationTime]", True)
        except Exception:
            # Fallback: some stores reject sort on Drafts; iterate as-is
            pass

        total = items.Count
        results = []
        # Iterate sorted items, honoring offset and limit
        # items.Item is 1-indexed
        start = offset + 1
        end = min(offset + limit, total)
        for i in range(start, end + 1):
            try:
                item = items.Item(i)
                body = item.Body or ""
                preview = body[:200]
                if len(body) > 200:
                    preview += "..."
                results.append({
                    "entry_id": item.EntryID,
                    "subject": item.Subject or "(no subject)",
                    "to": item.To or "",
                    "cc": item.CC or "",
                    "bcc": item.BCC or "",
                    "last_modified": str(item.LastModificationTime),
                    "has_attachments": bool(item.Attachments.Count > 0),
                    "attachment_count": item.Attachments.Count,
                    "body_preview": preview,
                })
            except Exception:
                continue
        return json.dumps(results, indent=2, default=str)

    try:
        return await bridge.call(_list, limit, offset, account)
    except Exception as e:
        return f"Error listing drafts: {format_com_error(e)}"


# =====================================================================
# TOOL 1D: read_draft
# =====================================================================

@mcp.tool()
async def read_draft(
    draft_id: str,
    account: str = "",
) -> str:
    """Read the full content of a specific draft.

    Retrieves complete draft details including the full body, all recipients
    (To/CC/BCC), attachment file names, and last-modified time. Use this to
    review a draft before calling send_draft.

    Args:
        draft_id: The unique Outlook EntryID of the draft. Get this from
            list_drafts results.
        account: Optional. Account display name (or substring). Only needed
            if draft_id is ambiguous across stores.

    Returns:
        JSON object with full draft details (entry_id, subject, to, cc, bcc,
        body, html_body, attachments, last_modified, has_attachments).
    """
    def _read(outlook, namespace, draft_id, account):
        if account:
            store = _require_store(namespace, account)
            item = namespace.GetItemFromID(draft_id, store.StoreID)
        else:
            item = namespace.GetItemFromID(draft_id)
        if err := _check_item_class(item, _OL_CLASS_MAIL, "mail item"):
            return err

        attachments = []
        try:
            for i in range(item.Attachments.Count):
                try:
                    attachments.append(item.Attachments.Item(i + 1).FileName)
                except Exception:
                    continue
        except Exception:
            pass

        # HTMLBody may be very large; expose it but truncated like Body
        html_body = ""
        try:
            html_body = item.HTMLBody or ""
        except Exception:
            html_body = ""

        result = {
            "entry_id": item.EntryID,
            "subject": item.Subject or "(no subject)",
            "to": item.To or "",
            "cc": item.CC or "",
            "bcc": item.BCC or "",
            "body": item.Body or "",
            "html_body": html_body,
            "attachments": attachments,
            "has_attachments": len(attachments) > 0,
            "attachment_count": len(attachments),
            "last_modified": str(item.LastModificationTime),
        }
        return json.dumps(result, indent=2, default=str)

    try:
        return await bridge.call(_read, draft_id, account)
    except Exception as e:
        return f"Error reading draft: {format_com_error(e)}"


# =====================================================================
# TOOL 1E: send_draft
# =====================================================================

@mcp.tool()
async def send_draft(
    draft_id: str,
    account: str = "",
) -> str:
    """Send an existing draft from Outlook's Drafts folder.

    Looks up the draft by EntryID and calls MailItem.Send() on it. After a
    successful send, Outlook automatically moves the message from Drafts to
    Sent Items (this is standard Outlook behavior, not something this tool
    does explicitly). The EntryID of the message changes when it moves
    folders, so the returned message_id reflects the post-send location when
    available.

    Args:
        draft_id: The unique Outlook EntryID of the draft to send. Get this
            from list_drafts results.
        account: Optional. Account display name (or substring). Only needed
            if draft_id is ambiguous across stores.

    Returns:
        Confirmation message including the sent subject and recipients, or
        an error.
    """
    def _send(outlook, namespace, draft_id, account):
        if account:
            store = _require_store(namespace, account)
            item = namespace.GetItemFromID(draft_id, store.StoreID)
        else:
            item = namespace.GetItemFromID(draft_id)
        if err := _check_item_class(item, _OL_CLASS_MAIL, "mail item"):
            return err

        subject = item.Subject or "(no subject)"
        to = item.To or ""
        # Send() triggers Outlook to deliver the message and (in standard
        # Outlook behavior) move it to the Sent Items folder. The original
        # EntryID is no longer valid afterward — Outlook assigns a new one.
        item.Send()
        return json.dumps({
            "status": "sent",
            "subject": subject,
            "to": to,
            "message": f"Draft sent: '{subject}' to {to}",
        }, indent=2, default=str)

    try:
        return await bridge.call(_send, draft_id, account)
    except Exception as e:
        return f"Error sending draft: {format_com_error(e)}"


# =====================================================================
# TOOL 2: list_emails
# =====================================================================

@mcp.tool()
async def list_emails(
    folder: str = "inbox",
    count: int = 10,
    unread_only: bool = False,
    start_date: str = "",
    end_date: str = "",
    account: str = "",
) -> str:
    """List recent emails from a specified Outlook folder.

    Returns a JSON array of email summaries sorted by received time (newest
    first). Each summary includes entry_id, subject, sender, sender_name,
    received_time, unread status, and attachment info.

    Use the entry_id from results to read full content with read_email,
    or to perform actions like mark_as_read, move_email, or reply_email.

    Args:
        folder: The folder to list. Case-insensitive names: "inbox" (default),
            "sent"/"sentmail", "drafts", "deleted"/"trash", "junk"/"spam",
            "outbox", "archive", or any custom folder name visible in
            list_folders output.
        count: Maximum number of emails to return. Default 10, max recommended 50.
        unread_only: If true, only return unread emails. Default false.
        start_date: Optional. Only return emails received on or after this date.
            ISO 8601 format (e.g. "2026-03-10" or "2026-03-10 09:00").
        end_date: Optional. Only return emails received on or before this date.
            ISO 8601 format. Default: now (if start_date is provided).
        account: Optional. Account display name (or substring) to target.
            Default: primary account. Use list_accounts to see available accounts.

    Returns:
        JSON array of email summary objects.
    """
    def _list(outlook, namespace, folder, count, unread_only, start_date, end_date, account):
        count = min(max(1, count), 200)
        store = _require_store(namespace, account)
        target = _resolve_folder(namespace, folder, store)
        if not target:
            return json.dumps({"error": f"Folder '{folder}' not found"})

        items = target.Items
        items.Sort("[ReceivedTime]", True)

        # Build restriction filters
        restrictions = []
        if unread_only:
            restrictions.append("[UnRead] = True")
        if start_date:
            start = _parse_date(start_date)
            restrictions.append(f"[ReceivedTime] >= '{restrict_date(start)}'")
        if end_date:
            end = _parse_date(end_date)
            restrictions.append(f"[ReceivedTime] <= '{restrict_date(end)}'")
        elif start_date:
            # Default end to now when start is specified
            restrictions.append(f"[ReceivedTime] <= '{restrict_date(datetime.now())}'")

        if restrictions:
            items = items.Restrict(" AND ".join(restrictions))

        results = []
        limit = min(count, items.Count)
        for i in range(limit):
            try:
                results.append(format_email_summary(items.Item(i + 1)))
            except Exception:
                continue
        return json.dumps(results, indent=2, default=str)

    try:
        return await bridge.call(_list, folder, count, unread_only, start_date, end_date, account)
    except Exception as e:
        return f"Error listing emails: {format_com_error(e)}"


# =====================================================================
# TOOL 3: read_email
# =====================================================================

@mcp.tool()
async def read_email(
    entry_id: str = "",
    subject_search: str = "",
    folder: str = "inbox",
    account: str = "",
) -> str:
    """Read the full content of a specific email.

    Retrieves complete email details including body text, recipients, CC,
    and metadata. Provide EITHER entry_id (preferred, exact match) OR
    subject_search (finds most recent match by subject substring).

    Args:
        entry_id: The unique Outlook EntryID of the email. Most reliable way
            to identify a specific email. Get this from list_emails or
            search_emails results.
        subject_search: Alternative to entry_id. A case-insensitive substring
            to search for in email subjects. Returns the most recent match.
        folder: Folder to search when using subject_search. Ignored when
            entry_id is provided. Default "inbox".
        account: Optional. Account display name (or substring) to target.
            Default: primary account. Use list_accounts to see available accounts.

    Returns:
        JSON object with full email details (entry_id, subject, sender,
        sender_name, received_time, unread, to, cc, body, attachment info).
    """
    def _read(outlook, namespace, entry_id, subject_search, folder, account):
        if entry_id:
            item = namespace.GetItemFromID(entry_id)
            return json.dumps(format_email_full(item), indent=2, default=str)

        if not subject_search:
            return json.dumps({"error": "Provide either entry_id or subject_search"})

        store = _require_store(namespace, account)
        target = _resolve_folder(namespace, folder, store)
        if not target:
            return json.dumps({"error": f"Folder '{folder}' not found"})

        safe_query = _safe_dasl(subject_search)
        filter_str = (
            f"@SQL=\"urn:schemas:httpmail:subject\" LIKE '%{safe_query}%'"
        )
        items = target.Items.Restrict(filter_str)
        items.Sort("[ReceivedTime]", True)
        if items.Count == 0:
            return json.dumps({"error": f"No email found matching '{subject_search}'"})

        return json.dumps(format_email_full(items.Item(1)), indent=2, default=str)

    try:
        return await bridge.call(_read, entry_id, subject_search, folder, account)
    except Exception as e:
        return f"Error reading email: {format_com_error(e)}"


# =====================================================================
# TOOL 4: mark_as_read
# =====================================================================

@mcp.tool()
async def mark_as_read(entry_id: str, account: str = "") -> str:
    """Mark a specific email as read in Outlook.

    Changes the unread status to read, same as clicking on an email in Outlook.
    The change is persisted immediately and synced to the server.

    Args:
        entry_id: The unique Outlook EntryID of the email. Get this from
            list_emails or search_emails results.
        account: Optional. Account display name (or substring). Only needed
            if entry_id is ambiguous across stores.

    Returns:
        Confirmation message with the email subject, or an error.
    """
    def _mark(outlook, namespace, entry_id, account):
        if account:
            store = _require_store(namespace, account)
            item = namespace.GetItemFromID(entry_id, store.StoreID)
        else:
            item = namespace.GetItemFromID(entry_id)
        if err := _check_item_class(item, _OL_CLASS_MAIL, "mail item"):
            return err
        subject = item.Subject
        item.UnRead = False
        item.Save()
        return f"Marked as read: '{subject}'"

    try:
        return await bridge.call(_mark, entry_id, account)
    except Exception as e:
        return f"Error marking email as read: {format_com_error(e)}"


# =====================================================================
# TOOL 5: mark_as_unread
# =====================================================================

@mcp.tool()
async def mark_as_unread(entry_id: str, account: str = "") -> str:
    """Mark a specific email as unread in Outlook.

    Restores a previously read email to unread status. Useful for flagging
    emails that need follow-up attention. Persisted immediately.

    Args:
        entry_id: The unique Outlook EntryID of the email. Get this from
            list_emails or search_emails results.
        account: Optional. Account display name (or substring). Only needed
            if entry_id is ambiguous across stores.

    Returns:
        Confirmation message with the email subject, or an error.
    """
    def _mark(outlook, namespace, entry_id, account):
        if account:
            store = _require_store(namespace, account)
            item = namespace.GetItemFromID(entry_id, store.StoreID)
        else:
            item = namespace.GetItemFromID(entry_id)
        if err := _check_item_class(item, _OL_CLASS_MAIL, "mail item"):
            return err
        subject = item.Subject
        item.UnRead = True
        item.Save()
        return f"Marked as unread: '{subject}'"

    try:
        return await bridge.call(_mark, entry_id, account)
    except Exception as e:
        return f"Error marking email as unread: {format_com_error(e)}"


# =====================================================================
# TOOL 6: move_email
# =====================================================================

@mcp.tool()
async def move_email(
    entry_id: str,
    target_folder: str = "archive",
    account: str = "",
) -> str:
    """Move an email to a different Outlook folder.

    Moves the specified email from its current location to the target folder.
    IMPORTANT: After moving, the email gets a NEW entry_id — the old one
    becomes invalid. Common use: archiving emails after processing.

    Args:
        entry_id: The unique Outlook EntryID of the email to move.
        target_folder: Destination folder name. Default is "archive". Supports
            same names as list_emails: "archive", "inbox", "sent", "deleted"/
            "trash", "drafts", "junk"/"spam", or any custom folder name.
        account: Optional. Account display name (or substring) to resolve
            the target folder in. Default: primary account.

    Returns:
        Confirmation with email subject and destination, or an error.
    """
    def _move(outlook, namespace, entry_id, target_folder, account):
        item = namespace.GetItemFromID(entry_id)
        if err := _check_item_class(item, _OL_CLASS_MAIL, "mail item"):
            return err
        subject = item.Subject

        store = _require_store(namespace, account)
        dest = _resolve_folder(namespace, target_folder, store)
        if not dest:
            return f"Error: Target folder '{target_folder}' not found. Use list_folders to see available folders."

        item.Move(dest)
        return f"Moved '{subject}' to {target_folder}"

    try:
        return await bridge.call(_move, entry_id, target_folder, account)
    except Exception as e:
        return f"Error moving email: {format_com_error(e)}"


# =====================================================================
# TOOL 6B: move_by_filter
# =====================================================================

@mcp.tool()
async def move_by_filter(
    folder: str,
    dasl: str,
    target_folder: str,
    dry_run: bool = True,
    max_items: int = 500,
    account: str = "",
    target_account: str = "",
) -> str:
    """Move every item in a folder that matches a DASL filter, in bulk.

    Runs Items.Restrict with the given DASL condition on the source folder
    and moves all matches to target_folder. Defaults to dry_run=True, which
    only reports what would be moved — call again with dry_run=False to
    actually move. Use search_emails with the same `filter` to preview
    individual messages first.

    CAUTION: With dry_run=False this moves messages immediately and every
    moved item gets a NEW entry_id.

    Args:
        folder: Source folder (same names as list_emails, e.g. "inbox",
            "Inbox/Newsletters").
        dasl: DASL condition passed to Items.Restrict as-is ("@SQL=" prefix
            optional). Example: "urn:schemas:httpmail:fromemail" LIKE '%@example.com'
        target_folder: Destination folder name or slash path.
        dry_run: If true (default), only count and sample the matches.
        max_items: Maximum number of items to move in one call. Default 500,
            capped at 5000. Call again to process the rest.
        account: Optional. Account holding the source folder. Default: primary.
        target_account: Optional. Account holding the target folder.
            Default: same as account.

    Returns:
        JSON object with matched/moved counts and a sample of subjects.
    """
    def _move_all(outlook, namespace, folder, dasl, target_folder, dry_run,
                  max_items, account, target_account):
        max_items = min(max(1, max_items), 5000)
        condition = _raw_dasl(dasl)
        if not condition:
            return json.dumps({"error": "dasl filter is required"})

        store = _require_store(namespace, account)
        source = _resolve_folder(namespace, folder, store)
        if not source:
            return json.dumps({"error": f"Folder '{folder}' not found"})
        dest_store = _require_store(namespace, target_account or account)
        dest = _resolve_folder(namespace, target_folder, dest_store)
        if not dest:
            return json.dumps({"error": f"Target folder '{target_folder}' not found. Use list_folders to see available folders."})
        if dest.EntryID == source.EntryID:
            return json.dumps({"error": "Source and target folder are the same"})

        matches = source.Items.Restrict("@SQL=" + condition)
        matched = matches.Count
        # Snapshot the items before moving: moving shrinks the restricted
        # collection and would make index-based iteration skip items.
        batch = [matches.Item(i + 1) for i in range(min(matched, max_items))]

        sample = []
        for item in batch[:20]:
            try:
                sample.append(item.Subject)
            except Exception:
                continue

        result = {
            "dry_run": dry_run,
            "folder": folder,
            "target_folder": target_folder,
            "matched": matched,
            "sample_subjects": sample,
        }
        if dry_run:
            result["would_move"] = len(batch)
            return json.dumps(result, indent=2, default=str)

        logger.warning(
            "move_by_filter: moving %d items from '%s' to '%s'",
            len(batch), folder, target_folder,
        )
        moved, errors = 0, []
        for item in batch:
            try:
                item.Move(dest)
                moved += 1
            except Exception as e:
                if len(errors) < 10:
                    errors.append(format_com_error(e))
        result["moved"] = moved
        result["remaining"] = matched - moved
        if errors:
            result["errors"] = errors
        return json.dumps(result, indent=2, default=str)

    try:
        return await bridge.call(_move_all, folder, dasl, target_folder, dry_run,
                                 max_items, account, target_account)
    except Exception as e:
        return f"Error moving emails by filter: {format_com_error(e)}"


# =====================================================================
# TOOL 7: reply_email
# =====================================================================

@mcp.tool()
async def reply_email(
    entry_id: str,
    body: str,
    reply_all: bool = False,
    account: str = "",
) -> str:
    """Reply to an email in Outlook.

    Creates and sends a reply, preserving the original message thread.
    Use reply_all=True to reply to all recipients (sender + CC list).

    Args:
        entry_id: The unique Outlook EntryID of the email to reply to.
        body: The reply message text. Prepended above the original message
            in the email thread.
        reply_all: If true, reply to all recipients (sender + all CC/To).
            If false (default), reply only to the sender.
        account: Optional. Account display name (or substring). Only needed
            if entry_id is ambiguous across stores.

    Returns:
        Confirmation indicating the reply was sent, or an error.
    """
    def _reply(outlook, namespace, entry_id, body, reply_all, account):
        if account:
            store = _require_store(namespace, account)
            item = namespace.GetItemFromID(entry_id, store.StoreID)
        else:
            item = namespace.GetItemFromID(entry_id)
        if err := _check_item_class(item, _OL_CLASS_MAIL, "mail item"):
            return err
        subject = item.Subject
        reply_item = item.ReplyAll() if reply_all else item.Reply()
        reply_item.Body = body + "\n\n" + reply_item.Body
        reply_item.Send()
        return f"Reply sent to '{subject}' (reply_all={reply_all})"

    try:
        return await bridge.call(_reply, entry_id, body, reply_all, account)
    except Exception as e:
        return f"Error replying to email: {format_com_error(e)}"


# =====================================================================
# TOOL 8: list_folders
# =====================================================================

@mcp.tool()
async def list_folders(folder: str = "", max_depth: int = 3, account: str = "") -> str:
    """List mail folders in the user's Outlook mailbox.

    When called with no folder argument, lists top-level folders. Provide a
    folder name to drill into its subfolders — use this to browse the full
    folder tree step by step (e.g. first call with no folder to see top-level,
    then call with folder="Inbox" to see Inbox children, then
    folder="Inbox/Projects" to go deeper).

    Folder names from this output can be used directly in list_emails,
    move_email, search_emails, etc. Use slash-delimited paths for nested
    folders (e.g. "Inbox/Receipts/2026").

    Args:
        folder: Optional. Folder to list children of. Supports folder names
            ("Inbox"), slash paths ("Inbox/Receipts"), or built-in names
            ("sent", "drafts"). When empty, lists from the mailbox root.
        max_depth: How many levels deep to recurse below the starting folder.
            Default 3. Set to 1 to see only immediate children.
        account: Optional. Account display name (or substring) to target.
            Default: primary account. Use list_accounts to see available accounts.

    Returns:
        JSON array of folder objects with name, full_path, item_count,
        unread_count, and subfolders (if any).
    """
    def _list(outlook, namespace, folder, max_depth, account):
        max_depth = min(max(1, max_depth), 10)
        store = _require_store(namespace, account)

        if folder:
            start = _resolve_folder(namespace, folder, store)
            if not start:
                return json.dumps({"error": f"Folder '{folder}' not found"})
            base_path = folder
        else:
            start = store.GetRootFolder()
            base_path = ""

        def walk(f, depth, path_prefix):
            current_path = f"{path_prefix}/{f.Name}" if path_prefix else f.Name
            result = {
                "name": f.Name,
                "full_path": current_path,
                "item_count": f.Items.Count,
                "unread_count": f.UnReadItemCount,
            }
            if depth < max_depth:
                children = []
                for i in range(f.Folders.Count):
                    try:
                        child = f.Folders.Item(i + 1)
                        children.append(walk(child, depth + 1, current_path))
                    except Exception:
                        continue
                if children:
                    result["subfolders"] = children
            return result

        folders = []
        for i in range(start.Folders.Count):
            try:
                child = start.Folders.Item(i + 1)
                folders.append(walk(child, 1, base_path))
            except Exception:
                continue
        return json.dumps(folders, indent=2, default=str)

    try:
        return await bridge.call(_list, folder, max_depth, account)
    except Exception as e:
        return f"Error listing folders: {format_com_error(e)}"


# =====================================================================
# TOOL 9: search_emails
# =====================================================================

@mcp.tool()
async def search_emails(
    query: str = "",
    folder: str = "inbox",
    count: int = 10,
    start_date: str = "",
    end_date: str = "",
    account: str = "",
    sender: str = "",
    filter: str = "",
) -> str:
    """Search for emails in Outlook using text, sender and/or a raw DASL filter.

    Searches email subjects and bodies using Outlook's DASL filter.
    All given criteria are combined with AND; at least one of query,
    sender, filter or a date is required. Results are sorted by received
    time (newest first). Each result includes entry_id for further operations.

    Args:
        query: Optional. The search term (case-insensitive substring match
            on subject and body). Examples: "budget report", "quarterly".
        folder: Folder to search in. Default "inbox". Supports same
            names as list_emails.
        count: Maximum results to return. Default 10.
        start_date: Optional. Only return emails received on or after this date.
            ISO 8601 format (e.g. "2026-03-10" or "2026-03-10 09:00").
        end_date: Optional. Only return emails received on or before this date.
            ISO 8601 format. Default: now (if start_date is provided).
        account: Optional. Account display name (or substring) to target.
            Default: primary account. Use list_accounts to see available accounts.
        sender: Optional. Substring matched against the sender's email
            address or display name (e.g. "newsletter@", "Mario Rossi").
        filter: Optional. Raw DASL condition passed to Items.Restrict as-is
            (the "@SQL=" prefix is optional), for anything the other
            parameters don't cover. Example:
            "urn:schemas:httpmail:fromemail" LIKE '%@example.com'
            AND "urn:schemas:httpmail:hasattachment" = 1

    Returns:
        JSON array of matching email summaries, or an error.
    """
    def _search(outlook, namespace, query, folder, count, start_date, end_date,
                account, sender, filter):
        count = min(max(1, count), 200)
        if not (query or sender or filter.strip() or start_date or end_date):
            return json.dumps({"error": "Provide at least one of query, sender, filter, start_date or end_date"})
        store = _require_store(namespace, account)
        target = _resolve_folder(namespace, folder, store)
        if not target:
            return json.dumps({"error": f"Folder '{folder}' not found"})

        dasl_parts = []
        if query:
            safe_query = _safe_dasl(query)
            dasl_parts.append(
                f"(\"urn:schemas:httpmail:subject\" LIKE '%{safe_query}%' OR "
                f"\"urn:schemas:httpmail:textdescription\" LIKE '%{safe_query}%')"
            )
        if sender:
            safe_sender = _safe_dasl(sender)
            dasl_parts.append(
                f"(\"urn:schemas:httpmail:fromemail\" LIKE '%{safe_sender}%' OR "
                f"\"urn:schemas:httpmail:fromname\" LIKE '%{safe_sender}%')"
            )
        if raw := _raw_dasl(filter):
            dasl_parts.append(raw)
        if start_date:
            start = _parse_date(start_date)
            dasl_parts.append(
                f"\"urn:schemas:httpmail:datereceived\" >= '{restrict_date(start, utc=True)}'"
            )
        if end_date:
            end = _parse_date(end_date)
            dasl_parts.append(
                f"\"urn:schemas:httpmail:datereceived\" <= '{restrict_date(end, utc=True)}'"
            )
        elif start_date:
            dasl_parts.append(
                f"\"urn:schemas:httpmail:datereceived\" <= '{restrict_date(datetime.now(), utc=True)}'"
            )

        filter_str = "@SQL=" + " AND ".join(dasl_parts)
        items = target.Items.Restrict(filter_str)
        items.Sort("[ReceivedTime]", True)

        results = []
        limit = min(count, items.Count)
        for i in range(limit):
            try:
                results.append(format_email_summary(items.Item(i + 1)))
            except Exception:
                continue
        return json.dumps(results, indent=2, default=str)

    try:
        return await bridge.call(_search, query, folder, count, start_date, end_date,
                                 account, sender, filter)
    except Exception as e:
        return f"Error searching emails: {format_com_error(e)}"


# =====================================================================
# CALENDAR TOOLS
# =====================================================================


# --- Helper: parse ISO date string ---

def _parse_date(date_str: str) -> datetime:
    """Parse ISO 8601 date string like '2026-02-25 14:00' or '2026-02-25T14:00:00'."""
    return datetime.fromisoformat(date_str)


# =====================================================================
# TOOL 10: list_events
# =====================================================================

@mcp.tool()
async def list_events(
    start_date: str = "",
    end_date: str = "",
    count: int = 20,
    account: str = "",
    calendar: str = "",
) -> str:
    """List upcoming calendar events from Outlook.

    Returns a JSON array of event summaries within a date range, sorted by
    start time. Includes recurring event occurrences. Each summary has
    entry_id, subject, start, end, duration, location, organizer, attendees,
    and status info.

    Use entry_id from results with get_event, update_event, delete_event,
    or respond_to_meeting.

    Args:
        start_date: Start of date range in ISO 8601 format (e.g. "2026-02-25"
            or "2026-02-25 09:00"). Default: now.
        end_date: End of date range. Default: 7 days from start_date.
        count: Maximum number of events to return. Default 20.
        account: Optional. Account display name (or substring) to target.
            Default: primary account. Use list_accounts to see available accounts.
        calendar: Optional. Name or slash-delimited path of a specific
            calendar folder to read, e.g. "Your family" or
            "Kalender/Your family". Use list_calendars to discover available
            calendar folders (including shared/synced ones). Default: the
            account's default Calendar folder.

    Returns:
        JSON array of event summary objects.
    """
    def _list(outlook, namespace, start_date, end_date, count, account, calendar):
        count = min(max(1, count), 200)
        store = _require_store(namespace, account)
        if calendar:
            cal_folder = _resolve_folder(namespace, calendar, store)
            if cal_folder is None:
                return (
                    f"Error: Calendar folder '{calendar}' not found. "
                    "Use list_calendars to see available calendar folders."
                )
        else:
            cal_folder = store.GetDefaultFolder(OL_FOLDER_CALENDAR)
        items = cal_folder.Items

        # CRITICAL ORDER: Sort BEFORE IncludeRecurrences BEFORE Restrict
        items.Sort("[Start]")
        items.IncludeRecurrences = True

        start = _parse_date(start_date) if start_date else datetime.now()
        end = _parse_date(end_date) if end_date else start + timedelta(days=7)

        restrict = (
            f"[Start] >= '{restrict_date(start)}' "
            f"AND [Start] <= '{restrict_date(end)}'"
        )
        filtered = items.Restrict(restrict)

        results = []
        n = 0
        for item in filtered:
            n += 1
            try:
                results.append(format_event_summary(item))
            except Exception:
                continue
            if n >= count:
                break

        return json.dumps(results, indent=2, default=str)

    try:
        return await bridge.call(_list, start_date, end_date, count, account, calendar)
    except Exception as e:
        return f"Error listing events: {format_com_error(e)}"


# =====================================================================
# TOOL 11: get_event
# =====================================================================

@mcp.tool()
async def get_event(entry_id: str, account: str = "") -> str:
    """Read the full details of a specific calendar event.

    Retrieves complete event information including body/description,
    attendees, recurrence status, reminders, and response status.

    Args:
        entry_id: The unique Outlook EntryID of the event. Get this from
            list_events or search_events results.
        account: Optional. Account display name (or substring). Only needed
            if entry_id is ambiguous across stores.

    Returns:
        JSON object with full event details.
    """
    def _get(outlook, namespace, entry_id, account):
        if account:
            store = _require_store(namespace, account)
            item = namespace.GetItemFromID(entry_id, store.StoreID)
        else:
            item = namespace.GetItemFromID(entry_id)
        return json.dumps(format_event_full(item), indent=2, default=str)

    try:
        return await bridge.call(_get, entry_id, account)
    except Exception as e:
        return f"Error reading event: {format_com_error(e)}"


# =====================================================================
# TOOL 12: create_event
# =====================================================================

@mcp.tool()
async def create_event(
    subject: str,
    start: str,
    end: str,
    location: str = "",
    body: str = "",
    all_day: bool = False,
    reminder_minutes: int = 15,
    account: str = "",
    calendar: str = "",
) -> str:
    """Create a personal calendar appointment (no attendees).

    Creates and saves an appointment on the user's calendar. This is a
    personal event — no meeting invitations are sent. Use create_meeting
    instead if you need to invite attendees.

    Args:
        subject: The event title.
        start: Start time in ISO 8601 format. Examples: "2026-02-25 14:00",
            "2026-02-25T14:00:00". For all-day events, use just the date:
            "2026-02-25".
        end: End time in ISO 8601 format. For all-day events, use the next
            day: "2026-02-26".
        location: Optional. Event location (e.g. "Conference Room A",
            "Microsoft Teams Meeting").
        body: Optional. Description or notes for the event.
        all_day: If true, creates an all-day event. Default false.
        reminder_minutes: Minutes before the event to show a reminder.
            Default 15. Set to 0 to disable reminder.
        account: Optional. Account display name (or substring) to create
            the event in. Default: primary account.
        calendar: Optional. Name or slash-delimited path of a specific
            calendar folder to create the event in, e.g. "Your family" or
            "Kalender/Your family". Use list_calendars to discover available
            calendar folders (including shared/synced ones). Default: the
            account's default Calendar folder.

    Returns:
        Confirmation with event subject and entry_id, or an error.
    """
    def _create(outlook, namespace, subject, start, end, location, body,
                all_day, reminder_minutes, account, calendar):
        # Create directly inside the target calendar folder if one was
        # requested. (Note: creating via outlook.CreateItem() and then
        # calling .Move() on the unsaved item fails with a COM exception —
        # Move only works on items that already exist in a folder.)
        if account or calendar:
            store = _require_store(namespace, account)
            if calendar:
                cal = _resolve_folder(namespace, calendar, store)
                if cal is None:
                    return (
                        f"Error: Calendar folder '{calendar}' not found. "
                        "Use list_calendars to see available calendar folders."
                    )
            else:
                cal = store.GetDefaultFolder(OL_FOLDER_CALENDAR)
            appt = cal.Items.Add(OL_APPOINTMENT_ITEM)
        else:
            appt = outlook.CreateItem(OL_APPOINTMENT_ITEM)
        appt.Subject = subject
        appt.Start = start
        appt.End = end
        if location:
            appt.Location = location
        if body:
            appt.Body = body
        appt.AllDayEvent = all_day
        if reminder_minutes > 0:
            appt.ReminderSet = True
            appt.ReminderMinutesBeforeStart = reminder_minutes
        else:
            appt.ReminderSet = False
        appt.Save()
        return json.dumps({
            "status": "created",
            "subject": appt.Subject,
            "start": str(appt.Start),
            "end": str(appt.End),
            "entry_id": appt.EntryID,
            "calendar": calendar or "(default)",
        }, indent=2, default=str)

    try:
        return await bridge.call(
            _create, subject, start, end, location, body, all_day,
            reminder_minutes, account, calendar,
        )
    except Exception as e:
        return f"Error creating event: {format_com_error(e)}"


# =====================================================================
# TOOL: list_calendars
# =====================================================================

@mcp.tool()
async def list_calendars(account: str = "") -> str:
    """List all calendar folders available in an account.

    Includes the default Calendar plus any additional calendar folders —
    e.g. a shared "Your family" calendar, or calendars synced in from
    Google/other providers. Use the returned "full_path" (or a unique
    "name") as the `calendar` argument to create_event to target a
    specific calendar.

    Args:
        account: Optional. Account display name (or substring) to target.
            Default: primary account. Use list_accounts to see available accounts.

    Returns:
        JSON array of calendar folder objects with name, full_path, and item_count.
    """
    def _list(outlook, namespace, account):
        store = _require_store(namespace, account)
        root = store.GetRootFolder()
        results = []

        def walk(f, path_prefix):
            current_path = f"{path_prefix}/{f.Name}" if path_prefix else f.Name
            try:
                is_calendar = f.DefaultItemType == OL_APPOINTMENT_ITEM
            except Exception:
                is_calendar = False
            if is_calendar:
                try:
                    item_count = f.Items.Count
                except Exception:
                    item_count = None
                results.append({
                    "name": f.Name,
                    "full_path": current_path,
                    "item_count": item_count,
                })
            for i in range(f.Folders.Count):
                try:
                    walk(f.Folders.Item(i + 1), current_path)
                except Exception:
                    continue

        for i in range(root.Folders.Count):
            try:
                walk(root.Folders.Item(i + 1), "")
            except Exception:
                continue

        return json.dumps(results, indent=2, default=str)

    try:
        return await bridge.call(_list, account)
    except Exception as e:
        return f"Error listing calendars: {format_com_error(e)}"


# =====================================================================
# TOOL 13: create_meeting
# =====================================================================

@mcp.tool()
async def create_meeting(
    subject: str,
    start: str,
    end: str,
    required_attendees: str,
    location: str = "",
    body: str = "",
    optional_attendees: str = "",
    account: str = "",
) -> str:
    """Create a meeting and send invitations to attendees.

    Creates a calendar meeting and immediately sends meeting requests to
    all specified attendees. The meeting will appear on the organizer's
    calendar and attendees will receive an invitation they can accept,
    decline, or tentatively accept.

    Args:
        subject: The meeting title.
        start: Start time in ISO 8601 format (e.g. "2026-02-25 14:00").
        end: End time in ISO 8601 format (e.g. "2026-02-25 15:00").
        required_attendees: Required attendee email addresses, separated by
            semicolons. Example: "alice@example.com; bob@example.com"
        location: Optional. Meeting location (e.g. "Teams", "Room 301").
        body: Optional. Meeting description or agenda.
        optional_attendees: Optional. Optional attendee emails, separated
            by semicolons.
        account: Optional. Account display name (or substring) to send from.
            Default: primary account. Use list_accounts to see available accounts.

    Returns:
        Confirmation that the meeting was created and invitations sent.
    """
    def _create(outlook, namespace, subject, start, end, required_attendees,
                location, body, optional_attendees, account):
        appt = outlook.CreateItem(OL_APPOINTMENT_ITEM)
        # Set sending account
        if account:
            store = _require_store(namespace, account)
            for acc in outlook.Session.Accounts:
                if acc.DeliveryStore.StoreID == store.StoreID:
                    appt._oleobj_.Invoke(*(64209, 0, 8, 0, acc))
                    break
        appt.Subject = subject
        appt.Start = start
        appt.End = end
        appt.MeetingStatus = OL_MEETING
        if location:
            appt.Location = location
        if body:
            appt.Body = body

        for addr in required_attendees.split(";"):
            addr = addr.strip()
            if addr:
                recip = appt.Recipients.Add(addr)
                recip.Type = OL_REQUIRED

        if optional_attendees:
            for addr in optional_attendees.split(";"):
                addr = addr.strip()
                if addr:
                    recip = appt.Recipients.Add(addr)
                    recip.Type = OL_OPTIONAL

        appt.Recipients.ResolveAll()
        appt.Send()
        return (
            f"Meeting '{subject}' created and invitations sent to "
            f"{required_attendees}"
        )

    try:
        return await bridge.call(
            _create, subject, start, end, required_attendees, location, body,
            optional_attendees, account,
        )
    except Exception as e:
        return f"Error creating meeting: {format_com_error(e)}"


# =====================================================================
# TOOL 14: update_event
# =====================================================================

@mcp.tool()
async def update_event(
    entry_id: str,
    subject: str = "",
    start: str = "",
    end: str = "",
    location: str = "",
    body: str = "",
    account: str = "",
) -> str:
    """Update an existing calendar event.

    Modifies properties of an appointment or meeting. Only the fields you
    provide will be updated — omitted fields remain unchanged. For meetings
    you organize, attendees will receive an update notification.

    Args:
        entry_id: The unique Outlook EntryID of the event to update.
        subject: Optional. New event title.
        start: Optional. New start time in ISO 8601 format.
        end: Optional. New end time in ISO 8601 format.
        location: Optional. New location.
        body: Optional. New description/notes.
        account: Optional. Account display name (or substring). Only needed
            if entry_id is ambiguous across stores.

    Returns:
        Confirmation with updated event details, or an error.
    """
    def _update(outlook, namespace, entry_id, subject, start, end, location, body, account):
        if account:
            store = _require_store(namespace, account)
            item = namespace.GetItemFromID(entry_id, store.StoreID)
        else:
            item = namespace.GetItemFromID(entry_id)
        if err := _check_item_class(item, _OL_CLASS_APPOINTMENT, "appointment/meeting item"):
            return err
        if subject:
            item.Subject = subject
        if start:
            item.Start = start
        if end:
            item.End = end
        if location:
            item.Location = location
        if body:
            item.Body = body
        item.Save()
        return json.dumps({
            "status": "updated",
            "subject": item.Subject,
            "start": str(item.Start),
            "end": str(item.End),
            "location": item.Location or "",
            "entry_id": item.EntryID,
        }, indent=2, default=str)

    try:
        return await bridge.call(
            _update, entry_id, subject, start, end, location, body, account,
        )
    except Exception as e:
        return f"Error updating event: {format_com_error(e)}"


# =====================================================================
# TOOL 15: delete_event
# =====================================================================

@mcp.tool()
async def delete_event(entry_id: str, account: str = "") -> str:
    """Delete a calendar event or cancel a meeting.

    For personal appointments, the event is simply deleted. For meetings
    you organized, this cancels the meeting and sends cancellation notices
    to all attendees. For meetings you received, this declines and removes
    the event from your calendar.

    Args:
        entry_id: The unique Outlook EntryID of the event to delete/cancel.
        account: Optional. Account display name (or substring). Only needed
            if entry_id is ambiguous across stores.

    Returns:
        Confirmation with the event subject, or an error.
    """
    def _delete(outlook, namespace, entry_id, account):
        if account:
            store = _require_store(namespace, account)
            item = namespace.GetItemFromID(entry_id, store.StoreID)
        else:
            item = namespace.GetItemFromID(entry_id)
        if err := _check_item_class(item, _OL_CLASS_APPOINTMENT, "appointment/meeting item"):
            return err
        subject = item.Subject
        meeting_status = item.MeetingStatus

        # If this is a meeting we organized, cancel it (sends notices)
        if meeting_status == OL_MEETING:
            item.MeetingStatus = OL_MEETING_CANCELED
            item.Send()
            return f"Meeting canceled: '{subject}' (cancellation sent to attendees)"

        # Otherwise just delete
        item.Delete()
        return f"Event deleted: '{subject}'"

    try:
        return await bridge.call(_delete, entry_id, account)
    except Exception as e:
        return f"Error deleting event: {format_com_error(e)}"


# =====================================================================
# TOOL 16: respond_to_meeting
# =====================================================================

@mcp.tool()
async def respond_to_meeting(
    entry_id: str,
    response: str,
    account: str = "",
) -> str:
    """Respond to a meeting invitation (accept, decline, or tentative).

    Sends your response to the meeting organizer. The meeting will be
    added to (or updated on) your calendar accordingly.

    Args:
        entry_id: The unique Outlook EntryID of the meeting to respond to.
            Get this from list_events or search_events.
        response: Your response. Must be one of: "accept", "decline",
            or "tentative".
        account: Optional. Account display name (or substring). Only needed
            if entry_id is ambiguous across stores.

    Returns:
        Confirmation of your response, or an error.
    """
    def _respond(outlook, namespace, entry_id, response, account):
        response_map = {
            "accept": OL_RESPONSE_ACCEPTED,
            "decline": OL_RESPONSE_DECLINED,
            "tentative": OL_RESPONSE_TENTATIVE,
        }
        response_lower = response.lower().strip()
        if response_lower not in response_map:
            return f"Error: response must be 'accept', 'decline', or 'tentative'. Got: '{response}'"

        if account:
            store = _require_store(namespace, account)
            item = namespace.GetItemFromID(entry_id, store.StoreID)
        else:
            item = namespace.GetItemFromID(entry_id)
        if err := _check_item_class(item, _OL_CLASS_APPOINTMENT, "appointment/meeting item"):
            return err
        subject = item.Subject
        response_item = item.Respond(response_map[response_lower])
        response_item.Send()
        return f"Responded '{response_lower}' to meeting: '{subject}'"

    try:
        return await bridge.call(_respond, entry_id, response, account)
    except Exception as e:
        return f"Error responding to meeting: {format_com_error(e)}"


# =====================================================================
# TOOL 17: search_events
# =====================================================================

@mcp.tool()
async def search_events(
    query: str,
    start_date: str = "",
    end_date: str = "",
    count: int = 10,
    account: str = "",
    calendar: str = "",
) -> str:
    """Search for calendar events by keyword.

    Searches event subjects within a date range. Results are sorted by
    start time. Includes recurring event occurrences.

    Args:
        query: The search term (case-insensitive substring match on subject).
            Examples: "standup", "review", "1:1".
        start_date: Start of search range in ISO 8601 format. Default: 30
            days ago.
        end_date: End of search range. Default: 30 days from now.
        count: Maximum results to return. Default 10.
        account: Optional. Account display name (or substring) to target.
            Default: primary account. Use list_accounts to see available accounts.
        calendar: Optional. Name or slash-delimited path of a specific
            calendar folder to search, e.g. "Your family" or
            "Kalender/Your family". Use list_calendars to discover available
            calendar folders (including shared/synced ones). Default: the
            account's default Calendar folder.

    Returns:
        JSON array of matching event summaries.
    """
    def _search(outlook, namespace, query, start_date, end_date, count, account, calendar):
        count = min(max(1, count), 200)
        store = _require_store(namespace, account)
        if calendar:
            cal_folder = _resolve_folder(namespace, calendar, store)
            if cal_folder is None:
                return (
                    f"Error: Calendar folder '{calendar}' not found. "
                    "Use list_calendars to see available calendar folders."
                )
        else:
            cal_folder = store.GetDefaultFolder(OL_FOLDER_CALENDAR)
        items = cal_folder.Items
        items.Sort("[Start]")
        items.IncludeRecurrences = True

        start = _parse_date(start_date) if start_date else datetime.now() - timedelta(days=30)
        end = _parse_date(end_date) if end_date else datetime.now() + timedelta(days=30)

        restrict = (
            f"[Start] >= '{restrict_date(start)}' "
            f"AND [Start] <= '{restrict_date(end)}'"
        )
        filtered = items.Restrict(restrict)

        query_lower = query.lower()
        results = []
        for item in filtered:
            if query_lower in (item.Subject or "").lower():
                try:
                    results.append(format_event_summary(item))
                except Exception:
                    continue
                if len(results) >= count:
                    break

        return json.dumps(results, indent=2, default=str)

    try:
        return await bridge.call(_search, query, start_date, end_date, count, account, calendar)
    except Exception as e:
        return f"Error searching events: {format_com_error(e)}"


# =====================================================================
# TASK TOOLS
# =====================================================================

@mcp.tool()
async def list_tasks(
    include_completed: bool = False,
    count: int = 20,
    account: str = "",
) -> str:
    """List tasks from the Outlook Tasks folder.

    Returns a JSON array of task summaries sorted by due date. Each task
    includes entry_id, subject, status, percent_complete, due_date,
    importance, and categories.

    Args:
        include_completed: If true, include completed tasks. Default false
            (only pending/in-progress tasks).
        count: Maximum number of tasks to return. Default 20.
        account: Optional. Account display name (or substring) to target.
            Default: primary account. Use list_accounts to see available accounts.

    Returns:
        JSON array of task summary objects.
    """
    def _list(outlook, namespace, include_completed, count, account):
        count = min(max(1, count), 200)
        store = _require_store(namespace, account)
        folder = store.GetDefaultFolder(OL_FOLDER_TASKS)
        items = folder.Items
        items.Sort("[DueDate]")

        if not include_completed:
            items = items.Restrict("[Complete] = False")

        results = []
        limit = min(count, items.Count)
        for i in range(limit):
            try:
                results.append(format_task_summary(items.Item(i + 1)))
            except Exception:
                continue
        return json.dumps(results, indent=2, default=str)

    try:
        return await bridge.call(_list, include_completed, count, account)
    except Exception as e:
        return f"Error listing tasks: {format_com_error(e)}"


@mcp.tool()
async def get_task(entry_id: str, account: str = "") -> str:
    """Read the full details of a specific task.

    Args:
        entry_id: The unique Outlook EntryID of the task.
        account: Optional. Account display name (or substring). Only needed
            if entry_id is ambiguous across stores.

    Returns:
        JSON object with full task details including body.
    """
    def _get(outlook, namespace, entry_id, account):
        if account:
            store = _require_store(namespace, account)
            item = namespace.GetItemFromID(entry_id, store.StoreID)
        else:
            item = namespace.GetItemFromID(entry_id)
        return json.dumps(format_task_full(item), indent=2, default=str)

    try:
        return await bridge.call(_get, entry_id, account)
    except Exception as e:
        return f"Error reading task: {format_com_error(e)}"


@mcp.tool()
async def create_task(
    subject: str,
    body: str = "",
    due_date: str = "",
    importance: str = "normal",
    reminder_minutes: int = 0,
    account: str = "",
) -> str:
    """Create a new task in Outlook.

    Args:
        subject: The task title.
        body: Optional. Task description or notes.
        due_date: Optional. Due date in ISO 8601 format (e.g. "2026-03-01").
        importance: Optional. "low", "normal" (default), or "high".
        reminder_minutes: Optional. Minutes before due date to remind.
            Default 0 (no reminder).
        account: Optional. Account display name (or substring) to create
            the task in. Default: primary account.

    Returns:
        Confirmation with task subject and entry_id.
    """
    def _create(outlook, namespace, subject, body, due_date, importance,
                reminder_minutes, account):
        # Create directly inside the target store's tasks folder if one was
        # requested. (Note: creating via outlook.CreateItem() and then
        # calling .Move() on the unsaved item fails with a COM exception —
        # Move only works on items that already exist in a folder.)
        if account:
            store = _require_store(namespace, account)
            tasks_folder = store.GetDefaultFolder(OL_FOLDER_TASKS)
            task = tasks_folder.Items.Add(OL_TASK_ITEM)
        else:
            task = outlook.CreateItem(OL_TASK_ITEM)
        task.Subject = subject
        if body:
            task.Body = body
        if due_date:
            task.DueDate = due_date
        imp_map = {"low": 0, "normal": 1, "high": 2}
        task.Importance = imp_map.get(importance.lower(), 1)
        if reminder_minutes > 0:
            task.ReminderSet = True
            task.ReminderMinutesBeforeStart = reminder_minutes
        else:
            task.ReminderSet = False
        task.Save()
        return json.dumps({
            "status": "created",
            "subject": task.Subject,
            "entry_id": task.EntryID,
            "due_date": str(task.DueDate) if due_date else None,
        }, indent=2, default=str)

    try:
        return await bridge.call(
            _create, subject, body, due_date, importance, reminder_minutes,
            account,
        )
    except Exception as e:
        return f"Error creating task: {format_com_error(e)}"


@mcp.tool()
async def complete_task(entry_id: str, account: str = "") -> str:
    """Mark a task as complete.

    Sets the task status to complete and percent to 100%.

    Args:
        entry_id: The unique Outlook EntryID of the task.
        account: Optional. Account display name (or substring). Only needed
            if entry_id is ambiguous across stores.

    Returns:
        Confirmation with the task subject.
    """
    def _complete(outlook, namespace, entry_id, account):
        if account:
            store = _require_store(namespace, account)
            item = namespace.GetItemFromID(entry_id, store.StoreID)
        else:
            item = namespace.GetItemFromID(entry_id)
        if err := _check_item_class(item, _OL_CLASS_TASK, "task item"):
            return err
        item.Status = OL_TASK_COMPLETE
        item.PercentComplete = 100
        item.Save()
        return f"Task completed: '{item.Subject}'"

    try:
        return await bridge.call(_complete, entry_id, account)
    except Exception as e:
        return f"Error completing task: {format_com_error(e)}"


@mcp.tool()
async def delete_task(entry_id: str, account: str = "") -> str:
    """Delete a task from Outlook.

    Args:
        entry_id: The unique Outlook EntryID of the task to delete.
        account: Optional. Account display name (or substring). Only needed
            if entry_id is ambiguous across stores.

    Returns:
        Confirmation with the task subject.
    """
    def _delete(outlook, namespace, entry_id, account):
        if account:
            store = _require_store(namespace, account)
            item = namespace.GetItemFromID(entry_id, store.StoreID)
        else:
            item = namespace.GetItemFromID(entry_id)
        if err := _check_item_class(item, _OL_CLASS_TASK, "task item"):
            return err
        subject = item.Subject
        item.Delete()
        return f"Task deleted: '{subject}'"

    try:
        return await bridge.call(_delete, entry_id, account)
    except Exception as e:
        return f"Error deleting task: {format_com_error(e)}"


# =====================================================================
# ATTACHMENT TOOLS
# =====================================================================

@mcp.tool()
async def list_attachments(entry_id: str, account: str = "") -> str:
    """List all attachments on an email or calendar event.

    Args:
        entry_id: The EntryID of the email or event to check for attachments.
        account: Optional. Account display name (or substring). Only needed
            if entry_id is ambiguous across stores.

    Returns:
        JSON array of attachment objects with index, filename, and size.
    """
    def _list(outlook, namespace, entry_id, account):
        if account:
            store = _require_store(namespace, account)
            item = namespace.GetItemFromID(entry_id, store.StoreID)
        else:
            item = namespace.GetItemFromID(entry_id)
        results = []
        for i in range(item.Attachments.Count):
            att = item.Attachments.Item(i + 1)
            results.append({
                "index": i + 1,
                "filename": att.FileName,
                "size": att.Size,
            })
        return json.dumps(results, indent=2, default=str)

    try:
        return await bridge.call(_list, entry_id, account)
    except Exception as e:
        return f"Error listing attachments: {format_com_error(e)}"


@mcp.tool()
async def save_attachment(
    entry_id: str,
    attachment_index: int = 1,
    save_directory: str = "",
    account: str = "",
) -> str:
    """Save an attachment from an email or event to disk.

    Downloads the specified attachment to a local directory.

    Args:
        entry_id: The EntryID of the email or event containing the attachment.
        attachment_index: Which attachment to save (1-based index). Default 1
            (first attachment). Use list_attachments to see available indices.
        save_directory: Directory to save the file to. Default: user's
            Downloads folder.
        account: Optional. Account display name (or substring). Only needed
            if entry_id is ambiguous across stores.

    Returns:
        The full file path where the attachment was saved, or an error.
    """
    def _save(outlook, namespace, entry_id, attachment_index, save_directory, account):
        if account:
            store = _require_store(namespace, account)
            item = namespace.GetItemFromID(entry_id, store.StoreID)
        else:
            item = namespace.GetItemFromID(entry_id)
        if attachment_index < 1 or item.Attachments.Count < attachment_index:
            return f"Error: Only {item.Attachments.Count} attachment(s), requested index {attachment_index}"

        att = item.Attachments.Item(attachment_index)
        if not save_directory:
            save_directory = os.path.join(os.path.expanduser("~"), "Downloads")

        # Resolve to real path before creating
        save_directory = os.path.realpath(save_directory)
        os.makedirs(save_directory, exist_ok=True)

        # Strip path separators and dangerous characters from filename
        safe_name = os.path.basename(att.FileName)
        safe_name = re.sub(r'[^\w\.\-_ ]', '_', safe_name)
        if not safe_name:
            safe_name = "attachment"

        save_path = os.path.join(save_directory, safe_name)

        # Ensure final path is still inside the intended directory
        if not os.path.realpath(save_path).startswith(save_directory + os.sep) and \
           os.path.realpath(save_path) != save_directory:
            return "Error: Attachment filename would escape the target directory."

        att.SaveAsFile(save_path)
        return json.dumps({
            "status": "saved",
            "filename": safe_name,
            "path": save_path,
            "size": att.Size,
        }, indent=2, default=str)

    try:
        return await bridge.call(_save, entry_id, attachment_index, save_directory, account)
    except Exception as e:
        return f"Error saving attachment: {format_com_error(e)}"


# =====================================================================
# CATEGORY TOOLS
# =====================================================================

@mcp.tool()
async def list_categories(account: str = "") -> str:
    """List all available Outlook categories.

    Returns the color categories configured in the user's Outlook profile.
    These can be applied to emails, events, tasks, and other items.

    Args:
        account: Optional. Account display name (or substring) to target.
            Default: primary account. Use list_accounts to see available accounts.

    Returns:
        JSON array of category objects with name and color index.
    """
    def _list(outlook, namespace, account):
        # Categories are profile-wide, not per-store, but we accept the param for consistency
        results = []
        for i in range(namespace.Categories.Count):
            cat = namespace.Categories.Item(i + 1)
            results.append({"name": cat.Name, "color": cat.Color})
        return json.dumps(results, indent=2, default=str)

    try:
        return await bridge.call(_list, account)
    except Exception as e:
        return f"Error listing categories: {format_com_error(e)}"


@mcp.tool()
async def set_category(
    entry_id: str,
    categories: str,
    account: str = "",
) -> str:
    """Set categories on an email, event, or task.

    Replaces any existing categories on the item. Use comma-separated
    values for multiple categories.

    Args:
        entry_id: The EntryID of the item to categorize.
        categories: Category name(s), comma-separated. Example:
            "Important" or "Work, Follow-up". Use an empty string to
            clear all categories.
        account: Optional. Account display name (or substring). Only needed
            if entry_id is ambiguous across stores.

    Returns:
        Confirmation with the item subject and applied categories.
    """
    def _set(outlook, namespace, entry_id, categories, account):
        if account:
            store = _require_store(namespace, account)
            item = namespace.GetItemFromID(entry_id, store.StoreID)
        else:
            item = namespace.GetItemFromID(entry_id)
        item.Categories = categories
        item.Save()
        return (
            f"Categories set on '{item.Subject}': "
            f"'{item.Categories or '(none)'}'"
        )

    try:
        return await bridge.call(_set, entry_id, categories, account)
    except Exception as e:
        return f"Error setting categories: {format_com_error(e)}"


# =====================================================================
# RULES TOOLS
# =====================================================================

@mcp.tool()
async def list_rules(account: str = "") -> str:
    """List all mail rules in Outlook.

    Returns the configured inbox rules with their names and enabled status.

    Args:
        account: Optional. Account display name (or substring) to target.
            Default: primary account. Use list_accounts to see available accounts.

    Returns:
        JSON array of rule objects with name, enabled status, and index.
    """
    def _list(outlook, namespace, account):
        store = _require_store(namespace, account)
        rules = store.GetRules()
        results = []
        for i in range(rules.Count):
            rule = rules.Item(i + 1)
            results.append({
                "index": i + 1,
                "name": rule.Name,
                "enabled": bool(rule.Enabled),
            })
        return json.dumps(results, indent=2, default=str)

    try:
        return await bridge.call(_list, account)
    except Exception as e:
        return f"Error listing rules: {format_com_error(e)}"


@mcp.tool()
async def toggle_rule(
    rule_name: str,
    enabled: bool,
    account: str = "",
) -> str:
    """Enable or disable a mail rule by name.

    CAUTION: This modifies live mail rules immediately. Confirm the rule name
    with list_rules before calling.

    Args:
        rule_name: The exact name of the rule to toggle. Use list_rules
            to see available rule names.
        enabled: True to enable the rule, False to disable it.
        account: Optional. Account display name (or substring) to target.
            Default: primary account. Use list_accounts to see available accounts.

    Returns:
        Confirmation with the rule name and new status.
    """
    def _toggle(outlook, namespace, rule_name, enabled, account):
        store = _require_store(namespace, account)
        rules = store.GetRules()
        for i in range(rules.Count):
            rule = rules.Item(i + 1)
            if rule.Name == rule_name:
                logger.warning(
                    "toggle_rule: setting rule '%s' enabled=%s", rule_name, enabled
                )
                rule.Enabled = enabled
                rules.Save()
                status = "enabled" if enabled else "disabled"
                return f"Rule '{rule_name}' {status}"
        return f"Error: Rule '{rule_name}' not found. Use list_rules to see available rules."

    try:
        return await bridge.call(_toggle, rule_name, enabled, account)
    except Exception as e:
        return f"Error toggling rule: {format_com_error(e)}"


_OL_RULE_RECEIVE = 0
_OL_RULE_EXECUTE_ALL_MESSAGES = 0
_OL_FOLDER_INBOX = 6


def _split_terms(value: str) -> list[str]:
    """Split a comma-separated parameter into a list of non-empty terms."""
    return [t.strip() for t in value.split(",") if t.strip()]


@mcp.tool()
async def create_rule(
    name: str,
    sender_contains: str = "",
    subject_contains: str = "",
    body_or_subject_contains: str = "",
    has_attachment: bool = False,
    except_subject_contains: str = "",
    move_to_folder: str = "",
    copy_to_folder: str = "",
    assign_categories: str = "",
    delete: bool = False,
    stop_processing: bool = False,
    enabled: bool = True,
    run_now: bool = False,
    account: str = "",
) -> str:
    """Create a new incoming-mail rule.

    All given conditions must match (AND); multiple comma-separated terms
    within one condition match if ANY term matches (OR). At least one
    condition and one action are required. The Outlook object model cannot
    create every rule the Rules wizard can — e.g. "mark as read" is not
    available programmatically.

    CAUTION: This modifies live mail rules immediately. Check list_rules
    first — rule names are not unique, so a duplicate name creates a second
    rule rather than replacing the existing one.

    Args:
        name: Rule name shown in Outlook's Rules and Alerts.
        sender_contains: Comma-separated text matched against the sender's
            address (e.g. "@newsletter.com, noreply@").
        subject_contains: Comma-separated words/phrases in the subject.
        body_or_subject_contains: Comma-separated words/phrases in subject or body.
        has_attachment: If true, only messages with attachments match.
        except_subject_contains: Comma-separated subject terms that EXCLUDE
            a message from the rule.
        move_to_folder: Action. Folder name or slash path to move matches to.
        copy_to_folder: Action. Folder name or slash path to copy matches to.
        assign_categories: Action. Comma-separated category names to assign.
        delete: Action. Move matches to Deleted Items.
        stop_processing: Action. Stop processing further rules after this one.
        enabled: Whether the rule is active after creation. Default true.
        run_now: If true, also apply the rule once to messages already in
            the Inbox (not subfolders).
        account: Optional. Account display name (or substring) to target.
            Default: primary account. Use list_accounts to see available accounts.

    Returns:
        JSON object describing the created rule, or an error.
    """
    def _create(outlook, namespace, name, sender_contains, subject_contains,
                body_or_subject_contains, has_attachment, except_subject_contains,
                move_to_folder, copy_to_folder, assign_categories, delete,
                stop_processing, enabled, run_now, account):
        senders = _split_terms(sender_contains)
        subjects = _split_terms(subject_contains)
        body_terms = _split_terms(body_or_subject_contains)
        except_subjects = _split_terms(except_subject_contains)
        categories = _split_terms(assign_categories)

        if not (senders or subjects or body_terms or has_attachment):
            return json.dumps({"error": "At least one condition is required"})
        if not (move_to_folder or copy_to_folder or categories or delete):
            return json.dumps({"error": "At least one action is required (move, copy, categories or delete)"})
        if move_to_folder and delete:
            return json.dumps({"error": "move_to_folder and delete are mutually exclusive"})

        store = _require_store(namespace, account)
        move_dest = copy_dest = None
        if move_to_folder:
            move_dest = _resolve_folder(namespace, move_to_folder, store)
            if not move_dest:
                return json.dumps({"error": f"Folder '{move_to_folder}' not found"})
        if copy_to_folder:
            copy_dest = _resolve_folder(namespace, copy_to_folder, store)
            if not copy_dest:
                return json.dumps({"error": f"Folder '{copy_to_folder}' not found"})

        rules = store.GetRules()
        rule = rules.Create(name, _OL_RULE_RECEIVE)

        cond = rule.Conditions
        if senders:
            cond.SenderAddress.Address = senders
            cond.SenderAddress.Enabled = True
        if subjects:
            cond.Subject.Text = subjects
            cond.Subject.Enabled = True
        if body_terms:
            cond.BodyOrSubject.Text = body_terms
            cond.BodyOrSubject.Enabled = True
        if has_attachment:
            cond.HasAttachment.Enabled = True
        if except_subjects:
            rule.Exceptions.Subject.Text = except_subjects
            rule.Exceptions.Subject.Enabled = True

        actions = rule.Actions
        if move_dest is not None:
            actions.MoveToFolder.Folder = move_dest
            actions.MoveToFolder.Enabled = True
        if copy_dest is not None:
            actions.CopyToFolder.Folder = copy_dest
            actions.CopyToFolder.Enabled = True
        if categories:
            actions.AssignToCategory.Categories = categories
            actions.AssignToCategory.Enabled = True
        if delete:
            actions.Delete.Enabled = True
        if stop_processing:
            actions.Stop.Enabled = True

        rule.Enabled = enabled
        logger.warning("create_rule: creating rule '%s' (enabled=%s)", name, enabled)
        rules.Save()

        if run_now:
            rule.Execute(False, store.GetDefaultFolder(_OL_FOLDER_INBOX), False,
                         _OL_RULE_EXECUTE_ALL_MESSAGES)

        return json.dumps({
            "status": "created",
            "name": name,
            "enabled": enabled,
            "conditions": {
                "sender_contains": senders,
                "subject_contains": subjects,
                "body_or_subject_contains": body_terms,
                "has_attachment": has_attachment,
            },
            "exceptions": {"subject_contains": except_subjects},
            "actions": {
                "move_to_folder": move_to_folder or None,
                "copy_to_folder": copy_to_folder or None,
                "assign_categories": categories,
                "delete": delete,
                "stop_processing": stop_processing,
            },
            "executed_on_inbox": run_now,
        }, indent=2)

    try:
        return await bridge.call(_create, name, sender_contains, subject_contains,
                                 body_or_subject_contains, has_attachment,
                                 except_subject_contains, move_to_folder,
                                 copy_to_folder, assign_categories, delete,
                                 stop_processing, enabled, run_now, account)
    except Exception as e:
        return f"Error creating rule: {format_com_error(e)}"


# =====================================================================
# OUT OF OFFICE TOOLS
# =====================================================================

@mcp.tool()
async def get_out_of_office(account: str = "") -> str:
    """Check the current Out of Office (auto-reply) status.

    Returns whether Out of Office is currently enabled.

    Args:
        account: Optional. Account display name (or substring) to target.
            Default: primary account. Use list_accounts to see available accounts.

    Returns:
        JSON object with the OOF status.
    """
    def _get(outlook, namespace, account):
        store = _require_store(namespace, account)
        try:
            prop_tag = "http://schemas.microsoft.com/mapi/proptag/0x661D000B"
            oof_state = store.PropertyAccessor.GetProperty(prop_tag)
            return json.dumps({
                "out_of_office": bool(oof_state),
                "status": "on" if oof_state else "off",
            }, indent=2)
        except Exception:
            return json.dumps({
                "out_of_office": None,
                "status": "unknown",
                "note": "Could not read OOF property. Check Outlook settings directly.",
            }, indent=2)

    try:
        return await bridge.call(_get, account)
    except Exception as e:
        return f"Error checking OOF status: {format_com_error(e)}"


# =====================================================================
# Entry point
# =====================================================================

def main():
    logger.info("Starting Outlook Desktop MCP server...")
    bridge.start()
    logger.info("COM bridge ready. Starting MCP stdio transport...")
    try:
        mcp.run(transport="stdio")
    finally:
        bridge.stop()


if __name__ == "__main__":
    main()
