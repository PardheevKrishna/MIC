import os
import re
import json
import time
import traceback
from datetime import datetime, timedelta
from dataclasses import dataclass
from typing import Any, Dict, List, Optional

import win32com.client
from dotenv import load_dotenv
from jsonschema import validate, ValidationError
import google.generativeai as genai

# ---------------------------
# Config & Utilities
# ---------------------------
load_dotenv()

GEMINI_API_KEY = os.getenv("GEMINI_API_KEY")
GEMINI_MODEL = os.getenv("GEMINI_MODEL", "gemini-1.5-flash")
SENDER_WHITELIST = {s.strip().lower() for s in os.getenv("SENDER_WHITELIST", "").split(",") if s.strip()}
OUTLOOK_EMAIL = os.getenv("OUTLOOK_EMAIL", "")
ATTACHMENTS_DIR = os.getenv("ATTACHMENTS_DIR", os.path.join(os.getcwd(), "attachments"))

# Create attachments directory if it doesn't exist
os.makedirs(ATTACHMENTS_DIR, exist_ok=True)
os.makedirs("logs", exist_ok=True)

def log(msg: str):
    stamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with open("logs/agent.log", "a", encoding="utf-8") as f:
        f.write(f"[{stamp}] {msg}\n")
    print(msg)

def normalize_email(addr: str) -> str:
    return re.sub(r".*<([^>]+)>.*", r"\1", addr).strip().lower()

# ---------------------------
# Structured Plan Schema
# ---------------------------
PLAN_SCHEMA = {
    "type": "object",
    "required": ["summary", "priority", "actions"],
    "properties": {
        "summary": {"type": "string", "minLength": 1},
        "priority": {"type": "string", "enum": ["low", "normal", "high", "urgent"]},
        "actions": {
            "type": "array",
            "items": {
                "type": "object",
                "required": ["type"],
                "properties": {
                    "type": {"type": "string", "enum": [
                        "draft_reply",
                        "create_calendar_event",
                        "create_todo"
                    ]},
                    # draft_reply
                    "reply_body": {"type": "string"},
                    "reply_to_all": {"type": "boolean"},
                    "attachment_files": {"type": "array", "items": {"type": "string"}},  # filenames from safe dir
                    # create_calendar_event
                    "title": {"type": "string"},
                    "start_iso": {"type": "string"},      # ISO 8601
                    "end_iso": {"type": "string"},        # ISO 8601
                    "location": {"type": "string"},
                    "attendees": {"type": "array", "items": {"type": "string"}},
                    # create_todo
                    "todo_title": {"type": "string"},
                    "due_iso": {"type": "string"}
                },
                "additionalProperties": False
            }
        }
    },
    "additionalProperties": False
}

# ---------------------------
# Gemini client
# ---------------------------
class GeminiPlanner:
    def __init__(self, api_key: str, model_name: str):
        genai.configure(api_key=api_key)
        self.model = genai.GenerativeModel(model_name)

    def _get_available_files_list(self, max_files: int = 50) -> str:
        """List all files in attachments directory recursively"""
        if not os.path.exists(ATTACHMENTS_DIR):
            return 'Directory not found'
        
        files = []
        for root, dirs, filenames in os.walk(ATTACHMENTS_DIR):
            for filename in filenames:
                # Get relative path from ATTACHMENTS_DIR
                full_path = os.path.join(root, filename)
                rel_path = os.path.relpath(full_path, ATTACHMENTS_DIR)
                # Convert backslashes to forward slashes for consistency
                rel_path = rel_path.replace('\\', '/')
                files.append(rel_path)
                if len(files) >= max_files:
                    break
            if len(files) >= max_files:
                break
        
        if not files:
            return 'No files available'
        
        if len(files) > max_files:
            return ', '.join(files[:max_files]) + f' ... and {len(files) - max_files} more'
        
        return ', '.join(files)

    def plan(self, sender: str, subject: str, body: str, attachments: List[Dict[str, str]] = None) -> Dict[str, Any]:
        system = (
            "You are an executive assistant agent. "
            "Read the email and output ONLY a compact JSON object matching the schema provided. "
            "If dates are ambiguous, prefer next plausible future weekday at 10:00. "
            "You MUST stick to the allowed action types. "
            "Never include commentary; output pure JSON."
        )

        schema_hint = json.dumps(PLAN_SCHEMA, indent=2)
        prompt = f"""{system}

Email metadata:
- From: {sender}
- Subject: {subject}

Email body (between triple backticks):
```
{body}
```

Attachments in email: {', '.join([f"{a['filename']} ({a['size']})" for a in (attachments or [])]) if attachments else 'None'}

Available files you can attach from safe directory (use relative paths like 'folder/file.txt'):
{self._get_available_files_list()}

Constraints:
- Only choose from the allowed actions in the schema.
- Keep 'summary' ≤ 120 words.
- IMPORTANT: If the email asks a question or requests information, use 'draft_reply' to answer it directly.
- If the email asks to schedule a meeting, use 'create_calendar_event'.
- Only use 'create_todo' for tasks that require physical actions you cannot perform (like "buy groceries" or "call someone").
- Use ISO 8601 for times (e.g., 2025-11-03T10:00:00).
- When drafting replies, be helpful, professional, and provide complete answers.
- If the email mentions or asks about attachments, reference them in your reply.
- CRITICAL for attachments: Use the EXACT file paths from the available files list above (e.g., 'IMAGES_750/image001.jpg' not just 'IMAGES_750').
- Include specific FILE paths in 'attachment_files' array, NOT folder names.
- If asked for files from a folder, pick relevant individual files from that folder.

JSON Schema (for your reference):
{schema_hint}

Respond with JSON only.
"""
        resp = self.model.generate_content(prompt)
        text = resp.text.strip()
        # Try to locate JSON (model sometimes adds fences)
        json_str = text
        fence = re.search(r"\{.*\}", text, re.S)
        if fence:
            json_str = fence.group(0)
        try:
            data = json.loads(json_str)
            validate(instance=data, schema=PLAN_SCHEMA)
            return data
        except Exception as e:
            log(f"[PLANNER] JSON parse/validate error: {e}\n{text}")
            # Fallback: safe no-op, but keep a short summary
            return {
                "summary": f"Auto-summary failed strictly; manual review needed. Subject: {subject}",
                "priority": "normal",
                "actions": []
            }

# ---------------------------
# Outlook bridge
# ---------------------------
@dataclass
class MailItem:
    message_id: str
    sender: str
    subject: str
    body: str
    received: datetime
    outlook_item: Any  # Store the actual Outlook mail item
    attachments: List[Dict[str, str]] = None  # List of {filename, mimetype, size}

class OutlookClient:
    def __init__(self):
        self.outlook = None
        self.namespace = None
        self.inbox = None
        self.calendar = None
        self._connect()

    def _connect(self):
        """Connect to Outlook via COM"""
        try:
            self.outlook = win32com.client.Dispatch("Outlook.Application")
            self.namespace = self.outlook.GetNamespace("MAPI")
            self.inbox = self.namespace.GetDefaultFolder(6)  # 6 = olFolderInbox
            self.calendar = self.namespace.GetDefaultFolder(9)  # 9 = olFolderCalendar
            log("[OUTLOOK] Successfully connected")
        except Exception as e:
            log(f"[OUTLOOK] Connection error: {e}")
            raise RuntimeError(
                "Failed to connect to Outlook. Make sure:\n"
                "1. Outlook desktop app (not Microsoft Store version) is installed\n"
                "2. You're using classic Outlook, not 'New Outlook for Windows'\n"
                "3. You have a configured email account in Outlook\n"
                "4. pywin32 is installed: pip install pywin32"
            )

    def fetch_candidates(self, max_items: int = 20) -> List[MailItem]:
        """Fetch unprocessed emails from whitelisted senders"""
        try:
            messages = self.inbox.Items
            messages.Sort("[ReceivedTime]", True)  # Descending
            
            mail_items = []
            count = 0
            
            for message in messages:
                if count >= max_items:
                    break
                
                try:
                    # Skip if not a mail item
                    if message.Class != 43:  # 43 = olMail
                        continue
                    
                    sender = normalize_email(message.SenderEmailAddress)
                    
                    # Check whitelist
                    if SENDER_WHITELIST and sender not in SENDER_WHITELIST:
                        continue
                    
                    # Check if already processed (has category "AI-Processed")
                    categories = message.Categories if message.Categories else ""
                    if "AI-Processed" in categories:
                        continue
                    
                    # Extract attachment info
                    attachments = []
                    if message.Attachments.Count > 0:
                        for attachment in message.Attachments:
                            size_kb = attachment.Size / 1024
                            attachments.append({
                                'filename': attachment.FileName,
                                'mimetype': 'application/octet-stream',  # Outlook doesn't expose MIME type easily
                                'size': f"{size_kb:.1f}KB"
                            })
                    
                    mail_items.append(MailItem(
                        message_id=message.EntryID,
                        sender=sender,
                        subject=message.Subject or "(No Subject)",
                        body=message.Body or "",
                        received=message.ReceivedTime,
                        outlook_item=message,
                        attachments=attachments if attachments else None
                    ))
                    
                    count += 1
                    
                except Exception as e:
                    log(f"[OUTLOOK] Error processing message: {e}")
                    continue
            
            return mail_items
        except Exception as error:
            log(f"[OUTLOOK] Error fetching messages: {error}")
            return []

    def tag_processed(self, item: MailItem, note: str):
        """Add AI-Processed category to message"""
        try:
            categories = item.outlook_item.Categories
            if categories:
                item.outlook_item.Categories = categories + ", AI-Processed"
            else:
                item.outlook_item.Categories = "AI-Processed"
            item.outlook_item.Save()
            log(f"[OUTLOOK] Tagged message '{item.subject}' as processed")
        except Exception as error:
            log(f"[OUTLOOK] Error tagging message: {error}")

    def draft_reply(self, item: MailItem, body: str, reply_to_all: bool, attachment_files: List[str] = None):
        """Send a reply email with optional attachments"""
        try:
            # Create reply
            if reply_to_all:
                reply = item.outlook_item.ReplyAll()
            else:
                reply = item.outlook_item.Reply()
            
            # Set body
            reply.Body = body
            
            # Attach files from safe directory
            if attachment_files:
                for filename in attachment_files:
                    # Normalize path separators (handle both / and \)
                    filename = filename.replace('/', os.sep).replace('\\', os.sep)
                    filepath = os.path.abspath(os.path.join(ATTACHMENTS_DIR, filename))
                    safe_root = os.path.abspath(ATTACHMENTS_DIR)
                    
                    # Security check: ensure file is in safe directory
                    if not filepath.startswith(safe_root + os.sep):
                        log(f"[OUTLOOK] ⚠️ Rejected attachment outside safe directory: {filename}")
                        continue
                    
                    if not os.path.isfile(filepath):
                        log(f"[OUTLOOK] ⚠️ Attachment file not found: {filename} (looked in {filepath})")
                        continue
                    
                    # Add attachment
                    reply.Attachments.Add(filepath)
                    log(f"[OUTLOOK] 📎 Attached file: {filename}")
            
            # Send the reply immediately
            reply.Send()
            
            log(f"[OUTLOOK] ✅ Sent reply to {item.sender} for message '{item.subject}'")
            return True
        except Exception as error:
            log(f"[OUTLOOK] Error sending reply: {error}")
            return False

    def create_calendar_event(self, title: str, start_iso: str, end_iso: str,
                              location: Optional[str], attendees: List[str]) -> bool:
        """Create an Outlook Calendar event"""
        try:
            appointment = self.outlook.CreateItem(1)  # 1 = olAppointmentItem
            appointment.Subject = title
            
            # Parse ISO dates
            start_dt = datetime.fromisoformat(start_iso.replace('Z', '+00:00'))
            end_dt = datetime.fromisoformat(end_iso.replace('Z', '+00:00'))
            
            appointment.Start = start_dt
            appointment.End = end_dt
            
            if location:
                appointment.Location = location
            
            # Add attendees
            if attendees:
                for email in attendees:
                    appointment.Recipients.Add(email)
                appointment.MeetingStatus = 1  # 1 = olMeeting (makes it a meeting with attendees)
            
            appointment.Save()
            
            if attendees:
                appointment.Send()  # Send meeting invites
                log(f"[CALENDAR] Created meeting and sent invites: {title}")
            else:
                log(f"[CALENDAR] Created appointment: {title}")
            
            return True
        except Exception as error:
            log(f"[CALENDAR] Error creating event: {error}")
            return False

# ---------------------------
# Action Registry
# ---------------------------
class ActionRegistry:
    def __init__(self, outlook: OutlookClient):
        self.outlook = outlook

    def run(self, item: MailItem, plan: Dict[str, Any]):
        results = []
        for idx, action in enumerate(plan.get("actions", []), start=1):
            t = action.get("type")
            try:
                if t == "draft_reply":
                    ok = self.outlook.draft_reply(
                        item,
                        body=action.get("reply_body", "Not specified."),
                        reply_to_all=bool(action.get("reply_to_all", False)),
                        attachment_files=action.get("attachment_files") or []
                    )
                    results.append((t, ok))
                elif t == "create_calendar_event":
                    ok = self.outlook.create_calendar_event(
                        title=action.get("title", "Untitled"),
                        start_iso=action.get("start_iso"),
                        end_iso=action.get("end_iso"),
                        location=action.get("location"),
                        attendees=action.get("attendees") or [],
                    )
                    results.append((t, ok))
                elif t == "create_todo":
                    todo = {
                        "title": action.get("todo_title", "Untitled"),
                        "due_iso": action.get("due_iso"),
                        "source_mail_subject": item.subject,
                        "source_mail_sender": item.sender,
                        "created_at": datetime.now().isoformat(),
                    }
                    with open("logs/todos.jsonl", "a", encoding="utf-8") as f:
                        f.write(json.dumps(todo) + "\n")
                    results.append((t, True))
                else:
                    results.append((t, False))
            except Exception as e:
                log(f"[ACTION] {t} error: {e}")
                results.append((t, False))
        return results

# ---------------------------
# Agent loop (poll)
# ---------------------------
class MailAgent:
    def __init__(self):
        assert GEMINI_API_KEY, "GEMINI_API_KEY not set"
        self.planner = GeminiPlanner(GEMINI_API_KEY, GEMINI_MODEL)
        self.outlook = OutlookClient()
        self.actions = ActionRegistry(self.outlook)

    def process_once(self, batch_size: int = 10):
        candidates = self.outlook.fetch_candidates(max_items=batch_size)
        log(f"[MAIN] Found {len(candidates)} unprocessed emails")
        for m in candidates:
            try:
                log(f"Processing: From={m.sender} Subject='{m.subject}'")
                if m.attachments:
                    log(f"  📎 Attachments: {', '.join([a['filename'] for a in m.attachments])}")
                plan = self.planner.plan(m.sender, m.subject, m.body, m.attachments)
                log(f"Plan:\n{json.dumps(plan, indent=2, ensure_ascii=False)}")
                results = self.actions.run(m, plan)
                note = f"Summary: {plan.get('summary','')[:120]} | Results: {results}"
                self.outlook.tag_processed(m, note)
            except Exception as e:
                log(f"[PROCESS] error: {e}\n{traceback.format_exc()}")

def main():
    agent = MailAgent()
    # Simple poller: every 60s. For production, use Windows Task Scheduler to run every 2–5 min.
    log("[MAIN] Mail agent started. Polling every 60 seconds. Press Ctrl+C to stop.")
    while True:
        try:
            agent.process_once(batch_size=10)
        except Exception as e:
            log(f"[MAIN] iteration error: {e}")
        time.sleep(60)

if __name__ == "__main__":
    main()
