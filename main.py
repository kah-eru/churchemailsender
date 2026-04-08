import base64
import csv
import json
import mimetypes
import os
import re
import smtplib
import threading
import time
import uuid
from datetime import datetime, timezone
from zoneinfo import ZoneInfo
from email import encoders
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

import webview

import db_manager

# ── Python API exposed to JavaScript ─────────────────────────────────────────

class Api:
    def get_contacts(self):
        rows = db_manager.get_contacts()
        group_map = db_manager.get_contact_groups()
        family_map = db_manager.get_contact_families()
        return [
            {"id": r[0], "name": r[1], "email": r[2], "category": r[3], "family_name": r[4],
             "families": family_map.get(r[0], []), "groups": group_map.get(r[0], [])}
            for r in rows
        ]

    def add_contact(self, name, email, category, family_id):
        try:
            fid = int(family_id) if family_id else None
            db_manager.add_contact(name, email, category, fid)
            return {"ok": True}
        except Exception as e:
            return {"ok": False, "error": str(e)}

    def update_contact(self, contact_id, name, email, category, family_id):
        try:
            fid = int(family_id) if family_id else None
            db_manager.update_contact(int(contact_id), name, email, category, fid)
            return {"ok": True}
        except Exception as e:
            return {"ok": False, "error": str(e)}

    def delete_contacts(self, ids):
        for cid in ids:
            db_manager.delete_contact(int(cid))
        return {"ok": True}

    def get_families(self):
        families = db_manager.get_families()
        result = []
        for fid, fname in families:
            members = db_manager.get_family_members_via_junction(fid)
            result.append({
                "id": fid,
                "name": fname,
                "members": [{"id": m[0], "name": m[1], "email": m[2]} for m in members],
            })
        return result

    def add_family_member(self, family_id, contact_id):
        try:
            db_manager.add_family_member(int(family_id), int(contact_id))
            return {"ok": True}
        except Exception as e:
            return {"ok": False, "error": str(e)}

    def remove_family_member(self, family_id, contact_id):
        db_manager.remove_family_member(int(family_id), int(contact_id))
        return {"ok": True}

    def add_family(self, name):
        try:
            db_manager.add_family(name)
            return {"ok": True}
        except Exception as e:
            return {"ok": False, "error": str(e)}

    def rename_family(self, family_id, new_name):
        try:
            db_manager.rename_family(int(family_id), new_name)
            return {"ok": True}
        except Exception as e:
            return {"ok": False, "error": str(e)}

    def delete_family(self, family_id):
        db_manager.delete_family(int(family_id))
        return {"ok": True}

    # ── Settings ──

    def get_settings(self):
        email = db_manager.get_setting("sender_email") or ""
        password = db_manager.get_setting("app_password") or ""
        timezone = db_manager.get_setting("timezone") or "US/Eastern"
        masked = ("*" * (len(password) - 4) + password[-4:]) if len(password) > 4 else "*" * len(password)
        return {"email": email, "app_password_masked": masked, "has_password": bool(password), "timezone": timezone}

    def save_settings(self, email, app_password):
        try:
            db_manager.set_setting("sender_email", email)
            if app_password:
                db_manager.set_setting("app_password", app_password)
            return {"ok": True}
        except Exception as e:
            return {"ok": False, "error": str(e)}

    def get_ui_setting(self, key):
        return db_manager.get_setting("ui_" + key) or ""

    def set_ui_setting(self, key, value):
        db_manager.set_setting("ui_" + key, value)

    def save_timezone(self, timezone):
        try:
            db_manager.set_setting("timezone", timezone)
            return {"ok": True}
        except Exception as e:
            return {"ok": False, "error": str(e)}

    def test_email_connection(self, email, app_password):
        try:
            with smtplib.SMTP("smtp.gmail.com", 587) as server:
                server.starttls()
                server.login(email, app_password)
            return {"ok": True}
        except Exception as e:
            return {"ok": False, "error": str(e)}

    # ── File Picker ──

    def pick_file(self):
        result = webview.windows[0].create_file_dialog(webview.FileDialog.OPEN, allow_multiple=True)
        if not result:
            return []
        files = []
        for path in result:
            files.append({"name": os.path.basename(path), "path": path, "size": os.path.getsize(path)})
        return files

    # ── Email Dispatch ──

    @staticmethod
    def _extract_inline_images(html_body):
        """Find base64 data-URI images in HTML, return (new_html, [(cid, mime_type, raw_bytes), ...])."""
        images = []
        def replacer(match):
            mime_type = match.group(1)
            b64_data = match.group(2)
            cid = uuid.uuid4().hex
            raw_bytes = base64.b64decode(b64_data)
            images.append((cid, mime_type, raw_bytes))
            return f'src="cid:{cid}"'
        new_html = re.sub(r'src="data:(image/[^;]+);base64,([^"]+)"', replacer, html_body)
        return new_html, images

    @staticmethod
    def _build_message(sender_email, to_email, subject, plain_text, processed_html, image_data, attachment_paths):
        """Build a fresh MIME message for one recipient."""

        msg = MIMEMultipart("mixed")
        msg["From"] = sender_email
        msg["To"] = to_email
        msg["Subject"] = subject

        # Related part (HTML + inline images)
        related = MIMEMultipart("related")
        alt = MIMEMultipart("alternative")
        alt.attach(MIMEText(plain_text, "plain"))
        alt.attach(MIMEText(processed_html, "html"))
        related.attach(alt)

        # Create fresh MIMEImage per recipient for each inline image
        for cid, mime_type, raw_bytes in image_data:
            subtype = mime_type.split("/", 1)[1] if "/" in mime_type else "png"
            img = MIMEImage(raw_bytes, _subtype=subtype)
            img.add_header("Content-ID", f"<{cid}>")
            img.add_header("Content-Disposition", "inline", filename=f"{cid}.{subtype}")
            related.attach(img)

        msg.attach(related)

        # File attachments with proper MIME type detection
        for fpath in (attachment_paths or []):
            if not os.path.isfile(fpath):
                continue
            filename = os.path.basename(fpath)
            content_type, _ = mimetypes.guess_type(fpath)
            if content_type is None:
                content_type = "application/octet-stream"
            maintype, subtype = content_type.split("/", 1)

            with open(fpath, "rb") as f:
                file_data = f.read()

            if maintype == "image":
                part = MIMEImage(file_data, _subtype=subtype)
            else:
                part = MIMEBase(maintype, subtype)
                part.set_payload(file_data)
                encoders.encode_base64(part)

            part.add_header("Content-Disposition", "attachment", filename=filename)
            msg.attach(part)

        return msg

    @staticmethod
    def _send_to_recipients(sender_email, app_password, recipients, subject, plain_text, html_body, attachment_paths):
        """Shared email-sending logic used by both dispatch and scheduler."""
        processed_html, image_data = Api._extract_inline_images(html_body)
        sent, failed = 0, 0
        error = None
        try:
            with smtplib.SMTP("smtp.gmail.com", 587) as server:
                server.starttls()
                server.login(sender_email, app_password)
                for name, email_addr in recipients:
                    try:
                        msg = Api._build_message(sender_email, email_addr, subject, plain_text, processed_html, image_data, attachment_paths)
                        server.sendmail(sender_email, email_addr, msg.as_string())
                        sent += 1
                    except Exception as e:
                        print(f"[Email] Failed to send to {email_addr}: {e}")
                        failed += 1
        except Exception as e:
            error = str(e)
        return {"sent": sent, "failed": failed, "error": error}

    def dispatch_emails(self, subject, html_body, plain_text, contact_ids, attachment_paths=None, target_type="all", target_id=None, manual_emails=None):
        sender_email = db_manager.get_setting("sender_email")
        app_password = db_manager.get_setting("app_password")
        if not sender_email or not app_password:
            return {"sent": 0, "failed": 0, "error": "Email credentials not configured. Go to Settings tab."}

        recipients = []
        if contact_ids:
            recipients += self.resolve_recipients("custom", contact_ids=contact_ids)
        if target_type and target_type != "manual":
            recipients += self.resolve_recipients(target_type, target_id=target_id)
        if manual_emails:
            for email in manual_emails:
                recipients.append((email, email))

        # Deduplicate by email address
        seen = set()
        unique = []
        for name, email in recipients:
            if email not in seen:
                seen.add(email)
                unique.append((name, email))
        recipients = unique

        if not recipients:
            return {"sent": 0, "failed": 0, "error": "No recipients specified."}

        result = self._send_to_recipients(sender_email, app_password, recipients, subject, plain_text, html_body, attachment_paths or [])
        target_desc = target_type or "manual"
        if manual_emails:
            target_desc = f"{len(manual_emails)} manual" + (f" + {target_type}" if target_type and target_type != "manual" else "")
        db_manager.log_email(subject, target_desc, len(recipients), result["sent"], result["failed"])
        return result

    # ── Groups ──

    def get_groups(self):
        groups = db_manager.get_groups()
        result = []
        for gid, gname in groups:
            members = db_manager.get_group_members(gid)
            result.append({
                "id": gid,
                "name": gname,
                "members": [{"id": m[0], "name": m[1], "email": m[2]} for m in members],
            })
        return result

    def add_group(self, name):
        try:
            db_manager.add_group(name)
            return {"ok": True}
        except Exception as e:
            return {"ok": False, "error": str(e)}

    def rename_group(self, group_id, new_name):
        try:
            db_manager.rename_group(int(group_id), new_name)
            return {"ok": True}
        except Exception as e:
            return {"ok": False, "error": str(e)}

    def delete_group(self, group_id):
        db_manager.delete_group(int(group_id))
        return {"ok": True}

    def add_group_member(self, group_id, contact_id):
        try:
            db_manager.add_group_member(int(group_id), int(contact_id))
            return {"ok": True}
        except Exception as e:
            return {"ok": False, "error": str(e)}

    def remove_group_member(self, group_id, contact_id):
        db_manager.remove_group_member(int(group_id), int(contact_id))
        return {"ok": True}

    def add_family_to_group(self, group_id, family_id):
        members = db_manager.get_contacts_by_family(int(family_id))
        added = 0
        for m in members:
            try:
                db_manager.add_group_member(int(group_id), m[0])
                added += 1
            except Exception:
                pass
        return {"ok": True, "added": added}

    # ── Scheduled Emails ──

    def schedule_email(self, subject, html_body, plain_text, target_type, target_id, contact_ids, attachment_paths, scheduled_at, recurrence=None, manual_emails=None):
        try:
            tid = int(target_id) if target_id else None
            db_manager.schedule_email(subject, html_body, plain_text, target_type, tid, contact_ids or [], attachment_paths or [], scheduled_at, recurrence=recurrence, manual_emails=manual_emails or [])
            return {"ok": True}
        except Exception as e:
            return {"ok": False, "error": str(e)}

    def get_scheduled_emails(self):
        rows = db_manager.get_scheduled_emails()
        return [
            {"id": r[0], "subject": r[1], "target_type": r[2], "target_id": r[3],
             "scheduled_at": r[4], "status": r[5], "sent_at": r[6], "result": r[7],
             "recurrence": json.loads(r[8]) if r[8] else None}
            for r in rows
        ]

    def get_scheduled_emails_with_recipients(self):
        rows = db_manager.get_scheduled_emails()
        result = []
        for r in rows:
            entry = {"id": r[0], "subject": r[1], "target_type": r[2], "target_id": r[3],
                     "scheduled_at": r[4], "status": r[5], "sent_at": r[6], "result": r[7],
                     "recurrence": json.loads(r[8]) if r[8] else None}
            contact_ids = json.loads(r[9]) if r[9] else []
            manual_emails = json.loads(r[10]) if r[10] else []
            recipients = self.resolve_recipients(entry["target_type"], entry["target_id"], contact_ids)
            for me in manual_emails:
                if not any(rec[1] == me for rec in recipients):
                    recipients.append(("", me))
            entry["recipients"] = [{"name": n, "email": e} for n, e in recipients]
            result.append(entry)
        return result

    def get_scheduled_email_detail(self, email_id):
        r = db_manager.get_scheduled_email_by_id(int(email_id))
        if not r:
            return None
        entry = {"id": r[0], "subject": r[1], "html_body": r[2], "plain_text": r[3],
                 "target_type": r[4], "target_id": r[5], "scheduled_at": r[8],
                 "status": r[9], "sent_at": r[10], "result": r[11],
                 "recurrence": json.loads(r[12]) if r[12] else None}
        contact_ids = json.loads(r[6]) if r[6] else []
        manual_emails = json.loads(r[13]) if r[13] else []
        recipients = self.resolve_recipients(entry["target_type"], entry["target_id"], contact_ids)
        for me in manual_emails:
            if not any(rec[1] == me for rec in recipients):
                recipients.append(("", me))
        entry["recipients"] = [{"name": n, "email": e} for n, e in recipients]
        return entry

    def cancel_scheduled_email(self, email_id):
        db_manager.cancel_scheduled_email(int(email_id))
        return {"ok": True}

    def resolve_recipients(self, target_type, target_id=None, contact_ids=None):
        if target_type == "group" and target_id:
            members = db_manager.get_group_members(int(target_id))
            return [(m[1], m[2]) for m in members]

        contacts = db_manager.get_contacts()
        if target_type == "all":
            return [(r[1], r[2]) for r in contacts]
        elif target_type == "family":
            return [(r[1], r[2]) for r in contacts if r[3] == "Family"]
        elif target_type == "single":
            return [(r[1], r[2]) for r in contacts if r[3] == "Single"]
        elif target_type == "custom" and contact_ids:
            id_set = set(int(c) for c in contact_ids)
            return [(r[1], r[2]) for r in contacts if r[0] in id_set]
        return []

    # ── Templates ──

    def get_templates(self):
        rows = db_manager.get_templates()
        return [{"id": r[0], "name": r[1], "subject": r[2], "html_body": r[3],
                 "recipients": json.loads(r[4]) if r[4] else None} for r in rows]

    def save_template(self, name, subject, html_body, recipients=None):
        try:
            db_manager.save_template(name, subject, html_body, recipients=recipients)
            return {"ok": True}
        except Exception as e:
            return {"ok": False, "error": str(e)}

    def update_template(self, template_id, subject, html_body, recipients=None):
        try:
            db_manager.update_template(int(template_id), subject, html_body, recipients=recipients)
            return {"ok": True}
        except Exception as e:
            return {"ok": False, "error": str(e)}

    def delete_template(self, template_id):
        db_manager.delete_template(int(template_id))
        return {"ok": True}

    # ── CSV Import/Export ──

    def import_csv(self):
        result = webview.windows[0].create_file_dialog(webview.FileDialog.OPEN, file_types=("CSV Files (*.csv)",))
        if not result:
            return {"ok": False, "error": "No file selected."}
        filepath = result[0]
        added, skipped = 0, 0
        try:
            family_cache = {fname: fid for fid, fname in db_manager.get_families()}
            with open(filepath, "r", encoding="utf-8-sig") as f:
                reader = csv.DictReader(f)
                for row in reader:
                    name = row.get("name", "").strip()
                    email = row.get("email", "").strip()
                    category = row.get("category", "Single").strip()
                    family_name = row.get("family_name", "").strip()
                    if not name or not email:
                        skipped += 1
                        continue
                    if category not in ("Family", "Single"):
                        category = "Single"
                    family_id = None
                    if category == "Family" and family_name:
                        if family_name not in family_cache:
                            try:
                                db_manager.add_family(family_name)
                                families = db_manager.get_families()
                                family_cache = {fname: fid for fid, fname in families}
                            except Exception:
                                pass
                        family_id = family_cache.get(family_name)
                    try:
                        db_manager.add_contact(name, email, category, family_id)
                        added += 1
                    except Exception:
                        skipped += 1
            return {"ok": True, "added": added, "skipped": skipped}
        except Exception as e:
            return {"ok": False, "error": str(e)}

    def export_csv(self):
        result = webview.windows[0].create_file_dialog(webview.FileDialog.SAVE, save_filename="contacts.csv")
        if not result:
            return {"ok": False, "error": "No location selected."}
        filepath = result if isinstance(result, str) else result[0]
        try:
            contacts = db_manager.get_contacts()
            with open(filepath, "w", newline="", encoding="utf-8") as f:
                writer = csv.writer(f)
                writer.writerow(["name", "email", "category", "family_name"])
                for r in contacts:
                    writer.writerow([r[1], r[2], r[3], r[4]])
            return {"ok": True, "count": len(contacts)}
        except Exception as e:
            return {"ok": False, "error": str(e)}

    # ── Email History ──

    def get_email_history(self):
        rows = db_manager.get_email_history()
        return [
            {"id": r[0], "subject": r[1], "target": r[2], "recipients": r[3],
             "sent": r[4], "failed": r[5], "sent_at": r[6]}
            for r in rows
        ]


# ── Background Scheduler ────────────────────────────────────────────────────

def run_scheduler(api_instance):
    """Check for due scheduled emails every 60 seconds."""
    print("[Scheduler] Started — checking every 60s")
    while True:
        try:
            tz_name = db_manager.get_setting("timezone") or "US/Eastern"
            try:
                tz = ZoneInfo(tz_name)
            except Exception:
                tz = ZoneInfo("US/Eastern")
            now_local = datetime.now(tz).replace(tzinfo=None).isoformat()
            due = db_manager.get_due_emails(now_iso=now_local)
            if due:
                print(f"[Scheduler] {len(due)} email(s) due at {now_local} ({tz_name})")
            for row in due:
                eid, subject, html_body, plain_text, target_type, target_id, contact_ids_json, attach_json, recurrence_json, manual_emails_json, scheduled_at_str = row
                contact_ids = json.loads(contact_ids_json) if contact_ids_json else []
                attachment_paths = json.loads(attach_json) if attach_json else []
                recurrence = json.loads(recurrence_json) if recurrence_json else None
                manual_emails = json.loads(manual_emails_json) if manual_emails_json else []

                recipients = api_instance.resolve_recipients(target_type, target_id, contact_ids)
                for email in manual_emails:
                    if not any(r[1] == email for r in recipients):
                        recipients.append((email, email))

                if not recipients:
                    db_manager.update_email_status(eid, "failed", {"sent": 0, "failed": 0, "error": "No recipients"})
                    continue

                sender_email = db_manager.get_setting("sender_email")
                app_password = db_manager.get_setting("app_password")
                if not sender_email or not app_password:
                    db_manager.update_email_status(eid, "failed", {"sent": 0, "failed": 0, "error": "No credentials"})
                    continue

                result = Api._send_to_recipients(sender_email, app_password, recipients, subject, plain_text, html_body, attachment_paths)
                status = "failed" if result["error"] else "sent"
                db_manager.update_email_status(eid, status, result)
                print(f"[Scheduler] Email '{subject}' — {status} (sent:{result['sent']}, failed:{result['failed']})")
                db_manager.log_email(subject, target_type, len(recipients), result["sent"], result["failed"])

                # Schedule next occurrence for recurring emails
                if recurrence and recurrence.get("type", "once") != "once":
                    try:
                        current_dt = datetime.fromisoformat(scheduled_at_str)
                        next_dt = db_manager.compute_next_occurrence(current_dt, recurrence)
                        if next_dt:
                            db_manager.schedule_email(
                                subject, html_body, plain_text, target_type, target_id,
                                contact_ids, attachment_paths, next_dt.isoformat(),
                                recurrence=recurrence, manual_emails=manual_emails
                            )
                            print(f"[Scheduler] Recurring email '{subject}' next at {next_dt.isoformat()}")
                    except Exception as e:
                        print(f"[Scheduler] Error scheduling next recurrence: {e}")
        except Exception as e:
            print(f"[Scheduler] Error: {e}")
        time.sleep(30)


# ── Inline HTML/CSS/JS Frontend ──────────────────────────────────────────────

HTML = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Church Roster</title>

<style>
  /* ── Theme variables ── */
  :root {
    --bg: #1a1a1a; --panel: #242424; --surface: #1e1e1e; --border: #3a3a3a;
    --text: #e0e0e0; --text-muted: #888; --text-dim: #555;
    --accent: #5a9e8f; --accent-hover: #6db5a5;
    --danger: #c44; --danger-hover: #d55;
    --success: #2d7a4f; --scrollbar: #4a4a4a;
    --tab-bg: #2e2e2e; --tab-text: #888;
    --card-bg: #2a2a2a; --row-hover: #303030;
    --toolbar-bg: #2a2a2a;
  }
  body.light {
    --bg: #f5f5f5; --panel: #ffffff; --surface: #fafafa; --border: #ddd;
    --text: #1a1a1a; --text-muted: #666; --text-dim: #999;
    --accent: #3d8b7a; --accent-hover: #2f7566;
    --danger: #c44; --danger-hover: #d55;
    --success: #2d7a4f; --scrollbar: #bbb;
    --tab-bg: #eee; --tab-text: #777;
    --card-bg: #f0f0f0; --row-hover: #e8e8e8;
    --toolbar-bg: #f0f0f0;
  }

  * { margin: 0; padding: 0; box-sizing: border-box; }
  body {
    font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
    background: var(--bg); color: var(--text); height: 100vh; overflow: hidden;
    transition: background 0.2s, color 0.2s;
  }
  .app { display: flex; height: 100vh; padding: 10px; gap: 0; }

  /* ── Panels ── */
  .left-panel, .right-panel {
    background: var(--panel); border-radius: 10px; padding: 16px; display: flex; flex-direction: column;
    transition: background 0.2s;
    min-width: 300px; overflow: hidden;
  }
  .left-panel { flex: 1 1 50%; }
  .right-panel { flex: 1 1 50%; }

  /* ── Resize divider ── */
  .panel-divider {
    width: 12px; flex-shrink: 0; cursor: col-resize; position: relative;
    display: flex; align-items: center; justify-content: center;
    -webkit-user-select: none; user-select: none; z-index: 10;
    touch-action: none; -webkit-app-region: no-drag;
  }
  .panel-divider::after {
    content: ''; display: block; width: 3px; height: 40px; border-radius: 2px;
    background: var(--border); transition: background 0.15s, height 0.15s;
    pointer-events: none;
  }
  .panel-divider:hover::after, .panel-divider.dragging::after {
    background: var(--accent); height: 60px;
  }

  /* ── Tabs ── */
  .tab-bar { display: flex; gap: 4px; margin-bottom: 12px; }
  .tab-btn {
    flex: 1; padding: 8px 0; border: none; border-radius: 6px; cursor: pointer;
    background: var(--tab-bg); color: var(--tab-text); font-size: 13px; font-weight: 600; transition: all 0.2s;
  }
  .tab-btn.active { background: var(--accent); color: #fff; }
  .tab-content { display: none; flex-direction: column; flex: 1; overflow: hidden; }
  .tab-content.active { display: flex; }

  /* ── Schedule Calendar ── */
  .sched-day-row { margin-bottom: 4px; border-radius: 6px; background: var(--card-bg); padding: 8px 10px; }
  .sched-day-row.sched-day-today { border: 2px solid var(--accent); }
  .sched-day-header { display: flex; justify-content: space-between; align-items: center; margin-bottom: 6px; }
  .sched-day-label { font-weight: 700; font-size: 13px; color: var(--text); }
  .sched-day-today .sched-day-label { color: var(--accent); }
  .sched-day-cards { display: flex; flex-direction: column; gap: 4px; }
  .sched-email-card {
    background: var(--surface); border-radius: 5px; padding: 7px 10px; font-size: 12px;
    border-left: 3px solid var(--accent); position: relative; cursor: default;
  }
  .sched-email-card[data-status="sent"] { border-left-color: var(--success); }
  .sched-email-card[data-status="failed"], .sched-email-card[data-status="cancelled"] { border-left-color: var(--danger); }
  .sched-card-subject { font-weight: 600; color: var(--text); white-space: nowrap; overflow: hidden; text-overflow: ellipsis; margin-bottom: 2px; }
  .sched-card-meta { color: var(--text-muted); font-size: 11px; display: flex; gap: 8px; align-items: center; flex-wrap: wrap; }
  .sched-card-status { font-weight: 700; text-transform: uppercase; font-size: 10px; }
  .sched-tooltip {
    display: none; position: absolute; z-index: 100; background: var(--card-bg); border: 1px solid var(--accent);
    border-radius: 6px; padding: 10px 12px; font-size: 12px; color: var(--text); min-width: 220px; max-width: 320px;
    box-shadow: 0 4px 12px rgba(0,0,0,0.25); left: 50%; top: 100%; transform: translateX(-50%); margin-top: 4px;
    white-space: normal; line-height: 1.5;
  }
  .sched-email-card:hover .sched-tooltip { display: block; }
  .sched-add-btn {
    background: none; border: 1px dashed var(--text-muted); border-radius: 5px; padding: 3px; margin-top: 4px;
    color: var(--text-muted); cursor: pointer; font-size: 16px; width: 100%; text-align: center; transition: all 0.2s;
  }
  .sched-add-btn:hover { border-color: var(--accent); color: var(--accent); }
  .sched-empty-day { color: var(--text-muted); font-size: 11px; font-style: italic; }
  .sched-email-card.clickable { cursor: pointer; }
  .sched-email-card.clickable:hover { filter: brightness(1.1); }

  /* ── Scrollable lists ── */
  .list-area { flex: 1; overflow-y: auto; margin-bottom: 10px; border-radius: 6px; background: var(--surface); padding: 6px; }
  .list-area::-webkit-scrollbar { width: 6px; }
  .list-area::-webkit-scrollbar-thumb { background: var(--scrollbar); border-radius: 3px; }

  .contact-row, .family-card {
    display: flex; align-items: center; gap: 8px; padding: 7px 10px; border-radius: 5px; font-size: 13px;
  }
  .contact-row { position: relative; }
  .contact-row:hover { background: var(--row-hover); }
  .contact-row input[type="checkbox"] { accent-color: var(--accent); width: 15px; height: 15px; cursor: pointer; }
  .contact-header {
    display: flex; align-items: center; gap: 8px; padding: 4px 10px; font-size: 11px;
    font-weight: 700; color: var(--text-muted); text-transform: uppercase; letter-spacing: 0.5px;
    border-bottom: 1px solid var(--border); margin-bottom: 2px;
  }
  .contact-header .h-check { width: 15px; }
  .contact-header .h-name { flex: 2; min-width: 0; }
  .contact-header .h-email { flex: 2; min-width: 0; }
  .contact-header .h-families { flex: 2; min-width: 0; }
  .contact-header .h-groups { flex: 2; min-width: 0; }
  .contact-header .h-edit { width: 30px; }
  .contact-row .name { flex: 2; font-weight: 500; min-width: 0; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
  .contact-row .email { flex: 2; color: var(--text-muted); font-size: 12px; min-width: 0; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
  .contact-row .col-families { flex: 2; display: flex; flex-wrap: wrap; gap: 3px; min-width: 0; }
  .contact-row .col-groups { flex: 2; display: flex; flex-wrap: wrap; gap: 3px; min-width: 0; }
  .contact-tag {
    display: inline-block; font-size: 10px; padding: 1px 7px; border-radius: 8px;
    white-space: nowrap; line-height: 1.4;
  }
  .contact-tag.fam-tag { background: var(--accent); color: #fff; opacity: 0.85; }
  .contact-tag.grp-tag { background: #b8860b; color: #fff; opacity: 0.85; }
  .contact-edit-btn {
    background: none; border: none; cursor: pointer; color: var(--text-muted); font-size: 14px;
    padding: 2px 5px; border-radius: 4px; transition: color 0.15s;
  }
  .contact-edit-btn:hover { color: var(--accent); }


  .family-card {
    flex-direction: column; align-items: flex-start; background: var(--card-bg); margin-bottom: 6px; padding: 10px 12px;
  }
  .family-card .fam-header { display: flex; width: 100%; justify-content: space-between; align-items: center; }
  .family-card .fam-name { font-weight: 600; font-size: 14px; }
  .family-card .fam-members { font-size: 12px; color: var(--text-muted); margin-top: 4px; display: flex; flex-direction: column; gap: 1px; }
  .fam-members .member-row { padding: 2px 0; }

  .empty-msg { text-align: center; color: var(--text-dim); padding: 30px 0; font-size: 13px; }

  /* ── Groups split layout ── */
  .groups-split { display: flex; gap: 10px; flex: 1; min-height: 0; }
  .groups-left { flex: 1; display: flex; flex-direction: column; min-height: 0; }
  .groups-right { flex: 1; display: flex; flex-direction: column; min-height: 0; border-left: 1px solid var(--border); padding-left: 10px; }
  .groups-right-header {
    display: flex; justify-content: space-between; align-items: center;
    margin-bottom: 6px; min-height: 28px;
  }
  .group-detail-title { font-weight: 700; font-size: 15px; color: var(--text); }
  .group-item {
    display: flex; align-items: center; justify-content: space-between;
    padding: 8px 10px; border-radius: 6px; cursor: pointer; font-size: 13px;
    margin-bottom: 2px; transition: background 0.1s;
  }
  .group-item:hover { background: var(--row-hover); }
  .group-item.active { background: var(--accent); color: #fff; }
  .group-item.active .group-count { color: rgba(255,255,255,0.7); }
  .group-item .group-name { font-weight: 500; }
  .group-item .group-count { font-size: 11px; color: var(--text-muted); }
  .group-member-row {
    display: flex; align-items: center; gap: 8px; padding: 6px 10px;
    border-radius: 5px; font-size: 13px;
  }
  .group-member-row:hover { background: var(--row-hover); }
  .group-member-row .gm-name { font-weight: 500; flex: 1; }
  .group-member-row .gm-email { color: var(--text-muted); font-size: 12px; }

  /* ── Form inputs ── */
  .form-row { display: flex; gap: 6px; margin-bottom: 6px; }
  .form-row input, .form-row select { flex: 1; }
  input[type="text"], input[type="password"], select {
    padding: 8px 10px; border: 1px solid var(--border); border-radius: 6px;
    background: var(--surface); color: var(--text); font-size: 13px; outline: none;
    transition: background 0.2s, color 0.2s, border-color 0.2s;
  }
  input[type="text"]:focus, select:focus { border-color: var(--accent); }
  input[type="text"]::placeholder { color: var(--text-dim); }
  select { cursor: pointer; }

  /* ── Buttons ── */
  .btn {
    padding: 8px 16px; border: none; border-radius: 6px; cursor: pointer;
    font-size: 13px; font-weight: 600; transition: all 0.15s;
  }
  .btn-primary { background: var(--accent); color: #fff; }
  .btn-primary:hover { background: var(--accent-hover); }
  .btn-danger { background: var(--danger); color: #fff; }
  .btn-danger:hover { background: var(--danger-hover); }
  .btn-dispatch {
    width: 100%; padding: 10px; margin-top: 8px; font-size: 14px; font-weight: 700;
    background: var(--accent); color: #fff; border: none; border-radius: 8px; cursor: pointer;
  }
  .btn-dispatch:hover { background: var(--accent-hover); }
  .btn-sm { padding: 5px 10px; font-size: 11px; }
  .btn-success { background: var(--success); }
  .btn-success:hover { background: var(--success); filter: brightness(1.1); }
  .btn-nowrap { white-space: nowrap; }
  .btn-row { display: flex; gap: 6px; }
  .btn-dispatch:disabled { opacity: 0.6; cursor: not-allowed; }

  /* ── Recurrence controls ── */
  .recurrence-row { display: flex; gap: 6px; align-items: center; flex-wrap: wrap; }
  .day-picker { display: flex; gap: 3px; }
  .day-btn {
    width: 30px; height: 28px; border: 1px solid var(--border); border-radius: 4px;
    background: var(--surface); color: var(--text-muted); font-size: 11px; font-weight: 600;
    cursor: pointer; transition: all 0.15s; display: flex; align-items: center; justify-content: center;
  }
  .day-btn.active { background: var(--accent); color: #fff; border-color: var(--accent); }
  .day-btn:hover { border-color: var(--accent); }
  .recurrence-extra { display: flex; gap: 6px; align-items: center; font-size: 12px; color: var(--text-muted); }
  .recurrence-extra input[type="number"] {
    width: 50px; padding: 4px 6px; border: 1px solid var(--border); border-radius: 4px;
    background: var(--surface); color: var(--text); font-size: 12px; text-align: center;
  }
  .recurrence-extra input[type="date"] {
    padding: 4px 6px; border: 1px solid var(--border); border-radius: 4px;
    background: var(--surface); color: var(--text); font-size: 12px;
  }
  .flex-row { display: flex; gap: 6px; align-items: center; }
  .section-label { font-size: 12px; color: var(--text-muted); margin-bottom: 4px; }
  .section-title { font-size: 14px; font-weight: 600; margin-bottom: 12px; }
  .settings-status { margin-top: 12px; font-size: 12px; color: var(--text-muted); }
  .composer-select {
    flex: 1; padding: 6px 8px; border: 1px solid var(--border); border-radius: 6px;
    background: var(--surface); color: var(--text); font-size: 12px; cursor: pointer;
    flex-shrink: 0;
  }
  .hidden { display: none; }

  /* ── Recipient field ── */
  .recipient-field {
    display: flex; gap: 0; margin-bottom: 6px; flex-shrink: 0;
  }
  .recipient-field input {
    flex: 1; padding: 6px 10px; border: 1px solid var(--border); border-radius: 6px 0 0 6px;
    background: var(--surface); color: var(--text); font-size: 12px; outline: none;
  }
  .recipient-field input:focus { border-color: var(--accent); }
  .recipient-field input::placeholder { color: var(--text-dim); }
  .recipient-field select {
    width: auto; padding: 6px 8px; border: 1px solid var(--border); border-left: none;
    border-radius: 0 6px 6px 0; background: var(--surface); color: var(--text);
    font-size: 12px; cursor: pointer;
  }
  .recipient-chips {
    display: flex; flex-wrap: wrap; gap: 4px; margin-bottom: 6px; flex-shrink: 0;
  }
  .recipient-chip {
    display: inline-flex; align-items: center; gap: 4px;
    background: var(--card-bg); padding: 3px 8px; border-radius: 4px; font-size: 11px;
  }
  .recipient-chip .remove {
    cursor: pointer; color: var(--danger); font-weight: bold; font-size: 12px;
  }

  /* ── Recipient autocomplete dropdown ── */
  .recipient-field { position: relative; }
  .autocomplete-dropdown {
    display: none; position: absolute; top: 100%; left: 0; right: 0;
    background: var(--panel); border: 1px solid var(--border); border-top: none;
    border-radius: 0 0 6px 6px; max-height: 220px; overflow-y: auto; z-index: 100;
    box-shadow: 0 4px 12px rgba(0,0,0,0.2);
  }
  .autocomplete-dropdown.show { display: block; }
  .ac-item {
    padding: 6px 10px; font-size: 12px; cursor: pointer; display: flex; align-items: center; gap: 8px;
  }
  .ac-item:hover, .ac-item.highlighted { background: var(--row-hover); }
  .ac-type {
    font-size: 10px; padding: 1px 5px; border-radius: 3px; font-weight: 600;
    flex-shrink: 0; text-transform: uppercase;
  }
  .ac-type.contact { background: var(--accent); color: #fff; }
  .ac-type.family { background: #7a6ab5; color: #fff; }
  .ac-type.group { background: #b58a3a; color: #fff; }
  .ac-name { flex: 1; }
  .ac-detail { color: var(--text-muted); font-size: 11px; }

  /* ── Tab search bars ── */
  .search-row { display: flex; gap: 6px; margin-bottom: 8px; align-items: center; }
  .search-row .tab-search { margin-bottom: 0; flex: 1; }
  .add-btn {
    width: 32px; height: 32px; border: none; border-radius: 6px; cursor: pointer;
    background: var(--accent); color: #fff; font-size: 20px; font-weight: 600;
    display: flex; align-items: center; justify-content: center; transition: background 0.15s;
    flex-shrink: 0; line-height: 1;
  }
  .add-btn:hover { background: var(--accent-hover); }
  .tab-search {
    width: 100%; padding: 6px 10px; margin-bottom: 8px; border: 1px solid var(--border);
    border-radius: 6px; background: var(--surface); color: var(--text); font-size: 12px;
    outline: none; transition: border-color 0.2s;
  }
  .tab-search:focus { border-color: var(--accent); }
  .tab-search::placeholder { color: var(--text-dim); }

  /* Create modal (used for new contact/family/group) */
  .create-overlay {
    display: none; position: fixed; top: 0; left: 0; width: 100%; height: 100%;
    background: rgba(0,0,0,0.35); z-index: 9999;
  }
  .create-overlay.show { display: block; }
  .create-modal {
    position: fixed; z-index: 10000;
    top: 50%; left: 50%; transform: translate(-50%, -50%) scale(0.3);
    background: var(--surface); border: 1px solid var(--accent); border-radius: 10px;
    padding: 18px 20px; min-width: 340px; max-width: 420px;
    box-shadow: 0 8px 32px rgba(0,0,0,0.3);
    display: none; flex-direction: column; gap: 8px;
    animation: createPopIn 0.3s cubic-bezier(0.34, 1.56, 0.64, 1) forwards;
  }
  .create-modal.show { display: flex; }
  @keyframes createPopIn {
    0% { opacity: 0; transform: translate(-50%, -50%) scale(0.3); }
    100% { opacity: 1; transform: translate(-50%, -50%) scale(1); }
  }
  .create-modal h3 { margin: 0 0 4px; font-size: 15px; color: var(--text); font-weight: 700; }
  .create-modal .edit-row { display: flex; gap: 6px; }
  .create-modal input, .create-modal select {
    flex: 1; padding: 7px 10px; font-size: 13px; border: 1px solid var(--border);
    border-radius: 6px; background: var(--bg); color: var(--text);
  }
  .create-modal .edit-actions { display: flex; gap: 6px; justify-content: flex-end; margin-top: 4px; }

  /* Edit modal shared styles (Groups & Families) */
  .edit-member-list {
    max-height: 150px; overflow-y: auto; margin: 4px 0; display: flex; flex-direction: column; gap: 2px;
  }
  .edit-member-pill {
    display: flex; align-items: center; justify-content: space-between;
    background: var(--bg); border: 1px solid var(--border); border-radius: 6px;
    padding: 5px 10px; font-size: 12px; color: var(--text);
  }
  .edit-member-pill .remove-x {
    cursor: pointer; color: var(--danger); font-weight: bold; font-size: 14px; line-height: 1;
    padding: 0 2px;
  }
  .edit-member-pill .remove-x:hover { opacity: 0.7; }
  .edit-search-results {
    max-height: 120px; overflow-y: auto; border: 1px solid var(--border);
    border-radius: 6px; margin-top: 4px; background: var(--bg);
  }
  .edit-search-results .esr-item {
    padding: 5px 8px; font-size: 12px; cursor: pointer; color: var(--text);
    border-bottom: 1px solid var(--border);
  }
  .edit-search-results .esr-item:last-child { border-bottom: none; }
  .edit-search-results .esr-item:hover { background: var(--row-hover); }
  .edit-search-results .esr-email { color: var(--text-muted); font-size: 11px; margin-left: 4px; }
  .edit-section-label { font-size: 12px; color: var(--text-muted); margin: 6px 0 2px; font-weight: 600; }

  /* ── Right panel: Composer ── */
  .composer-title { font-size: 16px; font-weight: 700; margin-bottom: 10px; }
  #subject {
    width: 100%; padding: 5px 8px; border: 1px solid var(--border); border-radius: 5px;
    background: var(--surface); color: var(--text); font-size: 12px; outline: none; margin-bottom: 6px;
    transition: background 0.2s, color 0.2s, border-color 0.2s;
    flex-shrink: 0; height: 30px;
  }
  #subject:focus { border-color: var(--accent); }
  #subject::placeholder { color: var(--text-dim); }

  /* ── Quill editor area ── */
  .editor-container {
    flex: 1; display: flex; flex-direction: column; border: 1px solid var(--border);
    border-radius: 8px; overflow: hidden; background: var(--surface);
    transition: background 0.2s, border-color 0.2s;
    min-height: 0;
  }

  /* Quill overrides moved to separate style block after quill.snow.css (see below) */

  /* ── Theme toggle ── */
  .theme-toggle {
    position: fixed; top: 12px; right: 12px; z-index: 9999;
    padding: 5px 12px; border: 1px solid var(--border); border-radius: 6px;
    background: var(--panel); color: var(--text-muted); font-size: 12px;
    cursor: pointer; transition: all 0.2s;
  }
  .theme-toggle:hover { border-color: var(--accent); color: var(--text); }

  /* ── Toast ── */
  .toast {
    position: fixed; bottom: 20px; right: 20px; padding: 12px 20px; border-radius: 8px;
    font-size: 13px; font-weight: 600; color: #fff; opacity: 0; transition: opacity 0.3s;
    z-index: 9999; max-width: 350px;
  }
  .toast.show { opacity: 1; }
  .toast.success { background: var(--success); }
  .toast.error { background: var(--danger); }

  /* ── Template Save Modal ── */
  .modal-overlay {
    display: none; position: fixed; top: 0; left: 0; width: 100%; height: 100%;
    background: rgba(0,0,0,0.5); z-index: 10000; align-items: center; justify-content: center;
  }
  .modal-overlay.show { display: flex; }
  .modal-box {
    background: var(--surface); border: 1px solid var(--border); border-radius: 12px;
    padding: 24px; min-width: 340px; max-width: 420px; box-shadow: 0 8px 32px rgba(0,0,0,0.3);
  }
  .modal-box h3 { margin: 0 0 12px; font-size: 15px; color: var(--text); }
  .modal-box p { margin: 0 0 16px; font-size: 13px; color: var(--text-secondary); }
  .modal-box input[type="text"] {
    width: 100%; padding: 8px 10px; border: 1px solid var(--border); border-radius: 6px;
    background: var(--bg); color: var(--text); font-size: 13px; margin-bottom: 16px;
    box-sizing: border-box;
  }
  .modal-actions { display: flex; gap: 8px; justify-content: flex-end; }
</style>

<link href="https://cdn.jsdelivr.net/npm/quill@2.0.3/dist/quill.snow.css" rel="stylesheet">
<style>
  /* Quill overrides — MUST load after quill.snow.css to win specificity */
  .editor-container .ql-snow .ql-toolbar {
    background: var(--toolbar-bg); border: none; border-bottom: 1px solid var(--border);
  }
  .editor-container .ql-toolbar.ql-snow {
    background: var(--toolbar-bg); border: none; border-bottom: 1px solid var(--border);
  }
  .editor-container .ql-snow .ql-stroke { stroke: var(--text-muted); }
  .editor-container .ql-snow .ql-fill { fill: var(--text-muted); }
  .editor-container .ql-snow .ql-picker-label { color: var(--text-muted); }
  .editor-container .ql-snow button:hover .ql-stroke,
  .editor-container .ql-snow .ql-picker-label:hover .ql-stroke { stroke: var(--text); }
  .editor-container .ql-snow button:hover .ql-fill,
  .editor-container .ql-snow .ql-picker-label:hover .ql-fill { fill: var(--text); }
  .editor-container .ql-snow button.ql-active .ql-stroke { stroke: var(--accent); }
  .editor-container .ql-snow button.ql-active .ql-fill { fill: var(--accent); }
  .editor-container .ql-snow .ql-picker-options {
    background: var(--panel); border-color: var(--border);
  }
  .editor-container .ql-snow .ql-picker-item { color: var(--text-muted); }
  .editor-container .ql-snow .ql-picker-item:hover { color: var(--text); }

  .editor-container .ql-snow.ql-container { border: none; flex: 1; overflow: hidden; display: flex; flex-direction: column; }
  .editor-container .ql-snow .ql-editor {
    overflow-y: auto; padding: 12px 14px; font-size: 14px; line-height: 1.6;
    flex: 1; color: var(--text); background: var(--surface);
  }
  .editor-container .ql-snow .ql-editor.ql-blank::before {
    color: var(--text-dim); font-style: normal;
  }
  .editor-container .ql-snow .ql-editor::-webkit-scrollbar { width: 6px; }
  .editor-container .ql-snow .ql-editor::-webkit-scrollbar-thumb { background: var(--scrollbar); border-radius: 3px; }
</style>
<script src="https://cdn.jsdelivr.net/npm/quill@2.0.3/dist/quill.js"></script>
</head>
<body>
<div class="app">

  <!-- ════ LEFT PANEL ════ -->
  <div class="left-panel" id="left-panel">
    <div class="tab-bar">
      <button class="tab-btn active" onclick="switchTab('contacts')">Contacts</button>
      <button class="tab-btn" onclick="switchTab('families')">Families</button>
      <button class="tab-btn" onclick="switchTab('groups')">Groups</button>
      <button class="tab-btn" onclick="switchTab('scheduled')">Scheduled</button>
      <button class="tab-btn" onclick="switchTab('history')">History</button>
      <button class="tab-btn" onclick="switchTab('settings')">Settings</button>
    </div>

    <!-- Contacts Tab -->
    <div id="contacts-tab" class="tab-content active">
      <div class="search-row">
        <input type="text" class="tab-search" id="search-contacts" placeholder="Search contacts..." oninput="filterContacts()">
        <button class="add-btn" onclick="openCreateContact()" title="Add Contact">+</button>
      </div>
      <div class="list-area" id="contact-list"></div>
      <div class="btn-row">
        <button class="btn btn-danger" onclick="deleteSelected()">Delete Selected</button>
        <button class="btn btn-primary btn-success" onclick="importCSV()">Import CSV</button>
        <button class="btn btn-primary" onclick="exportCSV()">Export CSV</button>
      </div>
    </div>

    <!-- Families Tab -->
    <div id="families-tab" class="tab-content">
      <div class="search-row">
        <input type="text" class="tab-search" id="search-families" placeholder="Search families..." oninput="filterFamilies()">
        <button class="add-btn" onclick="openCreateFamily()" title="Add Family">+</button>
      </div>
      <div class="list-area" id="family-list"></div>
    </div>

    <!-- Groups Tab -->
    <div id="groups-tab" class="tab-content">
      <div class="groups-split">
        <div class="groups-left">
          <div class="search-row">
            <input type="text" class="tab-search" id="search-groups" placeholder="Search groups or members..." oninput="filterGroups()">
            <button class="add-btn" onclick="openCreateGroup()" title="Add Group">+</button>
          </div>
          <div class="list-area" id="group-list"></div>
        </div>
        <div class="groups-right">
          <div class="groups-right-header" id="group-detail-header">
            <span class="group-detail-title" id="group-detail-title">Select a group</span>
            <span id="group-detail-actions"></span>
          </div>
          <input type="text" class="tab-search" id="search-group-members" placeholder="Search members..." oninput="filterGroupMembers()" style="margin-bottom:6px;">
          <div class="list-area" id="group-member-list">
            <div class="empty-msg">Click a group to view its members</div>
          </div>
        </div>
      </div>
    </div>

    <!-- Scheduled Tab -->
    <div id="scheduled-tab" class="tab-content">
      <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:8px;">
        <button class="btn btn-sm" onclick="schedChangeMonth(-1)">&#9664; Prev</button>
        <span id="sched-month-label" style="font-weight:700;font-size:15px;color:var(--text);"></span>
        <button class="btn btn-sm" onclick="schedChangeMonth(1)">Next &#9654;</button>
      </div>
      <div class="list-area" id="scheduled-calendar" style="flex:1;overflow-y:auto;"></div>
    </div>

    <!-- History Tab -->
    <div id="history-tab" class="tab-content">
      <div class="list-area" id="history-list"></div>
    </div>

    <!-- Settings Tab -->
    <div id="settings-tab" class="tab-content">
      <div style="padding: 10px 0;">
        <div class="section-title">Gmail SMTP Credentials</div>
        <div class="form-row">
          <input type="text" id="s-email" placeholder="Gmail address">
        </div>
        <div class="form-row">
          <input type="text" id="s-password" placeholder="App Password (leave blank to keep current)">
        </div>
        <div class="btn-row" style="margin-top: 6px;">
          <button class="btn btn-primary" onclick="saveSettings()">Save Settings</button>
          <button class="btn btn-primary btn-success" onclick="testConnection()">Test Connection</button>
        </div>
        <div id="settings-status" class="settings-status"></div>
        <div class="section-title" style="margin-top:20px;">Time Zone</div>
        <div class="form-row">
          <select id="s-timezone" style="flex:1;">
            <option value="US/Eastern">Eastern (US/Eastern)</option>
            <option value="US/Central">Central (US/Central)</option>
            <option value="US/Mountain">Mountain (US/Mountain)</option>
            <option value="US/Pacific">Pacific (US/Pacific)</option>
            <option value="US/Alaska">Alaska (US/Alaska)</option>
            <option value="US/Hawaii">Hawaii (US/Hawaii)</option>
            <option value="UTC">UTC</option>
            <option value="Europe/London">London (Europe/London)</option>
            <option value="Europe/Paris">Paris (Europe/Paris)</option>
            <option value="Europe/Berlin">Berlin (Europe/Berlin)</option>
            <option value="Asia/Tokyo">Tokyo (Asia/Tokyo)</option>
            <option value="Asia/Shanghai">Shanghai (Asia/Shanghai)</option>
            <option value="Asia/Kolkata">India (Asia/Kolkata)</option>
            <option value="Asia/Dubai">Dubai (Asia/Dubai)</option>
            <option value="Australia/Sydney">Sydney (Australia/Sydney)</option>
            <option value="Pacific/Auckland">Auckland (Pacific/Auckland)</option>
            <option value="America/Sao_Paulo">Sao Paulo (America/Sao_Paulo)</option>
            <option value="America/Mexico_City">Mexico City (America/Mexico_City)</option>
            <option value="America/Chicago">Chicago (America/Chicago)</option>
            <option value="America/Denver">Denver (America/Denver)</option>
            <option value="America/Los_Angeles">Los Angeles (America/Los_Angeles)</option>
            <option value="America/New_York">New York (America/New_York)</option>
            <option value="America/Anchorage">Anchorage (America/Anchorage)</option>
            <option value="America/Phoenix">Phoenix (America/Phoenix)</option>
          </select>
          <button class="btn btn-primary btn-nowrap" onclick="saveTimezone()">Save Timezone</button>
        </div>
        <div id="tz-status" class="settings-status"></div>
      </div>
    </div>
  </div>

  <!-- ════ DIVIDER ════ -->
  <div class="panel-divider" id="panel-divider"></div>

  <!-- ════ RIGHT PANEL ════ -->
  <div class="right-panel" id="right-panel">
    <div class="composer-title">Email Composer</div>
    <div class="flex-row" style="margin-bottom:8px;">
      <select id="template-select" class="composer-select" onchange="loadTemplate()">
        <option value="">Load template...</option>
      </select>
      <button class="btn btn-primary btn-sm" id="btn-save-template" onclick="saveAsTemplate()">Save Template</button>
      <button class="btn btn-danger btn-sm" onclick="deleteCurrentTemplate()">Del</button>
    </div>
    <div class="recipient-field">
      <input type="text" id="recipient-input" placeholder="Search name, email, group, family..." autocomplete="off">
      <select id="target-select" onchange="onTargetSelect()">
        <option value="">-- Quick select --</option>
        <option value="all">All Contacts</option>
        <option value="family">All Families</option>
        <option value="single">All Singles</option>
      </select>
      <div class="autocomplete-dropdown" id="ac-dropdown"></div>
    </div>
    <div class="recipient-chips" id="recipient-chips"></div>
    <input type="text" id="subject" placeholder="Subject">
    <div class="editor-container">
      <div id="editor"></div>
    </div>
    <div class="flex-row" style="margin-top:8px;">
      <button class="btn btn-primary" onclick="attachFiles()">Attach Files</button>
      <div id="attach-chips" style="display:flex; flex-wrap:wrap; gap:4px; flex:1;"></div>
    </div>
    <div style="margin-top:8px; flex-shrink:0;">
      <div class="flex-row" style="margin-bottom:6px;">
        <label style="font-size:12px; color:var(--text-muted); white-space:nowrap;">Start:</label>
        <input type="datetime-local" id="sched-datetime" class="composer-select" style="flex:1;">
      </div>
      <div class="recurrence-row" style="margin-bottom:6px;">
        <select id="recurrence-type" class="composer-select" onchange="onRecurrenceChange()" style="flex:0 0 auto; width:auto;">
          <option value="once">One-time</option>
          <option value="daily">Daily</option>
          <option value="every_other_day">Every other day</option>
          <option value="weekly">Weekly</option>
          <option value="every_other_week">Every other week</option>
          <option value="monthly">Monthly</option>
        </select>
        <div class="day-picker hidden" id="day-picker">
          <button type="button" class="day-btn" data-day="0">Su</button>
          <button type="button" class="day-btn" data-day="1">Mo</button>
          <button type="button" class="day-btn" data-day="2">Tu</button>
          <button type="button" class="day-btn" data-day="3">We</button>
          <button type="button" class="day-btn" data-day="4">Th</button>
          <button type="button" class="day-btn" data-day="5">Fr</button>
          <button type="button" class="day-btn" data-day="6">Sa</button>
        </div>
        <div class="recurrence-extra hidden" id="monthly-day">
          <span>Day</span>
          <input type="number" id="month-day" min="1" max="31" value="1">
        </div>
      </div>
      <div class="flex-row" style="margin-bottom:6px;">
        <div class="recurrence-extra" id="end-date-row">
          <label style="white-space:nowrap;">End date (optional):</label>
          <input type="datetime-local" id="recurrence-end-date" class="composer-select">
        </div>
        <div style="flex:1;"></div>
        <button class="btn btn-primary btn-nowrap" id="btn-schedule" onclick="scheduleEmail()">Save &amp; Schedule</button>
      </div>
    </div>
    <button class="btn-dispatch" id="btn-send-now" onclick="dispatchEmails()">Send Now</button>
  </div>
</div>

<button class="theme-toggle" id="theme-toggle" onclick="toggleTheme()">Light Mode</button>
<div class="toast" id="toast"></div>
<div class="create-overlay" id="create-overlay" onclick="closeCreateModal()"></div>
<div class="create-modal" id="create-modal"></div>

<div class="modal-overlay" id="template-save-modal">
  <div class="modal-box">
    <h3 id="template-modal-title">Save Template</h3>
    <p id="template-modal-desc"></p>
    <input type="text" id="template-modal-name" placeholder="Template name" class="hidden">
    <div class="modal-actions">
      <button class="btn btn-sm" id="template-modal-cancel" onclick="closeTemplateModal()">Cancel</button>
      <button class="btn btn-primary btn-sm hidden" id="template-modal-override">Override</button>
      <button class="btn btn-primary btn-sm hidden" id="template-modal-new">Save as New</button>
      <button class="btn btn-primary btn-sm hidden" id="template-modal-save">Save</button>
    </div>
  </div>
</div>

<script>
// ── Global error handler (surfaces JS errors in toast) ──
window.onerror = function(msg, src, line) {
  document.title = 'JS ERR: ' + msg + ' (line ' + line + ')';
};

// ── Quill editor init ──
var quill;
try {
  quill = new Quill('#editor', {
    theme: 'snow',
    placeholder: 'Compose your email...',
    modules: {
      toolbar: [
        ['bold', 'italic', 'underline'],
        [{ 'header': [1, 2, 3, false] }],
        [{ 'list': 'ordered' }, { 'list': 'bullet' }],
        [{ 'color': [] }, { 'background': [] }],
        ['link', 'image', 'clean']
      ]
    }
  });
} catch(e) {
  document.title = 'Quill init failed: ' + e.message;
}

// ── Theme toggle ──
window.toggleTheme = function() {
  document.body.classList.toggle('light');
  var isLight = document.body.classList.contains('light');
  document.getElementById('theme-toggle').textContent = isLight ? 'Dark Mode' : 'Light Mode';
  pywebview.api.set_ui_setting('theme', isLight ? 'light' : 'dark');
};
// Theme restored in initApp after API is ready

// ── Extract editor content ──
function getEditorHTML() {
  if (quill) return quill.root.innerHTML;
  return document.getElementById('editor').innerHTML;
}

function getEditorPlainText() {
  if (quill) return quill.getText();
  return document.getElementById('editor').textContent;
}

// ── Wait for pywebview bridge ──
var _apiReady = new Promise(function(resolve) {
  if (window.pywebview && window.pywebview.api) {
    resolve();
  } else {
    window.addEventListener('pywebviewready', resolve, { once: true });
  }
});
async function waitForApi() {
  await _apiReady;
}

// ── Toast notifications ──
function showToast(msg, type) {
  const el = document.getElementById('toast');
  el.textContent = msg;
  el.className = 'toast show ' + type;
  setTimeout(() => el.className = 'toast', 3000);
}

// ══════════════════════════════════════════════════════════════════════════════
// CONTACTS
// ══════════════════════════════════════════════════════════════════════════════

var cachedContacts = [];
async function loadContacts() {
  await waitForApi();
  cachedContacts = await pywebview.api.get_contacts();
  renderContactList(cachedContacts);
}

function renderContactList(contacts) {
  const el = document.getElementById('contact-list');
  if (!contacts.length) { renderEmpty(el, 'No contacts yet.'); return; }
  const header = `<div class="contact-header">
    <span class="h-check"></span>
    <span class="h-name">Name</span>
    <span class="h-email">Email</span>
    <span class="h-families">Families</span>
    <span class="h-groups">Groups</span>
    <span class="h-edit"></span>
  </div>`;
  const rows = contacts.map(c => {
    const famTags = (c.families && c.families.length)
      ? c.families.map(f => `<span class="contact-tag fam-tag">${esc(f.name)}</span>`).join('')
      : '<span style="font-size:11px;color:var(--text-dim);">-</span>';
    const grpTags = (c.groups && c.groups.length)
      ? c.groups.map(g => `<span class="contact-tag grp-tag">${esc(g)}</span>`).join('')
      : '<span style="font-size:11px;color:var(--text-dim);">-</span>';
    return `
    <div class="contact-row" id="contact-row-${c.id}">
      <input type="checkbox" data-id="${c.id}">
      <span class="name">${esc(c.name)}</span>
      <span class="email">${esc(c.email)}</span>
      <span class="col-families">${famTags}</span>
      <span class="col-groups">${grpTags}</span>
      <button class="contact-edit-btn" onclick="editContact(${c.id})" title="Edit">&#9998;</button>
    </div>`;
  }).join('');
  el.innerHTML = header + rows;
}

window.filterContacts = function() {
  const q = document.getElementById('search-contacts').value.toLowerCase().trim();
  if (!q) { renderContactList(cachedContacts); return; }
  const filtered = cachedContacts.filter(c =>
    c.name.toLowerCase().includes(q) || c.email.toLowerCase().includes(q) ||
    (c.families && c.families.some(f => f.name.toLowerCase().includes(q))) ||
    (c.groups && c.groups.some(g => g.toLowerCase().includes(q)))
  );
  renderContactList(filtered);
};

// ── Create modal helpers ──

function showCreateModal(html) {
  var modal = document.getElementById('create-modal');
  modal.innerHTML = html;
  // Remove old animation class, re-trigger
  modal.classList.remove('show');
  void modal.offsetWidth;
  modal.classList.add('show');
  requestAnimationFrame(function() {
    document.getElementById('create-overlay').classList.add('show');
  });
}

window.closeCreateModal = function() {
  document.getElementById('create-modal').classList.remove('show');
  document.getElementById('create-overlay').classList.remove('show');
};

window.openCreateContact = function() {
  var famOptions = '<option value="">No family</option>';
  cachedFamilies.forEach(function(f) {
    famOptions += '<option value="' + f.id + '">' + esc(f.name) + '</option>';
  });
  showCreateModal(
    '<h3>New Contact</h3>' +
    '<div class="edit-row"><input type="text" id="cm-name" placeholder="Name"><input type="text" id="cm-email" placeholder="Email"></div>' +
    '<div class="edit-row">' +
      '<select id="cm-category" onchange="document.getElementById(\'cm-family\').disabled = this.value !== \'Family\'"><option value="Single">Single</option><option value="Family">Family</option></select>' +
      '<select id="cm-family" disabled>' + famOptions + '</select>' +
    '</div>' +
    '<div class="edit-actions"><button class="btn btn-sm" onclick="closeCreateModal()">Cancel</button><button class="btn btn-primary btn-sm" onclick="submitCreateContact()">Add</button></div>'
  );
  document.getElementById('cm-name').focus();
};

window.submitCreateContact = async function() {
  var name = document.getElementById('cm-name').value.trim();
  var email = document.getElementById('cm-email').value.trim();
  var category = document.getElementById('cm-category').value;
  var familyId = document.getElementById('cm-family').value || null;
  if (!name || !email) { showToast('Enter both name and email.', 'error'); return; }
  var res = await pywebview.api.add_contact(name, email, category, familyId);
  if (!res.ok) { showToast(res.error, 'error'); return; }
  closeCreateModal();
  loadContacts();
  loadFamilies();
  refreshAcCache();
};

window.deleteSelected = async function() {
  const checks = document.querySelectorAll('#contact-list input[type="checkbox"]:checked');
  const ids = Array.from(checks).map(cb => parseInt(cb.dataset.id));
  if (!ids.length) { showToast('Select contacts to delete.', 'error'); return; }
  if (!confirm('Delete ' + ids.length + ' contact(s)?')) return;

  await pywebview.api.delete_contacts(ids);
  loadContacts();
  loadFamilies();
  refreshAcCache();
};

// ── Contact edit modal (with family + group management) ──
var _ceContactId = null;
var _ceOriginalFamilies = [];
var _ceOriginalGroups = [];
var _cePendingFamAdds = [];
var _cePendingFamRemoves = [];
var _cePendingGrpAdds = [];
var _cePendingGrpRemoves = [];

window.editContact = function(id) {
  var c = cachedContacts.find(x => x.id === id);
  if (!c) return;
  _ceContactId = id;
  _ceOriginalFamilies = (c.families || []).slice();
  _ceOriginalGroups = cachedGroups.filter(g => g.members.some(m => m.id === id)).map(g => ({id: g.id, name: g.name}));
  _cePendingFamAdds = [];
  _cePendingFamRemoves = [];
  _cePendingGrpAdds = [];
  _cePendingGrpRemoves = [];

  showCreateModal(
    '<h3>Edit Contact</h3>' +
    '<div class="edit-row">' +
      '<input type="text" id="ce-name" value="' + esc(c.name).replace(/"/g, '&quot;') + '" placeholder="Name">' +
      '<input type="text" id="ce-email" value="' + esc(c.email).replace(/"/g, '&quot;') + '" placeholder="Email">' +
    '</div>' +
    '<div class="edit-section-label">Families</div>' +
    '<div class="edit-member-list" id="ce-fam-list"></div>' +
    '<input type="text" id="ce-fam-search" placeholder="Search families to add..." oninput="ceFamSearch()" style="padding:6px 10px;font-size:12px;border:1px solid var(--border);border-radius:6px;background:var(--bg);color:var(--text);width:100%;box-sizing:border-box;">' +
    '<div class="edit-search-results" id="ce-fam-results"></div>' +
    '<div class="edit-section-label">Groups</div>' +
    '<div class="edit-member-list" id="ce-grp-list"></div>' +
    '<input type="text" id="ce-grp-search" placeholder="Search groups to add..." oninput="ceGrpSearch()" style="padding:6px 10px;font-size:12px;border:1px solid var(--border);border-radius:6px;background:var(--bg);color:var(--text);width:100%;box-sizing:border-box;">' +
    '<div class="edit-search-results" id="ce-grp-results"></div>' +
    '<div class="edit-actions">' +
      '<button class="btn btn-sm" onclick="closeCreateModal()">Cancel</button>' +
      '<button class="btn btn-primary btn-sm" onclick="saveEditContact()">Save</button>' +
    '</div>'
  );
  ceRenderFamilies();
  ceRenderGroups();
  document.getElementById('ce-name').focus();
};

window.ceRenderFamilies = function() {
  var el = document.getElementById('ce-fam-list');
  var current = _ceOriginalFamilies.filter(f => _cePendingFamRemoves.indexOf(f.id) === -1);
  var all = current.concat(_cePendingFamAdds);
  if (!all.length) { el.innerHTML = '<span style="font-size:11px;color:var(--text-muted);">None</span>'; return; }
  el.innerHTML = all.map(function(f) {
    var isNew = _cePendingFamAdds.some(a => a.id === f.id);
    return '<span class="edit-member-pill">' + esc(f.name) +
      ' <span class="remove-x" onclick="ceRemoveFam(' + f.id + ',' + (isNew ? 'true' : 'false') + ')">&times;</span></span>';
  }).join('');
};

window.ceRemoveFam = function(famId, isNew) {
  if (isNew) {
    _cePendingFamAdds = _cePendingFamAdds.filter(a => a.id !== famId);
  } else {
    if (_cePendingFamRemoves.indexOf(famId) === -1) _cePendingFamRemoves.push(famId);
  }
  ceRenderFamilies();
};

window.ceFamSearch = function() {
  var el = document.getElementById('ce-fam-results');
  var q = document.getElementById('ce-fam-search').value.toLowerCase().trim();
  if (!q) { el.innerHTML = ''; return; }
  var currentIds = _ceOriginalFamilies.filter(f => _cePendingFamRemoves.indexOf(f.id) === -1).map(f => f.id);
  var addedIds = _cePendingFamAdds.map(a => a.id);
  var exclude = currentIds.concat(addedIds);
  var matches = cachedFamilies.filter(function(f) {
    if (exclude.indexOf(f.id) !== -1) return false;
    return f.name.toLowerCase().includes(q);
  }).slice(0, 8);
  if (!matches.length) { el.innerHTML = '<div class="esr-item" style="color:var(--text-muted);cursor:default;">No matches</div>'; return; }
  el.innerHTML = matches.map(function(f) {
    return '<div class="esr-item" onclick="ceAddFam(' + f.id + ')">' + esc(f.name) + '</div>';
  }).join('');
};

window.ceAddFam = function(famId) {
  var f = cachedFamilies.find(x => x.id === famId);
  if (!f) return;
  var wasOriginal = _ceOriginalFamilies.some(o => o.id === famId);
  if (wasOriginal) {
    _cePendingFamRemoves = _cePendingFamRemoves.filter(id => id !== famId);
  } else {
    if (!_cePendingFamAdds.some(a => a.id === famId)) {
      _cePendingFamAdds.push({id: f.id, name: f.name});
    }
  }
  ceRenderFamilies();
  document.getElementById('ce-fam-search').value = '';
  document.getElementById('ce-fam-results').innerHTML = '';
};

window.ceRenderGroups = function() {
  var el = document.getElementById('ce-grp-list');
  var current = _ceOriginalGroups.filter(g => _cePendingGrpRemoves.indexOf(g.id) === -1);
  var all = current.concat(_cePendingGrpAdds);
  if (!all.length) { el.innerHTML = '<span style="font-size:11px;color:var(--text-muted);">None</span>'; return; }
  el.innerHTML = all.map(function(g) {
    var isNew = _cePendingGrpAdds.some(a => a.id === g.id);
    return '<span class="edit-member-pill">' + esc(g.name) +
      ' <span class="remove-x" onclick="ceRemoveGrp(' + g.id + ',' + (isNew ? 'true' : 'false') + ')">&times;</span></span>';
  }).join('');
};

window.ceRemoveGrp = function(grpId, isNew) {
  if (isNew) {
    _cePendingGrpAdds = _cePendingGrpAdds.filter(a => a.id !== grpId);
  } else {
    if (_cePendingGrpRemoves.indexOf(grpId) === -1) _cePendingGrpRemoves.push(grpId);
  }
  ceRenderGroups();
};

window.ceGrpSearch = function() {
  var el = document.getElementById('ce-grp-results');
  var q = document.getElementById('ce-grp-search').value.toLowerCase().trim();
  if (!q) { el.innerHTML = ''; return; }
  var currentIds = _ceOriginalGroups.filter(g => _cePendingGrpRemoves.indexOf(g.id) === -1).map(g => g.id);
  var addedIds = _cePendingGrpAdds.map(a => a.id);
  var exclude = currentIds.concat(addedIds);
  var matches = cachedGroups.filter(function(g) {
    if (exclude.indexOf(g.id) !== -1) return false;
    return g.name.toLowerCase().includes(q);
  }).slice(0, 8);
  if (!matches.length) { el.innerHTML = '<div class="esr-item" style="color:var(--text-muted);cursor:default;">No matches</div>'; return; }
  el.innerHTML = matches.map(function(g) {
    return '<div class="esr-item" onclick="ceAddGrp(' + g.id + ')">' + esc(g.name) + '</div>';
  }).join('');
};

window.ceAddGrp = function(grpId) {
  var g = cachedGroups.find(x => x.id === grpId);
  if (!g) return;
  var wasOriginal = _ceOriginalGroups.some(o => o.id === grpId);
  if (wasOriginal) {
    _cePendingGrpRemoves = _cePendingGrpRemoves.filter(id => id !== grpId);
  } else {
    if (!_cePendingGrpAdds.some(a => a.id === grpId)) {
      _cePendingGrpAdds.push({id: g.id, name: g.name});
    }
  }
  ceRenderGroups();
  document.getElementById('ce-grp-search').value = '';
  document.getElementById('ce-grp-results').innerHTML = '';
};

window.saveEditContact = async function() {
  var name = document.getElementById('ce-name').value.trim();
  var email = document.getElementById('ce-email').value.trim();
  if (!name || !email) { showToast('Enter both name and email.', 'error'); return; }

  // Update contact name/email (keep existing category/family_id for backwards compat)
  var c = cachedContacts.find(x => x.id === _ceContactId);
  var res = await pywebview.api.update_contact(_ceContactId, name, email, c ? c.category : 'Single', null);
  if (!res.ok) { showToast(res.error, 'error'); return; }

  // Family membership changes
  for (var i = 0; i < _cePendingFamRemoves.length; i++) {
    await pywebview.api.remove_family_member(_cePendingFamRemoves[i], _ceContactId);
  }
  for (var j = 0; j < _cePendingFamAdds.length; j++) {
    await pywebview.api.add_family_member(_cePendingFamAdds[j].id, _ceContactId);
  }

  // Group membership changes
  for (var k = 0; k < _cePendingGrpRemoves.length; k++) {
    await pywebview.api.remove_group_member(_cePendingGrpRemoves[k], _ceContactId);
  }
  for (var l = 0; l < _cePendingGrpAdds.length; l++) {
    await pywebview.api.add_group_member(_cePendingGrpAdds[l].id, _ceContactId);
  }

  closeCreateModal();
  loadContacts();
  loadFamilies();
  loadGroups();
  refreshAcCache();
};

// ══════════════════════════════════════════════════════════════════════════════
// FAMILIES
// ══════════════════════════════════════════════════════════════════════════════

var cachedFamilies = [];
async function loadFamilies() {
  await waitForApi();
  cachedFamilies = await pywebview.api.get_families();
  renderFamilyList(cachedFamilies);
}

function renderFamilyList(families) {
  const el = document.getElementById('family-list');
  if (!families.length) { renderEmpty(el, 'No families yet.'); return; }
  el.innerHTML = families.map(f => {
    const members = f.members.length ? f.members.map(m => '<div class="member-row">' + esc(m.name) + '</div>').join('') : '<div class="member-row" style="color:var(--text-dim);">No members</div>';
    return `
      <div class="family-card">
        <div class="fam-header">
          <span class="fam-name" id="fam-name-${f.id}">${esc(f.name)}</span>
          <span>
            <button class="contact-edit-btn" onclick="editFamily(${f.id})" title="Edit">&#9998;</button>
            <button class="btn btn-danger btn-sm" onclick="deleteFamily(${f.id})">Delete</button>
          </span>
        </div>
        <div class="fam-members">${members}</div>
      </div>
    `;
  }).join('');
}

window.filterFamilies = function() {
  const q = document.getElementById('search-families').value.toLowerCase().trim();
  if (!q) { renderFamilyList(cachedFamilies); return; }
  const filtered = cachedFamilies.filter(f =>
    f.name.toLowerCase().includes(q) ||
    f.members.some(m => m.name.toLowerCase().includes(q))
  );
  renderFamilyList(filtered);
};

window.openCreateFamily = function() {
  showCreateModal(
    '<h3>New Family</h3>' +
    '<input type="text" id="cm-fname" placeholder="Family Name">' +
    '<div class="edit-actions"><button class="btn btn-sm" onclick="closeCreateModal()">Cancel</button><button class="btn btn-primary btn-sm" onclick="submitCreateFamily()">Add</button></div>'
  );
  document.getElementById('cm-fname').focus();
};

window.submitCreateFamily = async function() {
  var name = document.getElementById('cm-fname').value.trim();
  if (!name) { showToast('Enter a family name.', 'error'); return; }
  var res = await pywebview.api.add_family(name);
  if (!res.ok) { showToast(res.error || 'Family already exists.', 'error'); return; }
  closeCreateModal();
  loadFamilies();
  refreshAcCache();
};

// ── Family edit modal ──
var _efFamilyId = null;
var _efOriginalMembers = [];
var _efPendingRemoves = [];
var _efPendingAdds = [];

window.editFamily = function(id) {
  var family = cachedFamilies.find(f => f.id === id);
  if (!family) return;
  _efFamilyId = id;
  _efOriginalMembers = family.members.slice();
  _efPendingRemoves = [];
  _efPendingAdds = [];
  showCreateModal(
    '<h3>Edit Family</h3>' +
    '<input type="text" id="ef-name" value="' + esc(family.name).replace(/"/g, '&quot;') + '" placeholder="Family Name">' +
    '<div class="edit-section-label">Members</div>' +
    '<div class="edit-member-list" id="ef-members"></div>' +
    '<div class="edit-section-label">Add Members</div>' +
    '<input type="text" id="ef-search" placeholder="Search contacts..." oninput="efSearchContacts()" style="padding:7px 10px;font-size:13px;border:1px solid var(--border);border-radius:6px;background:var(--bg);color:var(--text);width:100%;box-sizing:border-box;">' +
    '<div class="edit-search-results" id="ef-search-results"></div>' +
    '<div class="edit-actions"><button class="btn btn-sm" onclick="closeCreateModal()">Cancel</button><button class="btn btn-primary btn-sm" onclick="saveEditFamily()">Save</button></div>'
  );
  efRenderMembers();
  document.getElementById('ef-search-results').innerHTML = '';
};

window.efRenderMembers = function() {
  var el = document.getElementById('ef-members');
  var current = _efOriginalMembers.filter(m => _efPendingRemoves.indexOf(m.id) === -1);
  var all = current.concat(_efPendingAdds);
  if (!all.length) { el.innerHTML = '<span style="font-size:11px;color:var(--text-muted);">No members</span>'; return; }
  el.innerHTML = all.map(function(m) {
    var isNew = _efPendingAdds.some(a => a.id === m.id);
    return '<span class="edit-member-pill">' + esc(m.name) +
      ' <span class="remove-x" onclick="efRemoveMember(' + m.id + ',' + (isNew ? 'true' : 'false') + ')">&times;</span></span>';
  }).join('');
};

window.efRemoveMember = function(memberId, isNew) {
  if (isNew) {
    _efPendingAdds = _efPendingAdds.filter(a => a.id !== memberId);
  } else {
    if (_efPendingRemoves.indexOf(memberId) === -1) _efPendingRemoves.push(memberId);
  }
  efRenderMembers();
  efSearchContacts();
};

window.efSearchContacts = function() {
  var el = document.getElementById('ef-search-results');
  var q = document.getElementById('ef-search').value.toLowerCase().trim();
  if (!q) { el.innerHTML = ''; return; }
  var currentIds = _efOriginalMembers.filter(m => _efPendingRemoves.indexOf(m.id) === -1).map(m => m.id);
  var addedIds = _efPendingAdds.map(a => a.id);
  var exclude = currentIds.concat(addedIds);
  var matches = cachedContacts.filter(function(c) {
    if (exclude.indexOf(c.id) !== -1) return false;
    return c.name.toLowerCase().includes(q) || c.email.toLowerCase().includes(q);
  }).slice(0, 10);
  if (!matches.length) { el.innerHTML = '<div class="esr-item" style="color:var(--text-muted);cursor:default;">No matches</div>'; return; }
  el.innerHTML = matches.map(function(c) {
    return '<div class="esr-item" onclick="efAddContact(' + c.id + ')">' + esc(c.name) + '<span class="esr-email">' + esc(c.email) + '</span></div>';
  }).join('');
};

window.efAddContact = function(contactId) {
  var c = cachedContacts.find(x => x.id === contactId);
  if (!c) return;
  var wasOriginal = _efOriginalMembers.some(m => m.id === contactId);
  if (wasOriginal) {
    _efPendingRemoves = _efPendingRemoves.filter(id => id !== contactId);
  } else {
    if (!_efPendingAdds.some(a => a.id === contactId)) {
      _efPendingAdds.push({id: c.id, name: c.name, email: c.email});
    }
  }
  efRenderMembers();
  document.getElementById('ef-search').value = '';
  document.getElementById('ef-search-results').innerHTML = '';
};

window.saveEditFamily = async function() {
  var newName = document.getElementById('ef-name').value.trim();
  if (!newName) { showToast('Family name cannot be empty.', 'error'); return; }
  var family = cachedFamilies.find(f => f.id === _efFamilyId);
  if (family && newName !== family.name) {
    var res = await pywebview.api.rename_family(_efFamilyId, newName);
    if (!res.ok) { showToast(res.error || 'Rename failed.', 'error'); return; }
  }
  // Remove members from family
  for (var i = 0; i < _efPendingRemoves.length; i++) {
    await pywebview.api.remove_family_member(_efFamilyId, _efPendingRemoves[i]);
  }
  // Add members to family
  for (var j = 0; j < _efPendingAdds.length; j++) {
    await pywebview.api.add_family_member(_efFamilyId, _efPendingAdds[j].id);
  }
  closeCreateModal();
  loadFamilies();
  loadContacts();
  refreshAcCache();
};

window.deleteFamily = async function(id) {
  if (!confirm('Delete this family? Members will keep their contacts but lose the family link.')) return;
  await pywebview.api.delete_family(id);
  loadFamilies();
  loadContacts();
  refreshAcCache();
};

// ══════════════════════════════════════════════════════════════════════════════
// TABS
// ══════════════════════════════════════════════════════════════════════════════

const TABS = ['contacts', 'families', 'groups', 'scheduled', 'history', 'settings'];
window.switchTab = function(tab) {
  document.querySelectorAll('.tab-btn').forEach((b, i) => {
    b.classList.toggle('active', TABS[i] === tab);
  });
  TABS.forEach(t => {
    document.getElementById(t + '-tab').classList.toggle('active', t === tab);
  });
  // Refresh data when switching to these tabs
  if (tab === 'scheduled') loadScheduled();
  if (tab === 'history') loadHistory();
};


// ══════════════════════════════════════════════════════════════════════════════
// EMAIL DISPATCH
// ══════════════════════════════════════════════════════════════════════════════

// ═════════���═════════════════════════════════��══════════════════════════════════
// ATTACHMENTS
// ═════════���═════════════════════���══════════════════════════════════════════════

var attachedFiles = [];

function renderAttachChips() {
  const el = document.getElementById('attach-chips');
  el.innerHTML = attachedFiles.map((f, i) => {
    const size = f.size < 1024 ? f.size + ' B' :
      f.size < 1048576 ? (f.size / 1024).toFixed(1) + ' KB' :
      (f.size / 1048576).toFixed(1) + ' MB';
    return '<span style="background:var(--card-bg);padding:3px 8px;border-radius:4px;font-size:11px;display:inline-flex;align-items:center;gap:4px;">' +
      esc(f.name) + ' <span style="color:var(--text-muted);">(' + size + ')</span>' +
      '<span style="cursor:pointer;color:var(--danger);font-weight:bold;" onclick="removeAttachment(' + i + ')">&times;</span></span>';
  }).join('');
}

window.attachFiles = async function() {
  await waitForApi();
  const files = await pywebview.api.pick_file();
  if (files && files.length) {
    attachedFiles = attachedFiles.concat(files);
    renderAttachChips();
  }
};

window.removeAttachment = function(idx) {
  attachedFiles.splice(idx, 1);
  renderAttachChips();
};

// ═══════���══════════════════════════════════════════════════════════════════════
// EMAIL DISPATCH
// ══════════════════���════════════════════════════════��══════════════════════════

window.dispatchEmails = async function() {
  const subject = document.getElementById('subject').value.trim();
  const htmlBody = getEditorHTML();
  const plainText = getEditorPlainText();

  if (!subject || !plainText.trim()) {
    showToast('Fill in both subject and body.', 'error');
    return;
  }

  // Check for checkbox overrides on contacts tab
  const checks = document.querySelectorAll('#contact-list input[type="checkbox"]:checked');
  let contactIds = Array.from(checks).map(cb => parseInt(cb.dataset.id));

  // Get recipients from the recipient field
  const { targetType, targetId, manualEmails } = getRecipientSelection();

  if (contactIds.length) {
    if (!confirm('Send to ' + contactIds.length + ' selected contact(s)?')) return;
  } else if (!recipientList.length) {
    showToast('Add at least one recipient.', 'error');
    return;
  } else {
    const count = recipientList.map(r => r.label).join(', ');
    if (!confirm('Send to: ' + count + '?')) return;
  }

  const paths = attachedFiles.map(f => f.path);
  const sendBtn = document.querySelector('.btn-dispatch');
  sendBtn.disabled = true;
  showToast('Sending emails...', 'success');
  try {
    const res = await pywebview.api.dispatch_emails(
      subject, htmlBody, plainText, contactIds, paths,
      targetType || 'manual', targetId, manualEmails
    );
    if (res.error) {
      showToast('Error: ' + res.error, 'error');
    } else {
      let msg = 'Sent: ' + res.sent;
      if (res.failed) msg += ' | Failed: ' + res.failed;
      showToast(msg, 'success');
    }
  } finally {
    sendBtn.disabled = false;
  }
};

// ── Helpers ──
function esc(s) {
  const d = document.createElement('div');
  d.textContent = s;
  return d.innerHTML;
}

function renderEmpty(el, msg) {
  el.innerHTML = '<div class="empty-msg">' + msg + '</div>';
}

function renderCard(title, statusHtml, detailHtml) {
  return '<div class="family-card"><div class="fam-header"><span class="fam-name">' + title +
    '</span>' + statusHtml + '</div><div class="fam-members">' + detailHtml + '</div></div>';
}

// ══════════════════════════════════════════════════════════════════════════════
// RECIPIENTS
// ══════════════════════════════════════════════════════════════════════════════

// Each entry: { type: 'email'|'all'|'family'|'single'|'group'|'contact', value: string, label: string }
var recipientList = [];
var acCache = { contacts: [], families: [], groups: [] };
var acHighlightIdx = -1;

function addRecipient(entry) {
  const key = entry.type + ':' + entry.value;
  if (recipientList.some(r => (r.type + ':' + r.value) === key)) return;
  recipientList.push(entry);
  renderRecipients();
}

function removeRecipient(idx) {
  recipientList.splice(idx, 1);
  renderRecipients();
}

function renderRecipients() {
  const el = document.getElementById('recipient-chips');
  el.innerHTML = recipientList.map((r, i) => {
    return '<span class="recipient-chip">' + esc(r.label) +
      ' <span class="remove" onclick="removeRecipient(' + i + ')">&times;</span></span>';
  }).join('');
}

// Build autocomplete search cache from loaded data
async function refreshAcCache() {
  await waitForApi();
  acCache.contacts = await pywebview.api.get_contacts();
  acCache.families = await pywebview.api.get_families();
  acCache.groups = await pywebview.api.get_groups();
}

function acSearch(query) {
  const q = query.toLowerCase();
  const results = [];

  // Search contacts (name + email)
  for (const c of acCache.contacts) {
    if (c.name.toLowerCase().includes(q) || c.email.toLowerCase().includes(q)) {
      results.push({ type: 'contact', id: c.id, name: c.name, detail: c.email, label: c.name + ' <' + c.email + '>' });
    }
  }

  // Search families
  for (const f of acCache.families) {
    if (f.name.toLowerCase().includes(q)) {
      const count = f.members ? f.members.length : 0;
      results.push({ type: 'family', id: f.id, name: f.name, detail: count + ' member' + (count !== 1 ? 's' : ''), label: 'Family: ' + f.name });
    }
  }

  // Search groups
  for (const g of acCache.groups) {
    if (g.name.toLowerCase().includes(q)) {
      const count = g.members ? g.members.length : 0;
      results.push({ type: 'group', id: g.id, name: g.name, detail: count + ' member' + (count !== 1 ? 's' : ''), label: 'Group: ' + g.name });
    }
  }

  return results.slice(0, 15);
}

function showAcDropdown(results) {
  const dd = document.getElementById('ac-dropdown');
  if (!results.length) { dd.classList.remove('show'); return; }
  acHighlightIdx = -1;
  dd.innerHTML = results.map((r, i) => {
    return '<div class="ac-item" data-idx="' + i + '">' +
      '<span class="ac-type ' + r.type + '">' + r.type + '</span>' +
      '<span class="ac-name">' + esc(r.name) + '</span>' +
      '<span class="ac-detail">' + esc(r.detail) + '</span>' +
      '</div>';
  }).join('');
  dd.classList.add('show');

  // Attach click handlers
  dd.querySelectorAll('.ac-item').forEach((item, i) => {
    item.addEventListener('mousedown', function(e) {
      e.preventDefault();  // prevent input blur
      selectAcResult(results[i]);
    });
  });
}

function hideAcDropdown() {
  document.getElementById('ac-dropdown').classList.remove('show');
  acHighlightIdx = -1;
}

function selectAcResult(r) {
  const input = document.getElementById('recipient-input');
  if (r.type === 'contact') {
    addRecipient({ type: 'email', value: r.detail, label: r.label });
  } else if (r.type === 'family') {
    // Add all family members
    const fam = acCache.families.find(f => f.id === r.id);
    if (fam && fam.members) {
      for (const m of fam.members) {
        addRecipient({ type: 'email', value: m.email, label: m.name + ' <' + m.email + '>' });
      }
    }
  } else if (r.type === 'group') {
    addRecipient({ type: 'group', value: String(r.id), label: r.label });
  }
  input.value = '';
  hideAcDropdown();
}

// Wire up input events
document.addEventListener('DOMContentLoaded', function() {
  const input = document.getElementById('recipient-input');
  if (!input) return;
  var acDebounce = null;

  input.addEventListener('input', function() {
    clearTimeout(acDebounce);
    const val = input.value.trim();
    if (val.length < 1) { hideAcDropdown(); return; }
    acDebounce = setTimeout(function() {
      const results = acSearch(val);
      showAcDropdown(results);
    }, 150);
  });

  input.addEventListener('keydown', function(e) {
    const dd = document.getElementById('ac-dropdown');
    const items = dd.querySelectorAll('.ac-item');

    if (dd.classList.contains('show') && items.length) {
      if (e.key === 'ArrowDown') {
        e.preventDefault();
        acHighlightIdx = Math.min(acHighlightIdx + 1, items.length - 1);
        items.forEach((it, i) => it.classList.toggle('highlighted', i === acHighlightIdx));
        items[acHighlightIdx].scrollIntoView({ block: 'nearest' });
        return;
      }
      if (e.key === 'ArrowUp') {
        e.preventDefault();
        acHighlightIdx = Math.max(acHighlightIdx - 1, 0);
        items.forEach((it, i) => it.classList.toggle('highlighted', i === acHighlightIdx));
        items[acHighlightIdx].scrollIntoView({ block: 'nearest' });
        return;
      }
      if (e.key === 'Enter' && acHighlightIdx >= 0) {
        e.preventDefault();
        const results = acSearch(input.value.trim());
        selectAcResult(results[acHighlightIdx]);
        return;
      }
      if (e.key === 'Escape') { hideAcDropdown(); return; }
    }

    if (e.key === 'Enter') {
      e.preventDefault();
      const val = input.value.trim();
      if (!val) return;
      val.split(/[,;]+/).forEach(function(email) {
        email = email.trim();
        if (email) addRecipient({ type: 'email', value: email, label: email });
      });
      input.value = '';
      hideAcDropdown();
    }
  });

  input.addEventListener('blur', function() {
    setTimeout(hideAcDropdown, 200);
  });

  input.addEventListener('focus', function() {
    const val = input.value.trim();
    if (val.length >= 1) {
      const results = acSearch(val);
      showAcDropdown(results);
    }
  });
});

// Handle dropdown selection
window.onTargetSelect = function() {
  const sel = document.getElementById('target-select');
  const val = sel.value;
  if (!val) return;
  const label = sel.options[sel.selectedIndex].textContent;
  if (val.startsWith('group:')) {
    addRecipient({ type: 'group', value: val.split(':')[1], label: label });
  } else {
    addRecipient({ type: val, value: val, label: label });
  }
  sel.value = '';  // Reset dropdown
};

function getRecipientSelection() {
  // Returns { targetType, targetId, contactIds, manualEmails } for dispatch
  const manualEmails = [];
  let targetType = null, targetId = null;

  for (const r of recipientList) {
    if (r.type === 'email') {
      manualEmails.push(r.value);
    } else if (r.type === 'group') {
      targetType = 'group';
      targetId = parseInt(r.value);
    } else {
      // all, family, single
      targetType = r.type;
    }
  }
  return { targetType, targetId, manualEmails };
}

// ══════════════════════════════════════════════════════════════════════════════
// SETTINGS
// ══════════════════════════════════════════════════════════════════════════════

async function loadSettings() {
  await waitForApi();
  const s = await pywebview.api.get_settings();
  document.getElementById('s-email').value = s.email;
  document.getElementById('s-password').placeholder =
    s.has_password ? 'App Password (saved — leave blank to keep)' : 'App Password';
  const tzSel = document.getElementById('s-timezone');
  if (s.timezone) {
    tzSel.value = s.timezone;
    // If the value doesn't match any option, it stays at the first
    if (tzSel.value !== s.timezone) {
      // Add it as a custom option
      const opt = document.createElement('option');
      opt.value = s.timezone;
      opt.textContent = s.timezone;
      tzSel.appendChild(opt);
      tzSel.value = s.timezone;
    }
  }
}

window.saveSettings = async function() {
  const email = document.getElementById('s-email').value.trim();
  const password = document.getElementById('s-password').value.trim();
  if (!email) { showToast('Enter an email address.', 'error'); return; }

  const res = await pywebview.api.save_settings(email, password);
  if (res.ok) {
    showToast('Settings saved.', 'success');
    document.getElementById('s-password').value = '';
    loadSettings();
  } else {
    showToast(res.error, 'error');
  }
};

window.testConnection = async function() {
  const email = document.getElementById('s-email').value.trim();
  const password = document.getElementById('s-password').value.trim();
  if (!email || !password) {
    showToast('Enter both email and password to test.', 'error');
    return;
  }
  document.getElementById('settings-status').textContent = 'Testing connection...';
  const res = await pywebview.api.test_email_connection(email, password);
  if (res.ok) {
    document.getElementById('settings-status').innerHTML = '<span style="color:var(--success)">Connection successful!</span>';
  } else {
    document.getElementById('settings-status').innerHTML = '<span style="color:var(--danger)">Failed: ' + esc(res.error) + '</span>';
  }
};

window.saveTimezone = async function() {
  const tz = document.getElementById('s-timezone').value;
  const res = await pywebview.api.save_timezone(tz);
  if (res.ok) {
    showToast('Timezone saved: ' + tz, 'success');
  } else {
    showToast(res.error, 'error');
  }
};

// ══════════════════════════════════════════════════════════════════════════════
// GROUPS
// ══════════════════════════════════════════════════════════════════════════════

var cachedGroups = [];
async function loadGroups() {
  await waitForApi();
  cachedGroups = await pywebview.api.get_groups();

  renderGroupList(cachedGroups);
  renderGroupDetail();

  // Update target selector with group options
  const tsel = document.getElementById('target-select');
  const base = '<option value="">-- Quick select --</option><option value="all">All Contacts</option><option value="family">All Families</option><option value="single">All Singles</option>';
  const groupOpts = cachedGroups.map(g => '<option value="group:' + g.id + '">Group: ' + esc(g.name) + '</option>').join('');
  tsel.innerHTML = base + groupOpts;
}

var _selectedGroupId = null;

function renderGroupList(groups) {
  const el = document.getElementById('group-list');
  if (!groups.length) {
    renderEmpty(el, 'No groups yet.');
  } else {
    el.innerHTML = groups.map(g => {
      var active = g.id === _selectedGroupId ? ' active' : '';
      return '<div class="group-item' + active + '" onclick="selectGroup(' + g.id + ')">' +
        '<span class="group-name">' + esc(g.name) + '</span>' +
        '<span class="group-count">' + g.members.length + '</span>' +
      '</div>';
    }).join('');
  }
}

window.selectGroup = function(id) {
  _selectedGroupId = id;
  // Re-render left list to update active state
  var q = document.getElementById('search-groups').value.toLowerCase().trim();
  if (q) { filterGroups(); } else { renderGroupList(cachedGroups); }
  // Render right detail panel
  renderGroupDetail();
};

function renderGroupDetail() {
  var titleEl = document.getElementById('group-detail-title');
  var actionsEl = document.getElementById('group-detail-actions');
  var listEl = document.getElementById('group-member-list');
  var searchEl = document.getElementById('search-group-members');

  if (!_selectedGroupId) {
    titleEl.textContent = 'Select a group';
    actionsEl.innerHTML = '';
    listEl.innerHTML = '<div class="empty-msg">Click a group to view its members</div>';
    searchEl.value = '';
    return;
  }

  var group = cachedGroups.find(g => g.id === _selectedGroupId);
  if (!group) { _selectedGroupId = null; renderGroupDetail(); return; }

  titleEl.textContent = group.name;
  actionsEl.innerHTML =
    '<button class="contact-edit-btn" onclick="editGroup(' + group.id + ')" title="Edit">&#9998;</button> ' +
    '<button class="btn btn-danger btn-sm" onclick="deleteGroup(' + group.id + ')">Delete</button>';

  var mq = searchEl.value.toLowerCase().trim();
  var members = group.members;
  if (mq) {
    members = members.filter(m => m.name.toLowerCase().includes(mq) || m.email.toLowerCase().includes(mq));
  }

  if (!members.length) {
    listEl.innerHTML = '<div class="empty-msg">' + (mq ? 'No matching members' : 'No members in this group') + '</div>';
  } else {
    listEl.innerHTML = members.map(m =>
      '<div class="group-member-row">' +
        '<span class="gm-name">' + esc(m.name) + '</span>' +
        '<span class="gm-email">' + esc(m.email) + '</span>' +
      '</div>'
    ).join('');
  }
}

window.filterGroupMembers = function() {
  renderGroupDetail();
};

window.filterGroups = function() {
  const q = document.getElementById('search-groups').value.toLowerCase().trim();
  if (!q) { renderGroupList(cachedGroups); return; }
  const filtered = cachedGroups.filter(g =>
    g.name.toLowerCase().includes(q) ||
    g.members.some(m => m.name.toLowerCase().includes(q))
  );
  renderGroupList(filtered);
};

window.openCreateGroup = function() {
  var contactOpts = '<option value="">Select contact</option>';
  cachedContacts.forEach(function(c) {
    contactOpts += '<option value="' + c.id + '">' + esc(c.name) + ' (' + esc(c.email) + ')</option>';
  });
  var famOpts = '<option value="">Select family</option>';
  cachedFamilies.forEach(function(f) {
    famOpts += '<option value="' + f.id + '">' + esc(f.name) + ' (' + f.members.length + ' members)</option>';
  });
  showCreateModal(
    '<h3>New Group</h3>' +
    '<input type="text" id="cm-gname" placeholder="Group Name">' +
    '<div style="border-top:1px solid var(--border);margin-top:4px;padding-top:8px;">' +
      '<div style="font-size:12px;color:var(--text-muted);margin-bottom:4px;">Add members (optional):</div>' +
      '<div class="edit-row">' +
        '<select id="cm-add-type" onchange="cmGroupTypeChange()" style="width:auto;"><option value="contact">Contact</option><option value="family">Family</option></select>' +
        '<select id="cm-g-contact">' + contactOpts + '</select>' +
        '<select id="cm-g-family" style="display:none;">' + famOpts + '</select>' +
      '</div>' +
      '<div id="cm-g-members" style="margin-top:4px;font-size:11px;color:var(--text-muted);"></div>' +
    '</div>' +
    '<div class="edit-actions"><button class="btn btn-sm" onclick="closeCreateModal()">Cancel</button><button class="btn btn-primary btn-sm" onclick="submitCreateGroup()">Add</button></div>'
  );
  document.getElementById('cm-gname').focus();
};

var _cmGroupMembers = [];
window.cmGroupTypeChange = function() {
  var t = document.getElementById('cm-add-type').value;
  document.getElementById('cm-g-contact').style.display = t === 'contact' ? '' : 'none';
  document.getElementById('cm-g-family').style.display = t === 'family' ? '' : 'none';
};

window.submitCreateGroup = async function() {
  var name = document.getElementById('cm-gname').value.trim();
  if (!name) { showToast('Enter a group name.', 'error'); return; }
  var res = await pywebview.api.add_group(name);
  if (!res.ok) { showToast(res.error || 'Group already exists.', 'error'); return; }
  // Get the new group's id by reloading
  var groups = await pywebview.api.get_groups();
  var newGroup = groups.find(function(g) { return g.name === name; });
  if (newGroup) {
    // Add selected member
    var addType = document.getElementById('cm-add-type').value;
    if (addType === 'family') {
      var fid = document.getElementById('cm-g-family').value;
      if (fid) await pywebview.api.add_family_to_group(newGroup.id, parseInt(fid));
    } else {
      var cid = document.getElementById('cm-g-contact').value;
      if (cid) await pywebview.api.add_group_member(newGroup.id, parseInt(cid));
    }
  }
  closeCreateModal();
  loadGroups();
  refreshAcCache();
};

window.deleteGroup = async function(id) {
  if (!confirm('Delete this group?')) return;
  if (_selectedGroupId === id) _selectedGroupId = null;
  await pywebview.api.delete_group(id);
  loadGroups();
  refreshAcCache();
};

// ── Group edit modal ──
var _egGroupId = null;
var _egOriginalMembers = [];
var _egPendingRemoves = [];
var _egPendingAdds = [];

window.editGroup = function(id) {
  var group = cachedGroups.find(g => g.id === id);
  if (!group) return;
  _egGroupId = id;
  _egOriginalMembers = group.members.slice();
  _egPendingRemoves = [];
  _egPendingAdds = [];
  showCreateModal(
    '<h3>Edit Group</h3>' +
    '<input type="text" id="eg-name" value="' + esc(group.name).replace(/"/g, '&quot;') + '" placeholder="Group Name">' +
    '<div class="edit-section-label">Members</div>' +
    '<div class="edit-member-list" id="eg-members"></div>' +
    '<div class="edit-section-label">Add Members</div>' +
    '<input type="text" id="eg-search" placeholder="Search contacts..." oninput="egSearchContacts()" style="padding:7px 10px;font-size:13px;border:1px solid var(--border);border-radius:6px;background:var(--bg);color:var(--text);width:100%;box-sizing:border-box;">' +
    '<div class="edit-search-results" id="eg-search-results"></div>' +
    '<div class="edit-actions"><button class="btn btn-sm" onclick="closeCreateModal()">Cancel</button><button class="btn btn-primary btn-sm" onclick="saveEditGroup()">Save</button></div>'
  );
  egRenderMembers();
  document.getElementById('eg-search-results').innerHTML = '';
};

window.egRenderMembers = function() {
  var el = document.getElementById('eg-members');
  var current = _egOriginalMembers.filter(m => _egPendingRemoves.indexOf(m.id) === -1);
  var all = current.concat(_egPendingAdds);
  if (!all.length) { el.innerHTML = '<span style="font-size:11px;color:var(--text-muted);">No members</span>'; return; }
  el.innerHTML = all.map(function(m) {
    var isNew = _egPendingAdds.some(a => a.id === m.id);
    return '<span class="edit-member-pill">' + esc(m.name) +
      ' <span class="remove-x" onclick="egRemoveMember(' + m.id + ',' + (isNew ? 'true' : 'false') + ')">&times;</span></span>';
  }).join('');
};

window.egRemoveMember = function(memberId, isNew) {
  if (isNew) {
    _egPendingAdds = _egPendingAdds.filter(a => a.id !== memberId);
  } else {
    if (_egPendingRemoves.indexOf(memberId) === -1) _egPendingRemoves.push(memberId);
  }
  egRenderMembers();
  egSearchContacts();
};

window.egSearchContacts = function() {
  var el = document.getElementById('eg-search-results');
  var q = document.getElementById('eg-search').value.toLowerCase().trim();
  if (!q) { el.innerHTML = ''; return; }
  var currentIds = _egOriginalMembers.filter(m => _egPendingRemoves.indexOf(m.id) === -1).map(m => m.id);
  var addedIds = _egPendingAdds.map(a => a.id);
  var exclude = currentIds.concat(addedIds);
  var matches = cachedContacts.filter(function(c) {
    if (exclude.indexOf(c.id) !== -1) return false;
    return c.name.toLowerCase().includes(q) || c.email.toLowerCase().includes(q);
  }).slice(0, 10);
  if (!matches.length) { el.innerHTML = '<div class="esr-item" style="color:var(--text-muted);cursor:default;">No matches</div>'; return; }
  el.innerHTML = matches.map(function(c) {
    return '<div class="esr-item" onclick="egAddContact(' + c.id + ')">' + esc(c.name) + '<span class="esr-email">' + esc(c.email) + '</span></div>';
  }).join('');
};

window.egAddContact = function(contactId) {
  var c = cachedContacts.find(x => x.id === contactId);
  if (!c) return;
  // If it was a removed original member, un-remove it
  var wasOriginal = _egOriginalMembers.some(m => m.id === contactId);
  if (wasOriginal) {
    _egPendingRemoves = _egPendingRemoves.filter(id => id !== contactId);
  } else {
    if (!_egPendingAdds.some(a => a.id === contactId)) {
      _egPendingAdds.push({id: c.id, name: c.name, email: c.email});
    }
  }
  egRenderMembers();
  document.getElementById('eg-search').value = '';
  document.getElementById('eg-search-results').innerHTML = '';
};

window.saveEditGroup = async function() {
  var newName = document.getElementById('eg-name').value.trim();
  if (!newName) { showToast('Group name cannot be empty.', 'error'); return; }
  var group = cachedGroups.find(g => g.id === _egGroupId);
  if (group && newName !== group.name) {
    var res = await pywebview.api.rename_group(_egGroupId, newName);
    if (!res.ok) { showToast(res.error || 'Rename failed.', 'error'); return; }
  }
  for (var i = 0; i < _egPendingRemoves.length; i++) {
    await pywebview.api.remove_group_member(_egGroupId, _egPendingRemoves[i]);
  }
  for (var j = 0; j < _egPendingAdds.length; j++) {
    await pywebview.api.add_group_member(_egGroupId, _egPendingAdds[j].id);
  }
  closeCreateModal();
  loadGroups();
  loadContacts();
  refreshAcCache();
};

// ══════════════════════════════════════════════════════════════════════════════
// SCHEDULED EMAILS
// ══════════════════════════════════════════════════════════════════════════════

var schedMonth = new Date().getMonth();
var schedYear = new Date().getFullYear();
var _schedEmails = [];

async function loadScheduled() {
  await waitForApi();
  _schedEmails = await pywebview.api.get_scheduled_emails_with_recipients();
  renderScheduleCalendar();
}

function renderScheduleCalendar() {
  const monthNames = ['January','February','March','April','May','June','July','August','September','October','November','December'];
  const dayNames = ['Sun','Mon','Tue','Wed','Thu','Fri','Sat'];
  document.getElementById('sched-month-label').textContent = monthNames[schedMonth] + ' ' + schedYear;
  const el = document.getElementById('scheduled-calendar');

  const daysInMonth = new Date(schedYear, schedMonth + 1, 0).getDate();
  const today = new Date(); today.setHours(0,0,0,0);
  const todayKey = today.getFullYear() + '-' + String(today.getMonth()+1).padStart(2,'0') + '-' + String(today.getDate()).padStart(2,'0');

  // Group emails by date
  const byDate = {};
  _schedEmails.forEach(function(e) {
    const key = e.scheduled_at.split('T')[0];
    if (!byDate[key]) byDate[key] = [];
    byDate[key].push(e);
  });

  var html = '';
  for (var day = 1; day <= daysInMonth; day++) {
    var d = new Date(schedYear, schedMonth, day);
    var key = schedYear + '-' + String(schedMonth+1).padStart(2,'0') + '-' + String(day).padStart(2,'0');
    var isToday = key === todayKey;
    var label = dayNames[d.getDay()] + ', ' + monthNames[schedMonth].substring(0,3) + ' ' + day;
    var dayEmails = byDate[key] || [];

    html += '<div class="sched-day-row' + (isToday ? ' sched-day-today' : '') + '" id="sched-day-' + key + '">';
    html += '<div class="sched-day-header"><span class="sched-day-label">' + label + (isToday ? ' (Today)' : '') + '</span>';
    html += '<button class="sched-add-btn" style="width:auto;margin:0;padding:2px 8px;font-size:14px;" onclick="scheduleNewForDate(\'' + key + '\')" title="Schedule email for this day">+</button>';
    html += '</div>';

    if (dayEmails.length) {
      html += '<div class="sched-day-cards">';
      dayEmails.forEach(function(e) {
        var statusColor = e.status === 'sent' ? 'var(--success)' : e.status === 'pending' ? 'var(--accent)' : 'var(--danger)';
        var time = (e.scheduled_at.split('T')[1] || '').substring(0, 5);
        var recipNames = e.recipients.map(function(r) { return r.name || r.email; });
        var recipShort = recipNames.length <= 2 ? recipNames.join(', ') : recipNames.slice(0,2).join(', ') + ' +' + (recipNames.length - 2) + ' more';
        var cancelBtn = e.status === 'pending' ? '<button class="btn btn-danger btn-sm" style="font-size:10px;padding:1px 6px;margin-left:6px;" onclick="event.stopPropagation();cancelScheduled(' + e.id + ')">Cancel</button>' : '';

        // Recurrence label
        var recInfo = '';
        if (e.recurrence && e.recurrence.type !== 'once') {
          var dn = ['Su','Mo','Tu','We','Th','Fr','Sa'];
          recInfo = e.recurrence.type.replace(/_/g, ' ');
          if (e.recurrence.days && e.recurrence.days.length) recInfo += ' (' + e.recurrence.days.map(function(d){return dn[d];}).join(', ') + ')';
          if (e.recurrence.day_of_month) recInfo += ' (day ' + e.recurrence.day_of_month + ')';
          if (e.recurrence.end_date) recInfo += ' until ' + e.recurrence.end_date.split('T')[0];
        }

        // Tooltip content
        var fullRecip = e.recipients.map(function(r) { return r.name ? esc(r.name) + ' &lt;' + esc(r.email) + '&gt;' : esc(r.email); }).join('<br>');
        var tipHtml = '<strong>Recipients (' + e.recipients.length + '):</strong><br>' + (fullRecip || '<em>None resolved</em>');
        if (recInfo) tipHtml += '<br><br><strong>Recurrence:</strong> ' + esc(recInfo);
        if (e.sent_at) tipHtml += '<br><strong>Sent at:</strong> ' + esc(e.sent_at);
        if (e.result) tipHtml += '<br><strong>Result:</strong> ' + esc(typeof e.result === 'string' ? e.result : JSON.stringify(e.result));

        var clickable = e.status !== 'pending';
        var clickAttr = clickable ? ' clickable" onclick="showEmailPreview(' + e.id + ')"' : '"';
        html += '<div class="sched-email-card' + (clickable ? ' clickable' : '') + '" data-status="' + e.status + '"' + (clickable ? ' onclick="showEmailPreview(' + e.id + ')"' : '') + '>';
        html += '<div class="sched-card-subject">' + esc(e.subject) + '</div>';
        html += '<div class="sched-card-meta">';
        html += '<span>' + esc(time) + '</span>';
        html += '<span class="sched-card-status" style="color:' + statusColor + '">' + e.status + '</span>';
        html += '<span>' + esc(recipShort) + '</span>';
        html += cancelBtn;
        html += '</div>';
        html += '<div class="sched-tooltip">' + tipHtml + '</div>';
        html += '</div>';
      });
      html += '</div>';
    }
    html += '</div>';
  }

  el.innerHTML = html;

  // Scroll to today if viewing current month
  if (schedMonth === today.getMonth() && schedYear === today.getFullYear()) {
    var todayEl = document.getElementById('sched-day-' + todayKey);
    if (todayEl) todayEl.scrollIntoView({ behavior: 'auto', block: 'start' });
  }
}

window.schedChangeMonth = function(delta) {
  schedMonth += delta;
  if (schedMonth > 11) { schedMonth = 0; schedYear++; }
  if (schedMonth < 0) { schedMonth = 11; schedYear--; }
  renderScheduleCalendar();
};

window.scheduleNewForDate = function(dateStr) {
  clearComposer();
  var dt = dateStr + 'T09:00';
  document.getElementById('sched-datetime').value = dt;
  document.getElementById('subject').focus();
};

function setComposerDisabled(disabled) {
  var btns = [document.getElementById('btn-save-template'), document.getElementById('btn-schedule'), document.getElementById('btn-send-now')];
  btns.forEach(function(b) {
    if (!b) return;
    b.disabled = disabled;
    b.style.opacity = disabled ? '0.4' : '';
    b.style.pointerEvents = disabled ? 'none' : '';
  });
}

function clearComposer() {
  document.getElementById('subject').value = '';
  quill.setText('');
  recipientList = [];
  renderRecipients();
  document.getElementById('sched-datetime').value = '';
  document.getElementById('recurrence-type').value = 'once';
  if (window.onRecurrenceChange) onRecurrenceChange();
  document.getElementById('recurrence-end-date').value = '';
  attachedFiles = [];
  renderAttachChips();
  document.getElementById('template-select').value = '';
  document.getElementById('target-select').value = '';
  setComposerDisabled(false);
}

window.showEmailPreview = async function(emailId) {
  var detail = await pywebview.api.get_scheduled_email_detail(emailId);
  if (!detail) return;
  // Fill subject
  document.getElementById('subject').value = detail.subject || '';
  // Fill editor body
  if (detail.html_body) {
    quill.root.innerHTML = detail.html_body;
  } else {
    quill.setText(detail.plain_text || '');
  }
  // Fill recipients
  recipientList = detail.recipients.map(function(r) {
    return { type: 'email', value: r.email, label: r.name ? r.name + ' <' + r.email + '>' : r.email };
  });
  renderRecipients();
  // Fill schedule datetime
  if (detail.scheduled_at) {
    document.getElementById('sched-datetime').value = detail.scheduled_at.substring(0, 16);
  }
  // Disable action buttons
  setComposerDisabled(true);
};

// ── Recurrence controls ──

window.onRecurrenceChange = function() {
  const rtype = document.getElementById('recurrence-type').value;
  const dayPicker = document.getElementById('day-picker');
  const monthlyDay = document.getElementById('monthly-day');
  dayPicker.classList.toggle('hidden', rtype !== 'weekly' && rtype !== 'every_other_week');
  monthlyDay.classList.toggle('hidden', rtype !== 'monthly');
};

// Day-of-week toggle buttons
document.querySelectorAll('.day-btn').forEach(btn => {
  btn.addEventListener('click', function() {
    this.classList.toggle('active');
  });
});

function getRecurrenceConfig() {
  const rtype = document.getElementById('recurrence-type').value;
  if (rtype === 'once') return null;

  const config = { type: rtype };

  if (rtype === 'weekly' || rtype === 'every_other_week') {
    const days = [];
    document.querySelectorAll('.day-btn.active').forEach(btn => {
      days.push(parseInt(btn.dataset.day));
    });
    if (!days.length) return null;  // will be caught in validation
    config.days = days;
  }

  if (rtype === 'monthly') {
    config.day_of_month = parseInt(document.getElementById('month-day').value) || 1;
  }

  const endDate = document.getElementById('recurrence-end-date').value;
  if (endDate) config.end_date = endDate;

  return config;
}

async function doScheduleEmail() {
  const subject = document.getElementById('subject').value.trim();
  const htmlBody = getEditorHTML();
  const plainText = getEditorPlainText();
  const dtVal = document.getElementById('sched-datetime').value;

  if (!subject || !plainText.trim()) { showToast('Fill in subject and body.', 'error'); return; }
  if (!dtVal) { showToast('Pick a date and time.', 'error'); return; }

  const rtype = document.getElementById('recurrence-type').value;
  if ((rtype === 'weekly' || rtype === 'every_other_week') && !document.querySelectorAll('.day-btn.active').length) {
    showToast('Select at least one day of the week.', 'error');
    return;
  }

  const scheduledAt = dtVal.length === 16 ? dtVal + ':00' : dtVal;
  const { targetType, targetId, manualEmails } = getRecipientSelection();

  if (!recipientList.length) {
    showToast('Add at least one recipient.', 'error');
    return;
  }

  const recurrence = getRecurrenceConfig();
  const paths = attachedFiles.map(f => f.path);
  const res = await pywebview.api.schedule_email(
    subject, htmlBody, plainText, targetType || 'manual', targetId, [], paths, scheduledAt,
    recurrence, manualEmails
  );
  if (res.ok) {
    const label = recurrence ? recurrence.type.replace(/_/g, ' ') : 'one-time';
    showToast('Email scheduled (' + label + ') starting ' + dtVal, 'success');
    loadScheduled();
  } else {
    showToast(res.error, 'error');
  }
}

window.scheduleEmail = async function() {
  const subject = document.getElementById('subject').value.trim();
  const htmlBody = getEditorHTML();
  if (!subject) { showToast('Enter a subject first.', 'error'); return; }

  const dtVal = document.getElementById('sched-datetime').value;
  if (!dtVal) { showToast('Pick a date and time.', 'error'); return; }

  const sel = document.getElementById('template-select');
  const currentId = sel.value ? parseInt(sel.value) : null;
  const currentName = sel.value ? sel.options[sel.selectedIndex].textContent : '';
  const nameInput = document.getElementById('template-modal-name');
  const recipients = recipientList.slice();

  if (currentId) {
    document.getElementById('template-modal-title').textContent = 'Save & Schedule';
    document.getElementById('template-modal-desc').textContent = 'Template "' + currentName + '" is loaded. Override it or save as new, then schedule?';
    nameInput.classList.add('hidden');
    nameInput.value = '';
    showModalButtons(true, true, false);

    setModalHandler('template-modal-override', async function() {
      await pywebview.api.update_template(currentId, subject, htmlBody, recipients);
      showToast('Template "' + currentName + '" updated.', 'success');
      loadTemplates();
      setTimeout(() => { document.getElementById('template-select').value = currentId; }, 100);
      closeTemplateModal();
      await doScheduleEmail();
    });

    setModalHandler('template-modal-new', function() {
      document.getElementById('template-modal-title').textContent = 'Save as New & Schedule';
      document.getElementById('template-modal-desc').textContent = 'Enter a name for the new template:';
      nameInput.classList.remove('hidden');
      nameInput.value = '';
      nameInput.focus();
      showModalButtons(false, false, true);
      setModalHandler('template-modal-save', async function() {
        const newName = nameInput.value.trim();
        if (!newName) { showToast('Enter a template name.', 'error'); return; }
        await pywebview.api.save_template(newName, subject, htmlBody, recipients);
        showToast('Template "' + newName + '" saved.', 'success');
        loadTemplates();
        closeTemplateModal();
        await doScheduleEmail();
      });
    });
  } else {
    document.getElementById('template-modal-title').textContent = 'Save & Schedule';
    document.getElementById('template-modal-desc').textContent = 'Enter a name to save this as a template, then schedule:';
    nameInput.classList.remove('hidden');
    nameInput.value = '';
    showModalButtons(false, false, true);

    setModalHandler('template-modal-save', async function() {
      const newName = nameInput.value.trim();
      if (!newName) { showToast('Enter a template name.', 'error'); return; }
      await pywebview.api.save_template(newName, subject, htmlBody, recipients);
      showToast('Template "' + newName + '" saved.', 'success');
      loadTemplates();
      closeTemplateModal();
      await doScheduleEmail();
    });
  }

  document.getElementById('template-save-modal').classList.add('show');
  if (!nameInput.classList.contains('hidden')) nameInput.focus();
};

window.cancelScheduled = async function(id) {
  if (!confirm('Cancel this scheduled email?')) return;
  await pywebview.api.cancel_scheduled_email(id);
  loadScheduled();
};

// ══════════════════════════════════════════════════════════════════════════════
// TEMPLATES
// ══════════════════════════════════════════════════════════════════════════════

async function loadTemplates() {
  await waitForApi();
  const templates = await pywebview.api.get_templates();
  const sel = document.getElementById('template-select');
  sel.innerHTML = '<option value="">Load template...</option>' +
    templates.map(t => '<option value="' + t.id + '" data-subject="' + esc(t.subject).replace(/"/g, '&quot;') +
      '" data-html="' + esc(t.html_body).replace(/"/g, '&quot;') + '">' + esc(t.name) + '</option>').join('');
}

window.loadTemplate = async function() {
  const sel = document.getElementById('template-select');
  const opt = sel.options[sel.selectedIndex];
  if (!opt.value) return;

  // Fetch fresh template data from API
  const templates = await pywebview.api.get_templates();
  const tmpl = templates.find(t => t.id === parseInt(opt.value));
  if (!tmpl) return;

  document.getElementById('subject').value = tmpl.subject;
  quill.root.innerHTML = tmpl.html_body;

  // Restore saved recipients
  recipientList = [];
  if (tmpl.recipients && Array.isArray(tmpl.recipients)) {
    for (const r of tmpl.recipients) {
      addRecipient(r);
    }
  }
  renderRecipients();
};

function closeTemplateModal() {
  document.getElementById('template-save-modal').classList.remove('show');
}

// Tracked handlers for modal buttons so we can remove them cleanly
let _modalHandlers = {};

function setModalHandler(id, handler) {
  const el = document.getElementById(id);
  if (_modalHandlers[id]) el.removeEventListener('click', _modalHandlers[id]);
  _modalHandlers[id] = handler;
  el.addEventListener('click', handler);
}

function showModalButtons(override, newBtn, save) {
  document.getElementById('template-modal-override').classList.toggle('hidden', !override);
  document.getElementById('template-modal-new').classList.toggle('hidden', !newBtn);
  document.getElementById('template-modal-save').classList.toggle('hidden', !save);
}

window.saveAsTemplate = async function() {
  const subject = document.getElementById('subject').value.trim();
  const htmlBody = getEditorHTML();
  if (!subject) { showToast('Enter a subject first.', 'error'); return; }

  const sel = document.getElementById('template-select');
  const currentId = sel.value ? parseInt(sel.value) : null;
  const currentName = sel.value ? sel.options[sel.selectedIndex].textContent : '';
  const recipients = recipientList.slice();

  const nameInput = document.getElementById('template-modal-name');

  if (currentId) {
    document.getElementById('template-modal-title').textContent = 'Save Template';
    document.getElementById('template-modal-desc').textContent = 'A template "' + currentName + '" is currently loaded. Override it or save as a new template?';
    nameInput.classList.add('hidden');
    nameInput.value = '';
    showModalButtons(true, true, false);

    setModalHandler('template-modal-override', async function() {
      const res = await pywebview.api.update_template(currentId, subject, htmlBody, recipients);
      if (res.ok) {
        showToast('Template "' + currentName + '" updated.', 'success');
        loadTemplates();
        setTimeout(() => { document.getElementById('template-select').value = currentId; }, 100);
      } else { showToast(res.error, 'error'); }
      closeTemplateModal();
    });

    setModalHandler('template-modal-new', function() {
      document.getElementById('template-modal-title').textContent = 'Save as New Template';
      document.getElementById('template-modal-desc').textContent = 'Enter a name for the new template:';
      nameInput.classList.remove('hidden');
      nameInput.value = '';
      nameInput.focus();
      showModalButtons(false, false, true);
      setModalHandler('template-modal-save', async function() {
        const newName = nameInput.value.trim();
        if (!newName) { showToast('Enter a template name.', 'error'); return; }
        const res = await pywebview.api.save_template(newName, subject, htmlBody, recipients);
        if (res.ok) {
          showToast('Template "' + newName + '" saved.', 'success');
          loadTemplates();
        } else { showToast(res.error, 'error'); }
        closeTemplateModal();
      });
    });
  } else {
    document.getElementById('template-modal-title').textContent = 'Save Template';
    document.getElementById('template-modal-desc').textContent = 'Enter a name for the new template:';
    nameInput.classList.remove('hidden');
    nameInput.value = '';
    showModalButtons(false, false, true);

    setModalHandler('template-modal-save', async function() {
      const newName = nameInput.value.trim();
      if (!newName) { showToast('Enter a template name.', 'error'); return; }
      const res = await pywebview.api.save_template(newName, subject, htmlBody, recipients);
      if (res.ok) {
        showToast('Template "' + newName + '" saved.', 'success');
        loadTemplates();
      } else { showToast(res.error, 'error'); }
      closeTemplateModal();
    });
  }

  document.getElementById('template-save-modal').classList.add('show');
  if (!nameInput.classList.contains('hidden')) nameInput.focus();
};

window.deleteCurrentTemplate = async function() {
  const sel = document.getElementById('template-select');
  if (!sel.value) { showToast('Select a template to delete.', 'error'); return; }
  if (!confirm('Delete this template?')) return;
  await pywebview.api.delete_template(parseInt(sel.value));
  loadTemplates();
  showToast('Template deleted.', 'success');
};

// ══════════════════════════════════════════════════════════════════════════════
// CSV IMPORT / EXPORT
// ══════════════════════════════════════════════════════════════════════════════

window.importCSV = async function() {
  await waitForApi();
  const res = await pywebview.api.import_csv();
  if (res.ok) {
    showToast('Imported ' + res.added + ' contacts (' + res.skipped + ' skipped).', 'success');
    loadContacts();
    loadFamilies();
    loadGroups();
    refreshAcCache();
  } else {
    showToast(res.error || 'Import failed.', 'error');
  }
};

window.exportCSV = async function() {
  await waitForApi();
  const res = await pywebview.api.export_csv();
  if (res.ok) {
    showToast('Exported ' + res.count + ' contacts.', 'success');
  } else {
    showToast(res.error || 'Export failed.', 'error');
  }
};

// ══════════════════════════════════════════════════════════════════════════════
// EMAIL HISTORY
// ══════════════════════════════════════════════════════════════════════════════

async function loadHistory() {
  await waitForApi();
  const history = await pywebview.api.get_email_history();
  const el = document.getElementById('history-list');
  if (!history.length) { renderEmpty(el, 'No email history yet.'); return; }
  el.innerHTML = history.map(h => {
    const sentColor = h.failed > 0 ? 'var(--accent)' : 'var(--success)';
    return renderCard(
      esc(h.subject),
      '<span style="font-size:11px;color:' + sentColor + '">' + h.sent + '/' + h.recipients + ' sent</span>',
      'Target: ' + esc(h.target || 'all') + ' | ' + esc(h.sent_at) +
        (h.failed > 0 ? ' | <span style="color:var(--danger)">' + h.failed + ' failed</span>' : '')
    );
  }).join('');
}

// ── Initial load (single wait, then sequential to avoid race conditions) ──
// ══════════════════════════════════════════════════════════════════════════════
// PANEL DIVIDER DRAG
// ══════════════════════════════════════════════════════════════════════════════

(function() {
  var divider = document.getElementById('panel-divider');
  var leftPanel = document.getElementById('left-panel');
  var rightPanel = document.getElementById('right-panel');
  var app = leftPanel.parentElement;
  var dragging = false;

  // Full-screen transparent overlay to block ALL elements (including Quill)
  // from stealing mouse events during drag
  var overlay = document.createElement('div');
  overlay.style.cssText = 'position:fixed;top:0;left:0;width:100vw;height:100vh;z-index:99999;cursor:col-resize;display:none;';
  document.body.appendChild(overlay);

  function applyRatio(r) {
    leftPanel.style.flex = '0 0 calc(' + (r * 100) + '% - 6px)';
    rightPanel.style.flex = '1 1 0%';
  }

  // Ratio restored in initApp after API is ready

  divider.addEventListener('mousedown', function(e) {
    e.preventDefault();
    dragging = true;
    overlay.style.display = 'block';
    divider.classList.add('dragging');
  });

  overlay.addEventListener('mousemove', function(e) {
    if (!dragging) return;
    var appRect = app.getBoundingClientRect();
    var x = e.clientX - appRect.left;
    var w = appRect.width;
    var ratio = x / w;
    if (ratio < 0.2) ratio = 0.2;
    if (ratio > 0.8) ratio = 0.8;
    applyRatio(ratio);
    pywebview.api.set_ui_setting('panelRatio', ratio.toFixed(4));
  });

  overlay.addEventListener('mouseup', function(e) {
    dragging = false;
    overlay.style.display = 'none';
    divider.classList.remove('dragging');
  });
})();

// ── Initial load (single wait, then sequential to avoid race conditions) ──
(async function initApp() {
  await waitForApi();
  await loadContacts();
  await loadFamilies();
  await loadGroups();
  await loadScheduled();
  await loadHistory();
  await loadTemplates();
  await loadSettings();
  await refreshAcCache();

  // Restore UI settings (theme + panel ratio)
  var theme = await pywebview.api.get_ui_setting('theme');
  if (theme === 'light') {
    document.body.classList.add('light');
    document.getElementById('theme-toggle').textContent = 'Dark Mode';
  }
  var ratio = await pywebview.api.get_ui_setting('panelRatio');
  if (ratio) {
    var r = parseFloat(ratio);
    if (r >= 0.2 && r <= 0.8) {
      document.getElementById('left-panel').style.flex = '0 0 calc(' + (r * 100) + '% - 6px)';
      document.getElementById('right-panel').style.flex = '1 1 0%';
    }
  }
})();
</script>
</body>
</html>"""


# ── App bootstrap ─────────────────────────────────────────────────────────────

if __name__ == "__main__":
    db_manager.init_db()
    api = Api()

    # Start background scheduler for scheduled emails
    scheduler_thread = threading.Thread(target=run_scheduler, args=(api,), daemon=True, name="EmailScheduler")
    scheduler_thread.start()

    window = webview.create_window(
        "Church Roster & Email Dispatcher",
        html=HTML,
        js_api=api,
        width=1150,
        height=750,
        min_size=(1000, 650),
    )
    webview.start()
