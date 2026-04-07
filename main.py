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
        return [
            {"id": r[0], "name": r[1], "email": r[2], "category": r[3], "family_name": r[4]}
            for r in rows
        ]

    def add_contact(self, name, email, category, family_id):
        try:
            fid = int(family_id) if family_id else None
            db_manager.add_contact(name, email, category, fid)
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
            members = db_manager.get_contacts_by_family(fid)
            result.append({
                "id": fid,
                "name": fname,
                "members": [{"id": m[0], "name": m[1], "email": m[2]} for m in members],
            })
        return result

    def add_family(self, name):
        try:
            db_manager.add_family(name)
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
  .app { display: flex; height: 100vh; gap: 8px; padding: 10px; }

  /* ── Panels ── */
  .left-panel, .right-panel {
    background: var(--panel); border-radius: 10px; padding: 16px; display: flex; flex-direction: column;
    transition: background 0.2s;
  }
  .left-panel { flex: 1; min-width: 420px; }
  .right-panel { flex: 1; min-width: 420px; overflow: hidden; }

  /* ── Tabs ── */
  .tab-bar { display: flex; gap: 4px; margin-bottom: 12px; }
  .tab-btn {
    flex: 1; padding: 8px 0; border: none; border-radius: 6px; cursor: pointer;
    background: var(--tab-bg); color: var(--tab-text); font-size: 13px; font-weight: 600; transition: all 0.2s;
  }
  .tab-btn.active { background: var(--accent); color: #fff; }
  .tab-content { display: none; flex-direction: column; flex: 1; overflow: hidden; }
  .tab-content.active { display: flex; }

  /* ── Scrollable lists ── */
  .list-area { flex: 1; overflow-y: auto; margin-bottom: 10px; border-radius: 6px; background: var(--surface); padding: 6px; }
  .list-area::-webkit-scrollbar { width: 6px; }
  .list-area::-webkit-scrollbar-thumb { background: var(--scrollbar); border-radius: 3px; }

  .contact-row, .family-card {
    display: flex; align-items: center; gap: 8px; padding: 7px 10px; border-radius: 5px; font-size: 13px;
  }
  .contact-row:hover { background: var(--row-hover); }
  .contact-row input[type="checkbox"] { accent-color: var(--accent); width: 15px; height: 15px; cursor: pointer; }
  .contact-row .name { flex: 1; font-weight: 500; }
  .contact-row .email { flex: 1; color: var(--text-muted); }
  .contact-row .cat { width: 55px; font-size: 11px; color: var(--text-muted); }
  .contact-row .fam { width: 80px; font-size: 11px; color: var(--text-dim); overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }

  .family-card {
    flex-direction: column; align-items: flex-start; background: var(--card-bg); margin-bottom: 6px; padding: 10px 12px;
  }
  .family-card .fam-header { display: flex; width: 100%; justify-content: space-between; align-items: center; }
  .family-card .fam-name { font-weight: 600; font-size: 14px; }
  .family-card .fam-members { font-size: 12px; color: var(--text-muted); margin-top: 4px; }

  .empty-msg { text-align: center; color: var(--text-dim); padding: 30px 0; font-size: 13px; }

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
  .tab-search {
    width: 100%; padding: 6px 10px; margin-bottom: 8px; border: 1px solid var(--border);
    border-radius: 6px; background: var(--surface); color: var(--text); font-size: 12px;
    outline: none; transition: border-color 0.2s;
  }
  .tab-search:focus { border-color: var(--accent); }
  .tab-search::placeholder { color: var(--text-dim); }

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
  <div class="left-panel">
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
      <input type="text" class="tab-search" id="search-contacts" placeholder="Search contacts..." oninput="filterContacts()">
      <div class="list-area" id="contact-list"></div>
      <div class="form-row">
        <input type="text" id="c-name" placeholder="Name">
        <input type="text" id="c-email" placeholder="Email">
      </div>
      <div class="form-row">
        <select id="c-category" onchange="onCategoryChange()">
          <option value="Single">Single</option>
          <option value="Family">Family</option>
        </select>
        <select id="c-family" disabled><option value="">No family</option></select>
      </div>
      <div class="btn-row">
        <button class="btn btn-primary" onclick="addContact()">Add Contact</button>
        <button class="btn btn-danger" onclick="deleteSelected()">Delete Selected</button>
        <button class="btn btn-primary btn-success" onclick="importCSV()">Import CSV</button>
        <button class="btn btn-primary" onclick="exportCSV()">Export CSV</button>
      </div>
    </div>

    <!-- Families Tab -->
    <div id="families-tab" class="tab-content">
      <input type="text" class="tab-search" id="search-families" placeholder="Search families..." oninput="filterFamilies()">
      <div class="list-area" id="family-list"></div>
      <div class="form-row">
        <input type="text" id="f-name" placeholder="New Family Name">
        <button class="btn btn-primary" onclick="createFamily()">Create Family</button>
      </div>
    </div>

    <!-- Groups Tab -->
    <div id="groups-tab" class="tab-content">
      <input type="text" class="tab-search" id="search-groups" placeholder="Search groups..." oninput="filterGroups()">
      <div class="list-area" id="group-list"></div>
      <div class="form-row">
        <input type="text" id="g-name" placeholder="New Group Name">
        <button class="btn btn-primary" onclick="createGroup()">Create Group</button>
      </div>
      <div style="margin-top:8px;">
        <div class="section-label">Add member to group:</div>
        <div class="form-row">
          <select id="g-select"><option value="">Select group</option></select>
          <select id="g-contact"><option value="">Select contact</option></select>
          <button class="btn btn-primary btn-sm" onclick="addMemberToGroup()">Add</button>
        </div>
      </div>
    </div>

    <!-- Scheduled Tab -->
    <div id="scheduled-tab" class="tab-content">
      <div class="list-area" id="scheduled-list"></div>
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

  <!-- ════ RIGHT PANEL ════ -->
  <div class="right-panel">
    <div class="composer-title">Email Composer</div>
    <div class="flex-row" style="margin-bottom:8px;">
      <select id="template-select" class="composer-select" onchange="loadTemplate()">
        <option value="">Load template...</option>
      </select>
      <button class="btn btn-primary btn-sm" onclick="saveAsTemplate()">Save Template</button>
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
        <button class="btn btn-primary btn-nowrap" onclick="scheduleEmail()">Save &amp; Schedule</button>
      </div>
    </div>
    <button class="btn-dispatch" onclick="dispatchEmails()">Send Now</button>
  </div>
</div>

<button class="theme-toggle" id="theme-toggle" onclick="toggleTheme()">Light Mode</button>
<div class="toast" id="toast"></div>

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
  try { localStorage.setItem('theme', isLight ? 'light' : 'dark'); } catch(e) {}
};
// Restore saved theme
try {
  if (localStorage.getItem('theme') === 'light') {
    document.body.classList.add('light');
    document.getElementById('theme-toggle').textContent = 'Dark Mode';
  }
} catch(e) {}

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
  el.innerHTML = contacts.map(c => `
    <div class="contact-row">
      <input type="checkbox" data-id="${c.id}">
      <span class="name">${esc(c.name)}</span>
      <span class="email">${esc(c.email)}</span>
      <span class="cat">${esc(c.category)}</span>
      <span class="fam">${c.family_name ? esc(c.family_name) : '-'}</span>
    </div>
  `).join('');
}

window.filterContacts = function() {
  const q = document.getElementById('search-contacts').value.toLowerCase().trim();
  if (!q) { renderContactList(cachedContacts); return; }
  const filtered = cachedContacts.filter(c =>
    c.name.toLowerCase().includes(q) || c.email.toLowerCase().includes(q) ||
    c.category.toLowerCase().includes(q) || (c.family_name && c.family_name.toLowerCase().includes(q))
  );
  renderContactList(filtered);
};

window.addContact = async function() {
  const name = document.getElementById('c-name').value.trim();
  const email = document.getElementById('c-email').value.trim();
  const category = document.getElementById('c-category').value;
  const familySel = document.getElementById('c-family');
  const familyId = familySel.value || null;

  if (!name || !email) { showToast('Enter both name and email.', 'error'); return; }

  const res = await pywebview.api.add_contact(name, email, category, familyId);
  if (!res.ok) { showToast(res.error, 'error'); return; }

  document.getElementById('c-name').value = '';
  document.getElementById('c-email').value = '';
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

// ══════════════════════════════════════════════════════════════════════════════
// FAMILIES
// ══════════════════════════════════════════════════════════════════════════════

var cachedFamilies = [];
async function loadFamilies() {
  await waitForApi();
  cachedFamilies = await pywebview.api.get_families();

  // Update family dropdown in contacts form
  const sel = document.getElementById('c-family');
  sel.innerHTML = '<option value="">No family</option>' +
    cachedFamilies.map(f => `<option value="${f.id}">${esc(f.name)}</option>`).join('');

  renderFamilyList(cachedFamilies);
}

function renderFamilyList(families) {
  const el = document.getElementById('family-list');
  if (!families.length) { renderEmpty(el, 'No families yet.'); return; }
  el.innerHTML = families.map(f => {
    const members = f.members.length ? f.members.map(m => esc(m.name)).join(', ') : 'No members';
    return `
      <div class="family-card">
        <div class="fam-header">
          <span class="fam-name">${esc(f.name)}</span>
          <button class="btn btn-danger btn-sm" onclick="deleteFamily(${f.id})">Delete</button>
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

window.createFamily = async function() {
  const name = document.getElementById('f-name').value.trim();
  if (!name) { showToast('Enter a family name.', 'error'); return; }

  const res = await pywebview.api.add_family(name);
  if (!res.ok) { showToast(res.error || 'Family already exists.', 'error'); return; }

  document.getElementById('f-name').value = '';
  loadFamilies();
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

window.onCategoryChange = function() {
  const cat = document.getElementById('c-category').value;
  document.getElementById('c-family').disabled = (cat !== 'Family');
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
  const contacts = await pywebview.api.get_contacts();

  renderGroupList(cachedGroups);

  // Update group select dropdown
  const gsel = document.getElementById('g-select');
  gsel.innerHTML = '<option value="">Select group</option>' +
    cachedGroups.map(g => '<option value="' + g.id + '">' + esc(g.name) + '</option>').join('');

  // Update contact select dropdown
  const csel = document.getElementById('g-contact');
  csel.innerHTML = '<option value="">Select contact</option>' +
    contacts.map(c => '<option value="' + c.id + '">' + esc(c.name) + '</option>').join('');

  // Update target selector with group options
  const tsel = document.getElementById('target-select');
  const base = '<option value="">-- Quick select --</option><option value="all">All Contacts</option><option value="family">All Families</option><option value="single">All Singles</option>';
  const groupOpts = cachedGroups.map(g => '<option value="group:' + g.id + '">Group: ' + esc(g.name) + '</option>').join('');
  tsel.innerHTML = base + groupOpts;
}

function renderGroupList(groups) {
  const el = document.getElementById('group-list');
  if (!groups.length) {
    renderEmpty(el, 'No groups yet.');
  } else {
    el.innerHTML = groups.map(g => {
      const members = g.members.length
        ? g.members.map(m => '<span style="display:inline-flex;align-items:center;gap:2px;">' + esc(m.name) +
            ' <span style="cursor:pointer;color:var(--danger);font-size:10px;" onclick="removeMember(' + g.id + ',' + m.id + ')">&times;</span></span>').join(', ')
        : 'No members';
      return renderCard(
        esc(g.name),
        '<button class="btn btn-danger btn-sm" onclick="deleteGroup(' + g.id + ')">Delete</button>',
        members
      );
    }).join('');
  }
}

window.filterGroups = function() {
  const q = document.getElementById('search-groups').value.toLowerCase().trim();
  if (!q) { renderGroupList(cachedGroups); return; }
  const filtered = cachedGroups.filter(g =>
    g.name.toLowerCase().includes(q) ||
    g.members.some(m => m.name.toLowerCase().includes(q))
  );
  renderGroupList(filtered);
};

window.createGroup = async function() {
  const name = document.getElementById('g-name').value.trim();
  if (!name) { showToast('Enter a group name.', 'error'); return; }
  const res = await pywebview.api.add_group(name);
  if (!res.ok) { showToast(res.error || 'Group already exists.', 'error'); return; }
  document.getElementById('g-name').value = '';
  loadGroups();
  refreshAcCache();
};

window.deleteGroup = async function(id) {
  if (!confirm('Delete this group?')) return;
  await pywebview.api.delete_group(id);
  loadGroups();
  refreshAcCache();
};

window.addMemberToGroup = async function() {
  const gid = document.getElementById('g-select').value;
  const cid = document.getElementById('g-contact').value;
  if (!gid || !cid) { showToast('Select both a group and a contact.', 'error'); return; }
  const res = await pywebview.api.add_group_member(parseInt(gid), parseInt(cid));
  if (!res.ok) { showToast(res.error, 'error'); return; }
  loadGroups();
};

window.removeMember = async function(gid, cid) {
  await pywebview.api.remove_group_member(gid, cid);
  loadGroups();
};

// ══════════════════════════════════════════════════════════════════════════════
// SCHEDULED EMAILS
// ══════════════════════════════════════════════════════════════════════════════

async function loadScheduled() {
  await waitForApi();
  const emails = await pywebview.api.get_scheduled_emails();
  const el = document.getElementById('scheduled-list');
  if (!emails.length) { renderEmpty(el, 'No scheduled emails.'); return; }
  el.innerHTML = emails.map(e => {
    const statusColor = e.status === 'sent' ? 'var(--success)' : e.status === 'pending' ? 'var(--accent)' : 'var(--danger)';
    const cancelBtn = e.status === 'pending'
      ? ' <button class="btn btn-danger btn-sm" onclick="cancelScheduled(' + e.id + ')">Cancel</button>'
      : '';
    let recLabel = '';
    if (e.recurrence && e.recurrence.type !== 'once') {
      const dayNames = ['Su','Mo','Tu','We','Th','Fr','Sa'];
      let detail = e.recurrence.type.replace(/_/g, ' ');
      if (e.recurrence.days && e.recurrence.days.length) {
        detail += ' (' + e.recurrence.days.map(d => dayNames[d]).join(', ') + ')';
      }
      if (e.recurrence.day_of_month) detail += ' (day ' + e.recurrence.day_of_month + ')';
      if (e.recurrence.end_date) detail += ' until ' + e.recurrence.end_date.split('T')[0];
      recLabel = ' | Repeat: ' + detail;
    }
    return renderCard(
      esc(e.subject),
      '<span style="font-size:11px;color:' + statusColor + '">' + e.status.toUpperCase() + '</span>',
      'Target: ' + esc(e.target_type) + ' | Scheduled: ' + esc(e.scheduled_at) +
        recLabel + (e.sent_at ? ' | Sent: ' + esc(e.sent_at) : '') + cancelBtn
    );
  }).join('');
}

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
