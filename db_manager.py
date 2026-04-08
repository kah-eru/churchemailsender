import json
import sqlite3
import os
from datetime import datetime, timedelta

DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "contacts.db")


def get_connection():
    conn = sqlite3.connect(DB_PATH)
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


def init_db():
    conn = get_connection()
    try:
        cursor = conn.cursor()

        cursor.execute("""
            CREATE TABLE IF NOT EXISTS families (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL UNIQUE
            )
        """)

        cursor.execute("""
            CREATE TABLE IF NOT EXISTS roster (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL,
                email TEXT NOT NULL,
                category TEXT NOT NULL CHECK(category IN ('Family', 'Single')),
                family_id INTEGER,
                FOREIGN KEY (family_id) REFERENCES families(id) ON DELETE SET NULL
            )
        """)

        cursor.execute("""
            CREATE TABLE IF NOT EXISTS settings (
                key TEXT PRIMARY KEY,
                value TEXT NOT NULL
            )
        """)

        cursor.execute("""
            CREATE TABLE IF NOT EXISTS groups_ (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL UNIQUE
            )
        """)

        cursor.execute("""
            CREATE TABLE IF NOT EXISTS group_members (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                group_id INTEGER NOT NULL,
                contact_id INTEGER NOT NULL,
                FOREIGN KEY (group_id) REFERENCES groups_(id) ON DELETE CASCADE,
                FOREIGN KEY (contact_id) REFERENCES roster(id) ON DELETE CASCADE,
                UNIQUE (group_id, contact_id)
            )
        """)

        cursor.execute("""
            CREATE TABLE IF NOT EXISTS family_members (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                family_id INTEGER NOT NULL,
                contact_id INTEGER NOT NULL,
                FOREIGN KEY (family_id) REFERENCES families(id) ON DELETE CASCADE,
                FOREIGN KEY (contact_id) REFERENCES roster(id) ON DELETE CASCADE,
                UNIQUE (family_id, contact_id)
            )
        """)

        # Migration: copy existing roster.family_id into family_members
        cursor.execute("""
            INSERT OR IGNORE INTO family_members (family_id, contact_id)
            SELECT family_id, id FROM roster WHERE family_id IS NOT NULL
        """)

        cursor.execute("""
            CREATE TABLE IF NOT EXISTS scheduled_emails (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                subject TEXT NOT NULL,
                html_body TEXT NOT NULL,
                plain_text TEXT NOT NULL,
                target_type TEXT NOT NULL,
                target_id INTEGER,
                contact_ids TEXT,
                attachment_paths TEXT,
                scheduled_at TEXT NOT NULL,
                status TEXT NOT NULL DEFAULT 'pending',
                created_at TEXT NOT NULL,
                sent_at TEXT,
                result TEXT
            )
        """)

        cursor.execute("""
            CREATE TABLE IF NOT EXISTS email_templates (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL UNIQUE,
                subject TEXT NOT NULL,
                html_body TEXT NOT NULL
            )
        """)

        cursor.execute("""
            CREATE TABLE IF NOT EXISTS email_history (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                subject TEXT NOT NULL,
                target_description TEXT,
                recipient_count INTEGER NOT NULL DEFAULT 0,
                sent_count INTEGER NOT NULL DEFAULT 0,
                failed_count INTEGER NOT NULL DEFAULT 0,
                sent_at TEXT NOT NULL
            )
        """)

        cursor.execute("""
            CREATE TABLE IF NOT EXISTS email_history_details (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                history_id INTEGER NOT NULL,
                recipient_name TEXT,
                recipient_email TEXT NOT NULL,
                status TEXT NOT NULL DEFAULT 'sent',
                error_message TEXT,
                FOREIGN KEY (history_id) REFERENCES email_history(id) ON DELETE CASCADE
            )
        """)

        # Indexes for frequently queried columns
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_ehd_history_id ON email_history_details(history_id)")
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_roster_family_id ON roster(family_id)")
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_scheduled_status_at ON scheduled_emails(status, scheduled_at)")

        # Migrations — add columns to existing tables
        migrations = [
            ("scheduled_emails", "recurrence", "TEXT"),
            ("scheduled_emails", "manual_emails", "TEXT"),
            ("email_templates", "recipients", "TEXT"),
            ("roster", "phone", "TEXT DEFAULT ''"),
            ("roster", "notes", "TEXT DEFAULT ''"),
            ("roster", "opt_out", "INTEGER DEFAULT 0"),
            ("roster", "created_at", "TEXT"),
            ("roster", "last_emailed_at", "TEXT"),
            ("roster", "email_count", "INTEGER DEFAULT 0"),
        ]
        for table, col, col_type in migrations:
            try:
                cursor.execute(f"ALTER TABLE {table} ADD COLUMN {col} {col_type}")
            except sqlite3.OperationalError:
                pass  # column already exists

        conn.commit()
    finally:
        conn.close()


# ── Settings CRUD ────────────────────────────────────────────────────────────

def get_setting(key):
    conn = get_connection()
    try:
        cursor = conn.cursor()
        cursor.execute("SELECT value FROM settings WHERE key = ?", (key,))
        row = cursor.fetchone()
        return row[0] if row else None
    finally:
        conn.close()


def set_setting(key, value):
    conn = get_connection()
    try:
        cursor = conn.cursor()
        cursor.execute("INSERT OR REPLACE INTO settings (key, value) VALUES (?, ?)", (key, value))
        conn.commit()
    finally:
        conn.close()


# ── Family CRUD ───────────────────────────────────────────────────────────────

def add_family(name):
    conn = get_connection()
    try:
        cursor = conn.cursor()
        cursor.execute("INSERT INTO families (name) VALUES (?)", (name,))
        conn.commit()
    finally:
        conn.close()


def get_families():
    conn = get_connection()
    try:
        cursor = conn.cursor()
        cursor.execute("SELECT id, name FROM families ORDER BY name")
        return cursor.fetchall()
    finally:
        conn.close()


def rename_family(family_id, new_name):
    conn = get_connection()
    try:
        cursor = conn.cursor()
        cursor.execute("UPDATE families SET name = ? WHERE id = ?", (new_name, family_id))
        conn.commit()
    finally:
        conn.close()


def delete_family(family_id):
    conn = get_connection()
    try:
        cursor = conn.cursor()
        cursor.execute("DELETE FROM families WHERE id = ?", (family_id,))
        conn.commit()
    finally:
        conn.close()


# ── Contact CRUD ──────────────────────────────────────────────────────────────

def add_contact(name, email, category, family_id=None, phone="", notes=""):
    conn = get_connection()
    try:
        cursor = conn.cursor()
        cursor.execute(
            "INSERT INTO roster (name, email, category, family_id, phone, notes, created_at) VALUES (?, ?, ?, ?, ?, ?, ?)",
            (name, email, category, family_id, phone, notes, datetime.now().isoformat()),
        )
        contact_id = cursor.lastrowid
        if family_id:
            cursor.execute(
                "INSERT OR IGNORE INTO family_members (family_id, contact_id) VALUES (?, ?)",
                (family_id, contact_id),
            )
        conn.commit()
    finally:
        conn.close()


def get_contacts():
    conn = get_connection()
    try:
        cursor = conn.cursor()
        cursor.execute("""
            SELECT r.id, r.name, r.email, r.category, COALESCE(f.name, ''),
                   COALESCE(r.phone, ''), COALESCE(r.notes, ''), COALESCE(r.opt_out, 0),
                   r.created_at, r.last_emailed_at, COALESCE(r.email_count, 0)
            FROM roster r
            LEFT JOIN families f ON r.family_id = f.id
            ORDER BY r.name
        """)
        return cursor.fetchall()
    finally:
        conn.close()


def get_contacts_by_family(family_id):
    conn = get_connection()
    try:
        cursor = conn.cursor()
        cursor.execute("""
            SELECT r.id, r.name, r.email, r.category, COALESCE(f.name, '')
            FROM family_members fm
            JOIN roster r ON fm.contact_id = r.id
            LEFT JOIN families f ON fm.family_id = f.id
            WHERE fm.family_id = ?
            ORDER BY r.name
        """, (family_id,))
        return cursor.fetchall()
    finally:
        conn.close()


def get_contact_groups():
    """Returns a dict mapping contact_id -> list of group names."""
    conn = get_connection()
    try:
        cursor = conn.cursor()
        cursor.execute("""
            SELECT gm.contact_id, g.name
            FROM group_members gm
            JOIN groups_ g ON gm.group_id = g.id
            ORDER BY g.name
        """)
        result = {}
        for cid, gname in cursor.fetchall():
            result.setdefault(cid, []).append(gname)
        return result
    finally:
        conn.close()


def get_contact_families():
    """Returns a dict mapping contact_id -> list of {id, name}."""
    conn = get_connection()
    try:
        cursor = conn.cursor()
        cursor.execute("""
            SELECT fm.contact_id, f.id, f.name
            FROM family_members fm
            JOIN families f ON fm.family_id = f.id
            ORDER BY f.name
        """)
        result = {}
        for cid, fid, fname in cursor.fetchall():
            result.setdefault(cid, []).append({"id": fid, "name": fname})
        return result
    finally:
        conn.close()


def add_family_member(family_id, contact_id):
    conn = get_connection()
    try:
        cursor = conn.cursor()
        cursor.execute("INSERT OR IGNORE INTO family_members (family_id, contact_id) VALUES (?, ?)", (family_id, contact_id))
        conn.commit()
    finally:
        conn.close()


def remove_family_member(family_id, contact_id):
    conn = get_connection()
    try:
        cursor = conn.cursor()
        cursor.execute("DELETE FROM family_members WHERE family_id = ? AND contact_id = ?", (family_id, contact_id))
        conn.commit()
    finally:
        conn.close()


def get_family_members_via_junction(family_id):
    """Get family members via the family_members junction table."""
    conn = get_connection()
    try:
        cursor = conn.cursor()
        cursor.execute("""
            SELECT r.id, r.name, r.email
            FROM family_members fm
            JOIN roster r ON fm.contact_id = r.id
            WHERE fm.family_id = ?
            ORDER BY r.name
        """, (family_id,))
        return cursor.fetchall()
    finally:
        conn.close()


def get_all_families_with_members():
    """Get all families with their members in a single query (avoids N+1)."""
    conn = get_connection()
    try:
        cursor = conn.cursor()
        cursor.execute("SELECT id, name FROM families ORDER BY name")
        families = cursor.fetchall()
        cursor.execute("""
            SELECT fm.family_id, r.id, r.name, r.email
            FROM family_members fm
            JOIN roster r ON fm.contact_id = r.id
            ORDER BY r.name
        """)
        members_map = {}
        for fid, rid, rname, remail in cursor.fetchall():
            members_map.setdefault(fid, []).append((rid, rname, remail))
        return [(fid, fname, members_map.get(fid, [])) for fid, fname in families]
    finally:
        conn.close()


def update_contact(contact_id, name, email, category, family_id=None, phone="", notes=""):
    conn = get_connection()
    try:
        cursor = conn.cursor()
        # Get old family_id before updating so we can clean up the junction table
        cursor.execute("SELECT family_id FROM roster WHERE id = ?", (contact_id,))
        row = cursor.fetchone()
        old_family_id = row[0] if row else None

        cursor.execute("UPDATE roster SET name = ?, email = ?, category = ?, family_id = ?, phone = ?, notes = ? WHERE id = ?",
                       (name, email, category, family_id, phone, notes, contact_id))
        # Sync family_members junction table: remove old primary family link, add new one
        if old_family_id and old_family_id != family_id:
            cursor.execute(
                "DELETE FROM family_members WHERE family_id = ? AND contact_id = ?",
                (old_family_id, contact_id),
            )
        if family_id:
            cursor.execute(
                "INSERT OR IGNORE INTO family_members (family_id, contact_id) VALUES (?, ?)",
                (family_id, contact_id),
            )
        conn.commit()
    finally:
        conn.close()


def set_contact_opt_out(contact_id, opt_out):
    conn = get_connection()
    try:
        cursor = conn.cursor()
        cursor.execute("UPDATE roster SET opt_out = ? WHERE id = ?", (1 if opt_out else 0, contact_id))
        conn.commit()
    finally:
        conn.close()


def update_contact_email_stats(email_addresses):
    """Increment email_count and set last_emailed_at for contacts matching these email addresses."""
    if not email_addresses:
        return
    conn = get_connection()
    try:
        cursor = conn.cursor()
        now = datetime.now().isoformat()
        placeholders = ",".join(["?"] * len(email_addresses))
        cursor.execute(
            f"UPDATE roster SET email_count = COALESCE(email_count, 0) + 1, last_emailed_at = ? WHERE email IN ({placeholders})",
            [now] + list(email_addresses)
        )
        conn.commit()
    finally:
        conn.close()


def bulk_update_category(contact_ids, category):
    conn = get_connection()
    try:
        cursor = conn.cursor()
        placeholders = ",".join(["?"] * len(contact_ids))
        cursor.execute(f"UPDATE roster SET category = ? WHERE id IN ({placeholders})", [category] + list(contact_ids))
        conn.commit()
    finally:
        conn.close()


def bulk_add_to_group(group_id, contact_ids):
    conn = get_connection()
    try:
        cursor = conn.cursor()
        for cid in contact_ids:
            cursor.execute("INSERT OR IGNORE INTO group_members (group_id, contact_id) VALUES (?, ?)", (group_id, cid))
        conn.commit()
    finally:
        conn.close()


def update_scheduled_email(email_id, subject, html_body, plain_text, target_type, target_id, contact_ids, attachment_paths, scheduled_at, recurrence=None, manual_emails=None):
    conn = get_connection()
    try:
        cursor = conn.cursor()
        cursor.execute("""
            UPDATE scheduled_emails SET subject=?, html_body=?, plain_text=?, target_type=?, target_id=?,
            contact_ids=?, attachment_paths=?, scheduled_at=?, recurrence=?, manual_emails=?
            WHERE id=? AND status='pending'
        """, (subject, html_body, plain_text, target_type, target_id,
              json.dumps(contact_ids) if contact_ids else None,
              json.dumps(attachment_paths) if attachment_paths else None,
              scheduled_at,
              json.dumps(recurrence) if recurrence else None,
              json.dumps(manual_emails) if manual_emails else None,
              email_id))
        conn.commit()
    finally:
        conn.close()


def duplicate_scheduled_email(email_id):
    conn = get_connection()
    try:
        cursor = conn.cursor()
        cursor.execute("SELECT subject, html_body, plain_text, target_type, target_id, contact_ids, attachment_paths, recurrence, manual_emails FROM scheduled_emails WHERE id = ?", (email_id,))
        row = cursor.fetchone()
        if not row:
            return None
        new_scheduled = (datetime.now() + timedelta(days=1)).isoformat()
        cursor.execute("""
            INSERT INTO scheduled_emails (subject, html_body, plain_text, target_type, target_id, contact_ids, attachment_paths, scheduled_at, status, created_at, recurrence, manual_emails)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, 'pending', ?, ?, ?)
        """, (row[0], row[1], row[2], row[3], row[4], row[5], row[6], new_scheduled, datetime.now().isoformat(), row[7], row[8]))
        conn.commit()
        return cursor.lastrowid
    finally:
        conn.close()


def duplicate_template(template_id):
    conn = get_connection()
    try:
        cursor = conn.cursor()
        cursor.execute("SELECT name, subject, html_body, recipients FROM email_templates WHERE id = ?", (template_id,))
        row = cursor.fetchone()
        if not row:
            return None
        new_name = row[0] + " (Copy)"
        cursor.execute("INSERT INTO email_templates (name, subject, html_body, recipients) VALUES (?, ?, ?, ?)",
                       (new_name, row[1], row[2], row[3]))
        conn.commit()
        return cursor.lastrowid
    finally:
        conn.close()


def get_email_history_filtered(start_date=None, end_date=None):
    conn = get_connection()
    try:
        cursor = conn.cursor()
        query = "SELECT id, subject, target_description, recipient_count, sent_count, failed_count, sent_at FROM email_history"
        params = []
        conditions = []
        if start_date:
            conditions.append("sent_at >= ?")
            params.append(start_date)
        if end_date:
            conditions.append("sent_at <= ?")
            params.append(end_date)
        if conditions:
            query += " WHERE " + " AND ".join(conditions)
        query += " ORDER BY sent_at DESC"
        cursor.execute(query, params)
        return cursor.fetchall()
    finally:
        conn.close()


def get_analytics():
    conn = get_connection()
    try:
        cursor = conn.cursor()
        # Emails per week (last 12 weeks)
        cursor.execute("""
            SELECT strftime('%Y-W%W', sent_at) as week, COUNT(*) as count,
                   SUM(sent_count) as sent, SUM(failed_count) as failed
            FROM email_history
            WHERE sent_at >= date('now', '-84 days')
            GROUP BY week ORDER BY week
        """)
        weekly = [{"week": r[0], "count": r[1], "sent": r[2], "failed": r[3]} for r in cursor.fetchall()]

        # Most failed recipients
        cursor.execute("""
            SELECT recipient_email, recipient_name, COUNT(*) as fail_count
            FROM email_history_details WHERE status = 'failed'
            GROUP BY recipient_email ORDER BY fail_count DESC LIMIT 10
        """)
        top_failed = [{"email": r[0], "name": r[1], "count": r[2]} for r in cursor.fetchall()]

        # Overall stats
        cursor.execute("SELECT COUNT(*), SUM(sent_count), SUM(failed_count) FROM email_history")
        r = cursor.fetchone()
        totals = {"emails_sent": r[0] or 0, "recipients_reached": r[1] or 0, "total_failures": r[2] or 0}

        # Failure rate
        if totals["recipients_reached"] + totals["total_failures"] > 0:
            totals["failure_rate"] = round(totals["total_failures"] / (totals["recipients_reached"] + totals["total_failures"]) * 100, 1)
        else:
            totals["failure_rate"] = 0

        return {"weekly": weekly, "top_failed": top_failed, "totals": totals}
    finally:
        conn.close()


def backup_database(backup_path):
    """Create a backup copy of the database."""
    import shutil
    shutil.copy2(DB_PATH, backup_path)


def restore_database(backup_path):
    """Restore database from a backup file."""
    import shutil
    shutil.copy2(backup_path, DB_PATH)


def delete_contact(contact_id):
    conn = get_connection()
    try:
        cursor = conn.cursor()
        cursor.execute("DELETE FROM roster WHERE id = ?", (contact_id,))
        conn.commit()
    finally:
        conn.close()


# ── Group CRUD ───────────────────────────────────────────────────────────────

def add_group(name):
    conn = get_connection()
    try:
        cursor = conn.cursor()
        cursor.execute("INSERT INTO groups_ (name) VALUES (?)", (name,))
        conn.commit()
    finally:
        conn.close()


def get_groups():
    conn = get_connection()
    try:
        cursor = conn.cursor()
        cursor.execute("SELECT id, name FROM groups_ ORDER BY name")
        return cursor.fetchall()
    finally:
        conn.close()


def rename_group(group_id, new_name):
    conn = get_connection()
    try:
        cursor = conn.cursor()
        cursor.execute("UPDATE groups_ SET name = ? WHERE id = ?", (new_name, group_id))
        conn.commit()
    finally:
        conn.close()


def delete_group(group_id):
    conn = get_connection()
    try:
        cursor = conn.cursor()
        cursor.execute("DELETE FROM groups_ WHERE id = ?", (group_id,))
        conn.commit()
    finally:
        conn.close()


def add_group_member(group_id, contact_id):
    conn = get_connection()
    try:
        cursor = conn.cursor()
        cursor.execute("INSERT OR IGNORE INTO group_members (group_id, contact_id) VALUES (?, ?)", (group_id, contact_id))
        conn.commit()
    finally:
        conn.close()


def remove_group_member(group_id, contact_id):
    conn = get_connection()
    try:
        cursor = conn.cursor()
        cursor.execute("DELETE FROM group_members WHERE group_id = ? AND contact_id = ?", (group_id, contact_id))
        conn.commit()
    finally:
        conn.close()


def get_all_groups_with_members():
    """Get all groups with their members in a single query (avoids N+1)."""
    conn = get_connection()
    try:
        cursor = conn.cursor()
        cursor.execute("SELECT id, name FROM groups_ ORDER BY name")
        groups = cursor.fetchall()
        cursor.execute("""
            SELECT gm.group_id, r.id, r.name, r.email
            FROM group_members gm
            JOIN roster r ON gm.contact_id = r.id
            ORDER BY r.name
        """)
        members_map = {}
        for gid, rid, rname, remail in cursor.fetchall():
            members_map.setdefault(gid, []).append((rid, rname, remail))
        return [(gid, gname, members_map.get(gid, [])) for gid, gname in groups]
    finally:
        conn.close()


def get_group_members(group_id):
    conn = get_connection()
    try:
        cursor = conn.cursor()
        cursor.execute("""
            SELECT r.id, r.name, r.email, r.category
            FROM roster r
            JOIN group_members gm ON r.id = gm.contact_id
            WHERE gm.group_id = ?
            ORDER BY r.name
        """, (group_id,))
        return cursor.fetchall()
    finally:
        conn.close()


# ── Scheduled Emails CRUD ────────────────────────────────────────────────────

def schedule_email(subject, html_body, plain_text, target_type, target_id, contact_ids, attachment_paths, scheduled_at, recurrence=None, manual_emails=None):
    conn = get_connection()
    try:
        cursor = conn.cursor()
        cursor.execute("""
            INSERT INTO scheduled_emails (subject, html_body, plain_text, target_type, target_id, contact_ids, attachment_paths, scheduled_at, status, created_at, recurrence, manual_emails)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, 'pending', ?, ?, ?)
        """, (subject, html_body, plain_text, target_type, target_id,
              json.dumps(contact_ids) if contact_ids else None,
              json.dumps(attachment_paths) if attachment_paths else None,
              scheduled_at, datetime.now().isoformat(),
              json.dumps(recurrence) if recurrence else None,
              json.dumps(manual_emails) if manual_emails else None))
        conn.commit()
    finally:
        conn.close()


def get_scheduled_emails():
    conn = get_connection()
    try:
        cursor = conn.cursor()
        cursor.execute("SELECT id, subject, target_type, target_id, scheduled_at, status, sent_at, result, recurrence, contact_ids, manual_emails FROM scheduled_emails ORDER BY scheduled_at ASC")
        return cursor.fetchall()
    finally:
        conn.close()


def get_scheduled_email_by_id(email_id):
    conn = get_connection()
    try:
        cursor = conn.cursor()
        cursor.execute("SELECT id, subject, html_body, plain_text, target_type, target_id, contact_ids, attachment_paths, scheduled_at, status, sent_at, result, recurrence, manual_emails FROM scheduled_emails WHERE id = ?", (email_id,))
        return cursor.fetchone()
    finally:
        conn.close()


def cancel_scheduled_email(email_id):
    conn = get_connection()
    try:
        cursor = conn.cursor()
        cursor.execute("UPDATE scheduled_emails SET status = 'cancelled' WHERE id = ? AND status = 'pending'", (email_id,))
        conn.commit()
    finally:
        conn.close()


def get_due_emails(now_iso=None):
    conn = get_connection()
    try:
        cursor = conn.cursor()
        now = now_iso or datetime.now().isoformat()
        cursor.execute("SELECT id, subject, html_body, plain_text, target_type, target_id, contact_ids, attachment_paths, recurrence, manual_emails, scheduled_at FROM scheduled_emails WHERE status = 'pending' AND scheduled_at <= ?", (now,))
        return cursor.fetchall()
    finally:
        conn.close()


def update_email_status(email_id, status, result=None):
    conn = get_connection()
    try:
        cursor = conn.cursor()
        sent_at = datetime.now().isoformat() if status == 'sent' else None
        cursor.execute("UPDATE scheduled_emails SET status = ?, result = ?, sent_at = COALESCE(?, sent_at) WHERE id = ?",
                       (status, json.dumps(result) if result else None, sent_at, email_id))
        conn.commit()
    finally:
        conn.close()


# ── Email Templates CRUD ─────────────────────────────────────────────────────

def save_template(name, subject, html_body, recipients=None):
    conn = get_connection()
    try:
        cursor = conn.cursor()
        cursor.execute("INSERT OR REPLACE INTO email_templates (name, subject, html_body, recipients) VALUES (?, ?, ?, ?)",
                       (name, subject, html_body, json.dumps(recipients) if recipients else None))
        conn.commit()
    finally:
        conn.close()


def update_template(template_id, subject, html_body, recipients=None):
    conn = get_connection()
    try:
        cursor = conn.cursor()
        cursor.execute("UPDATE email_templates SET subject = ?, html_body = ?, recipients = ? WHERE id = ?",
                       (subject, html_body, json.dumps(recipients) if recipients else None, template_id))
        conn.commit()
    finally:
        conn.close()


def get_templates():
    conn = get_connection()
    try:
        cursor = conn.cursor()
        cursor.execute("SELECT id, name, subject, html_body, recipients FROM email_templates ORDER BY name")
        return cursor.fetchall()
    finally:
        conn.close()


def delete_template(template_id):
    conn = get_connection()
    try:
        cursor = conn.cursor()
        cursor.execute("DELETE FROM email_templates WHERE id = ?", (template_id,))
        conn.commit()
    finally:
        conn.close()


# ── Email History CRUD ───────────────────────────────────────────────────────

def log_email(subject, target_description, recipient_count, sent_count, failed_count, recipient_details=None):
    conn = get_connection()
    try:
        cursor = conn.cursor()
        cursor.execute("""
            INSERT INTO email_history (subject, target_description, recipient_count, sent_count, failed_count, sent_at)
            VALUES (?, ?, ?, ?, ?, ?)
        """, (subject, target_description, recipient_count, sent_count, failed_count, datetime.now().isoformat()))
        history_id = cursor.lastrowid
        if recipient_details:
            for rd in recipient_details:
                cursor.execute("""
                    INSERT INTO email_history_details (history_id, recipient_name, recipient_email, status, error_message)
                    VALUES (?, ?, ?, ?, ?)
                """, (history_id, rd.get("name", ""), rd["email"], rd["status"], rd.get("error")))
        conn.commit()
    finally:
        conn.close()


def get_email_history():
    conn = get_connection()
    try:
        cursor = conn.cursor()
        cursor.execute("SELECT id, subject, target_description, recipient_count, sent_count, failed_count, sent_at FROM email_history ORDER BY sent_at DESC")
        return cursor.fetchall()
    finally:
        conn.close()


def get_email_history_details(history_id):
    conn = get_connection()
    try:
        cursor = conn.cursor()
        cursor.execute("SELECT recipient_name, recipient_email, status, error_message FROM email_history_details WHERE history_id = ? ORDER BY status DESC, recipient_name", (history_id,))
        return cursor.fetchall()
    finally:
        conn.close()


# ── Recurrence helpers ──────────────────────────────────────────────────────

def compute_next_occurrence(current_dt, recurrence):
    """Given the current scheduled datetime and a recurrence dict, return the next datetime or None."""
    rtype = recurrence.get("type", "once")
    if rtype == "once":
        return None

    end_date_str = recurrence.get("end_date")
    if end_date_str:
        end_date = datetime.fromisoformat(end_date_str)
    else:
        end_date = None

    if rtype == "daily":
        nxt = current_dt + timedelta(days=1)
    elif rtype == "every_other_day":
        nxt = current_dt + timedelta(days=2)
    elif rtype in ("weekly", "every_other_week"):
        days = recurrence.get("days", [])
        if not days:
            return None
        week_offset = 1 if rtype == "weekly" else 2
        # days are 0=Sun..6=Sat in JS convention → convert to Python (0=Mon..6=Sun)
        py_days = sorted([(d - 1) % 7 for d in days])
        current_weekday = current_dt.weekday()
        # find next day in the same week (or the following week(s))
        nxt = None
        for d in py_days:
            diff = d - current_weekday
            if diff > 0:
                candidate = current_dt + timedelta(days=diff)
                nxt = candidate
                break
        if nxt is None:
            # wrap to first day of next cycle
            days_until = (py_days[0] - current_weekday) % 7
            if days_until == 0:
                days_until = 7
            nxt = current_dt + timedelta(days=days_until + 7 * (week_offset - 1))
    elif rtype == "monthly":
        day_of_month = recurrence.get("day_of_month", current_dt.day)
        month = current_dt.month + 1
        year = current_dt.year
        if month > 12:
            month = 1
            year += 1
        # clamp day to valid range for target month
        import calendar
        max_day = calendar.monthrange(year, month)[1]
        day = min(day_of_month, max_day)
        nxt = current_dt.replace(year=year, month=month, day=day)
    else:
        return None

    if end_date and nxt > end_date:
        return None
    return nxt
