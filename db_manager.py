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

        # Indexes for frequently queried columns
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_roster_family_id ON roster(family_id)")
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_scheduled_status_at ON scheduled_emails(status, scheduled_at)")

        # Migrations — add columns to existing tables
        for table, col in [("scheduled_emails", "recurrence"), ("scheduled_emails", "manual_emails"), ("email_templates", "recipients")]:
            try:
                cursor.execute(f"ALTER TABLE {table} ADD COLUMN {col} TEXT")
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


def delete_family(family_id):
    conn = get_connection()
    try:
        cursor = conn.cursor()
        cursor.execute("DELETE FROM families WHERE id = ?", (family_id,))
        conn.commit()
    finally:
        conn.close()


# ── Contact CRUD ──────────────────────────────────────────────────────────────

def add_contact(name, email, category, family_id=None):
    conn = get_connection()
    try:
        cursor = conn.cursor()
        cursor.execute(
            "INSERT INTO roster (name, email, category, family_id) VALUES (?, ?, ?, ?)",
            (name, email, category, family_id),
        )
        conn.commit()
    finally:
        conn.close()


def get_contacts():
    conn = get_connection()
    try:
        cursor = conn.cursor()
        cursor.execute("""
            SELECT r.id, r.name, r.email, r.category, COALESCE(f.name, '')
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
            FROM roster r
            LEFT JOIN families f ON r.family_id = f.id
            WHERE r.family_id = ?
            ORDER BY r.name
        """, (family_id,))
        return cursor.fetchall()
    finally:
        conn.close()


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
        cursor.execute("SELECT id, subject, target_type, target_id, scheduled_at, status, sent_at, result, recurrence FROM scheduled_emails ORDER BY scheduled_at DESC")
        return cursor.fetchall()
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

def log_email(subject, target_description, recipient_count, sent_count, failed_count):
    conn = get_connection()
    try:
        cursor = conn.cursor()
        cursor.execute("""
            INSERT INTO email_history (subject, target_description, recipient_count, sent_count, failed_count, sent_at)
            VALUES (?, ?, ?, ?, ?, ?)
        """, (subject, target_description, recipient_count, sent_count, failed_count, datetime.now().isoformat()))
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
