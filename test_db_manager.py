"""Comprehensive tests for db_manager.py — all 50 public functions."""
import json
import os
import tempfile
import pytest
from datetime import datetime, timedelta

import db_manager


@pytest.fixture(autouse=True)
def tmp_db(tmp_path):
    """Point db_manager at a fresh temp database for every test."""
    db_file = str(tmp_path / "test.db")
    db_manager.DB_PATH = db_file
    db_manager.init_db()
    yield db_file


# ── init_db / get_connection ────────────────────────────────────────────────

class TestInitAndConnection:
    def test_init_creates_tables(self):
        conn = db_manager.get_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table' ORDER BY name")
        tables = {r[0] for r in cursor.fetchall()}
        conn.close()
        expected = {"families", "roster", "settings", "groups_", "group_members",
                    "family_members", "scheduled_emails", "email_templates",
                    "email_history", "email_history_details"}
        assert expected.issubset(tables)

    def test_init_idempotent(self):
        db_manager.init_db()
        db_manager.init_db()
        assert db_manager.get_families() == []

    def test_foreign_keys_enabled(self):
        conn = db_manager.get_connection()
        cursor = conn.cursor()
        cursor.execute("PRAGMA foreign_keys")
        assert cursor.fetchone()[0] == 1
        conn.close()

    def test_migration_columns_exist(self):
        conn = db_manager.get_connection()
        cursor = conn.cursor()
        cursor.execute("PRAGMA table_info(roster)")
        cols = {r[1] for r in cursor.fetchall()}
        conn.close()
        assert {"phone", "notes", "opt_out", "created_at", "last_emailed_at", "email_count"}.issubset(cols)

    def test_scheduled_emails_migration_columns(self):
        conn = db_manager.get_connection()
        cursor = conn.cursor()
        cursor.execute("PRAGMA table_info(scheduled_emails)")
        cols = {r[1] for r in cursor.fetchall()}
        conn.close()
        assert "recurrence" in cols
        assert "manual_emails" in cols


# ── Settings CRUD ───────────────────────────────────────────────────────────

class TestSettings:
    def test_get_missing_setting_returns_none(self):
        assert db_manager.get_setting("nonexistent") is None

    def test_set_and_get(self):
        db_manager.set_setting("sender_email", "a@b.com")
        assert db_manager.get_setting("sender_email") == "a@b.com"

    def test_overwrite(self):
        db_manager.set_setting("key", "v1")
        db_manager.set_setting("key", "v2")
        assert db_manager.get_setting("key") == "v2"

    def test_multiple_keys(self):
        db_manager.set_setting("a", "1")
        db_manager.set_setting("b", "2")
        assert db_manager.get_setting("a") == "1"
        assert db_manager.get_setting("b") == "2"


# ── Family CRUD ─────────────────────────────────────────────────────────────

class TestFamilies:
    def test_add_and_get(self):
        db_manager.add_family("Smith")
        fams = db_manager.get_families()
        assert len(fams) == 1
        assert fams[0][1] == "Smith"

    def test_sorted_by_name(self):
        db_manager.add_family("Zulu")
        db_manager.add_family("Alpha")
        fams = db_manager.get_families()
        assert fams[0][1] == "Alpha"
        assert fams[1][1] == "Zulu"

    def test_rename(self):
        db_manager.add_family("Old")
        fid = db_manager.get_families()[0][0]
        db_manager.rename_family(fid, "New")
        assert db_manager.get_families()[0][1] == "New"

    def test_delete(self):
        db_manager.add_family("Gone")
        fid = db_manager.get_families()[0][0]
        db_manager.delete_family(fid)
        assert db_manager.get_families() == []

    def test_unique_name_constraint(self):
        db_manager.add_family("Dup")
        with pytest.raises(Exception):
            db_manager.add_family("Dup")


# ── Contact CRUD ────────────────────────────────────────────────────────────

class TestContacts:
    def test_add_and_get(self):
        db_manager.add_contact("John", "john@x.com", "Single")
        contacts = db_manager.get_contacts()
        assert len(contacts) == 1
        assert contacts[0][1] == "John"
        assert contacts[0][2] == "john@x.com"
        assert contacts[0][3] == "Single"

    def test_add_with_family(self):
        db_manager.add_family("Doe")
        fid = db_manager.get_families()[0][0]
        db_manager.add_contact("Jane", "jane@x.com", "Family", family_id=fid)
        contacts = db_manager.get_contacts()
        assert contacts[0][4] == "Doe"  # family name

    def test_add_with_phone_and_notes(self):
        db_manager.add_contact("Bob", "bob@x.com", "Single", phone="555-1234", notes="Elder")
        c = db_manager.get_contacts()[0]
        assert c[5] == "555-1234"
        assert c[6] == "Elder"

    def test_created_at_populated(self):
        db_manager.add_contact("Tim", "tim@x.com", "Single")
        c = db_manager.get_contacts()[0]
        assert c[8] is not None  # created_at

    def test_opt_out_default_false(self):
        db_manager.add_contact("Amy", "amy@x.com", "Single")
        c = db_manager.get_contacts()[0]
        assert c[7] == 0  # opt_out

    def test_update_contact(self):
        db_manager.add_contact("Old", "old@x.com", "Single")
        cid = db_manager.get_contacts()[0][0]
        db_manager.update_contact(cid, "New", "new@x.com", "Family", phone="111", notes="Updated")
        c = db_manager.get_contacts()[0]
        assert c[1] == "New"
        assert c[2] == "new@x.com"
        assert c[3] == "Family"
        assert c[5] == "111"
        assert c[6] == "Updated"

    def test_delete_contact(self):
        db_manager.add_contact("Del", "del@x.com", "Single")
        cid = db_manager.get_contacts()[0][0]
        db_manager.delete_contact(cid)
        assert db_manager.get_contacts() == []

    def test_set_opt_out(self):
        db_manager.add_contact("Opt", "opt@x.com", "Single")
        cid = db_manager.get_contacts()[0][0]
        db_manager.set_contact_opt_out(cid, True)
        assert db_manager.get_contacts()[0][7] == 1
        db_manager.set_contact_opt_out(cid, False)
        assert db_manager.get_contacts()[0][7] == 0

    def test_email_count_default_zero(self):
        db_manager.add_contact("Zero", "zero@x.com", "Single")
        c = db_manager.get_contacts()[0]
        assert c[10] == 0  # email_count

    def test_update_email_stats(self):
        db_manager.add_contact("A", "a@x.com", "Single")
        db_manager.add_contact("B", "b@x.com", "Single")
        db_manager.update_contact_email_stats(["a@x.com", "b@x.com"])
        contacts = db_manager.get_contacts()
        for c in contacts:
            assert c[10] == 1  # email_count
            assert c[9] is not None  # last_emailed_at
        # Call again
        db_manager.update_contact_email_stats(["a@x.com"])
        contacts = db_manager.get_contacts()
        counts = {c[1]: c[10] for c in contacts}
        assert counts["A"] == 2
        assert counts["B"] == 1

    def test_update_email_stats_empty_list(self):
        db_manager.update_contact_email_stats([])  # should not error

    def test_category_constraint(self):
        with pytest.raises(Exception):
            db_manager.add_contact("Bad", "bad@x.com", "InvalidCategory")

    def test_get_contacts_by_family(self):
        db_manager.add_family("Fam")
        fid = db_manager.get_families()[0][0]
        db_manager.add_contact("In", "in@x.com", "Family", family_id=fid)
        db_manager.add_contact("Out", "out@x.com", "Single")
        members = db_manager.get_contacts_by_family(fid)
        assert len(members) == 1
        assert members[0][1] == "In"

    def test_bulk_update_category(self):
        db_manager.add_contact("A", "a@x.com", "Single")
        db_manager.add_contact("B", "b@x.com", "Single")
        ids = [c[0] for c in db_manager.get_contacts()]
        db_manager.bulk_update_category(ids, "Family")
        for c in db_manager.get_contacts():
            assert c[3] == "Family"

    def test_bulk_add_to_group(self):
        db_manager.add_group("G1")
        gid = db_manager.get_groups()[0][0]
        db_manager.add_contact("A", "a@x.com", "Single")
        db_manager.add_contact("B", "b@x.com", "Single")
        ids = [c[0] for c in db_manager.get_contacts()]
        db_manager.bulk_add_to_group(gid, ids)
        members = db_manager.get_group_members(gid)
        assert len(members) == 2

    def test_bulk_add_to_group_idempotent(self):
        db_manager.add_group("G1")
        gid = db_manager.get_groups()[0][0]
        db_manager.add_contact("A", "a@x.com", "Single")
        cid = db_manager.get_contacts()[0][0]
        db_manager.bulk_add_to_group(gid, [cid, cid])
        assert len(db_manager.get_group_members(gid)) == 1


# ── Family Members (junction) ──────────────────────────────────────────────

class TestFamilyMembers:
    def test_add_and_get(self):
        db_manager.add_family("F")
        fid = db_manager.get_families()[0][0]
        db_manager.add_contact("C", "c@x.com", "Single")
        cid = db_manager.get_contacts()[0][0]
        db_manager.add_family_member(fid, cid)
        members = db_manager.get_family_members_via_junction(fid)
        assert len(members) == 1
        assert members[0][1] == "C"

    def test_remove(self):
        db_manager.add_family("F")
        fid = db_manager.get_families()[0][0]
        db_manager.add_contact("C", "c@x.com", "Single")
        cid = db_manager.get_contacts()[0][0]
        db_manager.add_family_member(fid, cid)
        db_manager.remove_family_member(fid, cid)
        assert db_manager.get_family_members_via_junction(fid) == []

    def test_add_idempotent(self):
        db_manager.add_family("F")
        fid = db_manager.get_families()[0][0]
        db_manager.add_contact("C", "c@x.com", "Single")
        cid = db_manager.get_contacts()[0][0]
        db_manager.add_family_member(fid, cid)
        db_manager.add_family_member(fid, cid)  # should not error
        assert len(db_manager.get_family_members_via_junction(fid)) == 1

    def test_get_contact_families(self):
        db_manager.add_family("F1")
        db_manager.add_family("F2")
        fams = db_manager.get_families()
        db_manager.add_contact("C", "c@x.com", "Single")
        cid = db_manager.get_contacts()[0][0]
        db_manager.add_family_member(fams[0][0], cid)
        db_manager.add_family_member(fams[1][0], cid)
        cf = db_manager.get_contact_families()
        assert cid in cf
        assert len(cf[cid]) == 2

    def test_auto_junction_on_add_contact(self):
        db_manager.add_family("F")
        fid = db_manager.get_families()[0][0]
        db_manager.add_contact("C", "c@x.com", "Family", family_id=fid)
        members = db_manager.get_family_members_via_junction(fid)
        assert len(members) == 1

    def test_cascade_delete_family(self):
        db_manager.add_family("F")
        fid = db_manager.get_families()[0][0]
        db_manager.add_contact("C", "c@x.com", "Family", family_id=fid)
        cid = db_manager.get_contacts()[0][0]
        db_manager.add_family_member(fid, cid)
        db_manager.delete_family(fid)
        assert db_manager.get_family_members_via_junction(fid) == []


# ── Group CRUD ──────────────────────────────────────────────────────────────

class TestGroups:
    def test_add_and_get(self):
        db_manager.add_group("Youth")
        groups = db_manager.get_groups()
        assert len(groups) == 1
        assert groups[0][1] == "Youth"

    def test_rename(self):
        db_manager.add_group("Old")
        gid = db_manager.get_groups()[0][0]
        db_manager.rename_group(gid, "New")
        assert db_manager.get_groups()[0][1] == "New"

    def test_delete(self):
        db_manager.add_group("Gone")
        gid = db_manager.get_groups()[0][0]
        db_manager.delete_group(gid)
        assert db_manager.get_groups() == []

    def test_unique_name(self):
        db_manager.add_group("Dup")
        with pytest.raises(Exception):
            db_manager.add_group("Dup")

    def test_add_and_remove_member(self):
        db_manager.add_group("G")
        gid = db_manager.get_groups()[0][0]
        db_manager.add_contact("M", "m@x.com", "Single")
        cid = db_manager.get_contacts()[0][0]
        db_manager.add_group_member(gid, cid)
        members = db_manager.get_group_members(gid)
        assert len(members) == 1
        db_manager.remove_group_member(gid, cid)
        assert db_manager.get_group_members(gid) == []

    def test_add_member_idempotent(self):
        db_manager.add_group("G")
        gid = db_manager.get_groups()[0][0]
        db_manager.add_contact("M", "m@x.com", "Single")
        cid = db_manager.get_contacts()[0][0]
        db_manager.add_group_member(gid, cid)
        db_manager.add_group_member(gid, cid)
        assert len(db_manager.get_group_members(gid)) == 1

    def test_get_contact_groups(self):
        db_manager.add_group("G1")
        db_manager.add_group("G2")
        groups = db_manager.get_groups()
        db_manager.add_contact("C", "c@x.com", "Single")
        cid = db_manager.get_contacts()[0][0]
        db_manager.add_group_member(groups[0][0], cid)
        db_manager.add_group_member(groups[1][0], cid)
        cg = db_manager.get_contact_groups()
        assert cid in cg
        assert len(cg[cid]) == 2

    def test_cascade_delete_group(self):
        db_manager.add_group("G")
        gid = db_manager.get_groups()[0][0]
        db_manager.add_contact("C", "c@x.com", "Single")
        cid = db_manager.get_contacts()[0][0]
        db_manager.add_group_member(gid, cid)
        db_manager.delete_group(gid)
        assert db_manager.get_group_members(gid) == []

    def test_cascade_delete_contact(self):
        db_manager.add_group("G")
        gid = db_manager.get_groups()[0][0]
        db_manager.add_contact("C", "c@x.com", "Single")
        cid = db_manager.get_contacts()[0][0]
        db_manager.add_group_member(gid, cid)
        db_manager.delete_contact(cid)
        assert db_manager.get_group_members(gid) == []


# ── Scheduled Emails CRUD ──────────────────────────────────────────────────

class TestScheduledEmails:
    def _schedule_one(self, subject="Test", scheduled_at=None, recurrence=None, manual_emails=None):
        sa = scheduled_at or (datetime.now() + timedelta(hours=1)).isoformat()
        db_manager.schedule_email(subject, "<p>body</p>", "plain", "all", None, [1], [], sa,
                                  recurrence=recurrence, manual_emails=manual_emails)

    def test_schedule_and_get(self):
        self._schedule_one("Hello")
        emails = db_manager.get_scheduled_emails()
        assert len(emails) == 1
        assert emails[0][1] == "Hello"
        assert emails[0][5] == "pending"

    def test_get_by_id(self):
        self._schedule_one("Detail")
        eid = db_manager.get_scheduled_emails()[0][0]
        e = db_manager.get_scheduled_email_by_id(eid)
        assert e is not None
        assert e[1] == "Detail"

    def test_get_by_id_missing(self):
        assert db_manager.get_scheduled_email_by_id(999) is None

    def test_cancel(self):
        self._schedule_one()
        eid = db_manager.get_scheduled_emails()[0][0]
        db_manager.cancel_scheduled_email(eid)
        assert db_manager.get_scheduled_emails()[0][5] == "cancelled"

    def test_cancel_only_pending(self):
        self._schedule_one()
        eid = db_manager.get_scheduled_emails()[0][0]
        db_manager.update_email_status(eid, "sent")
        db_manager.cancel_scheduled_email(eid)
        assert db_manager.get_scheduled_emails()[0][5] == "sent"  # unchanged

    def test_get_due_emails(self):
        past = (datetime.now() - timedelta(hours=1)).isoformat()
        future = (datetime.now() + timedelta(hours=1)).isoformat()
        self._schedule_one("Past", scheduled_at=past)
        self._schedule_one("Future", scheduled_at=future)
        due = db_manager.get_due_emails()
        assert len(due) == 1
        assert due[0][1] == "Past"

    def test_update_email_status(self):
        self._schedule_one()
        eid = db_manager.get_scheduled_emails()[0][0]
        db_manager.update_email_status(eid, "sent", {"sent": 5, "failed": 0})
        e = db_manager.get_scheduled_emails()[0]
        assert e[5] == "sent"
        assert e[6] is not None  # sent_at

    def test_update_scheduled_email(self):
        self._schedule_one("Old Subject")
        eid = db_manager.get_scheduled_emails()[0][0]
        new_at = (datetime.now() + timedelta(days=2)).isoformat()
        db_manager.update_scheduled_email(eid, "New Subject", "<p>new</p>", "new plain",
                                          "all", None, None, None, new_at)
        e = db_manager.get_scheduled_email_by_id(eid)
        assert e[1] == "New Subject"

    def test_update_only_pending(self):
        self._schedule_one("Sent")
        eid = db_manager.get_scheduled_emails()[0][0]
        db_manager.update_email_status(eid, "sent")
        db_manager.update_scheduled_email(eid, "Changed", "<p></p>", "p", "all", None, None, None,
                                          datetime.now().isoformat())
        e = db_manager.get_scheduled_email_by_id(eid)
        assert e[1] == "Sent"  # unchanged

    def test_duplicate(self):
        self._schedule_one("Original")
        eid = db_manager.get_scheduled_emails()[0][0]
        new_id = db_manager.duplicate_scheduled_email(eid)
        assert new_id is not None
        emails = db_manager.get_scheduled_emails()
        assert len(emails) == 2

    def test_duplicate_missing(self):
        assert db_manager.duplicate_scheduled_email(999) is None

    def test_with_recurrence(self):
        rec = {"type": "daily"}
        self._schedule_one(recurrence=rec)
        e = db_manager.get_scheduled_emails()[0]
        assert e[8] is not None  # recurrence column

    def test_with_manual_emails(self):
        self._schedule_one(manual_emails=["extra@x.com"])
        e = db_manager.get_scheduled_emails()[0]
        assert e[10] is not None  # manual_emails column


# ── Email Templates CRUD ───────────────────────────────────────────────────

class TestTemplates:
    def test_save_and_get(self):
        db_manager.save_template("Welcome", "Hi", "<p>Hi</p>")
        tpls = db_manager.get_templates()
        assert len(tpls) == 1
        assert tpls[0][1] == "Welcome"

    def test_update(self):
        db_manager.save_template("T", "Old", "<p>Old</p>")
        tid = db_manager.get_templates()[0][0]
        db_manager.update_template(tid, "New Subject", "<p>New</p>")
        t = db_manager.get_templates()[0]
        assert t[2] == "New Subject"

    def test_delete(self):
        db_manager.save_template("Gone", "S", "<p></p>")
        tid = db_manager.get_templates()[0][0]
        db_manager.delete_template(tid)
        assert db_manager.get_templates() == []

    def test_with_recipients(self):
        recipients = {"target_type": "group", "target_id": 1}
        db_manager.save_template("R", "S", "<p></p>", recipients=recipients)
        t = db_manager.get_templates()[0]
        assert json.loads(t[4]) == recipients

    def test_duplicate(self):
        db_manager.save_template("Orig", "S", "<p></p>")
        tid = db_manager.get_templates()[0][0]
        new_id = db_manager.duplicate_template(tid)
        assert new_id is not None
        tpls = db_manager.get_templates()
        assert len(tpls) == 2
        names = {t[1] for t in tpls}
        assert "Orig (Copy)" in names

    def test_duplicate_missing(self):
        assert db_manager.duplicate_template(999) is None

    def test_sorted_by_name(self):
        db_manager.save_template("Zeta", "S", "<p></p>")
        db_manager.save_template("Alpha", "S", "<p></p>")
        tpls = db_manager.get_templates()
        assert tpls[0][1] == "Alpha"


# ── Email History CRUD ──────────────────────────────────────────────────────

class TestEmailHistory:
    def test_log_and_get(self):
        db_manager.log_email("Subject", "all contacts", 10, 9, 1)
        history = db_manager.get_email_history()
        assert len(history) == 1
        assert history[0][1] == "Subject"
        assert history[0][3] == 10  # recipient_count
        assert history[0][4] == 9   # sent_count
        assert history[0][5] == 1   # failed_count

    def test_with_details(self):
        details = [
            {"email": "a@x.com", "name": "A", "status": "sent"},
            {"email": "b@x.com", "name": "B", "status": "failed", "error": "timeout"},
        ]
        db_manager.log_email("S", "desc", 2, 1, 1, recipient_details=details)
        hid = db_manager.get_email_history()[0][0]
        d = db_manager.get_email_history_details(hid)
        assert len(d) == 2
        statuses = {r[2] for r in d}
        assert statuses == {"sent", "failed"}

    def test_history_ordered_desc(self):
        db_manager.log_email("First", "d", 1, 1, 0)
        db_manager.log_email("Second", "d", 1, 1, 0)
        history = db_manager.get_email_history()
        assert history[0][1] == "Second"

    def test_filtered_by_date(self):
        db_manager.log_email("Old", "d", 1, 1, 0)
        # Manually insert one with a known date
        conn = db_manager.get_connection()
        conn.execute("UPDATE email_history SET sent_at = '2020-01-01T00:00:00'")
        conn.commit()
        conn.close()
        db_manager.log_email("New", "d", 1, 1, 0)
        results = db_manager.get_email_history_filtered(start_date="2024-01-01")
        assert len(results) == 1
        assert results[0][1] == "New"

    def test_filtered_end_date(self):
        db_manager.log_email("Recent", "d", 1, 1, 0)
        results = db_manager.get_email_history_filtered(end_date="2020-01-01")
        assert len(results) == 0

    def test_filtered_no_params(self):
        db_manager.log_email("Any", "d", 1, 1, 0)
        results = db_manager.get_email_history_filtered()
        assert len(results) == 1


# ── Analytics ───────────────────────────────────────────────────────────────

class TestAnalytics:
    def test_empty(self):
        a = db_manager.get_analytics()
        assert a["totals"]["emails_sent"] == 0
        assert a["totals"]["failure_rate"] == 0
        assert a["weekly"] == []
        assert a["top_failed"] == []

    def test_with_data(self):
        db_manager.log_email("S1", "d", 10, 8, 2, [
            {"email": "fail@x.com", "name": "F", "status": "failed", "error": "err"},
            {"email": "ok@x.com", "name": "O", "status": "sent"},
        ])
        a = db_manager.get_analytics()
        assert a["totals"]["emails_sent"] == 1
        assert a["totals"]["recipients_reached"] == 8
        assert a["totals"]["total_failures"] == 2
        assert a["totals"]["failure_rate"] > 0
        assert len(a["top_failed"]) == 1
        assert a["top_failed"][0]["email"] == "fail@x.com"


# ── Backup / Restore ───────────────────────────────────────────────────────

class TestBackupRestore:
    def test_backup_and_restore(self, tmp_path):
        db_manager.add_contact("Before", "before@x.com", "Single")
        backup_path = str(tmp_path / "backup.db")
        db_manager.backup_database(backup_path)
        assert os.path.exists(backup_path)

        # Add more data after backup
        db_manager.add_contact("After", "after@x.com", "Single")
        assert len(db_manager.get_contacts()) == 2

        # Restore
        db_manager.restore_database(backup_path)
        assert len(db_manager.get_contacts()) == 1
        assert db_manager.get_contacts()[0][1] == "Before"


# ── Recurrence ──────────────────────────────────────────────────────────────

class TestRecurrence:
    def test_once_returns_none(self):
        dt = datetime(2025, 6, 1, 9, 0)
        assert db_manager.compute_next_occurrence(dt, {"type": "once"}) is None

    def test_daily(self):
        dt = datetime(2025, 6, 1, 9, 0)
        nxt = db_manager.compute_next_occurrence(dt, {"type": "daily"})
        assert nxt == datetime(2025, 6, 2, 9, 0)

    def test_every_other_day(self):
        dt = datetime(2025, 6, 1, 9, 0)
        nxt = db_manager.compute_next_occurrence(dt, {"type": "every_other_day"})
        assert nxt == datetime(2025, 6, 3, 9, 0)

    def test_weekly_next_day_same_week(self):
        # 2025-06-02 is Monday (weekday=0). days=[3] = JS Wed → Python (3-1)%7=2 = Wed
        dt = datetime(2025, 6, 2, 9, 0)  # Monday
        nxt = db_manager.compute_next_occurrence(dt, {"type": "weekly", "days": [3]})
        assert nxt == datetime(2025, 6, 4, 9, 0)  # Wednesday

    def test_weekly_wrap_to_next_week(self):
        # 2025-06-06 is Friday (weekday=4). days=[2] = JS Monday → Python Sunday (6)
        dt = datetime(2025, 6, 6, 9, 0)  # Friday
        nxt = db_manager.compute_next_occurrence(dt, {"type": "weekly", "days": [2]})
        assert nxt is not None
        assert nxt > dt

    def test_weekly_no_days_returns_none(self):
        dt = datetime(2025, 6, 1, 9, 0)
        assert db_manager.compute_next_occurrence(dt, {"type": "weekly", "days": []}) is None

    def test_every_other_week(self):
        # days=[2] = JS Mon → Python (2-1)%7=1 = Tue. From Monday, next Tue is +1 day.
        # "every_other_week" only applies the 2-week offset when wrapping, not same-week matches.
        dt = datetime(2025, 6, 2, 9, 0)  # Monday
        nxt = db_manager.compute_next_occurrence(dt, {"type": "every_other_week", "days": [2]})
        assert nxt is not None
        assert nxt == datetime(2025, 6, 3, 9, 0)  # Tuesday (same week match)

    def test_every_other_week_wraps(self):
        # From Friday, days=[2] (JS Mon → Python Tue). No same-week match, wraps with 2-week offset.
        dt = datetime(2025, 6, 6, 9, 0)  # Friday (weekday=4)
        nxt = db_manager.compute_next_occurrence(dt, {"type": "every_other_week", "days": [2]})
        assert nxt is not None
        assert (nxt - dt).days >= 7  # should skip to 2 weeks out

    def test_monthly(self):
        dt = datetime(2025, 6, 15, 9, 0)
        nxt = db_manager.compute_next_occurrence(dt, {"type": "monthly", "day_of_month": 15})
        assert nxt == datetime(2025, 7, 15, 9, 0)

    def test_monthly_clamps_day(self):
        dt = datetime(2025, 1, 31, 9, 0)
        nxt = db_manager.compute_next_occurrence(dt, {"type": "monthly", "day_of_month": 31})
        assert nxt.month == 2
        assert nxt.day == 28  # Feb doesn't have 31

    def test_monthly_dec_to_jan(self):
        dt = datetime(2025, 12, 15, 9, 0)
        nxt = db_manager.compute_next_occurrence(dt, {"type": "monthly", "day_of_month": 15})
        assert nxt == datetime(2026, 1, 15, 9, 0)

    def test_end_date_stops_recurrence(self):
        dt = datetime(2025, 6, 1, 9, 0)
        nxt = db_manager.compute_next_occurrence(dt, {"type": "daily", "end_date": "2025-06-01T23:59:59"})
        assert nxt is None

    def test_end_date_allows_within_range(self):
        dt = datetime(2025, 6, 1, 9, 0)
        nxt = db_manager.compute_next_occurrence(dt, {"type": "daily", "end_date": "2025-06-10T00:00:00"})
        assert nxt == datetime(2025, 6, 2, 9, 0)

    def test_unknown_type_returns_none(self):
        dt = datetime(2025, 6, 1, 9, 0)
        assert db_manager.compute_next_occurrence(dt, {"type": "unknown"}) is None

    def test_weekly_multiple_days(self):
        # 2025-06-02 is Monday. days=[2,4,6] = JS Mon/Wed/Fri → Python Sun/Tue/Thu = [6,1,3]
        # sorted py_days = [1, 3, 6]. Current weekday=0 (Mon). diff for 1=1>0 → Tuesday
        dt = datetime(2025, 6, 2, 9, 0)  # Monday
        nxt = db_manager.compute_next_occurrence(dt, {"type": "weekly", "days": [2, 4, 6]})
        assert nxt is not None
        assert nxt > dt
        # Should pick the nearest future day in the list
        assert nxt == datetime(2025, 6, 3, 9, 0)  # Tuesday (py_day 1)

    def test_weekly_multiple_days_picks_closest(self):
        # 2025-06-04 is Wednesday (weekday=2). py_days=[1,3,6]. diff for 3=1>0 → Thursday
        dt = datetime(2025, 6, 4, 9, 0)  # Wednesday
        nxt = db_manager.compute_next_occurrence(dt, {"type": "weekly", "days": [2, 4, 6]})
        assert nxt == datetime(2025, 6, 5, 9, 0)  # Thursday (py_day 3)


# ── Additional edge cases ─────────────────────────────────────────────────

class TestContactDuplicateEmails:
    def test_two_contacts_same_email_allowed(self):
        db_manager.add_contact("Alice", "shared@x.com", "Single")
        db_manager.add_contact("Bob", "shared@x.com", "Single")
        contacts = db_manager.get_contacts()
        assert len(contacts) == 2
        emails = [c[2] for c in contacts]
        assert emails.count("shared@x.com") == 2


class TestUpdateContactFamilySync:
    def test_update_contact_syncs_junction(self):
        db_manager.add_family("Fam")
        fid = db_manager.get_families()[0][0]
        db_manager.add_contact("C", "c@x.com", "Single")
        cid = db_manager.get_contacts()[0][0]
        # No family_members entry yet
        assert db_manager.get_family_members_via_junction(fid) == []
        # Update contact to add family
        db_manager.update_contact(cid, "C", "c@x.com", "Family", family_id=fid)
        # Junction should now have the member
        members = db_manager.get_family_members_via_junction(fid)
        assert len(members) == 1
        assert members[0][0] == cid


class TestUpdateContactFamilyChange:
    def test_changing_family_removes_old_junction(self):
        db_manager.add_family("FamA")
        db_manager.add_family("FamB")
        families = db_manager.get_families()
        fid_a = [f for f in families if f[1] == "FamA"][0][0]
        fid_b = [f for f in families if f[1] == "FamB"][0][0]
        db_manager.add_contact("C", "c@x.com", "Family", family_id=fid_a)
        cid = db_manager.get_contacts()[0][0]
        assert len(db_manager.get_family_members_via_junction(fid_a)) == 1
        # Change family from A to B
        db_manager.update_contact(cid, "C", "c@x.com", "Family", family_id=fid_b)
        # Old family should have no members, new family should have one
        assert db_manager.get_family_members_via_junction(fid_a) == []
        assert len(db_manager.get_family_members_via_junction(fid_b)) == 1

    def test_removing_family_cleans_junction(self):
        db_manager.add_family("Fam")
        fid = db_manager.get_families()[0][0]
        db_manager.add_contact("C", "c@x.com", "Family", family_id=fid)
        cid = db_manager.get_contacts()[0][0]
        assert len(db_manager.get_family_members_via_junction(fid)) == 1
        # Remove family assignment
        db_manager.update_contact(cid, "C", "c@x.com", "Single", family_id=None)
        assert db_manager.get_family_members_via_junction(fid) == []


class TestDeleteContactFamilyFK:
    def test_delete_contact_nulls_family_id(self):
        db_manager.add_family("Fam")
        fid = db_manager.get_families()[0][0]
        db_manager.add_contact("C", "c@x.com", "Family", family_id=fid)
        cid = db_manager.get_contacts()[0][0]
        db_manager.delete_contact(cid)
        # Family should still exist
        assert len(db_manager.get_families()) == 1
        # Junction should be cleaned up (cascade)
        assert db_manager.get_family_members_via_junction(fid) == []


class TestGetDueEmailsCustomNow:
    def test_custom_now_iso(self):
        past = "2025-01-01T00:00:00"
        future = "2025-12-31T23:59:59"
        db_manager.schedule_email("Past", "<p></p>", "t", "all", None, [], [], past)
        db_manager.schedule_email("Future", "<p></p>", "t", "all", None, [], [], future)
        # Query with a time between the two
        due = db_manager.get_due_emails(now_iso="2025-06-01T00:00:00")
        assert len(due) == 1
        assert due[0][1] == "Past"

    def test_custom_now_iso_returns_all_past(self):
        db_manager.schedule_email("A", "<p></p>", "t", "all", None, [], [], "2025-01-01T00:00:00")
        db_manager.schedule_email("B", "<p></p>", "t", "all", None, [], [], "2025-03-01T00:00:00")
        due = db_manager.get_due_emails(now_iso="2025-06-01T00:00:00")
        assert len(due) == 2


class TestUpdateEmailStatusFailed:
    def test_sent_at_null_for_failed(self):
        sa = (datetime.now() + timedelta(hours=1)).isoformat()
        db_manager.schedule_email("Fail", "<p></p>", "t", "all", None, [], [], sa)
        eid = db_manager.get_scheduled_emails()[0][0]
        db_manager.update_email_status(eid, "failed", {"error": "boom"})
        e = db_manager.get_scheduled_emails()[0]
        assert e[5] == "failed"
        assert e[6] is None  # sent_at should be None for failed

    def test_sent_at_set_for_sent(self):
        sa = (datetime.now() + timedelta(hours=1)).isoformat()
        db_manager.schedule_email("OK", "<p></p>", "t", "all", None, [], [], sa)
        eid = db_manager.get_scheduled_emails()[0][0]
        db_manager.update_email_status(eid, "sent", {"sent": 1})
        e = db_manager.get_scheduled_emails()[0]
        assert e[5] == "sent"
        assert e[6] is not None  # sent_at should be set


class TestBackupEdgeCases:
    def test_backup_overwrites_existing(self, tmp_path):
        backup_path = str(tmp_path / "backup.db")
        db_manager.add_contact("First", "first@x.com", "Single")
        db_manager.backup_database(backup_path)
        # Modify and backup again
        db_manager.add_contact("Second", "second@x.com", "Single")
        db_manager.backup_database(backup_path)
        # Restore and verify latest backup
        db_manager.restore_database(backup_path)
        contacts = db_manager.get_contacts()
        assert len(contacts) == 2

    def test_restore_replaces_current_data(self, tmp_path):
        backup_path = str(tmp_path / "backup.db")
        db_manager.backup_database(backup_path)  # empty DB
        db_manager.add_contact("Added", "added@x.com", "Single")
        assert len(db_manager.get_contacts()) == 1
        db_manager.restore_database(backup_path)
        # After restore, should be empty again
        assert len(db_manager.get_contacts()) == 0


class TestLogEmailEdgeCases:
    def test_log_with_empty_details_list(self):
        db_manager.log_email("Subj", "desc", 0, 0, 0, recipient_details=[])
        history = db_manager.get_email_history()
        assert len(history) == 1
        hid = history[0][0]
        details = db_manager.get_email_history_details(hid)
        assert details == []
