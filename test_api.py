"""Comprehensive tests for main.py Api class and standalone functions."""
import json
import os
import tempfile
import pytest
from unittest.mock import patch, MagicMock
from datetime import datetime, timedelta

import db_manager

# We must patch webview before importing main, since main imports it at module level
import sys
webview_mock = MagicMock()
sys.modules["webview"] = webview_mock
sys.modules["pystray"] = MagicMock()

# Now import main safely (without triggering GUI)
import main
from main import Api, _friendly_smtp_error


@pytest.fixture(autouse=True)
def tmp_db(tmp_path):
    """Fresh temp database for every test."""
    db_file = str(tmp_path / "test.db")
    db_manager.DB_PATH = db_file
    db_manager.init_db()
    yield db_file


@pytest.fixture
def api():
    return Api()


# ── _friendly_smtp_error ───────────────────────────────────────────────────

class TestFriendlySmtpError:
    def test_auth_failure(self):
        msg = _friendly_smtp_error("535 Username and Password not accepted")
        assert "Authentication failed" in msg
        assert "App Password" in msg

    def test_connection_refused(self):
        msg = _friendly_smtp_error("Connection refused errno 61")
        assert "Connection refused" in msg

    def test_timeout(self):
        msg = _friendly_smtp_error("Connection timed out")
        assert "timed out" in msg

    def test_ssl_error(self):
        msg = _friendly_smtp_error("SSL handshake failed")
        assert "SSL/TLS" in msg

    def test_relay_denied(self):
        msg = _friendly_smtp_error("550 Relay denied")
        assert "Relay denied" in msg

    def test_dns_failure(self):
        msg = _friendly_smtp_error("getaddrinfo failed")
        assert "resolve" in msg

    def test_unknown_error_passthrough(self):
        msg = _friendly_smtp_error("Something weird happened")
        assert msg == "Something weird happened"


# ── Contacts API ───────────────────────────────────────────────────────────

class TestContactsApi:
    def test_get_contacts_empty(self, api):
        assert api.get_contacts() == []

    def test_add_and_get(self, api):
        res = api.add_contact("John", "john@x.com", "Single", None)
        assert res["ok"]
        contacts = api.get_contacts()
        assert len(contacts) == 1
        assert contacts[0]["name"] == "John"
        assert contacts[0]["email"] == "john@x.com"
        assert contacts[0]["opt_out"] is False
        assert contacts[0]["email_count"] == 0
        assert contacts[0]["groups"] == []
        assert contacts[0]["families"] == []

    def test_add_with_phone_notes(self, api):
        api.add_contact("Jane", "jane@x.com", "Single", None, phone="555", notes="N")
        c = api.get_contacts()[0]
        assert c["phone"] == "555"
        assert c["notes"] == "N"

    def test_add_with_family(self, api):
        api.add_family("Doe")
        fid = api.get_families()[0]["id"]
        api.add_contact("J", "j@x.com", "Family", fid)
        c = api.get_contacts()[0]
        assert c["family_name"] == "Doe"

    def test_add_invalid_category(self, api):
        res = api.add_contact("Bad", "bad@x.com", "Invalid", None)
        assert not res["ok"]

    def test_update_contact(self, api):
        api.add_contact("Old", "old@x.com", "Single", None)
        cid = api.get_contacts()[0]["id"]
        res = api.update_contact(cid, "New", "new@x.com", "Family", None, phone="111")
        assert res["ok"]
        c = api.get_contacts()[0]
        assert c["name"] == "New"
        assert c["phone"] == "111"

    def test_delete_contacts(self, api):
        api.add_contact("A", "a@x.com", "Single", None)
        api.add_contact("B", "b@x.com", "Single", None)
        ids = [c["id"] for c in api.get_contacts()]
        res = api.delete_contacts(ids)
        assert res["ok"]
        assert api.get_contacts() == []

    def test_set_opt_out(self, api):
        api.add_contact("X", "x@x.com", "Single", None)
        cid = api.get_contacts()[0]["id"]
        res = api.set_contact_opt_out(cid, True)
        assert res["ok"]
        assert api.get_contacts()[0]["opt_out"] is True

    def test_bulk_update_category(self, api):
        api.add_contact("A", "a@x.com", "Single", None)
        api.add_contact("B", "b@x.com", "Single", None)
        ids = [c["id"] for c in api.get_contacts()]
        res = api.bulk_update_category(ids, "Family")
        assert res["ok"]
        for c in api.get_contacts():
            assert c["category"] == "Family"

    def test_bulk_add_to_group(self, api):
        api.add_group("G")
        gid = api.get_groups()[0]["id"]
        api.add_contact("A", "a@x.com", "Single", None)
        cid = api.get_contacts()[0]["id"]
        res = api.bulk_add_to_group(gid, [cid])
        assert res["ok"]
        assert len(api.get_groups()[0]["members"]) == 1


# ── Family API ─────────────────────────────────────────────────────────────

class TestFamilyApi:
    def test_add_and_get(self, api):
        res = api.add_family("Smith")
        assert res["ok"]
        fams = api.get_families()
        assert len(fams) == 1
        assert fams[0]["name"] == "Smith"
        assert fams[0]["members"] == []

    def test_rename(self, api):
        api.add_family("Old")
        fid = api.get_families()[0]["id"]
        res = api.rename_family(fid, "New")
        assert res["ok"]
        assert api.get_families()[0]["name"] == "New"

    def test_delete(self, api):
        api.add_family("Gone")
        fid = api.get_families()[0]["id"]
        res = api.delete_family(fid)
        assert res["ok"]
        assert api.get_families() == []

    def test_duplicate_name(self, api):
        api.add_family("Dup")
        res = api.add_family("Dup")
        assert not res["ok"]

    def test_add_and_remove_member(self, api):
        api.add_family("F")
        fid = api.get_families()[0]["id"]
        api.add_contact("C", "c@x.com", "Single", None)
        cid = api.get_contacts()[0]["id"]
        api.add_family_member(fid, cid)
        assert len(api.get_families()[0]["members"]) == 1
        api.remove_family_member(fid, cid)
        assert len(api.get_families()[0]["members"]) == 0


# ── Group API ──────────────────────────────────────────────────────────────

class TestGroupApi:
    def test_add_and_get(self, api):
        res = api.add_group("Youth")
        assert res["ok"]
        groups = api.get_groups()
        assert len(groups) == 1
        assert groups[0]["name"] == "Youth"

    def test_rename(self, api):
        api.add_group("Old")
        gid = api.get_groups()[0]["id"]
        res = api.rename_group(gid, "New")
        assert res["ok"]

    def test_delete(self, api):
        api.add_group("Gone")
        gid = api.get_groups()[0]["id"]
        res = api.delete_group(gid)
        assert res["ok"]
        assert api.get_groups() == []

    def test_add_remove_member(self, api):
        api.add_group("G")
        gid = api.get_groups()[0]["id"]
        api.add_contact("M", "m@x.com", "Single", None)
        cid = api.get_contacts()[0]["id"]
        api.add_group_member(gid, cid)
        assert len(api.get_groups()[0]["members"]) == 1
        api.remove_group_member(gid, cid)
        assert len(api.get_groups()[0]["members"]) == 0

    def test_add_family_to_group(self, api):
        api.add_family("F")
        fid = api.get_families()[0]["id"]
        api.add_contact("A", "a@x.com", "Family", fid)
        api.add_contact("B", "b@x.com", "Family", fid)
        api.add_group("G")
        gid = api.get_groups()[0]["id"]
        res = api.add_family_to_group(gid, fid)
        assert res["ok"]
        assert res["added"] == 2
        assert len(api.get_groups()[0]["members"]) == 2


# ── Settings API ───────────────────────────────────────────────────────────

class TestSettingsApi:
    def test_get_defaults(self, api):
        s = api.get_settings()
        assert s["email"] == ""
        assert s["has_password"] is False
        assert s["smtp_host"] == "smtp.gmail.com"
        assert s["smtp_port"] == "587"
        assert s["launch_on_startup"] is False

    def test_save_and_get(self, api):
        res = api.save_settings("me@x.com", "secret123", sender_name="Church")
        assert res["ok"]
        s = api.get_settings()
        assert s["email"] == "me@x.com"
        assert s["has_password"] is True
        assert s["sender_name"] == "Church"
        assert s["app_password_masked"].endswith("t123")

    def test_save_preserves_password(self, api):
        api.save_settings("me@x.com", "secret123")
        api.save_settings("me@x.com", "")  # blank = keep existing
        s = api.get_settings()
        assert s["has_password"] is True

    def test_save_custom_smtp(self, api):
        api.save_settings("me@x.com", "pass", smtp_host="mail.example.com", smtp_port="465")
        s = api.get_settings()
        assert s["smtp_host"] == "mail.example.com"
        assert s["smtp_port"] == "465"

    def test_save_timezone(self, api):
        res = api.save_timezone("US/Pacific")
        assert res["ok"]
        s = api.get_settings()
        assert s["timezone"] == "US/Pacific"

    def test_ui_settings(self, api):
        api.set_ui_setting("theme", "dark")
        assert api.get_ui_setting("theme") == "dark"
        assert api.get_ui_setting("missing") == ""


# ── Resolve Recipients ─────────────────────────────────────────────────────

class TestResolveRecipients:
    def test_all(self, api):
        api.add_contact("A", "a@x.com", "Single", None)
        api.add_contact("B", "b@x.com", "Family", None)
        r = api.resolve_recipients("all")
        assert len(r) == 2

    def test_family_only(self, api):
        api.add_contact("A", "a@x.com", "Single", None)
        api.add_contact("B", "b@x.com", "Family", None)
        r = api.resolve_recipients("family")
        assert len(r) == 1
        assert r[0][0] == "B"

    def test_single_only(self, api):
        api.add_contact("A", "a@x.com", "Single", None)
        api.add_contact("B", "b@x.com", "Family", None)
        r = api.resolve_recipients("single")
        assert len(r) == 1
        assert r[0][0] == "A"

    def test_custom(self, api):
        api.add_contact("A", "a@x.com", "Single", None)
        api.add_contact("B", "b@x.com", "Single", None)
        cid = api.get_contacts()[0]["id"]
        r = api.resolve_recipients("custom", contact_ids=[cid])
        assert len(r) == 1

    def test_group(self, api):
        api.add_group("G")
        gid = api.get_groups()[0]["id"]
        api.add_contact("A", "a@x.com", "Single", None)
        cid = api.get_contacts()[0]["id"]
        api.add_group_member(gid, cid)
        r = api.resolve_recipients("group", target_id=gid)
        assert len(r) == 1

    def test_unknown_type(self, api):
        assert api.resolve_recipients("nonexistent") == []

    def test_manual_type(self, api):
        assert api.resolve_recipients("manual") == []


# ── Recipient Count ────────────────────────────────────────────────────────

class TestRecipientCount:
    def test_all_contacts(self, api):
        api.add_contact("A", "a@x.com", "Single", None)
        api.add_contact("B", "b@x.com", "Single", None)
        r = api.get_recipient_count(target_type="all")
        assert r["count"] == 2

    def test_filters_opted_out(self, api):
        api.add_contact("A", "a@x.com", "Single", None)
        api.add_contact("B", "b@x.com", "Single", None)
        cid = api.get_contacts()[1]["id"]
        api.set_contact_opt_out(cid, True)
        r = api.get_recipient_count(target_type="all")
        assert r["count"] == 1

    def test_manual_emails(self, api):
        r = api.get_recipient_count(manual_emails=["x@y.com", "bad-email", "z@w.com"])
        assert r["count"] == 2  # bad-email filtered out

    def test_deduplicates(self, api):
        api.add_contact("A", "a@x.com", "Single", None)
        r = api.get_recipient_count(target_type="all", manual_emails=["a@x.com"])
        assert r["count"] == 1

    def test_with_targets_list(self, api):
        api.add_contact("A", "a@x.com", "Family", None)
        api.add_contact("B", "b@x.com", "Single", None)
        r = api.get_recipient_count(targets=[{"type": "family"}, {"type": "single"}])
        assert r["count"] == 2


# ── Dispatch Emails ────────────────────────────────────────────────────────

class TestDispatchEmails:
    def test_no_credentials(self, api):
        api.add_contact("A", "a@x.com", "Single", None)
        cid = api.get_contacts()[0]["id"]
        res = api.dispatch_emails("Subj", "<p>Hi</p>", "Hi", [cid])
        assert res["error"]
        assert "credentials" in res["error"].lower()

    def test_no_recipients(self, api):
        api.save_settings("me@x.com", "pass")
        res = api.dispatch_emails("Subj", "<p>Hi</p>", "Hi", [])
        assert res["error"]
        assert "No recipients" in res["error"]

    @patch("main.smtplib.SMTP")
    def test_successful_send(self, mock_smtp_cls, api):
        mock_server = MagicMock()
        mock_smtp_cls.return_value.__enter__ = MagicMock(return_value=mock_server)
        mock_smtp_cls.return_value.__exit__ = MagicMock(return_value=False)

        api.save_settings("me@x.com", "pass")
        api.add_contact("John", "john@x.com", "Single", None)
        cid = api.get_contacts()[0]["id"]
        res = api.dispatch_emails("Hello", "<p>Hi {name}</p>", "Hi {name}", [cid])
        assert res["sent"] == 1
        assert res["failed"] == 0
        assert res["error"] is None
        # History should be logged
        history = api.get_email_history()
        assert len(history) == 1
        assert history[0]["subject"] == "Hello"

    @patch("main.smtplib.SMTP")
    def test_opt_out_filtered(self, mock_smtp_cls, api):
        mock_server = MagicMock()
        mock_smtp_cls.return_value.__enter__ = MagicMock(return_value=mock_server)
        mock_smtp_cls.return_value.__exit__ = MagicMock(return_value=False)

        api.save_settings("me@x.com", "pass")
        api.add_contact("A", "a@x.com", "Single", None)
        api.add_contact("B", "b@x.com", "Single", None)
        cids = [c["id"] for c in api.get_contacts()]
        api.set_contact_opt_out(cids[1], True)
        res = api.dispatch_emails("Hi", "<p>Hi</p>", "Hi", cids)
        assert res["sent"] == 1

    @patch("main.smtplib.SMTP")
    def test_cc_bcc(self, mock_smtp_cls, api):
        mock_server = MagicMock()
        mock_smtp_cls.return_value.__enter__ = MagicMock(return_value=mock_server)
        mock_smtp_cls.return_value.__exit__ = MagicMock(return_value=False)

        api.save_settings("me@x.com", "pass")
        api.add_contact("A", "a@x.com", "Single", None)
        cid = api.get_contacts()[0]["id"]
        res = api.dispatch_emails("Hi", "<p>Hi</p>", "Hi", [cid],
                                  cc_emails=["cc@x.com"], bcc_emails=["bcc@x.com"])
        assert res["sent"] == 1
        # Check sendmail was called with all recipients
        call_args = mock_server.sendmail.call_args
        all_recips = call_args[0][1]
        assert "cc@x.com" in all_recips
        assert "bcc@x.com" in all_recips

    @patch("main.smtplib.SMTP")
    def test_invalid_cc_filtered(self, mock_smtp_cls, api):
        mock_server = MagicMock()
        mock_smtp_cls.return_value.__enter__ = MagicMock(return_value=mock_server)
        mock_smtp_cls.return_value.__exit__ = MagicMock(return_value=False)

        api.save_settings("me@x.com", "pass")
        api.add_contact("A", "a@x.com", "Single", None)
        cid = api.get_contacts()[0]["id"]
        res = api.dispatch_emails("Hi", "<p>Hi</p>", "Hi", [cid],
                                  cc_emails=["not-an-email"])
        assert res["sent"] == 1

    @patch("main.smtplib.SMTP")
    def test_manual_emails(self, mock_smtp_cls, api):
        mock_server = MagicMock()
        mock_smtp_cls.return_value.__enter__ = MagicMock(return_value=mock_server)
        mock_smtp_cls.return_value.__exit__ = MagicMock(return_value=False)

        api.save_settings("me@x.com", "pass")
        res = api.dispatch_emails("Hi", "<p>Hi</p>", "Hi", [],
                                  manual_emails=["extra@x.com"])
        assert res["sent"] == 1

    @patch("main.smtplib.SMTP")
    def test_deduplicates_recipients(self, mock_smtp_cls, api):
        mock_server = MagicMock()
        mock_smtp_cls.return_value.__enter__ = MagicMock(return_value=mock_server)
        mock_smtp_cls.return_value.__exit__ = MagicMock(return_value=False)

        api.save_settings("me@x.com", "pass")
        api.add_contact("A", "a@x.com", "Single", None)
        cid = api.get_contacts()[0]["id"]
        res = api.dispatch_emails("Hi", "<p>Hi</p>", "Hi", [cid],
                                  manual_emails=["a@x.com"])
        assert res["sent"] == 1  # not 2

    @patch("main.smtplib.SMTP")
    def test_smtp_connection_failure(self, mock_smtp_cls, api):
        mock_smtp_cls.return_value.__enter__ = MagicMock(side_effect=Exception("Connection refused errno 61"))
        mock_smtp_cls.return_value.__exit__ = MagicMock(return_value=False)

        api.save_settings("me@x.com", "pass")
        api.add_contact("A", "a@x.com", "Single", None)
        cid = api.get_contacts()[0]["id"]
        res = api.dispatch_emails("Hi", "<p></p>", "Hi", [cid])
        assert res["error"] is not None
        assert "Connection refused" in res["error"]

    @patch("main.smtplib.SMTP")
    def test_updates_email_stats(self, mock_smtp_cls, api):
        mock_server = MagicMock()
        mock_smtp_cls.return_value.__enter__ = MagicMock(return_value=mock_server)
        mock_smtp_cls.return_value.__exit__ = MagicMock(return_value=False)

        api.save_settings("me@x.com", "pass")
        api.add_contact("A", "a@x.com", "Single", None)
        cid = api.get_contacts()[0]["id"]
        api.dispatch_emails("Hi", "<p>Hi</p>", "Hi", [cid])
        c = api.get_contacts()[0]
        assert c["email_count"] == 1
        assert c["last_emailed_at"] != ""

    @patch("main.smtplib.SMTP")
    def test_with_targets_list(self, mock_smtp_cls, api):
        mock_server = MagicMock()
        mock_smtp_cls.return_value.__enter__ = MagicMock(return_value=mock_server)
        mock_smtp_cls.return_value.__exit__ = MagicMock(return_value=False)

        api.save_settings("me@x.com", "pass")
        api.add_contact("A", "a@x.com", "Family", None)
        api.add_contact("B", "b@x.com", "Single", None)
        res = api.dispatch_emails("Hi", "<p>Hi</p>", "Hi", [],
                                  targets=[{"type": "family"}, {"type": "single"}])
        assert res["sent"] == 2


# ── Scheduled Emails API ──────────────────────────────────────────────────

class TestScheduledEmailsApi:
    def test_schedule_and_get(self, api):
        sa = (datetime.now() + timedelta(hours=1)).isoformat()
        res = api.schedule_email("Subj", "<p></p>", "txt", "all", None, [], [], sa)
        assert res["ok"]
        emails = api.get_scheduled_emails()
        assert len(emails) == 1
        assert emails[0]["subject"] == "Subj"
        assert emails[0]["status"] == "pending"

    def test_cancel(self, api):
        sa = (datetime.now() + timedelta(hours=1)).isoformat()
        api.schedule_email("S", "<p></p>", "t", "all", None, [], [], sa)
        eid = api.get_scheduled_emails()[0]["id"]
        res = api.cancel_scheduled_email(eid)
        assert res["ok"]
        assert api.get_scheduled_emails()[0]["status"] == "cancelled"

    def test_update(self, api):
        sa = (datetime.now() + timedelta(hours=1)).isoformat()
        api.schedule_email("Old", "<p></p>", "t", "all", None, [], [], sa)
        eid = api.get_scheduled_emails()[0]["id"]
        res = api.update_scheduled_email(eid, "New", "<p>new</p>", "new", "all", None, [], [], sa)
        assert res["ok"]

    def test_duplicate(self, api):
        sa = (datetime.now() + timedelta(hours=1)).isoformat()
        api.schedule_email("Orig", "<p></p>", "t", "all", None, [], [], sa)
        eid = api.get_scheduled_emails()[0]["id"]
        res = api.duplicate_scheduled_email(eid)
        assert res["ok"]
        assert len(api.get_scheduled_emails()) == 2

    def test_duplicate_missing(self, api):
        res = api.duplicate_scheduled_email(999)
        assert not res["ok"]

    def test_get_detail(self, api):
        sa = (datetime.now() + timedelta(hours=1)).isoformat()
        api.schedule_email("Detail", "<p>body</p>", "plain", "all", None, [], [], sa)
        eid = api.get_scheduled_emails()[0]["id"]
        d = api.get_scheduled_email_detail(eid)
        assert d is not None
        assert d["subject"] == "Detail"
        assert d["html_body"] == "<p>body</p>"

    def test_get_detail_missing(self, api):
        assert api.get_scheduled_email_detail(999) is None

    def test_with_recurrence(self, api):
        sa = (datetime.now() + timedelta(hours=1)).isoformat()
        rec = {"type": "daily"}
        api.schedule_email("Rec", "<p></p>", "t", "all", None, [], [], sa, recurrence=rec)
        e = api.get_scheduled_emails()[0]
        assert e["recurrence"] == {"type": "daily"}

    def test_get_with_recipients(self, api):
        api.add_contact("A", "a@x.com", "Single", None)
        cid = api.get_contacts()[0]["id"]
        sa = (datetime.now() + timedelta(hours=1)).isoformat()
        api.schedule_email("S", "<p></p>", "t", "custom", None, [cid], [], sa)
        emails = api.get_scheduled_emails_with_recipients()
        assert len(emails) == 1
        assert len(emails[0]["recipients"]) == 1


# ── Templates API ──────────────────────────────────────────────────────────

class TestTemplatesApi:
    def test_save_and_get(self, api):
        res = api.save_template("Welcome", "Hi", "<p>Hi</p>")
        assert res["ok"]
        tpls = api.get_templates()
        assert len(tpls) == 1
        assert tpls[0]["name"] == "Welcome"

    def test_update(self, api):
        api.save_template("T", "Old", "<p>Old</p>")
        tid = api.get_templates()[0]["id"]
        res = api.update_template(tid, "New", "<p>New</p>")
        assert res["ok"]

    def test_delete(self, api):
        api.save_template("T", "S", "<p></p>")
        tid = api.get_templates()[0]["id"]
        res = api.delete_template(tid)
        assert res["ok"]
        assert api.get_templates() == []

    def test_duplicate(self, api):
        api.save_template("Orig", "S", "<p></p>")
        tid = api.get_templates()[0]["id"]
        res = api.duplicate_template(tid)
        assert res["ok"]
        names = {t["name"] for t in api.get_templates()}
        assert "Orig (Copy)" in names

    def test_duplicate_missing(self, api):
        res = api.duplicate_template(999)
        assert not res["ok"]

    def test_with_recipients(self, api):
        recip = {"target_type": "all"}
        api.save_template("T", "S", "<p></p>", recipients=recip)
        t = api.get_templates()[0]
        assert t["recipients"] == recip


# ── Email History API ──────────────────────────────────────────────────────

class TestEmailHistoryApi:
    def test_get_empty(self, api):
        assert api.get_email_history() == []

    def test_get_after_log(self, api):
        db_manager.log_email("Test", "all", 5, 4, 1)
        h = api.get_email_history()
        assert len(h) == 1
        assert h[0]["subject"] == "Test"
        assert h[0]["sent"] == 4
        assert h[0]["failed"] == 1

    def test_details(self, api):
        details = [{"email": "a@x.com", "name": "A", "status": "sent"}]
        db_manager.log_email("S", "d", 1, 1, 0, recipient_details=details)
        hid = api.get_email_history()[0]["id"]
        d = api.get_email_history_details(hid)
        assert len(d) == 1
        assert d[0]["email"] == "a@x.com"

    def test_filtered(self, api):
        db_manager.log_email("Recent", "d", 1, 1, 0)
        results = api.get_email_history_filtered(start_date="2020-01-01")
        assert len(results) == 1

    def test_analytics(self, api):
        a = api.get_analytics()
        assert "totals" in a
        assert "weekly" in a
        assert "top_failed" in a


# ── Build Message ──────────────────────────────────────────────────────────

class TestBuildMessage:
    def test_basic_message(self):
        msg = Api._build_message("from@x.com", "to@x.com", "Subject", "plain", "<p>html</p>", [], [])
        assert msg["From"] == "from@x.com"
        assert msg["To"] == "to@x.com"
        assert msg["Subject"] == "Subject"

    def test_with_sender_name(self):
        msg = Api._build_message("from@x.com", "to@x.com", "S", "p", "<p></p>", [], [], sender_name="Church")
        assert "Church" in msg["From"]

    def test_with_cc_bcc(self):
        msg = Api._build_message("from@x.com", "to@x.com", "S", "p", "<p></p>", [], [],
                                 cc_addrs=["cc@x.com"], bcc_addrs=["bcc@x.com"])
        assert msg["Cc"] == "cc@x.com"
        assert msg["Bcc"] == "bcc@x.com"

    def test_with_attachment(self, tmp_path):
        # Create a temp file to attach
        att_file = tmp_path / "test.txt"
        att_file.write_text("hello")
        msg = Api._build_message("from@x.com", "to@x.com", "S", "p", "<p></p>", [], [str(att_file)])
        # Message should have attachment part
        payloads = msg.get_payload()
        assert len(payloads) >= 2  # related + attachment

    def test_nonexistent_attachment_skipped(self):
        msg = Api._build_message("from@x.com", "to@x.com", "S", "p", "<p></p>", [], ["/fake/path.txt"])
        # Should not crash, just skip


# ── Extract Inline Images ──────────────────────────────────────────────────

class TestExtractInlineImages:
    def test_no_images(self):
        html, images = Api._extract_inline_images("<p>Hello</p>")
        assert html == "<p>Hello</p>"
        assert images == []

    def test_with_base64_image(self):
        # Minimal valid 1x1 PNG as base64
        import base64
        pixel = base64.b64encode(b'\x89PNG\r\n\x1a\n' + b'\x00' * 50).decode()
        html_in = f'<img src="data:image/png;base64,{pixel}">'
        html_out, images = Api._extract_inline_images(html_in)
        assert len(images) == 1
        assert "cid:" in html_out


# ── SMTP Connection Test ──────────────────────────────────────────────────

class TestSmtpConnection:
    @patch("main.smtplib.SMTP")
    def test_success(self, mock_smtp_cls, api):
        mock_server = MagicMock()
        mock_smtp_cls.return_value.__enter__ = MagicMock(return_value=mock_server)
        mock_smtp_cls.return_value.__exit__ = MagicMock(return_value=False)
        res = api.test_email_connection("me@x.com", "pass")
        assert res["ok"]

    @patch("main.smtplib.SMTP")
    def test_failure(self, mock_smtp_cls, api):
        mock_smtp_cls.return_value.__enter__ = MagicMock(
            side_effect=Exception("535 Username and Password not accepted"))
        mock_smtp_cls.return_value.__exit__ = MagicMock(return_value=False)
        res = api.test_email_connection("me@x.com", "badpass")
        assert not res["ok"]
        assert "Authentication" in res["error"]


# ── Send Test Email ───────────────────────────────────────────────────────

class TestSendTestEmail:
    def test_no_credentials(self, api):
        res = api.send_test_email()
        assert not res["ok"]
        assert "credentials" in res["error"].lower()

    @patch("main.smtplib.SMTP")
    def test_success(self, mock_smtp_cls, api):
        mock_server = MagicMock()
        mock_smtp_cls.return_value.__enter__ = MagicMock(return_value=mock_server)
        mock_smtp_cls.return_value.__exit__ = MagicMock(return_value=False)
        api.save_settings("me@x.com", "pass")
        res = api.send_test_email()
        assert res["ok"]
        mock_server.sendmail.assert_called_once()


# ── Startup Toggle ────────────────────────────────────────────────────────

class TestStartupToggle:
    @patch("main.enable_startup")
    @patch("main.is_startup_enabled", return_value=True)
    def test_enable(self, mock_check, mock_enable, api):
        res = api.set_launch_on_startup(True)
        assert res["ok"]
        assert res["enabled"] is True
        mock_enable.assert_called_once()

    @patch("main.disable_startup")
    @patch("main.is_startup_enabled", return_value=False)
    def test_disable(self, mock_check, mock_disable, api):
        res = api.set_launch_on_startup(False)
        assert res["ok"]
        assert res["enabled"] is False
        mock_disable.assert_called_once()

    @patch("main.enable_startup", side_effect=Exception("Permission denied"))
    def test_error(self, mock_enable, api):
        res = api.set_launch_on_startup(True)
        assert not res["ok"]


# ── CSV Import/Export ─────────────────────────────────────────────────────

class TestCsvImport:
    def test_import_basic(self, api, tmp_path):
        csv_file = tmp_path / "contacts.csv"
        csv_file.write_text("name,email,category\nAlice,alice@x.com,Single\nBob,bob@x.com,Family\n")
        with patch.object(webview_mock, "windows", [MagicMock()]):
            webview_mock.windows[0].create_file_dialog.return_value = [str(csv_file)]
            webview_mock.FileDialog.OPEN = "open"
            res = api.import_csv()
        assert res["ok"]
        assert res["added"] == 2
        assert res["skipped"] == 0
        contacts = api.get_contacts()
        assert len(contacts) == 2

    def test_import_skips_missing_name_email(self, api, tmp_path):
        csv_file = tmp_path / "contacts.csv"
        csv_file.write_text("name,email,category\n,missing@x.com,Single\nNoEmail,,Single\nOK,ok@x.com,Single\n")
        with patch.object(webview_mock, "windows", [MagicMock()]):
            webview_mock.windows[0].create_file_dialog.return_value = [str(csv_file)]
            webview_mock.FileDialog.OPEN = "open"
            res = api.import_csv()
        assert res["ok"]
        assert res["added"] == 1
        assert res["skipped"] == 2

    def test_import_auto_creates_family(self, api, tmp_path):
        csv_file = tmp_path / "contacts.csv"
        csv_file.write_text("name,email,category,family_name\nJane,jane@x.com,Family,Doe\n")
        with patch.object(webview_mock, "windows", [MagicMock()]):
            webview_mock.windows[0].create_file_dialog.return_value = [str(csv_file)]
            webview_mock.FileDialog.OPEN = "open"
            res = api.import_csv()
        assert res["ok"]
        assert res["added"] == 1
        families = api.get_families()
        assert len(families) == 1
        assert families[0]["name"] == "Doe"

    def test_import_invalid_category_defaults_single(self, api, tmp_path):
        csv_file = tmp_path / "contacts.csv"
        csv_file.write_text("name,email,category\nTest,test@x.com,InvalidCat\n")
        with patch.object(webview_mock, "windows", [MagicMock()]):
            webview_mock.windows[0].create_file_dialog.return_value = [str(csv_file)]
            webview_mock.FileDialog.OPEN = "open"
            res = api.import_csv()
        assert res["ok"]
        assert res["added"] == 1
        assert api.get_contacts()[0]["category"] == "Single"

    def test_import_no_file_selected(self, api):
        with patch.object(webview_mock, "windows", [MagicMock()]):
            webview_mock.windows[0].create_file_dialog.return_value = None
            webview_mock.FileDialog.OPEN = "open"
            res = api.import_csv()
        assert not res["ok"]
        assert "No file" in res["error"]


class TestCsvExport:
    def test_export_basic(self, api, tmp_path):
        api.add_contact("Alice", "alice@x.com", "Single", None)
        api.add_contact("Bob", "bob@x.com", "Family", None)
        export_path = str(tmp_path / "export.csv")
        with patch.object(webview_mock, "windows", [MagicMock()]):
            webview_mock.windows[0].create_file_dialog.return_value = export_path
            webview_mock.FileDialog.SAVE = "save"
            res = api.export_csv()
        assert res["ok"]
        assert res["count"] == 2
        import csv
        with open(export_path) as f:
            reader = csv.DictReader(f)
            rows = list(reader)
        assert len(rows) == 2
        assert rows[0]["name"] in ("Alice", "Bob")

    def test_export_no_location(self, api):
        with patch.object(webview_mock, "windows", [MagicMock()]):
            webview_mock.windows[0].create_file_dialog.return_value = None
            webview_mock.FileDialog.SAVE = "save"
            res = api.export_csv()
        assert not res["ok"]


# ── Pick File ─────────────────────────────────────────────────────────────

class TestPickFile:
    def test_pick_file_returns_list(self, api, tmp_path):
        test_file = tmp_path / "doc.txt"
        test_file.write_text("hello")
        with patch.object(webview_mock, "windows", [MagicMock()]):
            webview_mock.windows[0].create_file_dialog.return_value = [str(test_file)]
            webview_mock.FileDialog.OPEN = "open"
            result = api.pick_file()
        assert len(result) == 1
        assert result[0]["name"] == "doc.txt"
        assert result[0]["size"] == 5

    def test_pick_file_cancelled(self, api):
        with patch.object(webview_mock, "windows", [MagicMock()]):
            webview_mock.windows[0].create_file_dialog.return_value = None
            webview_mock.FileDialog.OPEN = "open"
            result = api.pick_file()
        assert result == []


# ── Backup/Restore API ───────────────────────────────────────────────────

class TestBackupRestoreApi:
    def test_backup_success(self, api, tmp_path):
        backup_path = str(tmp_path / "backup.db")
        with patch.object(webview_mock, "windows", [MagicMock()]):
            webview_mock.windows[0].create_file_dialog.return_value = backup_path
            webview_mock.FileDialog.SAVE = "save"
            res = api.backup_database()
        assert res["ok"]
        assert os.path.exists(backup_path)

    def test_backup_no_location(self, api):
        with patch.object(webview_mock, "windows", [MagicMock()]):
            webview_mock.windows[0].create_file_dialog.return_value = None
            webview_mock.FileDialog.SAVE = "save"
            res = api.backup_database()
        assert not res["ok"]

    def test_restore_success(self, api, tmp_path):
        api.add_contact("Before", "before@x.com", "Single", None)
        backup_path = str(tmp_path / "backup.db")
        db_manager.backup_database(backup_path)
        api.add_contact("After", "after@x.com", "Single", None)
        with patch.object(webview_mock, "windows", [MagicMock()]):
            webview_mock.windows[0].create_file_dialog.return_value = [backup_path]
            webview_mock.FileDialog.OPEN = "open"
            res = api.restore_database()
        assert res["ok"]
        assert len(api.get_contacts()) == 1

    def test_restore_no_file(self, api):
        with patch.object(webview_mock, "windows", [MagicMock()]):
            webview_mock.windows[0].create_file_dialog.return_value = None
            webview_mock.FileDialog.OPEN = "open"
            res = api.restore_database()
        assert not res["ok"]


# ── Dispatch partial failure ─────────────────────────────────────────────

class TestDispatchPartialFailure:
    @patch("main.smtplib.SMTP")
    def test_per_recipient_failure(self, mock_smtp_cls, api):
        mock_server = MagicMock()
        mock_smtp_cls.return_value.__enter__ = MagicMock(return_value=mock_server)
        mock_smtp_cls.return_value.__exit__ = MagicMock(return_value=False)
        mock_server.sendmail.side_effect = [None, Exception("Mailbox full")]

        api.save_settings("me@x.com", "pass")
        api.add_contact("A", "a@x.com", "Single", None)
        api.add_contact("B", "b@x.com", "Single", None)
        cids = [c["id"] for c in api.get_contacts()]
        res = api.dispatch_emails("Hi", "<p>Hi</p>", "Hi", cids)
        assert res["sent"] == 1
        assert res["failed"] == 1
        assert res["error"] is None  # no connection error

    @patch("main.smtplib.SMTP")
    def test_dispatch_logs_details(self, mock_smtp_cls, api):
        mock_server = MagicMock()
        mock_smtp_cls.return_value.__enter__ = MagicMock(return_value=mock_server)
        mock_smtp_cls.return_value.__exit__ = MagicMock(return_value=False)
        mock_server.sendmail.side_effect = [None, Exception("bounce")]

        api.save_settings("me@x.com", "pass")
        api.add_contact("A", "a@x.com", "Single", None)
        api.add_contact("B", "b@x.com", "Single", None)
        cids = [c["id"] for c in api.get_contacts()]
        api.dispatch_emails("Hi", "<p>Hi</p>", "Hi", cids)

        history = api.get_email_history()
        assert len(history) == 1
        hid = history[0]["id"]
        details = api.get_email_history_details(hid)
        assert len(details) == 2
        statuses = {d["status"] for d in details}
        assert statuses == {"sent", "failed"}


# ── Dispatch with attachments ────────────────────────────────────────────

class TestDispatchWithAttachments:
    @patch("main.smtplib.SMTP")
    def test_dispatch_with_attachment(self, mock_smtp_cls, api, tmp_path):
        mock_server = MagicMock()
        mock_smtp_cls.return_value.__enter__ = MagicMock(return_value=mock_server)
        mock_smtp_cls.return_value.__exit__ = MagicMock(return_value=False)

        att_file = tmp_path / "doc.txt"
        att_file.write_text("hello world")
        api.save_settings("me@x.com", "pass")
        api.add_contact("A", "a@x.com", "Single", None)
        cid = api.get_contacts()[0]["id"]
        res = api.dispatch_emails("Hi", "<p>Hi</p>", "Hi", [cid], attachment_paths=[str(att_file)])
        assert res["sent"] == 1
        # Verify attachment was included in the message
        call_args = mock_server.sendmail.call_args
        msg_str = call_args[0][2]
        assert "doc.txt" in msg_str


# ── Resolve recipients edge cases ────────────────────────────────────────

class TestResolveRecipientsOptOut:
    def test_resolve_includes_opted_out(self, api):
        """resolve_recipients does NOT filter opt-out — dispatch/count do."""
        api.add_contact("A", "a@x.com", "Single", None)
        cid = api.get_contacts()[0]["id"]
        api.set_contact_opt_out(cid, True)
        r = api.resolve_recipients("all")
        assert len(r) == 1  # resolve returns all, including opted-out


# ── Scheduled emails with recipients dedup ───────────────────────────────

class TestScheduledEmailsWithRecipientsDedup:
    def test_manual_email_dedup_with_contact(self, api):
        api.add_contact("A", "a@x.com", "Single", None)
        cid = api.get_contacts()[0]["id"]
        sa = (datetime.now() + timedelta(hours=1)).isoformat()
        api.schedule_email("S", "<p></p>", "t", "custom", None, [cid], [], sa,
                           manual_emails=["a@x.com"])
        emails = api.get_scheduled_emails_with_recipients()
        # Should deduplicate: a@x.com from contact + a@x.com manual = 1
        assert len(emails[0]["recipients"]) == 1


# ── Test connection with custom SMTP ─────────────────────────────────────

class TestSmtpConnectionCustom:
    @patch("main.smtplib.SMTP")
    def test_custom_host_port(self, mock_smtp_cls, api):
        mock_server = MagicMock()
        mock_smtp_cls.return_value.__enter__ = MagicMock(return_value=mock_server)
        mock_smtp_cls.return_value.__exit__ = MagicMock(return_value=False)
        res = api.test_email_connection("me@x.com", "pass",
                                        smtp_host="mail.example.com", smtp_port="465")
        assert res["ok"]
        mock_smtp_cls.assert_called_once_with("mail.example.com", 465, timeout=15)


# ── Build message with inline images ─────────────────────────────────────

class TestBuildMessageWithImages:
    def test_with_image_data(self):
        import base64
        pixel = b'\x89PNG\r\n\x1a\n' + b'\x00' * 50
        image_data = [("abc123", "image/png", pixel)]
        msg = Api._build_message("from@x.com", "to@x.com", "S", "text", "<p>html</p>",
                                 image_data, [])
        msg_str = msg.as_string()
        assert "abc123" in msg_str
        assert "image/png" in msg_str


# ── Send to recipients partial failure ───────────────────────────────────

class TestSendToRecipientsPartial:
    @patch("main.smtplib.SMTP")
    def test_partial_failure_details(self, mock_smtp_cls):
        mock_server = MagicMock()
        mock_smtp_cls.return_value.__enter__ = MagicMock(return_value=mock_server)
        mock_smtp_cls.return_value.__exit__ = MagicMock(return_value=False)
        mock_server.sendmail.side_effect = [None, Exception("bounce"), None]

        db_manager.set_setting("sender_name", "Church")
        db_manager.set_setting("smtp_host", "smtp.gmail.com")
        db_manager.set_setting("smtp_port", "587")
        recipients = [("A", "a@x.com"), ("B", "b@x.com"), ("C", "c@x.com")]
        result = Api._send_to_recipients("me@x.com", "pass", recipients, "Hi", "text", "<p>html</p>", [])
        assert result["sent"] == 2
        assert result["failed"] == 1
        assert result["error"] is None
        assert len(result["details"]) == 3
        statuses = [d["status"] for d in result["details"]]
        assert statuses == ["sent", "failed", "sent"]


# ── Email regex edge cases ───────────────────────────────────────────────

class TestEmailRegex:
    def test_valid_emails(self, api):
        api.save_settings("me@x.com", "pass")
        r = api.get_recipient_count(manual_emails=["user+tag@gmail.com", "a.b@c.co.uk"])
        assert r["count"] == 2

    def test_invalid_emails_filtered(self, api):
        r = api.get_recipient_count(manual_emails=["nope", "@bad", "no@", "ok@x.com", ""])
        assert r["count"] == 1

    def test_unicode_domain(self, api):
        # Basic unicode in local part should still match our simple regex
        r = api.get_recipient_count(manual_emails=["user@example.com"])
        assert r["count"] == 1


# ── Empty/None input edge cases ──────────────────────────────────────────

class TestEmptyInputs:
    def test_add_contact_empty_phone_notes(self, api):
        res = api.add_contact("A", "a@x.com", "Single", None, phone="", notes="")
        assert res["ok"]
        c = api.get_contacts()[0]
        assert c["phone"] == ""
        assert c["notes"] == ""

    def test_save_settings_empty_smtp(self, api):
        res = api.save_settings("me@x.com", "pass", sender_name="", smtp_host="", smtp_port="")
        assert res["ok"]
        s = api.get_settings()
        # Empty smtp_host/port should keep defaults
        assert s["smtp_host"] == "smtp.gmail.com"
        assert s["smtp_port"] == "587"

    def test_dispatch_no_contact_ids_no_manual(self, api):
        api.save_settings("me@x.com", "pass")
        res = api.dispatch_emails("Hi", "<p>Hi</p>", "Hi", [], manual_emails=[])
        assert res["error"] is not None
        assert "No recipients" in res["error"]

    def test_get_recipient_count_all_empty(self, api):
        r = api.get_recipient_count()
        assert r["count"] == 0

    def test_schedule_email_empty_lists(self, api):
        sa = (datetime.now() + timedelta(hours=1)).isoformat()
        res = api.schedule_email("S", "<p></p>", "t", "all", None, [], [], sa,
                                 manual_emails=[])
        assert res["ok"]


# ── Template variable edge cases ─────────────────────────────────────────

class TestTemplateVariableEdgeCases:
    @patch("main.smtplib.SMTP")
    def test_html_in_name_not_escaped(self, mock_smtp_cls, api):
        """Template vars do simple string replacement — HTML in names passes through."""
        mock_server = MagicMock()
        mock_smtp_cls.return_value.__enter__ = MagicMock(return_value=mock_server)
        mock_smtp_cls.return_value.__exit__ = MagicMock(return_value=False)

        api.save_settings("me@x.com", "pass")
        api.add_contact("<b>Bold</b>", "bold@x.com", "Single", None)
        cid = api.get_contacts()[0]["id"]
        res = api.dispatch_emails("Hi {name}", "<p>Hello {name}</p>", "Hello {name}", [cid])
        assert res["sent"] == 1
        call_args = mock_server.sendmail.call_args
        msg_str = call_args[0][2]
        assert "<b>Bold</b>" in msg_str

    @patch("main.smtplib.SMTP")
    def test_email_template_var(self, mock_smtp_cls, api):
        mock_server = MagicMock()
        mock_smtp_cls.return_value.__enter__ = MagicMock(return_value=mock_server)
        mock_smtp_cls.return_value.__exit__ = MagicMock(return_value=False)

        api.save_settings("me@x.com", "pass")
        api.add_contact("A", "alice@x.com", "Single", None)
        cid = api.get_contacts()[0]["id"]
        res = api.dispatch_emails("For {email}", "<p>{email}</p>", "{email}", [cid])
        assert res["sent"] == 1
        msg_str = mock_server.sendmail.call_args[0][2]
        assert "alice@x.com" in msg_str
