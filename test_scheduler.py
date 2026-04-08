"""Tests for the background scheduler (run_scheduler) and edge cases."""
import json
import os
import pytest
from unittest.mock import patch, MagicMock, call
from datetime import datetime, timedelta

import db_manager

import sys
sys.modules["webview"] = MagicMock()
sys.modules["pystray"] = MagicMock()

import main
from main import Api, run_scheduler


@pytest.fixture(autouse=True)
def tmp_db(tmp_path):
    db_file = str(tmp_path / "test.db")
    db_manager.DB_PATH = db_file
    db_manager.init_db()
    yield db_file


@pytest.fixture
def api():
    return Api()


def _schedule_due(subject="Test", target_type="all", contact_ids=None, recurrence=None, manual_emails=None):
    """Helper: schedule an email that is already due (1 hour ago)."""
    sa = (datetime.now() - timedelta(hours=1)).isoformat()
    db_manager.schedule_email(subject, "<p>Hi {name}</p>", "Hi {name}", target_type, None,
                              contact_ids or [], [], sa, recurrence=recurrence,
                              manual_emails=manual_emails)


class TestSchedulerBasic:
    @patch("main.time.sleep", side_effect=StopIteration)  # stop after one loop
    @patch("main.smtplib.SMTP")
    def test_sends_due_email(self, mock_smtp_cls, mock_sleep, api):
        mock_server = MagicMock()
        mock_smtp_cls.return_value.__enter__ = MagicMock(return_value=mock_server)
        mock_smtp_cls.return_value.__exit__ = MagicMock(return_value=False)

        db_manager.set_setting("sender_email", "me@x.com")
        db_manager.set_setting("app_password", "pass")
        db_manager.add_contact("John", "john@x.com", "Single")
        _schedule_due("Hello")

        with pytest.raises(StopIteration):
            run_scheduler(api)

        # Email should be sent
        mock_server.sendmail.assert_called_once()
        # Status should be updated
        emails = db_manager.get_scheduled_emails()
        assert emails[0][5] == "sent"
        assert emails[0][6] is not None  # sent_at

    @patch("main.time.sleep", side_effect=StopIteration)
    @patch("main.smtplib.SMTP")
    def test_skips_future_emails(self, mock_smtp_cls, mock_sleep, api):
        mock_server = MagicMock()
        mock_smtp_cls.return_value.__enter__ = MagicMock(return_value=mock_server)
        mock_smtp_cls.return_value.__exit__ = MagicMock(return_value=False)

        db_manager.set_setting("sender_email", "me@x.com")
        db_manager.set_setting("app_password", "pass")
        db_manager.add_contact("A", "a@x.com", "Single")

        # Schedule far enough in the future to be safe across any timezone
        sa = (datetime.now() + timedelta(days=30)).isoformat()
        db_manager.schedule_email("Future", "<p></p>", "t", "all", None, [], [], sa)

        with pytest.raises(StopIteration):
            run_scheduler(api)

        mock_server.sendmail.assert_not_called()
        assert db_manager.get_scheduled_emails()[0][5] == "pending"

    @patch("main.time.sleep", side_effect=StopIteration)
    def test_no_credentials(self, mock_sleep, api):
        db_manager.add_contact("A", "a@x.com", "Single")
        _schedule_due("NoCreds")

        with pytest.raises(StopIteration):
            run_scheduler(api)

        emails = db_manager.get_scheduled_emails()
        assert emails[0][5] == "failed"

    @patch("main.time.sleep", side_effect=StopIteration)
    def test_no_recipients(self, mock_sleep, api):
        db_manager.set_setting("sender_email", "me@x.com")
        db_manager.set_setting("app_password", "pass")
        # Schedule with target_type that has no contacts
        _schedule_due("Empty", target_type="group")

        with pytest.raises(StopIteration):
            run_scheduler(api)

        emails = db_manager.get_scheduled_emails()
        assert emails[0][5] == "failed"


class TestSchedulerRecurrence:
    @patch("main.time.sleep", side_effect=StopIteration)
    @patch("main.smtplib.SMTP")
    def test_creates_next_occurrence(self, mock_smtp_cls, mock_sleep, api):
        mock_server = MagicMock()
        mock_smtp_cls.return_value.__enter__ = MagicMock(return_value=mock_server)
        mock_smtp_cls.return_value.__exit__ = MagicMock(return_value=False)

        db_manager.set_setting("sender_email", "me@x.com")
        db_manager.set_setting("app_password", "pass")
        db_manager.add_contact("A", "a@x.com", "Single")
        _schedule_due("Recurring", recurrence={"type": "daily"})

        with pytest.raises(StopIteration):
            run_scheduler(api)

        emails = db_manager.get_scheduled_emails()
        assert len(emails) == 2  # original + next
        statuses = {e[5] for e in emails}
        assert "sent" in statuses
        assert "pending" in statuses

    @patch("main.time.sleep", side_effect=StopIteration)
    @patch("main.smtplib.SMTP")
    def test_once_does_not_recur(self, mock_smtp_cls, mock_sleep, api):
        mock_server = MagicMock()
        mock_smtp_cls.return_value.__enter__ = MagicMock(return_value=mock_server)
        mock_smtp_cls.return_value.__exit__ = MagicMock(return_value=False)

        db_manager.set_setting("sender_email", "me@x.com")
        db_manager.set_setting("app_password", "pass")
        db_manager.add_contact("A", "a@x.com", "Single")
        _schedule_due("OneTime", recurrence={"type": "once"})

        with pytest.raises(StopIteration):
            run_scheduler(api)

        assert len(db_manager.get_scheduled_emails()) == 1

    @patch("main.time.sleep", side_effect=StopIteration)
    @patch("main.smtplib.SMTP")
    def test_expired_recurrence_stops(self, mock_smtp_cls, mock_sleep, api):
        mock_server = MagicMock()
        mock_smtp_cls.return_value.__enter__ = MagicMock(return_value=mock_server)
        mock_smtp_cls.return_value.__exit__ = MagicMock(return_value=False)

        db_manager.set_setting("sender_email", "me@x.com")
        db_manager.set_setting("app_password", "pass")
        db_manager.add_contact("A", "a@x.com", "Single")
        # end_date is in the past → no next occurrence
        _schedule_due("Expired", recurrence={
            "type": "daily",
            "end_date": (datetime.now() - timedelta(days=1)).isoformat()
        })

        with pytest.raises(StopIteration):
            run_scheduler(api)

        assert len(db_manager.get_scheduled_emails()) == 1  # no next created


class TestSchedulerManualEmails:
    @patch("main.time.sleep", side_effect=StopIteration)
    @patch("main.smtplib.SMTP")
    def test_includes_manual_emails(self, mock_smtp_cls, mock_sleep, api):
        mock_server = MagicMock()
        mock_smtp_cls.return_value.__enter__ = MagicMock(return_value=mock_server)
        mock_smtp_cls.return_value.__exit__ = MagicMock(return_value=False)

        db_manager.set_setting("sender_email", "me@x.com")
        db_manager.set_setting("app_password", "pass")
        _schedule_due("Manual", target_type="manual", manual_emails=["extra@x.com"])

        with pytest.raises(StopIteration):
            run_scheduler(api)

        mock_server.sendmail.assert_called_once()
        emails = db_manager.get_scheduled_emails()
        assert emails[0][5] == "sent"

    @patch("main.time.sleep", side_effect=StopIteration)
    @patch("main.smtplib.SMTP")
    def test_deduplicates_manual_and_contact(self, mock_smtp_cls, mock_sleep, api):
        mock_server = MagicMock()
        mock_smtp_cls.return_value.__enter__ = MagicMock(return_value=mock_server)
        mock_smtp_cls.return_value.__exit__ = MagicMock(return_value=False)

        db_manager.set_setting("sender_email", "me@x.com")
        db_manager.set_setting("app_password", "pass")
        db_manager.add_contact("A", "a@x.com", "Single")
        _schedule_due("Dedup", manual_emails=["a@x.com"])

        with pytest.raises(StopIteration):
            run_scheduler(api)

        # Should only send once despite duplicate
        assert mock_server.sendmail.call_count == 1


class TestSchedulerTemplateVars:
    @patch("main.time.sleep", side_effect=StopIteration)
    @patch("main.smtplib.SMTP")
    def test_substitutes_name_and_email(self, mock_smtp_cls, mock_sleep, api):
        mock_server = MagicMock()
        mock_smtp_cls.return_value.__enter__ = MagicMock(return_value=mock_server)
        mock_smtp_cls.return_value.__exit__ = MagicMock(return_value=False)

        db_manager.set_setting("sender_email", "me@x.com")
        db_manager.set_setting("app_password", "pass")
        db_manager.add_contact("John Doe", "john@x.com", "Single")
        _schedule_due("Hi {name}")

        with pytest.raises(StopIteration):
            run_scheduler(api)

        # Check that the sent message contains "John Doe" in the subject
        call_args = mock_server.sendmail.call_args
        msg_str = call_args[0][2]  # third arg is the message string
        assert "John Doe" in msg_str


class TestSchedulerSmtpFailure:
    @patch("main.time.sleep", side_effect=StopIteration)
    @patch("main.smtplib.SMTP")
    def test_per_recipient_failure(self, mock_smtp_cls, mock_sleep, api):
        mock_server = MagicMock()
        mock_smtp_cls.return_value.__enter__ = MagicMock(return_value=mock_server)
        mock_smtp_cls.return_value.__exit__ = MagicMock(return_value=False)
        # First call succeeds, second fails
        mock_server.sendmail.side_effect = [None, Exception("Mailbox full")]

        db_manager.set_setting("sender_email", "me@x.com")
        db_manager.set_setting("app_password", "pass")
        db_manager.add_contact("A", "a@x.com", "Single")
        db_manager.add_contact("B", "b@x.com", "Single")
        _schedule_due("Partial")

        with pytest.raises(StopIteration):
            run_scheduler(api)

        emails = db_manager.get_scheduled_emails()
        result = json.loads(emails[0][7])
        assert result["sent"] == 1
        assert result["failed"] == 1

    @patch("main.time.sleep", side_effect=StopIteration)
    @patch("main.smtplib.SMTP")
    def test_total_smtp_failure(self, mock_smtp_cls, mock_sleep, api):
        mock_smtp_cls.return_value.__enter__ = MagicMock(
            side_effect=Exception("Connection refused"))
        mock_smtp_cls.return_value.__exit__ = MagicMock(return_value=False)

        db_manager.set_setting("sender_email", "me@x.com")
        db_manager.set_setting("app_password", "pass")
        db_manager.add_contact("A", "a@x.com", "Single")
        _schedule_due("ConnFail")

        with pytest.raises(StopIteration):
            run_scheduler(api)

        emails = db_manager.get_scheduled_emails()
        assert emails[0][5] == "failed"


class TestSchedulerTimezone:
    @patch("main.time.sleep", side_effect=StopIteration)
    @patch("main.smtplib.SMTP")
    def test_uses_configured_timezone(self, mock_smtp_cls, mock_sleep, api):
        mock_server = MagicMock()
        mock_smtp_cls.return_value.__enter__ = MagicMock(return_value=mock_server)
        mock_smtp_cls.return_value.__exit__ = MagicMock(return_value=False)

        db_manager.set_setting("sender_email", "me@x.com")
        db_manager.set_setting("app_password", "pass")
        db_manager.set_setting("timezone", "US/Pacific")
        db_manager.add_contact("A", "a@x.com", "Single")
        _schedule_due("TZ")

        with pytest.raises(StopIteration):
            run_scheduler(api)

        # Should still process (the email is overdue regardless of timezone)
        assert db_manager.get_scheduled_emails()[0][5] == "sent"

    @patch("main.time.sleep", side_effect=StopIteration)
    @patch("main.smtplib.SMTP")
    def test_invalid_timezone_falls_back(self, mock_smtp_cls, mock_sleep, api):
        mock_server = MagicMock()
        mock_smtp_cls.return_value.__enter__ = MagicMock(return_value=mock_server)
        mock_smtp_cls.return_value.__exit__ = MagicMock(return_value=False)

        db_manager.set_setting("sender_email", "me@x.com")
        db_manager.set_setting("app_password", "pass")
        db_manager.set_setting("timezone", "Invalid/Timezone")
        db_manager.add_contact("A", "a@x.com", "Single")
        _schedule_due("BadTZ")

        with pytest.raises(StopIteration):
            run_scheduler(api)

        # Falls back to US/Eastern, still processes
        assert db_manager.get_scheduled_emails()[0][5] == "sent"


class TestSchedulerHistoryLogging:
    @patch("main.time.sleep", side_effect=StopIteration)
    @patch("main.smtplib.SMTP")
    def test_logs_to_history(self, mock_smtp_cls, mock_sleep, api):
        mock_server = MagicMock()
        mock_smtp_cls.return_value.__enter__ = MagicMock(return_value=mock_server)
        mock_smtp_cls.return_value.__exit__ = MagicMock(return_value=False)

        db_manager.set_setting("sender_email", "me@x.com")
        db_manager.set_setting("app_password", "pass")
        db_manager.add_contact("A", "a@x.com", "Single")
        _schedule_due("Logged")

        with pytest.raises(StopIteration):
            run_scheduler(api)

        history = db_manager.get_email_history()
        assert len(history) == 1
        assert history[0][1] == "Logged"
        assert history[0][4] == 1  # sent_count

    @patch("main.time.sleep", side_effect=StopIteration)
    @patch("main.smtplib.SMTP")
    def test_logs_history_details(self, mock_smtp_cls, mock_sleep, api):
        mock_server = MagicMock()
        mock_smtp_cls.return_value.__enter__ = MagicMock(return_value=mock_server)
        mock_smtp_cls.return_value.__exit__ = MagicMock(return_value=False)

        db_manager.set_setting("sender_email", "me@x.com")
        db_manager.set_setting("app_password", "pass")
        db_manager.add_contact("A", "a@x.com", "Single")
        db_manager.add_contact("B", "b@x.com", "Single")
        _schedule_due("DetailLog")

        with pytest.raises(StopIteration):
            run_scheduler(api)

        history = db_manager.get_email_history()
        assert len(history) == 1
        hid = history[0][0]
        details = db_manager.get_email_history_details(hid)
        assert len(details) == 2


class TestSchedulerOptOut:
    @patch("main.time.sleep", side_effect=StopIteration)
    @patch("main.smtplib.SMTP")
    def test_opted_out_contacts_excluded(self, mock_smtp_cls, mock_sleep, api):
        """Scheduler should filter out opted-out contacts, matching dispatch_emails behavior."""
        mock_server = MagicMock()
        mock_smtp_cls.return_value.__enter__ = MagicMock(return_value=mock_server)
        mock_smtp_cls.return_value.__exit__ = MagicMock(return_value=False)

        db_manager.set_setting("sender_email", "me@x.com")
        db_manager.set_setting("app_password", "pass")
        db_manager.add_contact("A", "a@x.com", "Single")
        db_manager.add_contact("B", "b@x.com", "Single")
        # Opt out contact B
        contacts = db_manager.get_contacts()
        for c in contacts:
            if c[1] == "B":
                db_manager.set_contact_opt_out(c[0], True)
        _schedule_due("OptOutTest")

        with pytest.raises(StopIteration):
            run_scheduler(api)

        # Scheduler now filters opted-out contacts — only A should receive
        assert mock_server.sendmail.call_count == 1


class TestSchedulerRecurrenceTypes:
    @patch("main.time.sleep", side_effect=StopIteration)
    @patch("main.smtplib.SMTP")
    def test_weekly_recurrence(self, mock_smtp_cls, mock_sleep, api):
        mock_server = MagicMock()
        mock_smtp_cls.return_value.__enter__ = MagicMock(return_value=mock_server)
        mock_smtp_cls.return_value.__exit__ = MagicMock(return_value=False)

        db_manager.set_setting("sender_email", "me@x.com")
        db_manager.set_setting("app_password", "pass")
        db_manager.add_contact("A", "a@x.com", "Single")
        _schedule_due("Weekly", recurrence={"type": "weekly", "days": [2, 4]})

        with pytest.raises(StopIteration):
            run_scheduler(api)

        emails = db_manager.get_scheduled_emails()
        assert len(emails) == 2
        statuses = {e[5] for e in emails}
        assert "sent" in statuses
        assert "pending" in statuses

    @patch("main.time.sleep", side_effect=StopIteration)
    @patch("main.smtplib.SMTP")
    def test_monthly_recurrence(self, mock_smtp_cls, mock_sleep, api):
        mock_server = MagicMock()
        mock_smtp_cls.return_value.__enter__ = MagicMock(return_value=mock_server)
        mock_smtp_cls.return_value.__exit__ = MagicMock(return_value=False)

        db_manager.set_setting("sender_email", "me@x.com")
        db_manager.set_setting("app_password", "pass")
        db_manager.add_contact("A", "a@x.com", "Single")
        _schedule_due("Monthly", recurrence={"type": "monthly", "day_of_month": 15})

        with pytest.raises(StopIteration):
            run_scheduler(api)

        emails = db_manager.get_scheduled_emails()
        assert len(emails) == 2
        pending = [e for e in emails if e[5] == "pending"]
        assert len(pending) == 1

    @patch("main.time.sleep", side_effect=StopIteration)
    @patch("main.smtplib.SMTP")
    def test_every_other_day_recurrence(self, mock_smtp_cls, mock_sleep, api):
        mock_server = MagicMock()
        mock_smtp_cls.return_value.__enter__ = MagicMock(return_value=mock_server)
        mock_smtp_cls.return_value.__exit__ = MagicMock(return_value=False)

        db_manager.set_setting("sender_email", "me@x.com")
        db_manager.set_setting("app_password", "pass")
        db_manager.add_contact("A", "a@x.com", "Single")
        _schedule_due("EveryOther", recurrence={"type": "every_other_day"})

        with pytest.raises(StopIteration):
            run_scheduler(api)

        emails = db_manager.get_scheduled_emails()
        assert len(emails) == 2

    @patch("main.time.sleep", side_effect=StopIteration)
    @patch("main.smtplib.SMTP")
    def test_every_other_week_recurrence(self, mock_smtp_cls, mock_sleep, api):
        mock_server = MagicMock()
        mock_smtp_cls.return_value.__enter__ = MagicMock(return_value=mock_server)
        mock_smtp_cls.return_value.__exit__ = MagicMock(return_value=False)

        db_manager.set_setting("sender_email", "me@x.com")
        db_manager.set_setting("app_password", "pass")
        db_manager.add_contact("A", "a@x.com", "Single")
        _schedule_due("BiWeekly", recurrence={"type": "every_other_week", "days": [3]})

        with pytest.raises(StopIteration):
            run_scheduler(api)

        emails = db_manager.get_scheduled_emails()
        assert len(emails) == 2


class TestSchedulerMultipleDue:
    @patch("main.time.sleep", side_effect=StopIteration)
    @patch("main.smtplib.SMTP")
    def test_processes_all_due_emails(self, mock_smtp_cls, mock_sleep, api):
        mock_server = MagicMock()
        mock_smtp_cls.return_value.__enter__ = MagicMock(return_value=mock_server)
        mock_smtp_cls.return_value.__exit__ = MagicMock(return_value=False)

        db_manager.set_setting("sender_email", "me@x.com")
        db_manager.set_setting("app_password", "pass")
        db_manager.add_contact("A", "a@x.com", "Single")
        _schedule_due("Email1")
        _schedule_due("Email2")
        _schedule_due("Email3")

        with pytest.raises(StopIteration):
            run_scheduler(api)

        emails = db_manager.get_scheduled_emails()
        sent = [e for e in emails if e[5] == "sent"]
        assert len(sent) == 3
        assert mock_server.sendmail.call_count == 3


class TestSchedulerPartialFailureStatus:
    @patch("main.time.sleep", side_effect=StopIteration)
    @patch("main.smtplib.SMTP")
    def test_mixed_results_status_sent(self, mock_smtp_cls, mock_sleep, api):
        """When some recipients succeed and some fail, status should be 'sent' (not 'failed')."""
        mock_server = MagicMock()
        mock_smtp_cls.return_value.__enter__ = MagicMock(return_value=mock_server)
        mock_smtp_cls.return_value.__exit__ = MagicMock(return_value=False)
        mock_server.sendmail.side_effect = [None, Exception("bounce"), None]

        db_manager.set_setting("sender_email", "me@x.com")
        db_manager.set_setting("app_password", "pass")
        db_manager.add_contact("A", "a@x.com", "Single")
        db_manager.add_contact("B", "b@x.com", "Single")
        db_manager.add_contact("C", "c@x.com", "Single")
        _schedule_due("Mixed")

        with pytest.raises(StopIteration):
            run_scheduler(api)

        emails = db_manager.get_scheduled_emails()
        assert emails[0][5] == "sent"  # not "failed" because some succeeded
        result = json.loads(emails[0][7])
        assert result["sent"] == 2
        assert result["failed"] == 1

    @patch("main.time.sleep", side_effect=StopIteration)
    @patch("main.smtplib.SMTP")
    def test_all_recipients_fail_status_failed(self, mock_smtp_cls, mock_sleep, api):
        mock_server = MagicMock()
        mock_smtp_cls.return_value.__enter__ = MagicMock(return_value=mock_server)
        mock_smtp_cls.return_value.__exit__ = MagicMock(return_value=False)
        mock_server.sendmail.side_effect = Exception("all fail")

        db_manager.set_setting("sender_email", "me@x.com")
        db_manager.set_setting("app_password", "pass")
        db_manager.add_contact("A", "a@x.com", "Single")
        _schedule_due("AllFail")

        with pytest.raises(StopIteration):
            run_scheduler(api)

        emails = db_manager.get_scheduled_emails()
        assert emails[0][5] == "failed"


class TestSchedulerWithAttachments:
    @patch("main.time.sleep", side_effect=StopIteration)
    @patch("main.smtplib.SMTP")
    def test_sends_with_attachments(self, mock_smtp_cls, mock_sleep, api, tmp_path):
        mock_server = MagicMock()
        mock_smtp_cls.return_value.__enter__ = MagicMock(return_value=mock_server)
        mock_smtp_cls.return_value.__exit__ = MagicMock(return_value=False)

        att_file = tmp_path / "report.txt"
        att_file.write_text("data")

        db_manager.set_setting("sender_email", "me@x.com")
        db_manager.set_setting("app_password", "pass")
        db_manager.add_contact("A", "a@x.com", "Single")

        sa = (datetime.now() - timedelta(hours=1)).isoformat()
        db_manager.schedule_email("WithAttach", "<p>Hi</p>", "Hi", "all", None,
                                  [], [str(att_file)], sa)

        with pytest.raises(StopIteration):
            run_scheduler(api)

        assert mock_server.sendmail.call_count == 1
        msg_str = mock_server.sendmail.call_args[0][2]
        assert "report.txt" in msg_str
