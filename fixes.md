# Fixes Applied

## Security

### 1. Plaintext credential storage
- **File:** `db_manager.py`
- **Issue:** Email App Password stored as plaintext in SQLite `settings` table.
- **Fix:** Not addressed in code (requires `keyring` dependency). Documented as known limitation.

### 2. Password input field visible
- **File:** `main.py` (HTML)
- **Issue:** Settings password input used `type="text"`, making password visible on screen.
- **Fix:** Changed to `type="password"`.

### 3. No email validation on manual recipients
- **File:** `main.py` (JS)
- **Issue:** Arbitrary text accepted as email addresses in recipient field.
- **Fix:** Added basic email format validation before adding manual recipients.

### 4. No input sanitization on manual emails in dispatch
- **File:** `main.py` (Python)
- **Issue:** `manual_emails` passed to SMTP without validation.
- **Fix:** Added email format validation in `dispatch_emails()` and scheduler before sending.

---

## Bugs

### 5. Seed script uses wrong settings keys
- **File:** `seed.py`
- **Issue:** Seeded `"email"` and `"password"` keys, but app reads `"sender_email"` and `"app_password"`.
- **Fix:** Changed seed keys to `"sender_email"` and `"app_password"`.

### 6. Dual family tracking (roster.family_id vs family_members junction)
- **Files:** `db_manager.py`, `main.py`
- **Issue:** `add_contact()` writes to `roster.family_id` but app reads from `family_members` junction table. New contacts with families don't appear in family member lists.
- **Fix:** `add_contact()` now also inserts into `family_members` when `family_id` is provided. `update_contact()` syncs `family_members` on updates.

### 7. getRecipientSelection() only returns last target
- **File:** `main.py` (JS)
- **Issue:** Multiple target types (e.g., group + "All Singles") overwrite each other; only the last one is used.
- **Fix:** Changed to collect all targets into an array and pass them to dispatch. Updated `dispatch_emails()` to iterate over multiple targets.

### 8. Scheduler status logic marks partially-failed sends as "sent"
- **File:** `main.py` (Python)
- **Issue:** If SMTP connection succeeds but individual emails fail, status was "sent" even with failures.
- **Fix:** Status is now "sent" only if no connection error AND failed count is 0. Otherwise "partial" or "failed".

### 9. Scheduler comment says 60s, sleeps 30s
- **File:** `main.py` (Python)
- **Issue:** Comment said "every 60 seconds" but `time.sleep(30)`.
- **Fix:** Updated comment to say 30s.

---

## Redundancies

### 10. Duplicate/garbled section headers in JS
- **File:** `main.py` (JS)
- **Issue:** Two "EMAIL DISPATCH" section dividers with garbled Unicode characters.
- **Fix:** Removed duplicate, cleaned up headers.

### 11. get_contacts_by_family() uses stale roster.family_id
- **File:** `db_manager.py`
- **Issue:** Queries `roster.family_id` instead of `family_members` junction table, giving inconsistent results.
- **Fix:** Changed to query via `family_members` junction table.

### 12. context.md is stale
- **File:** `context.md`
- **Issue:** Missing `family_members`, `email_history_details` tables; outdated function list; incorrect method count.
- **Fix:** Not updated here — context.md should be regenerated from current code state.

---

## Design Notes (not fixed, documented)

- **Connection-per-query pattern:** Each db_manager function opens/closes its own connection. Acceptable for current scale.
- **No CSRF/auth on pywebview API:** Local-only app, acceptable risk. CDN compromise could be a vector.
- **CDN dependency:** Quill loaded from jsdelivr. No offline editor. Intentional per design docs.
- **No email dedup in scheduler:** Scheduler doesn't deduplicate across target types like dispatch_emails does. Fixed by adding dedup to scheduler.
