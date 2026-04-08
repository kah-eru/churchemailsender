# Church Roster & Email Dispatcher — Context

## Overview
A standalone Python desktop app for managing a church contact roster and sending batch emails with rich text formatting. Runs natively on macOS via terminal and is designed to compile into a Windows `.exe` later. Supports light and dark themes.

## Architecture
Strict **two-file** architecture (plus test files):

| File | Role |
|------|------|
| `db_manager.py` | Data layer — SQLite CRUD for contacts, families, groups, settings, templates, schedules, history |
| `main.py` | Presentation & logic — pywebview window, JS/HTML frontend, email dispatch, background scheduler |
| `seed.py` | Dev utility — populates DB with realistic fake data (12 families, 15 singles, 6 groups, 8 templates, 20 history entries, 5 scheduled emails) |
| `test_db_manager.py` | Unit tests for all db_manager functions |
| `test_api.py` | Unit tests for all Api class methods |
| `test_scheduler.py` | Unit tests for background scheduler and recurrence |
| `test_ui.py` | Playwright end-to-end UI tests |
| `conftest_ui.py` | Shared fixtures for Playwright tests (HTTP bridge between browser JS and Python Api) |

### db_manager.py
- Initializes `contacts.db` in the project directory
- **Tables:**
  - `families` — `id` (INTEGER PK), `name` (TEXT UNIQUE)
  - `roster` — `id` (INTEGER PK), `name` (TEXT), `email` (TEXT), `category` (TEXT: 'Family' or 'Single'), `family_id` (FK to families, nullable), `phone`, `notes`, `opt_out`, `created_at`, `last_emailed_at`, `email_count`
  - `settings` — `key` (TEXT PK), `value` (TEXT) — stores email credentials, timezone, SMTP config, UI preferences
  - `groups_` — `id` (INTEGER PK), `name` (TEXT UNIQUE) — note trailing underscore to avoid SQL reserved word
  - `group_members` — `id` (INTEGER PK), `group_id` (FK), `contact_id` (FK), unique pair constraint
  - `family_members` — junction table: `id` (INTEGER PK), `family_id` (FK), `contact_id` (FK), unique pair — a contact can belong to multiple families
  - `scheduled_emails` — `id`, `subject`, `html_body`, `plain_text`, `target_type`, `target_id`, `contact_ids` (JSON), `attachment_paths` (JSON), `scheduled_at`, `status` ('pending'/'sent'/'failed'/'cancelled'), `created_at`, `sent_at`, `result`, `recurrence` (JSON), `manual_emails` (JSON)
  - `email_templates` — `id`, `name` (UNIQUE), `subject`, `html_body`, `recipients` (JSON)
  - `email_history` — `id`, `subject`, `target_description`, `recipient_count`, `sent_count`, `failed_count`, `sent_at`
  - `email_history_details` — per-recipient send results with status and error messages
- **Indexes:** `idx_roster_family_id` on `roster(family_id)`, `idx_scheduled_status_at` on `scheduled_emails(status, scheduled_at)`, `idx_ehd_history_id` on `email_history_details(history_id)`
- **Migrations:** `init_db()` adds columns via `ALTER TABLE ADD COLUMN` wrapped in try/except for idempotency
- **Batch queries:** `get_all_families_with_members()` and `get_all_groups_with_members()` fetch all families/groups with their members in 2 queries total (avoids N+1). Single-entity versions (`get_family_members_via_junction`, `get_group_members`) still available for targeted lookups.
- **Family sync:** `update_contact()` properly cleans up the old `family_members` junction row when a contact's primary family changes, preventing stale memberships

### main.py
- **Python backend (`class Api`)** — ~35 methods exposed to JavaScript via pywebview's `js_api` bridge:
  - Contact/family CRUD: `get_contacts`, `add_contact`, `update_contact`, `delete_contacts`, `set_contact_opt_out`, `bulk_update_category`, `bulk_add_to_group`, `get_families`, `add_family`, `rename_family`, `delete_family`, `add_family_member`, `remove_family_member`
  - Settings: `get_settings`, `save_settings`, `save_timezone`, `test_email_connection`, `send_test_email`, `set_launch_on_startup`, `get_ui_setting`, `set_ui_setting`
  - Groups: `get_groups`, `add_group`, `rename_group`, `delete_group`, `add_group_member`, `remove_group_member`, `add_family_to_group`
  - Email dispatch: `dispatch_emails` (with attachments, inline images, CC/BCC, target resolution, opt-out filtering), `get_recipient_count`
  - Scheduling: `schedule_email`, `get_scheduled_emails`, `get_scheduled_emails_with_recipients`, `get_scheduled_email_detail`, `cancel_scheduled_email`, `update_scheduled_email`, `duplicate_scheduled_email`, `resolve_recipients`
  - Templates: `get_templates`, `save_template`, `update_template`, `delete_template`, `duplicate_template`
  - CSV: `import_csv`, `export_csv`
  - History: `get_email_history`, `get_email_history_details`, `get_email_history_filtered`, `get_analytics`
  - Database: `backup_database`, `restore_database`
  - File picker: `pick_file`
- **Static helpers:**
  - `_extract_inline_images(html_body)` — finds base64 data-URI images in HTML, returns `(new_html, [(cid, mime_type, raw_bytes), ...])` for per-recipient MIME building
  - `_build_message(...)` — builds a fresh MIME message per recipient with proper structure, inline CID images, and attachments with correct MIME type detection
  - `_send_to_recipients(...)` — shared email-sending logic used by both `dispatch_emails` and the scheduler, handles per-recipient failures gracefully, updates contact email stats
  - `_friendly_smtp_error(...)` — converts raw SMTP errors into user-friendly messages
- **Background scheduler** — daemon thread (`run_scheduler`) checks `get_due_emails()` every 30 seconds, resolves recipients, filters opted-out contacts, builds and sends emails, updates status, logs to email_history with per-recipient details, handles recurring emails via `compute_next_occurrence`, uses configured timezone via `ZoneInfo`
- **System tray** — minimizes to tray on window close (macOS via pyobjc NSStatusBarItem, Windows/Linux via pystray), with Show/Quit menu items
- **Startup registration** — optional launch-on-login via macOS LaunchAgent or Windows registry
- **Inline HTML/CSS/JS** — single `HTML` string containing the entire frontend:
  - CSS custom properties (variables) for light/dark theme switching
  - Two-column layout with responsive panels
  - Left panel: 6 tabs (Contacts / Families / Groups / Scheduled / History / Settings), Contacts/Families/Groups tabs have search bars for filtering
  - Right panel: email composer with template selector (saves/restores recipients), autocomplete recipient search field with color-coded dropdown (contacts=teal, families=purple, groups=gold), recipient chips, Quill rich text editor (with image support), file attachment area with chips, recurrence controls (daily/every other day/weekly/every other week/monthly with day selection and optional end date), datetime-local picker, Save & Schedule button (opens template save modal first), and Send Now button
  - Theme toggle button (top-right) persists choice to localStorage
- **Bootstrap** — `webview.create_window(html=HTML, js_api=api)` renders the app in a native window; scheduler thread starts before window

## Tech Stack
| Component | Technology |
|-----------|-----------|
| Language | Python 3 |
| GUI | pywebview (native window with embedded web view, WKWebView on macOS) |
| Rich Text Editor | Quill 2.x (Snow theme) — loaded from jsDelivr CDN |
| Database | sqlite3 (local `contacts.db` file) |
| Email | smtplib + email.mime (Gmail SMTP, TLS on port 587) |
| Styling | CSS custom properties, neutral gray palette with teal accent, light/dark mode |
| Timezone | `zoneinfo.ZoneInfo` (Python 3.9+ stdlib) for timezone-aware scheduling |

## Key Design Decisions
- **pywebview over customtkinter**: Chosen to enable a JavaScript WYSIWYG editor (Quill). customtkinter has no equivalent rich text library. pywebview provides a native window with a real browser engine.
- **Quill 2.x via CDN (jsDelivr)**: The app requires internet for email anyway. Quill 2.x is loaded via two CDN tags (CSS + JS). The Snow theme provides a complete built-in toolbar with formatting, headings, lists, text color, background color, links, and images. Quill 2.x uses inline styles for color/background by default — no attributor registration needed (unlike Quill 1.x).
- **CSS custom properties for theming**: All colors are defined as CSS variables in `:root` (dark) and `body.light` (light). Quill Snow CSS overrides are in a separate `<style>` block that loads after `quill.snow.css` to win specificity. Theme preference persists via `localStorage`.
- **Per-recipient MIME message building**: Inline images are extracted as raw bytes once, then a fresh `MIMEImage` is created per recipient to avoid object-reuse issues. Attachments use `mimetypes.guess_type()` for proper MIME detection and RFC-compliant `Content-Disposition` headers.
- **pywebviewready event for bridge sync**: Instead of polling for `window.pywebview.api`, the app listens for the `pywebviewready` event (fired by pywebview when the JS bridge is ready). All initial data loads run sequentially after this event to avoid race conditions.
- **All HTML/JS inline in Python**: Maintains the two-file constraint. No separate `.html` or `.js` files.
- **`<input type="datetime-local">` for scheduling**: `<input type="date">` and `<input type="time">` return empty `.value` in pywebview's WKWebView on macOS. A single `datetime-local` input works correctly.
- **Recurrence engine**: `compute_next_occurrence()` in `db_manager.py` handles daily, every-other-day, weekly (multi-day), every-other-week, and monthly patterns. JS day convention (0=Sun..6=Sat) converted to Python (0=Mon..6=Sun) via `(d - 1) % 7`. Monthly clamps `day_of_month` to valid range using `calendar.monthrange`.
- **Template-first scheduling**: "Save & Schedule" button opens template save modal (override/save as new) before scheduling. Recipients are saved to templates.
- **Client-side autocomplete**: Cached contact/family/group data with debounced search, keyboard navigation, color-coded type badges. Selecting a family expands to individual member email chips.

## Email Configuration
Credentials are configured via the **Settings tab** in the app UI. They are stored in the `settings` table in `contacts.db`. On first launch, go to Settings, enter your Gmail address and App Password (generated at Google Account > Security > App Passwords), click "Test Connection" to verify, then "Save Settings".

## Email MIME Structure
```
multipart/mixed
├── multipart/related
│   ├── multipart/alternative
│   │   ├── text/plain
│   │   └── text/html (with cid: references for inline images)
│   ├── image/* (inline CID attachment, one per embedded image)
│   └── ...
├── application/* (file attachment with proper MIME type)
└── ...
```

## Dependencies
Installed in a local virtual environment (`venv/`):
```
pywebview     — native desktop window with embedded web view
Pillow        — tray icon generation (PIL.Image, PIL.ImageDraw)
pystray       — system tray support (Windows/Linux)
```
Dev/test dependencies:
```
pytest        — test runner
pytest-cov    — coverage reporting
playwright    — end-to-end UI testing
```
All other dependencies (`sqlite3`, `smtplib`, `email`, `csv`, `mimetypes`, `threading`, `json`, `base64`, `uuid`, `re`, `zoneinfo`) are Python stdlib.

## Running the App
```bash
cd "Project For Church"
source venv/bin/activate
python3 main.py
```

## Compiling to Windows .exe
```bash
pip install pyinstaller
pyinstaller --onefile --windowed main.py
```
The inline HTML/JS is bundled automatically. `db_manager.py` is included as an import.

---

## Implemented Features

### 1. GUI Settings for Email Credentials
- Settings tab with email and app password inputs
- "Test Connection" button verifies SMTP login before saving
- Credentials stored in `settings` table (no more hardcoded constants)

### 2. Email Attachments + Inline Images
- "Attach Files" button opens native file picker via `webview.FileDialog.OPEN` (multiple files)
- File chips with size display and remove button
- Quill image toolbar button for inline images (paste or insert)
- Base64 data-URI images extracted and converted to CID-attached inline images per recipient
- File attachments use `mimetypes.guess_type()` for correct MIME types and RFC-compliant Content-Disposition
- MIME structure: mixed > related > alternative (text + html) + inline CID images + file attachments

### 3. Scheduled & Group-Targeted Emails
- Custom groups with member management (Groups tab)
- Target selector dropdown in composer: All Contacts, All Families, All Singles, or specific group
- Checkbox selection on contacts tab overrides target dropdown
- `datetime-local` picker + "Save & Schedule" button for deferred send (opens template save modal first)
- **Recurring schedules**: daily, every other day, weekly (multi-day selection), every other week (multi-day), monthly (day-of-month), with optional end date
- Background scheduler daemon thread checks every 30s for due emails, uses configured timezone
- Scheduler logs sent emails to history and creates next occurrences for recurring emails
- Scheduled tab shows pending/sent/failed emails with cancel option (auto-refreshes on tab switch)

### 4. Email Templates
- Save current composer (subject + editor body + recipients) as a named template
- Load template dropdown populates subject + editor HTML + recipient chips
- Override existing or save as new template via modal
- Delete templates from dropdown

### 5. CSV Import/Export
- Import CSV with columns: `name`, `email`, `category`, `family_name`
- Auto-creates family records during import if they don't exist
- Export full roster to CSV via native save dialog (`webview.FileDialog.SAVE`)

### 6. Email History / Analytics
- Every dispatch (immediate or scheduled) logged to `email_history` table with per-recipient details
- History tab shows subject, target type, recipient/sent/failed counts, timestamp
- Click a history entry to see per-recipient status (sent/failed) and error messages
- Date-range filtering for history
- Analytics dashboard: weekly send volume (last 12 weeks), top failed recipients, overall totals and failure rate

### 7. Autocomplete Recipient Search
- Type-ahead search in the send-to field matches contacts, families, and groups
- Color-coded dropdown: contacts (teal), families (purple), groups (gold)
- Keyboard navigation (arrow keys, Enter, Escape)
- Selecting a family expands all members as individual recipient chips
- Cached data refreshes on contact/family/group mutations and CSV import

### 8. Tab Search Bars
- Contacts, Families, and Groups tabs each have a search bar
- Filters lists in real-time by name/email

### 9. Timezone Setting
- Settings tab has a timezone dropdown with common US/international timezones
- Stored in `settings` table, used by scheduler for timezone-aware send times
- Defaults to `US/Eastern`

### 10. Light/Dark Theme
- CSS custom properties define neutral gray palette with teal accent
- Dark mode: charcoal backgrounds (#1a1a1a/#242424), teal accent (#5a9e8f)
- Light mode: white/light gray (#f5f5f5/#ffffff), darker teal accent (#3d8b7a)
- Toggle button in top-right corner, persists to localStorage
- All inline styles in HTML and JS-generated markup use `var(--*)` references

### 11. Contact Opt-Out
- Per-contact opt-out flag — opted-out contacts are excluded from both immediate dispatch and scheduled sends
- Opt-out status visible in contact list

### 12. CC/BCC Support
- Composer supports CC and BCC email fields for dispatch

### 13. Database Backup/Restore
- Backup database to a file via native save dialog
- Restore from a backup file via native open dialog

### 14. System Tray & Startup
- App minimizes to system tray on window close (macOS + Windows)
- Optional launch-on-login via Settings toggle

---

## Testing

**248 unit tests** across 4 test files, all passing:

| File | Tests | Coverage Target |
|------|-------|-----------------|
| `test_db_manager.py` | ~100 | `db_manager.py` — 100% line coverage |
| `test_api.py` | ~120 | `main.py` Api class — all methods tested |
| `test_scheduler.py` | ~25 | `run_scheduler` — recurrence, opt-out, failures, timezone, history logging |
| `test_ui.py` | Playwright | End-to-end UI flows via HTTP bridge to real Api |

Run tests: `source venv/bin/activate && python -m pytest test_db_manager.py test_api.py test_scheduler.py -v`

Run with coverage: `python -m pytest test_db_manager.py test_api.py test_scheduler.py --cov=db_manager --cov=main --cov-report=term-missing`

**Test architecture:** Tests use `tmp_path` fixtures for isolated temp databases per test. SMTP is mocked via `unittest.mock.patch`. Playwright UI tests use an HTTP bridge (`conftest_ui.py`) that serves the app HTML with a mock `pywebview.api` proxy backed by the real Python Api class.

## Performance

All backend operations complete in < 2ms with 500 contacts. The app uses ~5MB Python memory (plus pywebview's browser engine). Key optimizations:
- Batch queries for families/groups avoid N+1 (2 queries instead of 1+N)
- SQLite indexes on frequently queried columns
- Scheduler polls every 30s with minimal CPU overhead
