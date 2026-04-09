# Church Roster & Email Dispatcher — Context

## Overview
A standalone Python desktop app (v1.2.0) for managing a church contact roster and sending batch emails with rich text formatting. Built with pywebview for a native desktop window with an embedded web view. Compiles to macOS `.app` (DMG) and Windows `.exe` via PyInstaller. Supports light and dark themes, system tray integration, and automatic update checking via GitHub Releases.

**Repository:** https://github.com/kah-eru/churchemailsender

## Architecture
Strict **two-file** architecture (plus test/support files):

| File | Lines | Role |
|------|-------|------|
| `main.py` | ~4,580 | Presentation & logic — pywebview window, full inline HTML/CSS/JS frontend, Python `Api` class, email dispatch, background scheduler, system tray, update checker |
| `db_manager.py` | ~905 | Data layer — SQLite CRUD for contacts, families, groups, settings, templates, schedules, history, recurrence engine |
| `seed.py` | ~387 | Dev utility — populates DB with realistic fake data (12 families, 15 singles, 6 groups, 8 templates, 20 history entries, 5 scheduled emails) |
| `test_api.py` | ~1,330 | 146 unit tests for all Api class methods (mocked SMTP, mocked HTTP for update checks, email presets) |
| `test_db_manager.py` | ~851 | 102 unit tests for all db_manager functions |
| `test_scheduler.py` | ~556 | 25 unit tests for background scheduler and recurrence |
| `test_ui.py` | ~1,180 | 111 Playwright end-to-end UI tests covering every tab, modal, interaction, email provider presets, and help panel |
| `conftest_ui.py` | ~165 | Shared fixtures for Playwright tests — HTTP bridge between browser JS and Python Api class |
| `context.md` | — | This file |
| `requirements.txt` | — | Runtime dependencies: `pywebview>=4.0`, `pystray>=0.19`, `Pillow>=9.0` |
| `.github/workflows/build.yml` | — | GitHub Actions CI — builds macOS DMG + Windows exe on tag push, creates GitHub Release |

### db_manager.py
- Initializes `contacts.db` in the project directory
- **Tables:**
  - `families` — `id` (INTEGER PK), `name` (TEXT UNIQUE)
  - `roster` — `id` (INTEGER PK), `name` (TEXT), `email` (TEXT), `category` (TEXT: 'Family' or 'Single'), `family_id` (FK to families, nullable), `phone`, `notes`, `opt_out`, `created_at`, `last_emailed_at`, `email_count`
  - `settings` — `key` (TEXT PK), `value` (TEXT) — stores email credentials, timezone, SMTP config, UI preferences, setup banner state
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
- **Recurrence engine:** `compute_next_occurrence()` handles daily, every-other-day, weekly (multi-day), every-other-week, and monthly patterns. JS day convention (0=Sun..6=Sat) converted to Python (0=Mon..6=Sun) via `(d - 1) % 7`. Monthly clamps `day_of_month` to valid range using `calendar.monthrange`.

### main.py
- **Constants:** `APP_VERSION = "1.2.0"`, `GITHUB_REPO = "kah-eru/churchemailsender"`, `APP_NAME`, `IS_FROZEN`, `APP_DIR`
- **Python backend (`class Api`)** — ~40+ methods exposed to JavaScript via pywebview's `js_api` bridge:
  - Contact/family CRUD: `get_contacts`, `add_contact`, `update_contact`, `delete_contacts`, `set_contact_opt_out`, `bulk_update_category`, `bulk_add_to_group`, `get_families`, `add_family`, `rename_family`, `delete_family`, `add_family_member`, `remove_family_member`
  - Settings: `get_settings`, `save_settings`, `save_timezone`, `test_email_connection`, `send_test_email`, `set_launch_on_startup`, `get_ui_setting`, `set_ui_setting`, `get_email_presets`
  - First-time setup: `check_email_setup` (returns `{configured, dismissed}`), `dismiss_setup_banner`
  - Update check: `get_app_version`, `check_for_updates` (queries GitHub Releases API, compares version tuples)
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
  - `_friendly_smtp_error(...)` — converts raw SMTP errors into user-friendly messages (auth failure, connection refused, timeout, SSL, relay denied, DNS)
- **Background scheduler** — daemon thread (`run_scheduler`) checks `get_due_emails()` every 30 seconds, resolves recipients, filters opted-out contacts, builds and sends emails, updates status, logs to email_history with per-recipient details, handles recurring emails via `compute_next_occurrence`, uses configured timezone via `ZoneInfo`
- **System tray** — minimizes to tray on window close (macOS via pyobjc NSStatusBarItem, Windows/Linux via pystray), with Show/Quit menu items
- **Startup registration** — optional launch-on-login via macOS LaunchAgent or Windows registry
- **Inline HTML/CSS/JS** — single `HTML` string containing the entire frontend (~2,800 lines of HTML/CSS/JS):
  - CSS custom properties (variables) for light/dark theme switching
  - Two-column layout with resizable panels (drag divider)
  - Left panel: 7 tabs (Contacts / Families / Groups / Scheduled / History / Analytics / Settings)
  - Right panel: email composer with template selector, autocomplete recipient search, Quill rich text editor, file attachments, recurrence controls, scheduling
  - First-time setup overlay (full-screen modal blocking app until dismissed)
  - Side reminder notification (top-right, persistent until credentials configured)
  - Email provider presets dropdown with auto-fill, "? Help" button with provider-specific instructions
  - Update checker section in Settings tab
  - Theme toggle button (top-right)
- **Bootstrap** — `webview.create_window(html=HTML, js_api=api)` renders the app in a native window; scheduler thread starts before window

## Tech Stack
| Component | Technology |
|-----------|-----------|
| Language | Python 3.12+ |
| GUI | pywebview (native window with embedded web view, WKWebView on macOS, EdgeChromium on Windows) |
| Rich Text Editor | Quill 2.x (Snow theme) — loaded from jsDelivr CDN |
| Database | sqlite3 (local `contacts.db` file) |
| Email | smtplib + email.mime (SMTP with TLS, port 587 by default; presets for Gmail, Outlook, Yahoo, iCloud, Zoho, or custom domain) |
| Styling | CSS custom properties, neutral gray palette with teal accent, light/dark mode |
| Timezone | `zoneinfo.ZoneInfo` (Python 3.9+ stdlib) for timezone-aware scheduling |
| Update Check | `urllib.request` against GitHub Releases API (`api.github.com`) |
| Build | PyInstaller (`--onedir` macOS, `--onefile` Windows) |
| CI/CD | GitHub Actions — builds on tag push (`v*`), creates GitHub Release with DMG + exe |
| Testing | pytest + pytest-playwright + unittest.mock |

## Key Design Decisions
- **pywebview over customtkinter**: Chosen to enable a JavaScript WYSIWYG editor (Quill). customtkinter has no equivalent rich text library. pywebview provides a native window with a real browser engine.
- **Quill 2.x via CDN (jsDelivr)**: The app requires internet for email anyway. Quill 2.x is loaded via two CDN tags (CSS + JS). The Snow theme provides a complete built-in toolbar with formatting, headings, lists, text color, background color, links, and images. Quill 2.x uses inline styles for color/background by default — no attributor registration needed (unlike Quill 1.x).
- **CSS custom properties for theming**: All colors are defined as CSS variables in `:root` (dark) and `body.light` (light). Quill Snow CSS overrides are in a separate `<style>` block that loads after `quill.snow.css` to win specificity. Theme preference persists via `localStorage`.
- **Per-recipient MIME message building**: Inline images are extracted as raw bytes once, then a fresh `MIMEImage` is created per recipient to avoid object-reuse issues. Attachments use `mimetypes.guess_type()` for proper MIME detection and RFC-compliant `Content-Disposition` headers.
- **pywebviewready event for bridge sync**: Instead of polling for `window.pywebview.api`, the app listens for the `pywebviewready` event (fired by pywebview when the JS bridge is ready). All initial data loads run sequentially after this event to avoid race conditions.
- **All HTML/JS inline in Python**: Maintains the two-file constraint. No separate `.html` or `.js` files.
- **`<input type="datetime-local">` for scheduling**: `<input type="date">` and `<input type="time">` return empty `.value` in pywebview's WKWebView on macOS. A single `datetime-local` input works correctly.
- **Template-first scheduling**: "Save & Schedule" button opens template save modal (override/save as new) before scheduling. Recipients are saved to templates.
- **Client-side autocomplete**: Cached contact/family/group data with debounced search, keyboard navigation, color-coded type badges. Selecting a family expands to individual member email chips.
- **Update check via GitHub Releases API**: Uses `urllib.request` (stdlib) — no third-party dependency. Compares semantic version tuples. Returns download link to the GitHub Release page. Gracefully handles network errors, timeouts, and malformed responses.

## Email Configuration
Credentials are configured via the **Settings tab** in the app UI. They are stored in the `settings` table in `contacts.db`.

On first launch, a **full-screen setup overlay** blocks the app and prompts the user to configure email settings. After dismissing the overlay, a **persistent side reminder** notification appears in the top-right corner until credentials are saved.

### Email Provider Presets
A **provider dropdown** in Settings lets users select from built-in presets that auto-fill SMTP host and port:

| Provider | SMTP Host | Port |
|----------|-----------|------|
| Gmail | `smtp.gmail.com` | `587` |
| Outlook / Hotmail | `smtp-mail.outlook.com` | `587` |
| Yahoo Mail | `smtp.mail.yahoo.com` | `587` |
| iCloud Mail | `smtp.mail.me.com` | `587` |
| Zoho Mail | `smtp.zoho.com` | `587` |
| Custom / Own Domain | *(user-provided)* | *(user-provided)* |

When a preset provider is selected, the SMTP host/port fields are hidden (auto-filled from the preset). Selecting "Custom / Own Domain" reveals manual host/port input fields. The dropdown auto-detects the current provider on load by matching the saved `smtp_host` against known presets.

A **"? Help" button** next to the section title toggles an instructional panel with step-by-step setup instructions and provider-specific guidance (e.g., how to generate an App Password for Gmail, where to find SMTP settings for custom domains). The help text updates dynamically as the user switches providers.

API method: `get_email_presets()` — returns a list of `{id, name, host, port, help}` objects.

Settings stored in DB:
- `sender_email` — Email address (any SMTP provider)
- `app_password` — App Password (provider-specific; for Gmail: generated at Google Account > Security > App Passwords)
- `sender_name` — Display name for outgoing emails
- `smtp_host` — SMTP server hostname (default: `smtp.gmail.com`)
- `smtp_port` — SMTP port (default: `587`)
- `timezone` — Timezone for scheduling (default: `US/Eastern`)
- `setup_banner_dismissed` — Whether the first-time setup overlay has been dismissed

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
Runtime (`requirements.txt`):
```
pywebview>=4.0    — native desktop window with embedded web view
pystray>=0.19     — system tray support (Windows/Linux)
Pillow>=9.0       — tray icon generation (PIL.Image, PIL.ImageDraw)
```

macOS additionally requires: `pyobjc-framework-Cocoa` (for native tray)

Dev/test dependencies:
```
pytest            — test runner
pytest-cov        — coverage reporting
playwright        — end-to-end UI testing
pytest-playwright — pytest integration for playwright
```

All other imports (`sqlite3`, `smtplib`, `email`, `csv`, `mimetypes`, `threading`, `json`, `base64`, `uuid`, `re`, `zoneinfo`, `urllib.request`, `urllib.error`, `platform`, `subprocess`, `os`, `sys`, `time`) are Python stdlib.

## Running the App
```bash
cd "Project For Church"
source venv/bin/activate
python3 main.py
```

## Building Installers

### Local build
```bash
pip install pyinstaller
# macOS
pyinstaller --onedir --windowed --name "Church Roster" --add-data "db_manager.py:." main.py
# Windows
pyinstaller --onefile --windowed --name "ChurchRoster" --add-data "db_manager.py;." main.py
```

### CI/CD (GitHub Actions)
The `.github/workflows/build.yml` workflow:
1. Triggers on tag push matching `v*` (e.g., `git tag v1.1.0 && git push origin v1.1.0`)
2. Builds macOS `.app` bundle → creates DMG via `create-dmg`
3. Builds Windows `.exe` via PyInstaller `--onefile`
4. Creates a GitHub Release with both artifacts attached
5. Auto-generates release notes from commits

### Release process
```bash
# 1. Update APP_VERSION in main.py
# 2. Commit and push
git add main.py && git commit -m "Bump version to X.Y.Z" && git push origin main
# 3. Tag and push to trigger build
git tag -a vX.Y.Z -m "vX.Y.Z — description" && git push origin vX.Y.Z
```

---

## Implemented Features

### 1. First-Time Email Setup Flow
- **Full-screen overlay** on first launch blocks the app with a "Welcome!" modal prompting users to configure email settings
- Single "Go to Settings" button dismisses the overlay and navigates to the Settings tab
- Overlay dismissal is persisted in the database (`setup_banner_dismissed` setting) — does not reappear
- **Persistent side reminder** notification appears in the top-right corner after dismissing the overlay, as long as credentials remain unconfigured
- Side reminder has a close button (session-only dismiss — reappears on reload) and a "Configure Email Settings" button
- Both notifications automatically disappear once valid credentials are saved
- API methods: `check_email_setup()`, `dismiss_setup_banner()`

### 2. Update Checker
- **Settings tab** shows current app version and a "Check for Updates" button
- Queries `api.github.com/repos/kah-eru/churchemailsender/releases/latest`
- Compares semantic version tuples (handles `v` prefix, malformed versions)
- Shows color-coded result banner:
  - **Green** — new version available with clickable download link to GitHub Release page
  - **Blue** — app is up to date
  - **Red** — network/timeout error with message
- API methods: `get_app_version()`, `check_for_updates()`
- No auto-download — users click the link to download the new DMG/exe from GitHub

### 3. GUI Settings for Email Credentials
- Settings tab with email provider dropdown (presets for Gmail, Outlook, Yahoo, iCloud, Zoho, custom domain)
- Selecting a preset auto-fills SMTP host/port; "Custom / Own Domain" reveals manual host/port inputs
- **"? Help" button** toggles an instructional panel with provider-specific setup guidance (app passwords, SMTP details)
- Help panel content updates dynamically when the selected provider changes
- "Test Connection" button verifies SMTP login before saving
- "Send Test Email" button sends a test message to the configured sender address
- Credentials stored in `settings` table (no hardcoded constants)
- Friendly SMTP error messages for common failures (auth, connection refused, timeout, SSL, relay denied, DNS)
- Validation: custom provider requires an SMTP host before saving

### 4. Email Attachments + Inline Images
- "Attach Files" button opens native file picker via `webview.FileDialog.OPEN` (multiple files)
- File chips with size display and remove button
- Quill image toolbar button for inline images (paste or insert)
- Base64 data-URI images extracted and converted to CID-attached inline images per recipient
- File attachments use `mimetypes.guess_type()` for correct MIME types and RFC-compliant Content-Disposition
- MIME structure: mixed > related > alternative (text + html) + inline CID images + file attachments

### 5. Scheduled & Group-Targeted Emails
- Custom groups with member management (Groups tab)
- Target selector dropdown in composer: All Contacts, All Families, All Singles, or specific group
- Checkbox selection on contacts tab overrides target dropdown
- `datetime-local` picker + "Save & Schedule" button for deferred send (opens template save modal first)
- **Recurring schedules**: daily, every other day, weekly (multi-day selection), every other week (multi-day), monthly (day-of-month), with optional end date
- Background scheduler daemon thread checks every 30s for due emails, uses configured timezone
- Scheduler logs sent emails to history and creates next occurrences for recurring emails
- Scheduled tab shows calendar view with pending/sent/failed emails, cancel and duplicate options

### 6. Email Templates
- Save current composer (subject + editor body + recipients) as a named template
- Load template dropdown populates subject + editor HTML + recipient chips
- Override existing or save as new template via modal
- Delete and duplicate templates from dropdown

### 7. CSV Import/Export
- Import CSV with columns: `name`, `email`, `category`, `family_name`
- Auto-creates family records during import if they don't exist
- Invalid categories default to 'Single'
- Export full roster to CSV via native save dialog (`webview.FileDialog.SAVE`)

### 8. Email History / Analytics
- Every dispatch (immediate or scheduled) logged to `email_history` table with per-recipient details
- History tab shows subject, target type, recipient/sent/failed counts, timestamp
- Click a history entry to see per-recipient status (sent/failed) and error messages
- Search and date-range filtering for history
- Analytics tab: weekly send volume (last 12 weeks), top failed recipients, overall totals and failure rate

### 9. Autocomplete Recipient Search
- Type-ahead search in the send-to field matches contacts, families, and groups
- Color-coded dropdown: contacts (teal), families (purple), groups (gold)
- Keyboard navigation (arrow keys, Enter, Escape)
- Selecting a family expands all members as individual recipient chips
- Cached data refreshes on contact/family/group mutations and CSV import

### 10. Tab Search Bars
- Contacts, Families, and Groups tabs each have a search bar
- Filters lists in real-time by name/email

### 11. Timezone Setting
- Settings tab has a timezone dropdown with common US/international timezones
- Stored in `settings` table, used by scheduler for timezone-aware send times
- Defaults to `US/Eastern`

### 12. Light/Dark Theme
- CSS custom properties define neutral gray palette with teal accent
- Dark mode: charcoal backgrounds (#1a1a1a/#242424), teal accent (#5a9e8f)
- Light mode: white/light gray (#f5f5f5/#ffffff), darker teal accent (#3d8b7a)
- Toggle button in top-right corner, persists to localStorage
- All inline styles in HTML and JS-generated markup use `var(--*)` references

### 13. Contact Opt-Out
- Per-contact opt-out flag — opted-out contacts are excluded from both immediate dispatch and scheduled sends
- Opt-out filter in contacts tab (All / Active only / Opted-out only)

### 14. CC/BCC Support
- Expandable CC and BCC email fields in the composer
- Validated with email regex, invalid addresses filtered out

### 15. Database Backup/Restore
- Backup database to a file via native save dialog
- Restore from a backup file via native open dialog

### 16. System Tray & Startup
- App minimizes to system tray on window close (macOS + Windows)
- Optional launch-on-login via Settings toggle (macOS LaunchAgent / Windows registry)

---

## Testing

**384 tests** across 4 test files, all passing:

| File | Tests | What it covers |
|------|-------|----------------|
| `test_api.py` | 146 | All Api class methods — contacts, families, groups, settings, email presets, email setup check, update checker, dispatch, scheduling, templates, history, CSV, file picker, backup/restore |
| `test_db_manager.py` | 102 | All db_manager functions — init, settings, families, contacts, groups, family members, scheduled emails, templates, email history, analytics, backup/restore, recurrence |
| `test_scheduler.py` | 25 | Background scheduler — due email processing, recurrence, opt-out filtering, failures, timezone, history logging, attachments |
| `test_ui.py` | 111 | Playwright end-to-end — tab navigation, contacts CRUD, families, groups, settings, email provider presets, email help panel, setup overlay, setup reminder, update check, theme toggle, composer, scheduled tab, history, analytics, layout, modals, autocomplete, toasts, bulk operations, edge cases |

### Running tests
```bash
source venv/bin/activate

# All tests (unit + Playwright)
python -m pytest test_api.py test_db_manager.py test_scheduler.py test_ui.py -v

# Unit tests only (~1 second)
python -m pytest test_api.py test_db_manager.py test_scheduler.py -v

# Playwright UI tests only (~3 minutes)
python -m pytest test_ui.py -v

# With coverage
python -m pytest test_api.py test_db_manager.py test_scheduler.py --cov=db_manager --cov=main --cov-report=term-missing
```

### Test architecture
- **Isolation:** Tests use `tmp_path` fixtures for isolated temp databases per test. Each test gets a fresh DB.
- **SMTP mocking:** `unittest.mock.patch("main.smtplib.SMTP")` for all email dispatch tests.
- **HTTP mocking:** `unittest.mock.patch("main.urllib.request.urlopen")` for update checker tests.
- **Playwright HTTP bridge:** `conftest_ui.py` starts two HTTP servers — one serves the app HTML (with a mock `pywebview.api` JS proxy injected), and one forwards JS API calls to the real Python `Api` class. This allows full end-to-end testing of the JS frontend against the real backend.
- **Setup overlay handling:** The `ui_db` fixture auto-dismisses the setup banner by default (sets `setup_banner_dismissed=true`) so the overlay doesn't block non-setup tests. Tests in `TestSetupOverlay` opt out via the `fresh_setup_db` fixture to test the overlay behavior.

## Performance

All backend operations complete in < 2ms with 500 contacts. The app uses ~5MB Python memory (plus pywebview's browser engine). Key optimizations:
- Batch queries for families/groups avoid N+1 (2 queries instead of 1+N)
- SQLite indexes on frequently queried columns
- Scheduler polls every 30s with minimal CPU overhead
- Update checker uses 10s timeout, only runs on demand (not automatic)

## Version History

| Version | Tag | Changes |
|---------|-----|---------|
| 1.0.0 | `v1.0.0` | Initial release — contacts, families, groups, email dispatch, scheduling, templates, history, analytics, CSV, themes, tray |
| 1.1.0 | `v1.1.0` | First-time email setup overlay + side reminder, update checker in Settings tab, 30 new tests |
| 1.2.0 | `v1.2.0` | Email provider presets (Gmail, Outlook, Yahoo, iCloud, Zoho, Custom), help panel with provider-specific instructions, 25 new tests |
