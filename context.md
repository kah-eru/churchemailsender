# Church Roster & Email Dispatcher — Context

## Overview
A standalone Python desktop app for managing a church contact roster and sending batch emails with rich text formatting. Runs natively on macOS via terminal and is designed to compile into a Windows `.exe` later. Supports light and dark themes.

## Architecture
Strict **two-file** architecture:

| File | Role |
|------|------|
| `db_manager.py` | Data layer — SQLite CRUD for contacts, families, groups, settings, templates, schedules, history |
| `main.py` | Presentation & logic — pywebview window, JS/HTML frontend, email dispatch, background scheduler |

### db_manager.py
- Initializes `contacts.db` in the project directory
- **Tables:**
  - `families` — `id` (INTEGER PK), `name` (TEXT UNIQUE)
  - `roster` — `id` (INTEGER PK), `name` (TEXT), `email` (TEXT), `category` (TEXT: 'Family' or 'Single'), `family_id` (FK to families, nullable)
  - `settings` — `key` (TEXT PK), `value` (TEXT) — stores email credentials and timezone
  - `groups_` — `id` (INTEGER PK), `name` (TEXT UNIQUE) — note trailing underscore to avoid SQL reserved word
  - `group_members` — `id` (INTEGER PK), `group_id` (FK), `contact_id` (FK), unique pair constraint
  - `scheduled_emails` — `id`, `subject`, `html_body`, `plain_text`, `target_type`, `target_id`, `contact_ids` (JSON), `attachment_paths` (JSON), `scheduled_at`, `status` ('pending'/'sent'/'failed'/'cancelled'), `created_at`, `sent_at`, `result`, `recurrence` (JSON), `manual_emails` (JSON)
  - `email_templates` — `id`, `name` (UNIQUE), `subject`, `html_body`, `recipients` (JSON)
  - `email_history` — `id`, `subject`, `target_description`, `recipient_count`, `sent_count`, `failed_count`, `sent_at`
- **Indexes:** `idx_roster_family_id` on `roster(family_id)`, `idx_scheduled_status_at` on `scheduled_emails(status, scheduled_at)`
- **Migrations:** `init_db()` adds columns via `ALTER TABLE ADD COLUMN` wrapped in try/except for idempotency: `scheduled_emails.recurrence`, `scheduled_emails.manual_emails`, `email_templates.recipients`
- **Functions:** `init_db()`, settings CRUD (`get_setting`, `set_setting`), family CRUD, contact CRUD, group CRUD (`add_group`, `get_groups`, `delete_group`, `add_group_member`, `remove_group_member`, `get_group_members`), scheduled email CRUD (`schedule_email`, `get_scheduled_emails`, `cancel_scheduled_email`, `get_due_emails`, `update_email_status`), template CRUD (`save_template`, `update_template`, `get_templates`, `delete_template`), history (`log_email`, `get_email_history`), recurrence (`compute_next_occurrence`)

### main.py
- **Python backend (`class Api`)** — ~30 methods exposed to JavaScript via pywebview's `js_api` bridge:
  - Contact/family CRUD: `get_contacts`, `add_contact`, `delete_contacts`, `get_families`, `add_family`, `delete_family`
  - Settings: `get_settings`, `save_settings`, `save_timezone`, `test_email_connection`
  - Groups: `get_groups`, `add_group`, `delete_group`, `add_group_member`, `remove_group_member`
  - Email dispatch: `dispatch_emails` (with attachments, inline images, target resolution)
  - Scheduling: `schedule_email` (with recurrence and manual_emails), `get_scheduled_emails`, `cancel_scheduled_email`, `resolve_recipients`
  - Templates: `get_templates`, `save_template`, `update_template`, `delete_template` (templates store recipients)
  - CSV: `import_csv`, `export_csv`
  - History: `get_email_history`
  - File picker: `pick_file`
- **Static helpers:**
  - `_extract_inline_images(html_body)` — finds base64 data-URI images in HTML, returns `(new_html, [(cid, mime_type, raw_bytes), ...])` for per-recipient MIME building
  - `_build_message(...)` — builds a fresh MIME message per recipient with proper structure, inline CID images, and attachments with correct MIME type detection
- **Background scheduler** — daemon thread (`run_scheduler`) checks `get_due_emails()` every 30 seconds, resolves recipients, builds and sends emails, updates status, logs to email_history, handles recurring emails via `compute_next_occurrence`, uses configured timezone via `ZoneInfo`
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
pywebview
```
All other dependencies (`sqlite3`, `smtplib`, `email`, `csv`, `mimetypes`, `threading`, `json`, `base64`, `uuid`, `re`) are Python stdlib.

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
- Every dispatch (immediate or scheduled) logged to `email_history` table
- History tab shows subject, target type, recipient/sent/failed counts, timestamp

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
