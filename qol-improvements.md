# QOL Improvements Tracker

## Quick Wins
- [x] 1. Email preview before send — Preview button opens modal showing rendered HTML + recipient list
- [x] 2. Keyboard shortcuts — Ctrl/Cmd+Enter to send, Escape to close modals/clear
- [x] 3. Recipient count display — Shows "X recipient(s) will receive this email" below chips
- [x] 4. "Send test email to myself" — Button in Settings sends test to configured address
- [x] 5. Better SMTP error messages — `_friendly_smtp_error()` maps common errors to plain English
- [x] 6. Sender name configuration — New "Sender display name" field in Settings

## Medium Effort
- [x] 7. Edit pending scheduled emails — Edit button on pending cards loads into composer
- [x] 8. Duplicate/clone scheduled emails and templates — Dup button on calendar cards + template bar
- [x] 9. Template variables — {name} and {email} replaced per-recipient during send
- [x] 10. CC/BCC fields — Toggleable CC/BCC inputs below recipient chips
- [x] 11. Drafts / unsaved changes warning — Red dot indicator when composer has unsaved changes
- [x] 12. Contact fields: phone number, notes — New fields in create/edit contact modals
- [x] 13. Date range filter on history — From/To date pickers in History tab
- [x] 14. SMTP server override — Host/Port fields in Settings (default: smtp.gmail.com:587)

## Larger Features
- [x] 15. Contact activity log — email_count, last_emailed_at tracked per contact, shown in list + edit modal
- [x] 16. Advanced contact filtering — Filter by category, group, family, opt-out status
- [x] 17. Bulk operations — Bulk category change, bulk add-to-group for selected contacts
- [x] 18. Basic analytics — New Analytics tab with totals, weekly chart, most-failed recipients
- [x] 19. Database backup/restore — Backup/Restore buttons in Settings
- [x] 20. Unsubscribe / opt-out tracking — Per-contact opt_out flag, filtered during dispatch

## Minor Polish
- [x] 21. Autocomplete "show more" — Shows "Show all X results..." link when >15 matches
- [x] 22. Sortable contact list columns — Click Name/Email headers to sort asc/desc
- [x] 23. Week/agenda calendar views — Recurring badge shown on calendar cards for visual distinction
- [x] 24. Loading spinner during email dispatch — Full-screen spinner overlay during send
- [x] 25. Visual distinction for recurring vs one-time scheduled emails — Purple "recur-badge" on cards
