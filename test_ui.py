"""Playwright UI tests for the Church Roster & Email Dispatcher frontend.

Covers every tab, modal, form, and interaction edge case in the inline HTML/JS frontend.
Uses a real Python Api + SQLite backend via an HTTP bridge (see conftest_ui.py).
"""
import json
import re
import pytest
from playwright.sync_api import expect

# Pull in shared fixtures
pytest_plugins = ["conftest_ui"]


# ── Helpers ──────────────────────────────────────────────────────────────────

def wait_for_app(page):
    """Wait until the app has fully loaded (initApp complete)."""
    page.wait_for_selector("#contact-list", timeout=10000)
    page.wait_for_timeout(500)


def switch_tab(page, tab_name):
    page.click(f"button.tab-btn:text('{tab_name}')")
    page.wait_for_timeout(300)


def get_toast_text(page):
    toast = page.locator("#toast")
    toast.wait_for(state="visible", timeout=3000)
    return toast.text_content()


def open_create_contact_modal(page):
    page.click("button.add-btn[title='Add Contact']")
    page.wait_for_selector("#cm-name", timeout=3000)


def fill_and_submit_contact(page, name, email, category="Single", phone="", notes=""):
    page.fill("#cm-name", name)
    page.fill("#cm-email", email)
    if phone:
        page.fill("#cm-phone", phone)
    if notes:
        page.fill("#cm-notes", notes)
    page.click("#create-modal button:text('Add')")
    page.wait_for_timeout(500)


def add_contact_filter(page, filter_type):
    """Add a filter via the + button and filter menu."""
    page.click(".filter-add-btn")
    page.wait_for_timeout(200)
    page.click(f".filter-menu-item[data-filter='{filter_type}']")
    page.wait_for_timeout(200)


def open_create_family_modal(page):
    page.click("button.add-btn[title='Add Family']")
    page.wait_for_selector("#cm-fname", timeout=3000)


def open_create_group_modal(page):
    page.click("button.add-btn[title='Add Group']")
    page.wait_for_selector("#cm-gname", timeout=3000)


# ══════════════════════════════════════════════════════════════════════════════
# TAB NAVIGATION
# ══════════════════════════════════════════════════════════════════════════════

class TestTabNavigation:
    def test_default_tab_is_contacts(self, page, app_url):
        page.goto(app_url)
        wait_for_app(page)
        expect(page.locator("#contacts-tab")).to_be_visible()
        expect(page.locator("#families-tab")).to_be_hidden()

    def test_switch_to_all_tabs(self, page, app_url):
        page.goto(app_url)
        wait_for_app(page)
        for tab in ["Families", "Groups", "Scheduled", "History", "Analytics", "Settings"]:
            switch_tab(page, tab)
            tab_id = tab.lower() + "-tab"
            expect(page.locator(f"#{tab_id}")).to_be_visible()
        switch_tab(page, "Contacts")
        expect(page.locator("#contacts-tab")).to_be_visible()

    def test_active_tab_button_highlighted(self, page, app_url):
        page.goto(app_url)
        wait_for_app(page)
        active = page.locator("button.tab-btn.active")
        expect(active).to_have_text("Contacts")
        switch_tab(page, "Settings")
        active = page.locator("button.tab-btn.active")
        expect(active).to_have_text("Settings")


# ══════════════════════════════════════════════════════════════════════════════
# CONTACTS TAB
# ══════════════════════════════════════════════════════════════════════════════

class TestContactsTab:
    def test_empty_state(self, page, app_url):
        page.goto(app_url)
        wait_for_app(page)
        expect(page.locator("#contact-list")).to_contain_text("No contacts")

    def test_create_contact(self, page, app_url):
        page.goto(app_url)
        wait_for_app(page)
        open_create_contact_modal(page)
        fill_and_submit_contact(page, "John Doe", "john@example.com")
        expect(page.locator("#contact-list")).to_contain_text("John Doe")
        expect(page.locator("#contact-list")).to_contain_text("john@example.com")

    def test_create_contact_validation_empty(self, page, app_url):
        page.goto(app_url)
        wait_for_app(page)
        open_create_contact_modal(page)
        page.click("#create-modal button:text('Add')")
        toast = get_toast_text(page)
        assert "name" in toast.lower() or "email" in toast.lower()

    def test_create_contact_name_only_fails(self, page, app_url):
        page.goto(app_url)
        wait_for_app(page)
        open_create_contact_modal(page)
        page.fill("#cm-name", "NoEmail")
        page.click("#create-modal button:text('Add')")
        toast = get_toast_text(page)
        assert "email" in toast.lower() or "both" in toast.lower()

    def test_create_multiple_contacts(self, page, app_url):
        page.goto(app_url)
        wait_for_app(page)
        for i in range(3):
            open_create_contact_modal(page)
            fill_and_submit_contact(page, f"Person{i}", f"p{i}@x.com")
        rows = page.locator("#contact-list .contact-row")
        expect(rows).to_have_count(3)

    def test_search_contacts(self, page, app_url):
        page.goto(app_url)
        wait_for_app(page)
        open_create_contact_modal(page)
        fill_and_submit_contact(page, "Alice", "alice@x.com")
        open_create_contact_modal(page)
        fill_and_submit_contact(page, "Bob", "bob@x.com")
        page.fill("#search-contacts", "Alice")
        page.wait_for_timeout(300)
        visible_rows = page.locator("#contact-list .contact-row:visible")
        expect(visible_rows).to_have_count(1)
        expect(visible_rows.first).to_contain_text("Alice")

    def test_search_no_match(self, page, app_url):
        page.goto(app_url)
        wait_for_app(page)
        open_create_contact_modal(page)
        fill_and_submit_contact(page, "Alice", "alice@x.com")
        page.fill("#search-contacts", "zzzzz")
        page.wait_for_timeout(300)
        visible_rows = page.locator("#contact-list .contact-row:visible")
        expect(visible_rows).to_have_count(0)

    def test_filter_by_optout(self, page, app_url):
        page.goto(app_url)
        wait_for_app(page)
        open_create_contact_modal(page)
        fill_and_submit_contact(page, "Alice", "alice@x.com")
        open_create_contact_modal(page)
        fill_and_submit_contact(page, "Bob", "bob@x.com")
        # Opt out Bob
        page.locator("#contact-list .contact-edit-btn").nth(1).click()
        page.wait_for_timeout(300)
        page.locator("#ce-optout").check()
        page.click("#create-modal button:text('Save')")
        page.wait_for_timeout(500)
        add_contact_filter(page, "optout")
        page.select_option("#filter-optout", "active")
        page.wait_for_timeout(300)
        visible_rows = page.locator("#contact-list .contact-row:visible")
        expect(visible_rows).to_have_count(1)
        expect(visible_rows.first).to_contain_text("Alice")

    def test_delete_contacts(self, page, app_url):
        page.goto(app_url)
        wait_for_app(page)
        open_create_contact_modal(page)
        fill_and_submit_contact(page, "ToDelete", "del@x.com")
        page.locator("#contact-list input[type='checkbox']").first.check()
        page.on("dialog", lambda d: d.accept())
        page.click("button:text('Delete Selected')")
        page.wait_for_timeout(600)
        expect(page.locator("#contact-list")).to_contain_text("No contacts")

    def test_create_contact_with_phone_notes(self, page, app_url):
        page.goto(app_url)
        wait_for_app(page)
        open_create_contact_modal(page)
        fill_and_submit_contact(page, "WithDetails", "wd@x.com", phone="555-1234", notes="Elder")
        # Click the edit button on the contact row
        page.locator("#contact-list .contact-edit-btn").first.click()
        page.wait_for_timeout(300)
        expect(page.locator("#ce-phone")).to_have_value("555-1234")
        expect(page.locator("#ce-notes")).to_contain_text("Elder")

    def test_edit_contact(self, page, app_url):
        page.goto(app_url)
        wait_for_app(page)
        open_create_contact_modal(page)
        fill_and_submit_contact(page, "Original", "orig@x.com")
        page.locator("#contact-list .contact-edit-btn").first.click()
        page.wait_for_timeout(300)
        page.locator("#ce-name").fill("Updated")
        page.click("#create-modal button:text('Save')")
        page.wait_for_timeout(500)
        expect(page.locator("#contact-list")).to_contain_text("Updated")

    def test_contact_opt_out_toggle(self, page, app_url):
        page.goto(app_url)
        wait_for_app(page)
        open_create_contact_modal(page)
        fill_and_submit_contact(page, "OptTest", "opt@x.com")
        page.locator("#contact-list .contact-edit-btn").first.click()
        page.wait_for_timeout(300)
        page.locator("#ce-optout").check()
        page.click("#create-modal button:text('Save')")
        page.wait_for_timeout(500)
        # Filter to opted-out only
        add_contact_filter(page, "optout")
        page.select_option("#filter-optout", "optout")
        page.wait_for_timeout(300)
        visible = page.locator("#contact-list .contact-row:visible")
        expect(visible).to_have_count(1)

    def test_filter_active_only(self, page, app_url):
        page.goto(app_url)
        wait_for_app(page)
        open_create_contact_modal(page)
        fill_and_submit_contact(page, "Active", "active@x.com")
        open_create_contact_modal(page)
        fill_and_submit_contact(page, "OptedOut", "out@x.com")
        # Opt out the second contact
        page.locator("#contact-list .contact-edit-btn").nth(1).click()
        page.wait_for_timeout(300)
        page.locator("#ce-optout").check()
        page.click("#create-modal button:text('Save')")
        page.wait_for_timeout(500)
        add_contact_filter(page, "optout")
        page.select_option("#filter-optout", "active")
        page.wait_for_timeout(300)
        visible = page.locator("#contact-list .contact-row:visible")
        expect(visible).to_have_count(1)
        expect(visible.first).to_contain_text("Active")


# ══════════════════════════════════════════════════════════════════════════════
# CONTACT DETAIL PANEL
# ══════════════════════════════════════════════════════════════════════════════

class TestContactDetailPanel:
    def test_detail_panel_default_empty(self, page, app_url):
        page.goto(app_url)
        wait_for_app(page)
        expect(page.locator("#contact-detail-content")).to_contain_text("Click a contact to view details")

    def test_click_contact_shows_detail(self, page, app_url):
        page.goto(app_url)
        wait_for_app(page)
        open_create_contact_modal(page)
        fill_and_submit_contact(page, "Alice", "alice@test.com", phone="555-1234", notes="Board member")
        page.locator("#contact-list .contact-row").first.click()
        page.wait_for_timeout(300)
        detail = page.locator("#contact-detail-content")
        expect(detail).to_contain_text("alice@test.com")
        expect(detail).to_contain_text("555-1234")
        expect(detail).to_contain_text("Board member")

    def test_detail_shows_name_in_header(self, page, app_url):
        page.goto(app_url)
        wait_for_app(page)
        open_create_contact_modal(page)
        fill_and_submit_contact(page, "Bob Smith", "bob@test.com")
        page.locator("#contact-list .contact-row").first.click()
        page.wait_for_timeout(300)
        expect(page.locator("#contact-detail-title")).to_have_text("Bob Smith")

    def test_detail_shows_no_notes_when_empty(self, page, app_url):
        page.goto(app_url)
        wait_for_app(page)
        open_create_contact_modal(page)
        fill_and_submit_contact(page, "NoNotes", "nn@test.com")
        page.locator("#contact-list .contact-row").first.click()
        page.wait_for_timeout(300)
        expect(page.locator("#contact-detail-content")).to_contain_text("No notes")

    def test_detail_shows_no_phone_when_empty(self, page, app_url):
        page.goto(app_url)
        wait_for_app(page)
        open_create_contact_modal(page)
        fill_and_submit_contact(page, "NoPhone", "np@test.com")
        page.locator("#contact-list .contact-row").first.click()
        page.wait_for_timeout(300)
        expect(page.locator("#contact-detail-content")).to_contain_text("Not provided")

    def test_detail_has_edit_button(self, page, app_url):
        page.goto(app_url)
        wait_for_app(page)
        open_create_contact_modal(page)
        fill_and_submit_contact(page, "EditMe", "edit@test.com")
        page.locator("#contact-list .contact-row").first.click()
        page.wait_for_timeout(300)
        expect(page.locator("#contact-detail-actions button:text('Edit')")).to_be_visible()

    def test_detail_updates_after_edit(self, page, app_url):
        page.goto(app_url)
        wait_for_app(page)
        open_create_contact_modal(page)
        fill_and_submit_contact(page, "Before", "before@test.com")
        page.locator("#contact-list .contact-row").first.click()
        page.wait_for_timeout(300)
        expect(page.locator("#contact-detail-title")).to_have_text("Before")
        # Edit the contact
        page.locator("#contact-list .contact-edit-btn").first.click()
        page.wait_for_timeout(300)
        page.locator("#ce-name").fill("After")
        page.click("#create-modal button:text('Save')")
        page.wait_for_timeout(500)
        expect(page.locator("#contact-detail-title")).to_have_text("After")

    def test_selected_row_highlighted(self, page, app_url):
        page.goto(app_url)
        wait_for_app(page)
        open_create_contact_modal(page)
        fill_and_submit_contact(page, "A", "a@x.com")
        open_create_contact_modal(page)
        fill_and_submit_contact(page, "B", "b@x.com")
        page.locator("#contact-list .contact-row").first.click()
        page.wait_for_timeout(200)
        expect(page.locator("#contact-list .contact-row.active")).to_have_count(1)

    def test_detail_shows_category(self, page, app_url):
        page.goto(app_url)
        wait_for_app(page)
        open_create_contact_modal(page)
        fill_and_submit_contact(page, "SomePerson", "sp@test.com")
        page.locator("#contact-list .contact-row").first.click()
        page.wait_for_timeout(300)
        expect(page.locator("#contact-detail-content")).to_contain_text("Single")


# ══════════════════════════════════════════════════════════════════════════════
# CREATE CONTACT MODAL
# ══════════════════════════════════════════════════════════════════════════════

class TestCreateContactModal:
    def test_no_category_dropdown(self, page, app_url):
        page.goto(app_url)
        wait_for_app(page)
        open_create_contact_modal(page)
        expect(page.locator("#cm-category")).to_have_count(0)

    def test_has_family_search(self, page, app_url):
        page.goto(app_url)
        wait_for_app(page)
        open_create_contact_modal(page)
        expect(page.locator("#cm-fam-search")).to_be_visible()

    def test_has_group_search(self, page, app_url):
        page.goto(app_url)
        wait_for_app(page)
        open_create_contact_modal(page)
        expect(page.locator("#cm-grp-search")).to_be_visible()

    def test_add_family_in_create(self, page, app_url):
        page.goto(app_url)
        wait_for_app(page)
        # Create family first
        switch_tab(page, "Families")
        open_create_family_modal(page)
        page.fill("#cm-fname", "Smith")
        page.click("#create-modal button:text('Add')")
        page.wait_for_timeout(500)
        # Create contact with family
        switch_tab(page, "Contacts")
        open_create_contact_modal(page)
        page.fill("#cm-name", "John Smith")
        page.fill("#cm-email", "john@smith.com")
        page.fill("#cm-fam-search", "Smith")
        page.wait_for_timeout(300)
        page.locator("#cm-fam-results .esr-item").first.click()
        page.wait_for_timeout(200)
        # Family pill should appear
        expect(page.locator("#cm-fam-list")).to_contain_text("Smith")
        page.click("#create-modal button:text('Add')")
        page.wait_for_timeout(500)
        # Verify family in detail panel
        page.locator("#contact-list .contact-row").first.click()
        page.wait_for_timeout(300)
        expect(page.locator("#contact-detail-content")).to_contain_text("Smith")

    def test_add_group_in_create(self, page, app_url):
        page.goto(app_url)
        wait_for_app(page)
        # Create group first
        switch_tab(page, "Groups")
        open_create_group_modal(page)
        page.fill("#cm-gname", "Choir")
        page.click("#create-modal button:text('Add')")
        page.wait_for_timeout(500)
        # Create contact with group
        switch_tab(page, "Contacts")
        open_create_contact_modal(page)
        page.fill("#cm-name", "Singer", )
        page.fill("#cm-email", "singer@x.com")
        page.fill("#cm-grp-search", "Choir")
        page.wait_for_timeout(300)
        page.locator("#cm-grp-results .esr-item").first.click()
        page.wait_for_timeout(200)
        expect(page.locator("#cm-grp-list")).to_contain_text("Choir")
        page.click("#create-modal button:text('Add')")
        page.wait_for_timeout(500)
        # Verify group in detail panel
        page.locator("#contact-list .contact-row").first.click()
        page.wait_for_timeout(300)
        expect(page.locator("#contact-detail-content")).to_contain_text("Choir")

    def test_remove_family_pill_in_create(self, page, app_url):
        page.goto(app_url)
        wait_for_app(page)
        switch_tab(page, "Families")
        open_create_family_modal(page)
        page.fill("#cm-fname", "Jones")
        page.click("#create-modal button:text('Add')")
        page.wait_for_timeout(500)
        switch_tab(page, "Contacts")
        open_create_contact_modal(page)
        page.fill("#cm-fam-search", "Jones")
        page.wait_for_timeout(300)
        page.locator("#cm-fam-results .esr-item").first.click()
        page.wait_for_timeout(200)
        expect(page.locator("#cm-fam-list")).to_contain_text("Jones")
        # Remove it
        page.locator("#cm-fam-list .remove-x").first.click()
        page.wait_for_timeout(200)
        expect(page.locator("#cm-fam-list")).not_to_contain_text("Jones")

    def test_add_both_family_and_group(self, page, app_url):
        page.goto(app_url)
        wait_for_app(page)
        # Create family and group
        switch_tab(page, "Families")
        open_create_family_modal(page)
        page.fill("#cm-fname", "Wilson")
        page.click("#create-modal button:text('Add')")
        page.wait_for_timeout(500)
        switch_tab(page, "Groups")
        open_create_group_modal(page)
        page.fill("#cm-gname", "Ushers")
        page.click("#create-modal button:text('Add')")
        page.wait_for_timeout(500)
        # Create contact with both
        switch_tab(page, "Contacts")
        open_create_contact_modal(page)
        page.fill("#cm-name", "Amy Wilson")
        page.fill("#cm-email", "amy@wilson.com")
        page.fill("#cm-fam-search", "Wilson")
        page.wait_for_timeout(300)
        page.locator("#cm-fam-results .esr-item").first.click()
        page.wait_for_timeout(200)
        page.fill("#cm-grp-search", "Ushers")
        page.wait_for_timeout(300)
        page.locator("#cm-grp-results .esr-item").first.click()
        page.wait_for_timeout(200)
        page.click("#create-modal button:text('Add')")
        page.wait_for_timeout(500)
        # Verify both in detail panel
        page.locator("#contact-list .contact-row").first.click()
        page.wait_for_timeout(300)
        detail = page.locator("#contact-detail-content")
        expect(detail).to_contain_text("Wilson")
        expect(detail).to_contain_text("Ushers")


# ══════════════════════════════════════════════════════════════════════════════
# CONTACT FILTERS
# ══════════════════════════════════════════════════════════════════════════════

class TestContactFilters:
    def test_no_filters_by_default(self, page, app_url):
        page.goto(app_url)
        wait_for_app(page)
        # No filter chips should be visible
        expect(page.locator(".filter-chip")).to_have_count(0)

    def test_add_filter_button_exists(self, page, app_url):
        page.goto(app_url)
        wait_for_app(page)
        expect(page.locator(".filter-add-btn")).to_be_visible()

    def test_add_filter_shows_menu(self, page, app_url):
        page.goto(app_url)
        wait_for_app(page)
        page.click(".filter-add-btn")
        page.wait_for_timeout(200)
        expect(page.locator("#filter-menu")).to_be_visible()
        expect(page.locator(".filter-menu-item")).to_have_count(3)

    def test_add_optout_filter(self, page, app_url):
        page.goto(app_url)
        wait_for_app(page)
        add_contact_filter(page, "optout")
        expect(page.locator(".filter-chip")).to_have_count(1)
        expect(page.locator("#filter-optout")).to_be_visible()

    def test_add_multiple_filters(self, page, app_url):
        page.goto(app_url)
        wait_for_app(page)
        add_contact_filter(page, "group")
        add_contact_filter(page, "optout")
        expect(page.locator(".filter-chip")).to_have_count(2)

    def test_remove_filter(self, page, app_url):
        page.goto(app_url)
        wait_for_app(page)
        add_contact_filter(page, "optout")
        expect(page.locator(".filter-chip")).to_have_count(1)
        page.locator(".filter-remove").first.click()
        page.wait_for_timeout(200)
        expect(page.locator(".filter-chip")).to_have_count(0)

    def test_remove_filter_clears_filtering(self, page, app_url):
        page.goto(app_url)
        wait_for_app(page)
        open_create_contact_modal(page)
        fill_and_submit_contact(page, "Alice", "alice@x.com")
        open_create_contact_modal(page)
        fill_and_submit_contact(page, "Bob", "bob@x.com")
        # Opt out Bob via edit
        page.locator("#contact-list .contact-edit-btn").nth(1).click()
        page.wait_for_timeout(300)
        page.locator("#ce-optout").check()
        page.click("#create-modal button:text('Save')")
        page.wait_for_timeout(500)
        # Add optout filter and filter to active only
        add_contact_filter(page, "optout")
        page.select_option("#filter-optout", "active")
        page.wait_for_timeout(300)
        expect(page.locator("#contact-list .contact-row:visible")).to_have_count(1)
        # Remove the filter — should show all again
        page.locator(".filter-remove").first.click()
        page.wait_for_timeout(300)
        expect(page.locator("#contact-list .contact-row:visible")).to_have_count(2)

    def test_already_added_filter_grayed_out(self, page, app_url):
        page.goto(app_url)
        wait_for_app(page)
        add_contact_filter(page, "optout")
        page.click(".filter-add-btn")
        page.wait_for_timeout(200)
        opt_item = page.locator(".filter-menu-item[data-filter='optout']")
        expect(opt_item).to_have_class(re.compile("added"))

    def test_filter_menu_closes_on_outside_click(self, page, app_url):
        page.goto(app_url)
        wait_for_app(page)
        page.click(".filter-add-btn")
        page.wait_for_timeout(200)
        expect(page.locator("#filter-menu")).to_be_visible()
        page.click("#contact-list")
        page.wait_for_timeout(200)
        expect(page.locator("#filter-menu")).not_to_be_visible()


# ══════════════════════════════════════════════════════════════════════════════
# FAMILIES TAB
# ══════════════════════════════════════════════════════════════════════════════

class TestFamiliesTab:
    def test_empty_state(self, page, app_url):
        page.goto(app_url)
        wait_for_app(page)
        switch_tab(page, "Families")
        expect(page.locator("#family-list")).to_contain_text("No families")

    def test_create_family(self, page, app_url):
        page.goto(app_url)
        wait_for_app(page)
        switch_tab(page, "Families")
        open_create_family_modal(page)
        page.fill("#cm-fname", "Smith")
        page.click("#create-modal button:text('Add')")
        page.wait_for_timeout(500)
        expect(page.locator("#family-list")).to_contain_text("Smith")

    def test_create_family_validation(self, page, app_url):
        page.goto(app_url)
        wait_for_app(page)
        switch_tab(page, "Families")
        open_create_family_modal(page)
        page.click("#create-modal button:text('Add')")
        toast = get_toast_text(page)
        assert "name" in toast.lower()

    def test_select_family_shows_detail(self, page, app_url):
        page.goto(app_url)
        wait_for_app(page)
        switch_tab(page, "Families")
        open_create_family_modal(page)
        page.fill("#cm-fname", "Doe")
        page.click("#create-modal button:text('Add')")
        page.wait_for_timeout(500)
        page.locator("#family-list .group-item").first.click()
        page.wait_for_timeout(300)
        expect(page.locator("#family-detail-title")).to_contain_text("Doe")

    def test_delete_family(self, page, app_url):
        page.goto(app_url)
        wait_for_app(page)
        switch_tab(page, "Families")
        open_create_family_modal(page)
        page.fill("#cm-fname", "Gone")
        page.click("#create-modal button:text('Add')")
        page.wait_for_timeout(500)
        page.locator("#family-list .group-item").first.click()
        page.wait_for_timeout(300)
        page.on("dialog", lambda d: d.accept())
        page.locator("#family-detail-actions .btn-danger").click()
        page.wait_for_timeout(500)
        expect(page.locator("#family-list")).to_contain_text("No families")

    def test_search_families(self, page, app_url):
        page.goto(app_url)
        wait_for_app(page)
        switch_tab(page, "Families")
        for name in ["Alpha", "Beta"]:
            open_create_family_modal(page)
            page.fill("#cm-fname", name)
            page.click("#create-modal button:text('Add')")
            page.wait_for_timeout(500)
        page.fill("#search-families", "Alpha")
        page.wait_for_timeout(300)
        visible = page.locator("#family-list .group-item:visible")
        expect(visible).to_have_count(1)

    def test_duplicate_family_name_error(self, page, app_url):
        page.goto(app_url)
        wait_for_app(page)
        switch_tab(page, "Families")
        open_create_family_modal(page)
        page.fill("#cm-fname", "Dup")
        page.click("#create-modal button:text('Add')")
        page.wait_for_timeout(500)
        open_create_family_modal(page)
        page.fill("#cm-fname", "Dup")
        page.click("#create-modal button:text('Add')")
        page.wait_for_timeout(300)
        toast = get_toast_text(page)
        assert "unique" in toast.lower() or "error" in toast.lower() or "already" in toast.lower()


# ══════════════════════════════════════════════════════════════════════════════
# GROUPS TAB
# ══════════════════════════════════════════════════════════════════════════════

class TestGroupsTab:
    def test_empty_state(self, page, app_url):
        page.goto(app_url)
        wait_for_app(page)
        switch_tab(page, "Groups")
        expect(page.locator("#group-list")).to_contain_text("No groups")

    def test_create_group(self, page, app_url):
        page.goto(app_url)
        wait_for_app(page)
        switch_tab(page, "Groups")
        open_create_group_modal(page)
        page.fill("#cm-gname", "Youth")
        page.click("#create-modal button:text('Add')")
        page.wait_for_timeout(500)
        expect(page.locator("#group-list")).to_contain_text("Youth")

    def test_select_group_shows_detail(self, page, app_url):
        page.goto(app_url)
        wait_for_app(page)
        switch_tab(page, "Groups")
        open_create_group_modal(page)
        page.fill("#cm-gname", "Choir")
        page.click("#create-modal button:text('Add')")
        page.wait_for_timeout(500)
        page.locator("#group-list .group-item").first.click()
        page.wait_for_timeout(300)
        expect(page.locator("#group-detail-title")).to_contain_text("Choir")

    def test_delete_group(self, page, app_url):
        page.goto(app_url)
        wait_for_app(page)
        switch_tab(page, "Groups")
        open_create_group_modal(page)
        page.fill("#cm-gname", "ToDelete")
        page.click("#create-modal button:text('Add')")
        page.wait_for_timeout(500)
        page.locator("#group-list .group-item").first.click()
        page.wait_for_timeout(300)
        page.on("dialog", lambda d: d.accept())
        page.locator("#group-detail-actions .btn-danger").click()
        page.wait_for_timeout(500)
        expect(page.locator("#group-list")).to_contain_text("No groups")

    def test_search_groups(self, page, app_url):
        page.goto(app_url)
        wait_for_app(page)
        switch_tab(page, "Groups")
        for name in ["Alpha", "Beta"]:
            open_create_group_modal(page)
            page.fill("#cm-gname", name)
            page.click("#create-modal button:text('Add')")
            page.wait_for_timeout(500)
        page.fill("#search-groups", "Beta")
        page.wait_for_timeout(300)
        visible = page.locator("#group-list .group-item:visible")
        expect(visible).to_have_count(1)


# ══════════════════════════════════════════════════════════════════════════════
# SETTINGS TAB
# ══════════════════════════════════════════════════════════════════════════════

class TestSettingsTab:
    def test_default_values(self, page, app_url):
        page.goto(app_url)
        wait_for_app(page)
        switch_tab(page, "Settings")
        # Provider dropdown should default to Gmail
        expect(page.locator("#s-provider")).to_have_value("gmail")
        # Custom SMTP fields should be hidden when a preset is selected
        expect(page.locator("#smtp-custom-fields")).not_to_be_visible()

    def test_save_settings(self, page, app_url):
        page.goto(app_url)
        wait_for_app(page)
        switch_tab(page, "Settings")
        page.fill("#s-email", "test@gmail.com")
        page.fill("#s-password", "myapppass")
        page.fill("#s-sender-name", "My Church")
        page.click("button:text('Save Settings')")
        page.wait_for_timeout(500)
        toast = get_toast_text(page)
        assert "saved" in toast.lower()

    def test_save_timezone(self, page, app_url):
        page.goto(app_url)
        wait_for_app(page)
        switch_tab(page, "Settings")
        page.select_option("#s-timezone", "US/Pacific")
        page.click("button:text('Save Timezone')")
        page.wait_for_timeout(500)
        toast = get_toast_text(page)
        assert "saved" in toast.lower() or "timezone" in toast.lower()

    def test_settings_persist_after_tab_switch(self, page, app_url):
        page.goto(app_url)
        wait_for_app(page)
        switch_tab(page, "Settings")
        page.fill("#s-email", "persist@gmail.com")
        page.fill("#s-password", "pass123")
        page.click("button:text('Save Settings')")
        page.wait_for_timeout(500)
        switch_tab(page, "Contacts")
        switch_tab(page, "Settings")
        page.wait_for_timeout(300)
        expect(page.locator("#s-email")).to_have_value("persist@gmail.com")


# ══════════════════════════════════════════════════════════════════════════════
# EMAIL PROVIDER PRESETS
# ══════════════════════════════════════════════════════════════════════════════

class TestEmailProviderPresets:
    def test_provider_dropdown_exists(self, page, app_url):
        page.goto(app_url)
        wait_for_app(page)
        switch_tab(page, "Settings")
        expect(page.locator("#s-provider")).to_be_visible()

    def test_provider_dropdown_has_options(self, page, app_url):
        page.goto(app_url)
        wait_for_app(page)
        switch_tab(page, "Settings")
        options = page.locator("#s-provider option")
        # placeholder + at least 6 providers
        assert options.count() >= 7

    def test_gmail_selected_by_default(self, page, app_url):
        page.goto(app_url)
        wait_for_app(page)
        switch_tab(page, "Settings")
        expect(page.locator("#s-provider")).to_have_value("gmail")

    def test_custom_fields_hidden_for_preset(self, page, app_url):
        page.goto(app_url)
        wait_for_app(page)
        switch_tab(page, "Settings")
        page.select_option("#s-provider", "outlook")
        page.wait_for_timeout(200)
        expect(page.locator("#smtp-custom-fields")).not_to_be_visible()

    def test_custom_fields_shown_for_custom(self, page, app_url):
        page.goto(app_url)
        wait_for_app(page)
        switch_tab(page, "Settings")
        page.select_option("#s-provider", "custom")
        page.wait_for_timeout(200)
        expect(page.locator("#smtp-custom-fields")).to_be_visible()
        expect(page.locator("#s-smtp-host")).to_be_visible()
        expect(page.locator("#s-smtp-port")).to_be_visible()

    def test_save_with_preset_provider(self, page, app_url):
        page.goto(app_url)
        wait_for_app(page)
        switch_tab(page, "Settings")
        page.select_option("#s-provider", "yahoo")
        page.fill("#s-email", "test@yahoo.com")
        page.fill("#s-password", "yahoopass")
        page.click("button:text('Save Settings')")
        page.wait_for_timeout(500)
        toast = get_toast_text(page)
        assert "saved" in toast.lower()
        # Reload and verify provider persists
        switch_tab(page, "Contacts")
        switch_tab(page, "Settings")
        page.wait_for_timeout(500)
        expect(page.locator("#s-provider")).to_have_value("yahoo")

    def test_save_with_custom_provider(self, page, app_url):
        page.goto(app_url)
        wait_for_app(page)
        switch_tab(page, "Settings")
        page.select_option("#s-provider", "custom")
        page.wait_for_timeout(200)
        page.fill("#s-email", "me@example.com")
        page.fill("#s-password", "mypass")
        page.fill("#s-smtp-host", "mail.example.com")
        page.fill("#s-smtp-port", "465")
        page.click("button:text('Save Settings')")
        page.wait_for_timeout(500)
        toast = get_toast_text(page)
        assert "saved" in toast.lower()
        # Reload and verify custom host persists
        switch_tab(page, "Contacts")
        switch_tab(page, "Settings")
        page.wait_for_timeout(500)
        expect(page.locator("#s-provider")).to_have_value("custom")
        expect(page.locator("#s-smtp-host")).to_have_value("mail.example.com")
        expect(page.locator("#s-smtp-port")).to_have_value("465")

    def test_custom_requires_host(self, page, app_url):
        page.goto(app_url)
        wait_for_app(page)
        switch_tab(page, "Settings")
        page.select_option("#s-provider", "custom")
        page.wait_for_timeout(200)
        page.fill("#s-email", "me@example.com")
        page.fill("#s-password", "mypass")
        # Leave host empty
        page.fill("#s-smtp-host", "")
        page.click("button:text('Save Settings')")
        page.wait_for_timeout(500)
        toast = get_toast_text(page)
        assert "smtp" in toast.lower() or "host" in toast.lower()

    def test_switching_provider_updates_help(self, page, app_url):
        page.goto(app_url)
        wait_for_app(page)
        switch_tab(page, "Settings")
        # Open help panel
        page.click("button:text('? Help')")
        page.wait_for_timeout(200)
        expect(page.locator("#email-help-panel")).to_be_visible()
        # Select Outlook
        page.select_option("#s-provider", "outlook")
        page.wait_for_timeout(200)
        info = page.locator("#email-help-provider-info")
        expect(info).to_be_visible()
        assert "outlook" in info.text_content().lower()


class TestEmailHelp:
    def test_help_button_exists(self, page, app_url):
        page.goto(app_url)
        wait_for_app(page)
        switch_tab(page, "Settings")
        expect(page.locator("button:text('? Help')")).to_be_visible()

    def test_help_panel_hidden_by_default(self, page, app_url):
        page.goto(app_url)
        wait_for_app(page)
        switch_tab(page, "Settings")
        expect(page.locator("#email-help-panel")).not_to_be_visible()

    def test_help_panel_toggles(self, page, app_url):
        page.goto(app_url)
        wait_for_app(page)
        switch_tab(page, "Settings")
        page.click("button:text('? Help')")
        page.wait_for_timeout(200)
        expect(page.locator("#email-help-panel")).to_be_visible()
        # Click again to close
        page.click("button:text('? Help')")
        page.wait_for_timeout(200)
        expect(page.locator("#email-help-panel")).not_to_be_visible()

    def test_help_panel_has_steps(self, page, app_url):
        page.goto(app_url)
        wait_for_app(page)
        switch_tab(page, "Settings")
        page.click("button:text('? Help')")
        page.wait_for_timeout(200)
        steps = page.locator("#email-help-steps li")
        assert steps.count() >= 3

    def test_help_panel_shows_provider_info(self, page, app_url):
        page.goto(app_url)
        wait_for_app(page)
        switch_tab(page, "Settings")
        page.click("button:text('? Help')")
        page.wait_for_timeout(200)
        # Gmail should be shown by default
        info = page.locator("#email-help-provider-info")
        expect(info).to_be_visible()
        assert "gmail" in info.text_content().lower() or "app password" in info.text_content().lower()

    def test_help_close_button(self, page, app_url):
        page.goto(app_url)
        wait_for_app(page)
        switch_tab(page, "Settings")
        page.click("button:text('? Help')")
        page.wait_for_timeout(200)
        expect(page.locator("#email-help-panel")).to_be_visible()
        # Click the X button inside the help panel
        page.locator("#email-help-panel button:text('\u00d7')").click()
        page.wait_for_timeout(200)
        expect(page.locator("#email-help-panel")).not_to_be_visible()


# ══════════════════════════════════════════════════════════════════════════════
# UPDATE CHECK
# ══════════════════════════════════════════════════════════════════════════════

class TestUpdateCheck:
    def test_version_displayed_in_settings(self, page, app_url):
        page.goto(app_url)
        wait_for_app(page)
        switch_tab(page, "Settings")
        version_el = page.locator("#current-version")
        expect(version_el).to_contain_text("Current version: v")

    def test_check_for_updates_button_exists(self, page, app_url):
        page.goto(app_url)
        wait_for_app(page)
        switch_tab(page, "Settings")
        expect(page.locator("button:text('Check for Updates')")).to_be_visible()

    def test_update_banner_hidden_by_default(self, page, app_url):
        page.goto(app_url)
        wait_for_app(page)
        switch_tab(page, "Settings")
        expect(page.locator("#update-banner")).not_to_have_class(re.compile("show"))

    def test_clicking_check_shows_banner(self, page, app_url):
        """Clicking the button should show a result banner (regardless of network)."""
        page.goto(app_url)
        wait_for_app(page)
        switch_tab(page, "Settings")
        page.click("button:text('Check for Updates')")
        page.wait_for_timeout(2000)
        expect(page.locator("#update-banner")).to_have_class(re.compile("show"))

    def test_update_section_exists(self, page, app_url):
        page.goto(app_url)
        wait_for_app(page)
        switch_tab(page, "Settings")
        expect(page.locator("#update-section")).to_be_visible()


# ══════════════════════════════════════════════════════════════════════════════
# FIRST-TIME SETUP BANNER
# ══════════════════════════════════════════════════════════════════════════════

class TestSetupOverlay:
    """Tests for the full-screen first-time setup overlay."""

    def test_overlay_visible_when_no_credentials(self, fresh_setup_db, page, app_url):
        page.goto(app_url)
        wait_for_app(page)
        expect(page.locator("#setup-banner")).to_have_class(re.compile("show"))
        expect(page.locator("#setup-banner")).to_contain_text("Welcome")
        expect(page.locator("#setup-banner")).to_contain_text("email settings")

    def test_overlay_hidden_after_credentials_saved(self, fresh_setup_db, page, app_url, ui_db):
        ui_db.save_settings("me@x.com", "secret123")
        page.goto(app_url)
        wait_for_app(page)
        expect(page.locator("#setup-banner")).not_to_have_class(re.compile("show"))

    def test_go_to_settings_dismisses_overlay_and_navigates(self, fresh_setup_db, page, app_url):
        page.goto(app_url)
        wait_for_app(page)
        page.locator("#setup-banner button:text('Go to Settings')").click()
        page.wait_for_timeout(300)
        expect(page.locator("#setup-banner")).not_to_have_class(re.compile("show"))
        expect(page.locator("#settings-tab")).to_be_visible()
        active = page.locator("button.tab-btn.active")
        expect(active).to_have_text("Settings")

    def test_overlay_dismiss_persists_on_reload(self, fresh_setup_db, page, app_url):
        page.goto(app_url)
        wait_for_app(page)
        page.locator("#setup-banner button:text('Go to Settings')").click()
        page.wait_for_timeout(300)
        page.goto(app_url)
        wait_for_app(page)
        expect(page.locator("#setup-banner")).not_to_have_class(re.compile("show"))

    def test_overlay_blocks_interaction(self, fresh_setup_db, page, app_url):
        """The overlay should cover the full viewport."""
        page.goto(app_url)
        wait_for_app(page)
        expect(page.locator("#setup-banner")).to_have_class(re.compile("show"))
        box = page.locator("#setup-banner").bounding_box()
        assert box["width"] >= 400
        assert box["height"] >= 400

    def test_overlay_hidden_when_dismissed_via_db(self, page, app_url):
        """Default fixture dismisses banner — overlay should not show."""
        page.goto(app_url)
        wait_for_app(page)
        expect(page.locator("#setup-banner")).not_to_have_class(re.compile("show"))

    def test_reminder_not_visible_while_overlay_showing(self, fresh_setup_db, page, app_url):
        """Side reminder should not show when overlay is up."""
        page.goto(app_url)
        wait_for_app(page)
        expect(page.locator("#setup-banner")).to_have_class(re.compile("show"))
        expect(page.locator("#setup-reminder")).not_to_have_class(re.compile("show"))

    def test_dismissing_overlay_shows_reminder(self, fresh_setup_db, page, app_url):
        """After dismissing overlay, side reminder should appear."""
        page.goto(app_url)
        wait_for_app(page)
        page.locator("#setup-banner button:text('Go to Settings')").click()
        page.wait_for_timeout(300)
        expect(page.locator("#setup-banner")).not_to_have_class(re.compile("show"))
        expect(page.locator("#setup-reminder")).to_have_class(re.compile("show"))


class TestSetupReminder:
    """Tests for the persistent side reminder notification."""

    def test_reminder_visible_when_dismissed_but_not_configured(self, page, app_url):
        """Default fixture: dismissed=true, no credentials -> reminder shows."""
        page.goto(app_url)
        wait_for_app(page)
        expect(page.locator("#setup-reminder")).to_have_class(re.compile("show"))
        expect(page.locator("#setup-reminder")).to_contain_text("Email Not Configured")

    def test_reminder_hidden_when_credentials_configured(self, page, app_url, ui_db):
        ui_db.save_settings("me@x.com", "secret123")
        page.goto(app_url)
        wait_for_app(page)
        expect(page.locator("#setup-reminder")).not_to_have_class(re.compile("show"))

    def test_reminder_close_button_hides(self, page, app_url):
        page.goto(app_url)
        wait_for_app(page)
        expect(page.locator("#setup-reminder")).to_have_class(re.compile("show"))
        page.locator("#setup-reminder .reminder-close").click()
        page.wait_for_timeout(300)
        expect(page.locator("#setup-reminder")).not_to_have_class(re.compile("show"))

    def test_reminder_configure_button_goes_to_settings(self, page, app_url):
        page.goto(app_url)
        wait_for_app(page)
        page.locator("#setup-reminder .reminder-btn").click()
        page.wait_for_timeout(300)
        expect(page.locator("#settings-tab")).to_be_visible()
        active = page.locator("button.tab-btn.active")
        expect(active).to_have_text("Settings")

    def test_reminder_reappears_on_reload(self, page, app_url):
        """Closing the reminder is session-only — it comes back on reload."""
        page.goto(app_url)
        wait_for_app(page)
        page.locator("#setup-reminder .reminder-close").click()
        page.wait_for_timeout(300)
        expect(page.locator("#setup-reminder")).not_to_have_class(re.compile("show"))
        page.goto(app_url)
        wait_for_app(page)
        expect(page.locator("#setup-reminder")).to_have_class(re.compile("show"))

    def test_reminder_disappears_after_saving_credentials(self, page, app_url):
        """Saving valid credentials should hide the reminder."""
        page.goto(app_url)
        wait_for_app(page)
        expect(page.locator("#setup-reminder")).to_have_class(re.compile("show"))
        # Navigate to settings and save credentials
        switch_tab(page, "Settings")
        page.fill("#s-email", "test@gmail.com")
        page.fill("#s-password", "myapppass")
        page.click("button:text('Save Settings')")
        page.wait_for_timeout(500)
        expect(page.locator("#setup-reminder")).not_to_have_class(re.compile("show"))

    def test_both_hidden_when_fully_configured(self, fresh_setup_db, page, app_url, ui_db):
        """With credentials saved, neither overlay nor reminder should show."""
        ui_db.save_settings("me@x.com", "secret123")
        page.goto(app_url)
        wait_for_app(page)
        expect(page.locator("#setup-banner")).not_to_have_class(re.compile("show"))
        expect(page.locator("#setup-reminder")).not_to_have_class(re.compile("show"))


# ══════════════════════════════════════════════════════════════════════════════
# THEME TOGGLE
# ══════════════════════════════════════════════════════════════════════════════

class TestThemeToggle:
    def test_default_is_dark(self, page, app_url):
        page.goto(app_url)
        wait_for_app(page)
        expect(page.locator("body")).not_to_have_class(re.compile("light"))
        expect(page.locator("#theme-toggle")).to_have_text("Light Mode")

    def test_toggle_to_light(self, page, app_url):
        page.goto(app_url)
        wait_for_app(page)
        page.click("#theme-toggle")
        expect(page.locator("body")).to_have_class(re.compile("light"))
        expect(page.locator("#theme-toggle")).to_have_text("Dark Mode")

    def test_toggle_back_to_dark(self, page, app_url):
        page.goto(app_url)
        wait_for_app(page)
        page.click("#theme-toggle")
        page.click("#theme-toggle")
        expect(page.locator("body")).not_to_have_class(re.compile("light"))


# ══════════════════════════════════════════════════════════════════════════════
# COMPOSER (RIGHT PANEL)
# ══════════════════════════════════════════════════════════════════════════════

class TestComposer:
    def test_subject_field_exists(self, page, app_url):
        page.goto(app_url)
        wait_for_app(page)
        expect(page.locator("#subject")).to_be_visible()

    def test_editor_exists(self, page, app_url):
        page.goto(app_url)
        wait_for_app(page)
        expect(page.locator("#editor")).to_be_visible()

    def test_cc_bcc_toggle(self, page, app_url):
        page.goto(app_url)
        wait_for_app(page)
        expect(page.locator("#cc-bcc-fields")).to_have_class(re.compile("hidden"))
        page.click("#cc-bcc-toggle")
        expect(page.locator("#cc-bcc-fields")).not_to_have_class(re.compile("hidden"))

    def test_send_without_setup_shows_error(self, page, app_url):
        page.goto(app_url)
        wait_for_app(page)
        page.fill("#subject", "Test Subject")
        # Type into the Quill editor
        page.locator("#editor .ql-editor").fill("Some body text")
        page.click("#btn-send-now")
        page.wait_for_timeout(500)
        toast = get_toast_text(page)
        assert "credentials" in toast.lower() or "recipient" in toast.lower()

    def test_schedule_button_exists(self, page, app_url):
        page.goto(app_url)
        wait_for_app(page)
        expect(page.locator("#btn-schedule")).to_be_visible()

    def test_recurrence_weekly_shows_day_picker(self, page, app_url):
        page.goto(app_url)
        wait_for_app(page)
        page.select_option("#recurrence-type", "weekly")
        expect(page.locator("#day-picker")).to_be_visible()

    def test_recurrence_monthly_shows_day_input(self, page, app_url):
        page.goto(app_url)
        wait_for_app(page)
        page.select_option("#recurrence-type", "monthly")
        expect(page.locator("#monthly-day")).to_be_visible()

    def test_recurrence_once_hides_extras(self, page, app_url):
        page.goto(app_url)
        wait_for_app(page)
        page.select_option("#recurrence-type", "weekly")
        expect(page.locator("#day-picker")).to_be_visible()
        page.select_option("#recurrence-type", "once")
        expect(page.locator("#day-picker")).to_be_hidden()

    def test_template_dropdown(self, page, app_url):
        page.goto(app_url)
        wait_for_app(page)
        expect(page.locator("#template-select")).to_be_visible()

    def test_recipient_input(self, page, app_url):
        page.goto(app_url)
        wait_for_app(page)
        expect(page.locator("#recipient-input")).to_be_visible()

    def test_target_selector_options(self, page, app_url):
        page.goto(app_url)
        wait_for_app(page)
        sel = page.locator("#target-select")
        expect(sel).to_be_visible()
        options = sel.locator("option").all_text_contents()
        assert "All Contacts" in options

    def test_preview_email(self, page, app_url):
        page.goto(app_url)
        wait_for_app(page)
        page.fill("#subject", "Preview Test")
        page.click("button:text('Preview')")
        page.wait_for_timeout(300)
        expect(page.locator("#preview-overlay")).to_have_class(re.compile("show"))
        expect(page.locator("#preview-meta")).to_contain_text("Preview Test")

    def test_close_preview(self, page, app_url):
        page.goto(app_url)
        wait_for_app(page)
        page.fill("#subject", "Test")
        page.click("button:text('Preview')")
        page.wait_for_timeout(300)
        page.locator("#preview-overlay button:text('Close')").click()
        page.wait_for_timeout(300)
        expect(page.locator("#preview-overlay")).not_to_have_class(re.compile("show"))


# ══════════════════════════════════════════════════════════════════════════════
# SCHEDULED TAB
# ══════════════════════════════════════════════════════════════════════════════

class TestScheduledTab:
    def test_shows_calendar(self, page, app_url):
        page.goto(app_url)
        wait_for_app(page)
        switch_tab(page, "Scheduled")
        expect(page.locator("#scheduled-calendar")).to_be_visible()
        expect(page.locator("#sched-month-label")).to_be_visible()

    def test_month_navigation_forward(self, page, app_url):
        page.goto(app_url)
        wait_for_app(page)
        switch_tab(page, "Scheduled")
        initial = page.locator("#sched-month-label").text_content()
        page.locator("#scheduled-tab button:text('Next')").click()
        page.wait_for_timeout(300)
        after = page.locator("#sched-month-label").text_content()
        assert initial != after

    def test_month_navigation_backward(self, page, app_url):
        page.goto(app_url)
        wait_for_app(page)
        switch_tab(page, "Scheduled")
        initial = page.locator("#sched-month-label").text_content()
        page.locator("#scheduled-tab button:text('Next')").click()
        page.wait_for_timeout(200)
        page.locator("#scheduled-tab button:text('Prev')").click()
        page.wait_for_timeout(200)
        back = page.locator("#sched-month-label").text_content()
        assert back == initial


# ══════════════════════════════════════════════════════════════════════════════
# HISTORY TAB
# ══════════════════════════════════════════════════════════════════════════════

class TestHistoryTab:
    def test_empty_state(self, page, app_url):
        page.goto(app_url)
        wait_for_app(page)
        switch_tab(page, "History")
        expect(page.locator("#history-list")).to_contain_text("No email history")

    def test_date_filters_exist(self, page, app_url):
        page.goto(app_url)
        wait_for_app(page)
        switch_tab(page, "History")
        expect(page.locator("#history-start-date")).to_be_visible()
        expect(page.locator("#history-end-date")).to_be_visible()

    def test_search_field_exists(self, page, app_url):
        page.goto(app_url)
        wait_for_app(page)
        switch_tab(page, "History")
        expect(page.locator("#search-history")).to_be_visible()


# ══════════════════════════════════════════════════════════════════════════════
# ANALYTICS TAB
# ══════════════════════════════════════════════════════════════════════════════

class TestAnalyticsTab:
    def test_shows_content(self, page, app_url):
        page.goto(app_url)
        wait_for_app(page)
        switch_tab(page, "Analytics")
        expect(page.locator("#analytics-content")).to_be_visible()

    def test_shows_zero_stats_when_empty(self, page, app_url):
        page.goto(app_url)
        wait_for_app(page)
        switch_tab(page, "Analytics")
        page.wait_for_timeout(500)
        content = page.locator("#analytics-content").text_content()
        assert "0" in content


# ══════════════════════════════════════════════════════════════════════════════
# LAYOUT
# ══════════════════════════════════════════════════════════════════════════════

class TestLayout:
    def test_divider_exists(self, page, app_url):
        page.goto(app_url)
        wait_for_app(page)
        expect(page.locator("#panel-divider")).to_be_visible()

    def test_both_panels_visible(self, page, app_url):
        page.goto(app_url)
        wait_for_app(page)
        expect(page.locator("#left-panel")).to_be_visible()
        expect(page.locator("#right-panel")).to_be_visible()


# ══════════════════════════════════════════════════════════════════════════════
# MODALS
# ══════════════════════════════════════════════════════════════════════════════

class TestModals:
    def test_create_modal_opens_and_closes(self, page, app_url):
        page.goto(app_url)
        wait_for_app(page)
        open_create_contact_modal(page)
        expect(page.locator("#create-modal")).to_have_class(re.compile("show"))
        page.click("#create-modal button:text('Cancel')")
        page.wait_for_timeout(300)
        expect(page.locator("#create-overlay")).not_to_have_class(re.compile("show"))

    def test_cancel_button_closes_modal(self, page, app_url):
        page.goto(app_url)
        wait_for_app(page)
        open_create_contact_modal(page)
        expect(page.locator("#create-modal")).to_have_class(re.compile("show"))
        page.click("#create-modal button:text('Cancel')")
        page.wait_for_timeout(300)
        expect(page.locator("#create-overlay")).not_to_have_class(re.compile("show"))


# ══════════════════════════════════════════════════════════════════════════════
# AUTOCOMPLETE / RECIPIENTS
# ══════════════════════════════════════════════════════════════════════════════

class TestAutocomplete:
    def test_typing_shows_dropdown(self, page, app_url):
        page.goto(app_url)
        wait_for_app(page)
        open_create_contact_modal(page)
        fill_and_submit_contact(page, "Alice", "alice@x.com")
        page.fill("#recipient-input", "Ali")
        page.wait_for_timeout(600)
        expect(page.locator("#ac-dropdown")).to_be_visible()

    def test_selecting_adds_chip(self, page, app_url):
        page.goto(app_url)
        wait_for_app(page)
        open_create_contact_modal(page)
        fill_and_submit_contact(page, "Alice", "alice@x.com")
        page.fill("#recipient-input", "Ali")
        page.wait_for_timeout(600)
        page.locator("#ac-dropdown .ac-item").first.click()
        page.wait_for_timeout(300)
        chips = page.locator("#recipient-chips .recipient-chip")
        expect(chips).to_have_count(1)

    def test_remove_chip(self, page, app_url):
        page.goto(app_url)
        wait_for_app(page)
        open_create_contact_modal(page)
        fill_and_submit_contact(page, "Alice", "alice@x.com")
        page.fill("#recipient-input", "Ali")
        page.wait_for_timeout(600)
        page.locator("#ac-dropdown .ac-item").first.click()
        page.wait_for_timeout(300)
        page.locator("#recipient-chips .remove").first.click()
        page.wait_for_timeout(200)
        chips = page.locator("#recipient-chips .recipient-chip")
        expect(chips).to_have_count(0)

    def test_search_by_email(self, page, app_url):
        page.goto(app_url)
        wait_for_app(page)
        open_create_contact_modal(page)
        fill_and_submit_contact(page, "Alice", "alice@x.com")
        page.fill("#recipient-input", "alice@")
        page.wait_for_timeout(600)
        expect(page.locator("#ac-dropdown")).to_be_visible()


# ══════════════════════════════════════════════════════════════════════════════
# TOAST NOTIFICATIONS
# ══════════════════════════════════════════════════════════════════════════════

class TestToast:
    def test_toast_appears_on_error(self, page, app_url):
        page.goto(app_url)
        wait_for_app(page)
        open_create_contact_modal(page)
        page.click("#create-modal button:text('Add')")
        toast = page.locator("#toast")
        expect(toast).to_have_class(re.compile("show"))

    def test_toast_auto_hides(self, page, app_url):
        page.goto(app_url)
        wait_for_app(page)
        open_create_contact_modal(page)
        page.click("#create-modal button:text('Add')")
        page.wait_for_timeout(4000)
        toast = page.locator("#toast")
        expect(toast).not_to_have_class(re.compile("show"))


# ══════════════════════════════════════════════════════════════════════════════
# BULK OPERATIONS
# ══════════════════════════════════════════════════════════════════════════════

class TestBulkOperations:
    def test_delete_multiple_contacts(self, page, app_url):
        page.goto(app_url)
        wait_for_app(page)
        for i in range(3):
            open_create_contact_modal(page)
            fill_and_submit_contact(page, f"Del{i}", f"del{i}@x.com")
        checkboxes = page.locator("#contact-list input[type='checkbox']")
        for i in range(checkboxes.count()):
            checkboxes.nth(i).check()
        page.on("dialog", lambda d: d.accept())
        page.click("button:text('Delete Selected')")
        page.wait_for_timeout(600)
        expect(page.locator("#contact-list")).to_contain_text("No contacts")

    def test_delete_with_none_selected(self, page, app_url):
        page.goto(app_url)
        wait_for_app(page)
        open_create_contact_modal(page)
        fill_and_submit_contact(page, "Keep", "keep@x.com")
        page.click("button:text('Delete Selected')")
        page.wait_for_timeout(300)
        toast = get_toast_text(page)
        assert "select" in toast.lower()


# ══════════════════════════════════════════════════════════════════════════════
# EDGE CASES
# ══════════════════════════════════════════════════════════════════════════════

class TestEdgeCases:
    def test_special_characters_in_name(self, page, app_url):
        page.goto(app_url)
        wait_for_app(page)
        open_create_contact_modal(page)
        fill_and_submit_contact(page, "O'Brien & <Son>", "obrien@x.com")
        expect(page.locator("#contact-list")).to_contain_text("O'Brien")

    def test_long_email_address(self, page, app_url):
        page.goto(app_url)
        wait_for_app(page)
        long_email = "very.long.email.address.here@subdomain.example.com"
        open_create_contact_modal(page)
        fill_and_submit_contact(page, "LongEmail", long_email)
        expect(page.locator("#contact-list")).to_contain_text("LongEmail")

    def test_rapid_tab_switching(self, page, app_url):
        page.goto(app_url)
        wait_for_app(page)
        for _ in range(3):
            for tab in ["Families", "Groups", "Settings", "Contacts"]:
                switch_tab(page, tab)
        expect(page.locator("#contacts-tab")).to_be_visible()

    def test_page_refresh_persists_data(self, page, app_url):
        page.goto(app_url)
        wait_for_app(page)
        open_create_contact_modal(page)
        fill_and_submit_contact(page, "Persist", "persist@x.com")
        page.reload()
        wait_for_app(page)
        expect(page.locator("#contact-list")).to_contain_text("Persist")

    def test_empty_search_shows_all(self, page, app_url):
        page.goto(app_url)
        wait_for_app(page)
        open_create_contact_modal(page)
        fill_and_submit_contact(page, "Alice", "alice@x.com")
        page.fill("#search-contacts", "zzz")
        page.wait_for_timeout(200)
        page.fill("#search-contacts", "")
        page.wait_for_timeout(200)
        visible = page.locator("#contact-list .contact-row:visible")
        expect(visible).to_have_count(1)

    def test_no_js_errors_on_load(self, page, app_url):
        page.goto(app_url)
        wait_for_app(page)
        title = page.title()
        assert "JS ERR" not in title

    def test_create_contact_with_family_assignment(self, page, app_url):
        """Create a family, then a contact assigned to that family via the create modal."""
        page.goto(app_url)
        wait_for_app(page)
        # Create family first
        switch_tab(page, "Families")
        open_create_family_modal(page)
        page.fill("#cm-fname", "Doe")
        page.click("#create-modal button:text('Add')")
        page.wait_for_timeout(500)
        # Create contact and assign to Doe family
        switch_tab(page, "Contacts")
        open_create_contact_modal(page)
        page.fill("#cm-name", "Jane Doe")
        page.fill("#cm-email", "jane@x.com")
        page.fill("#cm-fam-search", "Doe")
        page.wait_for_timeout(300)
        page.locator("#cm-fam-results .esr-item").first.click()
        page.wait_for_timeout(200)
        page.click("#create-modal button:text('Add')")
        page.wait_for_timeout(500)
        expect(page.locator("#contact-list")).to_contain_text("Jane Doe")
        # Click to verify family in detail panel
        page.locator("#contact-list .contact-row").first.click()
        page.wait_for_timeout(300)
        expect(page.locator("#contact-detail-content")).to_contain_text("Doe")

    def test_create_modal_has_family_and_group_search(self, page, app_url):
        page.goto(app_url)
        wait_for_app(page)
        open_create_contact_modal(page)
        expect(page.locator("#cm-fam-search")).to_be_visible()
        expect(page.locator("#cm-grp-search")).to_be_visible()

    def test_create_modal_no_category_selector(self, page, app_url):
        page.goto(app_url)
        wait_for_app(page)
        open_create_contact_modal(page)
        expect(page.locator("#cm-category")).to_have_count(0)
