"""
Seed script — populates the church roster app with realistic fake data.

Usage:
    cd "Project For Church"
    source venv/bin/activate
    python3 seed.py

Deletes the existing contacts.db and creates a fresh one filled with:
  - 12 families with 2-5 members each
  - 15 singles
  - 6 groups with mixed membership
  - 8 email templates
  - 20 email history entries
  - 5 scheduled emails (3 pending, 1 sent, 1 failed)
  - App settings (email + timezone)
"""

import os
import json
import random
from datetime import datetime, timedelta

import db_manager

# ── Wipe and recreate ────────────────────────────────────────────────────────

DB = db_manager.DB_PATH
if os.path.exists(DB):
    os.remove(DB)
    print(f"Removed old {DB}")

db_manager.init_db()
print("Initialized fresh database")

# ── Helpers ──────────────────────────────────────────────────────────────────

conn = db_manager.get_connection()
cur = conn.cursor()


def insert(table, **kw):
    cols = ", ".join(kw.keys())
    placeholders = ", ".join(["?"] * len(kw))
    cur.execute(f"INSERT INTO {table} ({cols}) VALUES ({placeholders})", tuple(kw.values()))
    return cur.lastrowid


# ── Settings ─────────────────────────────────────────────────────────────────

settings = {
    "sender_email": "gracecommunity.church@gmail.com",
    "app_password": "abcd efgh ijkl mnop",
    "timezone": "US/Eastern",
    "sender_name": "Grace Community Church",
    "smtp_host": "smtp.gmail.com",
    "smtp_port": "587",
}
for k, v in settings.items():
    insert("settings", key=k, value=v)
print("Settings saved")

# ── Families & contacts ──────────────────────────────────────────────────────

FAMILY_DATA = {
    "Johnson":   [("Michael Johnson", "michael.johnson@email.com"),
                  ("Sarah Johnson", "sarah.johnson@email.com"),
                  ("Emma Johnson", "emma.johnson@email.com"),
                  ("Ethan Johnson", "ethan.johnson@email.com")],
    "Williams":  [("Robert Williams", "robert.williams@email.com"),
                  ("Linda Williams", "linda.williams@email.com"),
                  ("James Williams", "james.williams@email.com")],
    "Garcia":    [("Carlos Garcia", "carlos.garcia@email.com"),
                  ("Maria Garcia", "maria.garcia@email.com"),
                  ("Sofia Garcia", "sofia.garcia@email.com"),
                  ("Diego Garcia", "diego.garcia@email.com"),
                  ("Isabella Garcia", "isabella.garcia@email.com")],
    "Brown":     [("David Brown", "david.brown@email.com"),
                  ("Patricia Brown", "patricia.brown@email.com")],
    "Martinez":  [("Jose Martinez", "jose.martinez@email.com"),
                  ("Ana Martinez", "ana.martinez@email.com"),
                  ("Luis Martinez", "luis.martinez@email.com")],
    "Davis":     [("Thomas Davis", "thomas.davis@email.com"),
                  ("Jennifer Davis", "jennifer.davis@email.com"),
                  ("Ryan Davis", "ryan.davis@email.com"),
                  ("Ashley Davis", "ashley.davis@email.com")],
    "Anderson":  [("Mark Anderson", "mark.anderson@email.com"),
                  ("Susan Anderson", "susan.anderson@email.com")],
    "Taylor":    [("Kevin Taylor", "kevin.taylor@email.com"),
                  ("Rebecca Taylor", "rebecca.taylor@email.com"),
                  ("Nathan Taylor", "nathan.taylor@email.com")],
    "Thomas":    [("William Thomas", "william.thomas@email.com"),
                  ("Elizabeth Thomas", "elizabeth.thomas@email.com"),
                  ("Grace Thomas", "grace.thomas@email.com")],
    "Lee":       [("Daniel Lee", "daniel.lee@email.com"),
                  ("Michelle Lee", "michelle.lee@email.com"),
                  ("Andrew Lee", "andrew.lee@email.com"),
                  ("Hannah Lee", "hannah.lee@email.com")],
    "Robinson":  [("Paul Robinson", "paul.robinson@email.com"),
                  ("Karen Robinson", "karen.robinson@email.com")],
    "Clark":     [("Steven Clark", "steven.clark@email.com"),
                  ("Laura Clark", "laura.clark@email.com"),
                  ("Matthew Clark", "matthew.clark@email.com")],
}

SINGLES = [
    ("Rachel Foster", "rachel.foster@email.com"),
    ("Benjamin Hayes", "benjamin.hayes@email.com"),
    ("Olivia Chen", "olivia.chen@email.com"),
    ("Samuel Wright", "samuel.wright@email.com"),
    ("Megan Palmer", "megan.palmer@email.com"),
    ("Christopher Reed", "christopher.reed@email.com"),
    ("Victoria Stone", "victoria.stone@email.com"),
    ("Jonathan Blake", "jonathan.blake@email.com"),
    ("Natalie Cruz", "natalie.cruz@email.com"),
    ("Derek Simmons", "derek.simmons@email.com"),
    ("Alyssa Morgan", "alyssa.morgan@email.com"),
    ("Tyler Brooks", "tyler.brooks@email.com"),
    ("Jasmine Patel", "jasmine.patel@email.com"),
    ("Brandon Nguyen", "brandon.nguyen@email.com"),
    ("Stephanie Kim", "stephanie.kim@email.com"),
]

# Insert families
family_ids = {}
for fname in FAMILY_DATA:
    fid = insert("families", name=fname)
    family_ids[fname] = fid

# Insert family contacts + family_members junction
all_contacts = {}  # name -> id
for fname, members in FAMILY_DATA.items():
    fid = family_ids[fname]
    for name, email in members:
        cid = insert("roster", name=name, email=email, category="Family", family_id=fid)
        insert("family_members", family_id=fid, contact_id=cid)
        all_contacts[name] = cid

# Insert singles
for name, email in SINGLES:
    cid = insert("roster", name=name, email=email, category="Single", family_id=None)
    all_contacts[name] = cid

print(f"Created {len(FAMILY_DATA)} families with {sum(len(m) for m in FAMILY_DATA.values())} members")
print(f"Created {len(SINGLES)} singles")

# ── Groups ───────────────────────────────────────────────────────────────────

GROUPS = {
    "Worship Team": [
        "Sarah Johnson", "Grace Thomas", "Nathan Taylor", "Rachel Foster",
        "Victoria Stone", "Sofia Garcia", "Hannah Lee",
    ],
    "Youth Ministry": [
        "Emma Johnson", "Ethan Johnson", "James Williams", "Sofia Garcia",
        "Diego Garcia", "Ryan Davis", "Ashley Davis", "Nathan Taylor",
        "Andrew Lee", "Hannah Lee", "Matthew Clark", "Olivia Chen",
    ],
    "Bible Study - Tuesday": [
        "Michael Johnson", "Robert Williams", "David Brown", "Jose Martinez",
        "Thomas Davis", "Mark Anderson", "Benjamin Hayes", "Derek Simmons",
    ],
    "Bible Study - Thursday": [
        "Sarah Johnson", "Linda Williams", "Patricia Brown", "Ana Martinez",
        "Jennifer Davis", "Susan Anderson", "Rebecca Taylor", "Megan Palmer",
        "Alyssa Morgan",
    ],
    "Outreach Committee": [
        "Carlos Garcia", "Kevin Taylor", "Paul Robinson", "Steven Clark",
        "Christopher Reed", "Jonathan Blake", "Brandon Nguyen",
    ],
    "Sunday School Teachers": [
        "Elizabeth Thomas", "Karen Robinson", "Laura Clark", "Michelle Lee",
        "Natalie Cruz", "Jasmine Patel", "Stephanie Kim",
    ],
}

for gname, members in GROUPS.items():
    gid = insert("groups_", name=gname)
    for mname in members:
        cid = all_contacts.get(mname)
        if cid:
            insert("group_members", group_id=gid, contact_id=cid)

print(f"Created {len(GROUPS)} groups")

# ── Email templates ──────────────────────────────────────────────────────────

TEMPLATES = [
    {
        "name": "Weekly Announcements",
        "subject": "This Week at Grace Community",
        "html_body": "<h2>This Week at Grace Community</h2><p>Dear Church Family,</p><p>Here are the announcements for this week:</p><ul><li><strong>Sunday Service:</strong> 10:00 AM — Pastor Mike will continue the series on Ephesians</li><li><strong>Wednesday Prayer Meeting:</strong> 7:00 PM in the Fellowship Hall</li><li><strong>Youth Group:</strong> Friday 6:30 PM</li></ul><p>We look forward to seeing you!</p><p>Blessings,<br>Grace Community Church</p>",
        "recipients": json.dumps([{"type": "all", "value": "all", "label": "All Contacts"}]),
    },
    {
        "name": "Youth Event Invite",
        "subject": "Youth Group Special Event This Friday!",
        "html_body": "<h2>Youth Group Movie Night</h2><p>Hey everyone!</p><p>This Friday we're having a <strong>movie night</strong> at the church! Bring your friends, snacks, and a blanket.</p><p><strong>When:</strong> Friday, 6:30 PM<br><strong>Where:</strong> Fellowship Hall<br><strong>What to bring:</strong> A friend and your favorite snack to share</p><p>See you there!</p>",
        "recipients": json.dumps([{"type": "group", "value": "group:2", "label": "Group: Youth Ministry"}]),
    },
    {
        "name": "Potluck Reminder",
        "subject": "Church Potluck This Sunday!",
        "html_body": "<h2>Monthly Potluck Lunch</h2><p>Don't forget — this Sunday after the morning service we'll be having our monthly potluck!</p><p>If your last name starts with <strong>A-L</strong>, please bring a <em>main dish</em>.<br>If your last name starts with <strong>M-Z</strong>, please bring a <em>side dish or dessert</em>.</p><p>Paper goods and drinks will be provided.</p>",
        "recipients": json.dumps([{"type": "all", "value": "all", "label": "All Contacts"}]),
    },
    {
        "name": "Volunteer Sign-Up",
        "subject": "Volunteers Needed — Summer VBS",
        "html_body": "<h2>Summer VBS Volunteers</h2><p>We're gearing up for Vacation Bible School and we need YOUR help!</p><p>Dates: <strong>July 14-18</strong><br>Times: <strong>9:00 AM - 12:00 PM</strong></p><p>We need volunteers for:</p><ul><li>Station leaders</li><li>Snack prep</li><li>Music & worship</li><li>Registration</li></ul><p>Reply to this email if you can help!</p>",
        "recipients": json.dumps([]),
    },
    {
        "name": "Prayer Request Update",
        "subject": "Prayer Requests — Please Keep These in Prayer",
        "html_body": "<h2>Weekly Prayer Requests</h2><p>Please keep the following in your prayers this week:</p><ul><li>The Garcia family as they welcome a new baby</li><li>Mark Anderson recovering from surgery</li><li>Our missionaries overseas</li><li>Upcoming church building renovation</li></ul><p><em>\"Do not be anxious about anything, but in every situation, by prayer and petition, with thanksgiving, present your requests to God.\"</em> — Philippians 4:6</p>",
        "recipients": json.dumps([]),
    },
    {
        "name": "Bible Study Reminder",
        "subject": "Bible Study This Week — Don't Miss It!",
        "html_body": "<h2>Bible Study Reminder</h2><p>Just a reminder that Bible Study meets this week:</p><p><strong>Tuesday Men's Group:</strong> 7:00 PM at the church<br><strong>Thursday Women's Group:</strong> 10:00 AM at the church</p><p>We'll be continuing our study of the book of James. Please read chapters 3-4 beforehand.</p><p>See you there!</p>",
        "recipients": json.dumps([]),
    },
    {
        "name": "Christmas Service Invite",
        "subject": "You're Invited — Christmas Eve Service",
        "html_body": "<h2>Christmas Eve Candlelight Service</h2><p>Join us for a special evening of worship and celebration!</p><p><strong>Date:</strong> December 24<br><strong>Time:</strong> 7:00 PM<br><strong>Location:</strong> Main Sanctuary</p><p>The service will include carols, a short message, candle lighting, and communion. <strong>Invite your friends and neighbors!</strong></p><p>Childcare will be available for ages 0-3.</p>",
        "recipients": json.dumps([{"type": "all", "value": "all", "label": "All Contacts"}]),
    },
    {
        "name": "Giving Update",
        "subject": "Quarterly Giving Update — Thank You!",
        "html_body": "<h2>Quarterly Giving Report</h2><p>Dear Church Family,</p><p>Thank you for your generous giving this quarter. Here's a brief update:</p><ul><li><strong>General Fund:</strong> On track with budget</li><li><strong>Missions:</strong> 110% of quarterly goal</li><li><strong>Building Fund:</strong> 85% of renovation target</li></ul><p>Your faithfulness makes a difference. If you have questions about giving, please contact the church office.</p>",
        "recipients": json.dumps([]),
    },
]

for t in TEMPLATES:
    insert("email_templates", name=t["name"], subject=t["subject"], html_body=t["html_body"], recipients=t["recipients"])

print(f"Created {len(TEMPLATES)} email templates")

# ── Email history ────────────────────────────────────────────────────────────

now = datetime.now()
HISTORY = [
    ("This Week at Grace Community", "All Contacts", 55, 54, 1, now - timedelta(days=1, hours=3)),
    ("Youth Group Special Event This Friday!", "Group: Youth Ministry", 12, 12, 0, now - timedelta(days=2, hours=5)),
    ("Church Potluck This Sunday!", "All Contacts", 55, 55, 0, now - timedelta(days=5)),
    ("Prayer Requests — Please Keep These in Prayer", "All Contacts", 55, 53, 2, now - timedelta(days=7)),
    ("Bible Study This Week — Don't Miss It!", "Group: Bible Study - Tuesday", 8, 8, 0, now - timedelta(days=8)),
    ("Bible Study This Week — Don't Miss It!", "Group: Bible Study - Thursday", 9, 9, 0, now - timedelta(days=8)),
    ("This Week at Grace Community", "All Contacts", 55, 55, 0, now - timedelta(days=8, hours=2)),
    ("Volunteers Needed — Summer VBS", "All Contacts", 55, 54, 1, now - timedelta(days=10)),
    ("Sunday School Update", "Group: Sunday School Teachers", 7, 7, 0, now - timedelta(days=12)),
    ("This Week at Grace Community", "All Contacts", 55, 55, 0, now - timedelta(days=15)),
    ("Outreach Planning Meeting", "Group: Outreach Committee", 7, 7, 0, now - timedelta(days=16)),
    ("This Week at Grace Community", "All Contacts", 55, 54, 1, now - timedelta(days=22)),
    ("Prayer Requests — Please Keep These in Prayer", "All Contacts", 55, 55, 0, now - timedelta(days=23)),
    ("Quarterly Giving Update — Thank You!", "All Contacts", 55, 55, 0, now - timedelta(days=28)),
    ("Youth Group Lock-In Registration", "Group: Youth Ministry", 12, 12, 0, now - timedelta(days=30)),
    ("This Week at Grace Community", "All Contacts", 55, 55, 0, now - timedelta(days=29)),
    ("Church Workday — All Hands on Deck", "All Contacts", 55, 50, 5, now - timedelta(days=35)),
    ("This Week at Grace Community", "All Contacts", 55, 55, 0, now - timedelta(days=36)),
    ("Easter Service Times", "All Contacts", 55, 55, 0, now - timedelta(days=45)),
    ("Women's Retreat Registration", "Group: Bible Study - Thursday", 9, 9, 0, now - timedelta(days=50)),
]

FAIL_REASONS = [
    "Connection refused by remote host",
    "Mailbox full — user over quota",
    "Invalid recipient address",
    "SMTP timeout after 30 seconds",
    "550 User not found",
]

contact_list = [(name, name.lower().replace(" ", ".") + "@email.com") for name in all_contacts.keys()]

for subj, target, rcpt, sent, failed, ts in HISTORY:
    hid = insert("email_history", subject=subj, target_description=target,
                 recipient_count=rcpt, sent_count=sent, failed_count=failed,
                 sent_at=ts.strftime("%Y-%m-%d %H:%M:%S"))
    # Generate recipient details
    sample = random.sample(contact_list, min(rcpt, len(contact_list)))
    fail_indices = set(random.sample(range(len(sample)), min(failed, len(sample)))) if failed > 0 else set()
    for i, (cname, cemail) in enumerate(sample):
        if i in fail_indices:
            insert("email_history_details", history_id=hid, recipient_name=cname,
                   recipient_email=cemail, status="failed",
                   error_message=random.choice(FAIL_REASONS))
        else:
            insert("email_history_details", history_id=hid, recipient_name=cname,
                   recipient_email=cemail, status="sent", error_message=None)

print(f"Created {len(HISTORY)} email history entries with recipient details")

# ── Scheduled emails ─────────────────────────────────────────────────────────

all_ids = list(all_contacts.values())

SCHEDULED = [
    {
        "subject": "This Week at Grace Community",
        "html_body": "<h2>This Week at Grace Community</h2><p>Announcements coming soon...</p>",
        "plain_text": "This Week at Grace Community\nAnnouncements coming soon...",
        "target_type": "all", "target_id": None,
        "contact_ids": json.dumps(all_ids),
        "attachment_paths": json.dumps([]),
        "scheduled_at": (now + timedelta(days=2, hours=9)).strftime("%Y-%m-%d %H:%M:%S"),
        "status": "pending",
        "created_at": now.strftime("%Y-%m-%d %H:%M:%S"),
        "recurrence": json.dumps({"type": "weekly", "days": [0], "end": None}),
        "manual_emails": json.dumps([]),
    },
    {
        "subject": "Youth Group Reminder",
        "html_body": "<h2>Youth Group This Friday!</h2><p>Don't forget — 6:30 PM at the church.</p>",
        "plain_text": "Youth Group This Friday!\nDon't forget - 6:30 PM at the church.",
        "target_type": "group", "target_id": 2,
        "contact_ids": json.dumps([all_contacts[n] for n in GROUPS["Youth Ministry"] if n in all_contacts]),
        "attachment_paths": json.dumps([]),
        "scheduled_at": (now + timedelta(days=4, hours=10)).strftime("%Y-%m-%d %H:%M:%S"),
        "status": "pending",
        "created_at": now.strftime("%Y-%m-%d %H:%M:%S"),
        "recurrence": json.dumps({"type": "weekly", "days": [3], "end": None}),
        "manual_emails": json.dumps([]),
    },
    {
        "subject": "Prayer Meeting Reminder",
        "html_body": "<p>Prayer meeting tonight at 7:00 PM. See you there!</p>",
        "plain_text": "Prayer meeting tonight at 7:00 PM. See you there!",
        "target_type": "all", "target_id": None,
        "contact_ids": json.dumps(all_ids),
        "attachment_paths": json.dumps([]),
        "scheduled_at": (now + timedelta(days=1, hours=8)).strftime("%Y-%m-%d %H:%M:%S"),
        "status": "pending",
        "created_at": now.strftime("%Y-%m-%d %H:%M:%S"),
        "recurrence": json.dumps({"type": "weekly", "days": [2], "end": None}),
        "manual_emails": json.dumps([]),
    },
    {
        "subject": "This Week at Grace Community",
        "html_body": "<h2>Last week's announcements</h2><p>...</p>",
        "plain_text": "Last week's announcements...",
        "target_type": "all", "target_id": None,
        "contact_ids": json.dumps(all_ids),
        "attachment_paths": json.dumps([]),
        "scheduled_at": (now - timedelta(days=5, hours=3)).strftime("%Y-%m-%d %H:%M:%S"),
        "status": "sent",
        "created_at": (now - timedelta(days=6)).strftime("%Y-%m-%d %H:%M:%S"),
        "sent_at": (now - timedelta(days=5, hours=3)).strftime("%Y-%m-%d %H:%M:%S"),
        "recurrence": None,
        "manual_emails": json.dumps([]),
    },
    {
        "subject": "Men's Breakfast Invite",
        "html_body": "<p>Men's breakfast this Saturday at 8 AM!</p>",
        "plain_text": "Men's breakfast this Saturday at 8 AM!",
        "target_type": "group", "target_id": 3,
        "contact_ids": json.dumps([all_contacts[n] for n in GROUPS["Bible Study - Tuesday"] if n in all_contacts]),
        "attachment_paths": json.dumps([]),
        "scheduled_at": (now - timedelta(days=2)).strftime("%Y-%m-%d %H:%M:%S"),
        "status": "failed",
        "created_at": (now - timedelta(days=3)).strftime("%Y-%m-%d %H:%M:%S"),
        "result": "SMTP connection failed: invalid credentials",
        "recurrence": None,
        "manual_emails": json.dumps([]),
    },
]

for s in SCHEDULED:
    cols = ", ".join(s.keys())
    placeholders = ", ".join(["?"] * len(s))
    cur.execute(f"INSERT INTO scheduled_emails ({cols}) VALUES ({placeholders})", tuple(s.values()))

print(f"Created {len(SCHEDULED)} scheduled emails")

# ── Commit & done ────────────────────────────────────────────────────────────

conn.commit()
conn.close()

total_contacts = sum(len(m) for m in FAMILY_DATA.values()) + len(SINGLES)
print(f"\nDone! Database seeded with {total_contacts} contacts, {len(FAMILY_DATA)} families, {len(GROUPS)} groups.")
print(f"Run the app:  source venv/bin/activate && python3 main.py")
