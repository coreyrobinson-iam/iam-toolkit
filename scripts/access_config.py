"""
Shared configuration for the IAM Toolkit.

Both build_tracker.py and build_qr.py read from this file, so the roles,
approvers, approval rules and email templates stay identical in the
workbook and the quick reference. Edit this file for your organization.

Text can use {TOKENS}. fill() swaps each token for its value below.
"""

# ---------------- Application ----------------
APP = 'ExampleApp'                      # the licensed SaaS app
APP_URL = 'https://app.example.com/'    # where users sign in
ACCESS_GROUP = 'ExampleApp-Users'       # Entra ID / AD group that grants a seat
# How long a group change takes to reach the app, as it should read in emails.
PROVISIONING_DELAY = 'Allow about 30 minutes for the change to sync, then sign in'
PROVISIONING_MAX = 'If it still does not work after an hour, reply and let me know.'

# ---------------- Teams and owners ----------------
ORG = '[Organization]'
DEPT = '[Department]'
REQUEST_HANDLER = 'Access Coordinator'  # person or team working the requests
IDENTITY_ADMIN = 'Identity Admin'       # team that changes group membership
SERVICE_DESK = 'Service Desk'           # first line of support
INTAKE_MAILBOX = '[access-requests mailbox]'
REVIEW_OWNER = '[Access review owner]'  # runs the periodic access review

# ---------------- Roles ----------------
# Requester roles and how each is approved. The first role is pre-approved.
ROLE_PREAPPROVED = 'Pre-approved role'  # e.g. a job family licensed by default
ROLE_EMPLOYEE = 'Employee'
ROLE_CONTRACTOR = 'Contractor'
ROLE_VENDOR = 'Vendor/partner'
ROLES = [ROLE_PREAPPROVED, ROLE_EMPLOYEE, ROLE_CONTRACTOR, ROLE_VENDOR]

# ---------------- Approvers ----------------
APPROVER_CONTRACTOR = '[Contractor sponsor]'
APPROVER_VENDOR = '[Vendor relationship owner]'
APPROVER_DEFAULT = '[Default approver]'  # employees whose team has no approver
TEAM_APPROVERS = [                       # (team, code, team approver)
    ('Example Team', 'EXT', '[Team approver name]'),
    ('Second Example Team', 'EX2', '[Team approver name]'),
]

# ---------------- Training gate ----------------
TRAINING_REQUIRED_FOR_NEW = True
TRAINING_LOCATION = '[Learning platform]'
TRAININGS = [                            # (name, url) - complete any one
    ('Required training: option 1', 'https://example.com/training-1'),
    ('Required training: option 2', 'https://example.com/training-2'),
]
GUIDE_LINE = 'See the [user guide] for tips and upcoming training sessions: [guide link].'

# ---------------- Access review ----------------
REVIEW_FREQUENCY = 'Every 2 weeks'
INACTIVITY_RULE = ('Remove users who have held access for more than 20 days '
                   'and show no sign-in in the last 20 days.')

# ---------------- Approval rules ----------------
# Who approves each role. Same approver for new and restore requests.
APPROVAL_RULES = [
    (ROLE_PREAPPROVED, 'No approval needed', 'Pre-approved by role.'),
    (ROLE_EMPLOYEE, "Requester's team approver",
     f'From the Team Approvers list. Teams not listed go to {APPROVER_DEFAULT}.'),
    (ROLE_CONTRACTOR, APPROVER_CONTRACTOR, 'Sponsor confirms business need and contract end date.'),
    (ROLE_VENDOR, APPROVER_VENDOR, 'Relationship owner confirms need and scope.'),
]

NEW_ACCESS_STEPS = [
    'Confirm training. If no certificate of completion is on file, send <b>Template F</b> and wait.',
    'Check the role. <b>{ROLE_PREAPPROVED}:</b> skip to step 4.',
    '<b>All other roles:</b> send the requester <b>Template A</b> and the approver <b>Template B</b>. Wait for a decision.',
    'If approved, ask {IDENTITY_ADMIN} to add the user to <b>{ACCESS_GROUP}</b>. If declined, tell the requester and close.',
    'Once added, send <b>Template C</b> and record the date access was confirmed.',
]

RESTORE_STEPS = [
    'Send <b>Template D</b>. It explains the removal and asks whether they expect regular use.',
    'No current need: close the request. They can ask again later.',
    'Need confirmed: <b>{ROLE_PREAPPROVED}</b> goes straight to {IDENTITY_ADMIN}. All other roles need approval first (same approvers as new access).',
    'Once restored to <b>{ACCESS_GROUP}</b>, send <b>Template E</b>. No training needed.',
]

# ---------------- Email templates ----------------
TEMPLATES = [
    ('A', 'To requester: approval needed',
     'Thanks for your interest in {APP}. Licenses are limited, so this request needs approval. '
     "I'll follow up as soon as I hear back."),
    ('B', 'To approver',
     '[Name] ([Role], [Team]) has requested {APP} access. Do you approve?'),
    ('C', 'Access granted (new)',
     'Your {APP} access is set up. {PROVISIONING_DELAY} at {APP_URL} with your work email. '
     '{PROVISIONING_MAX}\n{GUIDE_LINE}\n'
     'Accounts are removed automatically after a period of inactivity.'),
    ('D', 'Access was removed (restore)',
     "To make the best use of a limited number of {APP} licenses, we regularly reassign seats that haven't been used. "
     'Your account showed no recent activity, so access was removed.\n'
     "Do you expect to use {APP} regularly for current work? If so, reply and we'll look at restoring it. "
     "If not, you're welcome to ask again whenever a need comes up."),
    ('E', 'Access restored',
     'Your {APP} access has been restored. You can sign in at {APP_URL} with your work email.\n{GUIDE_LINE}'),
    ('F', 'Training required (new)',
     'Thanks for your interest in {APP}. Before we set up access, please complete one of the required trainings in {TRAINING_LOCATION}:\n'
     '{TRAINING_LIST}\n'
     'When you finish, reply with your certificate of completion and we will get you set up.'),
]


def fill(text, newline='\n'):
    """Replace {TOKENS} with config values. Unknown braces are left alone."""
    tokens = {k: v for k, v in globals().items() if k.isupper() and isinstance(v, str)}
    tokens['TRAINING_LIST'] = newline.join(f'- {n}: {u}' for n, u in TRAININGS)
    out = text
    for _ in range(2):  # second pass catches tokens inside token values
        for k, v in tokens.items():
            out = out.replace('{' + k + '}', v)
    return out.replace('\n', newline)
