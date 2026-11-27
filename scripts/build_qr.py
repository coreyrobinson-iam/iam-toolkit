"""
Access Requests Quick Reference generator (generic, scrubbed).
Builds a 2-page PDF quick reference for staff who handle app access requests.
Usage: python build_qr.py "Quick Reference.pdf"   (requires reportlab)
Edit the CONFIG block for your organization.
"""
import sys
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, KeepTogether
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.lib.units import inch

OUT = sys.argv[1] if len(sys.argv) > 1 else 'Quick Reference.pdf'
APP = 'ExampleApp'                        # app name used in labels and templates
APP_URL = 'https://app.example.com/'
# ================= CONFIG =================
ORG_LINE = '[Department] | [Organization] | For staff handling {APP} provisioning | Revised [date]'
FOOTER = '{APP} Access Requests: Quick Reference | Revised [date]'
INTAKE_MAILBOX = '[team mailbox]'
LMS = '[LMS name]'
APPROVER_CONTRACT = '[Contract attorney approver]'
VENDOR_ROLE = 'Vendor/affiliate'
APPROVER_VENDOR = '[Vendor/affiliate approver]'
APPROVER_DEFAULT = '[Default approver]'
REMOVAL_OWNER = '[Removal owner]'
COPIED_TEAM = '[Your team]'
TRAIN_URL_1 = 'https://example.com/training-1'
TRAIN_URL_2 = 'https://example.com/training-2'
GUIDE_LINE = 'Please see our [guide name] for more information, including upcoming and previously recorded training sessions: [guide link].'
LIAISONS = [('Example Team (EXT)', '[Liaison name]'), ('Second Example Team (EX2)', '[Liaison name]')]
# ==========================================
NAVY = colors.HexColor('#1F3864')
GRID = colors.HexColor('#BFBFBF')
ss = getSampleStyleSheet()
base = ParagraphStyle('b', parent=ss['Normal'], fontName='Helvetica', fontSize=9.5, leading=12.5)
small = ParagraphStyle('s', parent=base, fontSize=8.5, leading=11)
title = ParagraphStyle('t', parent=base, fontName='Helvetica-Bold', fontSize=16, leading=20, textColor=NAVY)
sub = ParagraphStyle('st', parent=base, fontSize=9, textColor=colors.HexColor('#595959'))
h = ParagraphStyle('h', parent=base, fontName='Helvetica-Bold', fontSize=11.5, leading=15, textColor=NAVY, spaceBefore=10, spaceAfter=4)
th = ParagraphStyle('th', parent=small, fontName='Helvetica-Bold', textColor=colors.white)
tmpl_h = ParagraphStyle('tmh', parent=base, fontName='Helvetica-Bold', spaceBefore=6, spaceAfter=2)
quote = ParagraphStyle('q', parent=base, leftIndent=12, textColor=colors.HexColor('#262626'))

def fill(t): return t.replace('{APP_URL}', APP_URL).replace('{APP}', APP)
def P(t, s=base): return Paragraph(fill(t), s)
def bullets(items, s=base):
    return [P('&bull;&nbsp; ' + i, ParagraphStyle('bl', parent=s, leftIndent=12, firstLineIndent=-9, spaceAfter=2)) for i in items]
def steps(items):
    return [P(f'{n}.&nbsp; ' + i, ParagraphStyle('st2', parent=base, leftIndent=14, firstLineIndent=-12, spaceAfter=3)) for n, i in enumerate(items, 1)]
def table(data, widths, header=True):
    rows = [[P(c, th if (header and r == 0) else small) for c in row] for r, row in enumerate(data)]
    t = Table(rows, colWidths=widths, repeatRows=1 if header else 0)
    st = [('GRID', (0, 0), (-1, -1), 0.5, GRID), ('VALIGN', (0, 0), (-1, -1), 'TOP'),
          ('LEFTPADDING', (0, 0), (-1, -1), 5), ('RIGHTPADDING', (0, 0), (-1, -1), 5),
          ('TOPPADDING', (0, 0), (-1, -1), 3), ('BOTTOMPADDING', (0, 0), (-1, -1), 3)]
    if header:
        st.append(('BACKGROUND', (0, 0), (-1, 0), NAVY))
    t.setStyle(TableStyle(st))
    return t

LIT = TRAIN_URL_1
TRN = TRAIN_URL_2
LIB = GUIDE_LINE
link = lambda u, t: f'<link href="{u}" color="#1F4E9E"><u>{t}</u></link>'

s = []
s += [P('{APP} Access Requests: Quick Reference', title),
      P(ORG_LINE, sub),
      Spacer(1, 6)]

s += [P('1. Intake', h)] + bullets([
    f'The Help Desk forwards all {APP} questions to the {INTAKE_MAILBOX}.',
    'First question for every request: <b>Is this a new request, or a request to restore access?</b>',
    '<b>New access requires training:</b> a certificate of completion for one of the two {APP} training recordings in ' + LMS + '. Restores do not require training.',
])

s += [P('2. Who approves?', h), table([
    ['Requester', 'New access', 'Restore access'],
    ['Regular (non-contract) attorney', 'Approved automatically. No approval needed.', 'Approved automatically. Ask if they expect regular use first (Template D).'],
    ['Regular staff (non-attorney)', f"AI liaison for the requester's team (or {APPROVER_DEFAULT} if the team has no liaison)", f"AI liaison for the requester's team (or {APPROVER_DEFAULT} if the team has no liaison)"],
    ['Contract attorney', APPROVER_CONTRACT, APPROVER_CONTRACT],
    [VENDOR_ROLE, APPROVER_VENDOR, APPROVER_VENDOR],
], [1.9 * inch, 2.4 * inch, 2.7 * inch])]

s += [P('3. New access: steps', h)] + steps([
    'Confirm training. The requester must send a certificate of completion for one of the two training recordings. If not received, send <b>Template F</b> and wait for the certificate.',
    '<b>Attorney:</b> Send to User Admin (or reply to the Help Desk) that the user is approved and should be added to the {APP} group. Go to step 5.',
    '<b>Staff, contract attorney, or ' + VENDOR_ROLE + ':</b> Send the requester <b>Template A</b>. Send the approver <b>Template B</b>. Wait for a reply.',
    'If approved, send to User Admin to add the user to the {APP} group. If declined, tell the requester.',
    'Once added, send the requester <b>Template C</b>. It includes the inactivity notice.',
])

s += [P('4. Restore access: steps', h)] + steps([
    'Send <b>Template D</b> to explain why access was removed and ask whether they expect regular use.',
    '<b>Attorney who confirms a need:</b> request re-provisioning through User Admin. <b>Staff, contract attorney, or ' + VENDOR_ROLE + ':</b> get approval first (see section 2).',
    'Once re-added, send <b>Template E</b>. No training required.',
])

s += [P('5. Biweekly deprovisioning (internal)', h)] + bullets([
    f'<b>Run by:</b> {REMOVAL_OWNER}. {COPIED_TEAM} is copied on each removal request to User Admin.',
    'Every 2 weeks: remove users who have had access for more than 20 days and show no activity in the last 20 days.',
    'Messages to users do not state the inactivity period. Every new user is notified in Template C that accounts are deprovisioned after periods of inactivity.',
])

s += [P('6. Email templates', h)]
temps = [
    ('Template A: To requester (approval needed)',
     'Thanks for your interest in {APP}. Since {APP} licenses are limited and carry a significant cost, we need to seek approval for this request. I will get back to you once I hear back.'),
    ('Template B: To approver', '[Name] has requested {APP} access. Do you approve?'),
    ('Template C: Access created',
     'I have created {APP} access for you. Wait about 30 minutes and try logging in at {APP_URL} with your email address. It may take up to an hour for you to be able to access.<br/>'
     + LIB + '<br/>Please note that accounts are automatically deprovisioned after set periods of inactivity.'),
    ('Template D: Access was removed',
     "To make the best use of a limited number of {APP} licenses, the firm periodically reassigns seats that haven't been used recently. Our records showed no recent activity on your account, so your access was deactivated.<br/>"
     "Do you expect to use {APP} regularly for current or upcoming work? If so, please let me know and we'll look at restoring your access. If not, you're welcome to reach out whenever a need comes up."),
    ('Template E: Access restored',
     'I have restored {APP} access for you. You should now be able to log in at {APP_URL} with your email address.<br/>' + LIB),
    ('Template F: Training required',
     'Thanks for your interest in {APP}. Before we can set up your access, please complete one of the two required {APP} training recordings in ' + LMS + ':<br/>'
     f'&bull; {link(LIT, "{APP} Litigation Training")}<br/>&bull; {link(TRN, "{APP} Transactional Training")}<br/>'
     'Once you\'ve finished, reply to this email with your certificate of completion and we can get you set up.'),
]
for t, body in temps:
    s.append(KeepTogether([P(t, tmpl_h), P(body, quote)]))

s += [P('7. AI liaisons', h),
      P('Source: [liaison chart]. <b>For any team without a listed liaison, send the request to ' + APPROVER_DEFAULT + ' for approval.</b>', small),
      Spacer(1, 4)]
L = LIAISONS
half = (len(L) + 1) // 2
data = [['Team', 'AI Liaison', 'Team', 'AI Liaison']]
for i in range(half):
    a = L[i]; b = L[i + half] if i + half < len(L) else ('', '')
    data.append([a[0], a[1], b[0], b[1]])
s.append(table(data, [2.15 * inch, 1.35 * inch, 2.15 * inch, 1.35 * inch]))

def footer(c, d):
    c.saveState(); c.setFont('Helvetica', 7.5); c.setFillColor(colors.HexColor('#808080'))
    c.drawString(0.75 * inch, 0.5 * inch, fill(FOOTER))
    c.drawRightString(7.75 * inch, 0.5 * inch, f'Page {d.page}'); c.restoreState()

doc = SimpleDocTemplate(OUT, pagesize=letter, leftMargin=0.75 * inch, rightMargin=0.75 * inch, topMargin=0.7 * inch, bottomMargin=0.8 * inch,
                        title=fill('{APP} Access Requests: Quick Reference'), author='[Department]')
doc.build(s, onFirstPage=footer, onLaterPages=footer)
print('ok')
