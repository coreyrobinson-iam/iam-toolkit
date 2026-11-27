"""
Access Request Tracker builder (generic, scrubbed).
Builds an Excel workbook that tracks license access requests for a SaaS app
gated by an Entra ID / AD security group, with a removal log and a console
for checking new removal lists.

Usage:  python build_tracker.py "Access Request Tracker.xlsx"
Then recalculate formulas (open in Excel, or run LibreOffice headless recalc).
Edit the CONFIG block for your organization.
"""
import sys, datetime as dt
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule
from openpyxl.comments import Comment
from openpyxl.utils import get_column_letter

# ================= CONFIG =================
OUT = sys.argv[1] if len(sys.argv) > 1 else 'Access Request Tracker.xlsx'
N = 300                                   # tracker rows
APP = 'ExampleApp'                        # app name used in labels and templates
APPROVER_CONTRACT = '[Contract attorney approver]'
APPROVER_VENDOR = '[Vendor/affiliate approver]'
VENDOR_ROLE = 'Vendor/affiliate'          # e.g. an outsourced staffing group
APPROVER_DEFAULT = '[Default approver]'   # used when a team has no liaison
TRAINING_LOCATION = '[LMS name]'
TRAIN_URL_1 = '[Training recording 1 URL]'
TRAIN_URL_2 = '[Training recording 2 URL]'
TRAIN_NAME_1 = 'Litigation Training'
TRAIN_NAME_2 = 'Transactional Training'
APP_URL = 'https://app.example.com/'
GUIDE_LINE = 'Please see our [guide name] for more information, including upcoming and previously recorded training sessions: [guide link].'
INACTIVITY_RULE = 'access > 20 days and no activity in last 20 days (internal only)'
# ==========================================

F = 'Arial'
hdr_font = Font(name=F, bold=True, color='FFFFFF', size=10)
hdr_fill = PatternFill('solid', fgColor='1F3864')
calc_hdr_fill = PatternFill('solid', fgColor='595959')
inp_font = Font(name=F, color='0000FF', size=10)
calc_font = Font(name=F, color='000000', size=10)
bold = Font(name=F, bold=True, size=10)
body = Font(name=F, size=10)
title_font = Font(name=F, bold=True, size=14)
inp_fill = PatternFill('solid', fgColor='FFF9E5')
thin = Side(style='thin', color='D9D9D9')
box = Border(left=thin, right=thin, top=thin, bottom=thin)
wrap = Alignment(wrap_text=True, vertical='top')

wb = Workbook()

# ---------------- Tracker ----------------
ws = wb.active
ws.title = 'Tracker'
cols = [
    # header, width, kind(input/calc)
    ('Ticket #', 11, 'in'),
    ('Received', 11, 'in'),
    ('Requester', 22, 'in'),
    ('Email', 30, 'in'),
    ('Role', 17, 'in'),
    ('Team (staff only)', 26, 'in'),
    ('Request type', 12, 'in'),
    ('Last removed', 12, 'calc'),
    ('Training certificate received (new only)', 14, 'in'),
    ('Expects regular use (restore only)', 14, 'in'),
    ('Approver', 22, 'calc'),
    ('Approval', 12, 'in'),
    ('Sent to User Admin', 12, 'in'),
    ('Access confirmed', 12, 'in'),
    ('Next action', 48, 'calc'),
    ('Status', 14, 'calc'),
    ('Notes', 48, 'in'),
    ('Times removed', 10, 'calc'),
]
for i, (h, w, k) in enumerate(cols, 1):
    c = ws.cell(row=1, column=i, value=h)
    c.font = hdr_font
    c.fill = hdr_fill if k == 'in' else calc_hdr_fill
    c.alignment = Alignment(wrap_text=True, vertical='center', horizontal='center')
    c.border = box
    ws.column_dimensions[get_column_letter(i)].width = w
ws.row_dimensions[1].height = 44
ws.freeze_panes = 'D2'

def next_action(r):
    E, G, I, J, K, L, M, N_ = (f'{x}{r}' for x in 'EGIJKLMN')
    approve_path = (f'IF(L{r}="Declined","Tell requester it was declined; close",'
                    f'IF(OR(E{r}="Attorney",L{r}="Approved"),'
                    f'IF(M{r}="","{{UA}}","{{AFTER}}"),'
                    f'IF(L{r}="Pending","Waiting on approver ("&K{r}&")",'
                    f'"{{ASK}}")))')
    new_path = approve_path.replace('{UA}', 'Ask User Admin to add to {APP} group') \
        .replace('{AFTER}', 'Once added: send Template C, then enter Access confirmed date') \
        .replace('"{ASK}"', f'"Send Template A to requester + Template B to "&K{r}&"; set Approval to Pending"')
    res_path = approve_path.replace('{UA}', 'Ask User Admin to restore to {APP} group') \
        .replace('{AFTER}', 'Once restored: send Template E, then enter Access confirmed date') \
        .replace('"{ASK}"', f'"Send Template B to "&K{r}&" for approval; set Approval to Pending"')
    role_chk = f'IF(AND(E{r}="",L{r}<>"Approved"),"Confirm requester\'s role (attorney or not)",'
    return (f'=IF(C{r}="","",IF(N{r}<>"","Done",IF(G{r}="","Set Request type (New or Restore)",'
            f'IF(G{r}="New",'
            f'IF(I{r}<>"Yes","Send Template F (training links); mark Training certificate received when it arrives",'
            f'{role_chk}{new_path})),'
            f'IF(J{r}="","Send Template D; record their answer in Expects regular use",'
            f'IF(J{r}="No","No current need: close; they can reach out later",'
            f'{role_chk}{res_path})))))))')

def status(r):
    return (f'=IF(C{r}="","",IF(N{r}<>"","Done",'
            f'IF(OR(J{r}="No",L{r}="Declined"),"Closed - no access",'
            f'IF(OR(L{r}="Pending",AND(M{r}<>"",N{r}="")),"Waiting","Action needed"))))')

for r in range(2, N + 2):
    ws[f'H{r}'] = f'=IF(D{r}="","",IF(COUNTIF(\'Removal Log\'!$A:$A,D{r})=0,"Never",_xlfn.MAXIFS(\'Removal Log\'!$B:$B,\'Removal Log\'!$A:$A,D{r})))'
    ws[f'R{r}'] = f'=IF(D{r}="","",COUNTIF(\'Removal Log\'!$A:$A,D{r}))'
    ws[f'K{r}'] = (f'=IF(E{r}="","",IF(E{r}="Attorney","None needed",'
                   f'IF(E{r}="Contract attorney","{APPROVER_CONTRACT}",'
                   f'IF(E{r}="{VENDOR_ROLE}","{APPROVER_VENDOR}",'
                   f'IFERROR(INDEX(Liaisons!$C:$C,MATCH(F{r},Liaisons!$A:$A,0)),"{APPROVER_DEFAULT} (no liaison)")))))')
    ws[f'O{r}'] = next_action(r)
    ws[f'P{r}'] = status(r)
    for ci, (_, _, k) in enumerate(cols, 1):
        c = ws.cell(row=r, column=ci)
        c.border = box
        c.alignment = Alignment(vertical='top', wrap_text=(ci in (15, 17)))
        if k == 'in':
            c.font = inp_font
            c.fill = inp_fill
        else:
            c.font = calc_font
    for col in 'BHMN':
        ws[f'{col}{r}'].number_format = 'mm/dd/yyyy'

# one fictional example row showing the expected format
ws['A2'] = '100001'; ws['B2'] = dt.date(2026, 1, 5); ws['C2'] = 'Jane Example (sample row - delete)'
ws['D2'] = 'jane.example@example.com'; ws['E2'] = 'Attorney'; ws['G2'] = 'New'
ws['Q2'] = 'Sample row. Delete before use.'

# dropdowns
def dv(formula, rng, allow_other=False):
    v = DataValidation(type='list', formula1=formula, allow_blank=True)
    if allow_other:
        v.showErrorMessage = False
    ws.add_data_validation(v)
    v.add(rng)
dv(f'"Attorney,Contract attorney,Staff,{VENDOR_ROLE}"', f'E2:E{N+1}')
dv(f"=Liaisons!$A$2:$A$60", f'F2:F{N+1}', allow_other=True)
dv('"New,Restore"', f'G2:G{N+1}')
dv('"Yes,No"', f'I2:I{N+1}')
dv('"Yes,No"', f'J2:J{N+1}')
dv('"Pending,Approved,Declined"', f'L2:L{N+1}')

# status colours
rng = f'P2:P{N+1}'
ws.conditional_formatting.add(rng, FormulaRule(formula=['$P2="Action needed"'], fill=PatternFill('solid', fgColor='FBE3D6'), font=Font(name=F, bold=True, color='9C3B00')))
ws.conditional_formatting.add(rng, FormulaRule(formula=['$P2="Waiting"'], fill=PatternFill('solid', fgColor='FFF2CC')))
ws.conditional_formatting.add(rng, FormulaRule(formula=['$P2="Done"'], fill=PatternFill('solid', fgColor='E2EFDA')))
ws.conditional_formatting.add(rng, FormulaRule(formula=['$P2="Closed - no access"'], fill=PatternFill('solid', fgColor='EDEDED')))
ws.auto_filter.ref = f'A1:R{N+1}'

ws['H1'].comment = Comment('Most recent date this email appears in the Removal Log tab, or Never.', 'Admin')
ws['R1'].comment = Comment('How many times this email appears in the Removal Log tab.', 'Admin')
ws['K1'].comment = Comment('From Role and Team. Attorneys need no approval. Staff: team liaison from the Liaisons tab, or the default approver if the team has none. Contract attorney and vendor/affiliate: their configured approvers. Same approvers for new and restore.', 'Admin')
ws['I1'].comment = Comment('New users must send a certificate of completion for one of the required training recordings. Not required for restores.', 'Admin')
ws['J1'].comment = Comment("Their answer to Template D. Leave blank until they reply.", 'Admin')

# ---------------- How to use ----------------
hw = wb.create_sheet('How to use', 0)
hw.column_dimensions['A'].width = 4
hw.column_dimensions['B'].width = 100
lines = [
    ('{APP} Access Request Tracker', title_font),
    ('One row per request. Fill in the cream-colored cells (blue text); grey-header columns calculate themselves.', body),
    ('', body),
    ('Steps', bold),
    ('1. Enter Ticket #, Received, Requester, Email, Role and Request type. Staff requests: also pick the Team.', body),
    ('2. Read Next action and do it. Approver, Last removed and Times removed fill in on their own.', body),
    ('3. Record what happened: Training certificate received, Expects regular use, Approval, Sent to User Admin, Access confirmed.', body),
    ('4. Next action updates after each entry. Entering an Access confirmed date marks the row Done.', body),
    ('', body),
    ('How Next action decides', bold),
    ('New request: training first (Template F) > approval if not an attorney (Templates A + B) > User Admin adds > Template C.', body),
    ('Restore request: Template D asks about regular use > approval if not an attorney > User Admin restores > Template E. No training needed.', body),
    ('A row with Approval = Approved skips the role check (use for pre-approved cases, e.g. a user removed in error).', body),
    ('', body),
    ('Status colours', bold),
    ('Action needed = your move · Waiting = approver or User Admin · Done · Closed - no access.', body),
    ('', body),
    ('Tabs', bold),
    ('Tracker: the requests (row 2 is a sample; delete it).', body),
    ('Templates: email text A-F. Fill in the bracketed placeholders.', body),
    ('Liaisons: team approvers. Teams not listed go to the default approver.', body),
    ('Console: paste a new removal list to check it before logging it (steps are on that tab).', body),
    ('Removal Log: every removal, one row per person per removal. Add each new list at the bottom; the Tracker and Console update on their own.', body),
    ('', body),
    ('Generic template. Configure approvers, URLs and wording in the CONFIG block of build_tracker.py.', Font(name=F, size=9, italic=True, color='595959')),
]
for i, (t, f) in enumerate(lines, 1):
    c = hw.cell(row=i, column=2, value=t)
    c.font = f
    c.alignment = Alignment(wrap_text=True, vertical='top')

# ---------------- Templates ----------------
tw = wb.create_sheet('Templates')
for col, w in zip('ABCD', (10, 26, 18, 100)):
    tw.column_dimensions[col].width = w
for i, h in enumerate(('Template', 'Use', 'Status', 'Text'), 1):
    c = tw.cell(row=1, column=i, value=h)
    c.font = hdr_font; c.fill = hdr_fill; c.border = box
LIT = TRAIN_URL_1
TRN = TRAIN_URL_2
LIB = GUIDE_LINE
temps = [
    ('A', 'To requester: approval needed', 'In use',
     'Thanks for your interest in {APP}. Since {APP} licenses are limited and carry a significant cost, we need to seek approval for this request. I will get back to you once I hear back.'),
    ('B', 'To approver', 'In use', '[Name] has requested {APP} access. Do you approve?'),
    ('C', 'Access created (new)', 'In use',
     'I have created {APP} access for you. Wait about 30 minutes and try logging in at ' + APP_URL + ' with your email address. It may take up to an hour for you to be able to access.\n' + LIB + '\nPlease note that accounts are automatically deprovisioned after set periods of inactivity.'),
    ('D', 'Access was removed (restore)', 'In use',
     "To make the best use of a limited number of {APP} licenses, the firm periodically reassigns seats that haven't been used recently. Our records showed no recent activity on your account, so your access was deactivated.\nDo you expect to use {APP} regularly for current or upcoming work? If so, please let me know and we'll look at restoring your access. If not, you're welcome to reach out whenever a need comes up."),
    ('E', 'Access restored', 'In use',
     'I have restored {APP} access for you. You should now be able to log in at {APP_URL} with your email address.\n' + LIB),
    ('F', 'Training required (new)', 'In use',
     f"Thanks for your interest in {APP}. Before we can set up your access, please complete one of the two required {APP} training recordings in {TRAINING_LOCATION}:\n- {TRAIN_NAME_1}: {LIT}\n- {TRAIN_NAME_2}: {TRN}\nOnce you've finished, reply to this email with your certificate of completion and we can get you set up."),
]
for i, t in enumerate(temps, 2):
    for j, v in enumerate(t, 1):
        c = tw.cell(row=i, column=j, value=v)
        c.font = bold if j == 1 else body
        c.alignment = wrap; c.border = box
tw.freeze_panes = 'A2'

# ---------------- Liaisons ----------------
lw = wb.create_sheet('Liaisons')
for col, w in zip('ABC', (58, 10, 32)):
    lw.column_dimensions[col].width = w
for i, h in enumerate(('Team', 'Code', 'AI liaison'), 1):
    c = lw.cell(row=1, column=i, value=h)
    c.font = hdr_font; c.fill = hdr_fill; c.border = box
liaisons = [
    ('Example Team', 'EXT', '[Liaison name]'),
]
for i, t in enumerate(liaisons, 2):
    for j, v in enumerate(t, 1):
        c = lw.cell(row=i, column=j, value=v); c.font = body; c.border = box
nr = len(liaisons) + 3
lw.cell(row=nr, column=1, value='Add one row per team. Any team not listed goes to the default approver.').font = Font(name=F, size=9, italic=True, color='595959')

# ---------------- Removal Log ----------------
from openpyxl.worksheet.table import Table, TableStyleInfo
rw = wb.create_sheet('Removal Log')
for col, w in zip('ABCD', (40, 14, 40, 40)):
    rw.column_dimensions[col].width = w
for i, h_ in enumerate(('Email', 'Removal date', 'Source', 'Notes'), 1):
    rw.cell(row=1, column=i, value=h_)
REMOVED = ['sample.user@example.com']
SRC = '[Who requested the removal]'
for i, e in enumerate(REMOVED, 2):
    rw.cell(row=i, column=1, value=e).font = inp_font
    c = rw.cell(row=i, column=2, value=dt.date(2026, 1, 1)); c.font = inp_font; c.number_format = 'mm/dd/yyyy'
    rw.cell(row=i, column=3, value=SRC).font = inp_font
    rw.cell(row=i, column=4, value='Sample row. Delete before use.').font = inp_font
last = len(REMOVED) + 1
tbl = Table(displayName='RemovalLog', ref=f'A1:D{last}')
tbl.tableStyleInfo = TableStyleInfo(name='TableStyleMedium2', showRowStripes=True)
rw.add_table(tbl)
rw.freeze_panes = 'A2'
for r in range(2, last + 400):
    rw[f'B{r}'].number_format = 'mm/dd/yyyy'
rw['F1'] = 'How to add a list'
rw['F1'].font = bold
for k, t in enumerate([
    'Paste the cleaned emails (Console column B, pasted as values) into the first empty row under the table.',
    'Fill in Removal date and Source for the new rows. The table grows on its own.',
    'One row per person per removal. Never delete old rows: they are the history.',
], 2):
    c = rw.cell(row=k, column=6, value=f'{k-1}. {t}'); c.font = body
rw.column_dimensions['F'].width = 90

# ---------------- Console ----------------
cw = wb.create_sheet('Console', 1)
widths = (44, 38, 10, 13, 11, 16, 40)
for col, w in zip('ABCDEFG', widths):
    cw.column_dimensions[col].width = w
cw['A1'] = 'Removal list console'; cw['A1'].font = title_font
steps_ = [
    "1. Paste the new removal list into column A, one email per row. Text like 'name@example.com <mailto:...>' is fine.",
    "2. Check the Flag column. Raise any 'Open request' or 'Recently provisioned' rows with the removal owner before the removal goes through.",
    "3. Copy column B and paste as values into the first empty row of the Removal Log tab. Fill in Removal date and Source.",
    "4. Clear column A. The Tracker's Last removed and Times removed columns are now up to date.",
]
for k, t in enumerate(steps_, 2):
    c = cw.cell(row=k, column=1, value=t); c.font = body
for k, (lab, f_) in enumerate([
    ('Emails pasted', '=COUNTIF(B9:B1008,"?*")'),
    ('Need a look (open request or recently provisioned)', '=COUNTIF(G9:G1008,"Open request*")+COUNTIF(G9:G1008,"Recently provisioned*")'),
], 0):
    cw.cell(row=6, column=1 + k * 3, value=lab).font = bold
    c = cw.cell(row=6, column=2 + k * 3 + (1 if k else 0), value=f_); c.font = calc_font
heads = ['Paste emails here', 'Cleaned email', 'Times in log', 'Last removed', 'In tracker', 'Tracker status', 'Flag']
for i, h_ in enumerate(heads, 1):
    c = cw.cell(row=8, column=i, value=h_)
    c.font = hdr_font; c.fill = hdr_fill if i == 1 else calc_hdr_fill
    c.alignment = Alignment(wrap_text=True, vertical='center'); c.border = box
cw.freeze_panes = 'A9'
for r in range(9, 1009):
    a = f'A{r}'; b = f'B{r}'
    cw[a].fill = inp_fill; cw[a].font = inp_font
    cw[b] = f'=IF(TRIM({a})="","",LOWER(LEFT(TRIM({a}),FIND(" ",TRIM({a})&" ")-1)))'
    cw[f'C{r}'] = f'=IF({b}="","",COUNTIF(\'Removal Log\'!$A:$A,{b}))'
    cw[f'D{r}'] = f'=IF({b}="","",IF(C{r}=0,"Never",_xlfn.MAXIFS(\'Removal Log\'!$B:$B,\'Removal Log\'!$A:$A,{b})))'
    cw[f'D{r}'].number_format = 'mm/dd/yyyy'
    cw[f'E{r}'] = f'=IF({b}="","",IF(COUNTIF(Tracker!$D:$D,{b})>0,"Yes","No"))'
    cw[f'F{r}'] = f'=IF({b}="","",IF(E{r}<>"Yes","",INDEX(Tracker!$P:$P,MATCH({b},Tracker!$D:$D,0))))'
    cw[f'G{r}'] = (f'=IF({b}="","",IF(COUNTIF($B$9:{b},{b})>1,"Duplicate in this paste",'
                   f'IF(F{r}="Done","Recently provisioned: confirm before removing",'
                   f'IF(AND(E{r}="Yes",F{r}<>"Closed - no access"),"Open request: check before removing",'
                   f'IF(C{r}>0,"Removed before ("&C{r}&"x)","First removal")))))')
    for col in 'BCDEFG':
        cw[f'{col}{r}'].font = calc_font
rngc = 'G9:G1008'
cw.conditional_formatting.add(rngc, FormulaRule(formula=['LEFT($G9,4)="Open"'], fill=PatternFill('solid', fgColor='FBE3D6'), font=Font(name=F, bold=True, color='9C3B00')))
cw.conditional_formatting.add(rngc, FormulaRule(formula=['LEFT($G9,6)="Recent"'], fill=PatternFill('solid', fgColor='FBE3D6'), font=Font(name=F, bold=True, color='9C3B00')))
cw.conditional_formatting.add(rngc, FormulaRule(formula=['LEFT($G9,9)="Duplicate"'], fill=PatternFill('solid', fgColor='EDEDED')))
cw.conditional_formatting.add(rngc, FormulaRule(formula=['LEFT($G9,7)="Removed"'], fill=PatternFill('solid', fgColor='FFF2CC')))

wb.active = wb.sheetnames.index('Tracker')
# Fill {APP} / {APP_URL} placeholders everywhere text was written.
def _fill(v):
    return v.replace('{APP_URL}', APP_URL).replace('{APP}', APP) if isinstance(v, str) else v
for _ws in wb.worksheets:
    for _row in _ws.iter_rows():
        for _c in _row:
            if isinstance(_c.value, str) and '{APP' in _c.value:
                _c.value = _fill(_c.value)
            if _c.comment and '{APP' in _c.comment.text:
                _c.comment = Comment(_fill(_c.comment.text), _c.comment.author)
    if '{APP' in _ws.title:
        _ws.title = _fill(_ws.title)

wb.save(OUT)
print('saved', OUT)
