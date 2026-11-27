"""
Access Request Tracker builder.
Builds an Excel workbook that tracks license access requests for a SaaS app
gated by an Entra ID / AD security group, with a removal log and a console
for checking new removal lists.

Usage:  python build_tracker.py "Access Request Tracker.xlsx"
Then recalculate formulas (open in Excel, or run LibreOffice headless recalc).
Roles, approvers, rules and templates come from access_config.py.
"""
import sys, re, datetime as dt
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule
from openpyxl.comments import Comment
from openpyxl.utils import get_column_letter
import access_config as C

OUT = sys.argv[1] if len(sys.argv) > 1 else 'Access Request Tracker.xlsx'
N = 300                                   # tracker rows
PRE, EMP, CON, VEN = C.ROLE_PREAPPROVED, C.ROLE_EMPLOYEE, C.ROLE_CONTRACTOR, C.ROLE_VENDOR

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
    (f'Team ({EMP} only)', 26, 'in'),
    ('Request type', 12, 'in'),
    ('Last removed', 12, 'calc'),
    ('Training certificate received (new only)', 14, 'in'),
    ('Expects regular use (restore only)', 14, 'in'),
    ('Approver', 22, 'calc'),
    ('Approval', 12, 'in'),
    (f'Sent to {C.IDENTITY_ADMIN}', 12, 'in'),
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
                    f'IF(OR(E{r}="{PRE}",L{r}="Approved"),'
                    f'IF(M{r}="","{{UA}}","{{AFTER}}"),'
                    f'IF(L{r}="Pending","Waiting on approver ("&K{r}&")",'
                    f'"{{ASK}}")))')
    new_path = approve_path.replace('{UA}', f'Ask {C.IDENTITY_ADMIN} to add to {C.ACCESS_GROUP}') \
        .replace('{AFTER}', 'Once added: send Template C, then enter Access confirmed date') \
        .replace('"{ASK}"', f'"Send Template A to requester + Template B to "&K{r}&"; set Approval to Pending"')
    res_path = approve_path.replace('{UA}', f'Ask {C.IDENTITY_ADMIN} to restore to {C.ACCESS_GROUP}') \
        .replace('{AFTER}', 'Once restored: send Template E, then enter Access confirmed date') \
        .replace('"{ASK}"', f'"Send Template B to "&K{r}&" for approval; set Approval to Pending"')
    role_chk = f'IF(AND(E{r}="",L{r}<>"Approved"),"Set the requester\'s Role",'
    if C.TRAINING_REQUIRED_FOR_NEW:
        new_branch = (f'IF(I{r}<>"Yes","Send Template F (training links); mark Training certificate received when it arrives",'
                      f'{role_chk}{new_path})),')
    else:
        new_branch = f'{role_chk}{new_path}),'
    return (f'=IF(C{r}="","",IF(N{r}<>"","Done",IF(G{r}="","Set Request type (New or Restore)",'
            f'IF(G{r}="New",'
            + new_branch +
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
    ws[f'K{r}'] = (f'=IF(E{r}="","",IF(E{r}="{PRE}","None needed",'
                   f'IF(E{r}="{CON}","{C.APPROVER_CONTRACTOR}",'
                   f'IF(E{r}="{VEN}","{C.APPROVER_VENDOR}",'
                   f'IFERROR(INDEX(Approvers!$C:$C,MATCH(F{r},Approvers!$A:$A,0)),"{C.APPROVER_DEFAULT} (team not listed)")))))')
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
ws['D2'] = 'jane.example@example.com'; ws['E2'] = PRE; ws['G2'] = 'New'
ws['Q2'] = 'Sample row. Delete before use.'

# dropdowns
def dv(formula, rng, allow_other=False):
    v = DataValidation(type='list', formula1=formula, allow_blank=True)
    if allow_other:
        v.showErrorMessage = False
    ws.add_data_validation(v)
    v.add(rng)
dv('"' + ",".join(C.ROLES) + '"', f'E2:E{N+1}')
dv(f"=Approvers!$A$2:$A$60", f'F2:F{N+1}', allow_other=True)
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
ws['K1'].comment = Comment(f'From Role and Team. {PRE}: no approval. {EMP}: team approver from the Approvers tab, or the default approver if the team is not listed. {CON} and {VEN}: their configured approvers. Same approvers for new and restore. Full rules on the Access Rules tab.', 'Admin')
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
    (f'1. Enter Ticket #, Received, Requester, Email, Role and Request type. {EMP} requests: also pick the Team.', body),
    ('2. Read Next action and do it. Approver, Last removed and Times removed fill in on their own.', body),
    (f'3. Record what happened: Training certificate received, Expects regular use, Approval, Sent to {C.IDENTITY_ADMIN}, Access confirmed.', body),
    ('4. Next action updates after each entry. Entering an Access confirmed date marks the row Done.', body),
    ('', body),
    ('How Next action decides', bold),
    (f'New request: training first (Template F) > approval unless {PRE} (Templates A + B) > {C.IDENTITY_ADMIN} adds to {C.ACCESS_GROUP} > Template C.', body),
    (f'Restore request: Template D asks about regular use > approval unless {PRE} > {C.IDENTITY_ADMIN} restores > Template E. No training needed.', body),
    ('A row with Approval = Approved skips the role check (use for pre-approved cases, e.g. a user removed in error).', body),
    ('', body),
    ('Status colours', bold),
    (f'Action needed = your move · Waiting = approver or {C.IDENTITY_ADMIN} · Done · Closed - no access.', body),
    ('', body),
    ('Tabs', bold),
    ('Access Rules: who approves each role, and the steps for new and restore requests.', body),
    ('Tracker: the requests (row 2 is a sample; delete it).', body),
    ('Templates: email text A-F. Fill in the bracketed placeholders.', body),
    ('Approvers: team approvers for employees. Teams not listed go to the default approver.', body),
    ('Console: paste a new removal list to check it before logging it (steps are on that tab).', body),
    ('Removal Log: every removal, one row per person per removal. Add each new list at the bottom; the Tracker and Console update on their own.', body),
    ('', body),
    ('Generic template. Configure roles, approvers, URLs and wording in access_config.py.', Font(name=F, size=9, italic=True, color='595959')),
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
temps = [(k, use, 'In use', C.fill(txt)) for k, use, txt in C.TEMPLATES]
for i, t in enumerate(temps, 2):
    for j, v in enumerate(t, 1):
        c = tw.cell(row=i, column=j, value=v)
        c.font = bold if j == 1 else body
        c.alignment = wrap; c.border = box
tw.freeze_panes = 'A2'

# ---------------- Approvers ----------------
lw = wb.create_sheet('Approvers')
for col, w in zip('ABC', (40, 10, 32)):
    lw.column_dimensions[col].width = w
for i, h in enumerate(('Team', 'Code', 'Team approver'), 1):
    c = lw.cell(row=1, column=i, value=h)
    c.font = hdr_font; c.fill = hdr_fill; c.border = box
for i, t_ in enumerate(C.TEAM_APPROVERS, 2):
    for j, v in enumerate(t_, 1):
        c = lw.cell(row=i, column=j, value=v); c.font = body; c.border = box
nr = len(C.TEAM_APPROVERS) + 3
lw.cell(row=nr, column=1, value=f'Add one row per team. {EMP} requests from a team not listed go to {C.APPROVER_DEFAULT}.').font = Font(name=F, size=9, italic=True, color='595959')

# ---------------- Access Rules ----------------
strip = lambda s: re.sub(r'<[^>]+>', '', C.fill(s))
ar = wb.create_sheet('Access Rules', 1)
for col, w in zip('ABC', (24, 34, 60)):
    ar.column_dimensions[col].width = w
ar['A1'] = '{APP} Access Rules'; ar['A1'].font = title_font
ar['A2'] = f'Seats are granted only through membership in {C.ACCESS_GROUP}. {C.IDENTITY_ADMIN} makes every change.'
ar['A2'].font = body
row = 4
ar.cell(row=row, column=1, value='Who approves').font = bold; row += 1
for i, h in enumerate(('Role', 'Approver (new and restore)', 'Notes'), 1):
    c = ar.cell(row=row, column=i, value=h); c.font = hdr_font; c.fill = hdr_fill; c.border = box
row += 1
for rule in C.APPROVAL_RULES:
    for j, v in enumerate(rule, 1):
        c = ar.cell(row=row, column=j, value=strip(v)); c.font = body; c.border = box; c.alignment = wrap
    row += 1
for heading, items in (('New access: steps', C.NEW_ACCESS_STEPS), ('Restore access: steps', C.RESTORE_STEPS)):
    row += 1
    ar.cell(row=row, column=1, value=heading).font = bold; row += 1
    for n, s_ in enumerate(items, 1):
        c = ar.cell(row=row, column=1, value=f'{n}. {strip(s_)}'); c.font = body
        ar.merge_cells(start_row=row, start_column=1, end_row=row, end_column=3)
        c.alignment = wrap; ar.row_dimensions[row].height = 28; row += 1
row += 1
ar.cell(row=row, column=1, value='Access review').font = bold; row += 1
for s_ in (f'{C.REVIEW_FREQUENCY}, run by {C.REVIEW_OWNER}: {C.INACTIVITY_RULE}',
           'Removal messages do not state the inactivity period. New users are told in Template C that inactive accounts are removed.',
           'Log every removal on the Removal Log tab. Check each new list on the Console tab before it goes through.'):
    c = ar.cell(row=row, column=1, value=s_); c.font = body
    ar.merge_cells(start_row=row, start_column=1, end_row=row, end_column=3)
    c.alignment = wrap; ar.row_dimensions[row].height = 28; row += 1

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
# Fill {TOKENS} everywhere text was written.
for _ws in wb.worksheets:
    for _row in _ws.iter_rows():
        for _c in _row:
            if isinstance(_c.value, str) and '{' in _c.value:
                _c.value = C.fill(_c.value)
            if _c.comment and '{' in _c.comment.text:
                _c.comment = Comment(C.fill(_c.comment.text), _c.comment.author)

wb.save(OUT)
print('saved', OUT)
