"""
Access Requests Quick Reference generator.
Builds a 2-page PDF quick reference for staff who handle app access requests.
Usage: python build_qr.py "Quick Reference.pdf"   (requires reportlab)
Roles, approvers, rules and templates come from access_config.py.
"""
import sys
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, KeepTogether
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.lib.units import inch
import access_config as C

OUT = sys.argv[1] if len(sys.argv) > 1 else 'Quick Reference.pdf'
ORG_LINE = '{DEPT} | {ORG} | For staff handling {APP} access | Revised [date]'
FOOTER = '{APP} Access Requests: Quick Reference | Revised [date]'
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

def fill(t): return C.fill(t, newline='<br/>')
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

s = []
s += [P('{APP} Access Requests: Quick Reference', title),
      P(ORG_LINE, sub),
      Spacer(1, 6)]

s += [P('1. Intake', h)] + bullets([
    '{SERVICE_DESK} forwards all {APP} access questions to {INTAKE_MAILBOX}.',
    'First question for every request: <b>Is this new access, or a request to restore access?</b>',
    ('<b>New access requires training:</b> a certificate of completion for one of the required trainings in {TRAINING_LOCATION}. Restores do not.'
     if C.TRAINING_REQUIRED_FOR_NEW else 'No training is required before access is granted.'),
    'Seats are granted only through membership in <b>{ACCESS_GROUP}</b>. {IDENTITY_ADMIN} makes every change.',
])

s += [P('2. Who approves?', h), table(
    [['Role', 'Approver (new and restore)', 'Notes']] + [list(r) for r in C.APPROVAL_RULES],
    [1.6 * inch, 2.2 * inch, 3.2 * inch])]

s += [P('3. New access: steps', h)] + steps(C.NEW_ACCESS_STEPS)
s += [P('4. Restore access: steps', h)] + steps(C.RESTORE_STEPS)

s += [P('5. Access review (internal)', h)] + bullets([
    '<b>{REVIEW_FREQUENCY}</b>, run by {REVIEW_OWNER}: {INACTIVITY_RULE}',
    'Removal requests go to {IDENTITY_ADMIN}, with {REQUEST_HANDLER} copied so the Removal Log stays current.',
    'Messages to users do not state the inactivity period. New users are told in Template C that inactive accounts are removed.',
])

s += [P('6. Email templates', h)]
for k, use, body in C.TEMPLATES:
    s.append(KeepTogether([P(f'Template {k}: {use}', tmpl_h), P(body, quote)]))

s += [P('7. Team approvers', h),
      P('For <b>{ROLE_EMPLOYEE}</b> requests. <b>Any team not listed goes to {APPROVER_DEFAULT}.</b>', small),
      Spacer(1, 4)]
L = [(f'{n} ({c})', a) for n, c, a in C.TEAM_APPROVERS]
half = (len(L) + 1) // 2
data = [['Team', 'Approver', 'Team', 'Approver']]
for i in range(half):
    a_ = L[i]; b_ = L[i + half] if i + half < len(L) else ('', '')
    data.append([a_[0], a_[1], b_[0], b_[1]])
s.append(table(data, [2.15 * inch, 1.35 * inch, 2.15 * inch, 1.35 * inch]))

def footer(c, d):
    c.saveState(); c.setFont('Helvetica', 7.5); c.setFillColor(colors.HexColor('#808080'))
    c.drawString(0.75 * inch, 0.5 * inch, fill(FOOTER))
    c.drawRightString(7.75 * inch, 0.5 * inch, f'Page {d.page}'); c.restoreState()

doc = SimpleDocTemplate(OUT, pagesize=letter, leftMargin=0.75 * inch, rightMargin=0.75 * inch, topMargin=0.7 * inch, bottomMargin=0.8 * inch,
                        title=fill('{APP} Access Requests: Quick Reference'), author=fill('{DEPT}'))
doc.build(s, onFirstPage=footer, onLaterPages=footer)
print('ok')
