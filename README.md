# IAM Toolkit

Access request tracker and quick reference for group-based SaaS licensing: role-based approvals, training gates, periodic access reviews and an audit trail.

## Use case

An organization pays for a limited number of seats in a SaaS app, and access is granted through an Entra ID or Active Directory security group. Requests come in from different kinds of users, such as employees, contractors and vendors, and each needs a different approver. Some users must finish training before they get access. Seats also have to be reclaimed from inactive users on a regular schedule.

Without a system, that work lives in email threads and memory. Approvals get missed, people are granted access without the right sign-off, and when a removed user asks to come back, nobody can say why they were removed or how often.

The toolkit gives the access coordinator one workbook that tracks every request from intake to close. For each request it names the right approver, shows the next step and keeps a removal history. A matching one-page reference lets anyone covering the queue follow the same rules. Roles, approvers and email templates are set in one config file, so the tools fit any organization's policy without changing the code.

## What's here

| File | What it does |
|---|---|
| `scripts/access_config.py` | One place for the roles, approvers, approval rules and email templates |
| `scripts/build_tracker.py` | Builds an Excel workbook for tracking access requests from intake to close |
| `scripts/build_qr.py` | Builds a 2-page PDF quick reference for the staff who handle those requests |

Sample output from both is in `examples/`.

### Access Request Tracker (`build_tracker.py`)

One row per request. Staff fill in a few input cells, and the workbook works out the rest:

- **Next action.** Formulas read the request's state (role, request type, training, approval) and tell the handler exactly what to do next.
- **Approver routing.** A pre-approved role needs no approval. Employees route to their team approver from the `Approvers` tab, with a default approver as a fallback. Contractors and vendors go to their own sponsors.
- **Access Rules tab.** The approval matrix and the step-by-step new and restore procedures, inside the workbook.
- **Removal history.** A `Removal Log` tab records deprovisioning runs. Each request shows when that user was last removed and how many times, so repeat inactive users are easy to spot.
- **Templates.** Ready-to-send emails for each step: approval requests, training requirements, access granted, restore check-ins.
- **Console.** A tab for checking a new removal list against the log before acting on it.

### Quick Reference (`build_qr.py`)

A printable summary of the same process: intake rules, who approves what, step-by-step new and restore flows, the deprovisioning cycle, and the email templates.

## Default roles and approvals

| Role | Approver (new and restore) |
|---|---|
| Pre-approved role | None. Approved by role |
| Employee | Team approver, or the default approver if the team isn't listed |
| Contractor | Contractor sponsor |
| Vendor/partner | Vendor relationship owner |

**New access:** training certificate > approval (unless pre-approved) > Identity Admin adds the user to the access group > confirmation email.

**Restore:** confirm ongoing need > approval (unless pre-approved) > Identity Admin restores > confirmation email. No training needed.

## Controls this models

- **Least privilege:** access is granted only through group membership, only after the right approval.
- **Approval routing by role:** different approval paths for pre-approved roles, employees, contractors and vendors.
- **Training gate:** new users must show proof of training before access is granted.
- **Periodic access review:** inactive seats are deprovisioned on a regular cycle, and restores require the user to confirm ongoing need.
- **Audit trail:** every request, approval and removal is recorded with dates.

## Usage

```bash
pip install -r requirements.txt
cd scripts
python build_tracker.py "Access Request Tracker.xlsx"
python build_qr.py "Quick Reference.pdf"
```

Open the workbook in Excel so its formulas calculate (or recalculate it with LibreOffice headless).

## Configure for your organization

Edit `scripts/access_config.py`. It holds the app name and URL, the access group, the sync delay wording, role names, approvers, team approvers, training links, review schedule, approval steps and email templates. Both scripts read from it, so the workbook and the quick reference always match. Set `TRAINING_REQUIRED_FOR_NEW = False` to drop the training gate.

## Notes

- All names, emails and teams in the scripts are placeholders. The sample row is labeled for deletion.
- No credentials, tenant IDs or real user data are stored or needed. The tools build blank templates. They don't connect to Entra ID.

## License

MIT. See `LICENSE`.
