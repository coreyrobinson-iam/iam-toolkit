# IAM Toolkit

Tools for running license access requests for a SaaS app whose seats are granted through an Entra ID / Active Directory security group. Built from real day-to-day provisioning work and scrubbed of anything organization-specific.

## What's here

| Script | What it builds |
|---|---|
| `scripts/build_tracker.py` | An Excel workbook for tracking access requests from intake to close |
| `scripts/build_qr.py` | A 2-page PDF quick reference for the staff who handle those requests |

Sample output from both is in `examples/`.

### Access Request Tracker (`build_tracker.py`)

One row per request. Staff fill in a few input cells, and the workbook works out the rest:

- **Next action.** Formulas read the request's state (role, request type, training, approval) and tell the handler exactly what to do next.
- **Approver routing.** Attorneys are auto-approved. Staff route to their team's liaison from the `Liaisons` tab, with a default approver as a fallback. Contract attorneys and vendors go to their own configured approvers.
- **Removal history.** A `Removal Log` tab records deprovisioning runs. Each request shows when that user was last removed and how many times, so repeat inactive users are easy to spot.
- **Templates.** Ready-to-send emails for each step: approval requests, training requirements, access granted, restore check-ins.
- **Console.** A tab for checking a new removal list against the log before acting on it.

### Quick Reference (`build_qr.py`)

A printable summary of the same process: intake rules, who approves what, step-by-step new and restore flows, the deprovisioning cycle, and the email templates.

## Controls this models

- **Least privilege:** access is granted only through group membership, only after the right approval.
- **Approval routing by role:** different approval paths for attorneys, staff, contract attorneys and vendors.
- **Training gate:** new users must show proof of training before access is granted.
- **Periodic access review:** inactive seats are deprovisioned on a regular cycle, and restores require the user to confirm ongoing need.
- **Audit trail:** every request, approval and removal is recorded with dates.

## Usage

```bash
pip install -r requirements.txt
python scripts/build_tracker.py "Access Request Tracker.xlsx"
python scripts/build_qr.py "Quick Reference.pdf"
```

Open the workbook in Excel so its formulas calculate (or recalculate it with LibreOffice headless).

## Configure for your organization

Each script starts with a `CONFIG` block. Set the app name (`APP`), its login URL (`APP_URL`), approvers, team liaisons, training links and mailbox names there. Everything else picks those values up.

## Notes

- All names, emails and teams in the scripts are placeholders. The sample row is labeled for deletion.
- No credentials, tenant IDs or real user data are stored or needed. The tools build blank templates. They don't connect to Entra ID.

## License

MIT. See `LICENSE`.
