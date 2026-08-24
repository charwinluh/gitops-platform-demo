#!/usr/bin/env python3
"""Build the customer-facing certificate renewal activity schedule workbook."""

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

FONT = "Arial"

# ---------- palette ----------
C_TITLE_FILL = "1F3864"
C_HDR_FILL = "2F5496"
C_PHASE_FILL = "D9E2F3"
C_PENDING_FILL = "FFF2CC"
C_CONTINGENCY_FILL = "FBE5D6"
C_DOWNTIME_FILL = "FCE4E4"
C_INPUT_FILL = "FFFF00"

thin = Side(style="thin", color="B4C6E7")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)

wb = Workbook()

# =====================================================================
# Sheet 1 - Activity Schedule
# =====================================================================
ws = wb.active
ws.title = "Activity Schedule"

HEADERS = [
    "ID",
    "Phase",
    "Certificate / Component",
    "Activity / Task",
    "Description & Scope",
    "Est. Duration",
    "Customer Impact",
    "Downtime Required",
    "Planned Start Date",
    "Planned End Date",
    "Dependency",
    "Customer Action Required",
    "Status",
]

WIDTHS = [7, 34, 26, 34, 52, 17, 40, 22, 20, 20, 26, 38, 22]

# (id, component, task, description, duration, impact, downtime,
#  start, end, dependency, customer_action, status)
PHASES = [
    (
        "PHASE 0  -  PREPARATION & ACCESS ENABLEMENT   (no access to customer environment required)",
        [
            ("0.1", "All", "Kick-off and scheduling call",
             "Walk through scope of work, agree maintenance windows, confirm escalation contacts and communication channels.",
             "1 hour", "None", "No", "2026-08-26", "2026-08-26", "-",
             "Attend call; nominate technical and business point of contact", "Planned"),
            ("0.2", "All", "Confirm remote access and credentials",
             "Validate VPN / jump host / bastion connectivity and confirm engineer accounts hold the required privilege level.",
             "2 business days", "None", "No", "2026-08-27", "2026-08-28", "0.1",
             "Provision or approve access; whitelist engineer source IPs", "Planned"),
            ("0.3", "All", "Confirm change freeze and blackout dates",
             "Confirm no business-critical events, financial close or freeze periods conflict with the proposed windows.",
             "1 hour", "None", "No", "2026-08-27", "2026-08-27", "0.1",
             "Provide blackout / freeze calendar", "Planned"),
            ("0.4", "All", "Verify current backup and restore point",
             "Confirm platform backups are recent and restorable. This underpins the rollback plan for every later phase.",
             "0.5 business day", "None", "No", "2026-08-28", "2026-08-28", "0.2",
             "Confirm backup policy and last successful restore point", "Planned"),
        ],
    ),
    (
        "PHASE 1  -  ENVIRONMENT VERIFICATION & INFORMATION COLLECTION   (ManageOne, FusionSphere - change request approved)",
        [
            ("1.1", "ManageOne", "Pre-check health assessment",
             "Runbook-guided health check of the ManageOne cluster to establish a known-good baseline prior to any collection.",
             "3 hours", "None - read-only activity", "No", "2026-09-01", "2026-09-01", "0.2, 0.4",
             "Provide agreed access window", "Planned"),
            ("1.2", "ManageOne", "Certificate inventory and expiry collection",
             "Collect full certificate inventory: issuer, chain, expiry dates, deployment mode and dependent components.",
             "3 hours", "None - read-only activity", "No", "2026-09-01", "2026-09-01", "1.1",
             "-", "Planned"),
            ("1.3", "FusionSphere", "Pre-check health assessment",
             "Health check of FusionSphere management nodes, hosts and cluster status to establish a known-good baseline.",
             "4 hours", "None - read-only activity", "No", "2026-09-02", "2026-09-02", "0.2, 0.4",
             "Provide agreed access window", "Planned"),
            ("1.4", "FusionSphere", "Certificate inventory and expiry collection",
             "Collect certificate inventory and component topology across the virtualisation layer.",
             "3 hours", "None - read-only activity", "No", "2026-09-02", "2026-09-02", "1.3",
             "-", "Planned"),
            ("1.5", "ManageOne, FusionSphere", "System time / NTP synchronisation check",
             "Confirm NTP synchronisation across all nodes. Certificate validation is time-sensitive; clock drift is a common cause of post-change failures.",
             "1 hour", "None - read-only activity", "No", "2026-09-03", "2026-09-03", "1.2, 1.4",
             "Confirm authoritative NTP source", "Planned"),
            ("1.6", "All", "Extended certificate expiry scan",
             "Identify any further certificates due to expire within the next 6 months so they can be bundled into the same maintenance window rather than requiring a repeat outage.",
             "2 hours", "None - read-only activity", "No", "2026-09-03", "2026-09-03", "1.2, 1.4",
             "-", "Planned"),
            ("1.7", "ManageOne, FusionSphere", "Information collection report and handover",
             "Package all collected findings and submit to backend engineering for analysis.",
             "1 business day", "None", "No", "2026-09-03", "2026-09-04", "1.5, 1.6",
             "-", "Planned"),
        ],
    ),
    (
        "PHASE 2  -  ANALYSIS & FINAL RUNBOOK PREPARATION   (performed offline - no access to customer environment required)",
        [
            ("2.1", "ManageOne, FusionSphere", "Backend analysis of collected data",
             "Backend engineering analyses the environment-specific findings and designs the certificate update procedure.",
             "5 business days", "None", "No", "2026-09-07", "2026-09-11", "1.7",
             "-", "Planned"),
            ("2.2", "ManageOne, FusionSphere", "Final update runbook issued",
             "Environment-specific update runbook produced and passed through internal technical review.",
             "2 business days", "None", "No", "2026-09-14", "2026-09-15", "2.1",
             "-", "Planned"),
            ("2.3", "ManageOne, FusionSphere", "Method of Procedure and rollback plan review",
             "Joint walkthrough of the implementation steps, expected impact, verification criteria and rollback plan.",
             "2 hours", "None", "No", "2026-09-16", "2026-09-16", "2.2",
             "Review and approve the Method of Procedure; confirm maintenance window", "Planned"),
            ("2.4", "ManageOne, FusionSphere", "Change request submission for implementation",
             "Submit the implementation change request into the customer change management process.",
             "2 business days", "None", "No", "2026-09-16", "2026-09-17", "2.3",
             "Approve change request via customer change board", "Planned"),
        ],
    ),
    (
        "PHASE 3  -  IMPLEMENTATION: MANAGEONE & FUSIONSPHERE   (scheduled in weekend maintenance window)",
        [
            ("3.1", "ManageOne, FusionSphere", "Pre-implementation health check and snapshot",
             "Full health check plus snapshot / backup of all affected components taken immediately before any change is made.",
             "2 hours", "None", "No", "2026-09-19", "2026-09-19", "2.4",
             "Provide go / no-go confirmation", "Planned"),
            ("3.2", "ManageOne", "Certificate replacement",
             "Replace expired ManageOne certificates in accordance with the final update runbook.",
             "4 hours", "Management portal and API access interrupted during service restart. Hosted workloads are not expected to be affected.",
             "Yes - service restart", "2026-09-19", "2026-09-19", "3.1",
             "Notify end users of the maintenance window", "Planned"),
            ("3.3", "ManageOne", "Post-change functional verification",
             "Verify portal login, API availability, inter-component authentication and clearance of related alarms.",
             "2 hours", "Minimal - verification activity only", "No", "2026-09-19", "2026-09-19", "3.2",
             "Perform customer acceptance checks and confirm portal access", "Planned"),
            ("3.4", "FusionSphere", "Certificate replacement",
             "Replace expired FusionSphere certificates in accordance with the final update runbook.",
             "5 hours", "Management plane temporarily unavailable. Running virtual machine workloads are expected to remain in service.",
             "Yes - management plane", "2026-09-20", "2026-09-20", "3.3",
             "Notify end users of the maintenance window", "Planned"),
            ("3.5", "FusionSphere", "Post-change functional verification",
             "Verify management node health, host connectivity, cluster status and virtual machine lifecycle operations.",
             "3 hours", "Minimal - verification activity only", "No", "2026-09-20", "2026-09-20", "3.4",
             "Perform customer acceptance checks", "Planned"),
            ("3.6", "ManageOne, FusionSphere", "Rollback (contingency only)",
             "Invoked only if post-change verification fails. Restores affected components from the pre-change snapshot taken in task 3.1.",
             "4 hours", "Extended maintenance window if invoked", "Yes - only if invoked",
             "2026-09-20", "2026-09-20", "Invoked on failure of 3.3 or 3.5",
             "Joint go / no-go decision", "Contingency"),
            ("3.7", "ManageOne, FusionSphere", "Post-implementation monitoring",
             "Observation period to confirm platform stability and detect any latent issues following the change.",
             "5 business days", "None", "No", "2026-09-21", "2026-09-25", "3.5",
             "Report any observed anomalies", "Planned"),
            ("3.8", "ManageOne, FusionSphere", "Closure report and handover",
             "Completion report issued including updated certificate inventory and next expiry dates for forward planning.",
             "1 business day", "None", "No", "2026-09-28", "2026-09-28", "3.7",
             "Acknowledge completion and closure", "Planned"),
        ],
    ),
    (
        "PHASE 4  -  CLOUDSCOPELITE & CLOUDMIDDLEWAREKAFKA   (currently progressing through internal change governance - dates indicative)",
        [
            ("4.1", "CloudScopeLite", "Internal change assessment and approval",
             "This component is currently progressing through our internal change governance process. Confirmed dates will be issued to you as soon as this completes.",
             "To be confirmed", "None", "No", "To be confirmed", "To be confirmed",
             "Completion of internal change governance",
             "No action required at this stage", "In progress - internal"),
            ("4.2", "CloudScopeLite", "Environment verification and information collection",
             "Read-only environment assessment and certificate inventory collection, mirroring the Phase 1 approach.",
             "Est. 1 business day", "None - read-only activity", "No",
             "Indicative: w/c 05-Oct-2026", "To be confirmed", "4.1",
             "Provide access window once dates are confirmed", "Pending"),
            ("4.3", "CloudMiddleWareKafka", "Internal change assessment and approval",
             "This component is currently progressing through our internal change governance process. Confirmed dates will be issued to you as soon as this completes.",
             "To be confirmed", "None", "No", "To be confirmed", "To be confirmed",
             "Completion of internal change governance",
             "No action required at this stage", "In progress - internal"),
            ("4.4", "CloudMiddleWareKafka", "Environment verification and information collection",
             "Read-only environment assessment and certificate inventory collection, mirroring the Phase 1 approach.",
             "Est. 1 business day", "None - read-only activity", "No",
             "Indicative: w/c 05-Oct-2026", "To be confirmed", "4.3",
             "Provide access window once dates are confirmed", "Pending"),
            ("4.5", "CloudScopeLite, CloudMiddleWareKafka", "Analysis and final runbook preparation",
             "Backend engineering analysis and production of the environment-specific update runbook.",
             "Est. 5 business days", "None", "No", "To be confirmed", "To be confirmed", "4.2, 4.4",
             "-", "Pending"),
            ("4.6", "CloudScopeLite", "Certificate replacement",
             "Replace expired certificates in accordance with the final update runbook, once issued.",
             "To be confirmed",
             "Monitoring, reporting and alerting functions expected to be interrupted for the duration of the change.",
             "Likely - to be confirmed", "Indicative: w/c 19-Oct-2026", "To be confirmed",
             "4.5 and customer change approval",
             "Approve maintenance window when issued", "Pending"),
            ("4.7", "CloudMiddleWareKafka", "Certificate replacement",
             "Replace expired certificates in accordance with the final update runbook. Expected to require a rolling restart across brokers.",
             "To be confirmed",
             "Possible interruption to message processing. Applications consuming this service may require reconnection.",
             "Yes - rolling restart", "Indicative: w/c 19-Oct-2026", "To be confirmed", "4.6",
             "Identify dependent applications and arrange application team standby", "Pending"),
        ],
    ),
]

# ---- title block ----
last_col = len(HEADERS)
ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=last_col)
t = ws.cell(row=1, column=1, value="Certificate Renewal - Customer Activity Schedule")
t.font = Font(name=FONT, size=16, bold=True, color="FFFFFF")
t.fill = PatternFill("solid", fgColor=C_TITLE_FILL)
t.alignment = Alignment(horizontal="left", vertical="center", indent=1)
ws.row_dimensions[1].height = 30

ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=last_col)
s = ws.cell(row=2, column=1,
            value="Prepared for: [CUSTOMER NAME]     |     Prepared by: [ENGINEER NAME]     |     "
                  "Document date: 24-Aug-2026     |     Version: 0.1 (Draft for customer review)")
s.font = Font(name=FONT, size=10, italic=True, color="1F3864")
s.fill = PatternFill("solid", fgColor="D9E2F3")
s.alignment = Alignment(horizontal="left", vertical="center", indent=1)
ws.row_dimensions[2].height = 20

ws.merge_cells(start_row=3, start_column=1, end_row=3, end_column=last_col)
n = ws.cell(row=3, column=1,
            value="All dates are proposed and subject to customer confirmation. Phases 0 to 3 relate to components "
                  "for which internal change approval has been obtained. Phase 4 dates are indicative only and will "
                  "be confirmed on completion of our internal change governance process.")
n.font = Font(name=FONT, size=9, italic=True)
n.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1)
ws.row_dimensions[3].height = 28

# ---- header row ----
HDR_ROW = 5
for i, h in enumerate(HEADERS, start=1):
    c = ws.cell(row=HDR_ROW, column=i, value=h)
    c.font = Font(name=FONT, size=10, bold=True, color="FFFFFF")
    c.fill = PatternFill("solid", fgColor=C_HDR_FILL)
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    c.border = BORDER
ws.row_dimensions[HDR_ROW].height = 34

for i, w in enumerate(WIDTHS, start=1):
    ws.column_dimensions[get_column_letter(i)].width = w

# ---- data rows ----
r = HDR_ROW + 1
first_data_row = r
for phase_label, tasks in PHASES:
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=last_col)
    pc = ws.cell(row=r, column=1, value=phase_label)
    pc.font = Font(name=FONT, size=10, bold=True, color="1F3864")
    pc.fill = PatternFill("solid", fgColor=C_PHASE_FILL)
    pc.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    for cc in range(1, last_col + 1):
        ws.cell(row=r, column=cc).border = BORDER
    ws.row_dimensions[r].height = 22
    r += 1

    for (tid, comp, task, desc, dur, impact, downtime,
         start, end, dep, action, status) in tasks:
        values = [tid, "", comp, task, desc, dur, impact, downtime,
                  start, end, dep, action, status]
        for i, v in enumerate(values, start=1):
            c = ws.cell(row=r, column=i, value=v)
            c.font = Font(name=FONT, size=10)
            c.border = BORDER
            c.alignment = Alignment(
                horizontal="center" if i in (1, 6, 8, 9, 10, 13) else "left",
                vertical="top",
                wrap_text=True,
            )
        # emphasis
        if status == "Contingency":
            fill = PatternFill("solid", fgColor=C_CONTINGENCY_FILL)
            for i in range(1, last_col + 1):
                ws.cell(row=r, column=i).fill = fill
        elif status in ("Pending", "In progress - internal"):
            fill = PatternFill("solid", fgColor=C_PENDING_FILL)
            for i in range(1, last_col + 1):
                ws.cell(row=r, column=i).fill = fill

        dcell = ws.cell(row=r, column=8)
        if str(downtime).lower().startswith("yes"):
            dcell.fill = PatternFill("solid", fgColor=C_DOWNTIME_FILL)
            dcell.font = Font(name=FONT, size=10, bold=True, color="C00000")

        ws.cell(row=r, column=1).font = Font(name=FONT, size=10, bold=True)
        r += 1

last_data_row = r - 1

ws.freeze_panes = "A6"
ws.auto_filter.ref = f"A{HDR_ROW}:{get_column_letter(last_col)}{last_data_row}"
ws.sheet_view.zoomScale = 90
ws.page_setup.orientation = "landscape"
ws.page_setup.fitToWidth = 1
ws.sheet_properties.pageSetUpPr.fitToPage = True
ws.print_title_rows = f"{HDR_ROW}:{HDR_ROW}"

# =====================================================================
# Sheet 2 - Summary
# =====================================================================
sm = wb.create_sheet("Summary")
sm.column_dimensions["A"].width = 4
sm.column_dimensions["B"].width = 52
sm.column_dimensions["C"].width = 18
sm.column_dimensions["D"].width = 60

sm.merge_cells("B2:D2")
c = sm.cell(row=2, column=2, value="Engagement Summary")
c.font = Font(name=FONT, size=14, bold=True, color="FFFFFF")
c.fill = PatternFill("solid", fgColor=C_TITLE_FILL)
c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
sm.row_dimensions[2].height = 26

SHEET_REF = "'Activity Schedule'"
rng_status = f"{SHEET_REF}!$M${first_data_row}:$M${last_data_row}"
rng_down = f"{SHEET_REF}!$H${first_data_row}:$H${last_data_row}"
rng_id = f"{SHEET_REF}!$A${first_data_row}:$A${last_data_row}"

summary_rows = [
    ("Total activities in scope", f'=COUNTA({rng_status})',
     "All numbered tasks across Phases 0 to 4."),
    ("Activities with confirmed dates", f'=COUNTIF({rng_status},"Planned")',
     "Internal change approval obtained; dates proposed to customer."),
    ("Activities awaiting internal governance",
     f'=COUNTIF({rng_status},"In progress - internal")+COUNTIF({rng_status},"Pending")',
     "CloudScopeLite and CloudMiddleWareKafka. Dates indicative only."),
    ("Contingency activities", f'=COUNTIF({rng_status},"Contingency")',
     "Executed only if post-change verification fails."),
    ("Activities requiring downtime", f'=COUNTIF({rng_down},"Yes*")',
     "Each requires an agreed maintenance window and end-user notification."),
    ("Read-only activities (no service impact)",
     f'=COUNTIF({SHEET_REF}!$G${first_data_row}:$G${last_data_row},"None - read-only*")',
     "Safe to perform outside a maintenance window."),
]

hr = 4
for lbl in ["Metric", "Count", "Notes"]:
    c = sm.cell(row=hr, column=2 + ["Metric", "Count", "Notes"].index(lbl), value=lbl)
    c.font = Font(name=FONT, size=10, bold=True, color="FFFFFF")
    c.fill = PatternFill("solid", fgColor=C_HDR_FILL)
    c.alignment = Alignment(horizontal="center", vertical="center")
    c.border = BORDER
sm.row_dimensions[hr].height = 20

rr = hr + 1
for label, formula, note in summary_rows:
    a = sm.cell(row=rr, column=2, value=label)
    b = sm.cell(row=rr, column=3, value=formula)
    d = sm.cell(row=rr, column=4, value=note)
    for cc in (a, b, d):
        cc.font = Font(name=FONT, size=10)
        cc.border = BORDER
        cc.alignment = Alignment(vertical="center", wrap_text=True)
    b.alignment = Alignment(horizontal="center", vertical="center")
    b.font = Font(name=FONT, size=10, bold=True)
    rr += 1

# key milestone block
rr += 2
sm.merge_cells(start_row=rr, start_column=2, end_row=rr, end_column=4)
c = sm.cell(row=rr, column=2, value="Key Milestones Requiring Customer Confirmation")
c.font = Font(name=FONT, size=12, bold=True, color="FFFFFF")
c.fill = PatternFill("solid", fgColor=C_HDR_FILL)
c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
rr += 1

MILESTONES = [
    ("26-Aug-2026", "Kick-off call and scheduling agreement",
     "Confirm attendees and nominate points of contact."),
    ("28-Aug-2026", "Remote access confirmed and validated",
     "Access must be working before 01-Sep, or Phase 1 slips."),
    ("01-Sep to 04-Sep-2026", "Phase 1 information collection window",
     "Read-only. No downtime required."),
    ("16-Sep-2026", "Method of Procedure walkthrough and sign-off",
     "Customer approval gate before implementation."),
    ("19-Sep to 20-Sep-2026", "ManageOne and FusionSphere maintenance window",
     "Weekend window. Downtime required. End-user notification needed."),
    ("28-Sep-2026", "Closure report issued", "Includes forward certificate expiry schedule."),
    ("To be confirmed", "CloudScopeLite and CloudMiddleWareKafka schedule",
     "Issued on completion of internal change governance."),
]
for hdr in ["Date", "Milestone", "Notes"]:
    c = sm.cell(row=rr, column=2 + ["Date", "Milestone", "Notes"].index(hdr), value=hdr)
    c.font = Font(name=FONT, size=10, bold=True)
    c.fill = PatternFill("solid", fgColor=C_PHASE_FILL)
    c.border = BORDER
    c.alignment = Alignment(horizontal="center", vertical="center")
rr += 1
for dt, ms, note in MILESTONES:
    vals = [dt, ms, note]
    for i, v in enumerate(vals):
        c = sm.cell(row=rr, column=2 + i, value=v)
        c.font = Font(name=FONT, size=10)
        c.border = BORDER
        c.alignment = Alignment(vertical="center", wrap_text=True)
    rr += 1

sm.sheet_view.showGridLines = False

# =====================================================================
# Sheet 3 - Customer Prerequisites
# =====================================================================
pr = wb.create_sheet("Customer Prerequisites")
pr.column_dimensions["A"].width = 4
pr.column_dimensions["B"].width = 8
pr.column_dimensions["C"].width = 40
pr.column_dimensions["D"].width = 62
pr.column_dimensions["E"].width = 20
pr.column_dimensions["F"].width = 18
pr.column_dimensions["G"].width = 28

pr.merge_cells("B2:G2")
c = pr.cell(row=2, column=2, value="Customer Prerequisites and Actions Required")
c.font = Font(name=FONT, size=14, bold=True, color="FFFFFF")
c.fill = PatternFill("solid", fgColor=C_TITLE_FILL)
c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
pr.row_dimensions[2].height = 26

pr.merge_cells("B3:G3")
c = pr.cell(row=3, column=2,
            value="Yellow cells are for the customer to complete. Please return this tab ahead of the kick-off call.")
c.font = Font(name=FONT, size=9, italic=True)
c.alignment = Alignment(vertical="center", indent=1)

PRE_HDRS = ["Ref", "Prerequisite", "Why it is required", "Required by", "Status", "Customer owner / comment"]
hr = 5
for i, h in enumerate(PRE_HDRS):
    c = pr.cell(row=hr, column=2 + i, value=h)
    c.font = Font(name=FONT, size=10, bold=True, color="FFFFFF")
    c.fill = PatternFill("solid", fgColor=C_HDR_FILL)
    c.border = BORDER
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
pr.row_dimensions[hr].height = 28

PREREQS = [
    ("P01", "Remote access method confirmed (VPN / jump host / bastion)",
     "Engineers cannot begin any activity without validated connectivity. Access issues are the most common cause of schedule slippage.",
     "27-Aug-2026"),
    ("P02", "Engineer accounts provisioned with required privilege level",
     "Certificate operations require elevated privileges. Read-only accounts are sufficient for Phase 1 but not for implementation.",
     "27-Aug-2026"),
    ("P03", "Source IP addresses whitelisted",
     "Required where access is restricted by firewall or security policy.",
     "27-Aug-2026"),
    ("P04", "Technical and business points of contact nominated",
     "Needed for go / no-go decisions and escalation during the maintenance window.",
     "26-Aug-2026"),
    ("P05", "Escalation matrix and out-of-hours contact numbers provided",
     "Implementation runs in a weekend window; a reachable decision-maker is required.",
     "16-Sep-2026"),
    ("P06", "Change freeze / blackout calendar provided",
     "To ensure the proposed windows do not conflict with business-critical periods.",
     "27-Aug-2026"),
    ("P07", "Confirmation of a recent, valid and restorable backup",
     "This is the basis of the rollback plan. Implementation will not proceed without it.",
     "28-Aug-2026"),
    ("P08", "Authoritative NTP source confirmed",
     "Certificate validation is time-sensitive. Clock drift is a frequent cause of post-change authentication failures.",
     "03-Sep-2026"),
    ("P09", "End-user notification issued for the maintenance window",
     "Management portal and API access will be interrupted during the change.",
     "17-Sep-2026"),
    ("P10", "Inventory of applications dependent on the Kafka service",
     "Certificate replacement will require a rolling restart; dependent applications may need to reconnect.",
     "To be confirmed"),
    ("P11", "Customer acceptance test criteria agreed",
     "Defines what 'successful' looks like and what would trigger rollback.",
     "16-Sep-2026"),
    ("P12", "Method of Procedure reviewed and approved",
     "Formal approval gate prior to any change being made in the environment.",
     "16-Sep-2026"),
]

dv = DataValidation(type="list", formula1='"Not started,In progress,Complete,Not applicable"',
                    allow_blank=True)
pr.add_data_validation(dv)

rr = hr + 1
example_row = rr
for ref, name, why, by in PREREQS:
    vals = [ref, name, why, by, "Not started", ""]
    for i, v in enumerate(vals):
        c = pr.cell(row=rr, column=2 + i, value=v)
        c.font = Font(name=FONT, size=10)
        c.border = BORDER
        c.alignment = Alignment(vertical="center", wrap_text=True,
                                horizontal="center" if i in (0, 3, 4) else "left")
    for col in (6, 7):
        pr.cell(row=rr, column=col).fill = PatternFill("solid", fgColor=C_INPUT_FILL)
    dv.add(pr.cell(row=rr, column=6))
    rr += 1

pr.cell(row=example_row, column=7,
        value="Example: J. Tan - VPN access requested, ticket INC0012345")
pr.cell(row=example_row, column=7).font = Font(name=FONT, size=9, italic=True, color="808080")

pr.freeze_panes = "B6"
pr.sheet_view.showGridLines = False
pr.page_setup.orientation = "landscape"
pr.sheet_properties.pageSetUpPr.fitToPage = True
pr.page_setup.fitToWidth = 1

# =====================================================================
# Sheet 4 - Assumptions, Risks & Legend
# =====================================================================
ar = wb.create_sheet("Assumptions & Risks")
ar.column_dimensions["A"].width = 4
ar.column_dimensions["B"].width = 8
ar.column_dimensions["C"].width = 56
ar.column_dimensions["D"].width = 66

ar.merge_cells("B2:D2")
c = ar.cell(row=2, column=2, value="Assumptions, Risks and Legend")
c.font = Font(name=FONT, size=14, bold=True, color="FFFFFF")
c.fill = PatternFill("solid", fgColor=C_TITLE_FILL)
c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
ar.row_dimensions[2].height = 26


def section(row, title):
    ar.merge_cells(start_row=row, start_column=2, end_row=row, end_column=4)
    cc = ar.cell(row=row, column=2, value=title)
    cc.font = Font(name=FONT, size=11, bold=True, color="FFFFFF")
    cc.fill = PatternFill("solid", fgColor=C_HDR_FILL)
    cc.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    return row + 1


def table(row, headers, rows):
    for i, h in enumerate(headers):
        cc = ar.cell(row=row, column=2 + i, value=h)
        cc.font = Font(name=FONT, size=10, bold=True)
        cc.fill = PatternFill("solid", fgColor=C_PHASE_FILL)
        cc.border = BORDER
        cc.alignment = Alignment(horizontal="center", vertical="center")
    row += 1
    for vals in rows:
        for i, v in enumerate(vals):
            cc = ar.cell(row=row, column=2 + i, value=v)
            cc.font = Font(name=FONT, size=10)
            cc.border = BORDER
            cc.alignment = Alignment(vertical="center", wrap_text=True,
                                     horizontal="center" if i == 0 else "left")
        row += 1
    return row + 1


r2 = 4
r2 = section(r2, "Assumptions")
r2 = table(r2, ["Ref", "Assumption", "Impact if not valid"], [
    ("A01", "All dates are proposed and subject to customer confirmation.",
     "Schedule will be reissued following customer feedback."),
    ("A02", "Remote access is available and validated by 28-Aug-2026.",
     "Phase 1 start date of 01-Sep-2026 will move, shifting all downstream dates."),
    ("A03", "Phase 1 activities are read-only and carry no service impact.",
     "If write access proves necessary, a separate change approval will be sought first."),
    ("A04", "The environment is in a healthy state prior to commencement.",
     "Pre-existing faults must be resolved before certificate replacement proceeds."),
    ("A05", "A valid and restorable backup exists prior to implementation.",
     "Implementation will not proceed. Rollback capability cannot be assured without it."),
    ("A06", "The 19-20 Sep weekend window is acceptable to the customer.",
     "An alternative window will be agreed; Phase 3 dates would move accordingly."),
    ("A07", "Final runbook is issued by 15-Sep-2026 as scheduled.",
     "Implementation window would be deferred to the following weekend."),
    ("A08", "Virtual machine workloads remain in service during FusionSphere management-plane changes.",
     "To be confirmed against the final runbook once issued."),
    ("A09", "Phase 4 components follow the same phased approach as Phases 1 to 3.",
     "Approach will be confirmed when internal governance completes."),
])

r2 = section(r2, "Risks and Mitigations")
r2 = table(r2, ["Ref", "Risk", "Mitigation"], [
    ("R01", "Certificates already expired may prevent normal administrative login to affected components.",
     "Verified during Phase 1 pre-checks. Alternative access paths identified before implementation."),
    ("R02", "Clock drift across nodes causes certificate validation to fail after replacement.",
     "NTP synchronisation explicitly verified in task 1.5 before any change is made."),
    ("R03", "Undiscovered dependencies on the affected certificates cause unexpected service impact.",
     "Component topology and dependency mapping performed during Phase 1 information collection."),
    ("R04", "Further certificates expire shortly after this engagement, requiring a repeat outage.",
     "Extended 6-month expiry scan performed in task 1.6 so they can be bundled into the same window."),
    ("R05", "Post-change verification fails and rollback is required.",
     "Snapshot taken in task 3.1. Rollback procedure defined in task 3.6 and agreed at the walkthrough."),
    ("R06", "Maintenance window overruns beyond the agreed period.",
     "Go / no-go checkpoint before each component. Agreed cut-off time after which rollback is invoked."),
    ("R07", "Kafka certificate replacement disrupts dependent applications.",
     "Dependent application inventory collected in advance (P10). Application team on standby during change."),
    ("R08", "Access or credential issues delay the start of a booked window.",
     "Access validated in task 0.2, well in advance of the first scheduled activity."),
    ("R09", "Phase 4 internal governance completes later than anticipated.",
     "Phase 4 dates shown as indicative only. Confirmed schedule issued on completion."),
])

r2 = section(r2, "Legend")
r2 = table(r2, ["Marker", "Meaning", "Notes"], [
    ("Planned", "Internal change approval obtained; dates proposed to customer.",
     "Standard white rows in the Activity Schedule tab."),
    ("Pending", "Awaiting completion of internal change governance.",
     "Shown with an amber background. Dates indicative only."),
    ("In progress - internal", "Currently within our internal change process.",
     "No customer action required at this stage."),
    ("Contingency", "Executed only if a preceding verification step fails.",
     "Shown with an orange background."),
    ("Downtime: Yes", "An agreed maintenance window and end-user notification are required.",
     "Highlighted in red in the Downtime Required column."),
    ("Yellow cell", "Input cell for the customer to complete.",
     "Used on the Customer Prerequisites tab."),
])

ar.sheet_view.showGridLines = False
ar.page_setup.orientation = "landscape"
ar.sheet_properties.pageSetUpPr.fitToPage = True
ar.page_setup.fitToWidth = 1

OUT = "/home/user/gitops-platform-demo/out/Certificate_Renewal_Activity_Schedule.xlsx"
wb.active = 0
# openpyxl writes formulas without cached values; force Excel to compute on open.
wb.calculation.fullCalcOnLoad = True
wb.save(OUT)
print("saved:", OUT)
print("data rows:", first_data_row, "->", last_data_row)
