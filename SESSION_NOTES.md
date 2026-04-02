# Mechanisms Dashboard — Session Notes
# Paste this to Claude at the start of each session if it has forgotten context.

## Status (as of 2026-03-28)

Excel PoC pivot: built `generate_dashboard.py` which generates `mechanisms_dashboard.xlsx` — a working Excel dashboard with Power Query connections to all 5 SharePoint files, formula-driven 12-week grid, and colour-coded conditional formatting.

mechanisms_dashboard.xlsx has been uploaded to SharePoint and syncs fine.

## Excel PoC files
- `generate_dashboard.py` — Python script to regenerate the workbook
- `mechanisms_dashboard.xlsx` — the live working dashboard (uploaded to SharePoint)

## Power Query connections (all working)
- `q_Lookup` → `_Lookup` table — Mechanisms Lookup.xlsx
- `q_Sortly` → `_Sortly` table — Latest Sortly Mech Report.xlsx
- `q_Production` → `_Production` table — Production 2026 Dec-Nov.xlsx (multi-sheet, W/C date parsing fixed)
- `q_PO` → `_PO` table — PO Listing - with unit costs LATEST.xlsx (SEMINAR* suppliers excluded)
- `q_Seminar` → `_Seminar` table — Latest Repose Order Summary.xlsx / "Seminar open orders" sheet

## Correct SharePoint paths (confirmed)
- Site: https://reposefurniturelimited.sharepoint.com/sites/ReposeFurniture-PlanningRepose
- Base: Shared Documents/Planning Repose/
- Lookup: Mech Forecast/Mechanism Codes/
- Sortly: Mech Forecast/Sortly Reports/
- PO: Mech Forecast/Purchase Orders/
- Seminar: Mech Forecast/Seminar/
- Production: root of Shared Documents/Planning Repose/

## Known quirks fixed
- Production L2 date format is "W/C 01/12/2025 <extra>" — parsed with Text.Middle(Text.Start(...,14),4)
- PO file uses hyphen not en-dash in filename
- SEMINAR filter uses Text.StartsWith(Text.Upper(...), "SEMINAR") to catch all variants
- Power Query creates tables named q_*; must rename to _* after loading (convert placeholder to range first)
- Subtitle formulas use & not CONCAT (avoids @CONCAT #NAME? error)

## Vue web app (original build — on hold)
- All 12 tasks complete, pushed to github.com/MechUser-Rep/mech-planning-dashboard
- Blocked on Azure AD app registration (IT admin needed) and GitHub Pages confirmation
- config.js has placeholder CLIENT_ID / TENANT_ID

## Next steps
1. Test data refresh from SharePoint on the live file
2. Review dashboard output with real data — check calculations are correct
3. Resume Vue app when Azure AD admin access is available

## Source files info
- Production 2026 Dec-Nov.xlsx: sheets named "WK XX", week date in L2 as "W/C DD/MM/YYYY", headers row 4, columns "Mechanism - 1", "Mechanism - 2", "ITEMS"
- Sortly: col A=Entry Name, J=Quantity, L=Min Level, M=Price
- PO Listing: headers row 2, col C=AccountReference, D=PODate, H=Description, J=Quantity
- Seminar: sheet "Seminar open orders", col F=Description, G=Quantity, H=Due Date (Friday of delivery week)
- Mechanisms Lookup: sheet "Mechanisms lookup", col D=New Code, also Sage code/Sage Code 2, Seminar Code/Seminar Code 2, Lead-time
