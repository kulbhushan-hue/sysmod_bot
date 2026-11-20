from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
import kb
H = list(kb.ALIASES)
wb = Workbook(); ws = wb.active; ws.title = "SYSMOD_DATA"
ws.append(H)
ws.append(["EXAMPLE VENDOR", "EXAMPLE PRODUCT", "V1", "2.5, 3.1, 3.2", "XX12345", "PTF", "HXXX110",
           "XX12344", "2026-01-15", "XX12000", "", "Example description of the fix", "https://example.com", "EXAMPLE"])
for c in ws[1]:
    c.font = Font(name="Arial", bold=True, color="FFFFFF"); c.fill = PatternFill("solid", fgColor="002D72")
for row in ws.iter_rows(min_row=2):
    for c in row: c.font = Font(name="Arial", italic=True, color="808080")
for col in ws.columns: ws.column_dimensions[col[0].column_letter].width = 20
ws.freeze_panes = "A2"
lg = wb.create_sheet("Legend")
for r in [["How to use this file"],
          ["1. Paste one row per SYSMOD (PTF / APAR / USERMOD). Row 2 is an example (Status = EXAMPLE) and is ignored by the bot."],
          ["2. Vendor + Product must match the names in your ISV workbook (the match is tolerant: 'CA Cleanup for RACF' matches 'Cleanup for RACF')."],
          ["3. z/OS Release: any of 2.5, 3.1, 3.2 / V2R5, V3R1, V3R2, comma separated or '2.5-3.2'. Leave blank if unknown (shown as '?')."],
          ["4. Your existing scraped vendor files also work as-is: columns 'PTF ID', 'FMID', 'APAR ID', 'Prerequisites', 'Superseding PTF', 'Full description' are recognised."],
          ["5. Save the file into the kb_data/ folder, or upload it in the app sidebar."]]:
    lg.append(r)
lg["A1"].font = Font(name="Arial", bold=True, size=12)
for row in lg.iter_rows(min_row=2):
    for c in row: c.font = Font(name="Arial")
lg.column_dimensions["A"].width = 150
wb.save("SYSMOD_DATA_template.xlsx")
