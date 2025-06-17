# -*- coding: utf-8 -*-
# IronPython 2.7 script for pyRevit (Steps 1–3 only)

import os
import shutil
import clr
from datetime import datetime
from pyrevit import revit, DB, script, forms
from Autodesk.Revit.DB import (
    FilteredElementCollector,
    ViewSchedule,
    SectionType,
    BuiltInCategory
)

# ------------------------------------------------------------------------------
# STEP 1: Extension Location & TEMPLATE_PATH
# ------------------------------------------------------------------------------
# 1. Verify your script sits in:
#    …\TheButton.extension\TheButton.tab\Issue Sheets.panel\NewIssueSheet.pushbutton\script.py

# 2. Exact path to the .xlsm template:
TEMPLATE_PATH = r"C:\Users\A.Osipova\Desktop\WORKING FOLDER\ISSUE_SHEET\Issue Sheet.xlsm"

# 3. Verification placeholder: check the template exists
if not os.path.exists(TEMPLATE_PATH):
    print "❌ STEP 1 ERROR: TEMPLATE_PATH not found:", TEMPLATE_PATH
    script.exit()
print "✅ STEP 1: TEMPLATE_PATH exists."

# ------------------------------------------------------------------------------
# STEP 2: Prompt for Destination & Copy Template
# ------------------------------------------------------------------------------
# 2.1 Ask user to pick a folder
dest_folder = forms.pick_folder("Select folder to save the new Issue Sheet")
if not dest_folder:
    print "❌ STEP 2: No folder selected, exiting."
    script.exit()

# 2.2 Build filename (preserve .xlsm so macros stay intact)
today        = datetime.now().strftime("%Y%m%d")
new_filename = "{0}_Issue Sheet.xlsm".format(today)
dest_path    = os.path.join(dest_folder, new_filename)

# 2.3 Copy the template
shutil.copyfile(TEMPLATE_PATH, dest_path)

# 2.4 Verification placeholder: confirm the copy
if not os.path.exists(dest_path):
    print "❌ STEP 2 ERROR: Copied file not found at", dest_path
    script.exit()
print "✅ STEP 2: Template copied to", dest_path

# ------------------------------------------------------------------------------
# STEP 3: Open Workbook via COM & Verify Worksheet
# ------------------------------------------------------------------------------
# 3.1 Excel COM setup
clr.AddReference("Microsoft.Office.Interop.Excel")
import Microsoft.Office.Interop.Excel as XL

excel = XL.ApplicationClass()
excel.Visible       = False
excel.DisplayAlerts = False

# 3.2 Open the workbook
try:
    wb = excel.Workbooks.Open(dest_path)
except Exception as e:
    print "❌ STEP 3 ERROR: Failed to open workbook:", e
    excel.Quit()
    script.exit()

# 3.3 List available sheet names for debugging
sheet_names = [wb.Worksheets[i].Name for i in range(1, wb.Worksheets.Count + 1)]
print "📝 STEP 3: Available sheets in workbook:", sheet_names

# 3.4 Reference the first sheet by index
try:
    ws = wb.Worksheets[1]   # first tab
    print "✅ STEP 3: Referencing first worksheet:", ws.Name
except Exception as e:
    print "❌ STEP 3 ERROR: Cannot reference first sheet:", e
    wb.Close(False)
    excel.Quit()
    script.exit()

# 3.5 Verification placeholder: write 'X' into cell D9
try:
    ws.Cells(9, 4).Value = "X"
    print "✅ STEP 3: Placed 'X' in D9 on sheet:", ws.Name
except Exception as e:
    print "❌ STEP 3 ERROR: Cannot write to cell D9:", e
    wb.Close(False)
    excel.Quit()
    script.exit()

# 3.6 Save & clean up
wb.Save()
wb.Close(True)
excel.Quit()
del ws, wb, excel

# ------------------------------------------------------------------------------
# STEP 4: Read “EWP_Drawing List New” Schedule, filter by “Appears In Sheet List,”
#         and extract Drawing Name & Sheet Name for Excel
# ------------------------------------------------------------------------------
# --------------------------------------------------------------------------
# STEP 4: Read “EWP_Drawing List New” Schedule → Build structured records
# --------------------------------------------------------------------------
from Autodesk.Revit.DB import (
    FilteredElementCollector,
    ViewSchedule,
    SectionType,
    BuiltInCategory
)
from pyrevit import revit, script

# 4.1 – Find the schedule
sched = next(
    (s for s in FilteredElementCollector(revit.doc)
         .OfClass(ViewSchedule)
     if s.Name == "EWP_Drawing List New"),
    None
)
if not sched:
    print "❌ STEP 4 ERROR: Schedule not found."
    script.exit()
print "✅ STEP 4.1: Found schedule:", sched.Name

# 4.2 – Build field-name list:
#     • Param-backed fields → GetSchedulableField().GetName(...)
#     • Combined/calculated → ScheduleField.GetName()
field_count = sched.Definition.GetFieldCount()
hdr         = sched.GetTableData().GetSectionData(SectionType.Header)
all_fields  = []
field_src   = []

for idx in range(field_count):
    fld = sched.Definition.GetField(idx)
    try:
        # Real parameter
        sf   = fld.GetSchedulableField()
        name = sf.GetName(revit.doc)
        field_src.append("Param")
    except:
        # Combined / calculated → use the field’s display name
        name = fld.GetName().strip()
        field_src.append("ScheduleField")
    all_fields.append(name)

print "📝 STEP 4.2: Columns detected:"
for i, name in enumerate(all_fields):
    print "    • Col#{}: '{}' ({})".format(i, name, field_src[i])

# 4.3 – Map the exact names we need
required = ["Drawing Name", "Sheet Name", "Sheet Number"]
col_idx  = {}
for req in required:
    if req in all_fields:
        col_idx[req] = all_fields.index(req)
    else:
        print "❌ STEP 4 ERROR: '{}' not found among columns.".format(req)
        script.exit()
print "✅ STEP 4.3: Mapped indices:", col_idx

# 4.4 – Read body rows
body       = sched.GetTableData().GetSectionData(SectionType.Body)
nrows      = body.NumberOfRows
sheet_data = []

for r in range(nrows):
    drawing_name  = body.GetCellText(r, col_idx["Drawing Name"])
    drawing_title = body.GetCellText(r, col_idx["Sheet Name"])
    sheet_number  = body.GetCellText(r, col_idx["Sheet Number"])

    sheet_elem = next(
        (sh for sh in FilteredElementCollector(revit.doc)
            .OfCategory(BuiltInCategory.OST_Sheets)
         if sh.SheetNumber == sheet_number),
        None
    )

    sheet_data.append({
        "drawing_name":  drawing_name,
        "drawing_title": drawing_title,
        "sheet_number":  sheet_number,
        "sheet_element": sheet_elem
    })

print "✅ STEP 4.4: Collected {} records.".format(len(sheet_data))
for rec in sheet_data[:3]:
    print "   •", rec
