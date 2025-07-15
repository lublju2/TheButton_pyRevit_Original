# -*- coding: utf-8 -*-
# Script: Replace TextNotes from Excel (Partial-Match)
# Version: 1.2.0 – July 2025
# Author: AO

from __future__ import unicode_literals  # IronPython 2.7 compatibility
import clr
import sys

# ─────────────────────────────────────────────────────────────────────────────
# Simple console logger
# ─────────────────────────────────────────────────────────────────────────────
def log(msg):
    sys.stdout.write(u"{0}\n".format(msg))

# ─────────────────────────────────────────────────────────────────────────────
# Helpers: Excel reading
# ─────────────────────────────────────────────────────────────────────────────
clr.AddReference('Microsoft.Office.Interop.Excel')
import Microsoft.Office.Interop.Excel as Excel

def read_excel_cell(path, sheet_name, addr):
    """Read a single cell from Excel, then close the app."""
    log(u"📂  Opening Excel workbook: {0}".format(path))
    app = Excel.ApplicationClass()
    app.Visible = False
    app.DisplayAlerts = False
    try:
        wb = app.Workbooks.Open(path)
        ws = wb.Sheets.Item[sheet_name]
        val = ws.Range[addr].Value2
        log(u"✅  Read {0}!{1}: {2}".format(sheet_name, addr, val))
    except Exception as ex:
        log(u"❌  Excel read error: {0}".format(ex))
        val = None
    finally:
        if 'wb' in locals():
            wb.Close(False)
        app.Quit()
    return val

# ─────────────────────────────────────────────────────────────────────────────
# Constants
# ─────────────────────────────────────────────────────────────────────────────
EXCEL_PATH = r"I:\\BLU - Service Delivery\\11 Innovations\\Parametrics\\00 - DIG1\\03 - Tools\\19 - General Notes (Excel to Revit)\\GenNotes-EWP-XX-XX-PS-S-General_Notes (version 1).xlsm"
CELL_ADDR  = "J3"

# Dictionary mapping titles to their TextNote element IDs
textbox_id = {
    'General Notes': 471348,
    'Structural Design Philosophy': 471349,
    'Key Site Constraints': 471332,
}

# ─────────────────────────────────────────────────────────────────────────────
# Load Revit API
# ─────────────────────────────────────────────────────────────────────────────
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
from Autodesk.Revit.DB import (
    FilteredElementCollector,
    BuiltInCategory,
    TextNote,
    Transaction,
    ElementId,
)
from Autodesk.Revit.UI import TaskDialog

# ─────────────────────────────────────────────────────────────────────────────
# Gather all TextNotes in the active document
# ─────────────────────────────────────────────────────────────────────────────
uidoc = __revit__.ActiveUIDocument
doc   = uidoc.Document

# ─────────────────────────────────────────────────────────────────────────────
# Read from Excel and replace TextNotes
# ─────────────────────────────────────────────────────────────────────────────
tx = Transaction(doc, "Replace TextNotes from Excel")
tx.Start()
replaced = 0

for title, element_id in textbox_id.items():
    try:
        # Read from Excel sheet
        sheet_name = title.title()  # Convert to titlecase
        text = read_excel_cell(EXCEL_PATH, sheet_name, CELL_ADDR)
        if text is None:
            log(u"❌  Failed to read {0}!{1}".format(sheet_name, CELL_ADDR))
            continue
        
        # Get the specific TextNote by its element ID
        note = doc.GetElement(ElementId(element_id))
        if note and isinstance(note, TextNote):
            log(u"🔄  Replacing TextNote '{0}' (Id={1})".format(title, element_id))
            note.Text = text
            replaced += 1
        else:
            log(u"⚠️  TextNote '{0}' (Id={1}) not found or not a TextNote".format(title, element_id))
    except Exception as ex:
        log(u"❌  Error replacing TextNote '{0}' (Id={1}): {2}".format(title, element_id, ex))

tx.Commit()

log(u"✅  Done. Replaced {0} TextNotes.".format(replaced))
TaskDialog.Show("Finished", u"Replaced {0} TextNotes.".format(replaced))
