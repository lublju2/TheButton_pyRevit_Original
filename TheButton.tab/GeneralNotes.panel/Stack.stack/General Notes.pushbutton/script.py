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
EXCEL_PATH = r"C:\\Users\\A.Osipova\\Desktop\\WORKING FOLDER\\TEST.xlsx"
SHEET_NAME = "Sheet1"
CELL_ADDR  = "D7"

# This exact capitalization will be searched for in every TextNote
KEY_PHRASE = "General notes notes"

# ─────────────────────────────────────────────────────────────────────────────
# Read replacement text from Excel
# ─────────────────────────────────────────────────────────────────────────────
NEW_TEXT = read_excel_cell(EXCEL_PATH, SHEET_NAME, CELL_ADDR)
if NEW_TEXT is None:
    from Autodesk.Revit.UI import TaskDialog
    TaskDialog.Show("Error", u"Failed to read {0}!{1}".format(SHEET_NAME, CELL_ADDR))
    sys.exit()

# ─────────────────────────────────────────────────────────────────────────────
# Load Revit API
# ─────────────────────────────────────────────────────────────────────────────
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
from Autodesk.Revit.DB import (
    FilteredElementCollector,
    BuiltInCategory,
    TextNote,
    Transaction
)
from Autodesk.Revit.UI import TaskDialog

# ─────────────────────────────────────────────────────────────────────────────
# Gather all TextNotes in the active document
# ─────────────────────────────────────────────────────────────────────────────
uidoc = __revit__.ActiveUIDocument
doc   = uidoc.Document

notes = FilteredElementCollector(doc) \
    .OfCategory(BuiltInCategory.OST_TextNotes) \
    .WhereElementIsNotElementType() \
    .ToElements()

log(u"📋  Found {0} TextNotes.".format(len(notes)))

# ─────────────────────────────────────────────────────────────────────────────
# Replace any note containing our key phrase
# ─────────────────────────────────────────────────────────────────────────────
tx = Transaction(doc, "Replace TextNotes from Excel")
tx.Start()
replaced = 0

for note in notes:
    text = note.Text or ""
    if KEY_PHRASE in text:
        log(u"🔄  Replacing TextNote Id={0}".format(note.Id.IntegerValue))
        note.Text = NEW_TEXT
        replaced += 1

tx.Commit()

log(u"✅  Done. Replaced {0} TextNotes.".format(replaced))
TaskDialog.Show("Finished", u"Replaced {0} TextNotes.".format(replaced))
