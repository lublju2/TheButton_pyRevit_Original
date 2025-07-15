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

def read_excel_worksheets(path):
    """Read all worksheets and their data from Excel."""
    log(u"📂  Opening Excel workbook: {0}".format(path))
    app = Excel.ApplicationClass()
    app.Visible = False
    app.DisplayAlerts = False
    worksheets_data = []
    
    try:
        wb = app.Workbooks.Open(path)
        for ws in wb.Worksheets:
            sheet_name = ws.Name
            
            # Skip Splash Screen tab
            if sheet_name == 'Splash Screen':
                log(u"⏭️  Skipping '{0}' tab".format(sheet_name))
                continue
                
            # Check H1 cell
            h1_value = ws.Range["H1"].Value2
            if h1_value != 'Yes':
                log(u"⏭️  Skipping '{0}' tab (H1 = {1})".format(sheet_name, h1_value))
                continue
                
            # Read J3 cell content
            j3_value = ws.Range["J3"].Value2
            if j3_value:
                worksheets_data.append({
                    'title': sheet_name,
                    'content': j3_value
                })
                log(u"✅  Added '{0}' tab for processing".format(sheet_name))
            else:
                log(u"⚠️  '{0}' tab has empty J3 cell".format(sheet_name))
                
    except Exception as ex:
        log(u"❌  Excel read error: {0}".format(ex))
    finally:
        if 'wb' in locals():
            wb.Close(False)
        app.Quit()
    
    return worksheets_data

# ─────────────────────────────────────────────────────────────────────────────
# Constants
# ─────────────────────────────────────────────────────────────────────────────
EXCEL_PATH = r"I:\\BLU - Service Delivery\\11 Innovations\\Parametrics\\00 - DIG1\\03 - Tools\\19 - General Notes (Excel to Revit)\\GenNotes-EWP-XX-XX-PS-S-General_Notes (version 1).xlsm"

# Layout settings
START_X = 0.0  # Starting X position
START_Y = 0.0  # Starting Y position
TITLE_SPACING = 0.5  # Space between title and content
SECTION_SPACING = 1.0  # Space between sections
COLUMN_WIDTH = 8.0  # Width of each column
PAGE_HEIGHT = 11.0  # Height of page before starting new column

# ─────────────────────────────────────────────────────────────────────────────
# Load Revit API
# ─────────────────────────────────────────────────────────────────────────────
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
from Autodesk.Revit.DB import (
    TextNote,
    Transaction,
    XYZ,
    TextNoteOptions,
    BuiltInCategory,
    FilteredElementCollector
)
from Autodesk.Revit.UI import TaskDialog

# ─────────────────────────────────────────────────────────────────────────────
# Read Excel data
# ─────────────────────────────────────────────────────────────────────────────
# worksheets_data = read_excel_worksheets(EXCEL_PATH)
# if not worksheets_data:
#     TaskDialog.Show("No Data", "No valid worksheets found to process.")
#     sys.exit()

# ─────────────────────────────────────────────────────────────────────────────
# Create TextNotes in Revit
# ─────────────────────────────────────────────────────────────────────────────
uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document

tx = Transaction(doc, "Create TextNotes from Excel")
tx.Start()

try:
    # Get text note types
    title_type = None
    content_type = None
    
    text_note_types = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_TextNotes).WhereElementIsElementType().ToElements()

    log(u"📋  Found {0} text note types".format(len(text_note_types)))

    # Debug: List all available text note types
    log(u"📋  Available text note types:")
    for elem in text_note_types:
        if hasattr(elem, 'Name'):
            log(u"   - {0}".format(elem.Name))
            if elem.Name == 'EWP_3.5mm Arrow Masking':
                title_type = elem
            elif elem.Name == 'EWP_3.5mm Arrow':
                content_type = elem
    
    if not title_type or not content_type:
        log(u"❌  No text note types available")
        tx.RollBack()
        TaskDialog.Show("Error", "No text note types found in the document.")
        sys.exit()
    
    current_y = START_Y
    current_column = 0
    created_notes = 0
    
    for data in worksheets_data:
        # Check if we need to start a new column
        if current_y > PAGE_HEIGHT:
            current_column += 1
            current_y = START_Y
        
        current_x = START_X + (current_column * COLUMN_WIDTH)
        
        # Create title TextNote
        title_point = XYZ(current_x, current_y, 0)
        title_options = TextNoteOptions()
        title_options.TypeId = title_type.Id
        
        title_note = TextNote.Create(doc, doc.ActiveView.Id, title_point, data['title'], title_options)
        current_y -= TITLE_SPACING
        created_notes += 1
        
        # Create content TextNote
        content_point = XYZ(current_x, current_y, 0)
        content_options = TextNoteOptions()
        content_options.TypeId = content_type.Id
        
        content_note = TextNote.Create(doc, doc.ActiveView.Id, content_point, data['content'], content_options)
        current_y -= SECTION_SPACING
        created_notes += 1
        
        log(u"📝  Created notes for '{0}'".format(data['title']))
    
    tx.Commit()
    log(u"✅  Done. Created {0} TextNotes.".format(created_notes))
    TaskDialog.Show("Finished", u"Created {0} TextNotes from {1} Excel tabs.".format(created_notes, len(worksheets_data)))
    
except Exception as ex:
    tx.RollBack()
    log(u"❌  Error creating TextNotes: {0}".format(ex))
    TaskDialog.Show("Error", u"Error creating TextNotes: {0}".format(ex))
