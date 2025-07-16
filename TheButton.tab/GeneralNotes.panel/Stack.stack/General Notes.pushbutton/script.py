# -*- coding: utf-8 -*-
# Script: Replace TextNotes from Excel (Partial-Match)
# Version: 1.2.0 – July 2025
# Author: AO

from __future__ import unicode_literals     # IronPython 2.7 compatibility
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
START_X = 8.883660091           # Starting X position (top left corner)
START_Y = 4.440310784           # Starting Y position (top left corner)
SECTION_SPACING = 0.01          # Small gap between title and content
INTER_SECTION_SPACING = 0.2     # Gap between sections
COLUMN_WIDTH = 8.0              # Width of each column
PAGE_HEIGHT = 11.0              # Height of page before starting new column
TEXT_WIDTH = 0.3                # Width constraint for text notes


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
    FilteredElementCollector,
    TextNoteType,
    BuiltInParameter
)
from Autodesk.Revit.UI import TaskDialog


# ─────────────────────────────────────────────────────────────────────────────
# Helper function to calculate TextNote height
# ─────────────────────────────────────────────────────────────────────────────
def calculate_text_note_height(text_note_type, text_content, text_width):
    """Calculate the height of a TextNote based on its type and content, processing paragraph by paragraph."""
    try:
        # Get text size from the TextNoteType
        text_size_param = text_note_type.get_Parameter(BuiltInParameter.TEXT_SIZE)
        if text_size_param:
            text_size = text_size_param.AsDouble()
        else:
            log(u"⚠️  No TEXT_SIZE parameter found for TextNoteType '{0}'".format(text_note_type.Name))
            text_size = 0.1  # fallback text size
        
        # Calculate line height (typically 1.2 times the text size)
        line_height = text_size * 1.2
        
        # Estimate character width (rough approximation)
        avg_char_width = text_size * 0.55
        
        # Calculate approximate characters per line
        chars_per_line = max(1, int(text_width / avg_char_width))
        
        # Split text into paragraphs
        if text_content:
            paragraphs = text_content.split('\n')
        else:
            paragraphs = ['']
        
        total_lines = 0
        
        # Calculate lines for each paragraph
        for paragraph in paragraphs:
            if paragraph.strip():  # Non-empty paragraph
                paragraph_length = len(paragraph.strip())
                paragraph_lines = max(1, (paragraph_length + chars_per_line - 1) // chars_per_line)
                total_lines += paragraph_lines
                log(u"📏  Paragraph '{0}...' length: {1}, lines: {2}".format(
                    paragraph[:20], paragraph_length, paragraph_lines))
            else:  # Empty paragraph (line break)
                total_lines += 1
                log(u"📏  Empty paragraph (line break): 1 line")
        
        # Calculate total height
        total_height = total_lines * line_height
        
        log(u"📏  Text size: {0}, Line height: {1}, Chars per line: {2}, Total lines: {3}, Total height: {4}".format(
            text_size, line_height, chars_per_line, total_lines, total_height))
        
        return total_height
        
    except Exception as ex:
        log(u"⚠️  Error calculating TextNote height: {0}".format(ex))
        return 0.2  # fallback height


# ─────────────────────────────────────────────────────────────────────────────
# Read Excel data
# ─────────────────────────────────────────────────────────────────────────────
worksheets_data = read_excel_worksheets(EXCEL_PATH)
if not worksheets_data:
    TaskDialog.Show("No Data", "No valid worksheets found to process.")
    sys.exit()


# ─────────────────────────────────────────────────────────────────────────────
# Create TextNotes in Revit
# ─────────────────────────────────────────────────────────────────────────────
uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document

def delete_all_textnotes(doc):
    """Delete all TextNotes from the current view."""
    try:
        # Get all TextNotes in the current view
        text_notes = FilteredElementCollector(doc, doc.ActiveView.Id) \
            .OfClass(TextNote) \
            .ToElements()
        
        deleted_count = 0
        if text_notes:
            log(u"🗑️  Found {0} existing TextNotes to delete".format(len(text_notes)))
            for note in text_notes:
                doc.Delete(note.Id)
                deleted_count += 1
            log(u"✅  Deleted {0} existing TextNotes".format(deleted_count))
        else:
            log(u"ℹ️  No existing TextNotes found to delete")
        
        return deleted_count
        
    except Exception as ex:
        log(u"⚠️  Error deleting TextNotes: {0}".format(ex))
        return 0

tx = Transaction(doc, "Create TextNotes from Excel")
tx.Start()

try:
    # Delete all existing TextNotes first
    delete_all_textnotes(doc)
    
    # Get text note types (styles)
    title_type = None
    content_type = None

    text_note_types = FilteredElementCollector(doc) \
        .OfClass(TextNoteType) \
        .ToElements()

    log(u"📋  Found {0} text note types".format(len(text_note_types)))

    # Debug: list all available text note types safely
    log(u"📋  Available text note types:")
    for t in text_note_types:
        # Try .Name, fallback to SYMBOL_NAME_PARAM, else placeholder
        try:
            name = t.Name
        except AttributeError:
            param = t.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
            if param:
                name = param.AsString()
            else:
                name = u"<Unnamed TextNoteType>"

        log(u"   - {0}".format(name))

        # Match by that name
        if name == 'EWP_3.5mm Arrow Masking':
            title_type = t
        elif name == 'EWP_2.5mm Arrow':
            content_type = t

    if not title_type or not content_type:
        log(u"❌  No text note types available")
        tx.RollBack()
        TaskDialog.Show("Error", "No text note types found in the document.")
        sys.exit()

    # Validate and adjust text width for both types
    title_min_width = TextNote.GetMinimumAllowedWidth(doc, title_type.Id)
    title_max_width = TextNote.GetMaximumAllowedWidth(doc, title_type.Id)
    title_width = max(title_min_width, min(TEXT_WIDTH, title_max_width))
    
    content_min_width = TextNote.GetMinimumAllowedWidth(doc, content_type.Id)
    content_max_width = TextNote.GetMaximumAllowedWidth(doc, content_type.Id)
    content_width = max(content_min_width, min(TEXT_WIDTH, content_max_width))
    
    log(u"📏  Title width: {0} (min: {1}, max: {2})".format(title_width, title_min_width, title_max_width))
    log(u"📏  Content width: {0} (min: {1}, max: {2})".format(content_width, content_min_width, content_max_width))

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

        title_note = TextNote.Create(doc, doc.ActiveView.Id, title_point, title_width, data['title'], title_options)
        created_notes += 1
        
        # Calculate height of title note and adjust Y position
        title_height = calculate_text_note_height(title_type, data['title'], title_width)
        current_y -= (title_height + SECTION_SPACING)
        log(u"📏  Title '{0}' height: {1}".format(data['title'], title_height))

        # Create content TextNote
        content_point = XYZ(current_x, current_y, 0)
        content_options = TextNoteOptions()
        content_options.TypeId = content_type.Id

        content_note = TextNote.Create(doc, doc.ActiveView.Id, content_point, content_width, data['content'], content_options)
        created_notes += 1
        
        # Calculate height of content note and adjust Y position for next section
        content_height = calculate_text_note_height(content_type, data['content'], content_width)
        current_y -= (content_height + INTER_SECTION_SPACING)
        log(u"📏  Content height: {0}".format(content_height))

        log(u"📝  Created notes for '{0}' at Y: {1}".format(data['title'], current_y))

    tx.Commit()
    log(u"✅  Done. Created {0} TextNotes.".format(created_notes))
    TaskDialog.Show(
        "Finished",
        u"Created {0} TextNotes from {1} Excel tabs.".format(created_notes, len(worksheets_data))
    )

except Exception as ex:
    # Roll back the transaction
    tx.RollBack()

    # Get full traceback
    import traceback
    tb = traceback.format_exc()

    log(u"❌  Error creating TextNotes: {0}".format(tb))

    # Show full traceback in a dialog
    TaskDialog.Show(
        "Error creating TextNotes – full traceback",
        tb
    )
    sys.exit()
