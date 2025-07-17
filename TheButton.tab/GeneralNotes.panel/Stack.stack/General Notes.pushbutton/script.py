# -*- coding: utf-8 -*-
# Script: Replace TextNotes from Excel (Partial-Match)
# Version: 1.2.0 – July 2025
# Author: AO

from __future__ import unicode_literals     # IronPython 2.7 compatibility
import clr
import sys
import os
import re

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

# Layout settings
START_X = 8.883660091           # Starting X position (top left corner)
START_Y = 4.440310784           # Starting Y position (top left corner)
BOTTOM_Y = 2.794031111          # Bottom Y position of the page
SECTION_SPACING = 0.002         # Small gap between title and content
INTER_SECTION_SPACING = 0.1     # Gap between sections
COLUMN_WIDTH = 0.37             # Width of each column
PAGE_HEIGHT = START_Y - BOTTOM_Y  # Calculate page height from coordinates
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
    # Cannot simply be extracted from the TextNote because that information is not available until the transaction is committed. Therefore we need to estimate it based on the TextNoteType parameters and the content.
    try:
        # Get text size from the TextNoteType
        text_size_param = text_note_type.get_Parameter(BuiltInParameter.TEXT_SIZE)
        if text_size_param:
            text_size = text_size_param.AsDouble()
        else:
            log(u"⚠️  No TEXT_SIZE parameter found for TextNoteType '{0}'".format(text_note_type.Name))
            text_size = 0.1  # fallback text size
        
        # Estimate line height including spacing
        line_height = text_size * 1.6
        
        # Estimated width of each character
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


def check_text_note_fits(current_y, text_height, bottom_boundary):
    """Check if a TextNote would fit within the page boundaries."""
    return (current_y - text_height) >= bottom_boundary

# ─────────────────────────────────────────────────────────────────────────────
# Helper function to find the Excel file
# ─────────────────────────────────────────────────────────────────────────────

def find_excel_file(doc):
    """Find the Excel file based on the current Revit document's location."""
    try:
        # Get current Revit document path
        doc_path = doc.PathName
        if not doc_path:
            TaskDialog.Show("Error", 
                "Please save the Revit file first before running this script.")
            return None
        
        log(u"📁  Current Revit file: {0}".format(doc_path))
        
        # Get the folder containing the Revit file
        revit_folder = os.path.dirname(doc_path)
        log(u"📁  Revit folder: {0}".format(revit_folder))
        
        # Go up one folder
        parent_folder = os.path.dirname(revit_folder)
        log(u"📁  Parent folder: {0}".format(parent_folder))
        
        # Check if the Revit file is in a folder named "00 Revit Model (20XX)"
        revit_folder_name = os.path.basename(revit_folder)
        if not re.match(r'^00 Revit Model \(20\d{2}\)$', revit_folder_name):
            TaskDialog.Show("Folder Structure Error", 
                "The Revit file should be placed in a folder named '00 Revit Model (20XX)'.\n\n" +
                "Current location: {0}\n\n".format(revit_folder) +
                "Expected structure:\n" +
                "📁 01 Structural\n" +
                "   📁 00 Revit Model (20XX)\n" +
                "      📄 [Your Revit file]\n" +
                "   📁 01 Linked Files\n" +
                "      📁 EWP\n" +
                "         📄 GenNotes*.xlsm")
            return None
        
        # Check if parent folder is named "01 Structural"
        if not os.path.basename(parent_folder) == "01 Structural":
            TaskDialog.Show("Folder Structure Error", 
                "The Revit file should be placed in a '00 Revit Model (20XX)' folder within '01 Structural'.\n\n" +
                "Current location: {0}\n\n".format(parent_folder) +
                "Expected structure:\n" +
                "📁 01 Structural\n" +
                "   📁 00 Revit Model (20XX)\n" +
                "      📄 [Your Revit file]\n" +
                "   📁 01 Linked Files\n" +
                "      📁 EWP\n" +
                "         📄 GenNotes*.xlsm")
            return None
        
        # Navigate to "01 Linked Files" folder
        linked_files_folder = os.path.join(parent_folder, "01 Linked Files")
        if not os.path.exists(linked_files_folder):
            TaskDialog.Show("Folder Structure Error", 
                "The '01 Linked Files' folder was not found.\n\n" +
                "Expected location: {0}\n\n".format(linked_files_folder) +
                "Please create the following folder structure:\n" +
                "📁 01 Structural\n" +
                "   📁 00 Revit Model (20XX)\n" +
                "      📄 [Your Revit file]\n" +
                "   📁 01 Linked Files\n" +
                "      📁 EWP\n" +
                "         📄 GenNotes*.xlsm")
            return None
        
        # Navigate to "EWP" folder
        ewp_folder = os.path.join(linked_files_folder, "EWP")
        if not os.path.exists(ewp_folder):
            TaskDialog.Show("Folder Structure Error", 
                "The 'EWP' folder was not found.\n\n" +
                "Expected location: {0}\n\n".format(ewp_folder) +
                "Please create the following folder structure:\n" +
                "📁 01 Structural\n" +
                "   📁 00 Revit Model (20XX)\n" +
                "      📄 [Your Revit file]\n" +
                "   📁 01 Linked Files\n" +
                "      📁 EWP\n" +
                "         📄 GenNotes*.xlsm")
            return None
        
        # Look for Excel file starting with "GenNotes"
        excel_files = []
        for file in os.listdir(ewp_folder):
            if file.startswith("GenNotes") and file.endswith(".xlsm"):
                excel_files.append(file)
        
        if not excel_files:
            TaskDialog.Show("Excel File Not Found", 
                "No Excel file starting with 'GenNotes' and ending with '.xlsm' was found.\n\n" +
                "Expected location: {0}\n\n".format(ewp_folder) +
                "Please place the GenNotes Excel file in the EWP folder:\n" +
                "📁 01 Structural\n" +
                "   📁 00 Revit Model (20XX)\n" +
                "      📄 [Your Revit file]\n" +
                "   📁 01 Linked Files\n" +
                "      📁 EWP\n" +
                "         📄 GenNotes*.xlsm")
            return None
        
        # Use the first matching Excel file
        excel_file = excel_files[0]
        excel_path = os.path.join(ewp_folder, excel_file)
        
        if len(excel_files) > 1:
            log(u"⚠️  Multiple GenNotes files found, using: {0}".format(excel_file))
        
        log(u"✅  Found Excel file: {0}".format(excel_path))
        return excel_path
        
    except Exception as ex:
        log(u"❌  Error finding Excel file: {0}".format(ex))
        TaskDialog.Show("Error", 
            "Error finding Excel file: {0}".format(ex))
        return None


# ─────────────────────────────────────────────────────────────────────────────
# Start up document and find Excel file
# ─────────────────────────────────────────────────────────────────────────────
uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document

EXCEL_PATH = find_excel_file(doc)


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
        # Calculate heights before placing
        title_height = calculate_text_note_height(title_type, data['title'], title_width)
        content_height = calculate_text_note_height(content_type, data['content'], content_width)
        
        # Calculate total height needed for this section
        total_section_height = title_height + SECTION_SPACING + content_height + INTER_SECTION_SPACING
        
        # Check if the entire section would fit, if not move to next column
        if not check_text_note_fits(current_y, total_section_height, BOTTOM_Y):
            current_column += 1
            current_y = START_Y
            log(u"📄  Moving to column {0} for section '{1}'".format(current_column + 1, data['title']))

        current_x = START_X + (current_column * COLUMN_WIDTH)

        # Create title TextNote
        title_point = XYZ(current_x, current_y, 0)
        title_options = TextNoteOptions()
        title_options.TypeId = title_type.Id

        title_note = TextNote.Create(doc, doc.ActiveView.Id, title_point, title_width, data['title'], title_options)
        created_notes += 1
        
        # Adjust Y position after title
        current_y -= (title_height + SECTION_SPACING)
        log(u"📏  Title '{0}' height: {1}".format(data['title'], title_height))

        # Create content TextNote
        content_point = XYZ(current_x, current_y, 0)
        content_options = TextNoteOptions()
        content_options.TypeId = content_type.Id

        content_note = TextNote.Create(doc, doc.ActiveView.Id, content_point, content_width, data['content'], content_options)
        created_notes += 1
        
        # Adjust Y position for next section
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
