"""
Uses win32com to open 'HW Q3 Problem_Agency MBS Research 2.xlsx' in Excel,
add a 'Section 3 Charts' sheet, and insert the Section 3 chart images with
labeled titles in a neat 2-column layout.
"""

import os
import win32com.client as win32
import time

BASE_DIR = r"c:\Users\nduta\OneDrive\Desktop\Projects\mbs"
Q3_FILE  = os.path.join(BASE_DIR, "HW Q3 Problem_Agency MBS Research 2.xlsx")
CHARTS_DIR = os.path.join(BASE_DIR, "charts")

# Section 3 charts in order: (filename, label)
SECTION3_CHARTS = [
    ("q3_1_issuance.png",       "Q3.1 – Issuance by Year & Loan Purpose"),
    ("q3_3a_balance.png",       "Q3.3a – Current Balance Over Time"),
    ("q3_3b_cpr.png",           "Q3.3b – Historical CPR Over Time"),
    ("q3_4_wac_cpr.png",        "Q3.4 – WAC vs CPR (Line Chart)"),
    ("q3_4_wac_cpr_scatter.png","Q3.4 – WAC vs CPR (Scatter Chart)"),
]

# Layout: 2 columns, each chart ~9 columns wide, ~22 rows tall
# Left column starts at col 1 (A), right column at col 11 (K)
# Each chart block: 1 title row + 22 image rows + 1 gap row = 24 rows per block row
IMG_WIDTH_PTS  = 390   # image width in points
IMG_HEIGHT_PTS = 245   # image height in points
BLOCK_HEIGHT   = 260   # vertical spacing per block row in points
TITLE_ROW_H    = 18    # title row height in points

LEFT_COL_OFFSET  = 5   # left margin in points
RIGHT_COL_OFFSET = 415 # right column left offset in points
TOP_OFFSET       = 30  # top margin in points

def col_left_pos(is_left: bool) -> float:
    return LEFT_COL_OFFSET if is_left else RIGHT_COL_OFFSET

def row_top_pos(block_row: int) -> float:
    return TOP_OFFSET + block_row * BLOCK_HEIGHT

def main():
    excel = win32.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    wb = excel.Workbooks.Open(Q3_FILE)

    # Remove existing 'Section 3 Charts' sheet if present
    for sh in wb.Sheets:
        if sh.Name == "Section 3 Charts":
            sh.Delete()
            break

    # Add new sheet at end
    wb.Sheets.Add(After=wb.Sheets(wb.Sheets.Count))
    ws = wb.Sheets(wb.Sheets.Count)
    ws.Name = "Section 3 Charts"

    # Style the sheet tab color (light blue)
    ws.Tab.Color = 0xD6E4F0

    for idx, (fname, label) in enumerate(SECTION3_CHARTS):
        img_path = os.path.join(CHARTS_DIR, fname)
        if not os.path.exists(img_path):
            print(f"  [SKIP] {fname} not found")
            continue

        block_row = idx // 2
        is_left   = (idx % 2 == 0)

        x_pos = col_left_pos(is_left)
        y_top = row_top_pos(block_row)

        # ----- Add title text box -----
        title_box = ws.Shapes.AddTextbox(
            1,           # msoTextOrientationHorizontal
            x_pos,       # Left
            y_top,       # Top
            IMG_WIDTH_PTS,  # Width
            TITLE_ROW_H     # Height
        )
        title_box.TextFrame.Characters().Text = label
        title_box.TextFrame.Characters().Font.Bold = True
        title_box.TextFrame.Characters().Font.Size = 11
        title_box.TextFrame.Characters().Font.Color = 0x1F4E79  # Dark blue (BGR for win32)
        title_box.Fill.ForeColor.RGB = 0xF0E4D6   # Light blue fill (BGR)
        title_box.Fill.Visible = True
        title_box.Line.Visible = False

        # ----- Insert chart image -----
        img_top = y_top + TITLE_ROW_H + 3
        pic = ws.Shapes.AddPicture(
            img_path,    # Filename
            False,       # LinkToFile
            True,        # SaveWithDocument
            x_pos,       # Left
            img_top,     # Top
            IMG_WIDTH_PTS,   # Width
            IMG_HEIGHT_PTS   # Height
        )

        print(f"  [OK] '{label}' inserted at ({x_pos:.0f}, {y_top:.0f})")

    # Save and close
    wb.Save()
    wb.Close()
    excel.Quit()
    print(f"\nDone. Saved: {Q3_FILE}")

if __name__ == "__main__":
    main()
