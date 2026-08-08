"""College header block for the result Excel files (shared by the live fetch
and the batch-export paths).

Layout (rows 1-4 are the header, row 5 a spacer, data starts at row 6):
  A            | B ... last column ... | last column
  [college logo]  Gopalan College of Engineering and Management
                  Accredited by NAAC, Recognized under 2(f) by UGC, ISO 9001:2015 certified
                  Approved by All India Council for Technical Education (AICTE), New Delhi
                  Affiliated to Visvesvaraya Technological University (VTU), Belagavi,
                  Karnataka Recognized by Govt. of Karnataka
"""

import os

from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.utils.units import pixels_to_EMU
from openpyxl.drawing.spreadsheet_drawing import OneCellAnchor, AnchorMarker

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
COLLEGE_LOGO = os.path.join(BASE_DIR, "college_logo.png")

HEADER_ROWS = 5
DATA_HEADER_ROW = HEADER_ROWS + 1  # row 6

DARK_BLUE = "1F3864"
GRAY = "808080"

LOGO_HEIGHT_PX = 90
LOGO_Y_OFFSET_PX = 15

COLLEGE_NAME = "Gopalan College of Engineering and Management"
SUBLINES = [
    "Accredited by NAAC, Recognized under 2(f) by UGC, ISO 9001:2015 certified",
    "Approved by All India Council for Technical Education (AICTE), New Delhi Affiliated to Visvesvaraya Technological",
    "University (VTU), Belagavi, Karnataka Recognized by Govt. of Karnataka",
]
ROW_HEIGHTS = [36, 18, 18, 18, 8]  # rows 1-5


def apply_header(ws, last_col):
    """Write the college title block into rows 1-4 and anchor the logo at the
    left edge of the worksheet, vertically centered over the header block.
    last_col = number of data columns (int)."""
    try:
        last_letter = get_column_letter(max(last_col, 2))

        ws.merge_cells(f"B1:{last_letter}1")
        c = ws["B1"]
        c.value = COLLEGE_NAME
        c.font = Font(name="Calibri", size=18, bold=True, color=DARK_BLUE)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.row_dimensions[1].height = ROW_HEIGHTS[0]

        for i, text in enumerate(SUBLINES, start=2):
            ws.merge_cells(f"B{i}:{last_letter}{i}")
            c = ws[f"B{i}"]
            c.value = text
            c.font = Font(name="Calibri", size=9, color=GRAY)
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            ws.row_dimensions[i].height = ROW_HEIGHTS[i - 1]

        ws.row_dimensions[5].height = ROW_HEIGHTS[4]

        _blank_header_cells(ws, last_col)
        _anchor_logo(ws)
    except Exception as e:
        print("[Warn] College header skipped:", e)


def _blank_header_cells(ws, last_col):
    """Solid white fill over the header rows (1..HEADER_ROWS). Gridlines stay
    ON for the rest of the sheet, but Excel hides them behind a filled cell,
    so the header area renders as a clean white band while the result table
    below keeps its normal gridlines and borders."""
    try:
        white = PatternFill("solid", fgColor="FFFFFF")
        for r in range(1, HEADER_ROWS + 1):
            for c in range(1, max(last_col, 2) + 1):
                ws.cell(row=r, column=c).fill = white
    except Exception as e:
        print("[Warn] Header cell blanking skipped:", e)


def _anchor_logo(ws):
    """Place the college logo at absolute cell A1 using an explicit one-cell
    anchor, sized ~90px high so it visually spans the header block instead of
    being squeezed into the first cell."""
    try:
        from openpyxl.drawing.image import Image as XLImage
        from openpyxl.drawing.spreadsheet_drawing import OneCellAnchor, AnchorMarker
        from openpyxl.drawing.xdr import XDRPositiveSize2D

        if not os.path.exists(COLLEGE_LOGO):
            return

        img = XLImage(COLLEGE_LOGO)
        img.height = LOGO_HEIGHT_PX
        img.width = int(LOGO_HEIGHT_PX * _aspect(COLLEGE_LOGO))

        marker = AnchorMarker(
            col=0, colOff=0, row=0, rowOff=pixels_to_EMU(LOGO_Y_OFFSET_PX)
        )
        img.anchor = OneCellAnchor(
            _from=marker,
            ext=XDRPositiveSize2D(
                cx=pixels_to_EMU(img.width),
                cy=pixels_to_EMU(img.height),
            ),
        )
        ws.add_image(img)
    except Exception as e:
        print("[Warn] Logo embedding skipped:", e)


def _aspect(path):
    try:
        from PIL import Image as PILImage

        with PILImage.open(path) as im:
            w, h = im.size
            return (w / h) if h else 1.0
    except Exception:
        return 1.0
