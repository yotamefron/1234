from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName

wb = Workbook()

sheet_names = [
    "לוח בקרה",
    "מלאי",
    "השאלות",
    "תכולה",
    "לוח בקרה טכנו-מבצעי",
    "ערכים טכנולוגיים",
    "הגדרות",
    "מלאי רצוי",
]

# Rename the default sheet to the first name
wb.active.title = sheet_names[0]

# Create the remaining sheets
for name in sheet_names[1:]:
    wb.create_sheet(title=name)

# ---------------------------------------------------------------------------
# Styling helpers
# ---------------------------------------------------------------------------
HEADER_FONT = Font(bold=True, color="FFFFFF", size=11)
HEADER_FILL = PatternFill(start_color="1F3864", end_color="1F3864", fill_type="solid")
HEADER_ALIGN = Alignment(horizontal="right", vertical="center")
ROW_EVEN_FILL = PatternFill(start_color="D6E4F0", end_color="D6E4F0", fill_type="solid")
ROW_ODD_FILL = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
CELL_ALIGN = Alignment(horizontal="right", vertical="center")
THIN_BORDER = Border(
    left=Side(style="thin", color="B0B0B0"),
    right=Side(style="thin", color="B0B0B0"),
    top=Side(style="thin", color="B0B0B0"),
    bottom=Side(style="thin", color="B0B0B0"),
)


def style_header(cell):
    cell.font = HEADER_FONT
    cell.fill = HEADER_FILL
    cell.alignment = HEADER_ALIGN
    cell.border = THIN_BORDER


def style_data(cell, row_idx):
    cell.fill = ROW_EVEN_FILL if row_idx % 2 == 0 else ROW_ODD_FILL
    cell.alignment = CELL_ALIGN
    cell.border = THIN_BORDER


def write_table(ws, start_row, start_col, headers, data_rows, range_name):
    """Write a table with styled headers and data rows, and define a named range."""
    num_cols = len(headers)

    # Write headers
    for c, header in enumerate(headers):
        cell = ws.cell(row=start_row, column=start_col + c, value=header)
        style_header(cell)

    # Write data rows
    for r, row_data in enumerate(data_rows):
        for c, value in enumerate(row_data):
            cell = ws.cell(row=start_row + 1 + r, column=start_col + c, value=value)
            style_data(cell, r)

    # Auto-fit column widths (approximate)
    for c in range(num_cols):
        col_letter = get_column_letter(start_col + c)
        max_len = len(str(headers[c]))
        for row_data in data_rows:
            if c < len(row_data) and row_data[c] is not None:
                max_len = max(max_len, len(str(row_data[c])))
        ws.column_dimensions[col_letter].width = max(max_len * 1.5, 14)

    # Define named range (data only, excluding header)
    first_col = get_column_letter(start_col)
    last_col = get_column_letter(start_col + num_cols - 1)
    data_start = start_row + 1
    data_end = start_row + len(data_rows)
    if len(data_rows) == 0:
        data_end = data_start  # at least one row for empty tables
    ref = f"'הגדרות'!${first_col}${data_start}:${last_col}${data_end}"
    dn = DefinedName(range_name, attr_text=ref)
    wb.defined_names.add(dn)

    # Return the next free row (with 2-row gap)
    return start_row + 1 + len(data_rows) + 2


# ---------------------------------------------------------------------------
# Build the הגדרות (Settings) sheet
# ---------------------------------------------------------------------------
ws = wb["הגדרות"]
ws.sheet_view.rightToLeft = True

current_row = 1

# ===== TABLE 1 — Environments (סביבות) =====
t1_headers = ["סביבה"]
t1_data = [
    ["מדברי"],
    ["מיוער"],
    ["שלג"],
    ["ימי"],
    ["מבולדר/שטח בנוי"],
    ["אחר"],
]
current_row = write_table(ws, current_row, 1, t1_headers, t1_data, "סביבות")

# ===== TABLE 2 — Prints (הדפסים) =====
t2_headers = ["הדפס", "סביבה ראשית", "סביבה משנית", "תמונה"]
t2_data = [
    ["Desert", "מדברי", None, None],
    ["Desert-Green", "מדברי", None, None],
    ["Desert Coyote", "מדברי", None, None],
    ["Sand", "מדברי", None, None],
    ["Ever-Green", "מיוער", None, None],
    ["Forest Green", "מיוער", None, None],
    ["Kestrel", "מיוער", None, None],
    ["Woodland", "מיוער", None, None],
    ["Woodland Brown", "מיוער", None, None],
    ["Urban Green", "מיוער", "מבולדר/שטח בנוי", None],
    ["Snirt Snow", "שלג", None, None],
    ["Fresh Snow", "שלג", None, None],
    ["Marine", "ימי", None, None],
    ["Urban", "מבולדר/שטח בנוי", None, None],
    ["Boulder", "מבולדר/שטח בנוי", None, None],
]
current_row = write_table(ws, current_row, 1, t2_headers, t2_data, "הדפסים")

# ===== TABLE 3 — Product Lines (קווי מוצר) =====
t3_headers = ["קו מוצר"]
t3_data = [
    ["OverGarment"],
    ["Hide Site"],
    ["Platform Hide Site"],
    ["Uniform"],
    ["Blankets"],
    ["Urban"],
    ["Platform On-The-Move"],
    ["Accessories"],
]
current_row = write_table(ws, current_row, 1, t3_headers, t3_data, "קווי_מוצר")

# ===== TABLE 4 — Product-to-Product Line mapping (מוצרים) =====
t4_headers = ["קו מוצר", "שם מוצר", "גרסאות זמינות"]
t4_data = [
    ["OverGarment", "Hobey Elite", ""],
    ["OverGarment", "Gal Suit", ""],
    ["OverGarment", "Sniper Assault", ""],
    ["OverGarment", "SWAT 3-Part Suit", ""],
    ["OverGarment", "May Suit", ""],
    ["OverGarment", "May Long Shirt", ""],
    ["OverGarment", "Arctic Suit", ""],
    ["OverGarment", "Poncho SRV", ""],
    ["OverGarment", "Poncho Inbar", ""],
    ["OverGarment", "Poncho Inbar 3D", ""],
    ["OverGarment", "Poncho Sahar", ""],
    ["OverGarment", "Poncho Sahar 3D", ""],
    ["Hide Site", "Hide Site Standard", ""],
    ["Hide Site", "HS Floor", ""],
    ["Hide Site", "Window Kit", ""],
    ["Platform Hide Site", "Platform HS (including parts)", ""],
    ["Uniform", "Shirt", ""],
    ["Uniform", "Pants", ""],
    ["Blankets", "Survival Blanket", ""],
    ["Blankets", "יריעת החשכה", ""],
    ["Urban", "Loki", ""],
    ["Platform On-The-Move", "GMV", ""],
    ["Platform On-The-Move", "רשת הצללה", ""],
    ["Accessories", "Accessories", ""],
]
current_row = write_table(ws, current_row, 1, t4_headers, t4_data, "מוצרים")

# ===== TABLE 5 — Fabric Types (סוגי בד) =====
t5_headers = ["סוג בד"]
t5_data = [
    ["Sahar"],
    ["Inbar"],
    ["SRV"],
    ["Gabardine"],
    ["Meron"],
    ["Arber"],
    ["IRR"],
    ["Polar"],
    ["Nylon"],
    ["Mesh"],
    ["PVC"],
    ["Other"],
]
current_row = write_table(ws, current_row, 1, t5_headers, t5_data, "סוגי_בד")

# ===== TABLE 6 — Statuses (סטטוסים) =====
t6_headers = ["סטטוס"]
t6_data = [
    ["במלאי"],
    ["מושאל"],
    ["בחדר תצוגה"],
    ["בתיק הדגמה"],
]
current_row = write_table(ws, current_row, 1, t6_headers, t6_data, "סטטוסים")

# ===== TABLE 7 — Yes/No (כן/לא) =====
t7_headers = ["ערך"]
t7_data = [
    ["כן"],
    ["לא"],
]
current_row = write_table(ws, current_row, 1, t7_headers, t7_data, "כן_לא")

# ===== TABLE 8 — Sizes (מידות) =====
t8_headers = ["מידה"]
t8_data = [
    ["S"],
    ["M"],
    ["L"],
    ["XL"],
    ["XXL"],
    ["One Size"],
    ["N/A"],
]
current_row = write_table(ws, current_row, 1, t8_headers, t8_data, "מידות")

# ===== TABLE 9 — Product/Fabric/Version to SKU mapping (מק"טים) =====
t9_headers = ["מוצר", "סוג בד", "גרסה", 'מק"ט']
t9_data = [["", "", "", ""] for _ in range(20)]
current_row = write_table(ws, current_row, 1, t9_headers, t9_data, "מקטים")

# ---------------------------------------------------------------------------
# Save
# ---------------------------------------------------------------------------
wb.save("ametrine_inventory.xlsx")
