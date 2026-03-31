from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule
import datetime

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

wb.active.title = sheet_names[0]
for name in sheet_names[1:]:
    wb.create_sheet(title=name)

# ---------------------------------------------------------------------------
# Styling
# ---------------------------------------------------------------------------
HEADER_FONT = Font(bold=True, color="FFFFFF", size=11)
HEADER_FILL = PatternFill(start_color="1F3864", end_color="1F3864", fill_type="solid")
HEADER_ALIGN = Alignment(horizontal="right", vertical="center", wrap_text=True)
ROW_EVEN_FILL = PatternFill(start_color="D6E4F0", end_color="D6E4F0", fill_type="solid")
ROW_ODD_FILL = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
CELL_ALIGN = Alignment(horizontal="right", vertical="center")
THIN_BORDER = Border(
    left=Side(style="thin", color="B0B0B0"),
    right=Side(style="thin", color="B0B0B0"),
    top=Side(style="thin", color="B0B0B0"),
    bottom=Side(style="thin", color="B0B0B0"),
)
YELLOW_FILL = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")


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
    """Write a styled table and define a named range. Returns (next_row, data_start, data_end)."""
    num_cols = len(headers)
    for c, header in enumerate(headers):
        cell = ws.cell(row=start_row, column=start_col + c, value=header)
        style_header(cell)
    for r, row_data in enumerate(data_rows):
        for c, value in enumerate(row_data):
            cell = ws.cell(row=start_row + 1 + r, column=start_col + c, value=value)
            style_data(cell, r)
    for c in range(num_cols):
        col_letter = get_column_letter(start_col + c)
        max_len = len(str(headers[c]))
        for row_data in data_rows:
            if c < len(row_data) and row_data[c] is not None:
                max_len = max(max_len, len(str(row_data[c])))
        ws.column_dimensions[col_letter].width = max(max_len * 1.5, 14)

    data_start = start_row + 1
    data_end = start_row + max(len(data_rows), 1)
    first_col = get_column_letter(start_col)
    last_col = get_column_letter(start_col + num_cols - 1)
    ref = f"'הגדרות'!${first_col}${data_start}:${last_col}${data_end}"
    dn = DefinedName(range_name, attr_text=ref)
    wb.defined_names.add(dn)

    next_row = start_row + 1 + len(data_rows) + 2
    return next_row, data_start, data_end


def add_named(name, ref):
    wb.defined_names.add(DefinedName(name, attr_text=ref))


# ===========================================================================
# Build הגדרות (Settings) sheet
# ===========================================================================
ws_s = wb["הגדרות"]
ws_s.sheet_view.rightToLeft = True
cr = 1

# TABLE 1 — Environments
cr, _, _ = write_table(ws_s, cr, 1, ["סביבה"],
    [["מדברי"], ["מיוער"], ["שלג"], ["ימי"], ["מבולדר/שטח בנוי"], ["אחר"]],
    "סביבות")

# TABLE 2 — Prints
prints_data = [
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
cr, pr_start, pr_end = write_table(ws_s, cr, 1,
    ["הדפס", "סביבה ראשית", "סביבה משנית", "תמונה"], prints_data, "הדפסים")

# Single-column named range for print name dropdowns
add_named("שמות_הדפסים", f"'הגדרות'!$A${pr_start}:$A${pr_end}")

# TABLE 3 — Product Lines
cr, _, _ = write_table(ws_s, cr, 1, ["קו מוצר"],
    [["OverGarment"], ["Hide Site"], ["Platform Hide Site"], ["Uniform"],
     ["Blankets"], ["Urban"], ["Platform On-The-Move"], ["Accessories"]],
    "קווי_מוצר")

# TABLE 4 — Products
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
cr, prod_start, prod_end = write_table(ws_s, cr, 1,
    ["קו מוצר", "שם מוצר", "גרסאות זמינות"], t4_data, "מוצרים")

# Per-product-line named ranges for dependent dropdown (INDIRECT)
def sanitize(name):
    return name.replace(" ", "_").replace("-", "_").replace("/", "_")

cur_pl, pl_first = None, None
for i, row in enumerate(t4_data):
    rnum = prod_start + i
    if row[0] != cur_pl:
        if cur_pl is not None:
            add_named(sanitize(cur_pl), f"'הגדרות'!$B${pl_first}:$B${rnum - 1}")
        cur_pl = row[0]
        pl_first = rnum
add_named(sanitize(cur_pl), f"'הגדרות'!$B${pl_first}:$B${prod_start + len(t4_data) - 1}")

# TABLE 5 — Fabric Types
cr, _, _ = write_table(ws_s, cr, 1, ["סוג בד"],
    [["Sahar"], ["Inbar"], ["SRV"], ["Gabardine"], ["Meron"], ["Arber"],
     ["IRR"], ["Polar"], ["Nylon"], ["Mesh"], ["PVC"], ["Other"]],
    "סוגי_בד")

# TABLE 6 — Statuses
cr, _, _ = write_table(ws_s, cr, 1, ["סטטוס"],
    [["במלאי"], ["מושאל"], ["בחדר תצוגה"], ["בתיק הדגמה"]], "סטטוסים")

# TABLE 7 — Yes/No
cr, _, _ = write_table(ws_s, cr, 1, ["ערך"], [["כן"], ["לא"]], "כן_לא")

# TABLE 8 — Sizes
cr, _, _ = write_table(ws_s, cr, 1, ["מידה"],
    [["S"], ["M"], ["L"], ["XL"], ["XXL"], ["One Size"], ["N/A"]], "מידות")

# TABLE 9 — SKU mapping
cr, sku_start, sku_end = write_table(ws_s, cr, 1,
    ["מוצר", "סוג בד", "גרסה", 'מק"ט'],
    [["", "", "", ""] for _ in range(20)], "מקטים")


# ===========================================================================
# Build מלאי (Inventory) sheet
# ===========================================================================
ws = wb["מלאי"]
ws.sheet_view.rightToLeft = True

# --- Column headers (row 1) ---
inv_headers = [
    "קו מוצר",                   # A
    "שם מוצר",                    # B
    "גרסה",                       # C
    "מזהה ייחודי",                # D
    'מק"ט',                       # E
    "תאריך השאלה",                # F
    "סוג בד 1",                   # G
    "סוג בד 2",                   # H
    "מידה",                       # I
    "גרסה מיוחדת - פיתוח",       # J
    "גרסה סופית",                 # K
    "הערות גרסה/פיצ׳רים",        # L
    "הדפס צד A",                  # M
    "הדפס צד B",                  # N
    "תקין ויזואלית צד A",         # O
    "תקין ויזואלית צד B",         # P
    "סביבה A",                    # Q  (formula)
    "סביבה משנית A",              # R  (formula)
    "סביבה B",                    # S  (formula)
    "סביבה משנית B",              # T  (formula)
    "תקין תרמית צד A",            # U
    "תקין תרמית צד B",            # V
    "פיצ׳רים תקינים",             # W
    "סטטוס",                      # X
    "מושאל ל",                    # Y
    "מדינה",                      # Z
    "אחראי מהחברה",              # AA
    "החזרה משוערת",               # AB
    "הערות",                      # AC
    "פער בתכולה",                 # AD
    "מתאים_סביבה",               # AE (helper)
    "תקין_לגמרי",                # AF (helper)
    "תקין_ויזואלי_בלבד",         # AG (helper)
    "תקין_תרמי_בלבד",            # AH (helper)
    "באיחור",                     # AI (helper)
]

for i, h in enumerate(inv_headers):
    cell = ws.cell(row=1, column=i + 1, value=h)
    style_header(cell)

# Column widths
widths = [16, 26, 12, 16, 12, 16, 14, 14, 10, 24, 14, 26,
          18, 18, 22, 22, 14, 18, 14, 18, 22, 22, 18, 12,
          16, 12, 18, 18, 28, 16, 16, 16, 22, 22, 14]
for i, w in enumerate(widths):
    ws.column_dimensions[get_column_letter(i + 1)].width = w

# Freeze header row
ws.freeze_panes = "A2"

# Hide helper columns AE–AI (columns 31–35)
for c in range(31, 36):
    ws.column_dimensions[get_column_letter(c)].hidden = True

# --- Data Validations (rows 2–500) ---
MR = 500

# A: Product Line
dv = DataValidation(type="list", formula1="קווי_מוצר", allow_blank=True)
dv.sqref = f"A2:A{MR}"
ws.add_data_validation(dv)

# B: Product Name — dependent dropdown via INDIRECT
dv = DataValidation(
    type="list",
    formula1='INDIRECT(SUBSTITUTE(SUBSTITUTE(A2," ","_"),"-","_"))',
    allow_blank=True,
)
dv.sqref = f"B2:B{MR}"
ws.add_data_validation(dv)

# G–H: Fabric types
dv = DataValidation(type="list", formula1="סוגי_בד", allow_blank=True)
dv.sqref = f"G2:G{MR} H2:H{MR}"
ws.add_data_validation(dv)

# I: Sizes
dv = DataValidation(type="list", formula1="מידות", allow_blank=True)
dv.sqref = f"I2:I{MR}"
ws.add_data_validation(dv)

# J, K, O, P, U, V, W, AD: Yes/No
dv = DataValidation(type="list", formula1="כן_לא", allow_blank=True)
dv.sqref = (f"J2:J{MR} K2:K{MR} O2:O{MR} P2:P{MR} "
            f"U2:U{MR} V2:V{MR} W2:W{MR} AD2:AD{MR}")
ws.add_data_validation(dv)

# M–N: Prints
dv = DataValidation(type="list", formula1="שמות_הדפסים", allow_blank=True)
dv.sqref = f"M2:M{MR} N2:N{MR}"
ws.add_data_validation(dv)

# X: Status
dv = DataValidation(type="list", formula1="סטטוסים", allow_blank=True)
dv.sqref = f"X2:X{MR}"
ws.add_data_validation(dv)

# F, AB: Date format
for col_letter in ("F", "AB"):
    for r in range(2, MR + 1):
        ws[f"{col_letter}{r}"].number_format = "DD/MM/YYYY"

# --- Formulas (written per data row; sample rows 2–6, expandable) ---
PRINTS_RANGE = f"'הגדרות'!$A${pr_start}:$D${pr_end}"
SKU_A = f"'הגדרות'!$A${sku_start}:$A${sku_end}"
SKU_B = f"'הגדרות'!$B${sku_start}:$B${sku_end}"
SKU_C = f"'הגדרות'!$C${sku_start}:$C${sku_end}"
SKU_D = f"'הגדרות'!$D${sku_start}:$D${sku_end}"

FORMULA_ROWS = range(2, MR + 1)

for r in FORMULA_ROWS:
    # E: SKU — INDEX/MATCH on concatenated product+fabric+version
    ws[f"E{r}"] = (
        f'=IFERROR(INDEX({SKU_D},MATCH(B{r}&G{r}&C{r},'
        f'{SKU_A}&{SKU_B}&{SKU_C},0)),"")'
    )

    # Q: Primary environment for Side-A print
    ws[f"Q{r}"] = f'=IFERROR(VLOOKUP(M{r},{PRINTS_RANGE},2,FALSE),"")'
    # R: Secondary environment for Side-A print
    ws[f"R{r}"] = f'=IFERROR(VLOOKUP(M{r},{PRINTS_RANGE},3,FALSE),"")'
    # S: Primary environment for Side-B print
    ws[f"S{r}"] = f'=IFERROR(VLOOKUP(N{r},{PRINTS_RANGE},2,FALSE),"")'
    # T: Secondary environment for Side-B print
    ws[f"T{r}"] = f'=IFERROR(VLOOKUP(N{r},{PRINTS_RANGE},3,FALSE),"")'

    # AE: env_match — TRUE if any of Q/R/S/T matches the dashboard filter cell
    ws[f"AE{r}"] = (
        f"=OR(Q{r}='לוח בקרה'!$B$2,R{r}='לוח בקרה'!$B$2,"
        f"S{r}='לוח בקרה'!$B$2,T{r}='לוח בקרה'!$B$2)"
    )

    # AF: fully_ok — env matches AND visually+thermally OK on at least one side
    ws[f"AF{r}"] = (
        f'=AND(AE{r},OR('
        f'AND(O{r}="כן",U{r}="כן"),'
        f'AND(P{r}="כן",V{r}="כן")))'
    )

    # AG: visual_only_ok — env matches AND visually OK but NOT thermally OK
    ws[f"AG{r}"] = (
        f'=AND(AE{r},'
        f'OR(O{r}="כן",P{r}="כן"),'
        f'NOT(OR(AND(O{r}="כן",U{r}="כן"),AND(P{r}="כן",V{r}="כן"))))'
    )

    # AH: thermal_only_ok — env matches AND thermally OK but NOT visually OK
    ws[f"AH{r}"] = (
        f'=AND(AE{r},'
        f'OR(U{r}="כן",V{r}="כן"),'
        f'NOT(OR(O{r}="כן",P{r}="כן")))'
    )

    # AI: overdue — loaned AND past expected return
    ws[f"AI{r}"] = (
        f'=AND(X{r}="מושאל",AB{r}<>"",AB{r}<TODAY())'
    )

# --- Conditional formatting: yellow Notes (AC) when גרסה סופית (K) = לא ---
ws.conditional_formatting.add(
    f"AC2:AC{MR}",
    FormulaRule(
        formula=[f'K2="לא"'],
        fill=YELLOW_FILL,
    ),
)

# --- 5 Sample data rows ---
samples = [
    {  # Row 2 — Poncho Inbar with Desert print → special note
        "A": "OverGarment", "B": "Poncho Inbar", "C": "",
        "D": "INV-001", "F": datetime.date(2025, 11, 10),
        "G": "Inbar", "H": "", "I": "L",
        "J": "לא", "K": "כן", "L": "",
        "M": "Desert", "N": "Forest Green",
        "O": "לא", "P": "כן",
        "U": "כן", "V": "כן", "W": "כן",
        "X": "במלאי", "Y": "", "Z": "ישראל",
        "AA": "", "AC": "הדפס Desert ישן — אינו תקין ויזואלית",
        "AD": "לא",
    },
    {  # Row 3 — Hide Site, loaned, overdue
        "A": "Hide Site", "B": "Hide Site Standard", "C": "",
        "D": "INV-002", "F": datetime.date(2025, 8, 1),
        "G": "Sahar", "H": "Gabardine", "I": "One Size",
        "J": "לא", "K": "כן", "L": "",
        "M": "Woodland", "N": "",
        "O": "כן", "P": "",
        "U": "כן", "V": "", "W": "כן",
        "X": "מושאל", "Y": "צה\"ל", "Z": "ישראל",
        "AA": "יוסי כהן", "AB": datetime.date(2026, 2, 15),
        "AC": "", "AD": "לא",
    },
    {  # Row 4 — Uniform, in showroom, final version = לא → yellow note
        "A": "Uniform", "B": "Shirt", "C": "",
        "D": "INV-003", "F": "",
        "G": "Meron", "H": "", "I": "M",
        "J": "לא", "K": "לא", "L": "דורש בדיקה נוספת",
        "M": "Urban Green", "N": "Urban",
        "O": "כן", "P": "כן",
        "U": "כן", "V": "כן", "W": "כן",
        "X": "בחדר תצוגה", "Y": "", "Z": "ישראל",
        "AA": "", "AC": "גרסת ביניים — ממתין לאישור סופי",
        "AD": "לא",
    },
    {  # Row 5 — Blankets, in stock, dev version, final = לא → yellow note
        "A": "Blankets", "B": "Survival Blanket", "C": "",
        "D": "INV-004", "F": "",
        "G": "IRR", "H": "", "I": "One Size",
        "J": "כן", "K": "לא", "L": "פרוטוטייפ",
        "M": "Fresh Snow", "N": "",
        "O": "כן", "P": "",
        "U": "לא", "V": "", "W": "כן",
        "X": "במלאי", "Y": "", "Z": "ישראל",
        "AA": "", "AC": "דגם ניסיוני — לא לשלוח ללקוח",
        "AD": "לא",
    },
    {  # Row 6 — OverGarment, loaned abroad, overdue
        "A": "OverGarment", "B": "Hobey Elite", "C": "",
        "D": "INV-005", "F": datetime.date(2025, 12, 1),
        "G": "SRV", "H": "Arber", "I": "XL",
        "J": "לא", "K": "כן", "L": "",
        "M": "Sand", "N": "Ever-Green",
        "O": "כן", "P": "כן",
        "U": "כן", "V": "כן", "W": "כן",
        "X": "מושאל", "Y": "לוקהיד מרטין", "Z": "ארה\"ב",
        "AA": "דני לוי", "AB": datetime.date(2026, 1, 10),
        "AC": "", "AD": "לא",
    },
]

# Column letter → column index mapping
COL = {}
for i in range(1, 36):
    COL[get_column_letter(i)] = i

for idx, sample in enumerate(samples):
    r = idx + 2
    for col_letter, value in sample.items():
        if value == "":
            continue
        cell = ws.cell(row=r, column=COL[col_letter], value=value)
        style_data(cell, idx)
        if col_letter in ("F", "AB") and isinstance(value, datetime.date):
            cell.number_format = "DD/MM/YYYY"

    # Style remaining cells in the row (formula cells, empty cells)
    for c in range(1, 36):
        cell = ws.cell(row=r, column=c)
        if cell.fill == PatternFill():  # unstyled
            style_data(cell, idx)

# ===========================================================================
# Save
# ===========================================================================
wb.save("ametrine_inventory.xlsx")
