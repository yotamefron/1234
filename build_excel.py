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

    # AE: env_match — TRUE if "הכל" selected or any of Q/R/S/T matches env
    ws[f"AE{r}"] = (
        f"=IF('לוח בקרה'!$A$2=\"הכל\",TRUE,"
        f"OR(Q{r}='לוח בקרה'!$A$2,R{r}='לוח בקרה'!$A$2,"
        f"S{r}='לוח בקרה'!$A$2,T{r}='לוח בקרה'!$A$2))"
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
# Build השאלות (Loans) sheet
# ===========================================================================
ws_l = wb["השאלות"]
ws_l.sheet_view.rightToLeft = True

# --- Filter area (rows 1–2) ---
filter_labels = ["מדינה", "מושאל ל", "אחראי מהחברה", "באיחור"]
for i, label in enumerate(filter_labels):
    cell = ws_l.cell(row=1, column=i + 1, value=label)
    style_header(cell)
    # Empty filter value cell below
    val_cell = ws_l.cell(row=2, column=i + 1)
    val_cell.border = THIN_BORDER
    val_cell.alignment = CELL_ALIGN

# D2: Overdue filter — Yes/No dropdown
dv = DataValidation(type="list", formula1="כן_לא", allow_blank=True)
dv.sqref = "D2"
ws_l.add_data_validation(dv)

# Row 3: empty separator (no content)

# --- Display table headers (row 4) ---
loan_headers = [
    "מזהה ייחודי",     # A → Inventory D
    "שם מוצר",          # B → Inventory B
    "גרסה",             # C → Inventory C
    "מושאל ל",          # D → Inventory Y
    "מדינה",             # E → Inventory Z
    "אחראי מהחברה",    # F → Inventory AA
    "תאריך השאלה",      # G → Inventory F
    "החזרה משוערת",     # H → Inventory AB
    "באיחור",            # I → computed
    "הערות",             # J → Inventory AC
]
for i, h in enumerate(loan_headers):
    cell = ws_l.cell(row=4, column=i + 1, value=h)
    style_header(cell)

# Column K: hidden helper (row reference)
ws_l.cell(row=4, column=11, value="row_ref")
style_header(ws_l.cell(row=4, column=11))
ws_l.column_dimensions["K"].hidden = True

# Column widths
loan_widths = {"A": 16, "B": 24, "C": 12, "D": 18, "E": 14,
               "F": 18, "G": 16, "H": 16, "I": 14, "J": 28, "K": 10}
for col, w in loan_widths.items():
    ws_l.column_dimensions[col].width = w

# Freeze at row 5 (below headers)
ws_l.freeze_panes = "A5"

# --- Formulas for 100 display rows (rows 5–104) ---
# Inventory column references (data rows 2–500)
INV = "מלאי"
INV_RANGE = "$2:$500"

# Helper column K: SMALL/IF to find nth matching row
# Conditions: status=מושאל AND optional filters from row 2
# This is an implicit array formula (works in Excel 365; CSE in older)
OVERDUE_CHECK = (
    f"({INV}!$AB$2:$AB$500<>\"\")*({INV}!$AB$2:$AB$500<TODAY())"
)
FILTER_COND = (
    f"({INV}!$X$2:$X$500=\"מושאל\")"
    f"*(IF($A$2=\"\",1,{INV}!$Z$2:$Z$500=$A$2))"
    f"*(IF($B$2=\"\",1,{INV}!$Y$2:$Y$500=$B$2))"
    f"*(IF($C$2=\"\",1,{INV}!$AA$2:$AA$500=$C$2))"
    f"*(IF($D$2=\"\",1,IF($D$2=\"כן\",{OVERDUE_CHECK},1-{OVERDUE_CHECK})))"
)
ROW_INDEX_ARRAY = f"ROW({INV}!$A$2:$A$500)-ROW({INV}!$A$2)+1"

# Map: loans display column → inventory column letter
inv_col_map = {
    "A": "D",   # מזהה ייחודי
    "B": "B",   # שם מוצר
    "C": "C",   # גרסה
    "D": "Y",   # מושאל ל
    "E": "Z",   # מדינה
    "F": "AA",  # אחראי מהחברה
    "G": "F",   # תאריך השאלה
    "H": "AB",  # החזרה משוערת
    # I = computed (overdue)
    "J": "AC",  # הערות
}

for r in range(5, 105):
    n = r - 4  # nth match

    # K: helper — row index of nth matching inventory row
    ws_l[f"K{r}"] = (
        f"=IFERROR(SMALL(IF({FILTER_COND},{ROW_INDEX_ARRAY},\"\"),{n}),\"\")"
    )

    # A–H, J: pull data via INDEX using the helper row reference
    for disp_col, inv_col in inv_col_map.items():
        ws_l[f"{disp_col}{r}"] = (
            f'=IF($K{r}="","",INDEX({INV}!${inv_col}$2:${inv_col}$500,$K{r}))'
        )

    # I: Overdue status — computed from expected return date
    ws_l[f"I{r}"] = (
        f'=IF($K{r}="","",IF(AND('
        f'INDEX({INV}!$AB$2:$AB$500,$K{r})<>"",'
        f'INDEX({INV}!$AB$2:$AB$500,$K{r})<TODAY()),'
        f'"באיחור!","תקין"))'
    )

    # Date formatting for loan date (G) and expected return (H)
    ws_l[f"G{r}"].number_format = "DD/MM/YYYY"
    ws_l[f"H{r}"].number_format = "DD/MM/YYYY"

    # Alternating row fill for all cells in the row
    for c in range(1, 12):
        cell = ws_l.cell(row=r, column=c)
        if not cell.border or cell.border == Border():
            pass
        style_data(cell, r - 5)

# --- Conditional formatting ---
RED_FILL = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
RED_FONT = Font(bold=True, color="9C0006")

# Overdue text in column I — red bold font
ws_l.conditional_formatting.add(
    "I5:I104",
    FormulaRule(formula=['I5="באיחור!"'], font=RED_FONT, fill=RED_FILL),
)

# Entire row highlighted with light red when overdue
ws_l.conditional_formatting.add(
    "A5:J104",
    FormulaRule(formula=['$I5="באיחור!"'], fill=RED_FILL),
)


# ===========================================================================
# Build לוח בקרה (Dashboard) sheet
# ===========================================================================
ws_d = wb["לוח בקרה"]
ws_d.sheet_view.rightToLeft = True

# --- Dashboard-specific styles ---
FILTER_BG = PatternFill(start_color="2E75B6", end_color="2E75B6", fill_type="solid")
FILTER_FONT = Font(bold=True, color="FFFFFF", size=11)
FILTER_VAL_FILL = PatternFill(start_color="D6E4F0", end_color="D6E4F0", fill_type="solid")
TITLE_FONT_D = Font(bold=True, size=14, color="1F3864")
METRIC_LABEL_FONT = Font(bold=True, size=11, color="1F3864")
METRIC_VAL_FONT = Font(bold=True, size=16, color="1F3864")
SUMMARY_BG = PatternFill(start_color="E9EFF7", end_color="E9EFF7", fill_type="solid")

# --- Filter Area (rows 1–4) ---
# Row 1–2: first 6 filters
filter_r1 = ["סביבה", "קו מוצר", "מוצר", "גרסה", "סוג בד", "מידה"]
for i, label in enumerate(filter_r1):
    c = i + 1
    cell = ws_d.cell(row=1, column=c, value=label)
    cell.font = FILTER_FONT
    cell.fill = FILTER_BG
    cell.alignment = HEADER_ALIGN
    cell.border = THIN_BORDER
    val = ws_d.cell(row=2, column=c, value="הכל")
    val.fill = FILTER_VAL_FILL
    val.alignment = CELL_ALIGN
    val.border = THIN_BORDER

# Row 3–4: last 5 filters
filter_r2 = ["סטטוס", "מדינה", "תקינות ויזואלית", "תקינות תרמית", "גרסה סופית"]
for i, label in enumerate(filter_r2):
    c = i + 1
    cell = ws_d.cell(row=3, column=c, value=label)
    cell.font = FILTER_FONT
    cell.fill = FILTER_BG
    cell.alignment = HEADER_ALIGN
    cell.border = THIN_BORDER
    val = ws_d.cell(row=4, column=c, value="הכל")
    val.fill = FILTER_VAL_FILL
    val.alignment = CELL_ALIGN
    val.border = THIN_BORDER

# Data validations for filter dropdowns
dv_defs = [
    ("A2", '"הכל,מדברי,מיוער,שלג,ימי,מבולדר/שטח בנוי,אחר"'),
    ("B2", '"הכל,OverGarment,Hide Site,Platform Hide Site,Uniform,'
           'Blankets,Urban,Platform On-The-Move,Accessories"'),
    ("E2", '"הכל,Sahar,Inbar,SRV,Gabardine,Meron,Arber,IRR,'
           'Polar,Nylon,Mesh,PVC,Other"'),
    ("F2", '"הכל,S,M,L,XL,XXL,One Size,N/A"'),
    ("A4", '"הכל,במלאי,מושאל,בחדר תצוגה,בתיק הדגמה"'),
    ("C4", '"הכל,כן,לא"'),
    ("D4", '"הכל,כן,לא"'),
    ("E4", '"הכל,כן,לא"'),
]
for sqref, f1 in dv_defs:
    dv = DataValidation(type="list", formula1=f1, allow_blank=True)
    dv.sqref = sqref
    ws_d.add_data_validation(dv)
# C2 (product), D2 (version), B4 (country) — free text, default "הכל"

# Column widths
for col, w in {"A": 48, "B": 22, "C": 20, "D": 18, "E": 18,
               "F": 14, "G": 18, "H": 18, "I": 14, "J": 28, "K": 10}.items():
    ws_d.column_dimensions[col].width = w

# ---------------------------------------------------------------------------
# Summary Area (rows 5–18)
# ---------------------------------------------------------------------------
# Reusable formula building blocks
DI = "מלאי"                 # inventory sheet name
DR = "$2:$500"              # data row range


def dflt(dash_cell, inv_col):
    """Dashboard filter term: pass all when 'הכל', else exact match."""
    return f'(IF({dash_cell}="הכל",1,{DI}!${inv_col}{DR}={dash_cell}))'


F_DATA = f'({DI}!$D{DR}<>"")'
F_ENV  = f'(--{DI}!$AE{DR})'

# Build full base‐filter string  (data × env × 10 user filters)
BF = (f'{F_DATA}*{F_ENV}'
      f'*{dflt("$B$2","A")}*{dflt("$C$2","B")}*{dflt("$D$2","C")}'
      f'*{dflt("$E$2","G")}*{dflt("$F$2","I")}'
      f'*{dflt("$A$4","X")}*{dflt("$B$4","Z")}'
      f'*{dflt("$C$4","O")}*{dflt("$D$4","U")}*{dflt("$E$4","K")}')

# Base filter WITHOUT env (for "wrong‐print" metric)
BF_NE = (f'{F_DATA}'
         f'*{dflt("$B$2","A")}*{dflt("$C$2","B")}*{dflt("$D$2","C")}'
         f'*{dflt("$E$2","G")}*{dflt("$F$2","I")}'
         f'*{dflt("$A$4","X")}*{dflt("$B$4","Z")}'
         f'*{dflt("$C$4","O")}*{dflt("$D$4","U")}*{dflt("$E$4","K")}')

# Section title
ws_d["A5"] = "סיכום"
ws_d["A5"].font = TITLE_FONT_D

# Metric definitions: (row, label, formula)
metrics = [
    (6,  'סה"כ פריטים תואמים',
     f'=SUMPRODUCT({BF})'),
    (7,  "כמות תקין לגמרי",
     f'=SUMPRODUCT({BF}*(--{DI}!$AF{DR}))'),
    (8,  "כמות תקין ויזואלית בלבד",
     f'=SUMPRODUCT({BF}*(--{DI}!$AG{DR}))'),
    (9,  "כמות תקין תרמית בלבד",
     f'=SUMPRODUCT({BF}*(--{DI}!$AH{DR}))'),
    (10, "כמות במלאי",
     f'=SUMPRODUCT({BF}*({DI}!$X{DR}="במלאי"))'),
    (11, "כמות מושאלים",
     f'=SUMPRODUCT({BF}*({DI}!$X{DR}="מושאל"))'),
    (12, "כמות בחדר תצוגה",
     f'=SUMPRODUCT({BF}*({DI}!$X{DR}="בחדר תצוגה"))'),
    (13, "כמות בתיק הדגמה",
     f'=SUMPRODUCT({BF}*({DI}!$X{DR}="בתיק הדגמה"))'),
    (14, "כמות עם פער בתכולה",
     f'=SUMPRODUCT({BF}*({DI}!$AD{DR}="כן"))'),
    (15, "כמות גרסה מיוחדת - פיתוח",
     f'=SUMPRODUCT({BF}*({DI}!$J{DR}="כן"))'),
    (16, "תקינים ויזואלית אך הדפס לא מתאים לסביבה",
     f'=IF($A$2="הכל","—",'
     f'SUMPRODUCT({BF_NE}*({DI}!$AE{DR}=FALSE)'
     f'*(({DI}!$O{DR}="כן")+({DI}!$P{DR}="כן")>0)))'),
]

for row, label, formula in metrics:
    ws_d.merge_cells(f"A{row}:C{row}")
    cl = ws_d[f"A{row}"]
    cl.value = label
    cl.font = METRIC_LABEL_FONT
    cl.alignment = CELL_ALIGN
    cl.fill = SUMMARY_BG
    cl.border = THIN_BORDER
    cv = ws_d[f"D{row}"]
    cv.value = formula
    cv.font = METRIC_VAL_FONT
    cv.alignment = Alignment(horizontal="center", vertical="center")
    cv.fill = SUMMARY_BG
    cv.border = THIN_BORDER

# Row 17: print detail for wrong-env metric
ws_d.merge_cells("A17:C17")
ws_d["A17"].value = "פירוט הדפסים שאינם מתאימים"
ws_d["A17"].font = METRIC_LABEL_FONT
ws_d["A17"].alignment = CELL_ALIGN
ws_d["A17"].fill = SUMMARY_BG
ws_d["A17"].border = THIN_BORDER
ws_d.merge_cells("D17:J17")
ws_d["D17"] = (
    f'=IF($A$2="הכל","—",IFERROR(TEXTJOIN(", ",TRUE,UNIQUE(FILTER('
    f'{DI}!$M{DR},({DI}!$D{DR}<>"")*({DI}!$AE{DR}=FALSE)'
    f'*(({DI}!$O{DR}="כן")+({DI}!$P{DR}="כן")>0)))),""))'
)
ws_d["D17"].font = Font(size=10, color="1F3864")
ws_d["D17"].alignment = CELL_ALIGN
ws_d["D17"].fill = SUMMARY_BG
ws_d["D17"].border = THIN_BORDER

# Row 18: empty separator

# ---------------------------------------------------------------------------
# Filtered Product Detail (rows 19–320)
# ---------------------------------------------------------------------------
ws_d["A19"] = "פירוט פריטים מסוננים"
ws_d["A19"].font = TITLE_FONT_D

detail_headers = [
    "מזהה", "שם מוצר", "גרסה", "סוג בד", "הדפס A", "הדפס B",
    "תקין ויזואלי", "תקין תרמי", "סטטוס", "הערות",
]
for i, h in enumerate(detail_headers):
    cell = ws_d.cell(row=20, column=i + 1, value=h)
    style_header(cell)

# Hidden helper column K
style_header(ws_d.cell(row=20, column=11, value="row_ref"))
ws_d.column_dimensions["K"].hidden = True

# Freeze below detail headers
ws_d.freeze_panes = "A21"

# SMALL/IF row-index formula shared components
D_ROW_IDX = f"ROW({DI}!$A$2:$A$500)-ROW({DI}!$A$2)+1"

# Detail column → inventory column mapping
det_map = {
    "A": "D", "B": "B", "C": "C", "D": "G", "E": "M",
    "F": "N", "G": "O", "H": "U", "I": "X", "J": "AC",
}

for r in range(21, 321):
    n = r - 20  # nth match

    # K: helper — row index of nth matching inventory row
    ws_d[f"K{r}"] = f'=IFERROR(SMALL(IF({BF},{D_ROW_IDX},""),{n}),"")'

    # A–J: pull data via INDEX
    for dcol, icol in det_map.items():
        ws_d[f"{dcol}{r}"] = (
            f'=IF($K{r}="","",INDEX({DI}!${icol}{DR},$K{r}))'
        )

    # Style all cells in the row
    for c in range(1, 12):
        style_data(ws_d.cell(row=r, column=c), r - 21)

# ---------------------------------------------------------------------------
# Active Loans Detail (rows 322–423, always displayed)
# ---------------------------------------------------------------------------
LOAN_T = 322        # title row
LOAN_H = 323        # header row
LOAN_S = 324        # first data row
LOAN_E = 423        # last data row

ws_d[f"A{LOAN_T}"] = "השאלות פעילות"
ws_d[f"A{LOAN_T}"].font = TITLE_FONT_D

loan_d_headers = ["מזהה", "מוצר", "מושאל ל", "מדינה", "החזרה משוערת", "באיחור"]
for i, h in enumerate(loan_d_headers):
    cell = ws_d.cell(row=LOAN_H, column=i + 1, value=h)
    style_header(cell)

# Reuse column K as helper for the loans row range
style_header(ws_d.cell(row=LOAN_H, column=11, value="loan_ref"))

# Loan filter: status = מושאל (always, no dashboard filter)
LOAN_COND = f'({DI}!$D{DR}<>"")*({DI}!$X{DR}="מושאל")'

loan_d_map = {"A": "D", "B": "B", "C": "Y", "D": "Z", "E": "AB"}

for r in range(LOAN_S, LOAN_E + 1):
    n = r - LOAN_S + 1

    # K: helper
    ws_d[f"K{r}"] = f'=IFERROR(SMALL(IF({LOAN_COND},{D_ROW_IDX},""),{n}),"")'

    # A–E: pull data
    for dcol, icol in loan_d_map.items():
        ws_d[f"{dcol}{r}"] = (
            f'=IF($K{r}="","",INDEX({DI}!${icol}{DR},$K{r}))'
        )

    # F: overdue indicator
    ws_d[f"F{r}"] = (
        f'=IF($K{r}="","",IF(AND('
        f'INDEX({DI}!$AB{DR},$K{r})<>"",'
        f'INDEX({DI}!$AB{DR},$K{r})<TODAY()),'
        f'"באיחור!","תקין"))'
    )

    # Date formatting for expected return (E)
    ws_d[f"E{r}"].number_format = "DD/MM/YYYY"

    # Style
    for c in range(1, 12):
        style_data(ws_d.cell(row=r, column=c), r - LOAN_S)

# Conditional formatting — overdue loans red highlight
RED_FILL_D = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
RED_FONT_D = Font(bold=True, color="9C0006")
ws_d.conditional_formatting.add(
    f"F{LOAN_S}:F{LOAN_E}",
    FormulaRule(formula=[f'F{LOAN_S}="באיחור!"'], font=RED_FONT_D, fill=RED_FILL_D),
)
ws_d.conditional_formatting.add(
    f"A{LOAN_S}:F{LOAN_E}",
    FormulaRule(formula=[f'$F{LOAN_S}="באיחור!"'], fill=RED_FILL_D),
)


# ===========================================================================
# Save
# ===========================================================================
wb.save("ametrine_inventory.xlsx")
