from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, Protection
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
    "רצוי מול מצוי",
]

wb.active.title = sheet_names[0]
for name in sheet_names[1:]:
    wb.create_sheet(title=name)

# ---------------------------------------------------------------------------
# Styling
# ---------------------------------------------------------------------------
HEADER_FONT = Font(name="Calibri", bold=True, color="FFFFFF", size=11)
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


DATA_FONT = Font(name="Calibri", size=11)
LOCKED = Protection(locked=True)
UNLOCKED = Protection(locked=False)


def style_header(cell):
    cell.font = HEADER_FONT
    cell.fill = HEADER_FILL
    cell.alignment = HEADER_ALIGN
    cell.border = THIN_BORDER
    cell.protection = LOCKED


def style_data(cell, row_idx):
    cell.font = DATA_FONT
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


MR = 500  # max data row for formulas


def R(col):
    """Absolute column range for inventory/tech formulas: $COL$2:$COL$MR."""
    return f"${col}$2:${col}${MR}"


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
    ["NATO", "מיוער", None, None],
    ["Kestrel (אמריקאי)", "מיוער", None, None],
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
    ["OverGarment", "Arctic Jacket", ""],
    ["OverGarment", "Arctic Overpants", ""],
    ["OverGarment", "Arctic Poncho SRV", ""],
    ["OverGarment", "Gaiters", ""],
    ["OverGarment", "Poncho SRV Short", ""],
    ["Hide Site", "Hide Site Standard", ""],
    ["Hide Site", "HS Floor", ""],
    ["Hide Site", "Window Kit", ""],
    ["Hide Site", "Operators HS", ""],
    ["Hide Site", "OverWatch (ATGM) HS", ""],
    ["Hide Site", "6X6 Hide Site", ""],
    ["Hide Site", "עמדת חוד", ""],
    ["Platform Hide Site", "Platform HS (including parts)", ""],
    ["Platform Hide Site", "LC79", ""],
    ["Platform Hide Site", "GMV", ""],
    ["Platform Hide Site", "רשת הצללה", ""],
    ["Uniform", "Shirt", ""],
    ["Uniform", "Pants", ""],
    ["Uniform", "Headcover", ""],
    ["Uniform", "Balaclava", ""],
    ["Uniform", "Boonie hat", ""],
    ["Uniform", "Vest cover", ""],
    ["Blankets", "Survival Blanket", ""],
    ["Blankets", "יריעת החשכה", ""],
    ["Blankets", "Water Proof Survival Blanket", ""],
    ["Blankets", "3D Survival Blanket", ""],
    ["Blankets", "Arctic Survival Blanket", ""],
    ["Urban", "Loki", ""],
    ["Platform On-The-Move", "GMV", ""],
    ["Platform On-The-Move", "רשת הצללה", ""],
    ["Accessories", "Accessories", ""],
    ["Accessories", "HS Floor", ""],
    ["Accessories", "חלקים של עמדות", ""],
    ["Accessories", "Blackout hood", ""],
    ["Accessories", "Weapon Wraps", ""],
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
     ["IRR"], ["Polar"], ["Nylon"], ["Mesh"], ["PVC"], ["Other"],
     ["RIPSTOP"], ["RIPSTOP Double Layer"], ["Beti"], ["MRG"], ["3D"],
     ["WaterProof Blackout Fabric"], ["רשת רכב"], ["Stretch"],
     ["Durable"], ["SMT"], ["LOKI Material"], ["אחר"]],
    "סוגי_בד")

# TABLE 6 — Statuses
cr, _, _ = write_table(ws_s, cr, 1, ["סטטוס"],
    [["במלאי"], ["מושאל"], ["בחדר תצוגה"], ["בתיק הדגמה"]], "סטטוסים")

# TABLE 7 — Yes/No
cr, _, _ = write_table(ws_s, cr, 1, ["ערך"], [["כן"], ["לא"]], "כן_לא")

# TABLE 8 — Sizes
cr, _, _ = write_table(ws_s, cr, 1, ["מידה"],
    [["XS"], ["S"], ["M"], ["L"], ["XL"], ["XXL"], ["XXXL"], ["One Size"], ["N/A"]], "מידות")

# TABLE 9 — SKU mapping
sku_data = [
    ["Poncho Sahar", "Mesh", "V1", "7031-01"],
    ["Poncho Sahar", "Mesh", "V2", "7031-01"],
    ["Poncho Inbar", "RIPSTOP", "V1", "7030-01"],
    ["Poncho SRV", "RIPSTOP", "V1", "7018-00"],
    ["Poncho SRV", "RIPSTOP", "Short", "7018-00"],
    ["Poncho SRV", "RIPSTOP Double Layer", "Short V2", "7018-00"],
    ["Hobey Elite", "Mesh", "V1", "7021-00"],
    ["Hobey Elite", "RIPSTOP", "V1", "7021-00"],
    ["Gal Suit", "Mesh", "V1", "7004-01"],
    ["Sniper Assault", "3D", "V1", "7033-00"],
    ["SWAT 3-Part SUIT", "Mesh", "V1", "9097-00"],
    ["May Suit", "RIPSTOP Double Layer", "V1", "7040-01"],
    ["Arctic Jacket", "RIPSTOP Double Layer", "V1", "7051-00"],
    ["Arctic Overpants", "RIPSTOP Double Layer", "V1", "7052-00"],
    ["Gaiters", "Mesh", "V1", "5004-00"],
    ["Gaiters", "RIPSTOP Double Layer", "V1", "5004-00"],
    ["Headcover", "Stretch", "", "8030-00"],
    ["Headcover", "Meron", "", "8030-00"],
    ["Balaclava", "Stretch", "", "8006-00"],
    ["Operators HS", "רשת רכב", "V1", "9001-00"],
    ["Operators HS", "MRG", "V1", "9001-00"],
    ["Operators HS", "RIPSTOP Double Layer", "V1", "9001-00"],
    ["Operators HS", "Mesh", "V1", "9001-00"],
    ["OverWatch (ATGM) HS", "רשת רכב", "V1", "9004-01"],
    ["OverWatch (ATGM) HS", "Beti", "V1", "9004-01"],
    ["6X6 Hide Site", "MRG", "", "9095-00"],
    ["6X6 Hide Site", "Beti", "", "9095-00"],
    ["LC79", "רשת רכב", "V1", "9026-00"],
    ["Survival Blanket", "Mesh", "V1", "9021-00"],
    ["Survival Blanket", "RIPSTOP", "V1", "9021-00"],
    ["Water Proof Survival Blanket", "RIPSTOP", "V1", "9028-00"],
    ["3D Survival Blanket", "Mesh", "V1", "9029-00"],
]
cr, sku_start, sku_end = write_table(ws_s, cr, 1,
    ["מוצר", "סוג בד", "גרסה", 'מק"ט'], sku_data, "מקטים")


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

# --- Formulas (written per data row) ---
# Padded ranges for future entries
PRINTS_PAD_END = pr_end + 5   # room for ~5 more prints
SKU_PAD_END = sku_end + 50    # room for ~50 more SKUs
PRINTS_RANGE = f"'הגדרות'!$A${pr_start}:$D${PRINTS_PAD_END}"
SKU_A = f"'הגדרות'!$A${sku_start}:$A${SKU_PAD_END}"
SKU_B = f"'הגדרות'!$B${sku_start}:$B${SKU_PAD_END}"
SKU_C = f"'הגדרות'!$C${sku_start}:$C${SKU_PAD_END}"
SKU_D = f"'הגדרות'!$D${sku_start}:$D${SKU_PAD_END}"

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

# No sample data — inventory starts empty

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
    # Filter value cell with "הכל" default
    val_cell = ws_l.cell(row=2, column=i + 1, value="הכל")
    val_cell.fill = ROW_EVEN_FILL
    val_cell.border = THIN_BORDER
    val_cell.alignment = CELL_ALIGN

# A2 (מדינה): dropdown with common countries
dv = DataValidation(type="list",
    formula1='"הכל,ישראל,ארה\"ב,גרמניה,צרפת,בריטניה,הודו,קנדה,אוסטרליה,אחר"',
    allow_blank=True)
dv.sqref = "A2"
ws_l.add_data_validation(dv)

# D2 (באיחור): dropdown
dv = DataValidation(type="list", formula1='"הכל,כן,לא"', allow_blank=True)
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
    f"*(IF($A$2=\"הכל\",1,{INV}!$Z$2:$Z$500=$A$2))"
    f"*(IF($B$2=\"הכל\",1,{INV}!$Y$2:$Y$500=$B$2))"
    f"*(IF($C$2=\"הכל\",1,{INV}!$AA$2:$AA$500=$C$2))"
    f"*(IF($D$2=\"הכל\",1,IF($D$2=\"כן\",{OVERDUE_CHECK},1-{OVERDUE_CHECK})))"
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
RED_FONT = Font(name="Calibri", bold=True, color="9C0006")

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
FILTER_FONT = Font(name="Calibri", bold=True, color="FFFFFF", size=11)
FILTER_VAL_FILL = PatternFill(start_color="D6E4F0", end_color="D6E4F0", fill_type="solid")
TITLE_FONT_D = Font(name="Calibri", bold=True, size=14, color="1F3864")
METRIC_LABEL_FONT = Font(name="Calibri", bold=True, size=11, color="1F3864")
METRIC_VAL_FONT = Font(name="Calibri", bold=True, size=16, color="1F3864")
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
           'Polar,Nylon,Mesh,PVC,Other,RIPSTOP,RIPSTOP Double Layer,'
           'Beti,MRG,3D,WaterProof Blackout Fabric,רשת רכב,'
           'Stretch,Durable,SMT,LOKI Material,אחר"'),
    ("F2", '"הכל,XS,S,M,L,XL,XXL,XXXL,One Size,N/A"'),
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

# --- Results counter (visible in filter area) ---
ws_d.merge_cells("G1:J1")
ws_d["G1"] = "תוצאות סינון"
ws_d["G1"].font = Font(name="Calibri", bold=True, size=11, color="1F3864")
ws_d["G1"].alignment = Alignment(horizontal="center", vertical="center")

# Column widths
for col, w in {"A": 18, "B": 22, "C": 20, "D": 14, "E": 18,
               "F": 14, "G": 18, "H": 18, "I": 14, "J": 28, "K": 10}.items():
    ws_d.column_dimensions[col].width = w

# ---------------------------------------------------------------------------
# Summary Area (rows 5–18)
# ---------------------------------------------------------------------------
# Reusable formula building blocks
DI = "מלאי"                 # inventory sheet name


def dflt(dash_cell, inv_col):
    """Dashboard filter term: pass all when 'הכל', else exact match."""
    return f'(IF({dash_cell}="הכל",1,{DI}!{R(inv_col)}={dash_cell}))'


F_DATA = f'({DI}!{R("D")}<>"")'
F_ENV  = f'(--{DI}!{R("AE")})'

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
     f'=SUMPRODUCT({BF}*(--{DI}!{R("AF")}))'),
    (8,  "כמות תקין ויזואלית בלבד",
     f'=SUMPRODUCT({BF}*(--{DI}!{R("AG")}))'),
    (9,  "כמות תקין תרמית בלבד",
     f'=SUMPRODUCT({BF}*(--{DI}!{R("AH")}))'),
    (10, "כמות במלאי",
     f'=SUMPRODUCT({BF}*({DI}!{R("X")}="במלאי"))'),
    (11, "כמות מושאלים",
     f'=SUMPRODUCT({BF}*({DI}!{R("X")}="מושאל"))'),
    (12, "כמות בחדר תצוגה",
     f'=SUMPRODUCT({BF}*({DI}!{R("X")}="בחדר תצוגה"))'),
    (13, "כמות בתיק הדגמה",
     f'=SUMPRODUCT({BF}*({DI}!{R("X")}="בתיק הדגמה"))'),
    (14, "כמות עם פער בתכולה",
     f'=SUMPRODUCT({BF}*({DI}!{R("AD")}="כן"))'),
    (15, "כמות גרסה מיוחדת - פיתוח",
     f'=SUMPRODUCT({BF}*({DI}!{R("J")}="כן"))'),
    (16, "תקינים ויזואלית אך הדפס לא מתאים לסביבה",
     f'=IF($A$2="הכל","—",'
     f'SUMPRODUCT({BF_NE}*({DI}!{R("AE")}=FALSE)'
     f'*(({DI}!{R("O")}="כן")+({DI}!{R("P")}="כן")>0)))'),
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
    f'{DI}!{R("M")},({DI}!{R("D")}<>"")*({DI}!{R("AE")}=FALSE)'
    f'*(({DI}!{R("O")}="כן")+({DI}!{R("P")}="כן")>0)))),""))'
)
ws_d["D17"].font = Font(name="Calibri", size=10, color="1F3864")
ws_d["D17"].alignment = CELL_ALIGN
ws_d["D17"].fill = SUMMARY_BG
ws_d["D17"].border = THIN_BORDER

# --- Results counter in filter area G2:J2 ---
ws_d.merge_cells("G2:H2")
ws_d["G2"] = f'=SUMPRODUCT({BF})'
ws_d["G2"].font = Font(name="Calibri", bold=True, size=18, color="1F3864")
ws_d["G2"].alignment = Alignment(horizontal="center", vertical="center")
ws_d["G2"].fill = FILTER_VAL_FILL
ws_d["G2"].border = THIN_BORDER
ws_d["G2"].number_format = '0" פריטים"'
ws_d.merge_cells("I2:J2")
ws_d["I2"] = f'=COUNTIF({DI}!$D$2:$D$500,"<>")'
ws_d["I2"].font = Font(name="Calibri", size=11, color="808080")
ws_d["I2"].alignment = Alignment(horizontal="center", vertical="center")
ws_d["I2"].fill = FILTER_VAL_FILL
ws_d["I2"].border = THIN_BORDER
ws_d["I2"].number_format = '"מתוך "0'

# Summary section bottom border
for c in range(1, 11):
    cell = ws_d.cell(row=18, column=c)
    cell.border = Border(bottom=Side(style="medium", color="1F3864"))

# ---------------------------------------------------------------------------
# Filtered Product Detail (rows 19–320)
# ---------------------------------------------------------------------------
ws_d.merge_cells("A19:C19")
ws_d["A19"] = "פירוט פריטים מסוננים"
ws_d["A19"].font = Font(name="Calibri", bold=True, size=13, color="FFFFFF")
ws_d["A19"].fill = PatternFill(start_color="2E75B6", end_color="2E75B6", fill_type="solid")
ws_d["A19"].border = THIN_BORDER
for c in range(4, 11):
    ws_d.cell(row=19, column=c).fill = PatternFill(
        start_color="2E75B6", end_color="2E75B6", fill_type="solid")
    ws_d.cell(row=19, column=c).border = THIN_BORDER

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

# Freeze rows 1-4 (filters stay visible when scrolling)
ws_d.freeze_panes = "A5"

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
            f'=IF($K{r}="","",INDEX({DI}!{R(icol)},$K{r}))'
        )

    # Style all cells in the row
    for c in range(1, 12):
        style_data(ws_d.cell(row=r, column=c), r - 21)

# Detail conditional formatting — quality and status columns
GREEN_FILL_D = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
GREEN_FONT_D = Font(name="Calibri", bold=True, color="006100")
ORANGE_FILL_D = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
ORANGE_FONT_D = Font(name="Calibri", color="806000")

# G column (visual OK): green when "כן"
ws_d.conditional_formatting.add(
    "G21:G320",
    FormulaRule(formula=['G21="כן"'], font=GREEN_FONT_D, fill=GREEN_FILL_D),
)
# H column (thermal OK): green when "כן"
ws_d.conditional_formatting.add(
    "H21:H320",
    FormulaRule(formula=['H21="כן"'], font=GREEN_FONT_D, fill=GREEN_FILL_D),
)
# I column (status): color by status
ws_d.conditional_formatting.add(
    "I21:I320",
    FormulaRule(formula=['I21="מושאל"'], font=ORANGE_FONT_D, fill=ORANGE_FILL_D),
)
ws_d.conditional_formatting.add(
    "I21:I320",
    FormulaRule(formula=['I21="בחדר תצוגה"'],
               font=Font(name="Calibri", color="1F4E79"),
               fill=PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")),
)

# ---------------------------------------------------------------------------
# Active Loans Detail (rows 322–423, always displayed)
# ---------------------------------------------------------------------------
LOAN_T = 322        # title row
LOAN_H = 323        # header row
LOAN_S = 324        # first data row
LOAN_E = 423        # last data row

# Visual separator before loans section
LOAN_SEP_FILL = PatternFill(start_color="1F3864", end_color="1F3864", fill_type="solid")
for c in range(1, 11):
    cell = ws_d.cell(row=LOAN_T - 1, column=c)
    cell.fill = LOAN_SEP_FILL
    cell.border = THIN_BORDER

ws_d[f"A{LOAN_T}"] = "השאלות פעילות"
ws_d[f"A{LOAN_T}"].font = Font(name="Calibri", bold=True, size=14, color="FFFFFF")
ws_d[f"A{LOAN_T}"].fill = PatternFill(start_color="2E75B6", end_color="2E75B6", fill_type="solid")
for c in range(1, 11):
    ws_d.cell(row=LOAN_T, column=c).fill = PatternFill(
        start_color="2E75B6", end_color="2E75B6", fill_type="solid")
    ws_d.cell(row=LOAN_T, column=c).border = THIN_BORDER

loan_d_headers = ["מזהה", "מוצר", "גרסה", "מושאל ל", "מדינה", "החזרה משוערת", "באיחור"]
for i, h in enumerate(loan_d_headers):
    cell = ws_d.cell(row=LOAN_H, column=i + 1, value=h)
    style_header(cell)

# Reuse column K as helper for the loans row range
style_header(ws_d.cell(row=LOAN_H, column=11, value="loan_ref"))

# Loan filter: status = מושאל (always, no dashboard filter)
LOAN_COND = f'({DI}!{R("D")}<>"")*({DI}!{R("X")}="מושאל")'

loan_d_map = {"A": "D", "B": "B", "C": "C", "D": "Y", "E": "Z", "F": "AB"}

for r in range(LOAN_S, LOAN_E + 1):
    n = r - LOAN_S + 1

    # K: helper
    ws_d[f"K{r}"] = f'=IFERROR(SMALL(IF({LOAN_COND},{D_ROW_IDX},""),{n}),"")'

    # A–F: pull data
    for dcol, icol in loan_d_map.items():
        ws_d[f"{dcol}{r}"] = (
            f'=IF($K{r}="","",INDEX({DI}!{R(icol)},$K{r}))'
        )

    # G: overdue indicator
    ws_d[f"G{r}"] = (
        f'=IF($K{r}="","",IF(AND('
        f'INDEX({DI}!{R("AB")},$K{r})<>"",'
        f'INDEX({DI}!{R("AB")},$K{r})<TODAY()),'
        f'"באיחור!","תקין"))'
    )

    # Date formatting for expected return (F)
    ws_d[f"F{r}"].number_format = "DD/MM/YYYY"

    # Style
    for c in range(1, 12):
        style_data(ws_d.cell(row=r, column=c), r - LOAN_S)

# Conditional formatting — overdue loans red highlight
RED_FILL_D = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
RED_FONT_D = Font(name="Calibri", bold=True, color="9C0006")
ws_d.conditional_formatting.add(
    f"G{LOAN_S}:G{LOAN_E}",
    FormulaRule(formula=[f'G{LOAN_S}="באיחור!"'], font=RED_FONT_D, fill=RED_FILL_D),
)
ws_d.conditional_formatting.add(
    f"A{LOAN_S}:G{LOAN_E}",
    FormulaRule(formula=[f'$G{LOAN_S}="באיחור!"'], fill=RED_FILL_D),
)


# ===========================================================================
# Build ערכים טכנולוגיים (Tech Values) sheet
# ===========================================================================
ws_tv = wb["ערכים טכנולוגיים"]
ws_tv.sheet_view.rightToLeft = True

tv_headers = [
    "מזהה פריט",    # A — manual (key to inventory)
    "שם מוצר",       # B — formula from inventory
    "גרסה",          # C — formula
    "קו מוצר",       # D — formula
    "הדפס צד A",     # E — formula
    "MWIR צד A",     # F — manual numeric
    "LWIR צד A",     # G — manual numeric
    "NIR צד A",      # H — manual numeric
    "הדפס צד B",     # I — formula
    "MWIR צד B",     # J — manual numeric
    "LWIR צד B",     # K — manual numeric
    "NIR צד B",      # L — manual numeric
    "נובל",           # M — Yes/No dropdown
    "ממוצע MWIR",    # N — formula
    "ממוצע LWIR",    # O — formula
    "ממוצע משולב",   # P — formula
]
for i, h in enumerate(tv_headers):
    cell = ws_tv.cell(row=1, column=i + 1, value=h)
    style_header(cell)

tv_widths = [16, 24, 12, 18, 18, 14, 14, 14, 18, 14, 14, 14, 10, 16, 16, 16]
for i, w in enumerate(tv_widths):
    ws_tv.column_dimensions[get_column_letter(i + 1)].width = w

ws_tv.freeze_panes = "A2"

# Data validation: Novel (M) = Yes/No
dv = DataValidation(type="list", formula1="כן_לא", allow_blank=True)
dv.sqref = "M2:M500"
ws_tv.add_data_validation(dv)

# Number format for numeric columns
for col in ("F", "G", "H", "J", "K", "L", "N", "O", "P"):
    for r in range(2, 501):
        ws_tv[f"{col}{r}"].number_format = "0.00"

# Inventory ID range for MATCH
INV_ID = f"{DI}!$D$2:$D$500"

# Formulas for all rows (2–500): lookup from inventory + averages
for r in range(2, 501):
    MR_TV = f"MATCH(A{r},{INV_ID},0)"
    # B: product name
    ws_tv[f"B{r}"] = f'=IFERROR(INDEX({DI}!$B$2:$B$500,{MR_TV}),"")'
    # C: version
    ws_tv[f"C{r}"] = f'=IFERROR(INDEX({DI}!$C$2:$C$500,{MR_TV}),"")'
    # D: product line
    ws_tv[f"D{r}"] = f'=IFERROR(INDEX({DI}!$A$2:$A$500,{MR_TV}),"")'
    # E: print A
    ws_tv[f"E{r}"] = f'=IFERROR(INDEX({DI}!$M$2:$M$500,{MR_TV}),"")'
    # I: print B
    ws_tv[f"I{r}"] = f'=IFERROR(INDEX({DI}!$N$2:$N$500,{MR_TV}),"")'
    # N: avg MWIR (average both sides if B exists, else A only)
    ws_tv[f"N{r}"] = f'=IF(F{r}="","",IF(J{r}<>"",AVERAGE(F{r},J{r}),F{r}))'
    # O: avg LWIR
    ws_tv[f"O{r}"] = f'=IF(G{r}="","",IF(K{r}<>"",AVERAGE(G{r},K{r}),G{r}))'
    # P: combined avg
    ws_tv[f"P{r}"] = f'=IF(OR(N{r}="",O{r}=""),"",AVERAGE(N{r},O{r}))'

# No sample data — tech values starts empty


# ===========================================================================
# Build לוח בקרה טכנו-מבצעי (TechOps Dashboard) sheet
# ===========================================================================
ws_t = wb["לוח בקרה טכנו-מבצעי"]
ws_t.sheet_view.rightToLeft = True

# --- Filter Area (rows 1–6) ---
# Row 1–2: first 6 filters (identical to regular dashboard)
for i, label in enumerate(filter_r1):
    c = i + 1
    cell = ws_t.cell(row=1, column=c, value=label)
    cell.font = FILTER_FONT
    cell.fill = FILTER_BG
    cell.alignment = HEADER_ALIGN
    cell.border = THIN_BORDER
    val = ws_t.cell(row=2, column=c, value="הכל")
    val.fill = FILTER_VAL_FILL
    val.alignment = CELL_ALIGN
    val.border = THIN_BORDER

# Row 3–4: next 5 filters
for i, label in enumerate(filter_r2):
    c = i + 1
    cell = ws_t.cell(row=3, column=c, value=label)
    cell.font = FILTER_FONT
    cell.fill = FILTER_BG
    cell.alignment = HEADER_ALIGN
    cell.border = THIN_BORDER
    val = ws_t.cell(row=4, column=c, value="הכל")
    val.fill = FILTER_VAL_FILL
    val.alignment = CELL_ALIGN
    val.border = THIN_BORDER

# Row 5–6: tech-specific filters
tech_filter_labels = [
    "נובל", "MWIR מינ׳", "MWIR מקס׳",
    "LWIR מינ׳", "LWIR מקס׳", "משולב מינ׳", "משולב מקס׳",
]
for i, label in enumerate(tech_filter_labels):
    c = i + 1
    cell = ws_t.cell(row=5, column=c, value=label)
    cell.font = FILTER_FONT
    cell.fill = FILTER_BG
    cell.alignment = HEADER_ALIGN
    cell.border = THIN_BORDER
    val = ws_t.cell(row=6, column=c)
    val.value = "הכל" if i == 0 else None
    val.fill = FILTER_VAL_FILL
    val.alignment = CELL_ALIGN
    val.border = THIN_BORDER

# Data validations — same base filters as regular dashboard
for sqref, f1 in dv_defs:
    dv = DataValidation(type="list", formula1=f1, allow_blank=True)
    dv.sqref = sqref
    ws_t.add_data_validation(dv)

# Novel filter (A6)
dv = DataValidation(type="list", formula1='"הכל,כן,לא"', allow_blank=True)
dv.sqref = "A6"
ws_t.add_data_validation(dv)

# Column widths
for col, w in {"A": 48, "B": 22, "C": 20, "D": 18, "E": 18,
               "F": 14, "G": 14, "H": 14, "I": 14, "J": 28,
               "K": 10, "L": 10, "M": 14, "N": 14,
               "O": 14, "P": 14, "Q": 16}.items():
    ws_t.column_dimensions[col].width = w

# ---------------------------------------------------------------------------
# TechOps formula building blocks
# ---------------------------------------------------------------------------
TV = "'ערכים טכנולוגיים'"  # sheet name (quoted for formulas)

# Inline env match (not dependent on regular dashboard AE helper)
T_ENV = (f'IF($A$2="הכל",1,'
         f'({DI}!{R("Q")}=$A$2)+({DI}!{R("R")}=$A$2)+'
         f'({DI}!{R("S")}=$A$2)+({DI}!{R("T")}=$A$2)>0)')

# Base filter with inline env (identical logic to regular, self-contained)
BF_T = (f'{F_DATA}*({T_ENV})'
        f'*{dflt("$B$2","A")}*{dflt("$C$2","B")}*{dflt("$D$2","C")}'
        f'*{dflt("$E$2","G")}*{dflt("$F$2","I")}'
        f'*{dflt("$A$4","X")}*{dflt("$B$4","Z")}'
        f'*{dflt("$C$4","O")}*{dflt("$D$4","U")}*{dflt("$E$4","K")}')

# MATCH expression: find inventory item in tech values sheet
TMATCH = f'MATCH({DI}!{R("D")},{TV}!{R("A")},0)'

# Tech filter conditions (appended to BF_T)
TF_NOV = f'IF($A$6="הכל",1,IFERROR(--(INDEX({TV}!{R("M")},{TMATCH})=$A$6),0))'
TF_MW1 = f'IF($B$6="",1,IFERROR(--(INDEX({TV}!{R("N")},{TMATCH})>=$B$6),0))'
TF_MW2 = f'IF($C$6="",1,IFERROR(--(INDEX({TV}!{R("N")},{TMATCH})<=$C$6),0))'
TF_LW1 = f'IF($D$6="",1,IFERROR(--(INDEX({TV}!{R("O")},{TMATCH})>=$D$6),0))'
TF_LW2 = f'IF($E$6="",1,IFERROR(--(INDEX({TV}!{R("O")},{TMATCH})<=$E$6),0))'
TF_CB1 = f'IF($F$6="",1,IFERROR(--(INDEX({TV}!{R("P")},{TMATCH})>=$F$6),0))'
TF_CB2 = f'IF($G$6="",1,IFERROR(--(INDEX({TV}!{R("P")},{TMATCH})<=$G$6),0))'

TF_ALL = (f'*({TF_NOV})*({TF_MW1})*({TF_MW2})'
          f'*({TF_LW1})*({TF_LW2})*({TF_CB1})*({TF_CB2})')

BF_TECH = BF_T + TF_ALL  # full techops filter

# Inline quality conditions (not dependent on regular dashboard helpers)
FULLY_OK_T = (f'(({DI}!{R("O")}="כן")*({DI}!{R("U")}="כן")'
              f'+({DI}!{R("P")}="כן")*({DI}!{R("V")}="כן")>0)')
ANY_VIS_T = f'(({DI}!{R("O")}="כן")+({DI}!{R("P")}="כן")>0)'
ANY_THERM_T = f'(({DI}!{R("U")}="כן")+({DI}!{R("V")}="כן")>0)'

# BF_TECH without env for wrong-print metric
BF_TECH_NE = (f'{F_DATA}'
              f'*{dflt("$B$2","A")}*{dflt("$C$2","B")}*{dflt("$D$2","C")}'
              f'*{dflt("$E$2","G")}*{dflt("$F$2","I")}'
              f'*{dflt("$A$4","X")}*{dflt("$B$4","Z")}'
              f'*{dflt("$C$4","O")}*{dflt("$D$4","U")}*{dflt("$E$4","K")}'
              + TF_ALL)

# Inline env FALSE for wrong-print
T_ENV_FALSE = (f'IF($A$2="הכל",FALSE,'
               f'({DI}!{R("Q")}<>$A$2)*({DI}!{R("R")}<>$A$2)*'
               f'({DI}!{R("S")}<>$A$2)*({DI}!{R("T")}<>$A$2))')

# ---------------------------------------------------------------------------
# Summary Area (rows 8–20)
# ---------------------------------------------------------------------------
ws_t["A7"] = "סיכום"
ws_t["A7"].font = TITLE_FONT_D

t_metrics = [
    (8,  'סה"כ פריטים תואמים',
     f'=SUMPRODUCT({BF_TECH})'),
    (9,  "כמות תקין לגמרי",
     f'=SUMPRODUCT({BF_TECH}*{FULLY_OK_T})'),
    (10, "כמות תקין ויזואלית בלבד",
     f'=SUMPRODUCT({BF_TECH}*{ANY_VIS_T}*(1-{FULLY_OK_T}))'),
    (11, "כמות תקין תרמית בלבד",
     f'=SUMPRODUCT({BF_TECH}*{ANY_THERM_T}*(1-{ANY_VIS_T}))'),
    (12, "כמות במלאי",
     f'=SUMPRODUCT({BF_TECH}*({DI}!{R("X")}="במלאי"))'),
    (13, "כמות מושאלים",
     f'=SUMPRODUCT({BF_TECH}*({DI}!{R("X")}="מושאל"))'),
    (14, "כמות בחדר תצוגה",
     f'=SUMPRODUCT({BF_TECH}*({DI}!{R("X")}="בחדר תצוגה"))'),
    (15, "כמות בתיק הדגמה",
     f'=SUMPRODUCT({BF_TECH}*({DI}!{R("X")}="בתיק הדגמה"))'),
    (16, "כמות עם פער בתכולה",
     f'=SUMPRODUCT({BF_TECH}*({DI}!{R("AD")}="כן"))'),
    (17, "כמות גרסה מיוחדת - פיתוח",
     f'=SUMPRODUCT({BF_TECH}*({DI}!{R("J")}="כן"))'),
    (18, "תקינים ויזואלית אך הדפס לא מתאים לסביבה",
     f'=IF($A$2="הכל","—",'
     f'SUMPRODUCT({BF_TECH_NE}*({T_ENV_FALSE})*{ANY_VIS_T}))'),
]

for row, label, formula in t_metrics:
    ws_t.merge_cells(f"A{row}:C{row}")
    cl = ws_t[f"A{row}"]
    cl.value = label
    cl.font = METRIC_LABEL_FONT
    cl.alignment = CELL_ALIGN
    cl.fill = SUMMARY_BG
    cl.border = THIN_BORDER
    cv = ws_t[f"D{row}"]
    cv.value = formula
    cv.font = METRIC_VAL_FONT
    cv.alignment = Alignment(horizontal="center", vertical="center")
    cv.fill = SUMMARY_BG
    cv.border = THIN_BORDER

# Row 19: print detail
ws_t.merge_cells("A19:C19")
ws_t["A19"].value = "פירוט הדפסים שאינם מתאימים"
ws_t["A19"].font = METRIC_LABEL_FONT
ws_t["A19"].alignment = CELL_ALIGN
ws_t["A19"].fill = SUMMARY_BG
ws_t["A19"].border = THIN_BORDER
ws_t.merge_cells("D19:J19")
ws_t["D19"] = (
    f'=IF($A$2="הכל","—",IFERROR(TEXTJOIN(", ",TRUE,UNIQUE(FILTER('
    f'{DI}!{R("M")},({DI}!{R("D")}<>"")*({T_ENV_FALSE})'
    f'*(({DI}!{R("O")}="כן")+({DI}!{R("P")}="כן")>0)))),""))'
)
ws_t["D19"].font = Font(name="Calibri", size=10, color="1F3864")
ws_t["D19"].alignment = CELL_ALIGN
ws_t["D19"].fill = SUMMARY_BG
ws_t["D19"].border = THIN_BORDER

# ---------------------------------------------------------------------------
# TechOps Filtered Product Detail (rows 22–321)
# ---------------------------------------------------------------------------
ws_t["A21"] = "פירוט פריטים מסוננים"
ws_t["A21"].font = TITLE_FONT_D

t_det_headers = [
    "מזהה", "שם מוצר", "גרסה", "סוג בד", "הדפס A", "הדפס B",
    "תקין ויזואלי", "תקין תרמי", "סטטוס", "הערות",
    # hidden helper
    "row_ref",
    # tech columns
    "נובל", "MWIR A", "MWIR B", "LWIR A", "LWIR B", "ממוצע משולב",
]
for i, h in enumerate(t_det_headers):
    cell = ws_t.cell(row=22, column=i + 1, value=h)
    style_header(cell)

# K (col 11) = hidden helper
ws_t.column_dimensions["K"].hidden = True

# Freeze below detail headers
ws_t.freeze_panes = "A23"

# Detail column → inventory column mapping (A–J same as regular)
t_det_map = {
    "A": "D", "B": "B", "C": "C", "D": "G", "E": "M",
    "F": "N", "G": "O", "H": "U", "I": "X", "J": "AC",
}

# Tech detail columns → tech values sheet column mapping (updated for NIR layout)
t_tech_map = {
    "L": "M",  # נובל
    "M": "F",  # MWIR A
    "N": "J",  # MWIR B
    "O": "G",  # LWIR A
    "P": "K",  # LWIR B
    "Q": "P",  # ממוצע משולב
}

T_ROW_IDX = f"ROW({DI}!$A$2:$A$500)-ROW({DI}!$A$2)+1"

for r in range(23, 323):
    n = r - 22  # nth match

    # K: helper — row index of nth matching inventory row
    ws_t[f"K{r}"] = f'=IFERROR(SMALL(IF({BF_TECH},{T_ROW_IDX},""),{n}),"")'

    # A–J: pull data from inventory
    for dcol, icol in t_det_map.items():
        ws_t[f"{dcol}{r}"] = (
            f'=IF($K{r}="","",INDEX({DI}!{R(icol)},$K{r}))'
        )

    # L–Q: pull tech values via INDEX/MATCH on item ID (A column = ID)
    for dcol, tcol in t_tech_map.items():
        ws_t[f"{dcol}{r}"] = (
            f'=IF($K{r}="","",IFERROR(INDEX({TV}!{R(tcol)},'
            f'MATCH(A{r},{TV}!{R("A")},0)),""))'
        )

    # Style all cells
    for c in range(1, 18):
        style_data(ws_t.cell(row=r, column=c), r - 23)

# ---------------------------------------------------------------------------
# TechOps Active Loans Detail (rows 325–424)
# ---------------------------------------------------------------------------
T_LN_T = 324   # title
T_LN_H = 325   # header
T_LN_S = 326   # first data
T_LN_E = 425   # last data

ws_t[f"A{T_LN_T}"] = "השאלות פעילות"
ws_t[f"A{T_LN_T}"].font = TITLE_FONT_D

t_loan_headers = ["מזהה", "מוצר", "גרסה", "מושאל ל", "מדינה", "החזרה משוערת", "באיחור"]
for i, h in enumerate(t_loan_headers):
    cell = ws_t.cell(row=T_LN_H, column=i + 1, value=h)
    style_header(cell)
style_header(ws_t.cell(row=T_LN_H, column=11, value="loan_ref"))

T_LOAN_COND = f'({DI}!{R("D")}<>"")*({DI}!{R("X")}="מושאל")'
t_loan_map = {"A": "D", "B": "B", "C": "C", "D": "Y", "E": "Z", "F": "AB"}

for r in range(T_LN_S, T_LN_E + 1):
    n = r - T_LN_S + 1
    ws_t[f"K{r}"] = f'=IFERROR(SMALL(IF({T_LOAN_COND},{T_ROW_IDX},""),{n}),"")'
    for dcol, icol in t_loan_map.items():
        ws_t[f"{dcol}{r}"] = (
            f'=IF($K{r}="","",INDEX({DI}!{R(icol)},$K{r}))'
        )
    ws_t[f"G{r}"] = (
        f'=IF($K{r}="","",IF(AND('
        f'INDEX({DI}!{R("AB")},$K{r})<>"",'
        f'INDEX({DI}!{R("AB")},$K{r})<TODAY()),'
        f'"באיחור!","תקין"))'
    )
    ws_t[f"F{r}"].number_format = "DD/MM/YYYY"
    for c in range(1, 18):
        style_data(ws_t.cell(row=r, column=c), r - T_LN_S)

# Conditional formatting — overdue red
ws_t.conditional_formatting.add(
    f"G{T_LN_S}:G{T_LN_E}",
    FormulaRule(formula=[f'G{T_LN_S}="באיחור!"'], font=RED_FONT_D, fill=RED_FILL_D),
)
ws_t.conditional_formatting.add(
    f"A{T_LN_S}:G{T_LN_E}",
    FormulaRule(formula=[f'$G{T_LN_S}="באיחור!"'], fill=RED_FILL_D),
)


# ===========================================================================
# Build תכולה (Contents) sheet
# ===========================================================================
ws_c = wb["תכולה"]
ws_c.sheet_view.rightToLeft = True

cont_headers = [
    "שם מוצר",       # A — dropdown from products
    "פריט תכולה",    # B — free text
    "כמות נדרשת",    # C — numeric
    "הערות",          # D — free text
]
for i, h in enumerate(cont_headers):
    cell = ws_c.cell(row=1, column=i + 1, value=h)
    style_header(cell)

cont_widths = {"A": 28, "B": 30, "C": 16, "D": 40}
for col, w in cont_widths.items():
    ws_c.column_dimensions[col].width = w

ws_c.freeze_panes = "A2"

# Style 30 empty data-entry rows
for r in range(2, 32):
    for c in range(1, 5):
        style_data(ws_c.cell(row=r, column=c), r - 2)

# C column: integer number format
for r in range(2, 32):
    ws_c[f"C{r}"].number_format = "0"


# ===========================================================================
# Build מלאי רצוי (Desired Inventory) sheet
# ===========================================================================
ws_di = wb["מלאי רצוי"]
ws_di.sheet_view.rightToLeft = True

DI_ROWS = 121  # 120 data rows (2–121)

di_headers = [
    "שם מוצר",              # A — free text / dropdown
    "גרסה",                 # B — free text
    "סביבה/הדפס",           # C — free text (env or print name)
    "סביבה צד שני",          # D — free text (second environment)
    "סוג בד",               # E — free text / dropdown
    "יעד חדר תצוגה",        # F — numeric target
    "יעד תיק הדגמות",       # G — numeric target
    "יעד השאלות",            # H — numeric target
    "קיים - חדר תצוגה",     # I — COUNTIFS formula
    "קיים - תיק",           # J — COUNTIFS formula
    "קיים - השאלות",         # K — COUNTIFS formula
    "חסר - חדר תצוגה",      # L — gap formula
    "חסר - תיק",            # M — gap formula
    "חסר - השאלות",          # N — gap formula
    "הערות",                 # O — free text
]
for i, h in enumerate(di_headers):
    cell = ws_di.cell(row=1, column=i + 1, value=h)
    style_header(cell)

di_widths = {"A": 24, "B": 14, "C": 20, "D": 20, "E": 16,
             "F": 18, "G": 18, "H": 16, "I": 20, "J": 16,
             "K": 18, "L": 20, "M": 16, "N": 18, "O": 28}
for col, w in di_widths.items():
    ws_di.column_dimensions[col].width = w

ws_di.freeze_panes = "A2"

# Target columns: integer format
for col in ("F", "G", "H"):
    for r in range(2, DI_ROWS + 1):
        ws_di[f"{col}{r}"].number_format = "0"

# Formulas for 120 data-entry rows (rows 2–121)
# COUNTIFS: count inventory items matching product + version + fabric + status
# Hierarchical: specific filters (B, E) only apply when filled
for r in range(2, DI_ROWS + 1):
    # Base product match (mandatory)
    prod_cond = f'{DI}!$B$2:$B$500,A{r}'

    # Optional fabric condition suffix (fabric is now col E)
    fab_if = f'{DI}!$G$2:$G$500,E{r},'

    # Build COUNTIFS for each status
    # Pattern: IF(E<>"", IF(B<>"", COUNTIFS(prod,ver,fab,status), COUNTIFS(prod,fab,status)),
    #                    IF(B<>"", COUNTIFS(prod,ver,status), COUNTIFS(prod,status)))
    for out_col, status in [("I", "בחדר תצוגה"), ("J", "בתיק הדגמה"), ("K", "מושאל")]:
        stat_cond = f'{DI}!$X$2:$X$500,"{status}"'
        ws_di[f"{out_col}{r}"] = (
            f'=IF(A{r}="","",IF(E{r}<>"",'
            f'IF(B{r}<>"",COUNTIFS({prod_cond},{DI}!$C$2:$C$500,B{r},{fab_if}{stat_cond}),'
            f'COUNTIFS({prod_cond},{fab_if}{stat_cond})),'
            f'IF(B{r}<>"",COUNTIFS({prod_cond},{DI}!$C$2:$C$500,B{r},{stat_cond}),'
            f'COUNTIFS({prod_cond},{stat_cond}))))'
        )

    # Gap formulas: target - actual, minimum 0
    for target_col, actual_col, gap_col in [("F", "I", "L"), ("G", "J", "M"), ("H", "K", "N")]:
        ws_di[f"{gap_col}{r}"] = (
            f'=IF(OR({target_col}{r}="",A{r}=""),"",MAX(0,{target_col}{r}-{actual_col}{r}))'
        )

    # Style all cells
    for c in range(1, 16):
        style_data(ws_di.cell(row=r, column=c), r - 2)

# Conditional formatting: red fill when gap > 0
GAP_RED = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
GAP_RED_FONT = Font(name="Calibri", bold=True, color="9C0006")
for gap_col in ("L", "M", "N"):
    ws_di.conditional_formatting.add(
        f"{gap_col}2:{gap_col}{DI_ROWS}",
        FormulaRule(
            formula=[f'{gap_col}2>0'],
            font=GAP_RED_FONT,
            fill=GAP_RED,
        ),
    )


# ===========================================================================
# Protection & Final Touches
# ===========================================================================

# --- Task 5A: Hide secret sheets ---
wb["ערכים טכנולוגיים"].sheet_state = "hidden"
wb["לוח בקרה טכנו-מבצעי"].sheet_state = "hidden"

# --- Task 5B/5C: Hide helper columns ---
# Inventory AE-AI already hidden above (lines 267-268)
# Dashboard K already hidden above
# Loans K already hidden above

# --- Inventory sheet (ws) ---
# Lock all cells first, then unlock input columns
for r in range(1, MR + 1):
    for c in range(1, 36):
        ws.cell(row=r, column=c).protection = LOCKED
# Unlock input columns for data rows 2–500
# Input cols: A-D (1-4), F-P (6-16), U-W (21-23), X-AD (24-30)
inv_input_cols = list(range(1, 5)) + list(range(6, 17)) + list(range(21, 24)) + list(range(24, 31))
for r in range(2, MR + 1):
    for c in inv_input_cols:
        ws.cell(row=r, column=c).protection = UNLOCKED
ws.protection.sheet = True
ws.protection.password = "1998"
ws.protection.enable()

# --- Settings sheet (ws_s) — data cells unlocked, headers locked ---
ws_s = wb["הגדרות"]
# Unlock all data cells (non-header rows) in the settings sheet
for row in ws_s.iter_rows(min_row=1, max_row=ws_s.max_row, max_col=ws_s.max_column):
    for cell in row:
        if cell.font and cell.font.color and hasattr(cell.font.color, 'rgb') and cell.font.color.rgb == "00FFFFFF":
            # Header cell (white text) — keep locked
            cell.protection = LOCKED
        else:
            cell.protection = UNLOCKED
ws_s.protection.sheet = True
ws_s.protection.password = "1998"
ws_s.protection.enable()

# --- Dashboard (ws_d) — unlock only filter cells ---
# Filter cells: A2, B2, C2, D2, E2, F2, A4, B4, C4, D4, E4
dash_unlock = ["A2", "B2", "C2", "D2", "E2", "F2",
               "A4", "B4", "C4", "D4", "E4"]
for ref in dash_unlock:
    ws_d[ref].protection = UNLOCKED
ws_d.protection.sheet = True
ws_d.protection.password = "1998"
ws_d.protection.enable()

# --- TechOps Dashboard (ws_t) — unlock filter cells + tech filters ---
tech_unlock = dash_unlock + ["A6", "B6", "C6", "D6", "E6", "F6", "G6"]
for ref in tech_unlock:
    ws_t[ref].protection = UNLOCKED
ws_t.protection.sheet = True
ws_t.protection.password = "1998"
ws_t.protection.enable()

# --- Loans (ws_l) — unlock filter cells ---
for ref in ["A2", "B2", "C2", "D2"]:
    ws_l[ref].protection = UNLOCKED
ws_l.protection.sheet = True
ws_l.protection.password = "1998"
ws_l.protection.enable()

# --- Tech Values (ws_tv) — unlock input columns A, F, G, H, J, K, L, M ---
tv_input_cols = [1, 6, 7, 8, 10, 11, 12, 13]  # A, F, G, H, J, K, L, M
for r in range(2, 501):
    for c in tv_input_cols:
        ws_tv.cell(row=r, column=c).protection = UNLOCKED
ws_tv.protection.sheet = True
ws_tv.protection.password = "1998"
ws_tv.protection.enable()

# --- Contents (ws_c) — unlock all data cells A2:D31 ---
for r in range(2, 32):
    for c in range(1, 5):
        ws_c.cell(row=r, column=c).protection = UNLOCKED
ws_c.protection.sheet = True
ws_c.protection.password = "1998"
ws_c.protection.enable()

# --- Desired Inventory (ws_di) — unlock input cols A-H + O, lock I-N ---
for r in range(2, DI_ROWS + 1):
    for c in list(range(1, 9)) + [15]:  # A-H, O
        ws_di.cell(row=r, column=c).protection = UNLOCKED
ws_di.protection.sheet = True
ws_di.protection.password = "1998"
ws_di.protection.enable()


# ===========================================================================
# Build רצוי מול מצוי (Desired vs Actual) sheet — interactive filtered view
# ===========================================================================
ws_cmp = wb["רצוי מול מצוי"]
ws_cmp.sheet_view.rightToLeft = True

DI_SH = "'מלאי רצוי'"
ST_SH = "'הגדרות'"
CMP_DISPLAY = 120  # display row slots

# --- DI source ranges (rows 2-121 = 120 data rows) ---
DI_END = DI_ROWS  # 121
DI_A_R = f"{DI_SH}!$A$2:$A${DI_END}"
DI_B_R = f"{DI_SH}!$B$2:$B${DI_END}"
DI_C_R = f"{DI_SH}!$C$2:$C${DI_END}"
DI_D_R = f"{DI_SH}!$D$2:$D${DI_END}"
DI_E_R = f"{DI_SH}!$E$2:$E${DI_END}"
DI_F_R = f"{DI_SH}!$F$2:$F${DI_END}"
DI_G_R = f"{DI_SH}!$G$2:$G${DI_END}"
DI_H_R = f"{DI_SH}!$H$2:$H${DI_END}"
DI_O_R = f"{DI_SH}!$O$2:$O${DI_END}"

# Product line lookup array (maps each DI product name → product line)
PL_A = f"{ST_SH}!$A${prod_start}:$A${prod_end}"
PL_B = f"{ST_SH}!$B${prod_start}:$B${prod_end}"

# --- Filter condition (evaluated over all 120 DI rows) ---
CMP_HAS_DATA = f'({DI_A_R}<>"")'
CMP_F_ENV  = f'(IF($A$2="הכל",1,{DI_C_R}=$A$2))'
CMP_F_PL   = f'(IF($B$2="הכל",1,IFERROR(--(INDEX({PL_A},MATCH({DI_A_R},{PL_B},0))=$B$2),0)))'
CMP_F_FAB  = f'(IF($C$2="הכל",1,{DI_E_R}=$C$2))'
CMP_BF = f'{CMP_HAS_DATA}*{CMP_F_ENV}*{CMP_F_PL}*{CMP_F_FAB}'
CMP_ROW_IDX = f"ROW({DI_A_R})-ROW({DI_SH}!$A$2)+1"

# --- Filter area (rows 1-2) — styled filter bar ---
cmp_filter_labels = ["סביבה", "קו מוצר", "סוג בד"]
for i, label in enumerate(cmp_filter_labels):
    cell = ws_cmp.cell(row=1, column=i + 1, value=label)
    cell.font = FILTER_FONT
    cell.fill = FILTER_BG
    cell.alignment = HEADER_ALIGN
    cell.border = THIN_BORDER
    val = ws_cmp.cell(row=2, column=i + 1, value="הכל")
    val.fill = FILTER_VAL_FILL
    val.alignment = CELL_ALIGN
    val.border = THIN_BORDER
    val.protection = UNLOCKED

# Title in the filter row area
ws_cmp.merge_cells("E1:J1")
ws_cmp["E1"] = "ריכוז רצוי מול מצוי"
ws_cmp["E1"].font = Font(name="Calibri", bold=True, size=16, color="1F3864")
ws_cmp["E1"].alignment = Alignment(horizontal="center", vertical="center")

# Filter dropdowns
dv = DataValidation(type="list",
    formula1='"הכל,מדברי,מיוער,שלג,ימי,מבולדר/שטח בנוי,אחר"',
    allow_blank=True)
dv.sqref = "A2"
ws_cmp.add_data_validation(dv)
dv = DataValidation(type="list",
    formula1='"הכל,OverGarment,Hide Site,Platform Hide Site,Uniform,'
             'Blankets,Urban,Platform On-The-Move,Accessories"',
    allow_blank=True)
dv.sqref = "B2"
ws_cmp.add_data_validation(dv)
dv = DataValidation(type="list",
    formula1='"הכל,Sahar,Inbar,SRV,Gabardine,Meron,Arber,IRR,'
             'Polar,Nylon,Mesh,PVC,Other,RIPSTOP,RIPSTOP Double Layer,'
             'Beti,MRG,3D,WaterProof Blackout Fabric,רשת רכב,'
             'Stretch,Durable,SMT,LOKI Material,אחר"',
    allow_blank=True)
dv.sqref = "C2"
ws_cmp.add_data_validation(dv)

# --- Summary bar (row 3) — key metrics that respond to filters ---
CMP_SUM_FONT = Font(name="Calibri", bold=True, size=12, color="FFFFFF")
CMP_SUM_VAL_FONT = Font(name="Calibri", bold=True, size=14, color="FFFFFF")
CMP_SUM_FILL = PatternFill(start_color="2E75B6", end_color="2E75B6", fill_type="solid")
CMP_CARD_FILL_G = PatternFill(start_color="548235", end_color="548235", fill_type="solid")
CMP_CARD_FILL_R = PatternFill(start_color="C00000", end_color="C00000", fill_type="solid")

# Apply summary fill across row 3
for c in range(1, 16):
    ws_cmp.cell(row=3, column=c).fill = CMP_SUM_FILL
    ws_cmp.cell(row=3, column=c).border = THIN_BORDER

# Summary metrics
ws_cmp["A3"] = "שורות:"
ws_cmp["A3"].font = CMP_SUM_FONT
ws_cmp["B3"] = f'=SUMPRODUCT({CMP_BF})'
ws_cmp["B3"].font = CMP_SUM_VAL_FONT
ws_cmp["B3"].number_format = "0"

ws_cmp["D3"] = 'סה"כ יעדים:'
ws_cmp["D3"].font = CMP_SUM_FONT
ws_cmp["E3"] = (f'=SUMPRODUCT({CMP_BF}*{DI_F_R})'
                f'+SUMPRODUCT({CMP_BF}*{DI_G_R})'
                f'+SUMPRODUCT({CMP_BF}*{DI_H_R})')
ws_cmp["E3"].font = CMP_SUM_VAL_FONT
ws_cmp["E3"].number_format = "0"

ws_cmp["G3"] = 'סה"כ פערים:'
ws_cmp["G3"].font = Font(name="Calibri", bold=True, size=12, color="FFFFFF")
ws_cmp["H3"] = f'=SUMPRODUCT((A6:A{5+CMP_DISPLAY}<>"")*L6:L{5+CMP_DISPLAY})' \
               f'+SUMPRODUCT((A6:A{5+CMP_DISPLAY}<>"")*M6:M{5+CMP_DISPLAY})' \
               f'+SUMPRODUCT((A6:A{5+CMP_DISPLAY}<>"")*N6:N{5+CMP_DISPLAY})'
ws_cmp["H3"].font = CMP_SUM_VAL_FONT
ws_cmp["H3"].number_format = "0"
ws_cmp["G3"].fill = CMP_CARD_FILL_R
ws_cmp["H3"].fill = CMP_CARD_FILL_R

ws_cmp["J3"] = "% מילוי:"
ws_cmp["J3"].font = CMP_SUM_FONT
ws_cmp["K3"] = f'=IF(E3=0,"—",TEXT(1-H3/E3,"0%"))'
ws_cmp["K3"].font = CMP_SUM_VAL_FONT

# --- Row 4: separator ---
for c in range(1, 16):
    ws_cmp.cell(row=4, column=c).border = THIN_BORDER

# --- Headers (row 5) ---
cmp_headers = [
    "שם מוצר",           # A
    "סביבה",             # B
    "סביבה צד שני",       # C
    "סוג בד",            # D
    "יעד תצוגה",          # E
    "יעד תיק",           # F
    "יעד השאלות",         # G
    "מצוי תצוגה",         # H
    "מצוי תיק",          # I
    "מצוי השאלות",        # J
    'מצוי סה"כ',          # K
    "פער תצוגה",          # L
    "פער תיק",           # M
    "פער השאלות",         # N
    "הערות",             # O
]
for i, h in enumerate(cmp_headers):
    cell = ws_cmp.cell(row=5, column=i + 1, value=h)
    style_header(cell)

# Hidden helper column P
style_header(ws_cmp.cell(row=5, column=16, value="row_ref"))
ws_cmp.column_dimensions["P"].hidden = True

# Column widths
cmp_widths = {"A": 28, "B": 16, "C": 16, "D": 18, "E": 14, "F": 14,
              "G": 14, "H": 14, "I": 14, "J": 14, "K": 14, "L": 14,
              "M": 14, "N": 14, "O": 28, "P": 10}
for col, w in cmp_widths.items():
    ws_cmp.column_dimensions[col].width = w

ws_cmp.freeze_panes = "A6"

# --- Filtered data rows (6-125) using SMALL/IF ---
for r in range(6, 6 + CMP_DISPLAY):
    n = r - 5  # nth match

    # P: hidden helper — row index into DI for nth matching row
    ws_cmp[f"P{r}"] = f'=IFERROR(SMALL(IF({CMP_BF},{CMP_ROW_IDX},""),{n}),"")'

    # A-D: pull from desired inventory via INDEX
    ws_cmp[f"A{r}"] = f'=IF($P{r}="","",INDEX({DI_A_R},$P{r}))'
    ws_cmp[f"B{r}"] = f'=IF($P{r}="","",INDEX({DI_C_R},$P{r}))'
    ws_cmp[f"C{r}"] = f'=IF($P{r}="","",INDEX({DI_D_R},$P{r}))'
    ws_cmp[f"D{r}"] = f'=IF($P{r}="","",INDEX({DI_E_R},$P{r}))'
    # E-G: targets
    ws_cmp[f"E{r}"] = f'=IF($P{r}="","",INDEX({DI_F_R},$P{r}))'
    ws_cmp[f"F{r}"] = f'=IF($P{r}="","",INDEX({DI_G_R},$P{r}))'
    ws_cmp[f"G{r}"] = f'=IF($P{r}="","",INDEX({DI_H_R},$P{r}))'
    # O: notes
    ws_cmp[f"O{r}"] = f'=IF($P{r}="","",INDEX({DI_O_R},$P{r}))'

    # H-J: actual counts (COUNTIFS from inventory, using displayed product+fabric)
    prod_cond = f'{DI}!$B$2:$B$500,$A{r}'
    fab_if = f'{DI}!$G$2:$G$500,$D{r},'
    for out_col, status in [("H", "בחדר תצוגה"), ("I", "בתיק הדגמה"), ("J", "מושאל")]:
        stat_cond = f'{DI}!$X$2:$X$500,"{status}"'
        ws_cmp[f"{out_col}{r}"] = (
            f'=IF($A{r}="","",IF($D{r}<>"",'
            f'COUNTIFS({prod_cond},{fab_if}{stat_cond}),'
            f'COUNTIFS({prod_cond},{stat_cond})))'
        )

    # K: total actual (all items matching product+fabric with any ID)
    ws_cmp[f"K{r}"] = (
        f'=IF($A{r}="","",IF($D{r}<>"",'
        f'COUNTIFS({prod_cond},{fab_if}{DI}!$D$2:$D$500,"<>"),'
        f'COUNTIFS({prod_cond},{DI}!$D$2:$D$500,"<>")))'
    )

    # L-N: gaps (target - actual, min 0)
    for tgt, act, gap in [("E", "H", "L"), ("F", "I", "M"), ("G", "J", "N")]:
        ws_cmp[f"{gap}{r}"] = f'=IF(OR($A{r}="",$E{r}=""),"",MAX(0,{tgt}{r}-{act}{r}))'

    # Style all cells (including hidden P)
    for c in range(1, 17):
        style_data(ws_cmp.cell(row=r, column=c), r - 6)

# Number format for target/actual/gap columns
for col in ("E", "F", "G", "H", "I", "J", "K", "L", "M", "N"):
    for r in range(6, 6 + CMP_DISPLAY):
        ws_cmp[f"{col}{r}"].number_format = "0"

# --- Conditional formatting ---
CMP_LAST = 5 + CMP_DISPLAY  # 125

# Red when gap > 0
for gap_col in ("L", "M", "N"):
    ws_cmp.conditional_formatting.add(
        f"{gap_col}6:{gap_col}{CMP_LAST}",
        FormulaRule(
            formula=[f'{gap_col}6>0'],
            font=Font(name="Calibri", bold=True, color="9C0006"),
            fill=PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid"),
        ),
    )

# Green when gap = 0 and target > 0 (fully met)
for gap_col, tgt_col in [("L", "E"), ("M", "F"), ("N", "G")]:
    ws_cmp.conditional_formatting.add(
        f"{gap_col}6:{gap_col}{CMP_LAST}",
        FormulaRule(
            formula=[f'AND({gap_col}6=0,{tgt_col}6>0)'],
            font=Font(name="Calibri", bold=True, color="006100"),
            fill=PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid"),
        ),
    )

# --- Protection ---
for ref in ["A2", "B2", "C2"]:
    ws_cmp[ref].protection = UNLOCKED
ws_cmp.protection.sheet = True
ws_cmp.protection.password = "1998"
ws_cmp.protection.enable()


# ===========================================================================
# Save
# ===========================================================================
wb.save("ametrine_inventory.xlsx")
