from openpyxl import Workbook

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

wb.save("ametrine_inventory.xlsx")
