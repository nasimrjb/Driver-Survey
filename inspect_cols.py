import openpyxl, json, sys
sys.stdout.reconfigure(encoding='utf-8')

wb = openpyxl.load_workbook(r"D:\Work\Driver Survey\03) z. All g 52 - Routine.xlsx", data_only=True)
ws_survey = wb["Survey"]

# Get the header row (row 1) for specific columns we care about
target_cols = ["BK", "BU", "BY", "BP", "CP", "D", "E", "F", "B"]
for col_letter in target_cols:
    cell = ws_survey[f"{col_letter}1"]
    print(f"Survey!{col_letter}1 = '{cell.value}'")

# Also show a few data rows for these columns to understand values
print("\nSample data rows 2-4:")
for row in range(2, 5):
    for col_letter in target_cols:
        cell = ws_survey[f"{col_letter}{row}"]
        print(f"  Survey!{col_letter}{row} = {repr(cell.value)}")
    print()

# Also check what N1 and N2 contain in sheet #18
wb2 = openpyxl.load_workbook(r"D:\Work\Driver Survey\03) z. All g 52 - Routine.xlsx", data_only=True)
ws18 = wb2["#18"]
print(f"\n#18!N1 = {ws18['N1'].value}")
print(f"#18!N2 = {ws18['N2'].value}")
print(f"#18!B1 = {ws18['B1'].value}")
print(f"#18!B2 = {ws18['B2'].value}")
print(f"#18!B10 = {ws18['B10'].value}")  # city value example
print(f"#18!E10 = {ws18['E10'].value}")  # joint drivers
print(f"#18!D10 = {ws18['D10'].value}")  # total drivers
print(f"#18!F10 = {ws18['F10'].value}")  # who got message

with open(r"D:\Work\Driver Survey\Sources\column_rename_mapping.json", encoding="utf-8") as f:
    mapping = json.load(f)

# Print all keys with their "long" names and sections to find tapsi-related incentive columns
print("\n--- Incentive/Tapsi-related mapping entries ---")
for key, val in mapping.items():
    section = val.get("section", "")
    long_name = val.get("long", "")
    if any(kw in section.lower() or kw in long_name.lower() or kw in key.lower()
           for kw in ["tapsi_incentive", "commiss", "wheel", "free", "message", "incentive"]):
        print(f"  key={key!r}, long={long_name!r}, section={section!r}, type={val.get('type')}")
        if val.get("answers"):
            print(f"    answers: {val['answers']}")
