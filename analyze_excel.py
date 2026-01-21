import pandas as pd

file_path = r"c:\Users\jackd\OneDrive - YourGP\Claude\Phone Dashboard\Call Data 2026_01_18 CLeo.xlsx"
xl = pd.ExcelFile(file_path)

print("=" * 60)
print("SHEET NAMES:")
print("=" * 60)
for i, name in enumerate(xl.sheet_names):
    print(f"  {i+1}. {name}")

print("\n" + "=" * 60)
print("ANALYZING EACH SHEET:")
print("=" * 60)

for sheet_name in xl.sheet_names:
    print(f"\n--- Sheet: '{sheet_name}' ---")
    df = pd.read_excel(xl, sheet_name=sheet_name)
    print(f"Shape: {df.shape[0]} rows x {df.shape[1]} columns")
    print(f"Columns: {list(df.columns)}")
    if df.shape[0] > 0:
        print(f"First 3 rows preview:")
        print(df.head(3).to_string())
    print()
