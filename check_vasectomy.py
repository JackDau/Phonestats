import pandas as pd

file_path = r"c:\Users\jackd\OneDrive - YourGP\Claude\Phone Dashboard\Call Data 2026_01_18 CLeo.xlsx"
xl = pd.ExcelFile(file_path)

# Check unique values in key columns to understand filtering
print("=" * 60)
print("CHECKING HOW VASECTOMY CALLS ARE IDENTIFIED")
print("=" * 60)

# Check Data In and Out for unique values
df_all = pd.read_excel(xl, sheet_name='Data In and Out')
print("\nUnique TargetNumbers:", df_all['TargetNumber'].unique()[:20])
print("\nUnique CallAlertName:", df_all['CallAlertName'].unique())
print("\nUnique OfficeName:", df_all['OfficeName'].dropna().unique())

# Check if there's a specific number or identifier for vasectomy
print("\n" + "=" * 60)
print("CHECKING VASECTOMY SHEET FOR CLUES")
print("=" * 60)

df_vas = pd.read_excel(xl, sheet_name='Vasectomy')
print("\nVas Calls column unique values:", df_vas['Vas Calls'].unique())
print("\nSample rows where Vas Calls = 1:")
vas_calls = df_vas[df_vas['Vas Calls'] == 1]
if len(vas_calls) > 0:
    print(vas_calls[['Date', 'Time', 'Call Answered', 'Vas Calls']].head(10))
else:
    print("No rows with Vas Calls = 1")

# Check if there's a pattern in the raw data
print("\n" + "=" * 60)
print("CHECKING DATA IN SHEET")
print("=" * 60)
df_in = pd.read_excel(xl, sheet_name='Data In')
print("\nColumn names:", list(df_in.columns))
print("\nFirst 5 rows:")
print(df_in.head())
