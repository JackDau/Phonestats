import pandas as pd

file_path = r"c:\Users\jackd\OneDrive - YourGP\Claude\Phone Dashboard\Call Data 2026_01_18 CLeo.xlsx"
df = pd.read_excel(file_path, sheet_name='Data In and Out')

print("All columns in Data In and Out:")
print("-" * 40)
for col in df.columns:
    print(f"  {col}")

print("\n" + "=" * 60)
print("Checking for transfer/callback indicators...")
print("=" * 60)

# Check HangupReason values
print("\nHangupReason unique values:")
print(df['HangupReason'].value_counts())

# Check if there are any columns that might indicate transfers
print("\nSample of 5 rows (all columns):")
print(df.head().T)

# Look for patterns - same phone number calling multiple times
print("\n" + "=" * 60)
print("Checking for repeat callers (potential callbacks)...")
print("=" * 60)
if 'OriginNumber' in df.columns:
    incoming = df[df['Direction'] == 'In']
    caller_counts = incoming['OriginNumber'].value_counts()
    repeat_callers = caller_counts[caller_counts > 1]
    print(f"\nTotal unique incoming numbers: {len(caller_counts)}")
    print(f"Numbers that called more than once: {len(repeat_callers)}")
    print(f"\nTop repeat callers:")
    print(repeat_callers.head(10))
