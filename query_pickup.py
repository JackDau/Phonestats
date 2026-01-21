import pandas as pd

file_path = r"c:\Users\jackd\OneDrive - YourGP\Claude\Phone Dashboard\Call Data 2026_01_18 CLeo.xlsx"
df = pd.read_excel(file_path, sheet_name='Data In and Out')

# Filter for incoming calls that were answered (TimeToAnswer > 0)
answered = df[(df['Direction'] == 'In') & (df['TimeToAnswer'] > 0)]

print(f"Total answered incoming calls: {len(answered)}")
print()

# Find minimum pickup time
min_time = answered['TimeToAnswer'].min()
print(f"Shortest pickup time: {min_time} seconds")

# Who picked up fastest?
fastest = answered[answered['TimeToAnswer'] == min_time]
print(f"\nFastest pickup(s):")
print(fastest[['CallDateTime', 'UserName', 'TimeToAnswer', 'CallDuration']].to_string())

# Distribution of quick pickups
print("\n--- Pickup time distribution ---")
print(f"Under 5 seconds: {len(answered[answered['TimeToAnswer'] < 5])}")
print(f"5-10 seconds: {len(answered[(answered['TimeToAnswer'] >= 5) & (answered['TimeToAnswer'] < 10)])}")
print(f"10-20 seconds: {len(answered[(answered['TimeToAnswer'] >= 10) & (answered['TimeToAnswer'] < 20)])}")
print(f"20-30 seconds: {len(answered[(answered['TimeToAnswer'] >= 20) & (answered['TimeToAnswer'] < 30)])}")
print(f"30-60 seconds: {len(answered[(answered['TimeToAnswer'] >= 30) & (answered['TimeToAnswer'] < 60)])}")
print(f"Over 60 seconds: {len(answered[answered['TimeToAnswer'] >= 60])}")
