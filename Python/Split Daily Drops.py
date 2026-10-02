
import os
import pandas as pd

# 1. Define your exact folder path so Python doesn't get lost
base_path = r"C:\Users\REDTECH\Desktop\fraud_hub\data"

# 2. Update the script to use that base_path
os.makedirs(f"{base_path}\daily_drops", exist_ok=True)
pay = pd.read_csv(f"{base_path}\payments.csv")

pay["day"] = pay["txn_time"].str[:10]
days = sorted(pay["day"].unique())
resend = pay.sample(30, random_state=7)          # rows that get sent twice

total_rows = 0
for i, d in enumerate(days):
    chunk = pay[pay["day"] == d].drop(columns="day").copy()
    if i > 0:
        prev = resend[resend["day"] == days[i - 1]].drop(columns="day")
        chunk = pd.concat([chunk, prev])          # duplicates arrive next day
    if i % 6 == 5:
        chunk["txn_type"] = chunk["txn_type"].str.upper()
        chunk["status"] = chunk["status"] + "  "
    
    # 3. Save the new files into the correct folder
    chunk.to_csv(f"{base_path}\daily_drops\payments_{d}.csv", index=False)
    total_rows += len(chunk)

print(f"{len(days)} files written, {total_rows} rows total "
      f"({len(pay)} unique + {total_rows - len(pay)} duplicates)")