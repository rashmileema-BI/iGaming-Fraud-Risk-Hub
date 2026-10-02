"""Phase 1b: Rule-based fraud & risk detection. Output: output/daily_exceptions.csv"""
import os
import pandas as pd

os.makedirs("output", exist_ok=True)
players = pd.read_csv("data/players.csv")
devices = pd.read_csv("data/devices.csv")
payments = pd.read_csv("data/payments.csv", parse_dates=["txn_time"])
bonuses = pd.read_csv("data/bonuses.csv")

# Tunable thresholds (document these in your write-up)
DEVICE_MIN_PLAYERS = 3     # players on one device
IP_MIN_PLAYERS = 3         # players on one IP
CARD_MIN_PLAYERS = 3       # players using one card
VELOCITY_MIN_DEPOSITS = 5  # deposits inside 1 hour
WAGER_RATIO_MAX = 5        # wagered / bonus below this is suspicious...
WITHDRAW_HOURS_MAX = 24    # ...if cash-out happens this fast
WEIGHTS = {"MULTI_ACCOUNT_DEVICE": 50, "SHARED_IP": 15, "SHARED_CARD": 50,
           "PAYMENT_VELOCITY": 40, "BONUS_ABUSE": 45}

flags = []  # (player_id, rule, detail)

# 1. Multi-accounting: shared device / IP
for col, minp, rule in [("device_id", DEVICE_MIN_PLAYERS, "MULTI_ACCOUNT_DEVICE"),
                        ("ip_address", IP_MIN_PLAYERS, "SHARED_IP")]:
    n = devices.groupby(col)["player_id"].nunique()
    hot = n[n >= minp]
    for val, cnt in hot.items():
        for p in devices.loc[devices[col] == val, "player_id"].unique():
            flags.append((p, rule, f"{col}={val} shared by {cnt} players"))

# 2. Shared payment card
cards = payments[payments["card_hash"].notna() & (payments["card_hash"] != "")]
n = cards.groupby("card_hash")["player_id"].nunique()
for val, cnt in n[n >= CARD_MIN_PLAYERS].items():
    for p in cards.loc[cards["card_hash"] == val, "player_id"].unique():
        flags.append((p, "SHARED_CARD", f"card {val} used by {cnt} players"))

# 3. Deposit velocity: rolling 1-hour deposit count per player
dep = payments[payments["txn_type"] == "deposit"].sort_values("txn_time")
roll = (dep.set_index("txn_time").groupby("player_id")["amount"].rolling("1h").count())
peak = roll.groupby("player_id").max()
for p, cnt in peak[peak >= VELOCITY_MIN_DEPOSITS].items():
    flags.append((p, "PAYMENT_VELOCITY", f"{int(cnt)} deposits within 1 hour"))

# 4. Bonus abuse: low wagering + fast withdrawal
bn = bonuses.copy()
bn["wager_ratio"] = bn["wagered_amount"] / bn["bonus_amount"]
bad = bn[(bn["wager_ratio"] < WAGER_RATIO_MAX) &
         (pd.to_numeric(bn["hours_to_first_withdrawal"], errors="coerce") < WITHDRAW_HOURS_MAX)]
for _, r in bad.iterrows():
    flags.append((r["player_id"], "BONUS_ABUSE",
                  f"wagered {r['wager_ratio']:.1f}x, withdrew after {r['hours_to_first_withdrawal']}h"))

f = pd.DataFrame(flags, columns=["player_id", "rule", "detail"]).drop_duplicates(["player_id", "rule"])
f["points"] = f["rule"].map(WEIGHTS)
agg = f.groupby("player_id").agg(risk_score=("points", "sum"),
                                 rules=("rule", lambda s: "; ".join(sorted(s))),
                                 details=("detail", lambda s: " | ".join(s))).reset_index()
agg["risk_band"] = pd.cut(agg["risk_score"], [0, 39, 59, 1000], labels=["Low", "Medium", "High"])
out = agg.merge(players[["player_id", "full_name", "country"]], on="player_id", how="left")
out = out.sort_values("risk_score", ascending=False)
out.to_csv("output/daily_exceptions.csv", index=False)

# ---- validation against seeded ground truth ----
gt = pd.read_csv("data/ground_truth.csv")
caught = set(out["player_id"])
seeded = set(gt["player_id"])
print(f"Flagged players: {len(caught)} | Seeded fraud players: {len(seeded)}")
print(f"Recall (seeded caught): {len(caught & seeded) / len(seeded):.0%}")
print(f"Precision (flagged that are seeded): {len(caught & seeded) / len(caught):.0%}")
print(out["risk_band"].value_counts().to_string())
print(f["rule"].value_counts().to_string())