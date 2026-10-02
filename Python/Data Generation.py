import os
import numpy as np
import pandas as pd
from typing import Any
from datetime import datetime, timedelta

rng = np.random.default_rng(42)
N_PLAYERS = 2000
START = datetime(2026, 4, 1)
DAYS = 183
SCALE = DAYS / 60
os.makedirs("data", exist_ok=True)

FIRST = ["Aiden", "Maya", "Liam", "Sofia", "Noah", "Zara", "Ethan", "Layla", "Omar", "Chloe",
         "Ravi", "Emma", "Yusuf", "Nina", "Lucas", "Amara", "Dev", "Ivy", "Hugo", "Tara"]
LAST = ["Smith", "Perera", "Khan", "Silva", "Brown", "Nair", "Lopez", "Ahmed", "Novak", "Kim",
        "Fernando", "Jones", "Patel", "Garcia", "Wong", "Ali", "Costa", "Mendis", "Taylor", "Reid"]
COUNTRIES = ["GB", "DE", "CA", "IN", "BR", "SE", "FI", "NZ"]
METHODS = ["card", "ewallet", "bank_transfer", "crypto"]


def ip():
    return ".".join(str(x) for x in rng.integers(11, 250, 4))


def ts_between(start: datetime, seconds: int | float) -> datetime:
    return start + timedelta(seconds=int(seconds))


ids = [f"P{i:05d}" for i in range(1, N_PLAYERS + 1)]
order = list(rng.permutation(ids))


roles: dict[str, str] = {}
clusters: list[list[str]] = []        
idx = 0
for _ in range(25):
    size = int(rng.integers(3, 6))
    clusters.append(order[idx:idx + size])
    idx += size
for c in clusters:
    for p in c:
        roles[p] = "multi_account"
abusers = order[idx:idx + 60]; idx += 60
for p in abusers:
    roles[p] = "bonus_abuse"
burst = order[idx:idx + 30]; idx += 30
for p in burst:
    roles[p] = "payment_velocity"
card_groups = [order[idx + i * 3: idx + i * 3 + 3] for i in range(15)]; idx += 45
for g in card_groups:
    for p in g:
        roles[p] = "shared_card"


rows: list[list[str]] = []
cluster_reg: dict[str, int] = {}
for c in clusters:
    base = int(rng.integers(5, 400))
    for p in c:
        cluster_reg[p] = base
for p in ids:
    days_before = cluster_reg.get(p, int(rng.integers(1, 400)))
    reg = START - timedelta(days=days_before) + timedelta(hours=int(rng.integers(0, 48)))
    fn, ln = rng.choice(FIRST), rng.choice(LAST)
    rows.append([p, f"{fn} {ln}", f"{fn.lower()}.{ln.lower()}{rng.integers(1, 999)}@mail.test",
                 rng.choice(COUNTRIES), reg.strftime("%Y-%m-%d %H:%M:%S")])
players = pd.DataFrame(rows, columns=["player_id", "full_name", "email", "country", "registered_at"])


dev_rows: list[list[str]] = []
n_dev = 0
def new_dev():
    global n_dev
    n_dev += 1
    return f"D{n_dev:05d}"

cluster_dev: dict[str, tuple[str, str]] = {}
for c in clusters:
    d, i = new_dev(), ip()
    for p in c:
        cluster_dev[p] = (d, i)
household_ip: dict[str, str] = {}
normals = [p for p in ids if p not in roles]
for k in range(40):   # noise: innocent households sharing an IP only
    a, b = rng.choice(normals, 2, replace=False)
    shared = ip()
    household_ip[str(a)] = shared; household_ip[str(b)] = shared
for p in ids:
    if p in cluster_dev:
        d, i = cluster_dev[p]
        dev_rows.append([p, d, i])
    else:
        for _ in range(int(rng.integers(1, 3))):
            dev_rows.append([p, new_dev(), household_ip.get(p, ip())])
devices = pd.DataFrame(dev_rows, columns=["player_id", "device_id", "ip_address"])


pay: list[list[Any]] = []
tx: int = 0
card_of = {p: f"C{int(p[1:]):08d}" for p in ids}
for g in card_groups:
    shared_card = f"CSH{g[0][1:]}"
    for p in g:
        card_of[p] = shared_card

def add(p: str, t: datetime, kind: str, method: str, amount: float, status: str = "success") -> None:
    global tx
    tx += 1
    pay.append([f"T{tx:07d}", p, t.strftime("%Y-%m-%d %H:%M:%S"), kind, method,
                round(float(amount), 2), status, card_of[p] if method == "card" else ""])

for p in ids:
    method = rng.choice(METHODS, p=[0.5, 0.25, 0.15, 0.1])
    if roles.get(p) == "shared_card":
        method = "card"
    for _ in range(int(rng.integers(3, 14) * SCALE)):
        t = ts_between(START, int(rng.integers(0, DAYS * 86400)))
        add(p, t, "deposit", method, rng.lognormal(3.8, 0.7), "success" if rng.random() > 0.05 else "failed")
    for _ in range(int(rng.integers(0, 3))):
        t = ts_between(START, int(rng.integers(0, DAYS * 86400)))
        add(p, t, "withdrawal", method, rng.lognormal(4.2, 0.6))
for p in burst:
    for _ in range(2):
        t0 = ts_between(START, int(rng.integers(0, (DAYS - 1) * 86400)))
        for _ in range(int(rng.integers(7, 14))):
            add(p, t0 + timedelta(minutes=int(rng.integers(0, 55))), "deposit", "card",
                rng.uniform(10, 60), "success" if rng.random() > 0.4 else "failed")
payments = pd.DataFrame(pay, columns=["txn_id", "player_id", "txn_time", "txn_type", "method",
                                      "amount", "status", "card_hash"])


bon: list[list[Any]] = []
b: int = 0
for p in ids:
    if p in abusers or rng.random() < 0.5:
        b += 1
        amt = float(rng.choice([25, 50, 100]))
        claim = ts_between(START, int(rng.integers(0, (DAYS - 3) * 86400)))
        if p in abusers:
            wager_ratio = rng.uniform(0.5, 3.0)
            wd_hours = rng.uniform(0.5, 8)
        else:
            wager_ratio = rng.uniform(15, 45)
            wd_hours = rng.uniform(72, 600) if rng.random() < 0.4 else np.nan
        bon.append([f"B{b:06d}", p, claim.strftime("%Y-%m-%d %H:%M:%S"), amt,
                    round(amt * wager_ratio, 2),
                    round(wd_hours, 1) if not np.isnan(wd_hours) else ""])
bonuses = pd.DataFrame(bon, columns=["bonus_id", "player_id", "claimed_at", "bonus_amount",
                                     "wagered_amount", "hours_to_first_withdrawal"])


players.to_csv("data/players.csv", index=False)
devices.to_csv("data/devices.csv", index=False)
payments.to_csv("data/payments.csv", index=False)
bonuses.to_csv("data/bonuses.csv", index=False)
pd.DataFrame({"player_id": list(roles), "seeded_pattern": list(roles.values())}).to_csv(
    "data/ground_truth.csv", index=False)
print({k: len(v) for k, v in dict(players=players, devices=devices, payments=payments,
                                  bonuses=bonuses).items()})