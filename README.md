# 🎰 iGaming Fraud & Risk Operations Hub

An end-to-end operations automation project that replaces a manual, spreadsheet-driven fraud investigation process. **All data is synthetic**; no real players, payments or clients.

---

### 🎯 The Problem

Fraud & risk analysts consolidate daily payment files by hand, hunt for multi-accounting and bonus abuse in spreadsheets, log cases on a shared sheet, and rebuild the same report every morning.

---

### 🛠️ What I Built

| Layer | Tool | What it does |
| :--- | :--- | :--- |
| **Data generation** | Python (Pandas) | 2,000 players, ~55k payments over 183 days (1 Apr to 30 Sep 2026) with seeded fraud patterns |
| **Detection** | Python + SQL Server | Rule-based flags: shared device/IP, shared card, deposit velocity, bonus abuse; weighted risk score |
| **Consolidation** | Power Query | Combines 183 daily drop files, cleans text, removes re-sent duplicates, joins to SQL data |
| **Team tooling** | Excel VBA | One-click daily report with PDF export, review templates, PSP reconciliation |
| **Case management** | Power Apps + SharePoint | Queue, assignment, dispositions, SLA tracking, approval rule for high-risk cases |
| **Workflow** | Power Automate | **In progress** (ingest, assignment, approvals, SLA reminders) |
| **Reporting** | Power BI | Live queue, turnaround, throughput and fraud trend dashboard |

---

### 📐 Design Decisions

* **Staging and core schemas**, so bad input never breaks the load
* **Rule weights and thresholds** are parameters/data, not hard-coded
* **Thresholds of 3+ players** leave innocent shared-IP households unflagged
* **High-risk cases** cannot be closed by an analyst; they route to approval

---

### ⚙️ Run It

1. `python python/generate_data.py`, then `python python/detect_risk.py`
2. Run `sql/phase2_sqlserver.sql` in SSMS
3. Follow the steps in each folder's notes

---

### 📌 Status

Phases 1-5 complete. Phase 6 (Power Automate) in progress.
