import server
import datetime

print("=========================================================================")
print("=== Z1F ENGINEERING PROGRESS MEASUREMENT VERIFICATION REPORT ===")
print("=========================================================================\n")

t0 = datetime.datetime.now()
data = server.extract_all_data(force_refresh=True)
t1 = datetime.datetime.now()

print(f"Data Ingestion & Calculation Time: {(t1-t0).total_seconds():.3f} seconds\n")

sections = [
    ("executive", "Executive Overview across All Work Packages"),
    ("wp1_topside_excl", "WP-1 Topside Excl. Structure"),
    ("wp1_topside_structure", "WP-1 Topside Structure"),
    ("wp1_jacket", "WP-1 Jacket"),
    ("wp2_pipeline", "WP-2 Pipeline")
]

for key, name in sections:
    sec = data[key]
    kpi = sec["kpi"]
    fname = sec.get("filename", "Aggregated (4 files)")
    print(f"--- {name} [{fname}] ---")
    print(f"  Total Deliverables:        {kpi['total_docs']}")
    print(f"  Overall Progress (Plan vs Act): {kpi.get('plan_progress_pct', 0.0):.2f}% Baseline Plan vs. {kpi.get('actual_progress_pct', 0.0):.2f}% Achieved (Forecast: {kpi.get('forecast_progress_pct', 0.0):.2f}%)")
    print(f"  Schedule Performance Index: SPI = {kpi.get('spi', 1.0):.2f} (Variance: {kpi.get('variance_pct', 0.0):+.2f}%)")
    print(f"  Total Delayed Gates:       {kpi['delayed_count']} (Type 3.1 Submitted Late: {kpi['delayed_type1_count']} | Type 3.2 Overdue: {kpi['delayed_type2_count']})")
    print(f"  14-Day Lookahead Risk:     {kpi['lookahead_count']} gates due in 14 days ({kpi['lookahead_slipping_count']} slipping from baseline Plan)")
    
    # Count document statuses
    status_counts = {}
    for doc in sec.get("docs", []):
        st = doc["status"]
        status_counts[st] = status_counts.get(st, 0) + 1
    print(f"  Document Status Mix:       {status_counts}")
    print()

# Audit Delay Classification exactness
print("=== DELAY & STATUS CLASSIFICATION AUDIT ===")
type1_errors = 0
type2_errors = 0
status_errors = 0

for key, _ in sections[1:]:
    for d in data[key]["delayed_list"]:
        if d["delay_type_code"] == "3.1":
            if d["submit_date"] == "-" or d["submit_date"] is None:
                type1_errors += 1
        elif d["delay_type_code"] == "3.2":
            if d["submit_date"] != "-":
                type2_errors += 1
                
    for doc in data[key]["docs"]:
        st = doc["status"]
        afc_app_d = doc.get("afc_app_date")
        afc_sub_d = doc.get("afc_sub_date") or doc.get("afc_submit_date")
        if st == "COMPLETE":
            if not afc_app_d or afc_app_d == "-":
                status_errors += 1
        elif st == "PENDING FINAL APPROVAL":
            if (not afc_sub_d or afc_sub_d == "-") or (afc_app_d and afc_app_d != "-"):
                status_errors += 1

print(f"Type 3.1 Delay Errors: {type1_errors}")
print(f"Type 3.2 Delay Errors: {type2_errors}")
print(f"Status Classification Errors: {status_errors}")
if type1_errors == 0 and type2_errors == 0 and status_errors == 0:
    print("[PASS] All verification audits PASSED with 100% compliance!")
else:
    print("[FAIL] Verification audits found discrepancies!")
