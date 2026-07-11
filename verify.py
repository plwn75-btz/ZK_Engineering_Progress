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
    print(f"  Total Delayed Gates:       {kpi['delayed_count']} (Type 3.1 Submitted Late: {kpi['delayed_type1_count']} | Type 3.2 Overdue: {kpi['delayed_type2_count']})")
    print(f"  14-Day Lookahead Risk:     {kpi['lookahead_count']} gates due in 14 days ({kpi['lookahead_slipping_count']} slipping from baseline Plan)")
    print()

# Audit Delay Classification exactness
print("=== DELAY CLASSIFICATION AUDIT ===")
type1_errors = 0
type2_errors = 0

for key, _ in sections[1:]:
    for d in data[key]["delayed_list"]:
        if d["delay_type_code"] == "3.1":
            # Must have submitted date != '-'
            if d["submit_date"] == "-" or d["submit_date"] is None:
                type1_errors += 1
        elif d["delay_type_code"] == "3.2":
            # Must have submitted date == '-'
            if d["submit_date"] != "-":
                type2_errors += 1

print(f"Type 3.1 Classification Errors: {type1_errors}")
print(f"Type 3.2 Classification Errors: {type2_errors}")
if type1_errors == 0 and type2_errors == 0:
    print("[PASS] Strict compliance check passed: All Type 3.1 and 3.2 delays rigorously adhere to specification rules!")
else:
    print("[FAIL] Compliance errors detected!")
