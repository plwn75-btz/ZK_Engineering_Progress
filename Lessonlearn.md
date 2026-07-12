# Engineering Progress Tracking & Analytics Dashboard — Lessons Learned & Architecture Guide (`Lessonlearn.md`)

**Project:** Aung Sinkha Development Project Phase 1A (EPC-01) — Z1F Engineering Progress Measurement  
**Version:** 2.1.0 (Multi-Work Package Architecture with Official Summary Sheet Source of Truth & Dynamic Timeline Detection)  
**Date:** July 12, 2026  

---

## 1. Executive Summary & Core Engineering Directive

During the design, verification, and enhancement of the multi-contractor **Engineering Progress Tracking & Analytics Dashboard**, several foundational engineering and architectural challenges were identified, analyzed, and solved. This document formalizes the **Lessons Learned** across three major phases of project execution:
1. **Source of Truth Alignment (`Official Summary Sheets vs. Raw Deliverable Averaging`)**
2. **Dynamic Cut-off Timeline Detection (`Horizontal Schedule Shift Handling`)**
3. **Executive Presentation Architecture (`First Tab Prominence & Dual-Basis Scoping`)**

Adhering strictly to these documented lessons ensures long-term audit compliance, zero data discrepancy with contractor master filings, and a seamless operational workflow for project control engineers (`PJM` / `PTTEPI`).

---

## 2. Lesson 1: Official Summary Sheet Source of Truth vs. Raw Deliverable Averaging

### 2.1 Problem Statement & Risk Analysis
In initial dashboard prototypes, executive progress figures (`Baseline Plan Progress %`, `Achieved Actual Progress %`, `Forecast Expected Progress %`, and `SPI`) were calculated by averaging individual row progress across the ~1,461 deliverable rows inside the Master Document Register (`MDR-Engineering`, `FilteredDEL`, and `EDSR` sheets).

When cross-checked against official contractor executive progress submittals, notable discrepancies emerged:
* **WP-1 Topside Excl. Structure:** Raw deliverable average produced ~38% Actual, whereas the official contractor cut-off report (`00210 EDSR as of WE 27 03 Jul 26.xlsm`) certified **42.09% Actual vs. 36.73% Plan** (`SPI = 1.15`).
* **WP-1 Topside Structure:** Raw calculation differed from the official `Progress Summary` sheet certification of **48.22% Actual vs. 46.70% Plan** (`SPI = 1.03`).
* **WP-1 Jacket:** Raw calculation differed from the official `Progress Summary` sheet certification of **32.59% Actual vs. 27.37% Plan** (`SPI = 1.19`).
* **WP-2 Pipeline:** Raw calculation differed from the official `MDR-Detailed Design-A10` summary rows of **34.75% Actual vs. 35.16% Plan** (`SPI = 0.99`).

**Root Cause:** Contractors calculate overall progress not via simple arithmetic row averaging, but via **weighted man-hour earned-value models, discipline weighting factors, and milestone progress curves** consolidated directly on dedicated summary sheets and summary rows.

### 2.2 Engineering Solution & Architectural Implementation (`REQ-09`)
To eliminate calculation discrepancies and enforce 100% audit alignment with official project records, the ingestion engine (`server.py`) was refactored to implement `extract_official_summary_kpis()`. This function bypasses row averaging for package-level KPIs and directly extracts exact certified metrics from their exact official cell coordinates:

```python
# 1. WP-1 Topside Excl. Structure (EDSR Report Sheet, Row 2788 'Total :')
plan_pct = edsr_sheet["AH2788"].value  # Col 34 -> 36.73%
act_pct  = edsr_sheet["AE2788"].value  # Col 31 -> 42.09%

# 2. WP-1 Topside Structure (Progress Summary Sheet, Row 8 'Overall Topsides DDE Progress')
plan_pct = sum_sheet["O8"].value       # Col 15 -> 46.70%
act_pct  = sum_sheet["Q8"].value       # Col 17 -> 48.22%

# 3. WP-1 Jacket (Progress Summary Sheet, Row 12 'Overall Jacket & Pile DDE Progress')
plan_pct = sum_sheet["O12"].value      # Col 15 -> 27.37%
act_pct  = sum_sheet["Q12"].value      # Col 17 -> 32.59%
```

### 2.3 Operational Takeaway
**Never re-calculate high-level executive progress percentages via simple average across deliverable rows.** Always extract macro progress figures directly from the official summary tables certified by the contractor. Individual deliverable rows (`MDR` / `EDSR`) should be utilized strictly for granular document filtering, delay aging analysis (`Type 3.1` and `Type 3.2`), lookahead warnings, and document status tracking.

---

## 3. Lesson 2: Dynamic Cut-Off Timeline Column Detection for Horizontal Schedules (`WP-2 Pipeline`)

### 3.1 Problem Statement & Risk Analysis
For **WP-2 Pipeline (`Att.2. WP2 Engineering Progress Measurement_030726_Updated.xlsx`)**, progress numbers (`Plan Cumm`, `Actual Cumm`, `Forecast Cumm`) are stored horizontally across weekly timeline columns inside the `MDR-Detailed Design-A10` sheet (`Rows 214, 216, and 218`).

If the dashboard hardcoded a specific column index (such as `Col BW` / `Col 75` for `03 July-2026`), the system would break or show outdated data whenever the project moves to future weekly cut-off dates (e.g., `10 July-2026` in `Col BX`, `17 July-2026` in `Col BY`).

### 3.2 Engineering Solution & Dynamic Algorithm (`REQ-09 Dynamic Engine`)
To future-proof the application against weekly schedule progression without requiring code updates, a **Dynamic Cut-Off Column Detection Algorithm** was engineered in `server.py`:

```python
# Scan across horizontal timeline columns starting from Column 53 (Col BB)
target_col = None
for c in range(53, sheet.max_column + 1):
    val_act = sheet.cell(row=216, column=c).value  # Row 216: Actual Cumm
    if val_act is not None and isinstance(val_act, (int, float)) and val_act > 0:
        target_col = c  # Continues rightward until the last active actual progress column is found
```

* **How it works:** The algorithm iterates from `Column 53` (`BB`) across all timeline columns. It checks `Row 216` (`Actual Cumm`). The rightmost column containing valid numeric actual progress before empty/zero future cells is dynamically locked as the active cut-off date column.
* **Current Verification Result:** For the `03-Jul-26` workbook, the algorithm perfectly identifies `Col 75` (`BW`), extracting:
  * `Plan Progress (Row 214, Col BW):` **35.16%**
  * `Actual Progress (Row 216, Col BW):` **34.75%**
  * `Forecast Progress (Row 218, Col BW):` **34.10%** (`SPI = 0.99`)

### 3.3 Operational Takeaway
Horizontal timeline sheets must always be evaluated dynamically by identifying the rightmost column populated with actual execution progress. This guarantees that when users upload new weekly cut-off Excel files, the dashboard automatically shifts to the latest reporting week without developer intervention.

---

## 4. Lesson 3: First Tab Executive Visibility (`Executive Overview`)

### 4.1 Problem Statement & Risk Analysis
In initial multi-tab layouts, users had to click into individual Work Package tabs (`WP-1 Topside Excl.`, `WP-1 Topside Struct.`, `WP-1 Jacket`, `WP-2 Pipeline`) to inspect progress percentages. Executive management and project directors require immediate, high-level macro visibility the instant the application opens (`REQ-10`).

### 4.2 Engineering Solution (`REQ-10 Executive Overview Tab`)
The first tab (**`Executive Overview`**, `data-wp="executive"`) was structured to serve as the supreme project command center:
1. **Top-Level Overall Progress Section (`progress-section`):**
   * Prominently displays aggregated, weighted progress across all `1,461` project deliverables:
   * **Baseline Plan Progress:** `34.02%`
   * **Achieved Actual Progress:** `38.16%` (`+4.14% variance`)
   * **Forecast Expected Progress:** `33.92%`
   * **Schedule Performance Index:** `SPI = 1.12` (`🚀 Ahead of Schedule`)
2. **Interactive Work Package Summary Table (`wpSummaryTable`):**
   * Displays a side-by-side comparison matrix of all four work packages, allowing executives to immediately pinpoint performance variances and click any row to navigate directly to that package's deep-dive tab:
   * `WP-1 Topside Excl. Structure:` `42.09% Actual vs. 36.73% Plan` (`SPI = 1.15`, `+5.36% variance`)
   * `WP-1 Topside Structure:` `48.22% Actual vs. 46.70% Plan` (`SPI = 1.03`, `+1.52% variance`)
   * `WP-1 Jacket:` `32.59% Actual vs. 27.37% Plan` (`SPI = 1.19`, `+5.22% variance`)
   * `WP-2 Pipeline:` `34.75% Actual vs. 35.16% Plan` (`SPI = 0.99`, `-0.41% variance`)

---

## 5. Lesson 4: Dual-Basis Scoping (`Gate Milestones vs. Unique MDR Documents Baseline`)

### 5.1 Problem Statement & Risk Analysis
A major source of confusion in engineering reporting arises from conflating **Review Gate Milestones (`3 Revisions x Document Count`)** with **Unique Document Inventories (`Master Document Register Basis`)**. 
If a project has `1,461` documents, and each document passes through `IFR (50%)`, `IFA (20%)`, and `AFC (30%)`, the total gate volume is `4,383 gates`. If a dashboard displays `"Total Delayed: 1,859"`, stakeholders may incorrectly assume that 1,859 distinct engineering drawings are delayed, when in fact it represents 1,859 review gate milestones across 639 unique documents (`WP-1 Jacket`).

### 5.2 Engineering Solution (`Dual-Metric KPI Layout & Unique MDR Basis Charting`)
To ensure total clarity across every dashboard view:
1. **Dual-Metric KPI Cards:** Every KPI card (`Total Deliverables`, `On-Time Deliverables`, `Total Delayed Gates`, `Type 3.1 Late`, `Type 3.2 Overdue`, and `14-Day Lookahead`) displays two numbers side-by-side:
   * **Left Column (`Max Gates / Gates Basis`):** Total review gate milestones (`3 Revs per document`).
   * **Right Column (`MDR Rows / Docs Basis`):** Total unique document numbers (`doc_no`) affected by at least one delay breach.
2. **Chart & Badge Standardization:** High-level stacked charts (`Chart 1: Discipline Progress` and `Chart 2: Delay Severity`) and sub-tab count badges (`subCountDelay3_1`, `subCountDelay3_2`, `subCountLookahead`) strictly plot quantities on a **Unique MDR Documents Basis (`doc_no`)**. This prevents chart distortion and aligns 1-to-1 with the base document register.

---

## 6. Lesson 5: Sequential Lookahead Evaluation (`Latest Unsubmitted Revision Basis`)

### 6.1 Problem Statement & Risk Analysis
When evaluating documents for the **14-Day Lookahead Window (`Today < Forecast ≤ Today + 14 days`)**, checking all three gates (`IFR`, `IFA`, `AFC`) simultaneously creates false warnings. For example, if a document has not even been issued for `IFR` yet (`Actual IFR is empty`), flagging its future `AFC` gate as "due in 12 days" violates the sequential nature of engineering execution (`IFR must precede IFA, which must precede AFC`).

### 6.2 Engineering Solution (`Sequential Gate Screening Engine`)
The lookahead engine (`compute_delay_and_lookahead` in `server.py`) strictly isolates the **single latest pending revision right now**:
```python
# Identify the exact pending sequential gate where Actual Date is None
pending_gate = None
if entry["ifr_actual"] is None:
    pending_gate = ("IFR", entry["ifr_forecast"], entry["ifr_plan"])
elif entry["ifa_actual"] is None:
    pending_gate = ("IFA", entry["ifa_forecast"], entry["ifa_plan"])
elif entry["afc_actual"] is None:
    pending_gate = ("AFC", entry["afc_forecast"], entry["afc_plan"])

# Only evaluate that specific pending_gate against the 14-day lookahead window
if pending_gate and today < pending_gate[1] <= today + datetime.timedelta(days=14):
    # Flag lookahead item and evaluate if Forecast > Plan (⚠️ Slipping from Plan)
```
This guarantees that 100% of lookahead alerts are actionable, immediate, and sequence-compliant.

---

## 7. Summary Verification & Checklist for Future Enhancements

Whenever modifying data extraction or UI presentation in future project phases, engineers must verify:
* [x] **Check 1:** Do executive progress numbers (`Plan`, `Actual`, `SPI`) originate directly from official summary tables (`extract_official_summary_kpis()`), overriding row averages? (`Yes`)
* [x] **Check 2:** Does WP-2 Pipeline dynamically locate the rightmost active cut-off date column rather than hardcoding column indices? (`Yes`)
* [x] **Check 3:** Does the First Tab (`Executive Overview`) feature the overall project progress cards and Work Package comparison table immediately at the top? (`Yes`)
* [x] **Check 4:** Are delay charts and sub-tab pill badges standardized to unique `MDR Docs Basis` (`1,461 documents total inventory`)? (`Yes`)
* [x] **Check 5:** Does the 14-day lookahead engine evaluate strictly the single latest pending sequential revision? (`Yes`)
