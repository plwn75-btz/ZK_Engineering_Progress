# Engineering Progress Tracking & Analytics Dashboard — Lessons Learned & Architecture Guide (`Lessonlearn.md`)

**Project:** Aung Sinkha Development Project Phase 1A (EPC-01) — Z1F Engineering Progress Measurement  
**Version:** 3.0.0 (14-Aug Cut-off with Dual Gate Final Approval Lifecycle & Stream-Optimized Ingestion)  
**Date:** August 18, 2026  

---

## 1. Executive Summary & Core Engineering Directive

During the design, verification, and enhancement of the multi-contractor **Engineering Progress Tracking & Analytics Dashboard**, several foundational engineering and architectural challenges were identified, analyzed, and solved. This document formalizes the **Lessons Learned** across key phases of project execution:
1. **Dual Gate Final Approval Lifecycle (`COMPLETE` vs. `PENDING FINAL APPROVAL`)**
2. **Official Summary Sheet Source of Truth (`Macro vs. Row Averaging`)**
3. **OpenPyXL Stream-Optimized Ingestion in Read-Only Mode**
4. **Dynamic Cut-off Timeline Detection (`Horizontal Schedule Shift Handling`)**
5. **Dual-Basis Scoping (`Gate Milestones vs. Unique MDR Documents Baseline`)**
6. **Sequential Lookahead Evaluation (`Latest Unsubmitted Revision Basis`)**

Adhering strictly to these documented lessons ensures long-term audit compliance, zero data discrepancy with contractor master filings, and a seamless operational workflow for project control engineers (`PJM` / `PTTEPI`).

---

## 2. Lesson 1: Dual Gate Final Approval Lifecycle (`COMPLETE` vs. `PENDING FINAL APPROVAL`)

### 2.1 Problem Statement & Risk Analysis
In EPC engineering execution, a deliverable submitted as `AFC` (Approved for Construction) undergoes two distinct milestone gates:
1. **Contractor Final Issue (`AFC Sub / 10% Milestone` or `AFC Achieved`):** Contractor submits the final deliverable for client endorsement.
2. **Client Final Approval (`AFC App / 20% Milestone` or `AP 100% Final Gate`):** Client formally signs off and approves the document (`100% Gate Closed`).

Treating any submission of `AFC` as immediately "Complete" causes premature project closure tracking. Management must clearly see which documents are fully approved (`COMPLETE`) versus those sitting in client review pending final sign-off (`PENDING FINAL APPROVAL`).

### 2.2 Engineering Solution & Implementation
1. **`COMPLETE` (Final Approved / 100% Gate Closed):**
   * *00210 EDSR (Topside Excl. Structure):* `AP (100%)` Achieved date is filled with a valid completion date in `Gates` sheet (`63 completed`).
   * *Z1F Topsides & Jacket EDSR:* `AFC App. (20%)` Actual date is filled with a valid completion date (`36 Topsides, 48 Jacket`).
   * *WP-2 Pipeline:* `AFC` Actual date is filled with date (`6 completed`).
   * **Project Total: 153 Complete Deliverables**
2. **`PENDING FINAL APPROVAL` (10% Submitted, 20% Approval Pending):**
   * *00210 EDSR:* `AFC` Achieved date is filled with date **AND** `AP (100%)` Achieved date is blank/empty (`71 pending`).
   * *Z1F Topsides & Jacket EDSR:* `AFC Sub (10%)` Actual date is filled with date **AND** `AFC App. (20%)` is blank/empty (`11 Topsides, 55 Jacket`).
   * **Project Total: 137 Pending Final Approval Deliverables**

---

## 3. Lesson 2: Official Summary Sheet Source of Truth vs. Raw Deliverable Averaging

### 3.1 Problem Statement & Risk Analysis
In initial dashboard prototypes, executive progress figures (`Baseline Plan Progress %`, `Achieved Actual Progress %`, `Forecast Expected Progress %`, and `SPI`) were calculated by averaging individual row progress across deliverable rows inside the Master Document Register.

When cross-checked against official contractor cut-offs, notable discrepancies emerged:
* **WP-1 Topside Excl. Structure (14-Aug):** Official cut-off (`00210 EDSR as of WE 33 14 Aug 26.xlsm`, Row 2816) certified **58.19% Actual vs. 48.95% Plan** (`SPI = 1.19`).
* **WP-1 Topside Structure (14-Aug):** Official `Progress Summary` sheet (Row 8) certified **67.21% Actual vs. 67.56% Plan** (`SPI = 0.99`).
* **WP-1 Jacket (14-Aug):** Official `Progress Summary` sheet (Row 12) certified **49.81% Actual vs. 47.77% Plan** (`SPI = 1.04`).
* **WP-2 Pipeline:** Official timeline columns certified **34.75% Actual vs. 35.16% Plan** (`SPI = 0.99`).

**Root Cause:** Contractors calculate overall progress via weighted man-hour earned-value models, discipline weighting factors, and milestone progress curves consolidated on certified summary sheets.

### 3.2 Engineering Solution
The ingestion engine (`server.py`) implements `extract_official_summary_kpis()` to extract certified macro figures directly from designated summary sheets, ensuring 100% audit alignment.

---

## 4. Lesson 3: OpenPyXL Stream-Optimized Ingestion in Read-Only Mode

### 4.1 Problem Statement
When ingesting large multi-sheet `.xlsm` workbooks (`00210 EDSR as of WE 33 14 Aug 26.xlsm` is ~25MB with 2,800+ rows and multiple complex sheets), loading in standard DOM mode consumes 500MB+ RAM and takes over 60 seconds. However, switching to `read_only=True` and calling `sheet.cell(row=r, column=c)` causes the openpyxl XML parser to rewind from row 1 on every access, causing the server process to freeze or hang.

### 4.2 Engineering Solution
In `server.py`, all `read_only=True` workbooks use sequential `ws.iter_rows(values_only=True)`:
```python
wb = openpyxl.load_workbook(filepath, data_only=True, read_only=True)
for row in ws.iter_rows(min_row=4, values_only=True):
    # Process row in single forward pass
```
This reduced full 4-workbook ingestion time from 65+ seconds down to **<19 seconds**, with zero memory bloat.

---

## 5. Lesson 4: Dynamic Cut-Off Timeline Column Detection for Horizontal Schedules (`WP-2 Pipeline`)

For **WP-2 Pipeline (`Att.2. WP2 Engineering Progress Measurement_030726_Updated.xlsx`)**, progress numbers (`Plan Cumm`, `Actual Cumm`, `Forecast Cumm`) are stored horizontally across weekly timeline columns.

The ingestion engine scans timeline columns dynamically starting from Column 53 to locate the rightmost populated `Actual Cumm` value. This guarantees that when users upload new weekly cut-off Excel files, the dashboard automatically shifts to the latest reporting week without developer intervention.

---

## 6. Lesson 5: Dual-Basis Scoping (`Gate Milestones vs. Unique MDR Documents Baseline`)

Every KPI card (`Total Deliverables`, `On-Time Deliverables`, `Total Delayed Gates`, `Type 3.1 Late`, `Type 3.2 Overdue`, and `14-Day Lookahead`) displays two numbers side-by-side:
* **Left Column (`Max Gates / Gates Basis`):** Total review gate milestones (`3-4 Revs per document`).
* **Right Column (`MDR Rows / Docs Basis`):** Total unique document numbers (`doc_no`) affected by at least one delay breach.
* High-level stacked charts (`Chart 1: Discipline Progress` and `Chart 2: Delay Severity`) and sub-tab count badges strictly plot quantities on a **Unique MDR Documents Basis (`1,491 total inventory`)**.

---

## 7. Lesson 6: Sequential Lookahead Evaluation (`Latest Unsubmitted Revision Basis`)

Engineering progress tracking is strictly sequential (`IFR -> IFA -> AFC Sub -> AFC App`). The lookahead engine (`compute_delay_and_lookahead`) strictly isolates the **single latest pending revision right now** and evaluates if its reference date falls within `TODAY < Forecast <= TODAY + 14 days`. Completed documents (`COMPLETE`) are 100% excluded from lookahead warnings.
