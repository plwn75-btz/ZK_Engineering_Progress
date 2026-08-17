# Aung Sinkha Development Project Phase 1A (EPC-01)
## Engineering Progress Tracking & Analytics Dashboard — Technical & Operational Handoff Report (`Handoff.md`)

**Project:** Z1F Engineering Progress Measurement  
**Document Version:** 3.0.0 (14-Aug Cut-off with Dual Gate Final Approval Lifecycle)  
**Date:** August 18, 2026  

---

## 1. Conceptual Overview

### 1.1 Background & Purpose
In complex Offshore & Onshore Engineering, Procurement, and Construction (EPC) projects, tracking engineering deliverables across thousands of individual documents, multiple Work Packages (WPs), and diverse engineering disciplines is critical to preventing project slippage. 

The **Engineering Progress Tracking & Analytics Dashboard** replaces static, manual spreadsheet reporting with an automated, high-performance, real-time web application. It ingests weekly Master Document Register (MDR) Excel cut-offs from project contractors across 4 work packages and transforms raw tabular records into actionable executive insights, S-Curve variance analysis, granular discipline delays, and proactive 14-day lookahead warning alerts.

### 1.2 Active File Mapping (14-Aug-2026 Cut-off)
1. **WP-1 Topside Excl. Structure (`wp1_topside_excl`):** `00210 EDSR as of WE 33 14 Aug 26.xlsm` (458 Deliverables)
2. **WP-1 Topside Structure (`wp1_topside_structure`):** `Z1F Topsides EDSR cut-off_14Aug26.xlsx` (237 Deliverables)
3. **WP-1 Jacket (`wp1_jacket`):** `Z1F Jacket EDSR cut-off_14Aug26.xlsx` (653 Deliverables)
4. **WP-2 Pipeline (`wp2_pipeline`):** `Att.2. WP2 Engineering Progress Measurement_030726_Updated.xlsx` (143 Deliverables)
**Total Master Deliverables Inventory: 1,491 Deliverables**

---

## 2. Document Status Lifecycle Specifications

A document's current active lifecycle status is dynamically computed based on the user's defined business rules:

1. **`COMPLETE` (Final Approved / 100% Gate Closed):**
   * *00210 EDSR:* `AP (100%)` Achieved date is filled with date (`63 completed`).
   * *Z1F Topsides & Jacket EDSR:* `AFC App. (20%)` Actual date is filled with date (`36 Topsides, 48 Jacket`).
   * *WP-2 Pipeline:* `AFC` Actual date is filled with date (`6 completed`).
   * **Project Total: 153 Complete Deliverables**
2. **`PENDING FINAL APPROVAL` (Submitted for Final Approval / 10% Gate Closed, 20% Approval Pending):**
   * *00210 EDSR:* `AFC` Achieved date is filled with date **AND** `AP (100%)` Achieved date is blank/empty (`71 pending`).
   * *Z1F Topsides & Jacket EDSR:* `AFC Sub (10%)` Actual date is filled with date **AND** `AFC App. (20%)` is blank/empty (`11 Topsides, 55 Jacket`).
   * **Project Total: 137 Pending Final Approval Deliverables**
3. **`IFA Submitted` (Issued for Approval Review):**
   * `IFA` Actual date is filled with date, **AND** later gates are blank/empty (`420 deliverables`).
4. **`IFR Submitted` (Issued for Interdisciplinary Review):**
   * `IFR` Actual date is filled with date, **AND** later gates are blank/empty (`411 deliverables`).
5. **`Not Yet Submitted`:**
   * No actual submission date exists for any review milestone (`370 deliverables`).

---

## 3. Verified Official Progress Metrics (14-Aug Cut-off)

| Work Package | Deliverables | Official Plan % | Official Actual % | SPI | Status Summary |
| :--- | :--- | :--- | :--- | :--- | :--- |
| **WP-1 Topside Excl. Structure** | 458 | 48.95% | 58.19% | 1.19 | Complete: 63, Pending App: 71, IFA: 73, IFR: 149, Not Sub: 102 |
| **WP-1 Topside Structure** | 237 | 67.56% | 67.21% | 0.99 | Complete: 36, Pending App: 11, IFA: 102, IFR: 51, Not Sub: 37 |
| **WP-1 Jacket** | 653 | 47.77% | 49.81% | 1.04 | Complete: 48, Pending App: 55, IFA: 212, IFR: 175, Not Sub: 163 |
| **WP-2 Pipeline** | 143 | 35.16% | 34.75% | 0.99 | Complete: 6, Pending App: 0, IFA: 33, IFR: 36, Not Sub: 68 |
| **Executive Overall Weighted** | **1,491** | **52.68%** | **55.77%** | **1.06** | **Complete: 153, Pending App: 137, IFA: 420, IFR: 411, Not Sub: 370** |

---

## 4. Delay & Lookahead Classification Engine

Every deliverable gate milestone is evaluated against active tracking date `TODAY`:

$$\text{Reference Date} = \begin{cases} \text{Forecast Date} & \text{if Forecast Date is valid date} \\ \text{Plan Date} & \text{otherwise} \end{cases}$$

### 4.1 Delay Definitions
* **Type 3.1: Submitted Late (`#b45309` Warm Copper/Brown):**
  $$\text{Actual Date is present} \land \text{Actual Date} > \text{Reference Date}$$
  $$\text{Days Delayed} = \text{Actual Date} - \text{Reference Date}$$
* **Type 3.2: Overdue & Not Submitted (`#ef4444` Vibrant Red):**
  $$\text{Actual Date is empty} \land \text{TODAY} \ge \text{Reference Date}$$
  $$\text{Days Overdue} = \text{TODAY} - \text{Reference Date}$$
* **Critical Aging Flag (`Urgency: Critical`):**
  $$\text{Days Delayed / Overdue} > 30 \text{ days}$$

### 4.2 14-Day Lookahead Warning Rules (`#eab308` Crisp Yellow)
* Evaluates strictly the **next immediate unsubmitted milestone gate**.
* If document is already fully approved (`COMPLETE`), lookahead is bypassed.
* Triggers when:
  $$\text{TODAY} < \text{Reference Date} \le \text{TODAY} + 14 \text{ days}$$

---

## 5. UI Architecture & Visualization Specifications

### 5.1 Discipline Stacked Bar Chart (`Chart 1`)
Displays deliverable counts stacked by status across disciplines:
* **COMPLETE:** `#10b981` (Emerald Green)
* **PENDING FINAL APPROVAL:** `#8b5cf6` (Royal Purple)
* **IFA Submitted:** `#0284c7` (Sky Blue)
* **IFR Submitted:** `#06b6d4` (Cyan)
* **Not Yet Submitted:** `#64748b` (Slate Grey)

### 5.2 Delay & Lookahead Grouped Bar Chart (`Chart 2`)
Displays unique document counts affected by:
* **Type 3.1 Late Documents:** `#b45309` (Warm Copper)
* **Type 3.2 Overdue Documents:** `#ef4444` (Red)
* **14-Day Lookahead Risk Documents:** `#eab308` (Yellow)

### 5.3 Sub-Table Section Tabs
1. **Master Document Register (MDR):** Full deliverable register with search, discipline, milestone, and document status filters. Includes column headers for `IFR`, `IFA`, `AFC Sub (10%)`, and `AFC App (20% / AP)`.
2. **Delayed & Overdue Deliverables:** Full list of Type 3.1 and Type 3.2 delay breaches.
3. **14-Day Lookahead Warnings:** Upcoming milestones due within 14 days, highlighting plan slippage (`Forecast > Plan`).

---

## 6. Operational Verification & Server Startup

### 6.1 Server Startup
```powershell
# Start the Python Threaded HTTP Server (Default Port: 8090)
python server.py
```
Open [`http://localhost:8090`](http://localhost:8090) in your web browser.

### 6.2 Running Automated Verification Script
```powershell
python verify.py
```
The script validates all 1,491 deliverables across all 4 work packages, confirms zero classification errors for Type 3.1/3.2 delays, and verifies exact status counts for `COMPLETE` and `PENDING FINAL APPROVAL`.

---

## 7. Summary of Deliverables Handoff Checklist
* [x] **`backup_original/`**: Preserved original versions of all source code files before 14-Aug execution.
* [x] **`uploads_mapping.json`**: Updated to map the 3 new 14-Aug-2026 cut-off files.
* [x] **`server.py`**: Refactored with stream-optimized ingestion, dual-gate `COMPLETE` and `PENDING FINAL APPROVAL` logic, and certified summary progress extraction.
* [x] **`styles.css`**: Enhanced with `.badge-complete`, `.badge-pending-approval`, `.badge-ifa-sub`, `.badge-ifr-sub` styles.
* [x] **`index.html`**: Updated dropdown filter options and dual-gate table structure.
* [x] **`app.js`**: Updated stacked bar charts, MDR table rendering, and filtering engine.
* [x] **`verify.py`**: Automated audit script testing data ingestion, delay classifications, and status accuracy.
* [x] **`Spec.md`**, **`Lessonlearn.md`**, **`Handoff.md`**: Complete updated technical and operational documentation.
