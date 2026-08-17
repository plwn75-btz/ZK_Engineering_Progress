# Engineering Progress Tracking & Analytics Dashboard — Technical Specification (Spec.md)

**Project : Z1F Engineering Progress Measurement.**  
**Document Version:** 3.0.0 (14-Aug Cut-off with Final Approval Lifecycle & Dual Gate Tracking)  
**Status:** Approved for Implementation  

---

## 1. Executive Summary & Core Concept

In multi-contractor Offshore & Onshore Engineering, Procurement, and Construction (EPC) projects, tracking engineering progress across distinct work packages requires unifying disparate reporting formats into a single, cohesive analytics dashboard. 

The **Engineering Progress Tracking & Analytics Dashboard** automatically ingests, parses, normalizes, and analyzes weekly Master Document Register (MDR) and Engineering Deliverable Status Report (EDSR) Excel cut-offs from contractors across four major engineering domains:
1. **WP-1 Topside Excl. Structure** (`00210 EDSR as of WE 33 14 Aug 26.xlsm` — 458 Deliverables)
2. **WP-1 Topside Structure** (`Z1F Topsides EDSR cut-off_14Aug26.xlsx` — 237 Deliverables)
3. **WP-1 Jacket** (`Z1F Jacket EDSR cut-off_14Aug26.xlsx` — 653 Deliverables)
4. **WP-2 Pipeline** (`Att.2. WP2 Engineering Progress Measurement_030726_Updated.xlsx` — 143 Deliverables)
**Total Master Deliverables Inventory: 1,491 Deliverables**

The dashboard replaces manual spreadsheets with a high-performance, dark-themed glassmorphism web application providing real-time executive progress tracking, discipline slippage analysis, Schedule Performance Index (SPI) curves, and automated delay/lookahead alerts anchored to `TODAY` vs. `FORECAST` tracking dates.

Crucially, to ensure 100% audit compliance with official contractor reports, the dashboard extracts all Work Package macro progress percentages (`Plan %`, `Forecast %`, `Actual %`, and `SPI`) directly from the **Official Summary Sheets & Dynamic Cut-off Timeline Columns** in each Excel workbook, overriding deliverable-level row approximations.

---

## 2. Document Status Lifecycle Specification (`COMPLETE` & `PENDING FINAL APPROVAL`)

Deliverable lifecycle tracking distinguishes between initial contractor issue, intermediate review, final contractor submission, and final client approval (`COMPLETE`):

### 2.1 Status Hierarchy & Assignment Rules
1. **`COMPLETE` (Final Approved / 100% Gate Closed):**
   * *00210 EDSR (Topside Excl. Structure):* `AP` (`DocStatusID == 'AP'`) Achieved date is filled in `Gates` sheet.
   * *Z1F Topsides & Jacket EDSR:* `AFC App. (20%)` Actual date is filled with a valid completion date.
   * *WP-2 Pipeline:* `AFC` Actual date is filled with a valid completion date.
2. **`PENDING FINAL APPROVAL` (Submitted for Final Approval / 10% Gate Closed, 20% Approval Pending):**
   * *00210 EDSR (Topside Excl. Structure):* `AFC` Achieved date is filled with date, **AND** `AP (100%)` Achieved date is blank/empty.
   * *Z1F Topsides & Jacket EDSR:* `AFC Sub (10%)` Actual date is filled with date, **AND** `AFC App. (20%)` is blank/empty.
   * *WP-2 Pipeline:* `AFC` Actual date is filled (or marked COMPLETE).
3. **`IFA Submitted` (Issued for Approval Review):**
   * `IFA` Actual date is filled with date, **AND** later gates (`AFC Sub`, `AFC App`, `AP`) are blank/empty.
4. **`IFR Submitted` (Issued for Interdisciplinary Review):**
   * `IFR` Actual date is filled with date, **AND** later gates (`IFA`, `AFC Sub`, `AFC App`, `AP`) are blank/empty.
5. **`Not Yet Submitted`:**
   * No actual submission date exists for any review milestone.

---

## 3. Ingestion Engine Specification & Excel Mapping Schema

The backend ingestion engine (`server.py`) handles structural variations across all source Excel files without requiring manual modifications.

### 3.1 File 1: WP-1 Topside Excl. Structure (`00210 EDSR as of WE 33 14 Aug 26.xlsm`)
* **Primary Deliverable Source (`FilteredDEL` / `DEL` Sheet):**
  * `Doc No`: Column `DELID` (Col A)
  * `Title`: Column `DELName` (Col B)
  * `Discipline`: Column `CAName` (Col C) / `CTIDName` (Col E)
* **Milestone Dates Source (`Gates` Sheet):**
  * Relational records (`1 to N` per `DELID`).
  * Gates parsed: `IFR`, `IFA`, `AFC` (10% milestone), `AP` (100% Final Approval gate).
  * `Plan Date`: Column K (`Planned Finish`)
  * `Forecast Date`: Column O (`Forecast Finish`, fallback to Col K)
  * `Actual Date`: Column M (`Actual Finish`)
* **Official Progress Extraction (`EDSR Report` Sheet):**
  * Row 2816 (`Total :`): Column S (Actual Progress % = 58.19%), Column K (Planned Progress % = 48.95%), SPI = 1.19.

### 3.2 File 2: WP-1 Topside Structure (`Z1F Topsides EDSR cut-off_14Aug26.xlsx`)
* **Primary Deliverable Source (`EDSR` Sheet):**
  * Headers on Rows 1 to 3. Data starts at **Row 4**.
  * `Doc No`: Column 4 (0-index 3, `Deliverable Title`)
  * `Title`: Column 5 (0-index 4, `Description`)
  * `Discipline`: Column 2 (0-index 1, `Work Package`)
* **Milestone Dates Mapping (0-indexed):**
  * **IFR Gate:** Plan = Col 10, Forecast = Col 11, Actual = Col 12
  * **IFA Gate:** Plan = Col 13, Forecast = Col 14, Actual = Col 15
  * **AFC Sub (10%) Gate:** Plan = Col 16, Forecast = Col 17, Actual = Col 18
  * **AFC App (20%) Gate:** Plan = Col 19, Forecast = Col 20, Actual = Col 21
* **Official Progress Extraction (`Progress Summary` Sheet):**
  * Row 8 (`Overall Topsides`): Planned % = 67.56%, Actual % = 67.21%, SPI = 0.99.

### 3.3 File 3: WP-1 Jacket (`Z1F Jacket EDSR cut-off_14Aug26.xlsx`)
* **Primary Deliverable Source (`EDSR` Sheet):**
  * Headers on Rows 1 to 3. Data starts at **Row 4**.
  * `Doc No`: Column 3 (0-index 2, `Deliverable Title`)
  * `Title`: Column 4 (0-index 3, `Description`)
  * `Discipline`: Column 2 (0-index 1, `Work Package`)
* **Milestone Dates Mapping (0-indexed):**
  * **IFR Gate:** Plan = Col 9, Forecast = Col 10, Actual = Col 11
  * **IFA Gate:** Plan = Col 12, Forecast = Col 13, Actual = Col 14
  * **AFC Sub (10%) Gate:** Plan = Col 15, Forecast = Col 16, Actual = Col 17
  * **AFC App (20%) Gate:** Plan = Col 18, Forecast = Col 19, Actual = Col 20
* **Official Progress Extraction (`Progress Summary` Sheet):**
  * Row 12 (`Overall Jacket & Pile DDE Progress`): Planned % = 47.77%, Actual % = 49.81%, SPI = 1.04.

### 3.4 File 4: WP-2 Pipeline (`Att.2. WP2 Engineering Progress Measurement_030726_Updated.xlsx`)
* **Primary Deliverable Source (`MDR` Sheet):**
  * Data starts at **Row 7**.
  * `Doc No`: Column C (`DOCUMENT NUMBER`)
  * `Title`: Column D (`DOCUMENT TITLE`)
  * `Discipline`: Column B (`DISCIPLINE`)
* **Milestone Dates Mapping:**
  * **IFR Gate:** Plan = Col E, Forecast = Col F, Actual = Col G
  * **IFA Gate:** Plan = Col I, Forecast = Col J, Actual = Col K
  * **AFC Gate:** Plan = Col M, Forecast = Col N, Actual = Col O
* **Official Progress Extraction (`S-Curve` Sheet):**
  * Active Column 75 (`BW`): Planned % = 35.16%, Actual % = 34.75%, Forecast % = 34.10%, SPI = 0.99.

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

### 5.1 Dedicated Status Summary Boxes (Between Row 1 & Row 2)
Positioned directly in between **Row 1** (Deliverables & Delay Breaches) and **Row 2** (Plan, Actual, Forecast Progress & SPI):
1. **PENDING FINAL APPROVAL Box (`#8b5cf6` Purple):**
   * Primary metric: Total documents pending client approval (`137 Docs` in Executive).
   * Category breakdown format: **`ENG XX / TBE XX / MR XX`**
2. **COMPLETE Box (`#10b981` Emerald):**
   * Primary metric: Total documents with closed 100% final approval gates (`153 Docs (10.3%)` in Executive).
   * Category breakdown format: **`ENG XX (YY%) / TBE XX (YY%) / MR XX (YY%)`** where `YY%` is `% complete of each type of document`.

### 5.2 Discipline Stacked Bar Chart (`Chart 1`)
Displays deliverable counts stacked by status across disciplines:
* **COMPLETE:** `#10b981` (Emerald Green)
* **PENDING FINAL APPROVAL:** `#8b5cf6` (Royal Purple)
* **IFA Submitted:** `#0284c7` (Sky Blue)
* **IFR Submitted:** `#06b6d4` (Cyan)
* **Not Yet Submitted:** `#64748b` (Slate Grey)

### 5.3 Delay & Lookahead Grouped Bar Chart (`Chart 2`)
Displays unique document counts affected by:
* **Type 3.1 Late Documents:** `#b45309` (Warm Copper)
* **Type 3.2 Overdue Documents:** `#ef4444` (Red)
* **14-Day Lookahead Risk Documents:** `#eab308` (Yellow)

### 5.3 Filter Bar Dropdowns
* **Gate Milestone:** `All`, `IFR`, `IFA`, `AFC Sub`, `AFC App`
* **Document / Breach Status:** `All`, `COMPLETE`, `PENDING FINAL APPROVAL`, `IFA Submitted`, `IFR Submitted`, `Not Yet Submitted`, `On-Time`, `Delayed`, `Type 3.1`, `Type 3.2`, `Critical Aging`, `14-Day Lookahead`

---

## 6. Verification and Audit Metrics (14-Aug Cut-off)

| Work Package | Deliverables | Official Plan % | Official Actual % | SPI | Status Summary |
| :--- | :--- | :--- | :--- | :--- | :--- |
| **WP-1 Topside Excl. Structure** | 458 | 48.95% | 58.19% | 1.19 | Complete: 63, Pending App: 71, IFA: 73, IFR: 149, Not Sub: 102 |
| **WP-1 Topside Structure** | 237 | 67.56% | 67.21% | 0.99 | Complete: 36, Pending App: 11, IFA: 102, IFR: 51, Not Sub: 37 |
| **WP-1 Jacket** | 653 | 47.77% | 49.81% | 1.04 | Complete: 48, Pending App: 55, IFA: 212, IFR: 175, Not Sub: 163 |
| **WP-2 Pipeline** | 143 | 35.16% | 34.75% | 0.99 | Complete: 6, Pending App: 0, IFA: 33, IFR: 36, Not Sub: 68 |
| **Executive Overall Weighted** | **1,491** | **52.68%** | **55.77%** | **1.06** | **Complete: 153, Pending App: 137, IFA: 420, IFR: 411, Not Sub: 370** |
