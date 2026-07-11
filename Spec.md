# Engineering Progress Tracking & Analytics Dashboard — Technical Specification (Spec.md)

**Project : Z1F Engineering Progress Measurement.**  
**Document Version:** 2.1.0 (Multi-Work Package Architecture with Weekly Upload Scope)  
**Status:** Approved for Implementation  

---

## 1. Executive Summary & Core Concept

In multi-contractor Offshore & Onshore Engineering, Procurement, and Construction (EPC) projects, tracking engineering progress across distinct work packages requires unifying disparate reporting formats into a single, cohesive analytics dashboard. 

The **Engineering Progress Tracking & Analytics Dashboard** automatically ingests, parses, normalizes, and analyzes weekly Master Document Register (MDR) and Engineering Deliverable Status Report (EDSR) Excel cut-offs from contractors across four major engineering domains:
1. **WP-1 Topside Excl. Structure** (`00210 EDSR as of WE 27 03 Jul 26.xlsm`)
2. **WP-1 Topside Structure** (`Z1F Topsides EDSR cut-off_3Jul26.xlsx`)
3. **WP-1 Jacket** (`Z1F Jacket EDSR cut-off_3Jul26.xlsx`)
4. **WP-2 Pipeline** (`Att.2. WP2 Engineering Progress Measurement_030726_Updated.xlsx`)

The dashboard replaces manual spreadsheets with a high-performance, dark-themed glassmorphism web application providing real-time executive progress tracking, discipline slippage analysis, Schedule Performance Index (SPI) curves, and automated delay/lookahead alerts anchored to `TODAY` vs. `FORECAST` tracking dates.

---

## 2. Ingestion Engine Specification & Excel Mapping Schema

The backend ingestion engine (`server.py`) must handle the exact structural variations across the 4 source Excel files without requiring manual file modifications.

### 2.1 File 1: WP-1 Topside Excl. Structure (`00210 EDSR as of WE 27 03 Jul 26.xlsm`)
* **Primary Deliverable Source (`FilteredDEL` / `DEL` Sheet):**
  * `Doc No`: Column `DELID` (Col A)
  * `Title`: Column `DELName` (Col B)
  * `Discipline`: Derived from `CAName` / `CTIDName` (e.g., `Electrical`, `Mechanical`, `Process`)
* **Milestone Dates Source (`Gates` Sheet):**
  * Records are relational (`1 to N` per `DELID`).
  * Filter rows where `DocStatusID` $\in$ `['IFR', 'IFA', 'AFC']` (or `IFI`, `IDC` where applicable).
  * `Plan Date`: Column K (`Planned Finish`)
  * `Forecast Date`: Column O (`Forecast Finish`, falling back to Column K if `'-'` or empty)
  * `Actual Date`: Column M (`Actual Finish`, parsed if valid datetime/date string, empty if `'-'` or `None`)

### 2.2 File 2: WP-1 Topside Structure (`Z1F Topsides EDSR cut-off_3Jul26.xlsx`)
* **Primary Deliverable Source (`EDSR` Sheet):**
  * Exact structural match with `WP-1 Jacket` (`EDSR -TOPSIDES`).
  * Headers on Rows 1 to 3. Data starts at **Row 4**.
  * `Doc No`: Column D (`Deliverable Title` / Document Number)
  * `Title`: Column E (`Description`)
  * `Discipline`: Column B (`Work Package` / Discipline Code)
* **Milestone Dates Mapping (Columns 1-indexed):**
  * **IFR (50%) Gate:** `Plan Date` = Col 11, `Forecast Date` = Col 12, `Actual Date` = Col 13
  * **IFA (20%) Gate:** `Plan Date` = Col 14, `Forecast Date` = Col 15, `Actual Date` = Col 16
  * **AFC Sub / App Gate:** `Plan Date` = Col 17/20, `Forecast Date` = Col 18/21, `Actual Date` = Col 19/22

### 2.3 File 3: WP-1 Jacket (`Z1F Jacket EDSR cut-off_3Jul26.xlsx`)
* **Primary Deliverable Source (`EDSR` Sheet):**
  * Headers on Rows 1 to 3. Data starts at **Row 4**.
  * `Doc No`: Column D (`Deliverable Title` / Document Number)
  * `Title`: Column E (`Description`)
  * `Discipline`: Column B (`Work Package` / Discipline Code)
* **Milestone Dates Mapping (Columns 1-indexed):**
  * **IFR (50%) Gate:** `Plan Date` = Col 11, `Forecast Date` = Col 12, `Actual Date` = Col 13
  * **IFA (20%) Gate:** `Plan Date` = Col 14, `Forecast Date` = Col 15, `Actual Date` = Col 16
  * **AFC Sub / App Gate:** `Plan Date` = Col 17/20, `Forecast Date` = Col 18/21, `Actual Date` = Col 19/22

### 2.4 File 4: WP-2 Pipeline (`Att.2. WP2 Engineering Progress Measurement_030726_Updated.xlsx`)
* **Primary Deliverable Source (`MDR-Detailed Design-A10` Sheet):**
  * Data rows start from **Row 11** (Headers on Rows 9 & 10).
  * `Doc No`: Column I (`Document No.`)
  * `Title`: Column J (`Document Title`)
  * `Discipline`: Column D (`Disc`) or Column C (`Area`)
* **Milestone Dates Mapping (Columns 1-indexed):**
  * **IFR Gate:** `Plan Date` = Col 21 (`PLAN`), `Actual Date` = Col 22 (`ACTUAL`), `Forecast Date` = Col 23 (`Forecast`)
  * **IFA Gate:** `Plan Date` = Col 30 (`PLAN`), `Actual Date` = Col 31 (`ACTUAL`), `Forecast Date` = Col 32 (`Forecast`)
  * **AFC Gate:** `Plan Date` = Col 37 (`PLAN`), `Actual Date` = Col 38 (`ACTUAL`), `Forecast Date` = Col 39/40 (`Forecast`)

---

## 3. Weekly Manual Upload & Dynamic Ingestion Scope (`USER Upload Workflow`)

Per user operational requirements, the **USER will update these 4 Excel files on a weekly basis**. Contractors deliver weekly progress files named with varying date, month, and year conventions (e.g., `00210 EDSR as of WE 27 10 Jul 26.xlsm`, `Z1F Topsides EDSR cut-off_10Jul26.xlsx`).

To ensure maximum user-friendliness and zero technical friction when updating weekly data:
1. **Interactive Weekly Upload Center Modal:**
   * Clicking the header button `📁 Weekly Cut-off Upload` opens a user-friendly drag-and-drop upload center.
   * Users can upload individual cut-off files or drag-and-drop all 4 weekly Excel files simultaneously.
2. **Intelligent Auto-Detection & Assignment (`server.py` & `app.js`):**
   * When files are uploaded or dropped into the workspace folder, the server inspects the filename keywords (`00210`/`Excl`, `Topsides`, `Jacket`, `WP2`/`Pipeline`) and internal sheet names (`FilteredDEL`, `EDSR`, `MDR-Detailed Design-A10`) to **automatically map each uploaded file to its exact Work Package**.
   * If a filename is ambiguous, the upload modal provides simple dropdown selectors allowing the user to confirm:
     * `[File: 00210 EDSR...xlsm] -> Assign to: WP-1 Topside Excl. Structure`
     * `[File: Z1F Topsides...xlsx] -> Assign to: WP-1 Topside Structure`
     * `[File: Z1F Jacket...xlsx] -> Assign to: WP-1 Jacket`
     * `[File: Att.2. WP2...xlsx] -> Assign to: WP-2 Pipeline`
3. **Instant Cache Refresh & Live Notification:**
   * Upon upload, `server.py` immediately saves the file, clears the thread-safe memory cache, re-parses the new weekly records, and updates the dashboard instantly without requiring a server restart.
   * A clean toast notification confirms: *"✅ Successfully processed [File Name] for [Work Package] as of [Cut-off Date]"*.

---

## 4. Delay & Lookahead Mathematical Algorithms

All progress and slippage evaluations strictly adhere to the rules established in `HANDOFF_REPORT.md` and project quality standards.

### 4.1 Reference Tracking Date Hierarchy
For any engineering document and review gate ($g \in \{\text{IFR}, \text{IFA}, \text{AFC}\}$):
$$\text{RefDate}_g = \begin{cases} \text{ForecastDate}_g & \text{if } \text{ForecastDate}_g \text{ exists and is a valid date} \\ \text{PlanDate}_g & \text{otherwise (fallback)} \end{cases}$$

### 4.2 Universal Anchor
Every evaluation compares exact date deltas against:
$$\text{Anchor Date} = \text{TODAY} \quad (\text{e.g., } \texttt{datetime.date.today()})$$

### 4.3 Two-Pronged Delay Classification (`REQ-03` & `REQ-04`)
A document gate is classified as **Delayed** under two distinct mathematical criteria:

#### 1. Type 3.1 Delay (Submitted Late)
* **Condition:** The actual submission occurred after the baseline forecast date.
  $$\text{ActualDate}_g \neq \text{None} \quad \land \quad \text{ActualDate}_g > \text{RefDate}_g$$
* **Calculation:**
  $$\text{Delay Days} = \max((\text{ActualDate}_g - \text{RefDate}_g).\text{days}, 1)$$
* **UI Indicator:** 🟠 Orange Badge (`Type 3.1: Submitted Late`)

#### 2. Type 3.2 Delay (Not Submitted & Overdue)
* **Condition:** No actual submission date is recorded AND today's date is on or after the forecast date.
  $$\text{ActualDate}_g == \text{None} \quad \land \quad \text{TODAY} \ge \text{RefDate}_g$$
* **Calculation:**
  $$\text{Delay Days} = \max((\text{TODAY} - \text{RefDate}_g).\text{days}, 1)$$
* **UI Indicator:** 🔴 Red Badge (`Type 3.2: Overdue & Not Submitted`)

### 4.4 14-Day Lookahead Warning & Slippage (`REQ-05`)
* **Condition:** Document has not yet been submitted (`ActualDate == None`), and the forecast due date falls within the upcoming two weeks:
  $$\text{TODAY} < \text{RefDate}_g \le \text{TODAY} + 14 \text{ days}$$
* **Slippage Flag (`⚠️ Slipping`):** A lookahead document is highlighted for slippage if its forecast submission date has slipped beyond its original baseline plan:
  $$\text{ForecastDate}_g > \text{PlanDate}_g$$

---

## 5. Dashboard Sections & Presentation Architecture

To fulfill user requirements, the frontend Single-Page Application (`index.html`, `app.js`, `index.css`) must present data partitioned into **4 distinct work package sections** (plus a unified Executive Summary):

```
+-----------------------------------------------------------------------------------+
| EXECUTIVE HEADER: Global Progress KPIs | 📁 Weekly Cut-off Upload | Filter Bar   |
+-----------------------------------------------------------------------------------+
| SECTION TABS:                                                                     |
| [ 🏢 Executive Overview ] [ 🛠️ WP-1 Topside Excl. ] [ 🏗️ WP-1 Topside Structure ] |
| [ ⚓ WP-1 Jacket ]          [ 🛢️ WP-2 Pipeline ]                                     |
+-----------------------------------------------------------------------------------+
| ACTIVE SECTION CONTENT:                                                           |
| +-------------------------------------------------------------------------------+ |
| | KPI Summary Cards:                                                            | |
| | (Total Docs) | (Total Delayed) | (Type 3.1 Late) | (Type 3.2 Overdue) | (Lookahead)| |
| +-------------------------------------------------------------------------------+ |
| +-----------------------------------------------+ +-----------------------------+ |
| | Discipline Progress & Status Stacked Chart    | | SPI / S-Curve Progress Chart| |
| +-----------------------------------------------+ +-----------------------------+ |
| +-------------------------------------------------------------------------------+ |
| | Sub-Tabs for Active Section:                                                  | |
| | [ 📋 Master Register ] [ 🚨 Overdue & Delayed (Type 3.1 & 3.2) ] [ ⚠️ Lookahead ] | |
| +-------------------------------------------------------------------------------+ |
| | Interactive Filterable & Searchable Data Table with Status Chips & Badges    | |
| +-------------------------------------------------------------------------------+ |
+-----------------------------------------------------------------------------------+
```

### 5.1 Section Breakdown
1. **Executive Overview Tab:** Aggregates statistics, overall S-Curves, and high-level delay risks across all 4 work packages simultaneously.
2. **WP-1 Topside Excl. Structure Tab:** Focuses purely on data extracted from `00210 EDSR as of WE 27 03 Jul 26.xlsm` (or newly uploaded weekly updates for Topside Excl).
3. **WP-1 Topside Structure Tab:** Focuses purely on data extracted from `Z1F Topsides EDSR cut-off_3Jul26.xlsx` (or newly uploaded weekly updates for Topside Structure).
4. **WP-1 Jacket Tab:** Focuses purely on data extracted from `Z1F Jacket EDSR cut-off_3Jul26.xlsx` (or newly uploaded weekly updates for Jacket).
5. **WP-2 Pipeline Tab:** Focuses purely on data extracted from `Att.2. WP2 Engineering Progress Measurement_030726_Updated.xlsx` (or newly uploaded weekly updates for Pipeline).

### 5.2 Interactive Sub-Views within Each Section
Each of the 4 work package tabs features three localized sub-tables:
* **Master Document Register (`MDR`):** Complete deliverable inventory showing Document Number, Title, Discipline, Lifecycle Status (`Not Yet Submitted`, `IFR`, `IFA`, `AFC`), and all Plan/Forecast/Actual dates.
* **Overdue & Delayed Deliverables (`Delay Tab`):** Strictly filters down to documents triggering **Type 3.1 (Submitted Late)** or **Type 3.2 (Overdue & Not Submitted)** delay rules. Displays exact `Delay Days` and aging badges (`Critical > 30 days`, `High 15-30 days`, `Medium 7-14 days`, `Low < 7 days`).
* **14-Day Lookahead Warning (`Lookahead Tab`):** Displays documents due in the next 14 days, highlighting `Days Remaining` and explicitly flagging `⚠️ Slipping from Plan` when $\text{Forecast} > \text{Plan}$.

---

## 6. UI/UX Design Standards & Aesthetics

* **Visual Theme:** Premium dark glassmorphism (`#030712` base background, semi-transparent backdrop-blur card containers with subtle borders `#1e293b`).
* **Color Palette:**
  * **Brand Primary:** `--accent-cyan` (`#06b6d4` / `#22d3ee`) for headers and primary highlights.
  * **Success / Submitted:** `--accent-emerald` (`#10b981`) for completed milestones and Type 3.1 low severity.
  * **Warning / Lookahead:** `--accent-amber` (`#f59e0b`) for lookahead due dates and Type 3.1 late submissions.
  * **Alert / Overdue:** `--accent-rose` (`#f43f5e`) for Type 3.2 overdue documents and critical delays.
* **Responsive Interactions:** Client-side instant filtering without page reloads, Chart.js interactive tooltips with drill-down highlights, and clean typography using Outfit/Inter fonts.

---

## 7. Verification & Acceptance Criteria

1. **Multi-Source Accuracy:** Running `python server.py` must successfully load all 4 Excel files without errors and return normalized JSON datasets for each section (`topside_excl`, `topside_structure`, `jacket`, `wp2_pipeline`).
2. **Weekly Upload Functionality:** Uploading a newly named Excel cut-off via the dashboard modal must properly auto-detect its Work Package, update `server.py` in-memory cache immediately, and re-render the dashboard cleanly.
3. **Delay Classification Compliance:**
   * Every document classified as Type 3.1 must have a non-empty Actual Date greater than Forecast Date.
   * Every document classified as Type 3.2 must have an empty Actual Date and a Forecast Date $\le$ `TODAY`.
4. **UI Functionality:** Clicking between the 4 work package tabs (`Topside Excl. Structure`, `Topside Structure`, `Jacket`, `Pipeline`) must instantly switch summary KPIs, charts, and table contents to reflect only that specific work package's progress.

---

## 8. Dual-Basis Reporting Methodology (Gates vs. MDR Document Baseline)

To ensure comprehensive engineering transparency and prevent any ambiguity regarding deliverable totals versus gate review volumes, all KPI cards across all Work Package views display **two distinct metrics side-by-side within the exact same KPI box**:

1. **Gate Milestones (`Gates - 3 Revs Basis`)**:
   * Measures individual review gate events across **IFR (50% Gate)**, **IFA (20% Gate)**, and **AFC (Approval Gate)**.
   * Because each document passes through up to 3 gates, the total gate scope is up to $3 \times \text{Total MDR Docs}$.
   * When counting delays (`Total Delayed Gates`, `Type 3.1 Submitted Late`, `Type 3.2 Overdue`), this number tracks how many individual gate submissions breached their Forecast Date or are overdue.

2. **Master Document Register Baseline (`MDR Docs Basis`)**:
   * Measures unique engineering deliverable rows (`MDR` files) directly matching the baseline inventory count.
   * When counting delays or status (`Total Delayed Docs Basis`, `Type 3.1 Docs Basis`, `On-Time Docs Basis`), this number represents the exact count of unique deliverable documents (`doc_no`) affected by at least one gate condition.
   * This guarantees direct comparability against the `Total Deliverables` card without duplicate document counting.

---

## 9. Color Coding, Layout Swapping & Multi-Metric Visual Refinements

To optimize visual hierarchy and executive scannability across all dashboard elements:

1. **Total Deliverables Card Layout**:
   * The dual-metric layout places **Max Gates (3 Revs)** (`kpiTotalGates`) on the left-hand column (`kpi-value`) and **MDR Rows (Docs)** (`kpiTotalDocs`) on the right-hand column (`kpi-value-secondary`).

2. **Strict Breach Color Standards**:
   * **Type 3.1: Submitted Late (`#b45309` Warm Copper/Brown):** Distinguishes documents submitted after Forecast Date (`Actual > Forecast`) with dedicated golden-brown `.alert-brown` and `.alert-bg-brown` badges.
   * **Type 3.2: Overdue (`#ef4444` Vibrant Red):** Distinguishes critical unsubmitted review gates (`Actual is empty & TODAY ≥ Forecast`) with crisp `.alert-rose` / `.alert-red` highlights.
   * **14-Day Lookahead Risk (`#eab308` Crisp Yellow - Latest Unsubmitted Revision Basis):** Evaluates strictly the **latest/next unsubmitted revision right now** (`the sequential gate where Actual Date is empty: IFR -> IFA -> AFC`). A document appears in the 14-day lookahead warning list if and only if its latest unsubmitted revision is due within the upcoming 14-day window (`TODAY < Forecast <= TODAY + 14 days`). Fully submitted documents (`AFC Submitted`) or future revisions beyond the immediate pending gate are never flagged.

3. **Combined Delay & Lookahead Charting (Unique MDR Documents Basis)**:
   * The second chart (`Delay Severity, Overdue & 14-Day Lookahead Risk by Discipline`) strictly plots quantities on **MDR unique document basis (`NOT base on 3 revisions`)**, displaying a 3-bar grouped comparison per discipline showing how many unique documents (`doc_no`) are affected:
     * Bar 1: Type 3.1 Late Documents (`#b45309` Brown)
     * Bar 2: Type 3.2 Overdue Documents (`#ef4444` Red)
     * Bar 3: 14-Day Lookahead Risk Documents (`#eab308` Yellow)

4. **Sub-Tab Pill Badges (Unique MDR Documents Basis)**:
   * To maintain consistent comparison against the baseline Master Document Register (`MDR`), all sub-tab pill badges (`subCountDelay3_1`, `subCountDelay3_2`, and `subCountLookahead`) strictly display the **unique MDR document quantity (`MDR Docs Basis`)** rather than total review gate milestones:
     * **Brown Pill (`3.1: X`)**: Unique MDR documents (`doc_no`) with Type 3.1 late submissions.
     * **Red Pill (`3.2: Y`)**: Unique MDR documents (`doc_no`) with Type 3.2 overdue breaches.
     * **Yellow Lookahead Pill (`[Z]`)**: Unique MDR documents (`doc_no`) due within the 14-day lookahead window.

---

## 10. Lessons Learned, Operational Handoff & Cloud Deployment Guide

### 10.1 Key Engineering Lessons Learned (`Lessons Learned`)
1. **Unique MDR Documents Basis vs. Total Revisions/Gates Scope**:
   * *Lesson:* When presenting high-level charts (`Chart 1: Discipline Progress` and `Chart 2: Delay Severity & Lookahead Risk`) or summary pill badges to Executive Project Managers, counts must strictly align with the unique **Master Document Register (`MDR`) baseline inventory (`1,461 documents`)** rather than total review gate scopes (`3 Revisions x 1,461 = 4,383 gates`).
   * *Impact:* Counting on a 3-revision basis skewed visual quantities (causing disciplines like Drawing to show 1,400+ delayed gates on a 460-document inventory). Standardizing both charts and pill badges to unique `doc_no` counts provides instant, verifiable audit alignment.

2. **Sequential Gate Evaluation for 14-Day Lookahead (`Latest Unsubmitted Revision Basis`)**:
   * *Lesson:* Engineering progress tracking is strictly sequential (`IFR -> IFA -> AFC`). A document cannot be evaluated for future `IFA` or `AFC` lookahead risks if `IFR` has not even been submitted yet.
   * *Impact:* The lookahead evaluation engine (`compute_delay_and_lookahead`) must first identify the **single latest pending revision (`the sequential gate where Actual Date is empty`)**. Only if that specific pending revision's `Reference Date` (`Forecast Date over Plan Date`) falls due within `TODAY < Forecast <= TODAY + 14 days` does the document trigger a lookahead warning. Fully submitted documents (`AFC Submitted`) are 100% excluded.

3. **Reference Date Hierarchy (`Forecast Date` over `Plan Date`)**:
   * *Lesson:* Contractor execution schedules constantly evolve via weekly re-forecasting (`*Cut off*.xlsx`). Evaluating delays against static `Plan Date` alone causes false alarms on mutually agreed re-forecasts.
   * *Impact:* The backend dynamically enforces $\text{Reference Date} = \text{Forecast Date if present, else Plan Date}$, ensuring that delay days and lookahead remaining days accurately reflect active contractor commitments.

---

### 10.2 Operational Handoff Summary
* **Master Spreadsheets Preservation:** All original Work Package Excel cut-off files (`00210 EDSR...xlsm`, `Att.2. WP2...xlsx`, `Z1F Topsides EDSR...xlsx`, `Z1F Jacket EDSR...xlsx`) are kept intact in the project workspace as the single source of truth (`Keep master file`).
* **Performance Optimization:** The Python backend (`server.py`) uses a thread-safe `_cache_lock` with multi-file openpyxl streaming, ingesting all 1,461 MDR records across 4 work packages in `<9 seconds` and serving instant API responses (`304 Not Modified` / `200 OK` in `<50ms`).
* **Self-Contained Architecture:** Zero external cloud database setup required; the system operates directly on standard file systems and works immediately out of the box.

---

### 10.3 Git Repository Preparation & Render Cloud Deployment (`Git & Render Guide`)

#### A. Prepared Deployment Configuration Files
1. **`requirements.txt`**: Specifies strict Python dependencies:
   ```text
   openpyxl>=3.1.0
   ```
2. **`.gitignore`**: Excludes Python bytecode and OS caches while explicitly **preserving all master Excel spreadsheets (`*.xlsx`, `*.xlsm`)** inside the repository so cloud servers can ingest data immediately:
   ```text
   __pycache__/
   *.py[cod]
   venv/
   .env
   .DS_Store
   Thumbs.db
   ```
3. **`render.yaml` (Render Blueprint Spec)**: Enables automated 1-click Web Service creation on Render.com:
   ```yaml
   services:
     - type: web
       name: z1f-engineering-progress-dashboard
       env: python
       buildCommand: pip install -r requirements.txt
       startCommand: python server.py
       plan: free
       envVars:
         - key: PORT
           value: 8090
         - key: PYTHONUNBUFFERED
           value: true
   ```
4. **`Procfile`**: Provides cross-platform compatibility for Render, Heroku, Railway, and Fly.io:
   ```text
   web: python server.py
   ```

#### B. Step-by-Step Git Upload Instructions
Open terminal inside the workspace root (`c:\Users\pipes\OneDrive\Documents\Google_AntiGravity\Project\Engineering_Progress_Z1F`) and execute:
```powershell
# 1. Initialize Git repository (if not already initialized)
git init

# 2. Stage all code, configuration files, and master spreadsheets
git add .

# 3. Commit the production-ready build
git commit -m "feat: Z1F Engineering Progress Dashboard — Final Production Build with MDR Basis Charting & Sequential Lookahead Engine"

# 4. Connect to your GitHub repository and push
git branch -M main
git remote add origin https://github.com/<YOUR_USERNAME>/<YOUR_REPOSITORY_NAME>.git
git push -u origin main
```

#### C. Step-by-Step Render Cloud Deployment Instructions (`Render.com`)
1. Log in to [Render.com](https://render.com) using your GitHub account.
2. Click **New +** $\rightarrow$ **Blueprint** (or **Web Service**).
3. Select your newly pushed GitHub repository (`Engineering_Progress_Z1F`).
4. If using **Blueprint**, Render will automatically detect `render.yaml` and configure the service (`Build: pip install -r requirements.txt`, `Start: python server.py`, `Port: 8090`).
5. Click **Apply Blueprint** (or **Create Web Service**).
6. Within `60 seconds`, Render will build the environment, run `python server.py`, ingest all 4 master Work Package Excel sheets on startup, and provide your live production URL (e.g., `https://z1f-engineering-progress-dashboard.onrender.com`)!

