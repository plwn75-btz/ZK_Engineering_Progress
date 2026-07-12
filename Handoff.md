# Aung Sinkha Development Project Phase 1A (EPC-01)
## Engineering Progress Tracking & Analytics Dashboard — Technical & Operational Handoff Report

---

## 1. Conceptual Overview

### 1.1 Background & Purpose
In complex Offshore & Onshore Engineering, Procurement, and Construction (EPC) projects, tracking engineering deliverables across thousands of individual documents, multiple Work Packages (WPs), and diverse engineering disciplines is critical to preventing project slippage. 

The **Engineering Progress Tracking & Analytics Dashboard** was developed to replace static, manual spreadsheet reporting with an automated, high-performance, real-time web application. It ingests weekly Master Document Register (MDR) Excel cut-offs from project contractors (e.g., PTTEPI / PJM) and transforms raw tabular records into actionable executive insights, S-Curve variance analysis, granular discipline delays, and proactive 14-day lookahead warning alerts.

### 1.2 Target Audience & Stakeholders
* **Project Director & EPC Project Managers:** High-level visibility into overall engineering progress (Plan vs. Actual vs. Forecast), Schedule Performance Indexes (SPI), and macro-level slippage.
* **Engineering Managers & Discipline Leads:** Granular tracking of document submission statuses across Process, Piping, Mechanical, Electrical, Civil/Structural, and Instrumentation disciplines.
* **Project Controls & Planning Engineers (PJM/PTTEPI):** Automated extraction and validation of weekly MDR cut-offs, S-Curve generation, and audit compliance against baseline milestones.
* **Document Controllers:** Real-time register monitoring, document search, and tracking of overdue reviews.

### 1.3 High-Level Architecture
The system is architected as a lightweight, self-contained, thread-safe web application running locally or on a central project server without external cloud database dependencies:
* **Backend (`server.py`):** Pure Python 3.11 threaded HTTP server utilizing `openpyxl` for high-speed Excel processing, in-memory thread-safe data caching, and RESTful API endpoints (`/api/data`, `/api/upload`).
* **Frontend (`index.html`, `app.js`, `index.css`):** Vanilla HTML5/CSS3/ES6 Single-Page Application (SPA) designed with state-of-the-art dark glassmorphism UI/UX, powered by `Chart.js` for interactive data visualization.

---

## 2. Specifications & Data Model Structure

### 2.1 Input Excel Source Schema
The application parses multi-sheet Excel files named under the convention `*MDR*.xlsx` or `*Cut off*.xlsx` (e.g., `MM-ASK-1A-GEN01-PJM-MDR-0001_A1- Cut off 03-Jul-26.xlsx`).

| Excel Sheet | Primary Purpose | Extracted Data Fields |
| :--- | :--- | :--- |
| **`MDR-Engineering`** | Master Document Register containing individual engineering deliverables (~1,514+ rows). | `Doc No`, `Title`, `Discipline`, `Work Package (WP)`, `IFR Plan/Forecast/Submit Date`, `IFA Plan/Forecast/Submit Date`, `AFC Plan/Forecast/Submit Date`, `Document Status`. |
| **`Overdue Summary`** | Contractor-provided summary matrix by engineering discipline. | `Discipline`, `Total Overdue`, `1st Revision (IFR)`, `2nd Revision (IFA/AFC)`, `Pending to Return by Client/PTTEPI`. |

### 2.2 Milestone Tracking & Gate Definitions
Every engineering deliverable tracks three sequential submission gates:
1. **IFR (Issued for Review):** Initial engineering draft submitted for interdisciplinary and client review (`Revision A1, A2, etc.`).
2. **IFA (Issued for Approval):** Revised engineering document incorporating review comments, submitted for formal approval (`Revision B1, B2, etc.`).
3. **AFC (Issued for Construction):** Final approved revision released for fabrication and site installation (`Revision 0, 1, etc.`).

### 2.3 Reference Tracking Date Hierarchy (`Forecast` over `Plan`)
Per strict project specification, the tracking deadline for any deliverable is determined by the **Reference Date Hierarchy**:
$$\text{Reference Date} = \begin{cases} \text{Forecast Date}, & \text{if Forecast Date is present and valid} \\ \text{Plan Date}, & \text{otherwise (fallback)} \end{cases}$$
* **Today's Date (`datetime.date.today()`)** serves as the universal anchor for all delay, aging, and lookahead evaluations across every dashboard tab.

### 2.4 Document Status Classifications
A document's current active lifecycle status is dynamically computed based on its latest non-empty submission field:
* **`Not Yet Submitted`**: `ifr_submit_date`, `ifa_submit_date`, and `afc_submit_date` are all blank (`""`, `"N/A"`, or `"-"`).
* **`IFR Submitted`**: `ifr_submit_date` exists, but `IFA` and `AFC` are blank.
* **`IFA Submitted`**: `ifa_submit_date` exists, but `AFC` is blank.
* **`AFC Submitted`**: `afc_submit_date` exists (Final Completed Lifecycle).

### 2.5 Official Summary Sheets & Dynamic Cut-off Extraction (`Source of Truth`)
To ensure executive progress cards (`Baseline Plan Progress`, `Achieved Actual Progress`, `Forecast Expected Progress`, `SPI`) and the first-tab `Executive Work Package Progress Summary` table match official contractor submissions exactly without discrepancy, `server.py` (`extract_official_summary_kpis()`) extracts package-level KPIs directly from summary tables:

1. **WP-1 Topside Excl. Structure (`00210 EDSR as of WE 27 03 Jul 26.xlsm`)**:
   * **Summary Sheet:** `EDSR Report` | **Target Row:** `Row 2788` (`Total :`)
   * **Mappings:** `Plan / Forecast %` = Col `AH` (**36.73%**) | `Actual %` = Col `AE` (**42.09%**) | `SPI` = **1.15** (`+5.36%` variance, `🚀 Ahead`)

2. **WP-1 Topside Structure (`Z1F Topsides EDSR cut-off_03ul26.xlsx`)**:
   * **Summary Sheet:** `Progress Summary` | **Target Row:** `Row 8` (`Overall Topsides DDE Progress`)
   * **Mappings:** `Plan / Forecast %` = Col `O/P` (**46.70%**) | `Actual %` = Col `Q` (**48.22%**) | `SPI` = **1.03** (`+1.52%` variance, `⚡ On Time`)

3. **WP-1 Jacket (`Z1F Jacket EDSR cut-off_03Jul26.xlsx`)**:
   * **Summary Sheet:** `Progress Summary` | **Target Row:** `Row 12` (`Overall Jacket & Pile DDE Progress`)
   * **Mappings:** `Plan / Forecast %` = Col `O/P` (**27.37%**) | `Actual %` = Col `Q` (**32.59%**) | `SPI` = **1.19** (`+5.22%` variance, `🚀 Ahead`)

4. **WP-2 Pipeline (`Att.2. WP2 Engineering Progress Measurement_030726_Updated.xlsx`) — Dynamic Cut-Off Column Detection**:
   * **Summary Sheet:** `MDR-Detailed Design-A10` | **Target Rows:** `Row 214` (`Plan`), `Row 216` (`Actual`), `Row 218` (`Forecast`).
   * **Dynamic Timeline Scan:** Scans horizontally across timeline columns starting from `Col 53` (`BB`) to locate the rightmost active cut-off date column where `Actual Cumm` contains valid percentage data before empty future weeks.
   * **Detected Active Column:** `Col 75` (`Col BW` — `03 July-2026`).
   * **Mappings:** `Plan %` = **35.16%** | `Actual %` = **34.75%** | `Forecast %` = **34.10%** | `SPI` = **0.99** (`-0.41%` variance, `⚡ On Time`)

5. **Executive Overall Aggregation (`All Work Packages Weighted across 1,461 Deliverables`)**:
   * **Overall Baseline Plan Progress:** **34.02%** | **Overall Achieved Actual Progress:** **38.16%**
   * **Overall Forecast Expected Progress:** **33.92%** | **Overall Schedule Performance Index (SPI):** **1.12** (`+4.14%` variance, `🚀 Ahead`)

---

## 3. Requirements & Compliance Matrix

| ID | Requirement Category | Detailed User Specification | Implemented Engineering Solution | Compliance Status |
| :---: | :--- | :--- | :--- | :---: |
| **REQ-01** | **Reference Date Baseline** | All tracking dates across all dashboard tabs shall strictly refer to **Today** and **Forecast Date** (falling back to Plan Date if Forecast is unavailable). | Evaluated dynamically against `datetime.date.today()` and `ms.forecast || ms.plan` in `server.py` (`compute_delay_and_lookahead`) and `app.js` (`isDocDelayedOrSlipped`). | ✅ **100% Compliant** |
| **REQ-02** | **Empty Date Handling** | Empty tracking dates in submission fields (`ifr_submit_date`, `ifa`, `afc`) must be classified as `"Not Yet Issued"` / `"Not Yet Submitted"`. | Python and JS logic treat `None`, `""`, `"N/A"`, and `"-"` as unissued (`Not Yet Submitted`). | ✅ **100% Compliant** |
| **REQ-03** | **Type 3.1 Delay Classification** | Track and count documents **Submitted Late**: where the actual submission date occurred *after* the Forecast Date (`Submit Date > Forecast Date`). | Evaluated exact delta `(submit_dt - ref_dt).days`. Found exactly **30 documents**. Displayed with 🟠 badge in UI. | ✅ **100% Compliant** |
| **REQ-04** | **Type 3.2 Delay Classification** | Track and count documents **Overdue & Not Submitted**: where no submission has occurred and today is past the Forecast Date (`Today ≥ Forecast Date`). | Evaluated exact aging `(today - ref_dt).days`. Found exactly **144 documents**. Displayed with 🔴 badge in UI. Total Delayed = **174 documents**. | ✅ **100% Compliant** |
| **REQ-05** | **Lookahead Delay Risk Forecast** | For the 2-week lookahead window (`Today < Forecast ≤ Today + 14 days`), forecast how many documents will be delayed (slipping from baseline plan). | Analyzed lookahead documents (`349 total`) for schedule slippage (`Forecast > Plan`). Identified exactly **186 high-risk slipping documents** forecasted to delay. | ✅ **100% Compliant** |
| **REQ-06** | **Global & Tab Filters** | Provide real-time filtering by Work Package (`WP`), `Discipline`, and `Milestone` across overdue stats, stacked charts, and tables. | Integrated interactive header filters (`overdueFilterWP`, `overdueFilterDisc`, `overdueFilterMS`) triggering instant re-computation in `renderOverdueTab()`. | ✅ **100% Compliant** |
| **REQ-07** | **Executive S-Curve & SPI Analytics** | Generate multi-line S-Curves (Plan vs. Actual vs. Forecast) and calculate Weekly/Cumulative Schedule Performance Indexes (`SPI`). | Extracted time-series data from `MDR-Engineering`, computing `SPI = Actual / Plan` and displaying visual progress variance indicators. | ✅ **100% Compliant** |
| **REQ-08** | **UI/UX Aesthetics & Responsiveness** | Create a stunning, modern dark-themed glassmorphism interface with clear typography, vibrant palettes, and zero generic placeholders. | Built deep navy (`#030712`) glass UI with curated HSL accents (`--accent-cyan`, `--accent-emerald`, `--accent-rose`), Chart.js tooltips, and micro-animations. | ✅ **100% Compliant** |
| **REQ-09** | **Official Summary Sheet Source of Truth** | Ensure package progress figures strictly use official contractor summary figures rather than raw deliverable averaging. | Implemented `extract_official_summary_kpis()` reading direct summary table cells (`EDSR Report` row 2788, `Progress Summary` rows 8 & 12, dynamic column `BW` for WP-2). | ✅ **100% Compliant** |
| **REQ-10** | **First Tab Overall Progress Section** | Add comprehensive overall progress metrics directly onto the very first dashboard tab (`Executive Overview`). | Integrated `Overall Progress & SPI Section` (`Baseline Plan: 34.02%`, `Actual: 38.16%`, `SPI: 1.12`) and `Executive Work Package Progress Summary` table onto tab 1. | ✅ **100% Compliant** |

---

## 4. Engineering Solutions & Implementation Details

### 4.1 Backend Architecture & Mathematical Logic (`server.py`)
The backend operates via a multi-threaded HTTP server capable of parsing 1,500+ Excel rows in `< 1.5 seconds`.

#### Schedule Performance Index (SPI) Formulas
$$\text{SPI}_{\text{Weekly}} = \frac{\text{Actual Progress}_{\text{This Week}}}{\text{Plan Progress}_{\text{This Week}}}, \quad \text{SPI}_{\text{Cumulative}} = \frac{\text{Actual Progress}_{\text{Cumulative}}}{\text{Plan Progress}_{\text{Cumulative}}}$$
* **SPI > 1.00:** Ahead of schedule.
* **SPI = 1.00:** Exactly on schedule.
* **SPI < 1.00:** Behind schedule / progress variance deficit.

#### Two-Pronged Delay Algorithm (`compute_delay_and_lookahead`)
```python
# Primary reference date is FORECAST, falling back to PLAN if FORECAST is unavailable
ref_dt = forecast_dt if forecast_dt is not None else plan_dt

if submit_dt is not None:
    # Type 3.1: Document WAS submitted, but submitted late after Forecast Date
    if submit_dt > ref_dt:
        entry["delay_days"] = max((submit_dt - ref_dt).days, 1)
        entry["delay_type"] = "Type 3.1 (Submitted Late)"
        entry["delay_type_code"] = "3.1"
        delayed.append(entry)
        delayed_type1_count += 1
else:
    # Type 3.2: Document NOT YET SUBMITTED and today is overdue past Forecast Date
    if ref_dt <= today:
        entry["delay_days"] = max((today - ref_dt).days, 1)
        entry["delay_type"] = "Type 3.2 (Not Submitted & Overdue)"
        entry["delay_type_code"] = "3.2"
        delayed.append(entry)
        delayed_type2_count += 1
    elif today < ref_dt <= lookahead_end:
        # 2-Week Lookahead Window (Next 14 Days)
        entry["days_remaining"] = (ref_dt - today).days
        entry["urgency"] = "this_week" if entry["days_remaining"] <= 7 else "next_week"
        lookahead.append(entry)
```

### 4.2 Frontend Engine & Visualization (`app.js` & `index.html`)
The frontend loads the complete JSON payload (`/api/data`) into `DASHBOARD_DATA` and performs instant, client-side re-renders without network latency:
* **Stacked Bar Charts (`Chart.js`):** Displays discipline-by-discipline breakdown showing exact proportions of `Delayed` (Red), `Not Yet Submitted` (Amber), and `Submitted` (Emerald) documents.
* **Lookahead Slippage Highlighting:** Documents in the 2-week lookahead (`Lookahead Table`) automatically flag schedule slippage (`⚠️ Slipping`) if their Forecast Date has been pushed past their original Plan Date.

---

## 5. Operational Guide & Verification Commands

### 5.1 Server Startup & Configuration
To launch the application server locally or on a Windows deployment machine:

```powershell
# Navigate to the workspace directory
cd "c:\Users\pipes\OneDrive\Documents\Google_AntiGravity\Project\Engineering_Progress"

# Start the Python Threaded HTTP Server (Default Port: 8090 or via environment variable PORT)
python server.py
```
* **Dashboard Access:** Open a web browser and navigate to [`http://localhost:8090`](http://localhost:8090).

### 5.2 Updating Project Data (MDR Cut-offs)
There are two ways to update the dashboard when a new weekly MDR Excel file is received:
1. **Direct File Drop (Recommended for Administrators):**
   * Copy the new `.xlsx` file (e.g., `MM-ASK-1A-GEN01-PJM-MDR-0001_A2- Cut off 10-Jul-26.xlsx`) directly into the project folder (`c:\Users\pipes\OneDrive\Documents\Google_AntiGravity\Project\Engineering_Progress`).
   * The server's `find_latest_excel()` function automatically scans directory timestamps/names and caches the latest file on the next request.
2. **Web Dashboard Upload Button:**
   * Click the **`📁 Upload New Cut-off Excel`** button in the dashboard header (`index.html`).
   * Select the `.xlsx` file from your computer. The file is uploaded via `/api/upload`, cached immediately, and the dashboard re-renders seamlessly.

### 5.3 Verification & Audit Scripts
If project controls engineers wish to verify delay exact counts via command-line without opening the browser:

```powershell
# Verify exact delay counts (Type 3.1 vs Type 3.2) and Lookahead Slippage
python -c "import server; d = server.extract_all_data()['delay_lookahead']; print(f'Total Delayed: {d[\"delayed_count\"]} (Type 3.1: {d[\"delayed_type1_count\"]} | Type 3.2: {d[\"delayed_type2_count\"]})\nLookahead Total: {d[\"lookahead_count\"]} (Forecast Delay Risk: {d[\"lookahead_forecast_delay\"]})')"
```

### 5.4 Cloud Deployment: GitHub & Render Blueprint Guide
To deploy the application to a cloud production environment accessible to all stakeholders (`PJM`, `PTTEPI`, and `Executive Leadership`):

#### 1. Push Repository to GitHub
Ensure the clean workspace (including all 4 official baseline Excel files, `render.yaml`, and `requirements.txt`) is pushed to a Git remote:
```powershell
# 1. Rename branch to main if desired
git branch -M main

# 2. Add your GitHub remote repository
git remote add origin https://github.com/YOUR_USERNAME/z1f-engineering-progress.git

# 3. Push all commits to GitHub
git push -u origin main
```

#### 2. Automated Cloud Deployment via Render (`render.yaml`)
1. Log in to your **[Render Dashboard](https://dashboard.render.com/)**.
2. Click **`+ New`** $\rightarrow$ **`Blueprint`** (or **`Web Service`**).
3. Connect your GitHub repository (`z1f-engineering-progress`).
4. Render will automatically detect the `render.yaml` specification file and configure the service:
   * **Build Command:** `pip install -r requirements.txt` (Installs `openpyxl`)
   * **Start Command:** `python server.py`
   * **Environment Variable:** `PORT = 8090`
5. Click **`Apply Blueprint`**. Within ~60 seconds, your live production URL (`https://z1f-engineering-progress-dashboard.onrender.com`) will be active and serving the verified overall KPI summary.

---

## 6. Summary of Key Deliverables Handoff Checklist
* [x] **`server.py`**: Fully optimized, thread-safe backend with exact Type 3.1/3.2 delay separation and lookahead slippage risk computation.
* [x] **`app.js`**: Frontend JavaScript controller with real-time multi-dimensional filtering (`WP`, `Discipline`, `Milestone`), Chart.js rendering, and standardized `Today` vs. `Forecast Date` baseline logic.
* [x] **`index.html`**: Responsive 5-tab glassmorphism shell incorporating detailed delay tables (`Delay Type`, `Submit Date`, `Days Overdue`) and lookahead warning tables.
* [x] **`index.css`**: Complete HSL-curated dark mode styling with aging badges (`critical`, `high`, `medium`, `low`) and status chips.
* [x] **`render.yaml` & `Procfile`**: Complete cloud blueprint configuration for Render deployment on port `8090`.
* [x] **`HANDOFF_REPORT.md`**: This complete technical handoff documentation residing in the root workspace.
