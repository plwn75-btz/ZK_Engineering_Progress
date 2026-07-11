# Z1F Engineering Progress Measurement Dashboard (EPC-01)
### Advanced Real-Time Tracking, S-Curve Analytics & Lookahead Warning System

---

## 🌟 Executive Summary
The **Z1F Engineering Progress Measurement Dashboard** is a state-of-the-art, high-performance web application tailored for complex Offshore & Onshore Engineering, Procurement, and Construction (EPC) projects. It automates the weekly ingestion and transformation of Master Document Register (`MDR`) Excel cut-offs across multiple project contractors, delivering instant executive KPIs, multi-revision gate tracking (`IFR -> IFA -> AFC`), delay severity categorization (`Type 3.1` vs `Type 3.2`), and proactive `14-Day Lookahead Warning` alerts.

---

## 🏗️ Key Architectural & Engineering Highlights

### 1. Unique MDR Documents Basis (Single Source of Truth)
* Both interactive overview charts (**Chart 1: Discipline Progress Stacked Bar** and **Chart 2: Delay Severity & Lookahead Risk**) strictly operate on the **Unique Master Document Register (`MDR`) Baseline (`1,461 documents`)**, eliminating distortion from multi-revision gate multipliers.
* Sub-tab pill badges (`subCountDelay3_1`, `subCountDelay3_2`, and `subCountLookahead`) directly mirror unique `doc_no` counts for seamless audit verification.

### 2. Sequential Lookahead Evaluation (`Latest Unsubmitted Revision Basis`)
* Enforces strict engineering workflow sequence (`IFR -> IFA -> AFC`). The lookahead engine dynamically identifies each document's **single latest pending revision** where `Actual Date` is empty.
* A document triggers a `14-Day Lookahead Warning` if and only if that specific pending revision is due within the upcoming 14-day window (`TODAY < Forecast <= TODAY + 14 days`). Fully submitted documents (`AFC Submitted`) are automatically excluded.

### 3. Reference Date Hierarchy (`Forecast` over `Plan`)
* Dynamically evaluates all schedule milestones using:
  $$\text{Reference Date} = \begin{cases} \text{Forecast Date}, & \text{if Forecast Date is present and valid} \\ \text{Plan Date}, & \text{otherwise (fallback)} \end{cases}$$
* Ensures accurate alignment with active contractor re-forecast commitments while anchoring calculations to real-time execution dates.

---

## 📦 Project Structure & Prepared Deployment Files
```text
├── 00210 EDSR as of WE 27 10 Jul 26.xlsm               # Master Excel Cut-off (WP1 Topside Excl)
├── Att.2. WP2 Engineering Progress Measurement...xlsx   # Master Excel Cut-off (WP2 Pipeline)
├── Z1F Topsides EDSR cut-off_10Jul26.xlsx               # Master Excel Cut-off (WP1 Topside Structure)
├── Z1F Jacket EDSR cut-off_10Jul26.xlsx                 # Master Excel Cut-off (WP1 Jacket Structure)
├── server.py                                            # High-speed Python threaded backend + API
├── app.js                                               # SPA Frontend Controller & Chart.js Engine
├── index.html                                           # Dark Glassmorphism UI Structure
├── styles.css                                           # Design System & Color Standardization
├── Spec.md                                              # Comprehensive System Specification & Lessons Learned
├── requirements.txt                                     # Python dependencies (openpyxl>=3.1.0)
├── .gitignore                                           # Git exclusion rules (preserving master Excel files)
├── render.yaml                                          # Render.com Blueprint cloud deployment spec
└── Procfile                                             # Cloud execution command (web: python server.py)
```
*(Note: All master Work Package spreadsheets (`*.xlsx`, `*.xlsm`) are preserved intact inside the repository so cloud servers can ingest real project data immediately on startup.)*

---

## 🚀 Local Development & Execution
1. Ensure Python 3.11+ is installed on your Windows machine.
2. Install the required dependency:
   ```powershell
   pip install -r requirements.txt
   ```
3. Launch the threaded project server:
   ```powershell
   python server.py
   ```
4. Open your web browser to **`http://localhost:8090`**.

---

## ☁️ 1-Click Cloud Deployment (`Render.com`)
This repository is fully configured for zero-downtime cloud hosting on [Render](https://render.com):
1. **Push to GitHub**:
   ```powershell
   git init
   git add .
   git commit -m "feat: Z1F Engineering Progress Dashboard Production Build"
   git branch -M main
   git remote add origin https://github.com/<YOUR_USERNAME>/<YOUR_REPO_NAME>.git
   git push -u origin main
   ```
2. **Deploy on Render**:
   * Log in to your Render account and click **New +** $\rightarrow$ **Blueprint**.
   * Connect this GitHub repository. Render will auto-detect `render.yaml`, install `requirements.txt`, run `python server.py`, and bind to `$PORT` (`8090`).
   * Your dashboard will be live and ingesting master spreadsheets within seconds!
