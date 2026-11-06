# TCUA Perfect - BOM Comparison & Analysis System

TCUA Perfect is a high-performance Bill of Materials (BOM) comparison tool designed for automotive engineering teams. It provides a two-stage analysis pipeline that combines fuzzy semantic matching with hierarchical tree visualization.

## 🚀 Key Features

- **Smart Submodule Matching**: Uses SBERT semantic embeddings and structured part-ID segments to detect "Replacements" even when part numbers change.
- **Hierarchical Visualization**: Interactive tree explorer for navigating multi-level BOMs with property-based filtering (Duplicates, Hierarchy Changes, Path Changes).
- **Manual Override Support**: Respects manual classifications made in Excel report sheets, ensuring the dashboard matches your "Source of Truth."
- **Excel KPI Export**: Generates comprehensive Excel reports with submodule-level summaries and change classifications.

## 🛠️ Technology Stack

- **Backend**: FastAPI (Python)
- **Matching Engine**: SBERT (Sentence-Transformers), pandas, NumPy
- **Frontend**: Vanilla JS, HTML5, CSS3 (with Glassmorphism design)
- **Desktop Wrapper**: Electron.js

## 📁 Project Structure

- `app.py`: FastAPI server and worker launcher.
- `bom_comp_levelcs.py`: Core comparison engine (Fuzzy Matching, Barrel Analysis).
- `generate_json_hierarchy.py`: Hierarchy builder and Excel report injector.
- `static/`: Frontend assets (Dashboards, Visualizers).
- `config91112.json`: Business rules, segment definitions, and thresholds.
- `models/`: Bundled SBERT model for offline semantic analysis.
- `extra/`: Archive of test scripts and legacy utilities.

## 🏁 Getting Started

### Development Mode
1. Ensure you have Python 3.10+ installed.
2. Install dependencies: `pip install -r requirements.txt`
3. Run the server: `python app.py`
4. Access the UI at: `http://127.0.0.1:9112`

### Production Build (Desktop App)
1. Install Node.js dependencies: `npm install`
2. Build the portable executable: `npm run dist`
3. The packaged app will be available in the `Deployment/` folder.

## 📊 Business Logic: Change Classification

- **Newly Added**: Submodules present only in the New BOM.
- **Deleted**: Submodules present only in the Old BOM.
- **Modified**: Part ID matches, but attributes (Qty, Description) changed.
- **Replaced**: Old item matched to a New item via high-similarity fuzzy matching.
- **Carry-Over**: No changes detected.

---
*Maintained by the BOM Analytics Team.*
