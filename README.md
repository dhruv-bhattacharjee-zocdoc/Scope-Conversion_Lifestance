# Scope Conversion Lifestance - FINAL Documentation

## Project Overview
This project automates the transposition, transformation, and validation of provider and location data for Lifestance. Data flows from raw Excel input through a series of extraction, cleaning, enrichment, matching, and output scripts, producing fully validated and structured Excel output for business processes.

**Latest Update: This version is finalized and production-ready.**

## Key Features
- Full CLI and GUI (PyQt5) options (`_main_1.py` and `lifestance_ui.py`)
- Automated API lookups for provider specialties and locations
- Intelligent fuzzy matching, street suffix and address normalization
- Data validation, drop-down generation, and real-time highlighting for data anomalies
- Seamless integration with `Snowflake`, `Zocdoc` APIs, and local reference files
- All Python dependencies specified in `requirements.txt`

## Main Workflow (CLI)
1. **Input Files**: Place your data as `Excel Files/Input_.xlsx` (or `Input.xlsx` as expected by default)
2. **Template File**: `Excel Files/New Business Scope Sheet - Practice Locations and Providers.xlsx`
3. **Run Main Script**: Run `_main_1.py` to process the data and orchestrate all steps
   - Extracts fields (helper scripts per column/task)
   - Copies template for structure/validation
   - Populates Provider and Location sheets
   - API integrations to enrich specialty/location info
   - Adds dropdowns/validations via multiple scripts
   - Final output: `Excel Files/Output.xlsx`

### To Run (CLI)
```bash
pip install -r requirements.txt
python _main_1.py
```

---

## Desktop GUI Utility
- **File**: `lifestance_ui.py`
    - Modern PyQt5 graphical launcher for the pipeline
    - Browse/select input/output files
    - View logs, timers, and progress visually
    - Displays user info (from `credentials.json`)
    - Run with:
      ```bash
      python lifestance_ui.py
      ```

---

## API Integration Scripts
- `API_Datamerge.py`: Merges API-enriched data and post-processes output sheets
- `api_for_specialty.py`: Looks up specialty IDs via Snowflake and writes to NPI file
- `api_for_location.py`: Fetches/updates practice and location Cloud IDs from REST API and populates reference sheets

## Main Python Script Reference

| File Name                | Description                                                                                       |
|-------------------------|---------------------------------------------------------------------------------------------------|
| `_main_1.py`            | Main orchestrator: runs the full workflow, calls extract/transform scripts, manages output files. |
| `lifestance_ui.py`      | PyQt5 graphical desktop interface for process management, logs, and file handling.                |
| `API_Datamerge.py`      | Merges API-enriched location/specialty/provider data and post-processes output Excel sheets.      |
| `api_for_specialty.py`  | Retrieves provider specialties from Snowflake and updates NPI-specialty mapping files.            |
| `api_for_location.py`   | Fetches and updates practice/location Cloud IDs from the API and updates reference Excel sheets.   |
| `Location.py`           | Generates the Location output sheet, cleans/normalizes addresses, and applies suffix mapping.     |
| `locationmapping.py`    | Standardizes address suffixes and matches provider/locations using fuzzy rules and reference data.|
| `Name.py`               | Extracts provider names and gender fields from the raw input sheet.                               |
| `Npi.py`                | Extracts NPI numbers from input data for use in downstream matching and validation.               |
| `Headshot.py`           | Extracts headshot/photo URLs or image paths.                                                      |
| `professional_suffix.py`| Extracts and assigns provider professional suffixes (e.g., MD, DO, PhD) for validation.           |
| `Specialty.py`          | Extracts provider specialty information and prepares validations.                                 |
| `PatientsAccepted.py`   | Manages dropdowns for patient types accepted (Adult, Pediatric, Both).                           |
| `Education.py`          | Extracts provider education and school information.                                               |
| `Professional_statement.py` | Extracts and sanitizes provider bios/statements.                                             |
| `Board_certification.py`| Extracts board certification and subspecialty fields.                                             |
| `provider_dropdowns.py` | Applies provider data validation and dropdowns to output Excel files.                             |
| `specialtydropdown.py`  | Adds specialty-specific dropdowns via data validation.                                            |
| `ESF.py`                | Adds and manages Enterprise Scheduling Flag dropdowns.                                            |
| `optoutrating.py`       | Manages the Opt Out of Ratings field/dropdown.                                                    |
| `Langauge.py`           | Extracts and validates languages spoken by providers.                                             |
| `suffix_check.py`       | Highlights invalid professional suffixes in output to catch user entry/mapping errors.             |
| `_status _check.py`     | Checks and validates the structure and completeness of output files.                              |


## Major Data & Processing Scripts
- `_main_1.py`, `Location.py`, `Name.py`, `Npi.py`, `Headshot.py`, `professional_suffix.py`, `Specialty.py`, `PatientsAccepted.py`, `Education.py`, `Professional_statement.py`, `Board_certification.py`, `provider_dropdowns.py`, `specialtydropdown.py`, `ESF.py`, `optoutrating.py`, `Langauge.py`, `locationmapping.py`, `suffix_check.py`, and others

---

## Requirements / Setup
- Python 3.x
- All dependencies as in `requirements.txt`
- Credentials for APIs (configure `credentials.json` as described in the GUI)
- Place all Excel input/template/reference files in the `Excel Files/` folder

---

## Core Files & Inputs
| File/Folder                | Purpose                                                    |
|---------------------------|------------------------------------------------------------|
| `Excel Files/Input_.xlsx` | Input data (rename as needed, e.g., `Input.xlsx`)          |
| `Excel Files/New Business Scope Sheet - Practice Locations and Providers.xlsx` | Output structure and data validations |
| `Excel Files/C1 Street Suffix Abbreviations.xlsx` | Address normalization (reference)               |
| `Excel Files/Output.xlsx` | Main generated output                                      |
| `Excel Files/Mergedoutput.xlsx` | Special/merged output in advanced API workflows         |
| `credentials.json`        | User/role credentials for APIs and GUI display              |
| `requirements.txt`        | Python dependencies                                        |
| `Documentations/README.md`| Additional documentation, links to Google Docs              |

---

## Frequently Asked Questions / Troubleshooting
- **Missing columns or errors?** Ensure all required columns exist in your input and template files. Watch logs (in CLI or GUI).
- **API credentials errors?** Ensure `credentials.json` is present and filled.
- **Output file missing some validations/dropdowns?** Check for template or reference file version mismatches.
- **Change input filename/reference?** Adapt scripts or rename your input as needed.

---

## Contact / Authors
This solution is finalized as of October 2025 for internal Lifestance business use. For legacy support or enhancements, contact your project admin or team lead.
