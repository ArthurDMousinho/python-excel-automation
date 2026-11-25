# Python Excel & CSV Automation Pipelines

Automated data engineering scripts built with **Python (`pandas` & `openpyxl`)** to eliminate manual spreadsheet processing, standardize messy datasets, and generate ready-to-use corporate Excel reports.

## Included Scripts

### 1. `crm_data_cleaner.py`
* **Purpose:** Sanitizes and standardizes raw CRM lead exports.
* **Features:** Removes empty rows, trims whitespace, normalizes capitalization and email strings, standardizes phone numbers to `(XXX) XXX-XXXX`, maps country abbreviations, deduplicates records by unique email, and exports with auto-fitted columns and corporate styling.

### 2. `sales_merger.py`
* **Purpose:** Consolidates fragmented monthly e-commerce CSV exports into a single master report.
* **Features:** Merges multiple files, removes duplicate order IDs, casts text-formatted currency strings into native numeric accounting formats (`$#,##0.00`), and auto-adjusts layout dimensions.
