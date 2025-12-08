# Financial Grant Reporting Automation

## Overview
This tool automates the consolidation of university grant financial reports. It merges raw budget transaction logs with project metadata (PI names, grant titles, dates) to create a clean, formatted executive summary in Excel.

## The Problem
Financial data was previously siloed in three separate exports:
1. **Budget Files:** Transaction-level details (Expenses, Encumbrances).
2. **Project Files:** Metadata (Sponsors, PI Names, Dates).

Manually combining these for weekly reporting was time-consuming and prone to copy-paste errors.

## The Solution
This Python script:
1. **Ingests** multiple raw Excel dumps.
2. **Cleans** data types (handling text vs numeric Project IDs).
3. **Merges** datasets using a Left Join on Project ID.
4. **Formats** the final output with:
   - Conditional logic for missing data.
   - Professional styling (headers, alternating rows, borders).
   - Currency formatting.

## Usage
1. Install dependencies:
   ```bash
   pip install -r requirements.txt
