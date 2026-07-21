# MDTGPT Complaint Review

Streamlit app for QA review and prioritization of raw complaint product event records using MDTGPT.

## Purpose

This tool helps review raw complaint and product event data for potential inconsistencies across:

- Coding vs Event Description
- Coding vs Regulatory Decisions
- Coding vs Investigation Decisions

It also assigns a prioritization tier to support QA follow-up.

## Features

- Upload raw Excel workbook
- Detect main raw data sheet automatically
- Review records using MDTGPT
- Flag potential issues in coding, regulatory, and investigation logic
- Assign priority tiers:
  - Tier 1 - Critical QA Review
  - Tier 2 - Targeted QA Review
  - Tier 3 - Follow-Up Signal
  - Tier 4 - No Review Signal
- Export categorized results to Excel

## Expected Input Columns

The app is designed for raw data files with columns such as:

- Product Event ID
- PE - PLI #
- Complaint? - PE
- Country - PE
- Lot Number - PE PLI
- Serial Number - PE PLI
- Product Description - PE PLI
- RFR Code
- FDP Code
- Reportable
- Successful GFE Attempts - PE PLI Comm
- PLI Tasks
- Rationale for no return - PE PLI
- Investigation Required
- Brazil FER
- China FER
- Japan FER
- Korea FER
- Event Description - PE
- Summary of Investigation Results
- Full UDI - PE PLI

## Requirements

- Python 3.10+
- Streamlit
- pandas
- requests
- openpyxl

Install dependencies with:

```bash
py -m pip install -r requirements.txt
