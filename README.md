# Excel GST Automation & Tally Data Preparation Tool

A Python-based Excel automation tool designed to automate GST-wise invoice processing, freight calculation, GST summaries, dynamic workbook generation, and preparation of structured sales data for downstream accounting and Tally-related workflows.

---

## Project Overview

Manual processing of GST-based Excel invoices can be repetitive, time-consuming, and prone to data-entry errors.

This project automates the transformation of an Excel workbook containing billing and GST information into a structured workbook with separate GST-rate sheets and Tally-oriented data preparation sheets.

The application reads an Excel workbook, processes GST-specific transactions, calculates GST values, creates summary rows, generates structured output sheets, and prepares the data required for further accounting-system integration.

---

## Key Features

- Windows file browser for selecting the input Excel workbook
- Automatic Excel workbook loading
- Dynamic detection of required columns
- GST-wise transaction separation
- Separate processing for 5% GST transactions
- Separate processing for 18% GST transactions
- Dynamic freight-row calculation
- CGST and SGST calculation and summarization
- GST summary-row generation
- Automatic serial-number renumbering
- Dynamic `Temp_5` sheet generation
- Dynamic `Temp_18` sheet generation
- Formula-based data mapping between worksheets
- Date-only voucher date generation
- Dynamic voucher number generation
- Tally-oriented ledger and item mapping
- Automatic workbook sheet reordering
- Automatic generation of a revised output workbook
- Reusable logic designed to work with different Excel workbooks

---

## Workbook Processing Flow

The application follows this general workflow:

```text
Input Excel Workbook
        │
        ▼
Read Bill Details
        │
        ▼
Detect GST Rate
        │
        ├───────────────┐
        │               │
        ▼               ▼
     GST 5%          GST 18%
        │               │
        ▼               ▼
    Sheet 5          Sheet 18
        │               │
        ▼               ▼
   GST Summary      GST Summary
        │               │
        ▼               ▼
    Temp_5          Temp_18
        │               │
        └───────┬───────┘
                ▼
        Revised Excel Workbook

Final Workbook Structure

The generated workbook follows this structure:

5
18
Temp_5
Temp_18
Bill Details
ITEM

**Sheet Descriptions**
Bill Details
The source billing sheet containing transaction-level billing information, including:

Item details
HSN information
Quantity
UOM
Sale Rate
Item Value
GST Rate
CGST
SGST
Final Amount

**5**
Contains transactions identified as 5% GST transactions.

The sheet also includes:
- GST-specific transaction rows
- Dynamic calculations
- CGST summary
- SGST summary

**18**
Contains transactions identified as 18% GST transactions.

The sheet also includes:
- GST-specific transaction rows
- Freight transactions
- Dynamic calculations
- CGST summary
- SGST summary

**Temp_5**
A structured data-preparation sheet for 5% GST transactions.

The sheet includes fields such as:
- Ctr
- Voucher Date
- Voucher Type Name
- Voucher Number
- Ledger Name
- Ledger Amount
- Ledger Amount Dr/Cr
- Item Name
- Item Allocations - Godown Name
- Actual Quantity
- Billed Quantity
- Quantity UOM
- Item Rate
- Item Rate per
- Item Amount
- Change Mode

The row mapping follows:
Ctr 0  → Farmer Name / Ledger
Ctr 1  → Sheet 5, Sl 1
Ctr 2  → Sheet 5, Sl 2
...
Ctr N  → Sheet 5, Sl N

**Temp_18**
A structured data-preparation sheet for 18% GST transactions.

It follows the same dynamic mapping approach:
Ctr 0  → Farmer Name / Ledger
Ctr 1  → Sheet 18, Sl 1
Ctr 2  → Sheet 18, Sl 2
...
Ctr N  → Sheet 18, Sl N

The final special rows include:
- Freight charges
- Output CGST
- Output SGST


**Technology Stack**
Python
OpenPyXL
Tkinter
pathlib
logging
Windows file dialog
