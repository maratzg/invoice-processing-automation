# Invoice Processing Automation

A Python-based document processing and validation pipeline that converts logistics invoice PDFs into structured, validated data suitable for import into the MONEO accounting system.

The project focuses on safe financial-data automation: invoices are parsed, validated against configurable business rules and mappings, transformed into accounting import rows, and written to Excel for human review before import.

## Overview

Manual invoice processing often involves repeatedly reading PDF invoices, copying accounting fields, checking VAT treatment, matching customers and service articles, and entering the resulting data into an accounting system.

This project automates much of that workflow:

```text

PDF Invoice

    ↓

Text Extraction

    ↓

Invoice Parsing

    ↓

Structured Data Model

    ↓

Business Validation

    ↓

Customer / VAT / Article Mapping

    ↓

MONEO Import Row Generation

    ↓

Excel Output

    ↓

Human Review

```

The system deliberately does not automatically post transactions into the accounting system. Validation failures are rejected or flagged for review rather than silently accepted.

## Key Features

- PDF invoice text extraction using `pdfplumber` with optional PyMuPDF support

- Structured invoice, container, customer, VAT and accounting data models

- Logistics-specific invoice parsing

- Container and shipment information extraction

- Configurable customer, VAT and article mappings

- VAT and invoice-total validation

- Duplicate invoice detection

- Invalid date and negative amount detection

- Multi-VAT invoice handling

- Grouping of accounting rows by VAT rate

- Template-based Excel output generation

- Formula and template structure preservation

- Atomic output-file replacement

- Processing and validation reporting

- Regression tests for business-critical rules

## Project Structure

```text

invoice-processing-automation/

├── src/

│   ├── config.py

│   ├── exceptions.py

│   ├── file_utils.py

│   ├── invoice_models.py

│   ├── logging_setup.py

│   ├── main.py

│   ├── mapping_loader.py

│   ├── moneo_writer.py

│   ├── pdf_parser.py

│   └── validator.py

├── tests/

│   └── test_business_rules.py

├── mappings/

│   ├── Article_Mapping.xlsx

│   ├── Client_Mapping.xlsx

│   └── VAT_Mapping.xlsx

├── input_pdfs/

├── output/

├── logs/

├── templates/

├── config.example.json

├── requirements.txt

└── README.md

```

## Validation Philosophy

Financial automation should fail safely.

An invoice is not considered ready for import merely because some fields were successfully extracted. The validation layer checks conditions including:

- required invoice fields

- customer mapping

- VAT mapping

- article mapping

- invoice totals

- calculated VAT

- invoice and due dates

- negative or impossible amounts

- container presence

- duplicate invoice numbers

- unsupported or uncertain VAT treatment

Cases that cannot be resolved confidently are rejected or marked for manual review.

## Mapping System

Business-specific accounting rules are separated from the parser through Excel mapping files.

### Client Mapping

Maps invoice customer information to the corresponding accounting-system customer.

### VAT Mapping

Determines the appropriate VAT treatment according to customer type, country and VAT rate.

Potentially ambiguous VAT cases can explicitly require manual review.

### Article Mapping

Maps invoice VAT groups to accounting article/service definitions.

The mapping files included in this repository contain synthetic/example values only.

## Configuration

Copy the example configuration:

```bash

cp config.example.json config.json

```

Adjust the paths and settings for your environment.

`config.json` is intentionally ignored by Git so local or production configuration is not committed.

## MONEO Template

The writer is designed to work with an authorized MONEO invoice-import Excel template.

The official template is **not distributed with this repository**.

Place an authorized template at:

```text

templates/Rekinu_Imp_2024.xlsx

```

before using template-dependent import generation.

## Installation

Create a virtual environment:

```bash

python3 -m venv .venv

source .venv/bin/activate

```

Install dependencies:

```bash

pip install -r requirements.txt

```

## Running Tests

Run the regression suite with:

```bash

python3 -m unittest discover -s tests -v

```

The public test suite currently covers 11 business-rule and integration scenarios, including:

- invoices without containers

- incomplete VAT mappings

- shipping-table parsing

- VAT-grouped accounting rows

- customer matching by registration number

- invalid invoice/due-date relationships

- duplicate invoices

- negative line amounts

- template row/formula preservation

- template VAT-code consistency

- template column mapping

All test fixtures in the public repository use synthetic data.

## Limitations

This is not a universal invoice parser.

PDF layouts vary significantly between suppliers, and scanned/image-only invoices require OCR or another extraction strategy.

Accounting and VAT mappings must also be configured and reviewed for the organization using the system. The software should not be treated as a substitute for accounting or tax review.

## Portfolio Notice

This repository is an independent portfolio representation inspired by real-world logistics and accounting automation workflows.

It contains synthetic/example data only and does not include employer or customer invoices, operational logs, production outputs, credentials, internal configuration, or proprietary accounting templates.

MONEO is referenced solely as the target system for the integration. This repository is not affiliated with or endorsed by MONEO.

## Tech Stack

- Python

- `pdfplumber`

- PyMuPDF

- `openpyxl`

- Excel

- `unittest`

## Engineering Focus

The main engineering challenge was not simply extracting text from PDFs. It was building a pipeline where uncertain or inconsistent financial data could not silently become accounting input.

The project therefore emphasizes validation, explicit mappings, traceable intermediate data, deterministic transformation rules, safe file handling and human review.

