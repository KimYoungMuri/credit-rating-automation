# Private Financial Statement Extractor

## Overview
This module provides automated tools for extracting and analyzing financial statements from PDF documents. It's designed to handle various formats of financial statements and extract structured data for further analysis.

## Components

### FinancialStatementFinder (`find_fs.py`)
- Locates financial statements within PDF documents
- Supports detection of:
  - Balance Sheets
  - Income Statements
  - Cash Flow Statements
- Uses advanced NLP techniques for accurate identification
- Handles document continuations and multi-page statements

### TableExtractor (`extract_tables.py`)
- Extracts tabular data from identified financial statements
- Features multiple extraction strategies:
  - Text-based extraction
  - Line-based extraction
  - Mixed strategy approach
- Includes data cleaning and validation
- Exports to Excel format

## Usage

```python
from priv_financials_extractor import FinancialStatementFinder, TableExtractor

# Initialize the finder
finder = FinancialStatementFinder()

# Find statements in a PDF
statements = finder.find_statements("path/to/document.pdf")

# Extract tables
extractor = TableExtractor()
tables, pages = extractor.extract_tables("path/to/document.pdf")
```

## Installation
All required dependencies are listed in the root `requirements.txt` file.

## Dependencies
- pdfplumber
- pandas
- torch
- transformers
- numpy
- spacy
- python-Levenshtein
- fuzzywuzzy
- openpyxl 