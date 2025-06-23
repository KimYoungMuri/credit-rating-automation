# Page Confirmation System

## Overview

The page confirmation system has been implemented to allow users to verify and modify the automatically detected financial statement pages before proceeding with text extraction. This ensures higher accuracy in the extraction process by giving users control over which pages are processed.

## Features

### 1. Interactive Confirmation
- Displays all detected pages with confidence scores for all statement types
- Shows all detected pages for each statement type with their scores
- Single confirmation prompt for all statements at once

### 2. User Options
The system follows this flow:

1. **Display All Detections**: Shows detected pages for all 3 financial statements
2. **Single Confirmation**: Asks one "y/n" question for all statements
3. **If "y"**: Accepts all detected pages as-is
4. **If "n"**: Prompts for custom page numbers for each statement individually

### 3. Custom Page Entry
- When "n" is selected, users can enter custom page numbers for each statement
- Supports comma-separated page numbers (e.g., "5,6,7")
- Users can press Enter to skip a statement type
- Invalid input allows retry or skip option

## Integration Points

The confirmation system is integrated into the following scripts:

### 1. `final_find_fs.py`
- **New Method**: `confirm_statement_pages(filename)` 
- **Returns**: Dictionary of confirmed pages in format `{stmt_type: [page_nums]}`
- **Usage**: Called from main function after page detection

### 2. `final_extractor.py`
- **Updated**: Main function now uses confirmation system
- **Flow**: Detection → Confirmation → Extraction

### 3. `run_extractor.py`
- **Updated**: Main workflow now includes confirmation step
- **Flow**: Detection → Confirmation → Extraction → Template Mapping

### 4. `test_extraction.py`
- **Updated**: Test script now uses confirmation system
- **Purpose**: Testing the complete pipeline with user verification

## Usage Examples

### Running with Confirmation
```bash
# Run the main extractor (will prompt for confirmation)
python run_extractor.py

# Run individual components
python final_extractor.py
python final_find_fs.py your_file.pdf
```

### Sample Confirmation Session
```
============================================================
FINANCIAL STATEMENT PAGE CONFIRMATION
============================================================
File: example.pdf

Detected Financial Statement Pages:
----------------------------------------

Balance Sheet:
  Detected Pages: [5, 6]
  Confidence: 85.2%
  All detected pages with scores:
    Page 5: 85.2% (High)
    Page 6: 72.1% (Medium)

Income Statement:
  Detected Pages: [7]
  Confidence: 78.9%
  All detected pages with scores:
    Page 7: 78.9% (Medium)

Cash Flow Statement: Not detected

============================================================
CONFIRMATION
============================================================
Confirm all detected pages above? (y/n): n

Entering custom page numbers...

Enter page numbers for Balance Sheet (comma-separated, or press Enter to skip): 5,6,7
✓ Custom Balance Sheet: [5, 6, 7]

Enter page numbers for Income Statement (comma-separated, or press Enter to skip): 8,9
✓ Custom Income Statement: [8, 9]

Enter page numbers for Cash Flow Statement (comma-separated, or press Enter to skip): 10,11
✓ Custom Cash Flow Statement: [10, 11]

============================================================
CONFIRMATION SUMMARY:
============================================================
Balance Sheet: [5, 6, 7]
Income Statement: [8, 9]
Cash Flow Statement: [10, 11]
```

## Data Flow

1. **Detection**: `FinancialStatementFinder.extractContent()` detects pages
2. **Display**: Shows all detected pages with scores for all statements
3. **Confirmation**: Single y/n prompt for all statements
4. **Custom Entry**: If "n", prompts for custom pages for each statement
5. **Extraction**: Confirmed pages are passed to text extraction
6. **Mapping**: Extracted data is mapped to templates

## Benefits

1. **Higher Accuracy**: User verification reduces false positives/negatives
2. **Efficiency**: Single confirmation prompt instead of multiple questions
3. **Flexibility**: Users can correct detection errors or add missing pages
4. **Transparency**: Users see confidence scores and all detected options
5. **Manual Override**: Support for manual page specification
6. **Pipeline Integration**: Seamlessly integrated into existing workflow

## Testing

Use the test script to verify the confirmation system:
```bash
python test_confirmation.py
```

This will run the detection and confirmation process on the most recent PDF in the input_pdfs directory. 