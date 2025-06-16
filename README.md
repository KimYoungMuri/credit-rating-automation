# Credit Rating Automation

This project automates the extraction and analysis of financial statements from PDF documents for credit rating purposes.

## Features

- Automated extraction of financial statements from PDF documents
- Support for Balance Sheet, Income Statement, and Cash Flow Statement
- Intelligent template mapping for standardized financial data
- Semantic matching using FinBERT for accurate line item mapping
- Excel output generation with formatted financial data

## Installation

1. Clone the repository:
```bash
git clone https://github.com/KimYoungMuri/credit-rating-automation.git
cd credit-rating-automation
```

2. Create a virtual environment and activate it:
```bash
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate
```

3. Install the required packages:
```bash
pip install -r requirements.txt
```

## Usage

1. Place your PDF financial statements in the `input_pdfs` directory
2. Run the main script:
```bash
python src/priv_financials_extractor/run_extractor.py
```

3. The processed data will be saved in the `output_excel` directory

## Project Structure

```
credit-rating-automation/
├── src/
│   └── priv_financials_extractor/
│       ├── final_extractor.py      # PDF text extraction
│       ├── final_find_fs.py        # Financial statement detection
│       ├── final_template_mapper.py # Template mapping
│       └── run_extractor.py        # Main script
├── input_pdfs/                     # Input PDF directory
├── output_excel/                   # Output Excel directory
├── requirements.txt                # Project dependencies
└── README.md                       # Project documentation
```

## Dependencies

- Python 3.8+
- pdfplumber
- pandas
- openpyxl
- transformers
- torch
- numpy

## License

This project is licensed under the MIT License - see the LICENSE file for details.