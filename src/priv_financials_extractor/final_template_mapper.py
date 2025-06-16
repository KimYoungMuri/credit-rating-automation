'''
Line Items to be filled for 2 years worth of data:

ASSETS
Cash and equivalents
Accounts Receivable
Prepaid Expenses
Inventory
Investments
Other
Total Current Assets

Net PPE
Goodwill
Intangibles
Other

Total Non Current Assets
Total Assets

LIABILITIES
Accounts Payable
Accrued Interest
Short term Borrowing
Current Portion of Long Term Debt
Other
Total Current Liabilities

Long Term Debt
Deferred income taxes
Other

Total Non Current Liabilities
Total Liabilities

EQUITY
Common Stock
Retained Earnings
Paid in Capital
Other
Total Equity

Total Liabilities and Equity

INCOME STATEMENT
Revenue
Operating Expenses
Operating Income

Depreciation (-)
Amortization (-)
Asses gain(loss) impairments
Interest Expense (-)
Interest Income (+)
Other income(expenses)

Income Before Taxes
Tax expense
Other
Net Income

STATEMENT OF CASH FLOW
Operating Activities
Net Income
Changes in noncash items
Changes in Asses and Liabilities
Net Cash from(used) Operating Activities

Investing Activities
CapEx
Proceeds from asset sales
Others
Net cash from(used) for investing 

Financing Activities
Issuance of Debt (long+short term)
Retirement of Debt (long+short term)
Issuance of Stock
Dividends Paid
Other
Net cash from(used) for financing

Net change in Cash
Starting Cash
Ending Cash

'''

import pandas as pd
import numpy as np
from pathlib import Path
import torch
from transformers import BertTokenizer, BertForSequenceClassification
import logging
from typing import Dict, List, Tuple, Optional
import re
import pdfplumber

class TemplateMapper:
    def __init__(self):
        self.setup_logging()
        self.setup_finbert()
        self.load_template_items()
        
    def setup_logging(self):
        """Setup logging configuration"""
        # Remove any existing handlers
        for handler in logging.root.handlers[:]:
            logging.root.removeHandler(handler)
            
        # Create a formatter
        formatter = logging.Formatter('%(message)s')
        
        # Create file handler
        file_handler = logging.FileHandler('template_mapping.log', mode='w', encoding='utf-8')
        file_handler.setFormatter(formatter)
        
        # Create console handler
        console_handler = logging.StreamHandler()
        console_handler.setFormatter(formatter)
        
        # Set up the logger
        self.logger = logging.getLogger(__name__)
        self.logger.setLevel(logging.INFO)
        self.logger.addHandler(file_handler)
        self.logger.addHandler(console_handler)
        
    def log_info(self, message: str):
        """Helper function to log messages both to file and console"""
        self.logger.info(message)
        
    def setup_finbert(self):
        """Initialize FinBERT model for semantic similarity"""
        model_name = "ProsusAI/finbert"
        self.tokenizer = BertTokenizer.from_pretrained(model_name)
        self.model = BertForSequenceClassification.from_pretrained(model_name)
        self.model.eval()
        
    def load_template_items(self):
        """Load standardized template items with common variations"""
        # Define balance sheet items
        balance_sheet_assets = [
            'Cash and equivalents|Cash and cash equivalents',
            'Accounts Receivable|Accounts receivable|Trade receivables',
            'Prepaid Expenses|Prepaid expenses|Prepayments',
            'Inventory|Inventories',
            'Investments|Investment securities',
            'Other current assets',
            'Total Current Assets|Total current assets',
            'Net PPE|Property and equipment|Property, plant and equipment',
            'Goodwill|Goodwill—Net',
            'Intangibles|Intangible assets|OTHER INTANGIBLE ASSETS',
            'Other non-current assets|Other assets|Other noncurrent assets',
            'Total Non Current Assets|Total noncurrent assets',
            'Total Assets|TOTAL'
        ]
        
        balance_sheet_liabilities = [
            'Accounts Payable|Accounts payable|Trade payables',
            'Accrued Interest|Accrued interest payable',
            'Short term Borrowing|Short-term debt|Short-term borrowings',
            'Current Portion of Long Term Debt|Current portion of long-term debt',
            'Other current liabilities|Accrued liabilities|Other Current Liabilities',
            'Total Current Liabilities|Total current liabilities',
            'Long Term Debt|Long-term debt|Long-term borrowings',
            'Deferred income taxes|Deferred tax liabilities',
            'Other non-current liabilities|Other long-term liabilities',
            'Total Non Current Liabilities|Total noncurrent liabilities',
            'Total Liabilities|Total liabilities'
        ]
        
        balance_sheet_equity = [
            'Common Stock|Common stock|Share capital',
            'Retained Earnings|Retained earnings',
            'Paid in Capital|Additional paid-in capital',
            'Other equity|Other comprehensive income|NONCONTROLLING INTERESTS',
            "Total Equity|Total shareholders equity|TOTAL COMMON SHAREHOLDERS' EQUITY|Total stockholders equity",
            'Total Liabilities and Equity|TOTAL|Total liabilities and equity'
        ]
        
        # Define income statement items
        income_statement_items = [
            'Revenue|Net sales|Sales revenue',
            'Operating Expenses|Operating costs and expenses',
            'Operating Income|Operating profit|Income from operations',
            'Depreciation|Depreciation expense|Depreciation and amortization',
            'Amortization|Amortization expense',
            'Asset gain(loss) impairments|Impairment losses|Gain on sale of assets',
            'Interest Expense|Interest expense',
            'Interest Income|Interest income',
            'Other income(expenses)|Other income|Other expense',
            'Income Before Taxes|Income before income taxes',
            'Tax expense|Income tax expense|Provision for income taxes',
            'Other items|Other comprehensive income',
            'Net Income|Net income|Net profit'
        ]
        
        # Define cash flow items
        cash_flow_operating = [
            'Operating Activities|OPERATING ACTIVITIES',
            'Net Income|Net income',
            'Changes in noncash items|Depreciation and amortization',
            'Changes in Assets and Liabilities|Changes in operating assets and liabilities',
            'Net Cash from(used) Operating Activities|Net cash provided by operating activities'
        ]
        
        cash_flow_investing = [
            'Investing Activities|INVESTING ACTIVITIES',
            'CapEx|Capital expenditures|Purchase of property and equipment',
            'Proceeds from asset sales|Proceeds from sale of assets',
            'Other investing items|Other investing activities',
            'Net cash from(used) for investing|Net cash used in investing activities'
        ]
        
        cash_flow_financing = [
            'Financing Activities|FINANCING ACTIVITIES',
            'Issuance of Debt|Proceeds from borrowings|Proceeds from long-term debt',
            'Retirement of Debt|Repayment of debt|Payments of long-term debt',
            'Issuance of Stock|Proceeds from issuance of stock',
            'Dividends Paid|Dividend payments|Distributions to shareholders',
            'Other financing items|Other financing activities',
            'Net cash from(used) for financing|Net cash provided by financing activities',
            'Net change in Cash|Net increase (decrease) in cash',
            'Starting Cash|Cash at beginning of year|Beginning of year',
            'Ending Cash|Cash at end of year|End of year'
        ]
        
        # Construct the template items dictionary
        self.template_items = {
            'balance_sheet': {
                'assets': balance_sheet_assets,
                'liabilities': balance_sheet_liabilities,
                'equity': balance_sheet_equity
            },
            'income_statement': income_statement_items,
            'cash_flow': {
                'operating': cash_flow_operating,
                'investing': cash_flow_investing,
                'financing': cash_flow_financing
            }
        }
        
    def get_semantic_similarity(self, template_item: str, actual_text: str) -> float:
        """Get semantic similarity between template item and actual text"""
        # First check for exact matches in the template variations
        template_variations = template_item.split('|')
        for variation in template_variations:
            # Case insensitive exact match
            if variation.lower() == actual_text.lower():
                return 1.0
            # Case insensitive contains match
            if variation.lower() in actual_text.lower() or actual_text.lower() in variation.lower():
                return 0.9
        
        # If no exact match, use BERT for semantic similarity
        try:
            # Use the first variation for BERT comparison
            inputs = self.tokenizer(template_variations[0], actual_text, return_tensors="pt", padding=True, truncation=True)
            with torch.no_grad():
                outputs = self.model(**inputs)
                logits = outputs.logits
                probabilities = torch.nn.functional.softmax(logits, dim=1)
                similarity = probabilities[0][1].item()  # Probability of being similar
            return similarity
        except Exception as e:
            print(f"Error in semantic similarity: {str(e)}")
            return 0.0
        
    def find_best_match(self, source_item: str, target_items: List[str], threshold: float = 0.6) -> Optional[str]:
        """Find the best matching template item for a given source item"""
        best_match = None
        best_score = -1
        
        for target in target_items:
            # Get primary template name (first variation)
            primary_name = target.split('|')[0]
            score = self.get_semantic_similarity(target, source_item)
            if score > best_score and score >= threshold:
                best_score = score
                best_match = primary_name
                
        return best_match
        
    def process_excel(self, excel_path: str) -> Dict:
        """Process the extracted Excel file and map to template items"""
        try:
            self.log_info("\nStarting template mapping process...")
            
            # Define the correct page numbers for US Venture 2024
            correct_pages = {
                'balance_sheet': [7, 8],
                'income_statement': [9],
                'cash_flow': [11, 12]
            }
            
            # Read all sheets from Excel file
            all_sheets = pd.read_excel(excel_path, sheet_name=None)
            
            # Debug: Print available sheets
            self.log_info("\nDEBUG: Available sheets in Excel file:")
            for sheet_name in all_sheets.keys():
                self.log_info(f"Sheet: {sheet_name}")
                df = all_sheets[sheet_name]
                self.log_info(f"Number of rows: {len(df)}")
                self.log_info("First few rows:")
                self.log_info(str(df.head()))
                self.log_info("Columns:")
                self.log_info(str(df.columns.tolist()))
                self.log_info("-" * 50)
            
            # Initialize results dictionary with both years
            results = {
                'balance_sheet': {'2024': {}, '2023': {}},
                'income_statement': {'2024': {}, '2023': {}},
                'cash_flow': {'2024': {}, '2023': {}}
            }
            
            # Create a set to track used rows to prevent duplicates
            used_rows = {}  # Dictionary to track used rows per sheet
            
            # Process Balance Sheet
            if 'Balance Sheet' in all_sheets:
                df = all_sheets['Balance Sheet']
                used_rows['Balance Sheet'] = set()
                self.log_info(f"\nProcessing Balance Sheet items ({len(df)} rows)...")
                
                for section, items in self.template_items['balance_sheet'].items():
                    self.log_info(f"\nSearching for {section} items:")
                    for template_item in items:
                        self.log_info(f"Looking for: {template_item.split('|')[0]}")  # Show primary variation
                        best_match = None
                        best_score = -1
                        best_row_idx = None
                        
                        # Search through all rows that haven't been used
                        for idx, row in df.iterrows():
                            if idx in used_rows['Balance Sheet']:
                                continue
                                
                            description = str(row['Description']).strip()
                            if not description or description == 'nan':
                                continue
                                
                            score = self.get_semantic_similarity(template_item, description)
                            if score > best_score and score >= 0.6:  # Lower threshold
                                best_score = score
                                best_match = description
                                best_row_idx = idx
                        
                        if best_match:
                            value_2024 = df.loc[best_row_idx, 'Value_1']
                            value_2023 = df.loc[best_row_idx, 'Value_2']
                            self.log_info(f"Found match: '{best_match}' (score: {best_score:.2f}) = {value_2024} (2024), {value_2023} (2023)")
                            results['balance_sheet']['2024'][template_item.split('|')[0]] = value_2024
                            results['balance_sheet']['2023'][template_item.split('|')[0]] = value_2023
                            used_rows['Balance Sheet'].add(best_row_idx)
                        else:
                            self.log_info(f"No match found for {template_item.split('|')[0]}")
            
            # Process Income Statement
            if 'Income Statement' in all_sheets:
                df = all_sheets['Income Statement']
                used_rows['Income Statement'] = set()
                self.log_info(f"\nProcessing Income Statement items ({len(df)} rows)...")
                
                for template_item in self.template_items['income_statement']:
                    self.log_info(f"Looking for: {template_item.split('|')[0]}")
                    best_match = None
                    best_score = -1
                    best_row_idx = None
                    
                    for idx, row in df.iterrows():
                        if idx in used_rows['Income Statement']:
                            continue
                            
                        description = str(row['Description']).strip()
                        if not description or description == 'nan':
                            continue
                            
                        score = self.get_semantic_similarity(template_item, description)
                        if score > best_score and score >= 0.6:
                            best_score = score
                            best_match = description
                            best_row_idx = idx
                    
                    if best_match:
                        value_2024 = df.loc[best_row_idx, 'Value_1']
                        value_2023 = df.loc[best_row_idx, 'Value_2']
                        self.log_info(f"Found match: '{best_match}' (score: {best_score:.2f}) = {value_2024} (2024), {value_2023} (2023)")
                        results['income_statement']['2024'][template_item.split('|')[0]] = value_2024
                        results['income_statement']['2023'][template_item.split('|')[0]] = value_2023
                        used_rows['Income Statement'].add(best_row_idx)
                    else:
                        self.log_info(f"No match found for {template_item.split('|')[0]}")
            
            # Process Cash Flow Statement
            if 'Cash Flow' in all_sheets:
                df = all_sheets['Cash Flow']
                used_rows['Cash Flow'] = set()
                self.log_info(f"\nProcessing Cash Flow items ({len(df)} rows)...")
                
                for section, items in self.template_items['cash_flow'].items():
                    self.log_info(f"\nSearching for {section} items:")
                    for template_item in items:
                        self.log_info(f"Looking for: {template_item.split('|')[0]}")
                        best_match = None
                        best_score = -1
                        best_row_idx = None
                        
                        for idx, row in df.iterrows():
                            if idx in used_rows['Cash Flow']:
                                continue
                                
                            description = str(row['Description']).strip()
                            if not description or description == 'nan':
                                continue
                                
                            score = self.get_semantic_similarity(template_item, description)
                            if score > best_score and score >= 0.6:
                                best_score = score
                                best_match = description
                                best_row_idx = idx
                        
                        if best_match:
                            value_2024 = df.loc[best_row_idx, 'Value_1']
                            value_2023 = df.loc[best_row_idx, 'Value_2']
                            self.log_info(f"Found match: '{best_match}' (score: {best_score:.2f}) = {value_2024} (2024), {value_2023} (2023)")
                            results['cash_flow']['2024'][template_item.split('|')[0]] = value_2024
                            results['cash_flow']['2023'][template_item.split('|')[0]] = value_2023
                            used_rows['Cash Flow'].add(best_row_idx)
                        else:
                            self.log_info(f"No match found for {template_item.split('|')[0]}")
            
            self.log_info("\nTemplate mapping completed.")
            '''
            for stmt_type, years in results.items():
                total_items = len(self.template_items[stmt_type]) if isinstance(self.template_items[stmt_type], list) else sum(len(items) for items in self.template_items[stmt_type].values())
                self.log_info(f"\n{stmt_type.upper()}:")
                for year in ['2024', '2023']:
                    self.log_info(f"\n{year}:")
                    self.log_info(f"Found {len(years[year])} matches out of {total_items} items")
                    for item, value in years[year].items():
                        self.log_info(f"  {item}: {value}")
            '''
            self.log_info("\nSummary of Results:")
            
                    
            return results
            
        except Exception as e:
            self.log_info(f"Error processing Excel file: {str(e)}")
            import traceback
            self.log_info(traceback.format_exc())
            return None
            
def main():
    # Example usage
    mapper = TemplateMapper()
    
    # Get the most recent Excel file from the output directory
    output_dir = Path("output_excel")
    if not output_dir.exists():
        print("No output directory found")
        return
        
    excel_files = list(output_dir.glob("*.xlsx"))
    if not excel_files:
        print("No Excel files found in output directory")
        return
        
    # Sort by creation time and get most recent
    latest_file = max(excel_files, key=lambda x: x.stat().st_ctime)
    print(f"\nProcessing {latest_file}")
    
    # First run final_find_fs.py to get correct page numbers
    from final_find_fs import FinancialStatementFinder
    finder = FinancialStatementFinder()
    input_dir = Path("input_pdfs")
    pdf_path = input_dir / "US_Venture_2024.pdf"
    
    if not pdf_path.exists():
        print(f"Could not find {pdf_path}")
        return
        
    # Get the correct page numbers
    with pdfplumber.open(pdf_path) as pdf:
        final_scores = finder.extractContent(str(pdf_path))
        
        # Get pages with high confidence (>= 80%)
        statements = {}
        for stmt_type, pages in final_scores.items():
            high_conf_pages = {page: score for page, score in pages.items() if score >= 80}
            if high_conf_pages:
                statements[stmt_type] = sorted(high_conf_pages.keys())
    
    # Process the file with correct page numbers
    results = mapper.process_excel(str(latest_file))
    
    if results:
        print("\nMapping Results:")
        print("=" * 50)
        
        for stmt_type, years in results.items():
            print(f"\n{stmt_type.replace('_', ' ').title()}:")
            for year in ['2024', '2023']:
                print(f"\n{year}:")
                for item, value in years[year].items():
                    print(f"{item}: {value}")
                
if __name__ == "__main__":
    main()

