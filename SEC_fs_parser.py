import requests
import pandas as pd
from datetime import datetime, timedelta
import time
from openpyxl import load_workbook
import numpy as np
import json
from typing import Dict, List, Optional

class CompanyInfo:
    # Dictionary of company name to CIK mapping
    CIK_MAPPING = {
        "ADM": "0000007084",        # Archer-Daniels-Midland Company
        "AEP": "0000004904",        # American Electric Power
        "BP": "0000313807",         # BP plc
        "DTE": "0000936340",        # DTE Energy
        "MPC": "0001510295",        # Marathon Petroleum Corporation
        "PSX": "0001534701",        # Phillips 66
        "PCG": "0001004980",        # Pacific Gas and Electric
        # Add more companies as needed
    }
    
    @staticmethod
    def get_cik(company_name: str) -> Optional[str]:
        """Get CIK number for a company"""
        return CompanyInfo.CIK_MAPPING.get(company_name.upper())

class SECDataFetcher:
    def __init__(self, company_name: str, start_year: int, end_year: int):
        self.company_name = company_name
        self.cik = CompanyInfo.get_cik(company_name)
        if not self.cik:
            raise ValueError(f"Company {company_name} not found in CIK mapping")
            
        self.start_date = f"{start_year}-01-01"
        self.end_date = f"{end_year}-12-31"
        self.base_url = "https://data.sec.gov/api/xbrl/companyfacts/CIK{cik}.json"
        self.headers = {
            'User-Agent': 'Credit Rating Tool (yk3057@columbia.edu)',
            'Accept-Encoding': 'gzip, deflate',
            'Host': 'data.sec.gov',
            'Accept': 'application/json'
        }
        self.template_path = r"C:\Users\ykim\Downloads\Credit Rating Template NEW (1).xlsx"
        self.output_path = f"financial_statements_{company_name}_{start_year}_{end_year}.txt"
        
        # Initialize storage for financial statements
        self.financial_data = {
            'Balance Sheet': {},
            'Income Statement': {},
            'Cash Flow Statement': {}
        }
        
    def fetch_sec_data(self) -> Optional[Dict]:
        """Fetch financial data from SEC EDGAR"""
        cik_padded = str(self.cik).zfill(10)
        url = self.base_url.format(cik=cik_padded)
        
        try:
            response = requests.get(url, headers=self.headers)
            if response.status_code == 200:
                return response.json()
            else:
                print(f"Error fetching SEC data: Status code {response.status_code}")
                return None
        except Exception as e:
            print(f"Error: {str(e)}")
            return None

    def get_annual_values(self, values: List[Dict], start_year: int, end_year: int) -> Dict[int, float]:
        """Get annual values within the specified date range"""
        annual_data = {}
        
        if not values:
            print(f"Warning: No values provided")
            return annual_data

        print(f"Processing {len(values)} data points")
        
        for value in values:
            try:
                # Debug print
                print(f"Processing value entry: {value}")
                
                # Check if we have the required fields
                if not all(key in value for key in ['end', 'val']):
                    print(f"Warning: Missing required fields in value: {value}")
                    continue

                # Some SEC filings might have 'form' field indicating the report type
                form = value.get('form', '').upper()
                if form and not ('10-K' in form or '10-Q' in form):
                    print(f"Skipping non-annual/quarterly report: {form}")
                    continue

                end_date = value['end']
                if not end_date:
                    print(f"Warning: Empty end date in value: {value}")
                    continue

                try:
                    end_date = datetime.strptime(end_date, '%Y-%m-%d')
                except ValueError as e:
                    print(f"Warning: Invalid date format: {end_date} - {e}")
                    continue

                year = end_date.year
                if start_year <= year <= end_year:
                    # If we already have a value for this year, only update if this is a more recent filing
                    if year in annual_data:
                        existing_value = annual_data[year]
                        if abs(value['val']) > abs(existing_value):  # Use the larger value (assuming it's more recent/corrected)
                            print(f"Updating {year} value from {existing_value} to {value['val']}")
                            annual_data[year] = value['val']
                    else:
                        annual_data[year] = value['val']
                        print(f"Added value for {year}: {value['val']}")

            except Exception as e:
                print(f"Warning: Error processing value {value}: {str(e)}")
                continue
        
        if not annual_data:
            print("Warning: No annual data found in the provided values")
        else:
            print(f"Found data for years: {sorted(annual_data.keys())}")
            
        return annual_data

    def process_financial_data(self, sec_data: Dict):
        """Process and organize financial data by statement type"""
        if 'facts' not in sec_data:
            print("No facts found in SEC data")
            return

        print("\nProcessing SEC data...")
        print(f"Available fact categories: {list(sec_data['facts'].keys())}")
        if 'us-gaap' in sec_data['facts']:
            print(f"Available US GAAP concepts: {list(sec_data['facts']['us-gaap'].keys())}")

        # Define statement sections and their concepts
        financial_concepts = {
            'Balance Sheet': {
                'Assets': {
                    'Cash and Cash Equivalents': 'CashAndCashEquivalentsAtCarryingValue',
                    'Short Term Investments': 'MarketableSecuritiesCurrent',
                    'Accounts Receivable': 'AccountsReceivableNetCurrent',
                    'Inventory': 'InventoryNet',
                    'Total Current Assets': 'AssetsCurrent',
                    'Property, Plant & Equipment': 'PropertyPlantAndEquipmentNet',
                    'Goodwill': 'Goodwill',
                    'Total Assets': 'Assets'
                },
                'Liabilities': {
                    'Accounts Payable': 'AccountsPayableCurrent',
                    'Short Term Debt': 'ShortTermBorrowings',
                    'Total Current Liabilities': 'LiabilitiesCurrent',
                    'Long Term Debt': 'LongTermDebtNoncurrent',
                    'Total Liabilities': 'Liabilities'
                },
                'Equity': {
                    'Common Stock': 'CommonStockValue',
                    'Retained Earnings': 'RetainedEarningsAccumulatedDeficit',
                    'Total Shareholders Equity': 'StockholdersEquity'
                }
            },
            'Income Statement': {
                'Revenue': 'Revenues',
                'Cost of Goods Sold': 'CostOfGoodsAndServicesSold',
                'Gross Profit': 'GrossProfit',
                'Operating Income': 'OperatingIncomeLoss',
                'Interest Expense': 'InterestExpense',
                'Income Before Taxes': 'IncomeLossFromContinuingOperationsBeforeIncomeTaxesExtraordinaryItemsNoncontrollingInterest',
                'Net Income': 'NetIncomeLoss'
            },
            'Cash Flow Statement': {
                'Operating Cash Flow': 'NetCashProvidedByUsedInOperatingActivities',
                'Investing Cash Flow': 'NetCashProvidedByUsedInInvestingActivities',
                'Financing Cash Flow': 'NetCashProvidedByUsedInFinancingActivities',
                'Capital Expenditure': 'PaymentsToAcquirePropertyPlantAndEquipment'
            }
        }

        start_year = int(self.start_date[:4])
        end_year = int(self.end_date[:4])

        try:
            # Process each financial statement
            for statement, concepts in financial_concepts.items():
                print(f"\nProcessing {statement}...")
                
                if isinstance(concepts, dict):
                    if any(isinstance(v, dict) for v in concepts.values()):  # Balance Sheet
                        self.financial_data[statement] = {}
                        for section, section_items in concepts.items():
                            print(f"\nProcessing {section}...")
                            self.financial_data[statement][section] = {}
                            for item, concept in section_items.items():
                                print(f"\nLooking for concept: {concept}")
                                if concept in sec_data['facts'].get('us-gaap', {}):
                                    values = sec_data['facts']['us-gaap'][concept]['units'].get('USD', [])
                                    print(f"Found {len(values)} values for {item}")
                                    if values:
                                        annual_values = self.get_annual_values(values, start_year, end_year)
                                        if annual_values:
                                            self.financial_data[statement][section][item] = annual_values
                                            print(f"Found data for {statement} - {section} - {item}")
                                        else:
                                            print(f"No annual values found for {item}")
                                else:
                                    print(f"Concept {concept} not found in SEC data")
                    else:  # Income Statement and Cash Flow
                        for item, concept in concepts.items():
                            print(f"\nLooking for concept: {concept}")
                            if concept in sec_data['facts'].get('us-gaap', {}):
                                values = sec_data['facts']['us-gaap'][concept]['units'].get('USD', [])
                                print(f"Found {len(values)} values for {item}")
                                if values:
                                    annual_values = self.get_annual_values(values, start_year, end_year)
                                    if annual_values:
                                        if item not in self.financial_data[statement]:
                                            self.financial_data[statement][item] = {}
                                        self.financial_data[statement][item] = annual_values
                                        print(f"Found data for {statement} - {item}")
                                    else:
                                        print(f"No annual values found for {item}")
                            else:
                                print(f"Concept {concept} not found in SEC data")
                                        
            # Check if we found any data
            data_found = False
            for statement in self.financial_data.values():
                if statement:
                    data_found = True
                    break
                    
            if not data_found:
                print("\nWarning: No financial data was found for the specified years.")
                print("This might be because:")
                print("1. The company hasn't filed their latest reports yet")
                print("2. The data format in SEC filings is different from expected")
                print("3. The years requested are outside the available range")
                
        except Exception as e:
            print(f"Error processing financial data: {str(e)}")
            import traceback
            print("Full error trace:")
            print(traceback.format_exc())

    def save_to_text_file(self):
        """Save financial statements to a text file with improved formatting"""
        with open(self.output_path, 'w') as f:
            f.write(f"Financial Statements for {self.company_name}\n")
            f.write(f"Period: {self.start_date} to {self.end_date}\n")
            f.write("=" * 80 + "\n\n")

            # Helper function to format numbers in millions
            def format_number(num):
                if isinstance(num, (int, float)):
                    return f"${num/1000000:,.1f}M"
                return "N/A"

            for statement, data in self.financial_data.items():
                f.write(f"{statement}\n")
                f.write("-" * len(statement) + "\n")
                
                if isinstance(data, dict):
                    # Handle nested structure (Balance Sheet)
                    if any(isinstance(v, dict) for v in data.values()):
                        for section, items in data.items():
                            f.write(f"\n{section}:\n")
                            if isinstance(items, dict):
                                for item, values in items.items():
                                    if isinstance(values, dict):
                                        f.write(f"  {item}:\n")
                                        for year, value in sorted(values.items()):
                                            f.write(f"    {year}: {format_number(value)}\n")
                                    else:
                                        f.write(f"  {item}: {format_number(values)}\n")
                            else:
                                f.write(f"  {format_number(items)}\n")
                    else:
                        # Handle flat structure
                        for item, values in data.items():
                            if isinstance(values, dict):
                                f.write(f"\n{item}:\n")
                                for year, value in sorted(values.items()):
                                    f.write(f"  {year}: {format_number(value)}\n")
                            else:
                                f.write(f"{item}: {format_number(values)}\n")
                else:
                    f.write(f"{format_number(data)}\n")
                
                f.write("\n" + "=" * 80 + "\n\n")

    def update_excel_template(self):
        """Update the Excel template with the financial data"""
        # This method will be implemented after verification of the text output
        pass

def main():
    """Process financial statements for all companies"""
    companies = [
        "ADM",  # Archer-Daniels-Midland Company
        "AEP",  # American Electric Power
        "BP",   # BP plc
        "DTE",  # DTE Energy
        "MPC",  # Marathon Petroleum Corporation
        "PSX",  # Phillips 66
        "PCG"   # Pacific Gas and Electric
    ]
    
    start_year = 2020
    end_year = 2024
    
    print(f"Processing {len(companies)} companies for years {start_year}-{end_year}")
    print("=" * 80)
    
    for company in companies:
        try:
            print(f"\nProcessing {company}...")
            fetcher = SECDataFetcher(company, start_year, end_year)
            sec_data = fetcher.fetch_sec_data()
            
            if sec_data:
                fetcher.process_financial_data(sec_data)
                fetcher.save_to_text_file()
                print(f"\nFinancial statements for {company} have been saved to: {fetcher.output_path}")
            else:
                print(f"Failed to fetch data for {company}")
                
        except Exception as e:
            print(f"Error processing {company}: {str(e)}")
        
        print("=" * 80)
        # Add a small delay between requests to be nice to the SEC API
        time.sleep(0.1)
    
    print("\nAll companies processed. Please verify the data in the text files.")

if __name__ == "__main__":
    main()

# Note: Before using this script, please:
# 1. Replace 'Your Company Name' and 'yourname@email.com' in the headers
# 2. Respect SEC's fair access guidelines (https://www.sec.gov/os/accessing-edgar-data)
# 3. Add appropriate rate limiting between requests
