import requests
import json
import pandas as pd
import openpyxl
from pathlib import Path
from datetime import datetime
import re
from typing import Dict, List, Tuple, Optional
from collections import defaultdict

class ComprehensiveMapper:
    """
    Comprehensive mapper that processes entire financial statements in single LLM calls.
    Takes extracted tables and outputs complete mapped templates.
    """
    
    def __init__(self, ollama_url: str = "http://localhost:11434"):
        self.ollama_url = ollama_url
        self.model_name = "mistral:latest"
        
    def check_ollama_available(self) -> bool:
        """Check if Ollama is running and available"""
        try:
            response = requests.get(f"{self.ollama_url}/api/tags", timeout=5)
            return response.status_code == 200
        except:
            return False
    
    def create_balance_sheet_prompt(self, extracted_data: Dict, template_structure: Dict) -> str:
        """
        Create a comprehensive prompt for balance sheet mapping.
        """
        # Format extracted data
        years = list(extracted_data.keys())
        items_data = []
        
        for year in years:
            year_data = extracted_data[year]
            for desc, val in year_data.items():
                if not self.is_total_row(desc):
                    items_data.append({
                        'description': desc,
                        'year': year,
                        'value': val
                    })
        
        items_formatted = "\n".join([
            f"- {item['description']} ({item['year']}): {item['value']}"
            for item in items_data
        ])
        
        # Format template structure
        template_formatted = "\n".join([
            f"Section: {section}"
            for section, items in template_structure.items()
        ])
        
        prompt = f"""You are an expert financial analyst. Your task is to map an entire balance sheet to a standardized template.

EXTRACTED BALANCE SHEET DATA:
{items_formatted}

TEMPLATE STRUCTURE:
{template_formatted}

INSTRUCTIONS:
1. Analyze each line item and map it to the appropriate template section and row given its meaning and relative location in the balance sheet. 
2. Consider synonyms, abbreviations, and common variations
3. Group related items appropriately (e.g., multiple cash accounts → "Cash and equivalents")
4. Handle items that don't have exact matches by placing them in "Other" categories
5. Ensure all values are used without skip and assigned to the correct years

MAPPING RULES:
- Cash and cash equivalents: Map all cash-related items here
- Accounts receivable: Include trade receivables, notes receivable, etc.
- Net PPE: Property, plant, equipment (net of depreciation)
- Investments: Derivative assets, 
- Goodwill: Intangible assets like goodwill
- Accounts payable: Trade payables, accrued expenses
- Long-term debt: All long-term debt obligations
- Equity items: Common stock, retained earnings, paid-in capital

OUTPUT FORMAT:
Return a JSON object with this structure:
{{
  "mappings": {{
    "template_section": {{
      "template_row": {{
        "2022": value,
        "2023": value
      }}
    }}
  }},
  "unmapped_items": [
    {{
      "description": "item description",
      "year": "2022",
      "value": value,
      "reason": "why it couldn't be mapped"
    }}
  ]
}}

Example:
{{
  "mappings": {{
    "current_assets": {{
      "Cash and equivalents": {{
        "2022": 1500000,
        "2023": 1800000
      }},
      "Accounts Receivable": {{
        "2022": 2500000,
        "2023": 2800000
      }}
    }}
  }},
  "unmapped_items": [
    {{
      "description": "Unusual asset",
      "year": "2022", 
      "value": 50000,
      "reason": "No clear template match"
    }}
  ]
}}

Your response:"""
        
        return prompt
    
    def create_income_statement_prompt(self, extracted_data: Dict, template_structure: Dict) -> str:
        """
        Create a comprehensive prompt for income statement mapping.
        """
        # Format extracted data
        years = list(extracted_data.keys())
        items_data = []
        
        for year in years:
            year_data = extracted_data[year]
            for desc, val in year_data.items():
                if not self.is_total_row(desc):
                    items_data.append({
                        'description': desc,
                        'year': year,
                        'value': val
                    })
        
        items_formatted = "\n".join([
            f"- {item['description']} ({item['year']}): {item['value']}"
            for item in items_data
        ])
        
        prompt = f"""You are an expert financial analyst. Your task is to map an entire income statement to a standardized template.

EXTRACTED INCOME STATEMENT DATA:
{items_formatted}

TEMPLATE SECTIONS:
- Revenue: Sales, revenue, income from operations
- Operating Expenses: Cost of goods sold, operating expenses, depreciation, amortization
- Other Income/Expense: Interest income/expense, other income/expenses
- Tax and Net Income: Income tax expense, net income

INSTRUCTIONS:
1. Map each line item to the appropriate template section and row
2. Consider synonyms and common variations
3. Group related items appropriately
4. Handle items without exact matches by placing them in "Other" categories
5. Ensure proper sign conventions (expenses as negative, income as positive)

MAPPING RULES:
- Revenue: All sales and revenue items
- Cost of Sales: Direct costs related to revenue generation
- Operating Expenses: Selling, general, administrative expenses
- Depreciation/Amortization: All depreciation and amortization items
- Interest Expense: All interest-related expenses
- Income Tax: Tax expenses
- Net Income: Final net income/loss

OUTPUT FORMAT:
Return a JSON object with this structure:
{{
  "mappings": {{
    "revenue": {{
      "Revenue": {{
        "2022": value,
        "2023": value
      }}
    }},
    "operating_expenses": {{
      "Cost of Sales": {{
        "2022": value,
        "2023": value
      }}
    }}
  }},
  "unmapped_items": [
    {{
      "description": "item description",
      "year": "2022",
      "value": value,
      "reason": "why it couldn't be mapped"
    }}
  ]
}}

Your response:"""
        
        return prompt
    
    def create_cash_flow_prompt(self, extracted_data: Dict, template_structure: Dict) -> str:
        """
        Create a comprehensive prompt for cash flow statement mapping.
        """
        # Format extracted data
        years = list(extracted_data.keys())
        items_data = []
        
        for year in years:
            year_data = extracted_data[year]
            for desc, val in year_data.items():
                if not self.is_total_row(desc):
                    items_data.append({
                        'description': desc,
                        'year': year,
                        'value': val
                    })
        
        items_formatted = "\n".join([
            f"- {item['description']} ({item['year']}): {item['value']}"
            for item in items_data
        ])
        
        prompt = f"""You are an expert financial analyst. Your task is to map an entire cash flow statement to a standardized template.

EXTRACTED CASH FLOW DATA:
{items_formatted}

TEMPLATE SECTIONS:
- Operating Activities: Net income, depreciation, changes in working capital
- Investing Activities: Capital expenditures, asset purchases/sales, investments
- Financing Activities: Debt issuance/repayment, stock issuance/repurchase, dividends
- Cash Reconciliation: Net change in cash, beginning/ending cash balances

INSTRUCTIONS:
1. Map each line item to the appropriate template section and row
2. Consider synonyms and common variations
3. Group related items appropriately
4. Handle items without exact matches by placing them in "Other" categories
5. Ensure proper sign conventions (cash outflows as negative, inflows as positive)

MAPPING RULES:
- Net Income: Starting point for operating activities
- Depreciation/Amortization: Non-cash expenses added back
- Changes in Working Capital: Changes in current assets/liabilities
- Capital Expenditures: Property and equipment purchases
- Debt Activities: Borrowing and repayment
- Stock Activities: Issuance and repurchase
- Dividends: Dividend payments

OUTPUT FORMAT:
Return a JSON object with this structure:
{{
  "mappings": {{
    "operating_activities": {{
      "Net profit (loss)": {{
        "2022": value,
        "2023": value
      }}
    }},
    "investing_activities": {{
      "Purchases of property and equipment": {{
        "2022": value,
        "2023": value
      }}
    }}
  }},
  "unmapped_items": [
    {{
      "description": "item description",
      "year": "2022",
      "value": value,
      "reason": "why it couldn't be mapped"
    }}
  ]
}}

Your response:"""
        
        return prompt
    
    def is_total_row(self, description: str) -> bool:
        """Check if description is a total/subtotal row"""
        desc_lower = description.lower().strip()
        total_patterns = [
            r'^total\s',
            r'\stotal$',
            r'^sum\s',
            r'\ssum$',
            r'^subtotal',
            r'\ssubtotal$',
            r'^net\s(?!income|loss)',
            r'\snet$'
        ]
        return any(re.search(pattern, desc_lower) for pattern in total_patterns)
    
    def call_ollama_comprehensive(self, prompt: str, timeout: int = 120) -> Optional[Dict]:
        """
        Call Ollama with comprehensive prompt and parse JSON response.
        """
        try:
            print(f"[DEBUG] Sending comprehensive prompt to Ollama (length: {len(prompt)} chars)")
            
            response = requests.post(
                f"{self.ollama_url}/api/generate",
                json={
                    "model": self.model_name,
                    "prompt": prompt,
                    "stream": False,
                    "options": {
                        "temperature": 0.1,
                        "top_p": 0.9,
                        "max_tokens": 4096  # Large response for comprehensive mapping
                    }
                },
                timeout=timeout
            )
            
            if response.status_code != 200:
                print(f"[ERROR] Ollama API error: {response.status_code}")
                return None
            
            result = response.json()
            response_text = result.get("response", "").strip()
            
            # Parse JSON response
            try:
                # Find JSON object in response
                json_match = re.search(r'\{.*\}', response_text, re.DOTALL)
                if json_match:
                    json_str = json_match.group(0)
                    parsed = json.loads(json_str)
                    print(f"[DEBUG] Successfully parsed comprehensive response")
                    return parsed
                else:
                    print(f"[ERROR] No JSON found in response")
                    return None
                    
            except json.JSONDecodeError as e:
                print(f"[ERROR] Failed to parse JSON response: {e}")
                print(f"[DEBUG] Response text: {response_text[:500]}...")
                return None
                
        except Exception as e:
            print(f"[ERROR] Comprehensive Ollama call failed: {e}")
            return None
    
    def get_template_structure(self, statement_type: str) -> Dict:
        """
        Get the template structure for a given statement type.
        """
        if statement_type == 'balance_sheet':
            return {
                'current_assets': [
                    'Cash and equivalents',
                    'Accounts Receivable', 
                    'Prepaid Expenses',
                    'Inventory',
                    'Investments',
                    'Other'
                ],
                'noncurrent_assets': [
                    'Net PPE',
                    'Goodwill',
                    'Intangibles',
                    'Other'
                ],
                'current_liabilities': [
                    'Accounts Payable',
                    'Accrued Interest',
                    'Short term Borrowing',
                    'Current Portion of Long Term Debt',
                    'Other'
                ],
                'noncurrent_liabilities': [
                    'Long Term Debt',
                    'Deferred income taxes',
                    'Other'
                ],
                'equity': [
                    'Common Stock',
                    'Retained Earnings',
                    'Paid in Capital',
                    'Other'
                ]
            }
        elif statement_type == 'income_statement':
            return {
                'revenue': ['Revenue'],
                'operating_expenses': [
                    'Operating Expenses',
                    'Depreciation (-)',
                    'Amortization (-)',
                    'Assets gain(loss) impairments'
                ],
                'other_income_expense': [
                    'Interest Expense (-)',
                    'Interest Income (+)',
                    'Other income(expenses)'
                ],
                'tax_net_income': [
                    'Income Before Taxes',
                    'Tax expense',
                    'Net Income'
                ]
            }
        elif statement_type == 'cash_flow':
            return {
                'operating_activities': [
                    'Net Income',
                    'Changes in noncash items',
                    'Changes in Assets and Liabilities',
                    'Net Cash from(used) Operating Activities'
                ],
                'investing_activities': [
                    'CapEx',
                    'Proceeds from asset sales',
                    'Net cash from(used) for investing'
                ],
                'financing_activities': [
                    'Issuance of Debt (long+short term)',
                    'Retirement of Debt (long+short term)',
                    'Issuance of Stock',
                    'Dividends Paid',
                    'Net cash from(used) for financing'
                ],
                'cash_reconciliation': [
                    'Net change in Cash',
                    'Starting Cash',
                    'Ending Cash'
                ]
            }
        else:
            return {}
    
    def map_statement_comprehensive(self, extracted_data: Dict, statement_type: str) -> Tuple[Dict, List]:
        """
        Map an entire statement comprehensively using a single LLM call.
        Returns (mappings, unmapped_items)
        """
        if not self.check_ollama_available():
            print(f"[ERROR] Ollama not available for comprehensive mapping of {statement_type}")
            return {}, []
        
        template_structure = self.get_template_structure(statement_type)
        
        # Create appropriate prompt based on statement type
        if statement_type == 'balance_sheet':
            prompt = self.create_balance_sheet_prompt(extracted_data, template_structure)
        elif statement_type == 'income_statement':
            prompt = self.create_income_statement_prompt(extracted_data, template_structure)
        elif statement_type == 'cash_flow':
            prompt = self.create_cash_flow_prompt(extracted_data, template_structure)
        else:
            print(f"[ERROR] Unknown statement type: {statement_type}")
            return {}, []
        
        # Call Ollama
        result = self.call_ollama_comprehensive(prompt)
        
        if result:
            mappings = result.get("mappings", {})
            unmapped_items = result.get("unmapped_items", [])
            
            print(f"[DEBUG] Comprehensive mapping successful for {statement_type}:")
            print(f"[DEBUG]   Mappings: {len(mappings)} sections")
            print(f"[DEBUG]   Unmapped items: {len(unmapped_items)}")
            
            return mappings, unmapped_items
        else:
            print(f"[ERROR] Comprehensive mapping failed for {statement_type}")
            return {}, []
    
    def apply_mappings_to_template(self, mappings: Dict, template_path: str, statement_type: str) -> str:
        """
        Apply comprehensive mappings to the Excel template.
        """
        # Load template
        wb = openpyxl.load_workbook(template_path)
        
        if statement_type == 'balance_sheet':
            sheet = wb['BS']
            # Apply balance sheet mappings
            # This would need to be implemented based on your specific template structure
            pass
        elif statement_type == 'income_statement':
            sheet = wb['IS.CF']
            # Apply income statement mappings
            pass
        elif statement_type == 'cash_flow':
            sheet = wb['IS.CF']
            # Apply cash flow mappings
            pass
        
        # Save populated template
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_path = f"comprehensive_template_{statement_type}_{timestamp}.xlsx"
        wb.save(output_path)
        
        return output_path

def main():
    """Test the comprehensive mapper"""
    mapper = ComprehensiveMapper()
    
    if not mapper.check_ollama_available():
        print("❌ Ollama not available. Please install and run Ollama with Mistral model.")
        return
    
    print("✅ Ollama available for comprehensive mapping!")
    
    # Test with sample data
    sample_bs_data = {
        "2022": {
            "Cash and cash equivalents": 1500000,
            "Accounts receivable": 2500000,
            "Inventory": 1800000,
            "Property, plant and equipment": 5000000,
            "Accounts payable": 1200000,
            "Long-term debt": 3000000,
            "Common stock": 1000000,
            "Retained earnings": 5600000
        },
        "2023": {
            "Cash and cash equivalents": 1800000,
            "Accounts receivable": 2800000,
            "Inventory": 2000000,
            "Property, plant and equipment": 5200000,
            "Accounts payable": 1400000,
            "Long-term debt": 2800000,
            "Common stock": 1000000,
            "Retained earnings": 6800000
        }
    }
    
    print("\n🧪 Testing comprehensive balance sheet mapping...")
    mappings, unmapped = mapper.map_statement_comprehensive(sample_bs_data, 'balance_sheet')
    
    print(f"\n📊 Results:")
    print(f"Mappings: {json.dumps(mappings, indent=2)}")
    print(f"Unmapped items: {unmapped}")

if __name__ == "__main__":
    main() 