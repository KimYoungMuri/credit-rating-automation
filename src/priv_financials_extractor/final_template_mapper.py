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
Assets gain(loss) impairments
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
Changes in Assets and Liabilities
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
from sentence_transformers import SentenceTransformer
import shutil
from openpyxl import load_workbook
from collections import defaultdict
import openpyxl
from datetime import datetime
from llm_mapper import LLMMapper  # Import the LLM mapper
from sklearn.metrics.pairwise import cosine_similarity

# Place these at the module level, outside of any function
IS_SECTION_TEMPLATE = {
    'revenue': {
        'template': [
            'Revenue', 'Other Revenue'
        ]
    },
    'operating_expenses': {
        'template': [
            'Cost of revenue', 'Station operations costs', 'Payroll and related costs',
            'Depreciation and amortization', 'Impairment and other losses', 'Selling, general and administrative expenses', 'Other Operating Expenses'
        ]
    },
    'other_income_expense': {
        'template': [
            'Interest expense', 'Other (income) and expense, net', 'Other Income/Expense'
        ]
    },
    'tax_net_income': {
        'template': [
            'Income tax benefit', 'Net profit (loss)']
    }
}

CFS_SECTION_TEMPLATE = {
    'operating_activities': {
        'template': [
            'Net profit (loss)', 'Adjustments to reconcile net profit', 'Depreciation', 'Deferred income taxes', 'Impairment and other losses', 'Changes in operating assets and liabilities', 'Net cash provided by (used in) operating activities', 'Other Operating Activities'
        ]
    },
    'investing_activities': {
        'template': [
            'Purchases of property and equipment', 'Proceeds from sale of assets', 'Net cash used in investing activities', 'Other Investing Activities'
        ]
    },
    'financing_activities': {
        'template': [
            'Proceeds from issuance', 'Principal payments', 'Net cash provided by (used in) financing activities', 'Other Financing Activities'
        ]
    },
    'other': {
        'template': ['Other']
    },
    'cash_reconciliation': {
        'template': [
            'Net change in Cash',
            'Starting Cash',
            'Ending Cash'
        ]
    }
}

# Expand manual mapping for IS and CFS
manual_section_map = {}
manual_section_map.update({
    # IS Revenue
    'revenue net': 'revenue',
    'revenue': 'revenue',
    'other revenue': 'revenue',
    # IS Operating Expenses
    'cost of revenue': 'operating_expenses',
    'station operations costs': 'operating_expenses',
    'payroll and related costs': 'operating_expenses',
    'depreciation and amortization': 'operating_expenses',
    'impairment and other losses': 'operating_expenses',
    'selling general and administrative expenses': 'operating_expenses',
    'other operating expenses': 'operating_expenses',
    # IS Other Income/Expense
    'interest expense': 'other_income_expense',
    'other income and expense net': 'other_income_expense',
    'other income expense': 'other_income_expense',
    # IS Tax/Net Income
    'income tax benefit': 'tax_net_income',
    'net profit loss': 'tax_net_income',
    # CFS Operating Activities
    'net profit loss': 'operating_activities',
    'adjustments to reconcile net profit': 'operating_activities',
    'depreciation': 'operating_activities',
    'deferred income taxes': 'operating_activities',
    'impairment and other losses': 'operating_activities',
    'changes in operating assets and liabilities': 'operating_activities',
    'net cash provided by used in operating activities': 'operating_activities',
    'other operating activities': 'operating_activities',
    # CFS Investing Activities
    'purchases of property and equipment': 'investing_activities',
    'proceeds from sale of assets': 'investing_activities',
    'net cash used in investing activities': 'investing_activities',
    'other investing activities': 'investing_activities',
    # CFS Financing Activities
    'proceeds from issuance': 'financing_activities',
    'principal payments': 'financing_activities',
    'net cash provided by used in financing activities': 'financing_activities',
    'other financing activities': 'financing_activities',
})

# Add section boundary keywords for IS and CFS
section_boundary_keywords = {}
section_boundary_keywords['is_section_boundaries'] = [
    'revenue', 'cost of revenue', 'station operations costs', 'payroll and related costs',
    'depreciation and amortization', 'impairment and other losses', 'selling general and administrative expenses',
    'income from operations', 'interest expense', 'other income and expense net', 'income tax benefit', 'net profit loss'
]
section_boundary_keywords['cfs_section_boundaries'] = [
    'cash flows from operating activities', 'cash flows from investing activities', 'cash flows from financing activities',
    'net cash provided by used in operating activities', 'net cash used in investing activities', 'net cash provided by used in financing activities'
]

class TemplateMatcher:
    TOTAL_NET_PATTERNS = [
        r'^total\s+current\s+assets?$',
        r'^total\s+non[- ]?current\s+assets?$',
        r'^total\s+assets?$',
        r'^total\s+current\s+liabilities?$',
        r'^total\s+non[- ]?current\s+liabilities?$',
        r'^total\s+liabilities?$',
        r'^total\s+equity$',
        r'^net\s+cash\s+(?:from|used\s+in|provided\s+by)',
        r'^total\s+(?:current|non[- ]?current)\s+(?:assets?|liabilities?)$',
        r'^total$',
        r'^total\s+common\s+shareholders?\s+equity$',
        r'^total\s+shareholders?\s+equity$',
        r'^total\s+stockholders?\s+equity$',
        r'^total\s+debt$',
        r'^total\s+long[- ]term\s+debt$',
        r'^total\s+current\s+debt$',
        r'^total\s+investments?$',
        r'^total\s+property\s+and\s+equipment$',
        r'^total\s+ppe$',
        r'^total\s+intangible\s+assets?$',
        r'^total\s+goodwill$',
        r'^total\s+accounts?\s+receivable$',
        r'^total\s+accounts?\s+payable$',
        r'^total\s+accrued\s+liabilities?$',
        r'^total\s+inventory$',
        r'^total\s+cash$',
        r'^total\s+cash\s+and\s+cash\s+equivalents$',
        r'^total\s+current\s+portion\s+of\s+long[- ]term\s+debt$',
        r'^total\s+long[- ]term\s+incentive$',
        r'^total\s+deferred\s+compensation$',
        r'^total\s+finance\s+lease\s+liability$',
        r'^total\s+operating\s+lease\s+liability$',
        r'^total\s+deferred\s+income\s+taxes?$',
        r'^total\s+other\s+noncurrent\s+liabilities?$',
        r'total\s+other\s+current\s+liabilities?$',
        r'^total\s+other\s+current\s+assets?$',
        r'^total\s+other\s+noncurrent\s+assets?$',
        r'^total\s+noncontrolling\s+interests?$',
        r'^total\s+common\s+stock$',
        r'^total\s+retained\s+earnings?$',
        r'^total\s+paid[- ]in\s+capital$',
    ]

    def __init__(self, model_name="all-MiniLM-L6-v2"):
        """Initialize with Sentence Transformers model"""
        self.model = SentenceTransformer(model_name)
        self.setup_logging()
        self.used_items = set()  # Track used items globally
        self.extraction_logger = self.setup_extraction_logger()
        
        # Initialize LLM mapper for hybrid approach
        self.llm_mapper = LLMMapper()
        self.use_llm = self.llm_mapper.check_ollama_available()
        if self.use_llm:
            print("✅ LLM mapper available - using hybrid approach")
        else:
            print("⚠️  LLM mapper not available - using rule-based approach only")
            print("   Install Ollama: https://ollama.ai/")
            print("   Run: ollama pull mistral")
        
    def setup_logging(self):
        """Setup logging configuration"""
        logging.basicConfig(
            filename='template_mapping.log',
            level=logging.INFO,
            format='%(message)s',
            filemode='w'
        )

    def setup_extraction_logger(self):
        extraction_logger = logging.getLogger("extraction_logger")
        extraction_logger.setLevel(logging.DEBUG)
        # Remove any existing handlers
        if extraction_logger.hasHandlers():
            extraction_logger.handlers.clear()
        # Set up extraction.log in project root
        import os
        from pathlib import Path
        current_dir = Path(__file__).resolve().parent
        project_root = current_dir.parent.parent
        extraction_log_file = project_root / "extraction.log"
        handler = logging.FileHandler(extraction_log_file, mode='w', encoding='utf-8')
        formatter = logging.Formatter('%(message)s')
        handler.setFormatter(formatter)
        extraction_logger.addHandler(handler)
        return extraction_logger

    def get_similarity(self, text1: str, text2: str) -> float:
        """Get semantic similarity between two texts using sentence-transformers"""
        embeddings = self.model.encode([text1, text2])
        return np.dot(embeddings[0], embeddings[1]) / (np.linalg.norm(embeddings[0]) * np.linalg.norm(embeddings[1]))

    def find_best_match(self, source_item: str, target_items: List[str], section_context: str, threshold: float = 0.5) -> Tuple[Optional[str], float]:
        """Find best matching template item for a given source item within a section context"""
        best_match = None
        best_score = -1
        
        # Normalize source item
        source_item = source_item.lower().strip()
        
        # Direct match patterns for common items
        direct_matches = {
            'Cash and equivalents': [
                r'cash(?:\s+and\s+(?:cash\s+)?equivalents?)?',
                r'cash\s+equivalents?'
            ],
            'Accounts Receivable': [
                r'accounts?\s+receivable',
                r'(?:trade|net)\s+receivables?'
            ],
            'Inventory': [
                r'inventor(?:y|ies)(?:\s*[-—]\s*net)?',
                r'net\s+inventor(?:y|ies)'
            ],
            'Net PPE': [
                r'property(?:\s+and\s+equipment)?(?:\s*[-—]\s*net)?',
                r'(?:net\s+)?(?:ppe|property,?\s+plant\s+and\s+equipment)',
                r'right\s+of\s+use\s+assets?',
                r'finance\s+lease\s+assets?'
            ],
            'Goodwill': [
                r'goodwill(?:\s*[-—]\s*net)?',
                r'(?:net\s+)?goodwill'
            ]
        }
        
        # Check for direct matches first
        for target in target_items:
            if target in direct_matches:
                for pattern in direct_matches[target]:
                    if re.search(pattern, source_item):
                        return target, 1.0
        
        # Section-specific term mappings
        section_terms = {
            'assets': [
                'cash', 'receivable', 'inventory', 'prepaid', 'investment', 
                'property', 'equipment', 'ppe', 'goodwill', 'intangible',
                'lease', 'right of use', 'margin deposit', 'derivative'
            ],
            'liabilities': [
                'payable', 'debt', 'borrowing', 'loan', 'accrued', 'deferred',
                'liability', 'lease', 'obligation', 'tax'
            ],
            'equity': [
                'stock', 'capital', 'earning', 'dividend', 'share', 'equity',
                'retained', 'paid-in', 'comprehensive'
            ]
        }
        
        # Only proceed if item matches section context
        if section_context in section_terms:
            matches_section = any(term in source_item for term in section_terms[section_context])
            if not matches_section:
                return None, 0
        
        for target in target_items:
            target = target.lower().strip()
            base_score = self.get_similarity(source_item, target)
            
            # Add bonus for matching section context
            if section_context in section_terms:
                if any(term in target for term in section_terms[section_context]):
                    base_score += 0.2
            
            # Add bonus for exact word matches
            source_words = set(re.findall(r'\b\w+\b', source_item))
            target_words = set(re.findall(r'\b\w+\b', target))
            common_words = source_words.intersection(target_words)
            if common_words:
                base_score += 0.1 * len(common_words)
            
            # Add bonus for matching parenthetical terms
            source_parens = re.findall(r'\((.*?)\)', source_item)
            target_parens = re.findall(r'\((.*?)\)', target)
            if source_parens and target_parens:
                if any(s.strip() in [t.strip() for t in target_parens] for s in source_parens):
                    base_score += 0.1
            
            # Add bonus for matching numerical indicators
            if ('current' in source_item and 'current' in target) or \
               ('long term' in source_item and 'long term' in target) or \
               ('short term' in source_item and 'short term' in target):
                base_score += 0.15
            
            # Penalize matching to "Other" unless it's really a good match
            if 'other' in target and 'other' not in source_item:
                base_score *= 0.7
            
            if base_score > best_score and base_score >= threshold:
                best_score = base_score
                best_match = target
        
        return best_match, best_score

    def _map_section(self, sheet, data: Dict, col: str, start_row: int, end_row: int, template_items: List[str], section_context: str):
        """Map a section of data to the template"""
        # Track matched items to handle "Other" categories
        matched_items = set()
        other_values = 0.0
        
        # First pass: direct matches with improved semantic matching
        for item, value in data.items():
            # Skip if already used or if it's a total line
            if item in self.used_items or any(total_word in item.lower() for total_word in ['total', 'sum']):
                continue
            
            # Normalize value
            try:
                if isinstance(value, str):
                    value = float(value.replace(',', ''))
                else:
                    value = float(value)
            except (ValueError, TypeError):
                logging.warning(f"Could not convert value {value} to float for {item}")
                continue
            
            best_match, score = self.find_best_match(item, template_items, section_context)
            if best_match:
                # Find row for this item
                for row in range(start_row, end_row + 1):
                    template_item = sheet[f'A{row}'].value
                    if template_item and template_item.strip() == best_match:
                        sheet[f"{col}{row}"] = value
                        matched_items.add(item)
                        self.used_items.add(item)  # Mark as used globally
                        logging.info(f"Matched {item} to {best_match} with score {score}")
                        break
        
        # Second pass: try to match remaining items with more context
        for item, value in data.items():
            if item in self.used_items or item in matched_items:
                continue
            
            # Skip totals
            if any(total_word in item.lower() for total_word in ['total', 'sum']):
                continue
            
            try:
                if isinstance(value, str):
                    value = float(value.replace(',', ''))
                else:
                    value = float(value)
                
                # Try to match based on context
                context_score = 0
                best_match = None
                
                for template_item in template_items:
                    # Check for semantic similarity with section context
                    score = self.get_similarity(item.lower(), template_item.lower())
                    if section_context:
                        score = self.find_best_match(item, [template_item], section_context)[1]
                    
                    if score > context_score:
                        context_score = score
                        best_match = template_item
                
                if best_match and context_score >= 0.5:
                    # Find row for this item
                    for row in range(start_row, end_row + 1):
                        template_item = sheet[f'A{row}'].value
                        if template_item and template_item.strip() == best_match:
                            sheet[f"{col}{row}"] = value
                            matched_items.add(item)
                            self.used_items.add(item)  # Mark as used globally
                            logging.info(f"Context matched {item} to {best_match} with score {context_score}")
                            break
                else:
                    # Only add to Other if it matches section context
                    if self.find_best_match(item, ['Other'], section_context)[1] >= 0.3:
                        other_values += value
                        self.used_items.add(item)  # Mark as used globally
                        logging.info(f"Adding {item} ({value}) to Other category")
            except (ValueError, TypeError):
                continue
        
        # Find and populate "Other" row if we have unmatched items
        if other_values != 0:
            found_other = False
            for row in range(start_row, end_row + 1):
                template_item = sheet[f'A{row}'].value
                print(f"[DEBUG] Checking row {row} for 'Other': '{template_item}'")
                if template_item and template_item.strip() == 'Other':
                    print(f"[DEBUG] Writing {other_values} to {section_context}::Other at cell {col}{row}")
                    sheet[f"{col}{row}"] = other_values
                    found_other = True
                    break
            if not found_other:
                print(f"[WARNING] Could not find 'Other' row in {section_context} section (rows {start_row}-{end_row}) to write value {other_values}")
        
        # Calculate totals for sections that have them
        if end_row in [13, 20, 29, 35, 43]:  # Total rows
            total = 0
            for row in range(start_row, end_row):
                val = sheet[f"{col}{row}"].value
                try:
                    total += float(val) if val not in [None, ''] else 0
                except (ValueError, TypeError):
                    continue
            sheet[f"{col}{end_row}"] = total

    def analyze_subsections(self, item: str) -> str:
        """Determine which subsection an item belongs to based on its content"""
        item = item.lower().strip()
        
        # Define subsection patterns
        subsections = {
            'Current Assets': [
                r'cash', r'equivalent', r'margin\s+deposit', r'derivative\s+asset',
                r'(?:current\s+)?(?:account|trade)\s+receivable', r'current\s+portion',
                r'inventor(?:y|ies)', r'prepaid', r'current\s+asset'
            ],
            'Non-Current Assets': [
                r'(?:property|equipment|ppe)(?:\s*[-—]\s*net)?', r'right\s+of\s+use',
                r'finance\s+lease', r'goodwill', r'intangible', r'non-?current\s+asset',
                r'deferred\s+(?:tax|compensation)', r'long-term', r'investment'
            ],
            'Current Liabilities': [
                r'current\s+liabilit(?:y|ies)', r'accounts?\s+payable',
                r'accrued', r'short[- ]term', r'current\s+portion'
            ],
            'Non-Current Liabilities': [
                r'long[- ]term\s+(?:debt|lease|liability)',
                r'deferred\s+(?:tax|revenue|income)',
                r'non-?current\s+liabilit(?:y|ies)'
            ],
            'Equity': [
                r'(?:common|preferred)\s+stock', r'retained\s+earnings?',
                r'paid[- ]in\s+capital', r'shareholder', r'equity'
            ],
            'Revenue': [
                r'revenue', r'sales', r'income\s+from\s+operations'
            ],
            'Operating Expenses': [
                r'cost\s+of\s+(?:goods\s+sold|revenue|sales)',
                r'operating\s+expense', r'selling', r'administrative',
                r'depreciation', r'amortization'
            ],
            'Other Income/Expense': [
                r'interest\s+(?:income|expense)',
                r'other\s+(?:income|expense)',
                r'gain|loss'
            ],
            'Operating Activities': [
                r'operating\s+activit(?:y|ies)',
                r'cash\s+from\s+operations?',
                r'working\s+capital'
            ],
            'Investing Activities': [
                r'investing\s+activit(?:y|ies)',
                r'capital\s+expenditure',
                r'acquisition',
                r'purchase\s+of'
            ],
            'Financing Activities': [
                r'financing\s+activit(?:y|ies)',
                r'dividend',
                r'stock\s+(?:issue|repurchase)',
                r'debt\s+(?:issue|repayment)'
            ]
        }
        
        # Check each subsection's patterns
        for subsection, patterns in subsections.items():
            for pattern in patterns:
                if re.search(pattern, item):
                    return subsection
        
        return "Uncategorized"

    def print_categorization(self, extracted_data: Dict):
        """Print categorization of all extracted values"""
        print("\nCategorization of Extracted Values:")
        print("=" * 80)
        
        for statement_type, years in extracted_data.items():
            print(f"\n{statement_type.upper()}")
            print("-" * 80)
            
            for year, items in years.items():
                print(f"\n{year}:")
                
                # Group items by subsection
                categorized = defaultdict(list)
                for item, value in items.items():
                    if not any(total_word in item.lower() for total_word in ['total', 'sum']):
                        subsection = self.analyze_subsections(item)
                        categorized[subsection].append((item, value))
                
                # Print categorized items
                for subsection in sorted(categorized.keys()):
                    print(f"\n  {subsection}:")
                    for item, value in sorted(categorized[subsection]):
                        print(f"    {item}: {value}")

    def assign_sections_by_context(self, extracted_lines: list) -> list:
        """
        Assigns BS sections to each line item based on content and explicit headers.
        """
        assigned = []
        current_section = None
        
        # Keywords to identify major section headers
        section_headers = {
            'current_assets': ['current assets'],
            'noncurrent_assets': ['noncurrent assets', 'non-current assets'],
            'current_liabilities': ['current liabilities'],
            'noncurrent_liabilities': ['noncurrent liabilities', 'non-current liabilities'],
            'equity': ['equity', "stockholders' equity", "shareholders' equity"]
        }

        # Keywords for specific items that can also define a section
        manual_map = {
            'cash and cash equivalents': 'current_assets',
            'accounts receivable': 'current_assets',
            'inventories': 'current_assets',
            'property, plant and equipment': 'noncurrent_assets',
            'accounts payable': 'current_liabilities',
            'retained earnings': 'equity',
            'additional paid in capital': 'equity',
            'common stock': 'equity'
        }
        
        for line in extracted_lines:
            desc = line['description']
            desc_lower = desc.lower()
            
            # 1. Check for major section headers first
            found_header = False
            for section, keywords in section_headers.items():
                if any(keyword in desc_lower for keyword in keywords):
                    current_section = section
                    found_header = True
                    break
            
            # 2. If no header, check for specific item keywords that imply a section change
            if not found_header:
                for keyword, section in manual_map.items():
                    if keyword in desc_lower:
                        current_section = section
                                break
            
            assigned.append({
                'description': desc,
                'value': line.get('numbers', []),
                'section': current_section
            })
        
        return assigned

    def assign_sections_by_context_is(self, extracted_lines: list) -> list:
        """
        Assigns IS sections to each line item based on its content and context.
        Uses explicit header detection to improve accuracy.
        """
        import re
        
        assigned = []
        current_section = None # Start with no section
        
        # Section headers to look for (case-insensitive)
        section_keywords = {
            'revenue': ['net sales', 'revenue'],
            'operating_expenses': ['operating costs and expenses', 'operating expenses'],
            'operating_income': ['operating income'],
            'other_income_expense': ['other income (expense)', 'other income'],
            'tax_net_income': ['net income', 'income before', 'comprehensive income']
        }

        for line in extracted_lines:
            desc_lower = line['description'].lower()
            
            # --- 1. Detect section based on headers ---
            found_section = False
            for section, keywords in section_keywords.items():
                if any(re.search(r'\\b' + keyword + r'\\b', desc_lower) for keyword in keywords):
                    current_section = section
                    found_section = True
                    break
            
            # --- 2. If no header, check if it's a known non-op item ---
            if not found_section:
                # This part is simplified; can be expanded with more keywords
                if 'interest' in desc_lower:
                    current_section = 'other_income_expense'

            # --- 3. Assign section ---
                        assigned_section = current_section
            
            assigned.append({
                'description': line['description'],
                'value': line['numbers'],  # FIX: Was 'line'
                'section': assigned_section
            })
        
        return assigned

    def assign_sections_by_context_cfs(self, extracted_lines: list) -> list:
        """
        Assigns CFS sections to each line item using a more robust, keyword-based state machine.
        This correctly handles the final cash reconciliation section.
        """
        assigned = []
        current_section = None  # Start with no section

        # Define keywords that signal the start of a new section
        section_keywords = {
            'operating_activities': [
                'cash flow from operating activities',
                'cash flows from operating activities',
                'operating activities'
            ],
            'investing_activities': [
                'cash flow from investing activities',
                'cash flows from investing activities',
                'investing activities'
            ],
            'financing_activities': [
                'cash flow from financing activities',
                'cash flows from financing activities',
                'financing activities'
            ],
            'cash_reconciliation': [
                'net change in cash',
                'net increase in cash',
                'net decrease in cash',
                'effect of exchange rate',
                'cash at beginning',
                'cash and cash equivalents at beginning',
                'supplemental disclosure' # Often marks the end of main sections
            ]
        }
        
        for line in extracted_lines:
            desc_lower = line['description'].lower().strip()
            
            # Check if the line description indicates a new section
            found_new_section = False
            for section, keywords in section_keywords.items():
                if any(keyword in desc_lower for keyword in keywords):
                    current_section = section
                    self.extraction_logger.info(f"[SECTION SWITCH] '{line['description']}' -> {current_section}")
                    found_new_section = True
                    break
            
            # If no section is ever found, we can try a fallback
            if current_section is None:
                if 'net income' in desc_lower or 'net loss' in desc_lower:
                    current_section = 'operating_activities'
            
            assigned.append({
                'description': line['description'],
                'value': line.get('numbers', []),
                'section': current_section
            })
            
        # Add the new section to the template definition for mapping
        if 'cash_reconciliation' not in CFS_SECTION_TEMPLATE:
             CFS_SECTION_TEMPLATE['cash_reconciliation'] = {
                 'template': [
                     'Net change in Cash',
                     'Starting Cash',
                     'Ending Cash'
                 ]
             }
        
        return assigned

    def get_is_row_map(self, worksheet, section: str) -> dict:
        """
        Returns a mapping of template item descriptions to their row number 
        for a given IS section by reading directly from the worksheet.
        This is more robust than using a pre-processed list.
        """
        row_map = {}
        # Correct, non-overlapping sections based on template visual layout
        section_rows = {
            'revenue': [(6, 6)],
            'operating_expenses': [(10, 12)], # Dep, Amort, Impair
            'other_income_expense': [(13, 15), (19, 19)], # Interest, Other Income, Other
            'tax_net_income': [(18, 18)] # Tax Expense only
        }
        
        # Rows that are headers or calculated totals, and should not be mapped to
        skip_rows = {
            7: "Operating Expenses", 
            8: "Operating Income", 
            16: "Income Before Taxes", 
            17: "Income Before Taxes",
            20: "Net Income"
        }

        ranges = section_rows.get(section, [])
        if not ranges: return {}
        
        for start, end in ranges:
            for row_num in range(start, end + 1):
                if row_num in skip_rows:
                    continue
                item = worksheet.cell(row=row_num, column=1).value
                if item:
                    row_map[item] = row_num
        return row_map

    def get_cfs_row_map(self, template_items: list, section: str) -> dict:
        """
        Returns a mapping of template item descriptions to their row number for a given CFS section.
        """
        cfs_row_mapping = {
            'Net profit (loss)': 23,
            'Adjustments to reconcile net profit': 24,
            'Depreciation': 25,
            'Deferred income taxes': 26,
            'Impairment and other losses': 27,
            'Changes in operating assets and liabilities': 28,
            'Net cash provided by (used in) operating activities': 29,
            'Other Operating Activities': 30,
            'Purchases of property and equipment': 32,
            'Proceeds from sale of assets': 33,
            'Net cash used in investing activities': 34,
            'Other Investing Activities': 35,
            'Proceeds from issuance': 37,
            'Principal payments': 38,
            'Net cash provided by (used in) financing activities': 39,
            'Other Financing Activities': 40,
            'Other': 42
        }
        
        return {item: cfs_row_mapping.get(item) for item in template_items if item in cfs_row_mapping}

    def map_to_template(self, extracted_data: Dict, template_path: str) -> str:
        # Load the template workbook
        shutil.copy(template_path, "temp_template.xlsx")
        wb = load_workbook("temp_template.xlsx")

        # Get sheets for BS, IS, CFS
        bs_sheet = wb['BS']
        is_cf_sheet = wb['IS.CF']

        # Determine year columns from the template (assuming they are in B, C, D, E starting from row 6)
        year_cols = {}
        print("DEBUG: Starting year column detection...")
        # Heuristic: Find years in row 6 of the Balance Sheet
        for col_idx in range(2, 6): # Check columns B, C, D, E
            cell_val = bs_sheet.cell(row=6, column=col_idx).value
            col_letter = openpyxl.utils.get_column_letter(col_idx)
            print(f"DEBUG: Checking column {col_letter} (row 6): {cell_val} (type: {type(cell_val)})")
            
            if isinstance(cell_val, int) and 1990 <= cell_val <= 2050:
                year_cols[str(cell_val)] = col_letter
                print(f"DEBUG: Found direct year {cell_val} in column {col_letter}")
            elif isinstance(cell_val, str) and cell_val.startswith('='):
                # Handle Excel formulas like "=B6+1" - extract the year
                print(f"DEBUG: Found formula {cell_val}, attempting to parse...")
                try:
                    # Simple parsing for formulas like "=B6+1" where B6 contains 2020
                    if '+1' in cell_val:
                        base_cell = cell_val.split('+')[0][1:]  # Extract "B6" from "=B6+1"
                        base_col = base_cell[0]  # "B"
                        base_row = int(base_cell[1:])  # 6
                        base_year = bs_sheet[f"{base_col}{base_row}"].value
                        print(f"DEBUG: Base cell {base_col}{base_row} contains: {base_year}")
                        if isinstance(base_year, int):
                            # Calculate the year for this column
                            col_offset = col_idx - openpyxl.utils.column_index_from_string(base_col)
                            calculated_year = base_year + col_offset
                            print(f"DEBUG: Calculated year: {base_year} + {col_offset} = {calculated_year}")
                            if 1990 <= calculated_year <= 2050:
                                year_cols[str(calculated_year)] = col_letter
                                print(f"DEBUG: Found calculated year {calculated_year} in column {col_letter}")
                except Exception as e:
                    print(f"DEBUG: Error parsing formula {cell_val}: {e}")
                    pass

        if not year_cols:
            print("ERROR: Could not determine year columns from template. Aborting.")
            return ""

        print(f"DEBUG: Found year columns in template: {year_cols}")

        # --- Balance Sheet Mapping ---
        if 'balance_sheet' in extracted_data:
            print("\n--- Processing Balance Sheet ---")
            
            # --- Get Balance Sheet row maps ---
            row_maps = {
                'current_assets': self.get_bs_row_map(bs_sheet, 7, 12),
                'noncurrent_assets': self.get_bs_row_map(bs_sheet, 15, 18),
                'current_liabilities': self.get_bs_row_map(bs_sheet, 23, 28),
                'noncurrent_liabilities': self.get_bs_row_map(bs_sheet, 31, 34),
                'equity': self.get_bs_row_map(bs_sheet, 38, 41)
            }
            
            for year, year_data in extracted_data['balance_sheet'].items():
                if str(year) not in year_cols:
                    print(f"Warning: Year {year} from PDF not found in template columns. Skipping.")
                    continue
                
                col = year_cols[str(year)]
                print(f"\n--- Mapping year {year} to column {col} ---")
                
                # Assign sections to all items for this year
                # We need to convert the data to the format the section assigner expects
                bs_lines_for_year = [{'description': d, 'numbers': [v]} for d, v in year_data.items()]
                assigned_bs_lines = self.assign_sections_by_context(bs_lines_for_year)
                
                self.print_section_assignments(assigned_bs_lines, year, 'balance_sheet')

                # Group items by their assigned section
                section_data = defaultdict(list)
                for item in assigned_bs_lines:
                    # item is {'description': ..., 'value': [...], 'section': ...}
                    if item.get('section') and item.get('value'):
                        section_data[item['section']].append({
                            'description': item['description'],
                            'value': item['value'][0] # We process one year at a time
                        })

                # Map each section
                for section, items in section_data.items():
                    row_map = row_maps.get(section)
                    if not row_map:
                        print(f"Warning: No row map found for section '{section}'. Skipping.")
                        continue
                        
                    template_items = list(row_map.keys())
                    print(f"\n[DEBUG] Mapping section '{section}'. Template items: {template_items}")

                    accumulated_values = defaultdict(float)
                    other_sum = 0.0

                    for entry in items:
                        desc = entry['description']
                        val_str = entry['value']
                        if val_str is None: continue

                        try:
                            val = float(re.sub(r'[^\d\.-]', '', str(val_str)))
                        except (ValueError, TypeError):
                            print(f"Warning: Could not convert value to float for '{desc}': {val_str}")
                            continue

                        # Skip total/net rows from being mapped to specific items
                        if self.is_total_or_net_row(desc):
                            print(f"  [SKIP TOTAL] '{desc}' is a total/subtotal row.")
                    continue
                
                        # Use the hybrid mapping function
                        target_item, score, method = self.hybrid_map_item(desc, template_items, section, 'balance_sheet')

                        if target_item and score >= 0.4:
                            accumulated_values[target_item] += val
                            print(f"  [ACCUMULATE HYBRID] '{desc}' ({val}) -> {section}::{target_item} [conf: {score:.2f}, method: {method}]")
                            else:
                            # If it's not a confident match, add it to 'Other' for this section
                                    other_sum += val
                            print(f"  [ACCUMULATE OTHER] '{desc}' ({val}) -> {section}::Other [best match: {target_item}, conf: {score:.2f}]")

                    # Write accumulated values to template
                    for t_item, total_val in accumulated_values.items():
                        if t_item in row_map and total_val != 0:
                            row_idx = row_map[t_item]
                            bs_sheet[f"{col}{row_idx}"] = total_val
                    
                    # Write the sum of unmapped items to 'Other' for that section
                    if 'Other' in row_map and other_sum != 0:
                        row_idx = row_map['Other']
                        # Add to existing value if any
                        existing_val = bs_sheet[f"{col}{row_idx}"].value or 0
                        if isinstance(existing_val, str): existing_val = 0
                        bs_sheet[f"{col}{row_idx}"] = existing_val + other_sum
                        print(f"    [OTHER SUM] Writing {other_sum} to '{section}::Other' in cell {col}{row_idx}")

        # --- Income Statement Mapping ---
        if 'income_statement' in extracted_data:
            print("\n--- Processing Income Statement ---")
            
            # --- Get IS row map ---
            is_row_map = self.get_is_row_map(is_cf_sheet)
            
            for year, year_data in extracted_data['income_statement'].items():
                if str(year) not in year_cols:
                    print(f"Warning: Year {year} from PDF not found in template columns. Skipping.")
                    continue
                
                col = year_cols[str(year)]
                print(f"\n--- Mapping year {year} to column {col} ---")

                is_lines_for_year = [{'description': d, 'numbers': [v]} for d, v in year_data.items()]
                assigned_is_lines = self.assign_sections_by_context_is(is_lines_for_year)
                
                self.print_section_assignments(assigned_is_lines, year, 'income_statement')
                
                section_data = defaultdict(list)
                for item in assigned_is_lines:
                    if item.get('section') and item.get('value'):
                        section_data[item['section']].append({
                            'description': item['description'],
                            'value': item['value'][0]
                        })

                for section, items in section_data.items():
                    template_items = IS_SECTION_TEMPLATE.get(section, {}).get('template', [])
                    if not template_items: continue

                    print(f"\n[DEBUG] Mapping IS section '{section}'. Template items: {template_items}")
                    
                    accumulated_values = defaultdict(float)
                    other_sum = 0.0
                    
                    for entry in items:
                        desc = entry['description']
                        val_str = entry['value']
                        if val_str is None: continue
                        
                        try:
                            val = float(re.sub(r'[^\d\.-]', '', str(val_str)))
                        except (ValueError, TypeError):
                            continue
                        
                        if self.is_total_or_net_row(desc):
                            print(f"  [SKIP TOTAL] '{desc}' is a total/subtotal row.")
                            continue
                        
                        target_item, score, method = self.hybrid_map_item(desc, template_items, section, 'income_statement')
                        
                        if target_item and score >= 0.4:
                            accumulated_values[target_item] += val
                            print(f"  [ACCUMULATE HYBRID] '{desc}' ({val}) -> {section}::{target_item} [conf: {score:.2f}, method: {method}]")
                            else:
                                    other_sum += val
                            print(f"  [ACCUMULATE OTHER] '{desc}' ({val}) -> {section}::Other")
                    
                    # Write values
                    for t_item, total_val in accumulated_values.items():
                        if t_item in is_row_map:
                            row_idx = is_row_map[t_item]
                            is_cf_sheet[f"{col}{row_idx}"] = total_val
                    
                    # Write 'Other' sum
                    other_category_name = self.get_other_category_for_is_section(section)
                    if other_category_name and other_category_name in is_row_map and other_sum != 0:
                        row_idx = is_row_map[other_category_name]
                        existing_val = is_cf_sheet[f"{col}{row_idx}"].value or 0
                        if isinstance(existing_val, str): existing_val = 0
                        is_cf_sheet[f"{col}{row_idx}"] = existing_val + other_sum
                        print(f"    [OTHER SUM] Writing {other_sum} to '{other_category_name}' in cell {col}{row_idx}")

        # --- Cash Flow Statement Mapping ---
        if 'cash_flow' in extracted_data:
            print("\n--- Processing Cash Flow Statement ---")
            
            cfs_row_map = self.get_cfs_row_map(is_cf_sheet)

            for year, year_data in extracted_data['cash_flow'].items():
                if str(year) not in year_cols:
                    continue
                
                col = year_cols[str(year)]
                print(f"\n--- Mapping year {year} to column {col} ---")
                
                cfs_lines_for_year = [{'description': d, 'numbers': [v]} for d, v in year_data.items()]
                assigned_cfs_lines = self.assign_sections_by_context_cfs(cfs_lines_for_year)
                
                self.print_section_assignments(assigned_cfs_lines, year, 'cash_flow')
                
                section_data = defaultdict(list)
                for item in assigned_cfs_lines:
                    if item.get('section') and item.get('value'):
                        section_data[item['section']].append({
                            'description': item['description'],
                            'value': item['value'][0]
                        })
                
                for section, items in section_data.items():
                    template_items = CFS_SECTION_TEMPLATE.get(section, {}).get('template', [])
                    if not template_items: continue
                    
                    print(f"\n[DEBUG] Mapping CFS section '{section}'. Template items: {template_items}")
                    
                    accumulated_values = defaultdict(float)
                    other_sum = 0.0

                    for entry in items:
                        desc = entry['description']
                        val_str = entry['value']
                        if val_str is None: continue
                        
                        try:
                            val = float(re.sub(r'[^\d\.-]', '', str(val_str)))
                        except (ValueError, TypeError): continue
                        
                        if self.is_total_or_net_row(desc):
                            print(f"  [SKIP TOTAL] '{desc}' is a total/subtotal row.")
                            continue
                        
                        target_item, score, method = self.hybrid_map_item(desc, template_items, section, 'cash_flow')
                        
                        if target_item and score >= 0.4:
                            accumulated_values[target_item] += val
                            print(f"  [ACCUMULATE HYBRID] '{desc}' ({val}) -> {section}::{target_item} [conf: {score:.2f}, method: {method}]")
                            else:
                                    other_sum += val
                            print(f"  [ACCUMULATE OTHER] '{desc}' ({val}) -> {section}::Other")
                    
                    # Write values
                    for t_item, total_val in accumulated_values.items():
                        if t_item in cfs_row_map:
                            row_idx = cfs_row_map[t_item]
                            is_cf_sheet[f"{col}{row_idx}"] = total_val
                            
                    # Write 'Other' sum
                    other_category_name = self.get_other_category_for_cfs_section(section)
                    if other_category_name and other_category_name in cfs_row_map and other_sum != 0:
                        row_idx = cfs_row_map[other_category_name]
                        existing_val = is_cf_sheet[f"{col}{row_idx}"].value or 0
                        if isinstance(existing_val, str): existing_val = 0
                        is_cf_sheet[f"{col}{row_idx}"] = existing_val + other_sum
                        print(f"    [OTHER SUM] Writing {other_sum} to '{other_category_name}' in cell {col}{row_idx}")

        # Save the populated template
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        current_dir = Path(__file__).resolve().parent
        project_root = current_dir.parent.parent
        output_dir = project_root / "output_excel"
        output_dir.mkdir(exist_ok=True)
        
        populated_template_path = output_dir / f"populated_template_{timestamp}.xlsx"
        wb.save(populated_template_path)
        Path("temp_template.xlsx").unlink()
        print(f"\nTemplate populated and saved to: {populated_template_path}")
        return str(populated_template_path)

    def get_bs_row_map(self, sheet, start_row, end_row):
        """Helper to get row map for a BS section."""
        row_map = {}
        for row in range(start_row, end_row + 1):
            cell_val = sheet.cell(row=row, column=1).value
            if cell_val:
                row_map[cell_val] = row
        return row_map
    
    def get_is_row_map(self, sheet):
        """Helper to get row map for the entire IS."""
        row_map = {}
        # Iterate over all defined IS sections and their templates
        for section_info in IS_SECTION_TEMPLATE.values():
            for item in section_info['template']:
                # Find this item in column A
                for row in range(1, sheet.max_row + 1):
                    if sheet.cell(row=row, column=1).value == item:
                        row_map[item] = row
                        break
        return row_map
        
    def get_cfs_row_map(self, sheet):
        """Helper to get row map for the entire CFS."""
        row_map = {}
        for section_info in CFS_SECTION_TEMPLATE.values():
            for item in section_info['template']:
                for row in range(1, sheet.max_row + 1):
                    if sheet.cell(row=row, column=1).value == item:
                        row_map[item] = row
                        break
        return row_map
        
    def get_other_category_for_is_section(self, section: str) -> Optional[str]:
        if section == 'operating_expenses': return 'Other Operating Expenses'
        if section == 'other_income_expense': return 'Other Income/Expense'
        return None

    def get_other_category_for_cfs_section(self, section: str) -> Optional[str]:
        if section == 'operating_activities': return 'Other Operating Activities'
        if section == 'investing_activities': return 'Other Investing Activities'
        if section == 'financing_activities': return 'Other Financing Activities'
        return None

    def print_section_assignments(self, assigned_lines: list, year: str, statement_type: str):
        print(f"\nSection assignments for {statement_type} {year}:")
        for entry in assigned_lines:
            print(f"  {entry['description']}  -->  {entry['section']}  (value: {entry['value']})")

    def is_total_or_net_row(self, description: str) -> bool:
        """Check if description is a total or net row"""
        import re
        desc_lower = description.lower()
        detail_lines = [
            'other current assets',
            'other noncurrent assets', 
            'other current liabilities',
            'other noncurrent liabilities',
            'other'
        ]
        if any(detail in desc_lower for detail in detail_lines):
            return False
        return any(re.search(pattern, desc_lower) for pattern in self.TOTAL_NET_PATTERNS)

    def apply_rule_based_mapping(self, description: str) -> tuple[str, float]:
        """Apply rule-based mapping for balance sheet items"""
        import re
        desc_lower = description.lower()
        
        # Balance sheet rule-based mapping
        bs_rules = {
            r'cash\s+(?:and\s+)?(?:cash\s+)?equivalents?': 'Cash and equivalents',
            r'accounts?\s+receivable(?:[—-]net)?': 'Accounts Receivable',
            r'prepaid\s+expenses?': 'Prepaid Expenses',
            r'inventor(?:y|ies)(?:[—-]net)?': 'Inventory',
            r'property\s+(?:and\s+)?equipment(?:[—-]net)?': 'Net PPE',
            r'net\s+ppe': 'Net PPE',
            r'goodwill(?:[—-]net)?': 'Goodwill',
            r'(?:other\s+)?intangible\s+assets?(?:[—-]net)?': 'Intangibles',
            r'accounts?\s+payable': 'Accounts Payable',
            r'accrued\s+(?:liabilities?|interest)': 'Accrued Interest',
            r'long[- ]term\s+debt(?!.*current)': 'Long Term Debt',
            r'deferred\s+income\s+taxes?': 'Deferred income taxes',
            r'common\s+stock': 'Common Stock',
            r'retained\s+earnings?': 'Retained Earnings',
            r'paid[- ]in\s+capital': 'Paid in Capital'
        }
        
        for pattern, template_item in bs_rules.items():
            if re.search(pattern, desc_lower):
                return template_item, 0.9
        
        return None, 0.0

    def apply_rule_based_mapping_is(self, description: str) -> tuple[str, float]:
        """Apply rule-based mapping for income statement items"""
        import re
        desc_lower = description.lower()
        
        # Income statement rule-based mapping
        is_rules = {
            r'revenue(?:s)?': 'Revenue',
            r'cost\s+of\s+(?:goods\s+)?sales?': 'Cost of Sales',
            r'gross\s+profit': 'Gross Profit',
            r'operating\s+expenses?': 'Operating Expenses',
            r'operating\s+income': 'Operating Income',
            r'interest\s+expense': 'Interest Expense',
            r'income\s+tax(?:es)?': 'Income Tax Expense',
            r'net\s+income': 'Net Income'
        }
        
        for pattern, template_item in is_rules.items():
            if re.search(pattern, desc_lower):
                return template_item, 0.9
        
        return None, 0.0

    def apply_rule_based_mapping_cfs(self, description: str) -> tuple[str, float]:
        """Apply rule-based mapping for cash flow statement items"""
        import re
        desc_lower = description.lower()
        
        # Cash flow statement rule-based mapping
        cfs_rules = {
            r'net\s+income': 'Net Income',
            r'depreciation': 'Depreciation',
            r'amortization': 'Amortization',
            r'capital\s+expenditures?': 'Capital Expenditures',
            r'dividends?': 'Dividends',
            r'proceeds?\s+from\s+debt': 'Proceeds from Debt',
            r'repayment\s+of\s+debt': 'Repayment of Debt'
        }
        
        for pattern, template_item in cfs_rules.items():
            if re.search(pattern, desc_lower):
                return template_item, 0.9
        
        return None, 0.0

    def get_semantic_match(self, description: str, template_items: list, section_context: str) -> tuple[str, float]:
        """Get semantic match using sentence transformers"""
        if not template_items:
            return None, 0.0
        
        try:
            # Get embeddings
            desc_embedding = self.model.encode([description])
            template_embeddings = self.model.encode(template_items)
            
            # Calculate similarities
            similarities = cosine_similarity(desc_embedding, template_embeddings)[0]
            
            # Find best match
            best_idx = np.argmax(similarities)
            best_score = similarities[best_idx]
            
            if best_score > 0.3:  # Lower threshold for semantic matching
                return template_items[best_idx], best_score
            
            return None, 0.0
            
        except Exception as e:
            print(f"Error in semantic matching: {e}")
            return None, 0.0

    def hybrid_map_item(self, description: str, template_items: List[str], 
                       section_context: str, statement_type: str) -> Tuple[Optional[str], float, str]:
        """
        Hybrid mapping approach: rule-based first, then LLM for low confidence.
        Returns (template_item, confidence_score, method_used)
        """
        # Step 1: Try rule-based mapping first
        if statement_type == 'balance_sheet':
            rule_match, rule_confidence = self.apply_rule_based_mapping(description)
        elif statement_type == 'income_statement':
            rule_match, rule_confidence = self.apply_rule_based_mapping_is(description)
        elif statement_type == 'cash_flow':
            rule_match, rule_confidence = self.apply_rule_based_mapping_cfs(description)
        else:
            rule_match, rule_confidence = None, 0.0
        
        # If rule-based gives high confidence, use it
        if rule_match and rule_confidence >= 0.7:
            return rule_match, rule_confidence, "rule_based"
        
        # Step 2: Try semantic matching
        semantic_match, semantic_score = self.get_semantic_match(description, template_items, section_context)
        
        # If semantic gives high confidence, use it
        if semantic_match and semantic_score >= 0.6:
            return semantic_match, semantic_score, "semantic"
        
        # Step 3: Use LLM as tie-breaker for low confidence cases
        if self.use_llm and (rule_confidence < 0.5 or semantic_score < 0.5):
            llm_match, llm_confidence, reasoning = self.llm_mapper.map_with_llm(
                description, template_items, section_context, statement_type
            )
            
            if llm_match and llm_confidence > 0.4:
                return llm_match, llm_confidence, f"llm: {reasoning}"
        
        # Step 4: Return best available option
        if rule_match and rule_confidence > semantic_score:
            return rule_match, rule_confidence, "rule_based_fallback"
        elif semantic_match:
            return semantic_match, semantic_score, "semantic_fallback"
        else:
            return None, 0.0, "no_match"

def main():
    # Get project root directory
    current_dir = Path(__file__).resolve().parent
    project_root = current_dir.parent.parent
    
    # Get paths
    template_path = project_root / "templates" / "financial_template.xlsx"
    if not template_path.exists():
        print(f"Template not found at {template_path}")
        return
        
    # Get most recent extracted Excel file
    output_dir = project_root / "output_excel"
    if not output_dir.exists():
        print("No output directory found")
        return
        
    excel_files = [f for f in output_dir.glob("*.xlsx") if not f.name.startswith('~$')]
    if not excel_files:
        print("No valid (non-temporary) Excel files found in output directory")
        return
        
    # Sort by creation time and get most recent
    latest_file = max(excel_files, key=lambda x: x.stat().st_ctime)
    print(f"\nProcessing {latest_file}")
    
    # Read extracted data
    extracted_data = {}
    for sheet in pd.read_excel(latest_file, sheet_name=None).items():
        sheet_name, df = sheet
        statement_type = sheet_name.lower().replace(' ', '_')
        extracted_data[statement_type] = {'2024': {}, '2023': {}}
        
        for _, row in df.iterrows():
            desc = row['Description']
            if pd.notna(desc):
                if pd.notna(row.get('Value_1')):
                    extracted_data[statement_type]['2024'][desc] = row['Value_1']
                if pd.notna(row.get('Value_2')):
                    extracted_data[statement_type]['2023'][row['Description']] = row['Value_2']
    
    # Map to template
    matcher = TemplateMatcher()
    output_path = matcher.map_to_template(extracted_data, str(template_path))
    print(f"\nTemplate populated and saved to: {output_path}")

if __name__ == "__main__":
    main()

