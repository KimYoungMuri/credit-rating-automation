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
        Assigns a section to each line item based on its position in the extracted table.
        Uses manual mapping, regex, and semantic similarity (sentence-transformers) for section assignment.
        Returns a list of dicts: {description, value, section}
        """
        import re
        # Canonical section descriptions for semantic similarity
        section_semantics = {
            'current_assets': 'current assets',
            'noncurrent_assets': 'noncurrent assets',
            'current_liabilities': 'current liabilities',
            'noncurrent_liabilities': 'noncurrent liabilities',
            'equity': 'equity',
        }
        # Section header/total patterns (fuzzy, lowercased)
        section_patterns = [
            ('current_assets', [r'^current assets$', r'^total current assets$', r'^other current assets$', r'cash and cash equivalents', r'margin deposits', r'derivative assets', r'accounts receivable', r'inventories?', r'notes receivable', r'subchapter s income tax deposit']),
            ('noncurrent_assets', [r'^net ppe$', r'^property and equipment', r'^goodwill', r'^intangibles', r'^other noncurrent assets$', r'^total non current assets$', r'^noncurrent assets$', r'deferred compensation plan investments', r'right of use assets', r'finance lease assets']),
            ('current_liabilities', [r'^current liabilities$', r'^total current liabilities$', r'^other current liabilities$', r'^accounts payable$', r'^accrued liabilities$', r'^accrued interest$', r'^derivative liabilities$', r'^short term borrowing$', r'^current portion of long term debt$', r'^contingent consideration payable$', r'^finance lease liability—current portion$', r'^operating lease liability—current portion$', r'^long-term incentive—current portion$', r'^subchapter s income tax deposit obligation$']),
            ('noncurrent_liabilities', [r'^long[- ]term debt$', r'^deferred income taxes$', r'^other noncurrent liabilities$', r'^total non current liabilities$', r'^noncurrent liabilities$', r'^revolving lines of credit$', r'^long-term incentive$', r'^deferred compensation$', r'^finance lease liability$', r'^operating lease liability$']),
            ('equity', [r'^common stock$', r'^retained earnings$', r'^paid in capital$', r'^total equity$', r'^equity$', r'^shareholder', r'^noncontrolling interests$', r"^total common shareholders' equity$"])
        ]
        # Manual section overrides for common items (lowercase)
        manual_section_map = {
            'cash and cash equivalents': 'current_assets',
            'cash and equivalents': 'current_assets',
            'cash and cash equivalent': 'current_assets',
            'margin deposits': 'current_assets',
            'margin deposit': 'current_assets',
            'derivative assets': 'current_assets',
            'accounts receivable': 'current_assets',
            'accounts receivable—net': 'current_assets',
            'inventory': 'current_assets',
            'inventory—net': 'current_assets',
            'inventories—net': 'current_assets',
            'notes receivable': 'current_assets',
            'notes receivable—current portion': 'current_assets',
            'subchapter s income tax deposit—current portion': 'current_assets',
            'other current assets': 'current_assets',
            'total current assets': 'current_assets',
            'goodwill': 'noncurrent_assets',
            'goodwill—net': 'noncurrent_assets',
            'property and equipment—net': 'noncurrent_assets',
            'property and equipment': 'noncurrent_assets',
            'net ppe': 'noncurrent_assets',
            'finance lease assets—net': 'noncurrent_assets',
            'finance lease assets': 'noncurrent_assets',
            'right of use assets—net': 'noncurrent_assets',
            'right of use assets': 'noncurrent_assets',
            'intangibles': 'noncurrent_assets',
            'other intangible assets—net': 'noncurrent_assets',
            'other intangible assets': 'noncurrent_assets',
            'deferred compensation plan investments': 'noncurrent_assets',
            'other noncurrent assets': 'noncurrent_assets',
            'accounts payable': 'current_liabilities',
            'accrued liabilities': 'current_liabilities',
            'accrued interest': 'current_liabilities',
            'derivative liabilities': 'current_liabilities',
            'sales, excise and property taxes payable': 'current_liabilities',
            'contingent consideration payable': 'current_liabilities',
            'finance lease liability—current portion': 'current_liabilities',
            'operating lease liability—current portion': 'current_liabilities',
            'long-term incentive—current portion': 'current_liabilities',
            'subchapter s income tax deposit obligation': 'current_liabilities',
            'current portion of long term debt': 'current_liabilities',
            'long-term debt—current portion': 'current_liabilities',
            'total current liabilities': 'current_liabilities',
            'long term debt': 'noncurrent_liabilities',
            'long-term debt': 'noncurrent_liabilities',
            'revolving lines of credit': 'noncurrent_liabilities',
            'long-term incentive': 'noncurrent_liabilities',
            'deferred compensation': 'noncurrent_liabilities',
            'finance lease liability': 'noncurrent_liabilities',
            'operating lease liability': 'noncurrent_liabilities',
            'deferred income taxes': 'noncurrent_liabilities',
            'other noncurrent liabilities': 'noncurrent_liabilities',
            'total liabilities': 'noncurrent_liabilities',
            'common stock': 'equity',
            'retained earnings': 'equity',
            'paid in capital': 'equity',
            "total common shareholders' equity": 'equity',
            'total common shareholders equity': 'equity',
            'noncontrolling interests': 'equity',
            'total equity': 'equity',
            'total shareholders equity': 'equity',
            'total stockholders equity': 'equity',
            "total shareholders' equity": 'equity',
            'total shareholders equity': 'equity',
            "total stockholders' equity": 'equity',
            'total stockholders equity': 'equity',
        }
        def normalize(s):
            import re
            if not s:
                return ''
            s = s.lower()
            # Replace all dash types with a space
            s = re.sub(r'[\u2013\u2014\u2012\u2010\-]', ' ', s)  # en dash, em dash, figure dash, hyphen, minus
            s = re.sub(r'—', ' ', s)  # em dash (redundant, but explicit)
            s = re.sub(r'–', ' ', s)  # en dash (redundant, but explicit)
            s = re.sub(r'\s+', ' ', s)
            s = s.strip()
            return s
        assigned = []
        current_section = None
        section_confidence = 0  # Track confidence in current section
        
        # Track major section transitions
        major_sections = ['current_assets', 'noncurrent_assets', 'current_liabilities', 'noncurrent_liabilities', 'equity']
        section_transition_keywords = {
            'current_assets': ['total current assets', 'current assets'],
            'noncurrent_assets': ['total assets', 'total noncurrent assets', 'total non current assets'],
            'current_liabilities': ['total current liabilities', 'current liabilities'],
            'noncurrent_liabilities': ['total liabilities', 'total noncurrent liabilities'],
            'equity': ['total equity', 'total shareholders equity', 'total stockholders equity']
        }
        
        # Section boundary keywords that force transitions
        section_boundary_keywords = {
            'assets_to_liabilities': [
                'accounts payable', 'derivative liabilities', 'sales excise and property taxes payable',
                'accrued liabilities', 'long-term incentive', 'contingent consideration payable',
                'subchapter s income tax deposit obligation', 'finance lease liability',
                'operating lease liability', 'long-term debt', 'revolving lines of credit',
                'deferred compensation', 'other noncurrent liabilities',
                # Add all current liability items explicitly
                'accounts payable', 'derivative liabilities', 'sales, excise and property taxes payable',
                'accrued liabilities', 'long-term incentive—current portion', 'contingent consideration payable',
                'subchapter s income tax deposit obligation', 'finance lease liability—current portion',
                'operating lease liability—current portion', 'long-term debt—current portion',
                'long term debt—current portion', 'long-term debt-current portion', 'long term debt-current portion',
                'finance lease liability-current portion', 'operating lease liability-current portion',
                'long-term incentive-current portion', 'long term incentive-current portion',
            ],
            'liabilities_to_equity': [
                "total common shareholders' equity",
                'total common shareholders equity',
                "total shareholders' equity",
                'total shareholders equity',
                "total stockholders' equity",
                'total stockholders equity',
                'common stock', 'additional paid-in capital',
                'retained earnings', 'accumulated other comprehensive income',
                'treasury stock', 'noncontrolling interests', 'total equity', 'total shareholders equity'
            ]
        }
        
        for idx, line in enumerate(extracted_lines):
            desc = normalize(line['description'])
            original_desc = line['description']
            print(f"[DEBUG] Normalized desc: '{desc}' (original: '{original_desc}')")
            
            # Check for major section transitions first
            section_transition_detected = False
            for section, keywords in section_transition_keywords.items():
                if any(keyword in desc for keyword in keywords):
                    current_section = section
                    section_confidence = 9
                    section_transition_detected = True
                    self.extraction_logger.info(f"[SECTION TRANSITION] '{original_desc}' -> {section} (major_section_transition)")
                    break
            
            # Check for section boundary keywords that force transitions
            if not section_transition_detected:
                for boundary_type, keywords in section_boundary_keywords.items():
                    for kw in keywords:
                        if normalize(kw) == desc:
                            print(f"[DEBUG] Section boundary hit: '{desc}' matches '{kw}'")
                            if boundary_type == 'assets_to_liabilities':
                                if any(liability_keyword in desc for liability_keyword in ['current portion', 'payable', 'accrued']):
                                    current_section = 'current_liabilities'
                                else:
                                    current_section = 'noncurrent_liabilities'
                                section_confidence = 9
                                section_transition_detected = True
                                self.extraction_logger.info(f"[SECTION BOUNDARY] '{original_desc}' -> {current_section} (assets_to_liabilities_boundary)")
                                break
                            elif boundary_type == 'liabilities_to_equity':
                                current_section = 'equity'
                                section_confidence = 9
                                section_transition_detected = True
                                self.extraction_logger.info(f"[SECTION BOUNDARY] '{original_desc}' -> {current_section} (liabilities_to_equity_boundary)")
                                break
            
            # Manual override (highest priority)
            if desc in manual_section_map:
                print(f"[DEBUG] Manual map hit: '{desc}' -> {manual_section_map[desc]}")
                assigned_section = manual_section_map[desc]
                debug_reason = 'manual_map'
                current_section = assigned_section
                section_confidence = 10
                self.extraction_logger.info(f"[SECTION ASSIGN] '{original_desc}' -> {assigned_section} (manual_map)")
            else:
                # Check for explicit section headers first
                explicit_section_found = False
                for section, patterns in section_patterns:
                    for pat in patterns:
                        if pat.startswith('^') and pat.endswith('$'):
                            if re.search(pat, desc):
                                assigned_section = section
                                debug_reason = f'exact_regex:{pat}'
                                current_section = assigned_section
                                section_confidence = 8
                                explicit_section_found = True
                                self.extraction_logger.info(f"[SECTION ASSIGN] '{original_desc}' -> {assigned_section} (exact_regex:{pat})")
                                break
                    if explicit_section_found:
                        break
                if not explicit_section_found:
                    # Check for partial section matches (lower confidence)
                    partial_section_found = False
                    for section, patterns in section_patterns:
                        for pat in patterns:
                            if not pat.startswith('^') and not pat.endswith('$'):
                                if re.search(pat, desc):
                                    assigned_section = section
                                    debug_reason = f'partial_regex:{pat}'
                                    current_section = assigned_section
                                    section_confidence = 6
                                    partial_section_found = True
                                    self.extraction_logger.info(f"[SECTION ASSIGN] '{original_desc}' -> {assigned_section} (partial_regex:{pat})")
                                    break
                        if partial_section_found:
                            break
                # If no section header found, use context-based assignment
                if not explicit_section_found and not partial_section_found:
                    if current_section:
                        assigned_section = current_section
                        section_confidence = 7
                        debug_reason = 'context_continuation'
                        self.extraction_logger.info(f"[SECTION ASSIGN] '{original_desc}' -> {assigned_section} (context_continuation)")
                    else:
                        # No current section, try semantic fallback
                        assigned_section = None
                        debug_reason = 'no_context'
                        self.extraction_logger.info(f"[SECTION ASSIGN] '{original_desc}' -> None (no_context)")
            # Semantic fallback only if no section assigned and no current context
            if not assigned_section and not current_section:
                best_score = 0
                best_section = None
                for section, section_desc in section_semantics.items():
                    score = self.get_similarity(desc, section_desc)
                    print(f"[SECTION SEMANTIC DEBUG] '{desc}' vs '{section_desc}' -> score: {score:.3f}")
                    if score > best_score:
                        best_score = score
                        best_section = section
                if best_score > 0.3:
                    assigned_section = best_section
                    current_section = assigned_section
                    section_confidence = 5
                    debug_reason = f'semantic:{section_semantics[best_section]}:{best_score:.3f}'
                    self.extraction_logger.info(f"[SECTION ASSIGN] '{original_desc}' -> {assigned_section} (semantic:{section_semantics[best_section]}:{best_score:.3f})")
            # Debug output for section assignment
            if assigned_section is None:
                prev_desc = normalize(extracted_lines[idx-1]['description']) if idx > 0 else ''
                next_desc = normalize(extracted_lines[idx+1]['description']) if idx < len(extracted_lines)-1 else ''
                print(f"[DEBUG] Section=None for '{desc}'. Prev='{prev_desc}' Next='{next_desc}'")
                self.extraction_logger.info(f"[SECTION ASSIGN] '{original_desc}' -> None (no_match)")
            else:
                print(f"[SECTION ASSIGN] '{line['description']}' -> {assigned_section} ({debug_reason}) [confidence:{section_confidence}]")
            assigned.append({
                'description': line['description'],
                'value': line['numbers'][0] if line['numbers'] else None,
                'section': assigned_section
            })
        
        return assigned

    def assign_sections_by_context_is(self, extracted_lines: list) -> list:
        """
        Assigns IS sections to each line item based on its content and context.
        Returns a list of dicts: {description, value, section}
        """
        import re
        
        # IS section semantics for semantic similarity
        section_semantics = {
            'revenue': 'revenue sales income',
            'operating_expenses': 'operating expenses cost of goods sold depreciation amortization',
            'other_income_expense': 'interest expense other income expense gain loss',
            'tax_net_income': 'income tax benefit net profit loss'
        }
        
        def normalize(s):
            if not s:
                return ''
            s = s.lower()
            s = re.sub(r'[^a-z0-9 ]', ' ', s)  # Remove punctuation
            s = re.sub(r'\s+', ' ', s)
            s = s.strip()
            return s

        # Expanded manual section mapping for IS
        manual_section_map = {
            # Revenue items - robust coverage
            'revenue': 'revenue',
            'net revenue': 'revenue',
            'total revenue': 'revenue',
            'sales': 'revenue',
            'net sales': 'revenue',
            'gross sales': 'revenue',
            'total sales': 'revenue',
            'operating revenue': 'revenue',
            'service revenue': 'revenue',
            'product revenue': 'revenue',
            'advertising revenue': 'revenue',
            'broadcasting revenue': 'revenue',
            'media revenue': 'revenue',
            'subscription revenue': 'revenue',
            'license revenue': 'revenue',
            'commission revenue': 'revenue',
            'fee revenue': 'revenue',
            'other revenue': 'revenue',
            'other income': 'revenue',
            'miscellaneous revenue': 'revenue',
            'miscellaneous income': 'revenue',
        }

        assigned = []
        current_section = None
        for line in extracted_lines:
            desc = normalize(line['description'])
            original_desc = line['description']

            # Debug: print normalized description
            print(f"[DEBUG IS] Normalized desc: '{desc}' (original: '{original_desc}')")

            # Loosened regex for revenue lines
            revenue_match = re.search(r'(sales|revenue)', desc)
            print(f"[DEBUG IS] Revenue regex match: {bool(revenue_match)} for '{desc}'")
            if revenue_match:
                assigned_section = 'revenue'
                current_section = assigned_section
                self.extraction_logger.info(f"[SECTION ASSIGN] '{original_desc}' -> {assigned_section} (regex_revenue)")
            elif desc in manual_section_map:
                assigned_section = manual_section_map[desc]
                current_section = assigned_section
                self.extraction_logger.info(f"[SECTION ASSIGN] '{original_desc}' -> {assigned_section} (manual_map)")
            else:
                # Semantic matching
                best_score = 0
                best_section = None
                for section, section_desc in section_semantics.items():
                    score = self.get_similarity(desc, section_desc)
                    if score > best_score:
                        best_score = score
                        best_section = section
                if best_score > 0.3:
                    assigned_section = best_section
                    current_section = assigned_section
                    self.extraction_logger.info(f"[SECTION ASSIGN] '{original_desc}' -> {assigned_section} (semantic:{section_semantics[best_section]}:{best_score:.3f})")
                else:
                    if current_section:
                        assigned_section = current_section
                        self.extraction_logger.info(f"[SECTION ASSIGN] '{original_desc}' -> {assigned_section} (context_continuation)")
                    else:
                        assigned_section = None
                        self.extraction_logger.info(f"[SECTION ASSIGN] '{original_desc}' -> None (no_context)")
            assigned.append({
                'description': line['description'],
                'value': line['numbers'][0] if line['numbers'] else None,
                'section': assigned_section
            })
        return assigned

    def assign_sections_by_context_cfs(self, extracted_lines: list) -> list:
        """
        Assigns CFS sections to each line item based on its content and context.
        Returns a list of dicts: {description, value, section}
        """
        import re
        
        # CFS section semantics for semantic similarity
        section_semantics = {
            'operating_activities': 'operating activities net income depreciation changes in working capital',
            'investing_activities': 'investing activities capital expenditure purchase of assets proceeds from sale',
            'financing_activities': 'financing activities debt issuance stock issuance dividend payment',
            'other': 'other'
        }
        
        # Manual section mapping for CFS
        manual_section_map = {
            'net profit loss': 'operating_activities',
            'net income': 'operating_activities',
            'depreciation': 'operating_activities',
            'deferred income taxes': 'operating_activities',
            'impairment and other losses': 'operating_activities',
            'changes in operating assets and liabilities': 'operating_activities',
            'net cash provided by used in operating activities': 'operating_activities',
            'purchases of property and equipment': 'investing_activities',
            'capital expenditure': 'investing_activities',
            'proceeds from sale of assets': 'investing_activities',
            'net cash used in investing activities': 'investing_activities',
            'proceeds from issuance': 'financing_activities',
            'principal payments': 'financing_activities',
            'net cash provided by used in financing activities': 'financing_activities'
        }
        
        def normalize(s):
            if not s:
                return ''
            s = s.lower()
            s = re.sub(r'[\u2013\u2014\u2012\u2010\-]', ' ', s)
            s = re.sub(r'—', ' ', s)
            s = re.sub(r'–', ' ', s)
            s = re.sub(r'\s+', ' ', s)
            s = s.strip()
            return s
        
        assigned = []
        current_section = None
        
        for line in extracted_lines:
            desc = normalize(line['description'])
            original_desc = line['description']
            
            # Manual override (highest priority)
            if desc in manual_section_map:
                assigned_section = manual_section_map[desc]
                current_section = assigned_section
                self.extraction_logger.info(f"[SECTION ASSIGN] '{original_desc}' -> {assigned_section} (manual_map)")
            else:
                # Semantic matching
                best_score = 0
                best_section = None
                for section, section_desc in section_semantics.items():
                    score = self.get_similarity(desc, section_desc)
                    if score > best_score:
                        best_score = score
                        best_section = section
                
                if best_score > 0.3:
                    assigned_section = best_section
                    current_section = assigned_section
                    self.extraction_logger.info(f"[SECTION ASSIGN] '{original_desc}' -> {assigned_section} (semantic:{section_semantics[best_section]}:{best_score:.3f})")
                else:
                    # Use context continuation
                    if current_section:
                        assigned_section = current_section
                        self.extraction_logger.info(f"[SECTION ASSIGN] '{original_desc}' -> {assigned_section} (context_continuation)")
                    else:
                        assigned_section = None
                        self.extraction_logger.info(f"[SECTION ASSIGN] '{original_desc}' -> None (no_context)")
            
            assigned.append({
                'description': line['description'],
                'value': line['numbers'][0] if line['numbers'] else None,
                'section': assigned_section
            })
        
        return assigned

    def get_is_row_map(self, template_items: list, section: str) -> dict:
        """Get row mapping for IS template items"""
        # Map IS template items to Excel rows based on the template structure
        is_row_mapping = {
            'Revenue': 6,
            'Other Revenue': 7,
            'Cost of revenue': 9,
            'Station operations costs': 10,
            'Payroll and related costs': 11,
            'Depreciation and amortization': 12,
            'Impairment and other losses': 13,
            'Selling, general and administrative expenses': 14,
            'Other Operating Expenses': 15,
            'Interest expense': 16,
            'Other (income) and expense, net': 17,
            'Other Income/Expense': 18,
            'Income tax benefit': 19,
            'Net profit (loss)': 20
        }
        
        return {item: is_row_mapping.get(item) for item in template_items if item in is_row_mapping}

    def get_cfs_row_map(self, template_items: list, section: str) -> dict:
        """Get row mapping for CFS template items"""
        # Map CFS template items to Excel rows based on the template structure
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
        """
        Map extracted data to template structure using hybrid approach:
        1. Rule-based filters and regex matching
        2. Semantic similarity with sentence-transformers
        3. Confidence scoring and manual review flags
        """
        self.used_items = set()
        output_dir = Path("output_excel")
        output_dir.mkdir(exist_ok=True)
        timestamp = pd.Timestamp.now().strftime("%Y%m%d_%H%M%S")
        output_path = output_dir / f"populated_template_{timestamp}.xlsx"
        shutil.copy2(template_path, output_path)
        wb = load_workbook(output_path)
        bs_sheet = wb['BS']
        is_sheet = wb['IS.CF']  # Add IS/CF sheet definition
        year_cols = {'2024': 'F', '2023': 'E'}

        # --- Enhanced Section definitions with better categorization ---
        SECTION_TEMPLATE = {
            'current_assets': {
                'rows': list(range(7, 14)),
                'template': [
                    'Cash and equivalents', 'Accounts Receivable', 'Prepaid Expenses',
                    'Inventory', 'Investments', 'Other'
                ]
            },
            'noncurrent_assets': {
                'rows': list(range(15, 21)),
                'template': [
                    'Net PPE', 'Goodwill', 'Intangibles', 'Other'
                ]
            },
            'current_liabilities': {
                'rows': list(range(24, 30)),
                'template': [
                    'Accounts Payable', 'Accrued Interest', 'Short term Borrowing',
                    'Current Portion of Long Term Debt', 'Other'
                ]
            },
            'noncurrent_liabilities': {
                'rows': list(range(31, 36)),
                'template': [
                    'Long Term Debt', 'Deferred income taxes', 'Other'
                ]
            },
            'equity': {
                'rows': list(range(39, 44)),
                'template': [
                    'Common Stock', 'Retained Earnings', 'Paid in Capital', 'Other'
                ]
            }
        }

        # --- Enhanced rule-based mapping with regex patterns ---
        RULE_BASED_MAP = {
            # Current Assets
            r'cash\s+(?:and\s+)?(?:cash\s+)?equivalents?': 'Cash and equivalents',
            r'accounts?\s+receivable(?:[—-]net)?': 'Accounts Receivable',
            r'prepaid\s+expenses?': 'Prepaid Expenses',
            r'inventor(?:y|ies)(?:[—-]net)?': 'Inventory',
            r'notes?\s+receivable[—-]current\s+portion': 'Accounts Receivable',  # Current portion
            r'margin\s+deposits?': 'Investments',
            r'derivative\s+assets?': 'Investments',
            r'subchapter\s+s\s+income\s+tax\s+deposit[—-]current\s+portion': 'Other',  # Current portion
            r'other\s+current\s+assets?': 'Other',  # Add this line
            
            # Non-Current Assets
            r'property\s+(?:and\s+)?equipment(?:[—-]net)?': 'Net PPE',
            r'net\s+ppe': 'Net PPE',
            r'goodwill(?:[—-]net)?': 'Goodwill',
            r'(?:other\s+)?intangible\s+assets?(?:[—-]net)?': 'Intangibles',
            r'right\s+of\s+use\s+assets?(?:[—-]net)?': 'Other',  # Map to Other since no specific template line
            r'finance\s+lease\s+assets?(?:[—-]net)?': 'Other',  # Map to Other since no specific template line
            r'deferred\s+compensation\s+plan\s+investments?': 'Other',  # Map to Other since no specific template line
            r'other\s+noncurrent\s+assets?': 'Other',  # Add this line
            r'notes?\s+receivable(?!.*current)': 'Other',  # Noncurrent notes receivable (no "current" in description)
            r'subchapter\s+s\s+income\s+tax\s+deposit(?!.*current)': 'Other',  # Noncurrent tax deposit (no "current" in description)
            
            # Current Liabilities
            r'accounts?\s+payable': 'Accounts Payable',
            r'accrued\s+(?:liabilities?|interest)': 'Accrued Interest',
            r'revolving\s+lines?\s+of\s+credit': 'Short term Borrowing',
            r'long[- ]term\s+debt[—-]current\s+portion': 'Current Portion of Long Term Debt',
            r'derivative\s+liabilities?': 'Other',
            r'sales,?\s+excise\s+and\s+property\s+taxes?\s+payable': 'Other',
            r'long[- ]term\s+incentive[—-]current\s+portion': 'Other',
            r'contingent\s+consideration\s+payable': 'Other',
            r'finance\s+lease\s+liability[—-]current\s+portion': 'Other',
            r'operating\s+lease\s+liability[—-]current\s+portion': 'Other',
            r'other\s+current\s+liabilities?': 'Other',  # Add this line
            
            # Non-Current Liabilities
            r'long[- ]term\s+debt(?!.*current)': 'Long Term Debt',
            r'deferred\s+income\s+taxes?': 'Deferred income taxes',
            r'long[- ]term\s+incentive(?!.*current)': 'Other',
            r'deferred\s+compensation(?!.*plan)': 'Other',
            r'finance\s+lease\s+liability(?!.*current)': 'Other',
            r'operating\s+lease\s+liability(?!.*current)': 'Other',
            r'other\s+noncurrent\s+liabilities?': 'Other',  # Add this line
            
            # Equity
            r'common\s+stock': 'Common Stock',
            r'retained\s+earnings?': 'Retained Earnings',
            r'paid[- ]in\s+capital': 'Paid in Capital',
            r'total\s+common\s+shareholders?\s+equity': 'Retained Earnings',  # Map to Retained Earnings as main equity
            r'noncontrolling\s+interests?': 'Other (Equity)',  # Map to Other (Equity) specifically
        }

        def normalize(s):
            """Enhanced normalization with better preprocessing"""
            import re
            try:
                from nltk.stem import PorterStemmer
                stemmer = PorterStemmer()
                def stem_word(word):
                    return stemmer.stem(word)
            except ImportError:
                def stem_word(word):
                    return word.rstrip('s')
            
            if not s:
                return ''
            
            s = s.lower()
            # Remove common financial suffixes
            s = re.sub(r'[—-]net\b', '', s)
            s = re.sub(r'[—-]current\s+portion\b', '', s)
            s = re.sub(r'\bportion\b', '', s)
            s = re.sub(r'\bcurrent\b', '', s)
            s = re.sub(r'\bnoncurrent\b', '', s)
            s = re.sub(r'\bof\b', '', s)
            s = re.sub(r'\bthe\b', '', s)
            s = re.sub(r'\band\b', '', s)
            # Remove punctuation and normalize spaces
            s = re.sub(r'[^\w\s]', ' ', s)
            s = re.sub(r'\s+', ' ', s)
            s = s.strip()
            # Stem each word
            s = ' '.join(stem_word(word) for word in s.split())
            return s

        def apply_rule_based_mapping(description: str) -> tuple[str, float]:
            """Apply rule-based mapping with confidence scoring"""
            desc_lower = description.lower()
            
            # Check for exact matches first
            for pattern, template_item in RULE_BASED_MAP.items():
                if re.search(pattern, desc_lower, re.IGNORECASE):
                    return template_item, 1.0  # High confidence for rule-based matches
            
            return None, 0.0

        def get_semantic_match(description: str, template_items: list, section: str) -> tuple[str, float]:
            """Get semantic match with confidence scoring"""
            desc_norm = normalize(description)
            best_match = None
            best_score = 0.0
            
            for template_item in template_items:
                template_norm = normalize(template_item)
                score = self.get_similarity(desc_norm, template_norm)
                
                # Add section-specific bonuses
                if section == 'current_assets':
                    if any(term in desc_norm for term in ['cash', 'receivable', 'inventory', 'prepaid']):
                        score += 0.1
                elif section == 'noncurrent_assets':
                    if any(term in desc_norm for term in ['property', 'equipment', 'goodwill', 'intangible']):
                        score += 0.1
                elif section == 'current_liabilities':
                    if any(term in desc_norm for term in ['payable', 'accrued', 'debt', 'borrowing']):
                        score += 0.1
                elif section == 'noncurrent_liabilities':
                    if any(term in desc_norm for term in ['debt', 'deferred', 'lease', 'liability']):
                        score += 0.1
                elif section == 'equity':
                    if any(term in desc_norm for term in ['stock', 'equity', 'capital', 'earning']):
                        score += 0.1
                
                if score > best_score:
                    best_score = score
                    best_match = template_item
            
            return best_match, best_score

        # --- Enhanced mapping logic ---
        if 'balance_sheet' in extracted_data:
            for year in ['2024', '2023']:
                if year not in extracted_data['balance_sheet']:
                    continue
                
                # Build ordered list of lines for this year
                lines = []
                for desc, value in extracted_data['balance_sheet'][year].items():
                    lines.append({'description': desc, 'numbers': [value] if value is not None else []})
                
                # Assign sections by context
                assigned = self.assign_sections_by_context(lines)
                self.print_section_assignments(assigned, year, 'balance_sheet')
                
                # For each section, map to template
                for section, section_info in SECTION_TEMPLATE.items():
                    col = year_cols[year]
                    template_items = section_info['template']
                    row_map = {template_items[i]: section_info['rows'][i] for i in range(len(template_items))}
                    used_template_rows = set()
                    used_extracted = set()
                    other_sum = 0
                    manual_review_items = []  # Track items needing manual review
                    # --- NEW: Accumulate values for each template line ---
                    accumulated_values = {t: 0.0 for t in template_items}
                    # --- NEW: Track section sum for total detection ---
                    section_sum = 0.0
                    # Step 1: Prioritize Total/Net rows (but skip mapping them)
                    for entry in assigned:
                        if entry['section'] != section or entry['value'] is None:
                            continue
                        # --- NEW: Detect total rows ---
                        desc = entry['description']
                        val = float(entry['value'])
                        is_total = self.is_total_or_net_row(desc)
                        is_empty_total = (not desc or desc.strip() == '') and abs(val - section_sum) < 2.0 and section_sum > 0
                        if is_total or is_empty_total:
                            print(f"[SKIP TOTAL] '{desc}' ({val}) matches section sum ({section_sum}) or total pattern; skipping mapping.")
                            self.extraction_logger.info(f"[SKIP TOTAL] '{desc}' ({val}) - skipping total/net row in accumulation")
                            continue  # Skip mapping totals
                        # Step 2a: Try rule-based mapping first
                        rule_match, rule_confidence = apply_rule_based_mapping(desc)
                        if rule_match and rule_match in template_items:
                            accumulated_values[rule_match] += val
                            section_sum += val
                            used_extracted.add(desc)
                            print(f"[ACCUMULATE RULE] '{desc}' ({val}) -> {section}::{rule_match}")
                            self.extraction_logger.info(f"[TEMPLATE MAP] '{desc}' ({val}) -> {section}::{rule_match} (rule_based)")
                            # --- FIX: Also add to other_sum if rule_match is "Other" ---
                            if rule_match == 'Other':
                                other_sum += val
                            continue
                        # Step 2b: Try semantic matching
                        semantic_match, semantic_score = get_semantic_match(desc, template_items, section)
                        if semantic_match:
                            if semantic_score >= 0.4:
                                accumulated_values[semantic_match] += val
                                section_sum += val
                                used_extracted.add(desc)
                                print(f"[ACCUMULATE SEMANTIC] '{desc}' ({val}) -> {section}::{semantic_match} [confidence: {semantic_score:.2f}]")
                                self.extraction_logger.info(f"[TEMPLATE MAP] '{desc}' ({val}) -> {section}::{semantic_match} (semantic:{semantic_score:.2f})")
                                # --- FIX: Also add to other_sum if semantic_match is "Other" ---
                                if semantic_match == 'Other':
                                    other_sum += val
                            else:
                                # --- FIX: Only add to Other if not a total/net row ---
                                if not self.is_total_or_net_row(desc):
                                    other_sum += val
                                    section_sum += val
                                    used_extracted.add(desc)
                                    print(f"[ACCUMULATE OTHER] '{desc}' ({val}) -> {section}::Other [best semantic: {semantic_match} ({semantic_score:.2f})]")
                                    self.extraction_logger.info(f"[TEMPLATE MAP] '{desc}' ({val}) -> {section}::Other (semantic_fallback:{semantic_match}:{semantic_score:.2f})")
                        else:
                            # --- FIX: Only add to Other if not a total/net row ---
                            if not self.is_total_or_net_row(desc):
                                other_sum += val
                                section_sum += val
                                used_extracted.add(desc)
                                print(f"[ACCUMULATE OTHER] '{desc}' ({val}) -> {section}::Other (no semantic match)")
                                self.extraction_logger.info(f"[TEMPLATE MAP] '{desc}' ({val}) -> {section}::Other (no_match)")
                            else:
                                print(f"[SKIP TOTAL in OTHER] '{desc}' ({val}) would have gone to Other, but is total/net row.")
                                self.extraction_logger.info(f"[SKIP TOTAL in OTHER] '{desc}' ({val}) - would have gone to Other, but is total/net row")
                    # Write accumulated values to template
                    for t in template_items:
                        if t == 'Other':
                            continue  # Write Other separately
                        row = row_map[t]
                        if accumulated_values[t] != 0:
                            bs_sheet[f"{col}{row}"] = accumulated_values[t]
                    # Write Other sum if any
                    if other_sum != 0:
                        row = row_map['Other']
                        prev_val = bs_sheet[f"{col}{row}"].value
                        if prev_val is not None and prev_val != '':
                            try:
                                other_sum += float(prev_val)
                            except Exception:
                                pass
                        bs_sheet[f"{col}{row}"] = other_sum
                        print(f"[OTHER_SUM] {section}::Other (cell {col}{row}) = {other_sum}")
                        self.extraction_logger.info(f"[WRITE OTHER] {section}::Other (cell {col}{row}) = {other_sum}")
                    else:
                        self.extraction_logger.info(f"[NO OTHER] {section}::Other = 0 (no items mapped to Other)")
                    # Print manual review items (unchanged)
                    if manual_review_items:
                        print(f"\n[MANUAL REVIEW NEEDED] {section} {year}:")
                        for desc, template_item, score in manual_review_items:
                            print(f"  '{desc}' -> '{template_item}' (confidence: {score:.2f})")
        
        # --- Add IS and CFS section structures and templates ---
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
            }
        }

        # --- Add section boundary keywords for IS and CFS ---
        section_boundary_keywords['is_section_boundaries'] = [
            'revenue', 'cost of revenue', 'station operations costs', 'payroll and related costs',
            'depreciation and amortization', 'impairment and other losses', 'selling general and administrative expenses',
            'income from operations', 'interest expense', 'other income and expense net', 'income tax benefit', 'net profit loss'
        ]
        section_boundary_keywords['cfs_section_boundaries'] = [
            'cash flows from operating activities', 'cash flows from investing activities', 'cash flows from financing activities',
            'net cash provided by used in operating activities', 'net cash used in investing activities', 'net cash provided by used in financing activities'
        ]

        # --- Add IS and CFS mapping logic (mirroring BS logic) ---
        
        # Income Statement Mapping
        if 'income_statement' in extracted_data:
            for year in ['2024', '2023']:
                if year not in extracted_data['income_statement']:
                    continue
                
                # Build ordered list of lines for this year
                lines = []
                for desc, value in extracted_data['income_statement'][year].items():
                    lines.append({'description': desc, 'numbers': [value] if value is not None else []})
                
                # Assign sections by context (using IS-specific logic)
                assigned = self.assign_sections_by_context_is(lines)
                self.print_section_assignments(assigned, year, 'income_statement')
                
                # For each section, map to template
                for section, section_info in IS_SECTION_TEMPLATE.items():
                    col = year_cols[year]
                    template_items = section_info['template']
                    # Map template items to IS sheet rows
                    row_map = self.get_is_row_map(template_items, section)
                    used_template_rows = set()
                    used_extracted = set()
                    other_sum = 0
                    manual_review_items = []
                    accumulated_values = {t: 0.0 for t in template_items}
                    section_sum = 0.0
                    
                    # Process each assigned line
                    for entry in assigned:
                        if entry['section'] != section or entry['value'] is None:
                            continue
                        
                        desc = entry['description']
                        val = float(entry['value'])
                        is_total = self.is_total_or_net_row(desc)
                        is_empty_total = (not desc or desc.strip() == '') and abs(val - section_sum) < 2.0 and section_sum > 0
                        
                        if is_total or is_empty_total:
                            print(f"[SKIP TOTAL] '{desc}' ({val}) matches section sum ({section_sum}) or total pattern; skipping mapping.")
                            self.extraction_logger.info(f"[SKIP TOTAL] '{desc}' ({val}) - skipping total/net row in accumulation")
                            continue
                        
                        # Try rule-based mapping first
                        rule_match, rule_confidence = self.apply_rule_based_mapping_is(desc)
                        if rule_match and rule_match in template_items:
                            accumulated_values[rule_match] += val
                            section_sum += val
                            used_extracted.add(desc)
                            print(f"[ACCUMULATE RULE] '{desc}' ({val}) -> {section}::{rule_match}")
                            self.extraction_logger.info(f"[TEMPLATE MAP] '{desc}' ({val}) -> {section}::{rule_match} (rule_based)")
                            if rule_match == 'Other Operating Expenses':
                                other_sum += val
                            continue
                        
                        # Try semantic matching
                        semantic_match, semantic_score = self.get_semantic_match_is(desc, template_items, section)
                        if semantic_match:
                            if semantic_score >= 0.4:
                                accumulated_values[semantic_match] += val
                                section_sum += val
                                used_extracted.add(desc)
                                print(f"[ACCUMULATE SEMANTIC] '{desc}' ({val}) -> {section}::{semantic_match} [confidence: {semantic_score:.2f}]")
                                self.extraction_logger.info(f"[TEMPLATE MAP] '{desc}' ({val}) -> {section}::{semantic_match} (semantic:{semantic_score:.2f})")
                                if semantic_match == 'Other Operating Expenses':
                                    other_sum += val
                            else:
                                if not self.is_total_or_net_row(desc):
                                    other_sum += val
                                    section_sum += val
                                    used_extracted.add(desc)
                                    print(f"[ACCUMULATE OTHER] '{desc}' ({val}) -> {section}::Other Operating Expenses [best semantic: {semantic_match} ({semantic_score:.2f})]")
                                    self.extraction_logger.info(f"[TEMPLATE MAP] '{desc}' ({val}) -> {section}::Other Operating Expenses (semantic_fallback:{semantic_match}:{semantic_score:.2f})")
                        else:
                            if not self.is_total_or_net_row(desc):
                                other_sum += val
                                section_sum += val
                                used_extracted.add(desc)
                                print(f"[ACCUMULATE OTHER] '{desc}' ({val}) -> {section}::Other Operating Expenses (no semantic match)")
                                self.extraction_logger.info(f"[TEMPLATE MAP] '{desc}' ({val}) -> {section}::Other Operating Expenses (no_match)")
                            else:
                                print(f"[SKIP TOTAL in OTHER] '{desc}' ({val}) would have gone to Other, but is total/net row.")
                                self.extraction_logger.info(f"[SKIP TOTAL in OTHER] '{desc}' ({val}) - would have gone to Other, but is total/net row")
                    
                    # Write accumulated values to template
                    for t in template_items:
                        if t == 'Other Operating Expenses':
                            continue  # Write Other separately
                        if t in row_map:
                            row = row_map[t]
                            if accumulated_values[t] != 0:
                                is_sheet[f"{col}{row}"] = accumulated_values[t]
                                print(f"[WRITE IS] {section}::{t} (cell {col}{row}) = {accumulated_values[t]}")
                                self.extraction_logger.info(f"[WRITE IS] {section}::{t} (cell {col}{row}) = {accumulated_values[t]}")
                    
                    # Write Other sum if any
                    if other_sum != 0 and 'Other Operating Expenses' in row_map:
                        row = row_map['Other Operating Expenses']
                        prev_val = is_sheet[f"{col}{row}"].value
                        if prev_val is not None and prev_val != '':
                            try:
                                other_sum += float(prev_val)
                            except Exception:
                                pass
                        is_sheet[f"{col}{row}"] = other_sum
                        print(f"[OTHER_SUM] {section}::Other Operating Expenses (cell {col}{row}) = {other_sum}")
                        self.extraction_logger.info(f"[WRITE OTHER] {section}::Other Operating Expenses (cell {col}{row}) = {other_sum}")
                    else:
                        self.extraction_logger.info(f"[NO OTHER] {section}::Other Operating Expenses = 0 (no items mapped to Other)")
                    
                    # Print manual review items
                    if manual_review_items:
                        print(f"\n[MANUAL REVIEW NEEDED] {section} {year}:")
                        for desc, template_item, score in manual_review_items:
                            print(f"  '{desc}' -> '{template_item}' (confidence: {score:.2f})")

        # Cash Flow Statement Mapping
        if 'cash_flow' in extracted_data:
            for year in ['2024', '2023']:
                if year not in extracted_data['cash_flow']:
                    continue
                
                # Build ordered list of lines for this year
                lines = []
                for desc, value in extracted_data['cash_flow'][year].items():
                    lines.append({'description': desc, 'numbers': [value] if value is not None else []})
                
                # Assign sections by context (using CFS-specific logic)
                assigned = self.assign_sections_by_context_cfs(lines)
                self.print_section_assignments(assigned, year, 'cash_flow')
                
                # For each section, map to template
                for section, section_info in CFS_SECTION_TEMPLATE.items():
                    col = year_cols[year]
                    template_items = section_info['template']
                    # Map template items to CFS sheet rows
                    row_map = self.get_cfs_row_map(template_items, section)
                    used_template_rows = set()
                    used_extracted = set()
                    other_sum = 0
                    manual_review_items = []
                    accumulated_values = {t: 0.0 for t in template_items}
                    section_sum = 0.0
                    
                    # Process each assigned line
                    for entry in assigned:
                        if entry['section'] != section or entry['value'] is None:
                            continue
                        
                        desc = entry['description']
                        val = float(entry['value'])
                        is_total = self.is_total_or_net_row(desc)
                        is_empty_total = (not desc or desc.strip() == '') and abs(val - section_sum) < 2.0 and section_sum > 0
                        
                        if is_total or is_empty_total:
                            print(f"[SKIP TOTAL] '{desc}' ({val}) matches section sum ({section_sum}) or total pattern; skipping mapping.")
                            self.extraction_logger.info(f"[SKIP TOTAL] '{desc}' ({val}) - skipping total/net row in accumulation")
                            continue
                        
                        # Try rule-based mapping first
                        rule_match, rule_confidence = self.apply_rule_based_mapping_cfs(desc)
                        if rule_match and rule_match in template_items:
                            accumulated_values[rule_match] += val
                            section_sum += val
                            used_extracted.add(desc)
                            print(f"[ACCUMULATE RULE] '{desc}' ({val}) -> {section}::{rule_match}")
                            self.extraction_logger.info(f"[TEMPLATE MAP] '{desc}' ({val}) -> {section}::{rule_match} (rule_based)")
                            if 'Other' in rule_match:
                                other_sum += val
                            continue
                        
                        # Try semantic matching
                        semantic_match, semantic_score = self.get_semantic_match_cfs(desc, template_items, section)
                        if semantic_match:
                            if semantic_score >= 0.4:
                                accumulated_values[semantic_match] += val
                                section_sum += val
                                used_extracted.add(desc)
                                print(f"[ACCUMULATE SEMANTIC] '{desc}' ({val}) -> {section}::{semantic_match} [confidence: {semantic_score:.2f}]")
                                self.extraction_logger.info(f"[TEMPLATE MAP] '{desc}' ({val}) -> {section}::{semantic_match} (semantic:{semantic_score:.2f})")
                                if 'Other' in semantic_match:
                                    other_sum += val
                            else:
                                if not self.is_total_or_net_row(desc):
                                    other_sum += val
                                    section_sum += val
                                    used_extracted.add(desc)
                                    print(f"[ACCUMULATE OTHER] '{desc}' ({val}) -> {section}::Other [best semantic: {semantic_match} ({semantic_score:.2f})]")
                                    self.extraction_logger.info(f"[TEMPLATE MAP] '{desc}' ({val}) -> {section}::Other (semantic_fallback:{semantic_match}:{semantic_score:.2f})")
                        else:
                            if not self.is_total_or_net_row(desc):
                                other_sum += val
                                section_sum += val
                                used_extracted.add(desc)
                                print(f"[ACCUMULATE OTHER] '{desc}' ({val}) -> {section}::Other (no semantic match)")
                                self.extraction_logger.info(f"[TEMPLATE MAP] '{desc}' ({val}) -> {section}::Other (no_match)")
                            else:
                                print(f"[SKIP TOTAL in OTHER] '{desc}' ({val}) would have gone to Other, but is total/net row.")
                                self.extraction_logger.info(f"[SKIP TOTAL in OTHER] '{desc}' ({val}) - would have gone to Other, but is total/net row")
                    
                    # Write accumulated values to template
                    for t in template_items:
                        if 'Other' in t:
                            continue  # Write Other separately
                        if t in row_map:
                            row = row_map[t]
                            if accumulated_values[t] != 0:
                                is_sheet[f"{col}{row}"] = accumulated_values[t]
                                print(f"[WRITE CFS] {section}::{t} (cell {col}{row}) = {accumulated_values[t]}")
                                self.extraction_logger.info(f"[WRITE CFS] {section}::{t} (cell {col}{row}) = {accumulated_values[t]}")
                    
                    # Write Other sum if any
                    if other_sum != 0 and any('Other' in t for t in template_items):
                        other_template = next(t for t in template_items if 'Other' in t)
                        if other_template in row_map:
                            row = row_map[other_template]
                            prev_val = is_sheet[f"{col}{row}"].value
                            if prev_val is not None and prev_val != '':
                                try:
                                    other_sum += float(prev_val)
                                except Exception:
                                    pass
                            is_sheet[f"{col}{row}"] = other_sum
                            print(f"[OTHER_SUM] {section}::{other_template} (cell {col}{row}) = {other_sum}")
                            self.extraction_logger.info(f"[WRITE OTHER] {section}::{other_template} (cell {col}{row}) = {other_sum}")
                        else:
                            self.extraction_logger.info(f"[NO OTHER] {section}::Other = 0 (no items mapped to Other)")
                    
                    # Print manual review items
                    if manual_review_items:
                        print(f"\n[MANUAL REVIEW NEEDED] {section} {year}:")
                        for desc, template_item, score in manual_review_items:
                            print(f"  '{desc}' -> '{template_item}' (confidence: {score:.2f})")

        # --- Add debug output for IS and CFS mapping (mirroring BS logic) ---
        # (The rest of the code for section assignment and mapping should mirror the robust logic used for the balance sheet, using the new IS and CFS structures, manual mapping, and boundaries.)

        wb.save(output_path)
        self.extraction_logger.info("\n==== MAPPING TO TEMPLATE END ====")
        return str(output_path)

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

    def apply_rule_based_mapping_is(self, description: str) -> tuple[str, float]:
        """Apply rule-based mapping for IS items"""
        import re
        desc_lower = description.lower()
        
        # IS-specific rule-based mapping with comprehensive revenue coverage
        is_rules = {
            # Revenue items - all map to main Revenue line
            r'revenue': 'Revenue',
            r'sales': 'Revenue',
            r'net\s+(?:revenue|sales)': 'Revenue',
            r'gross\s+(?:revenue|sales)': 'Revenue',
            r'total\s+(?:revenue|sales)': 'Revenue',
            r'operating\s+revenue': 'Revenue',
            r'service\s+revenue': 'Revenue',
            r'product\s+revenue': 'Revenue',
            r'advertising\s+revenue': 'Revenue',
            r'broadcasting\s+revenue': 'Revenue',
            r'media\s+revenue': 'Revenue',
            r'subscription\s+revenue': 'Revenue',
            r'license\s+revenue': 'Revenue',
            r'commission\s+revenue': 'Revenue',
            r'fee\s+revenue': 'Revenue',
            r'other\s+revenue': 'Revenue',
            r'other\s+income': 'Revenue',
            r'miscellaneous\s+(?:revenue|income)': 'Revenue',
            
            # Operating expenses
            r'cost\s+of\s+(?:revenue|goods\s+sold|sales)': 'Cost of revenue',
            r'station\s+operations?\s+costs?': 'Station operations costs',
            r'payroll': 'Payroll and related costs',
            r'depreciation': 'Depreciation and amortization',
            r'amortization': 'Depreciation and amortization',
            r'impairment': 'Impairment and other losses',
            r'selling.*administrative': 'Selling, general and administrative expenses',
            r'sg&a': 'Selling, general and administrative expenses',
            r'general\s+and\s+administrative': 'Selling, general and administrative expenses',
            r'administrative\s+expenses?': 'Selling, general and administrative expenses',
            r'selling\s+expenses?': 'Selling, general and administrative expenses',
            r'research\s+and\s+development': 'Selling, general and administrative expenses',
            r'r&d': 'Selling, general and administrative expenses',
            
            # Other income/expense
            r'interest\s+expense': 'Interest expense',
            r'interest\s+income': 'Other income and expense, net',
            r'other\s+income\s+and\s+expense': 'Other income and expense, net',
            r'other\s+income': 'Other income and expense, net',
            r'other\s+expense': 'Other income and expense, net',
            r'gain\s+on\s+(?:sale|disposal)': 'Other income and expense, net',
            r'loss\s+on\s+(?:sale|disposal)': 'Other income and expense, net',
            r'foreign\s+exchange\s+(?:gain|loss)': 'Other income and expense, net',
            
            # Tax and net income
            r'income\s+tax': 'Income tax benefit',
            r'provision\s+for\s+income\s+taxes?': 'Income tax benefit',
            r'net\s+(?:profit|income|earnings|loss)': 'Net profit (loss)',
            r'earnings?\s+before\s+(?:income\s+)?taxes?': 'Net profit (loss)',
            r'ebt': 'Net profit (loss)'
        }
        
        for pattern, template_item in is_rules.items():
            if re.search(pattern, desc_lower):
                return template_item, 0.9
        
        return None, 0.0

    def apply_rule_based_mapping_cfs(self, description: str) -> tuple[str, float]:
        """Apply rule-based mapping for CFS items"""
        import re
        desc_lower = description.lower()
        
        # CFS-specific rule-based mapping
        cfs_rules = {
            r'net profit': 'Net profit (loss)',
            r'net income': 'Net profit (loss)',
            r'depreciation': 'Depreciation',
            r'deferred.*tax': 'Deferred income taxes',
            r'impairment': 'Impairment and other losses',
            r'changes.*operating': 'Changes in operating assets and liabilities',
            r'net cash.*operating': 'Net cash provided by (used in) operating activities',
            r'purchases.*equipment': 'Purchases of property and equipment',
            r'proceeds.*sale': 'Proceeds from sale of assets',
            r'net cash.*investing': 'Net cash used in investing activities',
            r'proceeds.*issuance': 'Proceeds from issuance',
            r'principal payments': 'Principal payments',
            r'net cash.*financing': 'Net cash provided by (used in) financing activities'
        }
        
        for pattern, template_item in cfs_rules.items():
            if re.search(pattern, desc_lower):
                return template_item, 0.9
        
        return None, 0.0

    def get_semantic_match_is(self, description: str, template_items: list, section: str) -> tuple[str, float]:
        """Get semantic match for IS items"""
        best_match = None
        best_score = 0.0
        
        for template_item in template_items:
            score = self.get_similarity(description, template_item)
            if score > best_score:
                best_score = score
                best_match = template_item
        
        return best_match, best_score

    def get_semantic_match_cfs(self, description: str, template_items: list, section: str) -> tuple[str, float]:
        """Get semantic match for CFS items"""
        best_match = None
        best_score = 0.0
        
        for template_item in template_items:
            score = self.get_similarity(description, template_item)
            if score > best_score:
                best_score = score
                best_match = template_item
        
        return best_match, best_score

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
        
    excel_files = list(output_dir.glob("*.xlsx"))
    if not excel_files:
        print("No Excel files found in output directory")
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

