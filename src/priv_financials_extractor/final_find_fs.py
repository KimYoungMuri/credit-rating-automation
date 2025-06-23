import pdfplumber
import logging
from pathlib import Path
import re 
import sys
from collections import defaultdict
import os
from typing import List, Optional, Tuple

class ScoringSystem:
    def __init__(self):
        self.logger = logging.getLogger(__name__)
        
    def score_toc_match(self, text: str, page_num: int, statement_type: str) -> float:
        """Score a TOC match based on various factors"""
        score = 80.0  # Base score for TOC match
        
        # Check for exact vs fuzzy match of statement name
        statement_patterns = {
            'balance_sheet': [
                (r'(?:consolidated\s+)?balance\s+sheets?', 10),
                (r'statements?\s+of\s+financial\s+position', 8),
                (r'statements?\s+of\s+financial\s+condition', 8),
                (r'statement\s+of\s+condition', 8)
            ],
            'income_statement': [
                (r'statements?\s+of\s+(?:comprehensive\s+)?income', 10),
                (r'income\s+statements?', 9),
                (r'statements?\s+of\s+(?:operations|earnings)', 8)
            ],
            'cash_flow': [
                (r'statements?\s+of\s+cash\s+flows?', 10),
                (r'cash\s+flow\s+statements?', 9)
            ]
        }
        
        # Check for exact matches first
        for pattern, points in statement_patterns.get(statement_type, []):
            if re.search(pattern, text.lower()):
                score += points
                break
                
        # Check for page number format
        if re.search(r'\d+(?:\s*[-–]\s*\d+)?', text):
            score += 5
            
        # Cap at 100
        return min(100, score)
        
    def score_direct_scan(self, text: str, page_num: int, statement_type: str, 
                         table_quality: tuple, line_items_count: int) -> float:
        """Score a direct scan match based on various factors"""
        score = 60.0  # Base score for direct scan
        
        # Header match quality (same patterns as TOC)
        header_text = '\n'.join(text.split('\n')[:5]).lower()
        statement_patterns = {
            'balance_sheet': [
                (r'(?:consolidated\s+)?balance\s+sheets?', 15),
                (r'statements?\s+of\s+financial\s+position', 12),
                (r'statements?\s+of\s+financial\s+condition', 12),
                (r'statement\s+of\s+condition', 12)
            ],
            'income_statement': [
                (r'statements?\s+of\s+(?:comprehensive\s+)?income', 15),
                (r'income\s+statements?', 13),
                (r'statements?\s+of\s+(?:operations|earnings)', 12),
                (r'(?:consolidated\s+)?statements?\s+of\s+comprehensive\s+income', 15),
            ],
            'cash_flow': [
                (r'statements?\s+of\s+cash\s+flows?', 15),
                (r'cash\s+flow\s+statements?', 13)
            ]
        }
        
        # Check header matches
        for pattern, points in statement_patterns.get(statement_type, []):
            if re.search(pattern, header_text):
                score += points
                break
                
        # Table quality (has_table, data_rows)
        has_table, data_rows = table_quality
        if has_table:
            score += min(15, data_rows / 2)  # Up to 15 points for table quality
            
        # Line items quality
        if line_items_count >= 8:
            score += 15
        elif line_items_count >= 5:
            score += 10
        elif line_items_count >= 3:
            score += 5
            
        # Check for continuation indicators
        continuation_patterns = [
            r'\(continued\)',
            r'\(cont(?:\'d|\.)\)',
            r'continued',
            r'\(concluded\)',
            r'\(continued\)\s*$',
            r'\(page\s+\d+\s+of\s+\d+\)'
        ]
        
        if any(re.search(pattern, text.lower()) for pattern in continuation_patterns):
            score += 5
            
        # Cap at 100
        return min(100, score)
        
    def combine_scores(self, toc_score: float, direct_score: float, has_toc: bool) -> float:
        """Combine TOC and direct scan scores based on whether TOC exists"""
        if has_toc:
            return (toc_score * 0.7) + (direct_score * 0.3)
        return direct_score

class FinancialStatementFinder:
    def __init__(self): 
        # Track the last detected statement type and its page
        self.last_statement_type = None
        self.last_statement_page = None
        self.logger = logging.getLogger(__name__)
        self.setup_logging()
        # Initialize sets to store pages found by each method
        self.toc_statement_pages = set()
        self.direct_statement_pages = set()
        self.scoring = ScoringSystem()
        # Store scores for each page
        self.page_scores = defaultdict(dict)
        # Store final scores for each statement type
        self.final_scores = defaultdict(dict)

    def setup_logging(self): 
        logging.basicConfig(
            level=logging.INFO, 
            format='%(asctime)s - %(levelname)s - %(message)s',
            handlers=[
                logging.StreamHandler()
            ]
        )

    def process_toc_page(self, text, page_num):
        statements = {
            'balance_sheet': {'page': None, 'page_range': None, 'score': 0},
            'income_statement': {'page': None, 'page_range': None, 'score': 0},
            'cash_flow': {'page': None, 'page_range': None, 'score': 0}
        }
        
        patterns = {
            'balance_sheet': [
                r'(?:consolidated\s+)?balance\s+sheets?.*?(\\d+)(?:\\s*(?:and|to|through|\\-|&)\\s*(\\d+))?(?!\\d)',
                r'balance\s+sheets?.*?(\d+)(?:[-–—]\s*(\d+))?',
                r'statements?\s+of\s+financial\s+position.*?(\d+)(?:[-–—]\s*(\d+))?',
                r'statements?\s+of\s+financial\s+condition.*?(\d+)(?:[-–—]\s*(\d+))?',
                r'statement\s+of\s+condition.*?(\d+)(?:[-–—]\s*(\d+))?',
            ],
            'income_statement': [
                r'statements?\s+of\s+(?:comprehensive\s+)?income.*?(\d+)(?:[-–—]\s*(\d+))?',
                r'statements?\s+of\s+(?:operations|earnings).*?(\d+)(?:[-–—]\s*(\d+))?',
                r'income\s+statements?.*?(\d+)(?:[-–—]\s*(\d+))?',
                r'statements?\s+of\s+earnings.*?(\d+)(?:[-–—]\s*(\d+))?',
                r'statements?\s+of\s+operations.*?(\d+)(?:[-–—]\s*(\d+))?',
                r'(?:consolidated\s+)?statements?\s+of\s+comprehensive\s+income.*?(\d+)(?:[-–—]\s*(\d+))?',
            ],
            'cash_flow': [
                r'statements?\s+of\s+cash\s+flows?.*?(\d+)(?:[-–—]\s*(\d+))?',
                r'cash\s+flow\s+statements?.*?(\d+)(?:[-–—]\s*(\d+))?',
            ]
        }
        
        for stmt_type in patterns:
            for pattern in patterns[stmt_type]:
                matches = re.finditer(pattern, text.lower(), re.IGNORECASE)
                for match in matches:
                    try:
                        page = int(match.group(1))
                        page_range = None
                        if match.group(2):  # If there's a second page number
                            page_range = (page, int(match.group(2)))
                        statements[stmt_type]['page'] = page
                        statements[stmt_type]['page_range'] = page_range
                        # Score the TOC match
                        statements[stmt_type]['score'] = self.scoring.score_toc_match(text, page_num, stmt_type)
                        break  # Take first match for each pattern
                    except (ValueError, IndexError, AttributeError):
                        continue
                if statements[stmt_type]['page'] is not None:
                    break  # Stop if we found a match

        return statements

    def find_page_offset(self, pdf):
        for page_num in range(min(5, len(pdf.pages))):
            page = pdf.pages[page_num]
            text = page.extract_text()
            # Look for page number at bottom
            found_num = self.extract_page_number_from_bottom(page)
            if found_num is not None:
                return page_num - (found_num - 1)  # -1 because found_num is 1-based
        return 0

    def extract_page_number_from_bottom(self, page):
        """Extract page number from the bottom of the page."""
        try:
            height = page.height
            width = page.width
            margin = 10  # Safety margin in points
            
            bottom_height = min(height * 0.1, height - margin)
            crop_box = (
                margin,  # left
                max(margin, height - bottom_height),  # top
                width - margin,  # right
                height - margin  # bottom
            )
            
            bottom_crop = page.crop(crop_box)
            text = bottom_crop.extract_text() or ""
            text = self.standardize_text(text)
            
            page_patterns = [
                r'(?:^|\s)(\d+)(?:\s|$)',  # Standalone number
                r'page\s*(\d+)',  # "Page X"
                r'pg\.\s*(\d+)',  # "Pg. X"
                r'[-‒–—]\s*(\d+)\s*[-‒–—]'  # Number between dashes
            ]
            
            for pattern in page_patterns:
                matches = re.finditer(pattern, text.lower())
                for match in matches:
                    try:
                        return int(match.group(1))
                    except ValueError:
                        continue
            return None
        except Exception as e:
            logging.error(f"Error extracting page number: {str(e)}")
        return None

    def extractContent(self, pdf_path):
        try:
            with pdfplumber.open(pdf_path) as pdf:
                ret = []
                toc_pages = []
                self.toc_statement_pages = set()  # Statements found from TOC
                self.direct_statement_pages = set()  # Statements found from direct scanning
                page_offset = self.find_page_offset(pdf)
                
                # Store potential fallback pages
                fallback_pages = []
                
                # First pass: find TOC pages and process them
                has_toc = False
                for page_num, page in enumerate(pdf.pages, 1):
                    text = page.extract_text()
                    is_toc = self.is_table_of_contents(text, page_num)
                    if is_toc:
                        has_toc = True
                        toc_pages.append(page_num)
                        statements = self.process_toc_page(text, page_num)
                        for stmt_type, info in statements.items():
                            if info['page'] is not None:
                                # Convert document page to PDF page
                                pdf_page = info['page'] + page_offset
                                self.toc_statement_pages.add(pdf_page)
                                # Store TOC score
                                self.page_scores[pdf_page][f"{stmt_type}_toc"] = info['score']
                                if info['page_range'] and info['page_range'][1] > info['page']:
                                    # Add continuation pages
                                    for p in range(info['page'] + 1, info['page_range'][1] + 1):
                                        p_pdf = p + page_offset
                                        self.toc_statement_pages.add(p_pdf)
                                        # Store continuation TOC score (slightly lower)
                                        self.page_scores[p_pdf][f"{stmt_type}_toc"] = max(60, info['score'] - 10)

                # Second pass: process each page
                for page_num, page in enumerate(pdf.pages, 1):
                    text = page.extract_text()
                    
                    # Check if this is a statement page through direct scanning
                    is_statement, direct_score = self.is_statement_page(text, page_num)
                    if is_statement:
                        if direct_score >= 60:  # High confidence matches
                            self.direct_statement_pages.add(page_num)
                            # Store direct scan score
                            self.page_scores[page_num][f"{self.last_statement_type}_direct"] = direct_score
                        else:  # Lower confidence matches stored as fallback
                            fallback_pages.append((page_num, self.last_statement_type, direct_score))
                    
                    ret.append(text)

                # Combine and evaluate final scores
                self.final_scores = defaultdict(dict)
                
                # First add high-confidence matches
                for page_num in set(self.toc_statement_pages).union(self.direct_statement_pages):
                    for stmt_type in ['balance_sheet', 'income_statement', 'cash_flow']:
                        toc_score = self.page_scores[page_num].get(f"{stmt_type}_toc", 0)
                        direct_score = self.page_scores[page_num].get(f"{stmt_type}_direct", 0)
                        if toc_score > 0 or direct_score > 0:
                            self.final_scores[stmt_type][page_num] = self.scoring.combine_scores(
                                toc_score, direct_score, has_toc
                            )
                
                # If we didn't find all statement types, add fallback pages
                if len(self.final_scores) < 3 and fallback_pages:
                    missing_types = set(['balance_sheet', 'income_statement', 'cash_flow']) - set(self.final_scores.keys())
                    for page_num, stmt_type, score in fallback_pages:
                        if stmt_type in missing_types and page_num not in self.final_scores.get(stmt_type, {}):
                            self.final_scores[stmt_type][page_num] = score

                # Print results
                print("\nFinancial Statement Detection Results:")
                print("=" * 50)
                
                # Check if we have any high confidence matches
                high_confidence_matches = False
                for stmt_type, pages in self.final_scores.items():
                    high_conf_pages = {page: score for page, score in pages.items() if score >= 50}
                    if high_conf_pages:
                        high_confidence_matches = True
                        print(f"\n{stmt_type.replace('_', ' ').title()}:")
                        for page, score in sorted(high_conf_pages.items(), key=lambda x: (-x[1], x[0])):
                            confidence = "High" if score >= 80 else "Medium" if score >= 60 else "Low"
                            print(f"  Page {page}: {score:.1f}% confidence ({confidence})")
                
                # If no high confidence matches, show top 3 pages as fallback suggestions
                if not high_confidence_matches:
                    print("\nNo high confidence matches found.")
                    print("\nFallback Suggestions:")
                    
                    # Collect all pages and scores
                    all_pages = []
                    for stmt_type, pages in self.final_scores.items():
                        for page, score in pages.items():
                            all_pages.append((page, score, stmt_type))
                    
                    # Sort by score and take top 3
                    top_pages = sorted(all_pages, key=lambda x: (-x[1], x[0]))[:3]
                    
                    if top_pages:
                        for page, score, stmt_type in top_pages:
                            print(f"  Page {page}: {score:.1f}% confidence (possible {stmt_type.replace('_', ' ')})")
                    else:
                        print("  No potential financial statement pages found.")
                
                print("\n" + "=" * 50)
                
                return ret, toc_pages, set(self.toc_statement_pages).union(self.direct_statement_pages)

        except Exception as e:
            logging.error(f"Error processing PDF: {str(e)}")
            import traceback
            traceback.print_exc()  # Print full traceback for debugging
            return None, None, None

    def standardize_text(self, text):
        if not isinstance(text, str): 
            return ""
        
        text = text.lower()
        
        lines = []
        for line in text.split('\n'):
            if not re.match(r'^\s*[a-z]\s*$', line, re.IGNORECASE):
                line = ' '.join(line.split())
                line = re.sub(r'(?<=\b\w)\s+(?=\w\b)', '', line)
                if line.strip():
                    lines.append(line)
        
        return '\n'.join(lines)

    def is_table_of_contents(self, text, page_num):
        """Determine if a page is likely a table of contents"""
        text = self.standardize_text(text)
        
        explicit_toc = any(re.search(pattern, text, re.IGNORECASE | re.MULTILINE) 
                         for pattern in [
                             r'^\s*table\s+of\s+contents\s*$',
                             r'^\s*contents\s*$'
                         ])
        
        if not explicit_toc: return False
            
        toc_indicators = [
            (r'page(?:\s+numbers?)?(?:\s*\d+\s*)+', 5), 
            (r'(?:\d+\.){2,}', 4),                      
            (r'^\s*\d+\s*\.\s+\w+.*?\.{3,}.*?\d+\s*$', 4), 
            (r'(?:consolidated\s+)?financial\s+statements?', 3),
            (r'notes\s+to.*statements', 3),
            (r'independent\s+auditors?\s*\'?\s*report', 3)
        ]
        
        # Strong indicators that this is NOT a TOC
        non_toc_indicators = [
            # Balance sheet specific
            (r'total\s+current\s+assets.*\d', 8),
            (r'total\s+assets.*\d', 8),
            (r'total\s+liabilities.*\d', 8),
            (r'total\s+equity.*\d', 8),
            
            # Income statement specific
            (r'total\s+revenue.*\d', 8),
            (r'operating\s+income.*\d', 8),
            (r'net\s+income.*\d', 8),
            
            # Cash flow specific
            (r'net\s+cash.*\d', 8),
            (r'cash\s+flows?.*\d', 8)
        ]
        
        # Check for non-TOC indicators
        non_toc_score = 0
        for pattern, weight in non_toc_indicators:
            if re.search(pattern, text, re.IGNORECASE | re.MULTILINE):
                non_toc_score += weight
        
        if non_toc_score >= 15:
            return False
            
        # Check for TOC structure
        toc_score = 0
        for pattern, weight in toc_indicators:
            if re.search(pattern, text, re.IGNORECASE | re.MULTILINE):
                toc_score += weight
        
        # Count numbered entries and page references
        lines = text.split('\n')
        numbered_entries = 0
        page_refs = 0
        
        for line in lines:
            line = line.strip()
            # Look for numbered entries
            if re.match(r'^\d+[\.\)]', line):
                numbered_entries += 1
            # Look for page number references at end of lines
            if re.search(r'.+\d+\s*$', line):
                page_refs += 1
        
        if numbered_entries >= 3:
            toc_score += 3
        if page_refs >= 3:
            toc_score += 3
            
        # Must have both good structure and explicit TOC header
        return toc_score >= 8

    def has_table_structure(self, text: str, page_num: int = None) -> tuple[bool, int]:
        """Determine if a page has table-like structure"""
        if not text:
            #self.logger.info(f"Page {page_num} - Rejected: No text content")
            return False, 0
            
        lines = text.split('\n')
        if len(lines) < 3:  # Reduced minimum lines from 4
            #self.logger.info(f"Page {page_num} - Rejected: Too few lines ({len(lines)})")
            return False, 0
            
        # Indicators of tabular data
        indicators = {
            'number_lines': 0,  # Lines containing numbers
            'aligned_numbers': 0,  # Lines with right-aligned numbers
            'dollar_amounts': 0,  # Lines with dollar amounts
            'total_lines': len(lines),
            'data_rows': 0  # Count of actual data rows
        }
        
        # Patterns for detection
        patterns = {
            'numbers': r'\d+(?:,\d{3})*(?:\.\d+)?',
            'dollar_amounts': r'\$\s*\d+(?:,\d{3})*(?:\.\d+)?',
            'right_aligned': r'.*\s+\d+(?:,\d{3})*(?:\.\d+)?\s*$'
        }
        
        prev_number_positions = []
        in_data_section = False
        
        for line in lines: 
            line = line.strip()
            if not line:
                continue
                
            # Check if line contains numbers or dollar amounts
            has_numbers = bool(re.search(patterns['numbers'], line))
            has_dollars = bool(re.search(patterns['dollar_amounts'], line))
            
            if has_numbers or has_dollars:
                indicators['number_lines'] += 1
                if in_data_section:
                    indicators['data_rows'] += 1
                else:
                    # First number line indicates start of data section
                    in_data_section = True
                    indicators['data_rows'] += 1
            
            if re.search(patterns['right_aligned'], line):
                indicators['aligned_numbers'] += 1
                
            number_positions = [m.start() for m in re.finditer(patterns['numbers'], line)]
            
            if prev_number_positions and number_positions:
                for pos in number_positions:
                    if any(abs(pos - prev_pos) <= 3 for prev_pos in prev_number_positions):  # Increased tolerance
                        indicators['aligned_numbers'] += 1
                        
            prev_number_positions = number_positions
            
        num_lines_ratio = indicators['number_lines'] / indicators['total_lines']
        dollar_ratio = indicators['dollar_amounts'] / indicators['total_lines']
        alignment_ratio = indicators['aligned_numbers'] / indicators['total_lines']
        '''
        if page_num:
            self.logger.debug(f"Page {page_num} table indicators:")
            self.logger.debug(f"Number lines ratio: {num_lines_ratio:.2f}")
            self.logger.debug(f"Dollar amount ratio: {dollar_ratio:.2f}")
            self.logger.debug(f"Alignment ratio: {alignment_ratio:.2f}")
            self.logger.debug(f"Data rows: {indicators['data_rows']}")      
        '''
        # More lenient criteria
        has_table = (
            (num_lines_ratio > 0.2 and alignment_ratio > 0.1) or   # Reduced thresholds
            (dollar_ratio > 0.1 and alignment_ratio > 0.1) or      # Reduced thresholds
            (num_lines_ratio > 0.3 and dollar_ratio > 0.1) or      # Reduced thresholds
            indicators['data_rows'] >= 8                           # Added absolute row count criteria
        )
        '''
        if not has_table:
            self.logger.info(f"Page {page_num} - Rejected: No table structure (numbers: {num_lines_ratio:.2f}, dollars: {dollar_ratio:.2f}, alignment: {alignment_ratio:.2f})")
        else:
            self.logger.info(f"Page {page_num} - Table structure found: {indicators['data_rows']} data rows")
            
        '''

        return has_table, indicators['data_rows']

    def has_typical_line_items(self, text: str, statement_type: str) -> bool:
        """Check if the text contains typical line items for the given statement type."""
        text = text.lower()
        
        # First identify the table region
        lines = text.split('\n')
        table_start = 0
        table_end = len(lines)
        
        # Find start of table - first line with numbers or dollar amounts
        for i, line in enumerate(lines):
            if re.search(r'\d+(?:,\d{3})*(?:\.\d+)?|\$\s*\d+(?:,\d{3})*(?:\.\d+)?', line):
                table_start = max(0, i - 1)  # Include one line before first number
                break
                
        # Find end of table - last line with numbers or dollar amounts
        for i in range(len(lines) - 1, -1, -1):
            if re.search(r'\d+(?:,\d{3})*(?:\.\d+)?|\$\s*\d+(?:,\d{3})*(?:\.\d+)?', lines[i]):
                table_end = min(len(lines), i + 2)  # Include one line after last number
                break
                
        # Get just the table region
        table_text = '\n'.join(lines[table_start:table_end])
        
        # Define required patterns for each statement type
        statement_requirements = {
            'balance_sheet': {
                'patterns': [
                    # Asset items
                    r'(?:total\s+)?assets?',
                    r'current\s+assets?',
                    r'cash(?:\s+and\s+(?:cash\s+)?equivalents?)?',
                    r'accounts?\s+receivable',
                    r'inventor(?:y|ies)',
                    r'prepaid\s+(?:expenses?|assets?)',
                    r'property(?:\s+and\s+equipment)?',
                    r'(?:other\s+)?(?:non-?current\s+)?assets?',
                    
                    # Liability items
                    r'(?:total\s+)?liabilit(?:y|ies)',
                    r'current\s+liabilit(?:y|ies)',
                    r'accounts?\s+payable',
                    r'(?:short|long)[-\s]term\s+(?:debt|borrowings?|loans?)',
                    r'accrued\s+(?:expenses?|liabilities)',
                    r'notes?\s+payable',
                    
                    # Equity items
                    r'(?:total\s+)?(?:shareholders?|stockholders?\s+)?equity',
                    r'(?:paid[- ]in|additional)\s+capital',
                    r'retained\s+earnings?',
                    r'common\s+stock',
                    r'preferred\s+stock',
                    r'treasury\s+stock',
                    r'accumulated\s+(?:other\s+)?comprehensive\s+(?:income|loss)'
                ],
                'min_matches': 4
            },
            'income_statement': {
                'patterns': [
                    # Revenue items
                    r'(?:net\s+)?(?:revenue|sales)',
                    r'total\s+revenue',
                    r'gross\s+(?:revenue|sales)',
                    r'operating\s+revenue',
                    
                    # Cost/Expense items
                    r'cost\s+of\s+(?:goods\s+sold|revenue|sales)',
                    r'gross\s+(?:profit|margin)',
                    r'operating\s+expenses?',
                    r'selling(?:\s+and|\s*,\s*)\s*general(?:\s+and)?\s+administrative',
                    r'depreciation(?:\s+and\s+amortization)?',
                    r'research\s+and\s+development',
                    
                    # Income/Profit items
                    r'(?:net\s+)?income(?:\s+\(?loss\)?)?(?:\s+before|attributable\s+to)?',
                    r'net\s+earnings?',
                    r'(?:operating|gross)\s+(?:income|profit|loss)',
                    r'income\s+from\s+operations',
                    r'income\s+before\s+(?:income\s+)?tax(?:es)?',
                    r'(?:provision|benefit)\s+for\s+income\s+tax(?:es)?',
                    r'earnings?\s+per\s+share',
                    r'ebitda'
                ],
                'min_matches': 4
            },
            'cash_flow': {
                'patterns': [
                    # Operating activities
                    r'operating\s+activities?',
                    r'cash\s+(?:flows?\s+)?(?:from|provided\s+by|used\s+in)\s+operating',
                    r'net\s+(?:cash\s+)?(?:provided|used)\s+by\s+operating',
                    r'depreciation(?:\s+and\s+amortization)?',
                    r'changes?\s+in\s+(?:operating|working)\s+capital',
                    r'accounts?\s+(?:receivable|payable)',
                    
                    # Investing activities
                    r'investing\s+activities?',
                    r'cash\s+(?:flows?\s+)?(?:from|provided\s+by|used\s+in)\s+investing',
                    r'capital\s+expenditures?',
                    r'(?:purchase|sale)\s+of\s+(?:property|equipment|investments?)',
                    r'acquisitions?(?:\s+of\s+business(?:es)?)?',
                    
                    # Financing activities
                    r'financing\s+activities?',
                    r'cash\s+(?:flows?\s+)?(?:from|provided\s+by|used\s+in)\s+financing',
                    r'(?:proceeds|payments?)\s+from\s+(?:debt|borrowings?)',
                    r'dividends?\s+paid',
                    r'(?:issuance|repurchase)\s+of\s+(?:common\s+)?stock',
                    r'net\s+(?:increase|decrease)\s+in\s+cash'
                ],
                'min_matches': 3
            }
        }
        
        if statement_type not in statement_requirements:
            return False
            
        # Count how many patterns match and log them
        matches = []
        for pattern in statement_requirements[statement_type]['patterns']:
            if re.search(pattern, table_text):
                matches.append(pattern)
                
        #self.logger.info(f"Line item check for {statement_type} - Found matches: {matches}")
        return len(matches) >= statement_requirements[statement_type]['min_matches']

    def is_continuation_page(self, text: str, page_num: int) -> tuple[bool, str]:
        """Check if this page is a continuation of the previous statement."""
        text = text.lower()
        
        # Common continuation indicators
        continuation_patterns = [
            r'\(continued\)',
            r'\(cont(?:\'d|\.)\)',
            r'continued',
            r'\(concluded\)',
            r'conclusion',
            r'\(continued\)\s*$',  # At end of line
            r'\(concluded\)\s*$',   # At end of line
            r'continued\s+from\s+(?:previous|preceding)\s+page',
            r'continued\s+on\s+next\s+page',
            r'\(continued\s+on\s+page\s+\d+\)',
            r'\(see\s+next\s+page\)',
            r'\(page\s+\d+\s+of\s+\d+\)',
            r'-\s*\d+\s*-'  # Page number format like "- 5 -"
        ]
        
        # Statement title patterns
        statement_patterns = {
            'balance_sheet': [
                r'(?:consolidated\s+)?balance\s+sheets?',
                r'(?:consolidated\s+)?statements?\s+of\s+financial\s+position',
            ],
            'income_statement': [
                r'(?:consolidated\s+)?statements?\s+of\s+(?:comprehensive\s+)?income',
                r'(?:consolidated\s+)?income\s+statements?',
                r'(?:consolidated\s+)?statements?\s+of\s+comprehensive\s+income',
            ],
            'cash_flow': [
                r'(?:consolidated\s+)?statements?\s+of\s+cash\s+flows?',
                r'(?:consolidated\s+)?cash\s+flow\s+statements?',
            ]
        }
        
        # First check if this page has a continuation marker
        is_continuation = any(re.search(pattern, text) for pattern in continuation_patterns)
        
        # If we have a previous statement and this is the next page, treat it as a potential continuation
        if (hasattr(self, 'last_statement_type') and self.last_statement_type and 
            hasattr(self, 'last_statement_page') and self.last_statement_page is not None):
            if abs(page_num - self.last_statement_page) <= 2:  # Allow up to 2 pages gap
                is_continuation = True
        
        # Get first few lines for header check
        header_text = '\n'.join(text.split('\n')[:5]).lower()
        
        # Check for main statement headers first
        for stmt_type, patterns in statement_patterns.items():
            for pattern in patterns:
                if re.search(pattern, header_text):
                    # If it has a clear statement header and no continuation marker,
                    # it's likely a main statement page, not a continuation
                    if not is_continuation:
                        #self.logger.info(f"Page {page_num} - Found main {stmt_type} header")
                        return False, stmt_type
                    else:
                        #self.logger.info(f"Page {page_num} - Found {stmt_type} continuation header")
                        return True, stmt_type
                        
        if is_continuation and hasattr(self, 'last_statement_type') and self.last_statement_type:
            #self.logger.info(f"Page {page_num} - Found continuation of {self.last_statement_type}")
            return True, self.last_statement_type
            
        return False, None

    def has_long_paragraph(self, text):
        """Check if the text contains a long paragraph (likely notes or disclosures)."""
        if not text:
            return False
            
        # Split into lines and normalize whitespace
        lines = [line.strip() for line in text.split('\n') if line.strip()]
        
        # Identify non-table lines (no strong alignment patterns, not mostly numbers)
        non_table_lines = []
        current_paragraph = []
        
        for line in lines:
            # Skip table-like lines:
            # - Lines with heavy number content (>50% of characters are numbers) - increased from 30%
            # - Lines with typical table alignments (numbers at end, dollar signs)
            # - Very short lines (likely headers or single cells)
            # - Lines that are dates or years
            num_count = sum(c.isdigit() for c in line)
            if (num_count / len(line) > 0.5 or                # Heavy number content - increased threshold
                re.search(r'^\s*\d|[\d,]+\s*$', line) or     # Table indicators (removed $ check)
                len(line.split()) < 3 or                      # Very short lines
                re.match(r'^\s*\d{4}\s*$', line) or          # Years
                re.match(r'^\s*\(?\d{1,2}/\d{1,2}/\d{2,4}\)?', line)):  # Dates
                # If we were building a paragraph, complete it
                if current_paragraph:
                    non_table_lines.append(' '.join(current_paragraph))
                    current_paragraph = []
                continue
                
            # Add line to current paragraph if it doesn't end with period
            # or starts with lowercase (continuation)
            if current_paragraph and (not current_paragraph[-1].endswith('.') or 
                                    line[0].islower()):
                current_paragraph.append(line)
            else:
                # If we were building a paragraph, complete it
                if current_paragraph:
                    non_table_lines.append(' '.join(current_paragraph))
                current_paragraph = [line]
        
        # Add final paragraph if exists
        if current_paragraph:
            non_table_lines.append(' '.join(current_paragraph))
        
        # Check if any paragraph is long enough - increased threshold significantly
        for paragraph in non_table_lines:
            # Only reject if paragraph is very long and contains specific disclosure keywords
            if len(paragraph.split()) > 50 and any(keyword in paragraph.lower() for keyword in [
                'note', 'disclosure', 'accounting policies', 'significant', 'management'
            ]):
                return True
                
        return False

    def is_statement_page(self, text, page_num):
        text_lower = text.lower()
        
        # NEW: Check for negative keywords indicating it's NOT a statement page
        negative_keywords = ["independent auditors' report", "auditors' responsibilities", "report of independent"]
        if any(keyword in text_lower for keyword in negative_keywords):
            return False, 0
        # END NEW

        # Check for statement keywords
        scores = {}
        if not text:
            #self.logger.info(f"Page {page_num} - Rejected: Empty text")
            return False, 0
            
        # Check for long paragraphs early - likely notes or disclosures
        if self.has_long_paragraph(text):
            #self.logger.info(f"Page {page_num} - Rejected: Contains long paragraph (likely notes/disclosures)")
            return False, 0
            
        # Step 1: First check if the page has a proper table structure
        table_quality = self.has_table_structure(text, page_num)
        has_table, data_rows = table_quality
        if not has_table:
            return False, 0
            
        if data_rows < 4:  # Minimum rows for any statement page
            #self.logger.info(f"Page {page_num} - Rejected: Insufficient rows ({data_rows})")
            return False, 0
            
        # Step 2: Look for statement headers
        statement_patterns = {
            'balance_sheet': [
                r'(?:consolidated\s+)?balance\s+sheets?',
                r'statements?\s+of\s+financial\s+position',
            ],
            'income_statement': [
                r'statements?\s+of\s+(?:comprehensive\s+)?income',
                r'statements?\s+of\s+(?:operations|earnings)',
                r'income\s+statements?',
                r'(?:consolidated\s+)?statements?\s+of\s+comprehensive\s+income',
            ],
            'cash_flow': [
                r'statements?\s+of\s+cash\s+flows?',
                r'cash\s+flow\s+statements?',
            ]
        }
            
        # Get first few lines for header check
        header_text = '\n'.join(text.split('\n')[:5]).lower()
        
        # Try to identify statement type from header
        identified_type = None
        for stmt_type, patterns in statement_patterns.items():
            for pattern in patterns:
                if re.search(pattern, header_text):
                    identified_type = stmt_type
                    #self.logger.info(f"Page {page_num} - Found {stmt_type} header")
                    break
            if identified_type:
                break

        if hasattr(self, 'last_statement_type') and self.last_statement_type:
            # Check if it's a continuation based on content, not just page number
            if self.has_typical_line_items(text, self.last_statement_type):
                self.last_statement_page = page_num  # Update last page
                #self.logger.info(f"Page {page_num} - Accepted: Continuation of {self.last_statement_type} (matching items)")
                # Score continuation page
                line_items = sum(1 for _ in re.finditer(r'total\s+|net\s+|gross\s+', text.lower()))
                score = self.scoring.score_direct_scan(text, page_num, self.last_statement_type, 
                                                     table_quality, line_items)
                return True, score
                
        # Step 4: For main statement pages, check line items
        if identified_type:
            line_items = sum(1 for _ in re.finditer(r'total\s+|net\s+|gross\s+', text.lower()))
            has_items = self.has_typical_line_items(text, identified_type)
            if has_items:
                self.last_statement_type = identified_type
                self.last_statement_page = page_num
                #self.logger.info(f"Page {page_num} - Accepted: Main {identified_type} with header and items")
                # Score main statement page
                score = self.scoring.score_direct_scan(text, page_num, identified_type, 
                                                     table_quality, line_items)
                return True, score
            
        # Step 5: If we haven't found a match but have good table structure, 
        # store it as a potential fallback with a lower score
        if data_rows >= 8:  # Good number of rows
            # Try to identify type from content
            best_type = None
            best_matches = 0
            for stmt_type in ['balance_sheet', 'income_statement', 'cash_flow']:
                if self.has_typical_line_items(text, stmt_type):
                    matches = sum(1 for _ in re.finditer(r'total\s+|net\s+|gross\s+', text.lower()))
                    if matches > best_matches:
                        best_matches = matches
                        best_type = stmt_type
            
            if best_type:
                #self.logger.info(f"Page {page_num} - Found potential {best_type} (no header but good structure)")
                # Lower base score for no-header matches
                score = 40 + min(20, data_rows / 2)  # Up to 20 points for table quality
                if best_matches >= 3:
                    score += min(20, best_matches * 3)  # Up to 20 points for line items
                return True, score
            
        #self.logger.info(f"Page {page_num} - Rejected: No statement header or valid continuation")
        return False, 0
    
    def get_statement_pages(self) -> dict:
        """
        Returns a dictionary containing the page numbers and highest confidence score 
        for each type of financial statement. Only includes statements with a 
        confidence score >= 50%.
        
        Returns:
            dict: A dictionary where keys are statement types and values are another
                  dictionary with 'pages' and 'confidence'.
                  e.g., {'balance_sheet': {'pages': [5], 'confidence': 0.92}}
        """
        statement_pages = {}
        for stmt_type, pages in self.final_scores.items():
            # Get pages with confidence >= 50%
            high_conf_pages = {page: score for page, score in pages.items() if score >= 50}
            if high_conf_pages:
                # Find the page with the highest score for this statement type
                best_page = max(high_conf_pages, key=high_conf_pages.get)
                confidence = high_conf_pages[best_page] / 100.0 # Normalize to 0-1
                
                # For now, we assume statements can span multiple pages if they are contiguous
                # and of the same type. This logic can be enhanced.
                sorted_pages = sorted(high_conf_pages.keys())
                
                statement_pages[stmt_type] = {
                    'pages': sorted_pages,
                    'confidence': confidence
                }
        return statement_pages
    
    def confirm_statement_pages(self, filename: str) -> dict:
        """
        Display detected financial statement pages and allow user confirmation/modification.
        
        Returns:
            dict: Confirmed statement pages in the format expected by the pipeline
                  e.g., {'balance_sheet': [5, 6], 'income_statement': [7], 'cash_flow': [8]}
        """
        print("\n" + "="*60)
        print("FINANCIAL STATEMENT PAGE CONFIRMATION")
        print("="*60)
        print(f"File: {filename}")
        
        # Get the detected pages
        detected_pages = self.get_statement_pages()
        
        # Display current detections for all statements first
        print("\nDetected Financial Statement Pages:")
        print("-" * 40)
        
        statement_names = {
            'balance_sheet': 'Balance Sheet',
            'income_statement': 'Income Statement', 
            'cash_flow': 'Cash Flow Statement'
        }
        
        # Show all detected pages first
        all_detected = {}
        for stmt_type in ['balance_sheet', 'income_statement', 'cash_flow']:
            stmt_name = statement_names[stmt_type]
            
            if stmt_type in detected_pages and detected_pages[stmt_type]['pages']:
                pages = detected_pages[stmt_type]['pages']
                confidence = detected_pages[stmt_type]['confidence'] * 100
                print(f"\n{stmt_name}:")
                print(f"  Detected Pages: {pages}")
                print(f"  Confidence: {confidence:.1f}%")
                
                # Show all pages with scores for this statement type
                if stmt_type in self.final_scores:
                    print("  All detected pages with scores:")
                    sorted_pages = sorted(self.final_scores[stmt_type].items(), 
                                        key=lambda x: (-x[1], x[0]))
                    for page, score in sorted_pages[:5]:  # Show top 5
                        conf_level = "High" if score >= 80 else "Medium" if score >= 60 else "Low"
                        print(f"    Page {page}: {score:.1f}% ({conf_level})")
                
                all_detected[stmt_type] = pages
            else:
                print(f"\n{stmt_name}: Not detected")
                all_detected[stmt_type] = []
        
        # Ask one y/n question for all statements
        print("\n" + "="*60)
        print("CONFIRMATION")
        print("="*60)
        
        while True:
            response = input("Confirm all detected pages above? (y/n): ").strip().lower()
            
            if response in ['y', 'yes']:
                # Use all detected pages
                confirmed_pages = {stmt_type: pages for stmt_type, pages in all_detected.items() if pages}
                print("✓ Confirmed all detected pages")
                break
            elif response in ['n', 'no']:
                # Ask for custom page numbers for each statement
                print("\nEntering custom page numbers...")
                confirmed_pages = {}
                
                for stmt_type in ['balance_sheet', 'income_statement', 'cash_flow']:
                    stmt_name = statement_names[stmt_type]
                    
                    while True:
                        custom_input = input(f"Enter page numbers for {stmt_name} (comma-separated, or press Enter to skip): ").strip()
                        
                        if not custom_input:  # Skip this statement
                            print(f"✗ Skipping {stmt_name}")
                            break
                        
                        try:
                            custom_pages = [int(p.strip()) for p in custom_input.split(',') if p.strip()]
                            if custom_pages:
                                confirmed_pages[stmt_type] = sorted(custom_pages)
                                print(f"✓ Custom {stmt_name}: {confirmed_pages[stmt_type]}")
                                break
                            else:
                                print("No valid page numbers entered. Please try again or press Enter to skip.")
                        except ValueError:
                            print("Invalid input. Please enter comma-separated numbers (e.g., 5,6,7) or press Enter to skip.")
                
                break
            else:
                print("Please enter 'y' (yes) or 'n' (no)")
        
        print("\n" + "="*60)
        print("CONFIRMATION SUMMARY:")
        print("="*60)
        for stmt_type, pages in confirmed_pages.items():
            stmt_name = statement_names[stmt_type]
            print(f"{stmt_name}: {pages}")
        
        if not confirmed_pages:
            print("No pages confirmed. Extraction will proceed without page filtering.")
        
        return confirmed_pages

def main():
    if len(sys.argv) != 2:
        print("Usage: python final_find_fs.py <pdf_path>")
        return

    filename = sys.argv[1]
    # Use project root's input_pdfs directory
    project_root = Path(__file__).resolve().parent.parent.parent
    input_dir = project_root / "input_pdfs"
    input_dir.mkdir(exist_ok=True)
    path = str((input_dir / filename).resolve())

    if not Path(path).exists():
        print(f"Error: PDF not found at {path}")
        return
    
    finder = FinancialStatementFinder()
    lines, toc_pages, statement_pages = finder.extractContent(path)
    
    if lines:
        print("\n" + "="*50)
        print(f"Results for {filename}:")
        print("="*50)
        print("\nTOC pages found:", toc_pages)
        print("\nStatement pages found:")
        print(f"  - From TOC: {sorted(finder.toc_statement_pages)}")
        print(f"  - From Direct Scanning: {sorted(finder.direct_statement_pages)}")
        print(f"  - Combined unique pages: {sorted(statement_pages)}")
        
        # Print statement pages with high confidence
        statement_pages = finder.get_statement_pages()
        print("\nHigh Confidence Statement Pages (>=50%):")
        for stmt_type, pages in statement_pages.items():
            print(f"  {stmt_type.replace('_', ' ').title()}: {pages}")
        
        # Use the confirmation system
        confirmed_pages = finder.confirm_statement_pages(filename)
        
        # Return the confirmed pages in the format expected by the pipeline
        return confirmed_pages
            
        print("\n" + "="*50 + "\n")
    else:
        print(f"No content extracted from {filename}")
        return {}
        
if __name__ == "__main__":
    main()