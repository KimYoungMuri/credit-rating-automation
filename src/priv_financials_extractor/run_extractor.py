from pathlib import Path
from final_extractor import TextExtractor
from final_template_mapper import TemplateMatcher
from final_find_fs import FinancialStatementFinder
import glob
import os
import pandas as pd
import sys
from datetime import datetime
import traceback
import shutil
from openpyxl import load_workbook
import torch
from transformers import BertTokenizer, BertForSequenceClassification
from typing import List, Dict, Tuple, Optional
import numpy as np
from tenacity import retry, stop_after_attempt, wait_random_exponential
from sentence_transformers import SentenceTransformer
from collections import defaultdict

class TeeOutput:
    def __init__(self, filename):
        self.terminal = sys.stdout
        self.logfile = open(filename, 'w', encoding='utf-8')
        
    def write(self, message):
        self.terminal.write(message)
        self.logfile.write(message)
        self.logfile.flush()
        
    def flush(self):
        self.terminal.flush()
        self.logfile.flush()

def print_flush(*args, **kwargs):
    """Print function that forces flushing"""
    kwargs['flush'] = True
    print(*args, **kwargs)

def debug_print(message, file=None):
    print(message)
    if file:
        print(message, file=file)
        file.flush()

def verify_all_items_used(extracted_data, flagged_extracted):
    """
    Verify that all extracted line items with numerical values have been used.
    Returns a list of unused items with numerical values.
    """
    unused_items = []
    
    for year in ['2024', '2023']:
        for stmt_type, items in extracted_data.items():
            # Skip if no items for this statement type
            if year not in items:
                continue
                
            for item, value in items[year].items():
                # Skip items without numerical values
                if value is None or value == '' or (isinstance(value, str) and not any(c.isdigit() for c in value)):
                    continue
                    
                # Check if item was used
                if item not in flagged_extracted.get(stmt_type, set()):
                    unused_items.append({
                        'statement': stmt_type,
                        'year': year,
                        'item': item,
                        'value': value
                    })
    
    return unused_items

class EmbeddingMatcher:
    def __init__(self, model_name="all-MiniLM-L6-v2"):
        """Initialize with Sentence Transformers model"""
        try:
            self.model = SentenceTransformer(model_name)
            self.use_sbert = True
        except Exception as e:
            print(f"Warning: Could not load Sentence Transformers model: {str(e)}")
            print("Falling back to FinBERT")
            self.use_sbert = False
            model_name = "ProsusAI/finbert"
            self.tokenizer = BertTokenizer.from_pretrained(model_name)
            self.model = BertForSequenceClassification.from_pretrained(model_name)
            self.model.eval()

    def get_embedding(self, text: str) -> np.ndarray:
        """Get embeddings using Sentence Transformers"""
        if self.use_sbert:
            return self.model.encode(text, convert_to_numpy=True)
        else:
            return self._get_finbert_embedding(text)

    def _get_finbert_embedding(self, text: str) -> np.ndarray:
        """Fallback to FinBERT embeddings"""
        try:
            inputs = self.tokenizer(text, return_tensors="pt", padding=True, truncation=True, max_length=128)
            with torch.no_grad():
                outputs = self.model(**inputs)
                # Use the last hidden state as embedding
                embeddings = outputs.hidden_states[-1].mean(dim=1)
                return embeddings[0].numpy()
        except Exception as e:
            print(f"FinBERT error: {str(e)}")
            return np.zeros(768)  # Return zero vector as fallback

    def get_similarity(self, text1: str, text2: str) -> float:
        """Get cosine similarity between two texts"""
        if self.use_sbert:
            # SentenceTransformer's encode handles batching efficiently
            embeddings = self.model.encode([text1, text2], convert_to_numpy=True)
            return self._cosine_similarity(embeddings[0], embeddings[1])
        else:
            emb1 = self.get_embedding(text1)
            emb2 = self.get_embedding(text2)
            return self._cosine_similarity(emb1, emb2)

    def _cosine_similarity(self, v1: np.ndarray, v2: np.ndarray) -> float:
        """Compute cosine similarity between two vectors"""
        dot_product = np.dot(v1, v2)
        norm1 = np.linalg.norm(v1)
        norm2 = np.linalg.norm(v2)
        return dot_product / (norm1 * norm2)

def populate_excel_template(extracted_data, template_path=None):
    """
    Populate Excel template with extracted financial data using subsection-aware, flagging, and 'Other(s)' logic.
    """
    # Initialize embedding matcher (will use OpenAI if API key available)
    matcher = EmbeddingMatcher()
    
    # Subsection keywords for contextual matching
    SUBSECTION_KEYWORDS = {
        'current_assets': [
            'current assets', 'cash', 'receivable', 'inventory', 'prepaid', 
            'short-term', 'marketable securities', 'investments'
        ],
        'non_current_assets': [
            'non-current assets', 'long-term', 'property', 'equipment', 'ppe', 
            'goodwill', 'intangible'
        ],
        'current_liabilities': [
            'current liabilities', 'accounts payable', 'accrued', 'short-term debt',
            'current portion', 'deferred revenue'
        ],
        'non_current_liabilities': [
            'non-current liabilities', 'long-term debt', 'long term', 'bonds payable',
            'deferred tax', 'lease obligations'
        ],
        'equity': [
            'equity', 'stock', 'capital', 'retained earnings', 'accumulated', 
            'treasury', 'shareholders'
        ],
        'operating_income': [
            'revenue', 'sales', 'operating expense', 'cost of goods', 'gross profit',
            'operating income'
        ],
        'pretax_income': [
            'interest', 'depreciation', 'amortization', 'non-operating', 'gain',
            'loss', 'impairment'
        ],
        'net_income': [
            'tax', 'income tax', 'net income', 'earnings per share', 'net earnings'
        ]
    }
    
    if template_path is None:
        template_path = "templates/financial_template.xlsx"
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_path = f"output_excel/populated_template_{timestamp}.xlsx"
    shutil.copy2(template_path, output_path)
    wb = load_workbook(output_path)
    bs_sheet = wb['BS']
    is_sheet = wb['IS.CF']
    year_cols = {'2024': 'F', '2023': 'E'}

    # --- Subsection definitions (row ranges and 'Other'/'Total' rows) ---
    BS_SUBSECTIONS = [
        {'name': 'current_assets', 'start': 7, 'end': 13, 'other_row': 12, 'total_row': 13},
        {'name': 'non_current_assets', 'start': 15, 'end': 20, 'other_row': 18, 'total_row': 20},
        {'name': 'total_assets', 'start': 20, 'end': 21, 'other_row': None, 'total_row': 21},  # Total Assets
        {'name': 'current_liabilities', 'start': 24, 'end': 29, 'other_row': 28, 'total_row': 29},
        {'name': 'non_current_liabilities', 'start': 31, 'end': 35, 'other_row': 33, 'total_row': 35},
        {'name': 'total_liabilities', 'start': 35, 'end': 36, 'other_row': None, 'total_row': 36},  # Total Liabilities
        {'name': 'equity', 'start': 39, 'end': 43, 'other_row': 42, 'total_row': 43},
        {'name': 'total_liabilities_and_equity', 'start': 45, 'end': 45, 'other_row': None, 'total_row': 45},  # Total Liabilities and Equity
    ]
    IS_SUBSECTIONS = [
        {'name': 'operating_income', 'start': 6, 'end': 8, 'other_row': 7, 'total_row': 8},  # Revenue to Operating Income
        {'name': 'pretax_income', 'start': 10, 'end': 17, 'other_row': 15, 'total_row': 17},  # Non-operating items to Income Before Taxes
        {'name': 'net_income', 'start': 17, 'end': 20, 'other_row': 19, 'total_row': 20},  # Tax expense to Net Income
    ]
    CF_SUBSECTIONS = [
        {'name': 'operating', 'start': 23, 'end': 27, 'other_row': None, 'total_row': 27},
        {'name': 'investing', 'start': 29, 'end': 33, 'other_row': 32, 'total_row': 33},
        {'name': 'financing', 'start': 35, 'end': 41, 'other_row': 40, 'total_row': 41},
        {'name': 'cash_flow_summary', 'start': 43, 'end': 45, 'other_row': None, 'total_row': 45},
    ]

    def normalize(s):
        """Normalize text for comparison"""
        return s.lower().replace('-', ' ').replace('_', ' ').replace('—', ' ').replace(':', '').strip() if s else ''

    def get_semantic_similarity(text1, text2):
        """Get semantic similarity score using embeddings"""
        return matcher.get_similarity(text1, text2)

    def find_subsection(item_desc: str, context_items: list, statement_type: str) -> list:
        """
        Find potential subsections for an item using embeddings and contextual information.
        
        Args:
            item_desc: The description of the line item
            context_items: List of items before and after the current item
            statement_type: Type of statement ('balance_sheet', 'income_statement', 'cash_flow')
        
        Returns:
            List of potential subsection names, sorted by confidence
        """
        potential_subsections = {}
        item_desc_lower = item_desc.lower()
        
        # Get relevant subsections based on statement type
        if statement_type == 'balance_sheet':
            relevant_sections = ['current_assets', 'non_current_assets', 'current_liabilities', 
                               'non_current_liabilities', 'equity']
        elif statement_type == 'income_statement':
            relevant_sections = ['operating_income', 'pretax_income', 'net_income']
        else:
            return []  # Cash flow handled separately
            
        # 1. Direct keyword matching
        for section in relevant_sections:
            score = 0
            for keyword in SUBSECTION_KEYWORDS[section]:
                if keyword in item_desc_lower:
                    score += 0.3
                # Check context items
                for ctx_item in context_items:
                    if keyword in ctx_item.lower():
                        score += 0.1
            if score > 0:
                potential_subsections[section] = max(potential_subsections.get(section, 0), score)
                
        # 2. Semantic matching using embeddings
        for section in relevant_sections:
            # Create section description from keywords
            section_desc = ' '.join(SUBSECTION_KEYWORDS[section])
            semantic_score = get_semantic_similarity(item_desc, section_desc)
            
            # Also check context items
            context_score = 0
            for ctx_item in context_items:
                context_score += get_semantic_similarity(ctx_item, section_desc) * 0.2
            
            total_score = semantic_score + min(context_score, 0.4)  # Cap context influence
            if total_score >= 0.4:  # Minimum threshold
                potential_subsections[section] = max(potential_subsections.get(section, 0), total_score)
        
        # 3. Special cases and heuristics
        if statement_type == 'balance_sheet':
            # Check for total indicators
            if 'total' in item_desc_lower:
                if any(word in item_desc_lower for word in ['current', 'short']):
                    potential_subsections['current_assets'] = max(potential_subsections.get('current_assets', 0), 0.9)
                    potential_subsections['current_liabilities'] = max(potential_subsections.get('current_liabilities', 0), 0.9)
                elif any(word in item_desc_lower for word in ['non', 'long']):
                    potential_subsections['non_current_assets'] = max(potential_subsections.get('non_current_assets', 0), 0.9)
                    potential_subsections['non_current_liabilities'] = max(potential_subsections.get('non_current_liabilities', 0), 0.9)
        
        elif statement_type == 'income_statement':
            # Check for operating vs non-operating indicators
            if any(word in item_desc_lower for word in ['operating', 'revenue', 'sales']):
                potential_subsections['operating_income'] = max(potential_subsections.get('operating_income', 0), 0.8)
            elif any(word in item_desc_lower for word in ['interest', 'depreciation', 'amortization']):
                potential_subsections['pretax_income'] = max(potential_subsections.get('pretax_income', 0), 0.8)
            elif any(word in item_desc_lower for word in ['tax', 'net income']):
                potential_subsections['net_income'] = max(potential_subsections.get('net_income', 0), 0.8)
        
        # Sort subsections by confidence score
        sorted_subsections = sorted(potential_subsections.items(), key=lambda x: x[1], reverse=True)
        
        # Return only subsections with score above threshold
        return [section for section, score in sorted_subsections if score >= 0.4]

    def best_match(item, candidates, flagged, context_items=None, statement_type=None):
        """Enhanced matching using both direct text and semantic similarity with subsection awareness"""
        item_norm = normalize(item)
        best_direct = None
        best_direct_score = 0
        best_semantic = None
        best_semantic_score = 0
        
        # First, find potential subsections
        if context_items and statement_type:
            potential_subsections = find_subsection(item, context_items, statement_type)
            # Filter candidates to only those in potential subsections
            if potential_subsections:
                filtered_candidates = {}
                for row, label in candidates.items():
                    # Get the subsection that contains this row
                    subsections = BS_SUBSECTIONS if statement_type == 'balance_sheet' else \
                                IS_SUBSECTIONS if statement_type == 'income_statement' else \
                                CF_SUBSECTIONS
                    
                    for sub in subsections:
                        if row in range(sub['start'], sub['end'] + 1) and \
                           sub['name'] in potential_subsections:
                            filtered_candidates[row] = label
                            break
                if filtered_candidates:
                    candidates = filtered_candidates
        
        for row, label in candidates.items():
            if row in flagged:
                continue
                
            label_norm = normalize(label)
            
            # Direct text matching (faster, more precise)
            if item_norm == label_norm:
                return row, 1.0, "exact"
            elif item_norm in label_norm or label_norm in item_norm:
                if best_direct_score < 0.8:
                    best_direct = row
                    best_direct_score = 0.8
            else:
                # Word overlap score
                overlap = len(set(item_norm.split()) & set(label_norm.split()))
                overlap_score = overlap / max(1, len(set(item_norm.split())))
                if overlap_score > best_direct_score:
                    best_direct = row
                    best_direct_score = overlap_score
            
            # Semantic matching using embeddings
            semantic_score = get_semantic_similarity(item, label)
            if semantic_score > best_semantic_score:
                best_semantic = row
                best_semantic_score = semantic_score
        
        if best_direct_score >= 0.6 and best_direct_score >= best_semantic_score:
            return best_direct, best_direct_score, "direct"
        elif best_semantic_score >= 0.8:
            return best_semantic, best_semantic_score, "semantic"
        elif best_direct_score >= 0.6:
            return best_direct, best_direct_score, "direct"
        else:
            return None, 0, None

    def process_subsections(sheet, extracted, subsections, year):
        """Process each subsection with enhanced matching"""
        flagged_extracted = set()
        flagged_template = set()
        match_details = []
        
        # Get all extracted items for context
        all_items = list(extracted[year].keys())
        
        for sub in subsections:
            template_items = {row: sheet[f'A{row}'].value 
                            for row in range(sub['start'], sub['end']+1) 
                            if sheet[f'A{row}'].value}
            
            # First pass: direct/semantic match with context
            for i, ex_item in enumerate(all_items):
                if ex_item in flagged_extracted:
                    continue
                    
                # Get context items (2 before and 2 after)
                context_start = max(0, i - 2)
                context_end = min(len(all_items), i + 3)
                context_items = all_items[context_start:i] + all_items[i+1:context_end]
                
                value = extracted[year][ex_item]
                row, score, match_type = best_match(
                    ex_item, 
                    template_items, 
                    flagged_template,
                    context_items=context_items,
                    statement_type=sub.get('statement_type', 'balance_sheet')
                )
                
                if row:
                    sheet[f"{year_cols[year]}{row}"] = value
                    flagged_extracted.add(ex_item)
                    flagged_template.add(row)
                    match_details.append({
                        'subsection': sub['name'],
                        'extracted_item': ex_item,
                        'template_item': template_items[row],
                        'score': score,
                        'match_type': match_type
                    })
            
            # Second pass: add unmatched extracted to 'Other(s)' if possible
            if sub['other_row']:
                other_sum = 0
                for ex_item, value in extracted[year].items():
                    if ex_item not in flagged_extracted:
                        try:
                            other_sum += float(value) if value not in [None, ''] else 0
                            flagged_extracted.add(ex_item)
                            match_details.append({
                                'subsection': sub['name'],
                                'extracted_item': ex_item,
                                'template_item': 'Other',
                                'score': None,
                                'match_type': 'other'
                            })
                        except (ValueError, TypeError):
                            print(f"Warning: Could not convert value '{value}' to float")
                
                if other_sum != 0:
                    sheet[f"{year_cols[year]}{sub['other_row']}"] = other_sum
            
            # Fill total row: if not matched, sum subsection
            if sub['total_row']:
                total_row = sub['total_row']
                if sheet[f"{year_cols[year]}{total_row}"].value in [None, '', 0]:
                    total = 0
                    for row in range(sub['start'], total_row):
                        v = sheet[f"{year_cols[year]}{row}"].value
                        try:
                            total += float(v) if v not in [None, ''] else 0
                        except (ValueError, TypeError):
                            continue
                    if total != 0:
                        sheet[f"{year_cols[year]}{total_row}"] = total
        
        # Print matching details
        print(f"\nMatching details for {year}:")
        for detail in match_details:
            print(f"Section: {detail['subsection']}")
            print(f"Extracted: {detail['extracted_item']}")
            print(f"Template: {detail['template_item']}")
            print(f"Score: {detail['score']}")
            print(f"Match type: {detail['match_type']}")
            print("-" * 50)
        
        return flagged_extracted, flagged_template

    # --- Process each statement ---
    all_flagged_extracted = {}
    for year in ['2024', '2023']:
        # Balance Sheet
        if 'balance_sheet' in extracted_data:
            flagged_extracted, _ = process_subsections(bs_sheet, extracted_data['balance_sheet'], BS_SUBSECTIONS, year)
            all_flagged_extracted['balance_sheet'] = flagged_extracted
            
        # Income Statement
        if 'income_statement' in extracted_data:
            flagged_extracted, _ = process_subsections(is_sheet, extracted_data['income_statement'], IS_SUBSECTIONS, year)
            all_flagged_extracted['income_statement'] = flagged_extracted
            
        # Cash Flow
        if 'cash_flow' in extracted_data:
            flagged_extracted, _ = process_subsections(is_sheet, extracted_data['cash_flow'], CF_SUBSECTIONS, year)
            all_flagged_extracted['cash_flow'] = flagged_extracted
    
    # Verify all items were used
    unused_items = verify_all_items_used(extracted_data, all_flagged_extracted)
    if unused_items:
        print("\nWARNING: Found unused line items with numerical values:")
        for item in unused_items:
            print(f"Statement: {item['statement']}")
            print(f"Year: {item['year']}")
            print(f"Item: {item['item']}")
            print(f"Value: {item['value']}")
            print("-" * 50)
    
    wb.save(output_path)
    print(f"\nPopulated Excel template saved to: {output_path}")
    return output_path

def export_to_template(results, output_dir):
    """Export the mapped results to an Excel template"""
    # Get project root directory
    current_dir = Path(__file__).resolve().parent
    project_root = current_dir.parent.parent
    
    # Use the template from templates directory
    template_path = project_root / "templates" / "financial_template.xlsx"
    
    if not template_path.exists():
        print(f"Template not found at {template_path}")
        return None
        
    # Convert results format to match what populate_excel_template expects
    converted_data = {}
    for stmt_type, years in results.items():
        converted_data[stmt_type] = years  # The format is already correct - both use {'2024': {}, '2023': {}}
        
    # Call populate_excel_template with the converted data
    return populate_excel_template(converted_data, str(template_path))

def main(): 
    """
    Main workflow for the financial statement extraction and mapping process.
    """
    # Get project root directory
    current_dir = Path(__file__).resolve().parent
    project_root = current_dir.parent.parent
    
    # Setup input and output directories
    input_pdfs_dir = project_root / "input_pdfs"
    output_excel_dir = project_root / "output_excel"
    input_pdfs_dir.mkdir(exist_ok=True)
    output_excel_dir.mkdir(exist_ok=True)
    
    # Redirect stdout to both terminal and log file
    log_file_path = project_root / "run_extractor.log"
    sys.stdout = TeeOutput(log_file_path)
    
    print_flush("--- Starting Financial Statement Extraction Workflow ---")
    
    try:
        # --- Get PDF file from command line or user input ---
        if len(sys.argv) > 1:
            pdf_filename = sys.argv[1]
        else:
            # Find the most recently modified PDF in the input directory
            pdf_files = list(input_pdfs_dir.glob("*.pdf"))
            if not pdf_files:
                print_flush("No PDF files found in the 'input_pdfs' directory.")
                return
            latest_pdf = max(pdf_files, key=lambda p: p.stat().st_mtime)
            pdf_filename = latest_pdf.name
            print_flush(f"No PDF specified, processing most recent: {pdf_filename}")

        pdf_path = input_pdfs_dir / pdf_filename
        if not pdf_path.exists():
            print_flush(f"ERROR: PDF file not found at {pdf_path}")
            return
            
        print_flush(f"\nProcessing file: {pdf_path}")
        
        # --- Step 1: Find financial statement pages ---
        print_flush("\n--- Step 1: Finding financial statement pages... ---")
        finder = FinancialStatementFinder()
        _, _, statement_pages = finder.extractContent(str(pdf_path))
        
        # Use the confirmation system to get user-verified pages
        confirmed_pages = finder.confirm_statement_pages(pdf_filename)
        
        if not confirmed_pages:
            print_flush("ERROR: No financial statement pages confirmed. Aborting.")
            return
            
        # Convert to the format expected by the extractor {stmt_type: [page_nums]}
        pages_to_extract = confirmed_pages
        
        # --- Step 2: Extract text from the identified pages ---
        print_flush("\n--- Step 2: Extracting text from identified pages... ---")
        extractor = TextExtractor()
        # The `extract_text` method now returns a dictionary in the correct format
        _, extracted_data = extractor.extract_text(str(pdf_path), process_numbers=True, statement_pages=pages_to_extract)

        if not extracted_data:
            print_flush("ERROR: Text extraction failed to produce data. Aborting.")
            return
        
        print_flush("\n--- Step 3: Mapping extracted data to template... ---")
        matcher = TemplateMatcher(use_llm_fallback=True)
        template_path = project_root / "templates" / "financial_template.xlsx"
        
        final_populated_path = matcher.map_to_template(extracted_data, str(template_path))
        
        print_flush(f"\n--- WORKFLOW COMPLETE ---")
        print_flush(f"Final populated template saved to: {final_populated_path}")

    except Exception as e:
        print_flush(f"\n--- An error occurred during the workflow ---")
        print_flush(f"Error: {str(e)}")
        print_flush("Traceback:")
        traceback.print_exc(file=sys.stdout)
    
    finally:
        # Restore original stdout
        if isinstance(sys.stdout, TeeOutput):
            sys.stdout.logfile.close()
            sys.stdout = sys.stdout.terminal
            
if __name__ == "__main__":
    main() 