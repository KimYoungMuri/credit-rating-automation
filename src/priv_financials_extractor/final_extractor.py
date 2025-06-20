import logging
from pathlib import Path
import pdfplumber
from final_find_fs import FinancialStatementFinder
import re
from typing import List, Dict, Tuple, Optional
from collections import Counter, defaultdict
import sys
import numpy as np
import pandas as pd
from datetime import datetime
from final_template_mapper import TemplateMatcher

class TextExtractor: 
    def __init__(self):
        self.setup_logging()
        self.finder = FinancialStatementFinder()
        
        # Get project root directory (two levels up from this file)
        current_dir = Path(__file__).resolve().parent
        project_root = current_dir.parent.parent
        
        # Set up output directory relative to project root
        self.output_dir = project_root / "output_excel"
        self.output_dir.mkdir(exist_ok=True)
        
        print("Debug log file created")
        print(f"Output directory created/verified: {self.output_dir}")

    def setup_logging(self):
        """Setup logging configuration"""
        # Get project root directory for log file
        current_dir = Path(__file__).resolve().parent
        project_root = current_dir.parent.parent
        log_file = project_root / "text_extraction.log"
        
        for handler in logging.root.handlers[:]:
            logging.root.removeHandler(handler)
        formatter = logging.Formatter('%(message)s')
        file_handler = logging.FileHandler(log_file, mode='w', encoding='utf-8')
        file_handler.setFormatter(formatter)
        logging.root.addHandler(file_handler)
        logging.root.setLevel(logging.INFO)
        logging.info("Logging initialized successfully")

    def clean_and_validate_number(self, text: str) -> Optional[str]:
        """
        Cleans and validates if a string is a parsable financial number.
        Handles parentheses, currency symbols, commas, and trailing dashes.
        Returns a clean number string (e.g., '-123.45') or None.
        """
        if not text:
            return None
        
        original_text = text
        text = text.strip()
        # Normalize different dash characters to a standard hyphen
        text = re.sub(r'[–—]', '-', text)
        
        # Remove currency symbols and commas for easier parsing
        text = text.replace('$', '').replace(',', '')
        # Handle cases like '1 234' -> '1234'
        text = re.sub(r'(\d)\s+(\d)', r'\1\2', text)

        # Handle parenthetical numbers, which can have spaces, e.g., "( 123.45 )" -> "-123.45"
        match = re.match(r'^\(\s*([\d\.]+)\s*\)$', text)
        if match:
            return '-' + match.group(1)

        # Handle trailing negative signs, e.g., "123.45-" -> "-123.45"
        if text.endswith('-'):
            text = '-' + text[:-1]
        
        # Final check if it's a valid floating point number
        try:
            float(text)
            return text
        except ValueError:
            # It's not a valid number
            return None

    def is_number(self, text: str) -> bool:
        """
        A lenient check to see if a word is potentially part of a number.
        The final validation is done in clean_and_validate_number.
        """
        # A word is potentially a number if it contains at least one digit.
        if not any(char.isdigit() for char in text):
            return False
        
        # And if it's composed only of characters common in financial numbers.
        # This is lenient and includes separators that might be part of the text.
        text_no_space = text.replace(' ', '')
        if all(c in '0123456789,.$()-–—' for c in text_no_space):
            return True
            
        return False

    def find_number_columns(self, words: List[dict]) -> List[Tuple[float, float]]:
        """
        Identifies distinct vertical columns of numbers on a page using clustering.
        This is robust against variations in horizontal spacing.
        """
        all_number_words = []
        for word in words:
            # We use is_number for a quick check, as clean_and_validate is more expensive
            if self.is_number(word['text']):
                all_number_words.append(word)

        if not all_number_words:
            return []

        # Cluster word x-positions to find columns
        x_coords = sorted([word['x0'] for word in all_number_words])
        
        if not x_coords:
            return []

        clusters = []
        if x_coords:
            current_cluster = [x_coords[0]]
            # A gap of ~4 characters is a good threshold for a new column
            COLUMN_GAP_THRESHOLD = 30 

            for x in x_coords[1:]:
                if x - current_cluster[-1] < COLUMN_GAP_THRESHOLD:
                    current_cluster.append(x)
                else:
                    clusters.append(current_cluster)
                    current_cluster = [x]
            clusters.append(current_cluster)

        # Define column boundaries from clusters
        final_columns = []
        for cluster in clusters:
            if not cluster: continue
            
            # Find all words that fall into this cluster's x-range
            min_x, max_x = min(cluster), max(cluster)
            cluster_words = [w for w in all_number_words if min_x <= w['x0'] <= max_x]
            
            # A real column must contain a minimum number of values
            if len(cluster_words) < 3:
                continue

            leftmost_x = min(w['x0'] for w in cluster_words)
            rightmost_x = max(w['x0'] + len(w['text']) * 7 for w in cluster_words) # Avg char width 7px
            
            final_columns.append((leftmost_x - 5, rightmost_x + 5)) # Add padding
        
        # Merge overlapping columns that were generated
        if not final_columns:
            return []
        
        final_columns.sort()
        merged_columns = [final_columns[0]]
        for current_start, current_end in final_columns[1:]:
            last_start, last_end = merged_columns[-1]
            if current_start < last_end: # Overlap detected
                merged_columns[-1] = (last_start, max(last_end, current_end))
            else:
                merged_columns.append((current_start, current_end))

        print(f"DEBUG: Found {len(merged_columns)} number columns: {merged_columns}")
        return merged_columns

    def process_numbers(self, lines: List[List[dict]], number_columns: List[Tuple[float, float]] = None, stmt_type: str = 'balance_sheet') -> List[dict]:
        """
        Processes lines of words into a structured format of {description, numbers}.
        It robustly separates text from numbers and places numbers in the correct column.
        """
        processed_lines = []
        if not number_columns:
            for line in lines:
                line_text = ' '.join(word['text'].replace('$', '').strip() for word in sorted(line, key=lambda w: w['x0']))
                if line_text.strip():
                    processed_lines.append({'description': line_text.strip(),'numbers': []})
            return processed_lines

        for line in lines:
            text_words = []
            potential_number_words = []

            # 1. Classify words into text or potential numbers
            for word in line:
                if self.is_number(word['text']):
                    potential_number_words.append(word)
                else:
                    text_words.append(word)

            # 2. Assemble the description from text words
            description = ' '.join(w['text'] for w in sorted(text_words, key=lambda w: w['x0']))
            
            # 3. Place numbers into their column buckets
            number_values = [None] * len(number_columns)
            
            for num_word in potential_number_words:
                cleaned_num = self.clean_and_validate_number(num_word['text'])
                if cleaned_num is None:
                    # If it looked like a number but wasn't, treat it as text
                    description += ' ' + num_word['text']
                    continue

                word_center_x = num_word['x0'] + (len(num_word['text']) * 7 / 2)
                
                best_col_idx = -1
                # Find which column this number belongs to
                for i, (col_start, col_end) in enumerate(number_columns):
                    if col_start <= word_center_x <= col_end:
                        best_col_idx = i
                        break
                
                if best_col_idx != -1:
                    number_values[best_col_idx] = cleaned_num
                else:
                    # If it didn't fit in a column, it's probably part of the description
                    description += ' ' + num_word['text']
            
            # Final cleanup of the description
            description = re.sub(r'\s+', ' ', description).strip()
            
            # Only add the line if it has content
            if description or any(v is not None for v in number_values):
                processed_lines.append({
                    'description': description,
                    'numbers': number_values
                })

        # --- Step 2: Assign sections to all lines (using a helper) ---
        matcher = TemplateMatcher()
        
        raw_lines_for_matcher = [
            {'description': l['description'], 'numbers': l['numbers']} for l in processed_lines
        ]

        if stmt_type == 'income_statement':
            assigned = matcher.assign_sections_by_context_is(raw_lines_for_matcher)
        elif stmt_type == 'cash_flow':
            assigned = matcher.assign_sections_by_context_cfs(raw_lines_for_matcher)
        else: # balance_sheet or default
            assigned = matcher.assign_sections_by_context(raw_lines_for_matcher)
        
        # --- Step 3: Robust row-combining logic ---
        final_processed_lines = []
        i = 0
        while i < len(assigned):
            curr = assigned[i]
            desc = curr['description']
            nums = curr['value'] if isinstance(curr['value'], list) else ([curr['value']] if curr['value'] else [])
            section = curr.get('section', None)
            
            # Combine description-only rows with the next row if it has numbers
            if desc and (not nums or all(n in [None, ''] for n in nums)) and i + 1 < len(assigned):
                next_row = assigned[i + 1]
                next_desc = next_row['description']
                next_nums = next_row['value'] if isinstance(next_row['value'], list) else ([next_row['value']] if next_row['value'] else [])
                next_section = next_row.get('section', None)

                if next_nums and section == next_section:
                    combined_desc = (desc + ' ' + next_desc).strip()
                    final_processed_lines.append({
                        'description': combined_desc,
                        'numbers': next_nums
                    })
                    i += 2
                    continue
            
            final_processed_lines.append({
                'description': desc.strip(),
                'numbers': nums
            })
            i += 1
            
        return final_processed_lines

    def export_to_excel(self, data: Dict[str, List[dict]], pdf_name: str):
        # Create a timestamp for the filename
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        excel_path = self.output_dir / f"{pdf_name}_{timestamp}.xlsx"
        
        print("\nDEBUG: Data to be written to Excel:")
        for stmt_type, lines in data.items():
            print(f"\nStatement type: {stmt_type}")
            print(f"Number of lines: {len(lines)}")
            print("First few lines:")
            for line in lines[:5]:
                print(f"Description: {line['description']}")
                print(f"Numbers: {line['numbers']}")
            print("-" * 50)
        
        # Create Excel writer
        with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
            for stmt_type, lines in data.items():
                if not lines:  # Skip empty statements
                    print(f"\nSkipping empty statement: {stmt_type}")
                    continue
                
                # Convert to DataFrame
                max_numbers = max(len(line['numbers']) for line in lines)
                df_data = {
                    'Description': [line['description'] for line in lines]
                }
                
                # Add number columns
                for i in range(max_numbers):
                    col_name = f'Value_{i+1}'
                    df_data[col_name] = [
                        line['numbers'][i] if i < len(line['numbers']) and line['numbers'][i] is not None else None
                        for line in lines
                    ]
                
                df = pd.DataFrame(df_data)
                
                print(f"\nWriting sheet: {stmt_type}")
                print(f"DataFrame shape: {df.shape}")
                print("DataFrame head:")
                print(df.head())
                
                # Write to Excel with formatting
                sheet_name = stmt_type.replace('_', ' ').title()
                print(f"Sheet name: {sheet_name}")
                df.to_excel(writer, sheet_name=sheet_name, index=False)
                
                # Get the worksheet
                worksheet = writer.sheets[sheet_name]
                
                # Format columns
                for idx, col in enumerate(df.columns):
                    column_letter = chr(65 + idx)
                    # Set column width
                    if col == 'Description':
                        worksheet.column_dimensions[column_letter].width = 60  # Width for description
                    else:
                        worksheet.column_dimensions[column_letter].width = 15  # Width for numbers
                
                # Format header
                from openpyxl.styles import Font, PatternFill
                header_font = Font(bold=True)
                header_fill = PatternFill(start_color='CCCCCC', end_color='CCCCCC', fill_type='solid')
                
                for cell in worksheet[1]:
                    cell.font = header_font
                    cell.fill = header_fill
        
        print(f"\nData exported to: {excel_path}")
        return excel_path

    def extract_text(self, pdf_path: str, process_numbers: bool = False, statement_pages: dict = None):
        try: 
            print(f"\nProcessing {pdf_path}")
            logging.info(f"\nProcessing {pdf_path}")
            
            # Initialize finder at the start since we need it for continuation pages
            finder = FinancialStatementFinder()
            
            if statement_pages:
                logging.info("\nReceived statement pages from final_find_fs.py:")
                for stmt_type, pages in statement_pages.items():
                    logging.info(f"{stmt_type}: {pages}")
            else:
                logging.info("No statement pages provided from final_find_fs.py")
            
            with pdfplumber.open(pdf_path) as pdf:
                # Initialize statements dictionary based on provided pages
                statements = {}
                if statement_pages:
                    for stmt_type, pages in statement_pages.items():
                        if pages:  # Only add if we have pages
                            statements[stmt_type] = {'pdf_page': pages[0], 'score': 100}  # Use first page if multiple
                            logging.info(f"Using page {pages[0]} for {stmt_type} from final_find_fs.py")
                
                # If no pages provided or found, fall back to direct scanning
                if not statements:
                    logging.info("\nFalling back to direct page scanning:")
                    
                    # First pass: find TOC and statement pages
                    toc_pages = []
                    has_toc = False
                    logging.info("\nDEBUG: First pass - Looking for TOC and statements")
                    logging.info(f"Total pages in PDF: {len(pdf.pages)}")
                    for page_num, page in enumerate(pdf.pages, 1):
                        text = page.extract_text()
                        logging.info(f"\nDEBUG: Checking page {page_num}")
                        is_toc = finder.is_table_of_contents(text, page_num)
                        if is_toc:
                            has_toc = True
                            toc_pages.append(page_num)
                            logging.info(f"DEBUG: Found TOC on page {page_num}")
                            # Process TOC to find statement pages
                            toc_statements = finder.process_toc_page(text, page_num)
                            logging.info("DEBUG: TOC statements found: %s", toc_statements)
                            for stmt_type, info in toc_statements.items():
                                if info['page'] is not None:
                                    if stmt_type not in statements:
                                        statements[stmt_type] = {'pdf_page': info['page'], 'score': info['score']}
                                        logging.info(f"DEBUG: Found {stmt_type} on page {info['page']} from TOC with score {info['score']}")
                                    elif info['score'] > statements[stmt_type]['score']:
                                        statements[stmt_type] = {'pdf_page': info['page'], 'score': info['score']}
                                        logging.info(f"DEBUG: Updated {stmt_type} to page {info['page']} from TOC with better score {info['score']}")
                    
                    # Second pass: direct scanning for statements
                    logging.info("\nDEBUG: Second pass - Direct scanning for statements")
                    for page_num, page in enumerate(pdf.pages, 1):
                        text = page.extract_text()
                        logging.info(f"\nDEBUG: Checking page {page_num}")
                        is_statement, score = finder.is_statement_page(text, page_num)
                        if is_statement and finder.last_statement_type:
                            stmt_type = finder.last_statement_type
                            logging.info(f"DEBUG: Found {stmt_type} on page {page_num} with score {score}")
                            if stmt_type not in statements or score > statements[stmt_type]['score']:
                                statements[stmt_type] = {'pdf_page': page_num, 'score': score}
                                logging.info(f"DEBUG: {'Updated' if stmt_type in statements else 'Added'} {stmt_type} to page {page_num} with score {score}")
                
                if not statements:
                    logging.info("No financial statements found")
                    return None, None
                
                logging.info("\nDEBUG: Final statement pages to be processed:")
                for stmt_type, info in statements.items():
                    logging.info(f"{stmt_type}: page {info['pdf_page']} (score: {info['score']}")
                
                # Extract and process each statement
                extracted_data = defaultdict(list)
                
                for stmt_type, info in statements.items():
                    # Handle multiple pages for a single statement type
                    pages_to_process = [info['pdf_page']]
                    if statement_pages and stmt_type in statement_pages:
                        pages_to_process = statement_pages[stmt_type]
                    
                    all_processed_lines = []
                    for page_num in pages_to_process:
                        if page_num > len(pdf.pages):
                            print(f"PAGE NUMBER ERROR: {pdf_path} has only {len(pdf.pages)} pages but exceeded")
                            continue
                            
                        print(f"\nExtracting {stmt_type} from page {page_num}")
                        page = pdf.pages[page_num - 1]
                        
                        words = page.extract_words(
                            x_tolerance=1, 
                            y_tolerance=1, 
                            keep_blank_chars=False,
                            use_text_flow=True, # More robust for reading order
                            extra_attrs=["fontname", "size"]
                        )
                        
                        # --- New Line Clustering Logic ---
                        if not words: continue
                        
                        # Sort words primarily by vertical position, then horizontal
                        words.sort(key=lambda w: (w['top'], w['x0']))
                        
                        lines = []
                        current_line = []
                        if words:
                            current_line.append(words[0])
                            for word in words[1:]:
                                # If word is on the same line (small vertical distance)
                                if abs(word['top'] - current_line[-1]['top']) < 5:
                                    current_line.append(word)
                                else:
                                    # New line detected
                                    lines.append(sorted(current_line, key=lambda w: w['x0']))
                                    current_line = [word]
                            lines.append(sorted(current_line, key=lambda w: w['x0']))
                        
                        filtered_lines = [line for line in lines if line] # Remove empty lines
                        
                        # Find number columns using all words on the page
                        number_columns = self.find_number_columns(words) if process_numbers else []
                        
                        # Process lines into structured data
                        processed_lines = self.process_numbers(filtered_lines, number_columns, stmt_type=stmt_type)
                        print(f"DEBUG: Processed {len(processed_lines)} lines for {stmt_type} on page {page_num}")
                        all_processed_lines.extend(processed_lines)

                    # Store the processed lines for the statement type
                    extracted_data[stmt_type].extend(all_processed_lines)
                    print(f"DEBUG: Total {len(all_processed_lines)} lines stored for {stmt_type}")
                    
                    # Print the table for the entire statement
                    if all_processed_lines:
                        # Find the maximum width for the description column
                        max_desc_width = max((len(line['description']) for line in all_processed_lines if line['description']), default=30)
                        max_desc_width = min(max_desc_width, 80)
                        
                        # Find the number of columns needed
                        max_numbers = max((len(line['numbers']) for line in all_processed_lines), default=0)
                        
                        # Print header
                        header = f"{'Description':<{max_desc_width}}"
                        for i in range(max_numbers):
                            header += f" | {'Value ' + str(i+1):>15}"
                        print("\n" + "=" * len(header))
                        print(header)
                        print("=" * len(header))
                        
                        # Print rows
                        for line in all_processed_lines:
                            desc = (line['description'] or '')[:max_desc_width]
                            row_str = f"{desc:<{max_desc_width}}"
                            for i in range(max_numbers):
                                num = line['numbers'][i] if i < len(line['numbers']) and line['numbers'][i] is not None else '-'
                                row_str += f" | {str(num):>15}"
                            print(row_str)
                        print("=" * len(header) + "\n")
                
                # Export all data to Excel
                if extracted_data:
                    pdf_name = Path(pdf_path).stem
                    excel_path = self.export_to_excel(extracted_data, pdf_name)
                    return excel_path, extracted_data
                
                return None, None

        except Exception as e:
            print(f"Error extracting text: {str(e)}")
            import traceback
            traceback.print_exc()
            return None, None

def main(): 
    # Get project root directory (two levels up from this file)
    current_dir = Path(__file__).resolve().parent
    project_root = current_dir.parent.parent
    
    # Set up input and output directories relative to project root
    input_pdfs_dir = project_root / "input_pdfs"
    input_pdfs_dir.mkdir(exist_ok=True)

    if len(sys.argv) > 1: 
        pdf_filename = sys.argv[1]
    else: 
        pdf_filename = input("Enter PDF filename: ")

    pdf_path = input_pdfs_dir / pdf_filename
    pdf_path = str(pdf_path.resolve())

    if not Path(pdf_path).exists():
        print(f"PDF file path not found: {pdf_path}")
        return
    
    # Step 1: Use final_find_fs.py to get statement pages
    print("Step 1: Finding financial statement pages...")
    finder = FinancialStatementFinder()
    lines, toc_pages, statement_pages = finder.extractContent(pdf_path)
    
    if statement_pages:
        print(f"Found statement pages: {statement_pages}")
        # Get the high confidence pages (>=50%)
        high_conf_pages = finder.get_statement_pages()
        print(f"High confidence pages (>=50%): {high_conf_pages}")
        
        # Convert to the format expected by extract_text
        statement_pages_dict = {}
        for stmt_type, pages in high_conf_pages.items():
            if pages:  # Only add if we have pages
                statement_pages_dict[stmt_type] = pages
    else:
        print("No statement pages found by final_find_fs.py")
        statement_pages_dict = None
    
    # Step 2: Extract text using the found pages
    print("\nStep 2: Extracting text from identified pages...")
    extractor = TextExtractor()
    excel_path, extracted_data = extractor.extract_text(pdf_path, process_numbers=True, statement_pages=statement_pages_dict)
    
    if excel_path:
        print(f"\nExtraction completed successfully!")
        print(f"Excel file saved to: {excel_path}")
    else:
        print("\nExtraction failed!")

if __name__ == "__main__":
    main()