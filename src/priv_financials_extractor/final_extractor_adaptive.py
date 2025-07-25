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
from sklearn.cluster import DBSCAN

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
        
        # --- NEW: Filter out year-like numbers ---
        cleaned_text = text.strip()
        if len(cleaned_text) == 4 and cleaned_text.isdigit() and 1990 <= int(cleaned_text) <= 2050:
            return None
        # --- END NEW ---
        
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
            # Filter out numbers that are likely to be years
            text = word['text'].strip()
            if len(text) == 4 and text.isdigit() and 1990 <= int(text) <= 2050:
                continue

            # We use is_number for a quick check, as clean_and_validate is more expensive
            if self.is_number(word['text']):
                all_number_words.append(word)

        if len(all_number_words) < 3: # Not enough numbers to form a column
            return []

        # Use DBSCAN to find clusters of numbers based on their x-position
        # The features are the x-coordinates of the start of each number word
        x_coords = np.array([w['x0'] for w in all_number_words]).reshape(-1, 1)

        # eps: The maximum distance between two samples for one to be considered as in the neighborhood of the other.
        # This is effectively our column width tolerance. 25px is a reasonable starting point.
        # min_samples: The number of samples in a neighborhood for a point to be considered as a core point.
        # This is our minimum number of values to form a column.
        db = DBSCAN(eps=25, min_samples=3).fit(x_coords)
        
        labels = db.labels_
        n_clusters = len(set(labels)) - (1 if -1 in labels else 0)
        
        print(f"DEBUG: DBSCAN found {n_clusters} clusters (columns).")

        final_columns = []
        for i in range(n_clusters):
            cluster_indices = np.where(labels == i)[0]
            
            cluster_words = [all_number_words[j] for j in cluster_indices]

            if not cluster_words:
                continue

            # Define column boundary from the words in the cluster
            leftmost_x = min(w['x0'] for w in cluster_words)
            # Calculate the rightmost edge by finding the end of the longest word in the cluster
            rightmost_x = max(w['x0'] + len(w['text']) * 7.5 for w in cluster_words) # Avg char width
            
            final_columns.append((leftmost_x - 5, rightmost_x + 5)) # Add padding
        
        final_columns.sort()
        
        print(f"DEBUG: Found {len(final_columns)} number columns: {final_columns}")
        return final_columns

    def find_column_headers(self, words: List[dict]) -> Dict[str, float]:
        """
        Finds potential year-based column headers (e.g., "2023", "2024") at the top of the page.
        Returns a dictionary mapping the year (as a string) to its horizontal position.
        """
        headers = {}
        # Heuristic: Scan the top 20% of the page for headers
        page_top_y = min(w['top'] for w in words) if words else 0
        page_bottom_y = max(w['bottom'] for w in words) if words else 0
        header_zone_y = page_top_y + (page_bottom_y - page_top_y) * 0.2
        
        for word in words:
            if word['top'] > header_zone_y:
                continue # Only look at the top of the page
                
            text = word['text'].strip()
            # Basic check for a 4-digit number that looks like a year
            if len(text) == 4 and text.isdigit() and 1990 <= int(text) <= 2050:
                # Use the center of the word as its position
                x_pos = word['x0'] + (word['x1'] - word['x0']) / 2
                headers[text] = x_pos
                
        # Sort by horizontal position to handle cases where years are out of order
        sorted_headers = dict(sorted(headers.items(), key=lambda item: item[1]))
        print(f"DEBUG: Found potential year headers: {sorted_headers}")
        return sorted_headers

    def find_and_label_columns(self, words: List[dict], year_headers: Dict[str, float]) -> Dict[int, str]:
        """
        Identifies number columns and labels them with the nearest year header.
        Returns a dictionary mapping a column index to its year label.
        """
        # Step 1: Find number columns using the existing DBSCAN approach
        number_columns_coords = self.find_number_columns(words) # This returns [(x0, x1), ...]
        if not number_columns_coords:
            return {}

        labeled_columns = {}
        
        # Step 2: Assign each found number column to the nearest year header
        for i, (col_x0, col_x1) in enumerate(number_columns_coords):
            col_center = col_x0 + (col_x1 - col_x0) / 2
            
            # Find the year header with the minimum horizontal distance to the column's center
            if year_headers:
                closest_year = min(
                    year_headers.keys(),
                    key=lambda year: abs(year_headers[year] - col_center)
                )
                labeled_columns[i] = closest_year
            else:
                # Fallback if no year headers were found
                labeled_columns[i] = f"Value_{i+1}"

        print(f"DEBUG: Labeled columns: {labeled_columns}")
        return labeled_columns

    def process_numbers(self, lines: List[List[dict]], number_columns: List[Tuple[float, float]], labeled_columns: Dict[int, str]) -> List[dict]:
        """
        Processes lines of words into a structured format of {description, numbers:{year: value}}.
        """
        processed_lines = []
        if not number_columns:
            # If no number columns are found, just return all text as descriptions.
            for line in lines:
                line_text = ' '.join(word['text'].replace('$', '').strip() for word in sorted(line, key=lambda w: w['x0']))
                if line_text.strip():
                    processed_lines.append({'description': line_text.strip(), 'numbers': {}})
            return processed_lines

        for line in lines:
            text_words = []
            # Use a dictionary for number_values to store year-based values
            number_values = {label: None for label in labeled_columns.values()}
            
            for word in line:
                word_center_x = word['x0'] + (len(word['text']) * 7 / 2) # Approximate center
                is_in_number_column = False
                
                for i, (col_start, col_end) in enumerate(number_columns):
                    if col_start <= word_center_x <= col_end:
                        cleaned_num = self.clean_and_validate_number(word['text'])
                        if cleaned_num is not None and i in labeled_columns:
                            year_label = labeled_columns[i]
                            number_values[year_label] = cleaned_num
                            is_in_number_column = True
                        break
                
                if not is_in_number_column:
                    text_words.append(word)

            description = ' '.join(w['text'] for w in sorted(text_words, key=lambda w: w['x0']))
            description = re.sub(r'\\s+', ' ', description).strip()
            
            if len(description) <= 2 and any(v is not None for v in number_values.values()):
                continue

            if description or any(v is not None for v in number_values.values()):
                    processed_lines.append({
                    'description': description,
                    'numbers': number_values
                })
        
        # The row-combining logic needs to be adapted for the new numbers structure
        final_processed_lines = []
        i = 0
        while i < len(processed_lines):
            curr = processed_lines[i]
            desc = curr['description']
            nums_dict = curr.get('numbers', {})

            # Combine description-only rows with the next row if it has numbers
            if desc and all(n is None for n in nums_dict.values()) and i + 1 < len(processed_lines):
                next_row = processed_lines[i + 1]
                next_desc = next_row['description']
                next_nums_dict = next_row.get('numbers', {})
                
                if any(n is not None for n in next_nums_dict.values()):
                    combined_desc = (desc + ' ' + next_desc).strip()
                    final_processed_lines.append({
                        'description': combined_desc,
                        'numbers': next_nums_dict
                    })
                    i += 2
                    continue
            
            final_processed_lines.append(curr)
            i += 1
            
        return final_processed_lines

    def merge_multiline_descriptions(self, df):
        """Merge multi-row descriptions where the value is only on the second row."""
        merged_rows = []
        prev_desc = None
        prev_row = None
        value_cols = [col for col in df.columns if col != 'Description']
        for idx, row in df.iterrows():
            desc = str(row['Description'])
            has_value = any(pd.notna(row[col]) for col in value_cols)
            if (desc.strip().startswith('and ') or (desc and desc[0].islower() and prev_desc)) and prev_row is not None:
                # Merge with previous
                merged_desc = prev_row['Description'] + ' ' + desc.strip()
                prev_row['Description'] = merged_desc
                # If this row has value, move it up
                for col in value_cols:
                    if pd.notna(row[col]):
                        prev_row[col] = row[col]
            else:
                prev_row = row.copy()
                prev_desc = desc
                merged_rows.append(prev_row)
        return pd.DataFrame(merged_rows)

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
                print(f"\n[DEBUG] Processing statement: {stmt_type}")
                print(f"[DEBUG] lines type: {type(lines)}")
                print(f"[DEBUG] lines content: {lines}")
                if not lines or not isinstance(lines, list) or not all(isinstance(line, dict) for line in lines):
                    print(f"[WARNING] Skipping statement: {stmt_type} (empty or invalid format)")
                    continue
                # Collect all years present in any line
                all_years = set()
                for line in lines:
                    all_years.update(line['numbers'].keys())
                all_years = sorted(all_years)
                df_data = {
                    'Description': [line['description'] for line in lines]
                }
                for year in all_years:
                    df_data[year] = [line['numbers'].get(year, None) for line in lines]
                df = pd.DataFrame(df_data)
                # --- NEW: Merge multi-row descriptions ---
                df = self.merge_multiline_descriptions(df)
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
            
            finder = FinancialStatementFinder()
            
            if not statement_pages:
                logging.info("No statement pages provided. Falling back to direct scanning.")
                # This fallback is simplified; in a real scenario, we'd run the full finder logic.
                return None, None
            
            with pdfplumber.open(pdf_path) as pdf:
                # This will hold the final data in the format the mapper expects
                final_extracted_data = {}

                for stmt_type, page_nums in statement_pages.items():
                    all_processed_lines = []
                    for page_num in page_nums:
                        if page_num > len(pdf.pages):
                            print(f"PAGE NUMBER ERROR: {pdf_path} has only {len(pdf.pages)} pages but exceeded")
                            continue
                            
                        print(f"\nExtracting {stmt_type} from page {page_num}")
                        page = pdf.pages[page_num - 1]
                        
                        # ADAPTIVE IMPROVEMENT: Adjust tolerances based on font size detection
                        chars = page.chars
                        if chars:
                            font_sizes = [c.get('size', 10) for c in chars if c.get('size')]
                            avg_font_size = sum(font_sizes) / len(font_sizes) if font_sizes else 10
                            print(f"DEBUG: Average font size on page {page_num}: {avg_font_size:.1f}")
                            
                            # Adaptive tolerances based on font size
                            if avg_font_size < 9:  # Small font (like CFS)
                                x_tolerance = 1.5  # Tighter horizontal tolerance
                                y_tolerance = 1.5  # Tighter vertical tolerance
                                print(f"DEBUG: Using SMALL FONT tolerances (x={x_tolerance}, y={y_tolerance})")
                            elif avg_font_size > 12:  # Large font
                                x_tolerance = 3.0
                                y_tolerance = 3.0
                                print(f"DEBUG: Using LARGE FONT tolerances (x={x_tolerance}, y={y_tolerance})")
                            else:  # Medium font
                                x_tolerance = 2.0
                                y_tolerance = 2.0
                                print(f"DEBUG: Using MEDIUM FONT tolerances (x={x_tolerance}, y={y_tolerance})")
                        else:
                            x_tolerance = 2.0
                            y_tolerance = 2.0
                            print("DEBUG: Using DEFAULT tolerances (no font info)")
                        
                        words = page.extract_words(
                            x_tolerance=x_tolerance, 
                            y_tolerance=y_tolerance, 
                            keep_blank_chars=False,
                            use_text_flow=True, 
                            extra_attrs=["fontname", "size"]
                        )
                        if not words: continue
                        
                        # ADAPTIVE IMPROVEMENT: Post-process to merge obvious split words
                        words = self.merge_split_words(words)
                        
                        words.sort(key=lambda w: (w['top'], w['x0']))
                        
                        lines = []
                        current_line = []
                        if words:
                            current_line.append(words[0])
                            for word in words[1:]:
                                # Adaptive line detection based on font size
                                line_tolerance = max(3, avg_font_size * 0.4) if 'avg_font_size' in locals() else 5
                                if abs(word['top'] - current_line[-1]['top']) < line_tolerance:
                                    current_line.append(word)
                                else:
                                    lines.append(sorted(current_line, key=lambda w: w['x0']))
                                    current_line = [word]
                            lines.append(sorted(current_line, key=lambda w: w['x0']))
                        
                        filtered_lines = [line for line in lines if line]
                        
                        # --- New Year-Aware Column Processing ---
                        year_headers = self.find_column_headers(words)
                        number_columns_coords = self.find_number_columns(words)
                        labeled_columns = self.find_and_label_columns(words, year_headers)
                        
                        processed_lines = self.process_numbers(filtered_lines, number_columns_coords, labeled_columns)
                        all_processed_lines.extend(processed_lines)

                    # Store the processed lines in order for this statement type
                    final_extracted_data[stmt_type] = all_processed_lines
                    # (Optional) Print the extracted table for debugging
                    if all_processed_lines:
                        pass 

                if final_extracted_data:
                    excel_path = self.export_to_excel(final_extracted_data, Path(pdf_path).stem)
                    return excel_path, final_extracted_data
                else:
                    return None, None

        except Exception as e:
            print(f"Error extracting text: {str(e)}")
            import traceback
            traceback.print_exc()
            return None, None

    def merge_split_words(self, words: List[dict]) -> List[dict]:
        """
        Post-process to merge words that were incorrectly split due to OCR issues.
        This is especially common in small fonts where spaces appear in the middle of words.
        """
        if not words:
            return words
            
        merged_words = []
        i = 0
        
        while i < len(words):
            current_word = words[i]
            current_text = current_word['text']
            
            # Look ahead to see if we can merge with next words
            merge_candidates = [current_word]
            j = i + 1
            
            # Try to merge up to 3 consecutive words if they look like parts of the same word
            while j < len(words) and j < i + 4:
                next_word = words[j]
                
                # Check if these words should be merged
                if self.should_merge_words(current_word, next_word, [w['text'] for w in merge_candidates]):
                    merge_candidates.append(next_word)
                    j += 1
                else:
                    break
            
            if len(merge_candidates) > 1:
                # Merge the words
                merged_text = ' '.join(w['text'] for w in merge_candidates)
                # Clean up common OCR artifacts in merged text
                merged_text = self.clean_merged_text(merged_text)
                
                # Create merged word with combined bounds
                merged_word = {
                    'text': merged_text,
                    'x0': merge_candidates[0]['x0'],
                    'x1': merge_candidates[-1]['x1'],
                    'top': min(w['top'] for w in merge_candidates),
                    'bottom': max(w['bottom'] for w in merge_candidates),
                    'fontname': current_word.get('fontname', ''),
                    'size': current_word.get('size', 10)
                }
                merged_words.append(merged_word)
                print(f"DEBUG: Merged '{' + '.join(w['text'] for w in merge_candidates)}' → '{merged_text}'")
            else:
                merged_words.append(current_word)
            
            i = j
        
        return merged_words
    
    def should_merge_words(self, word1: dict, word2: dict, candidate_texts: List[str]) -> bool:
        """
        Determine if two words should be merged based on various heuristics.
        """
        text1 = word1['text']
        text2 = word2['text']
        
        # Don't merge if either is empty
        if not text1.strip() or not text2.strip():
            return False
        
        # Don't merge if they're too far apart horizontally
        horizontal_distance = word2['x0'] - word1['x1']
        if horizontal_distance > 15:  # More than ~2 characters apart
            return False
            
        # Don't merge if they're on different lines (vertical distance too large)
        vertical_distance = abs(word1['top'] - word2['top'])
        if vertical_distance > 3:
            return False
        
        # Don't merge numbers with letters unless it's a clear pattern
        if text1.isdigit() and text2.isalpha():
            return False
        if text1.isalpha() and text2.isdigit() and len(text2) > 2:
            return False
            
        # Pattern-based merging for common OCR splits
        combined_text = text1 + text2
        
        # Common financial terms that get split
        financial_patterns = [
            r'depreci.*tion',
            r'amortiza.*tion', 
            r'opera.*ting',
            r'activi.*ties',
            r'equip.*ment',
            r'liabil.*ities',
            r'assets?',
            r'receiv.*able',
            r'invent.*ories',
            r'payable',
            r'financing',
            r'proceeds',
            r'purchase',
            r'distributions?'
        ]
        
        # Check if the combination forms a known financial term
        for pattern in financial_patterns:
            if re.search(pattern, combined_text.lower()):
                return True
        
        # Merge single letters with the next word (common OCR split)
        if len(text1) == 1 and text1.isalpha() and len(text2) > 1:
            return True
            
        # Merge if first word ends with partial word and second continues it
        if len(text1) >= 2 and text1[-1].isalpha() and text2[0].isalpha():
            # Check if it looks like a word was split in the middle
            if len(text1) <= 4 or len(text2) <= 4:
                return True
        
        return False
    
    def clean_merged_text(self, text: str) -> str:
        """
        Clean up common OCR artifacts in merged text.
        """
        # Fix common split patterns but preserve necessary spaces
        text = re.sub(r'\b(depreci)\s+(ation)\b', r'depreciation', text, flags=re.IGNORECASE)
        text = re.sub(r'\b(amortiz)\s+(ation)\b', r'amortization', text, flags=re.IGNORECASE)
        text = re.sub(r'\b(opera)\s+(ting)\b', r'operating', text, flags=re.IGNORECASE)
        text = re.sub(r'\b(activi)\s+(ties?)\b', r'activities', text, flags=re.IGNORECASE)
        text = re.sub(r'\b(equi)\s+(pment)\b', r'equipment', text, flags=re.IGNORECASE)
        text = re.sub(r'\b(liabil)\s+(ities?)\b', r'liabilities', text, flags=re.IGNORECASE)
        text = re.sub(r'\b(receiv)\s+(able?)\b', r'receivable', text, flags=re.IGNORECASE)
        text = re.sub(r'\b(invent)\s+(ori)\s+(es?)\b', r'inventories', text, flags=re.IGNORECASE)
        text = re.sub(r'\b(financ)\s+(ing)\b', r'financing', text, flags=re.IGNORECASE)
        
        # DON'T remove all spaces - be more selective
        # Only remove spaces between single characters that are clearly OCR errors
        text = re.sub(r'\b([a-z])\s+([a-z])\b', r'\1\2', text, flags=re.IGNORECASE)
        
        # Preserve important phrase separations
        text = re.sub(r'activities:\s*', 'activities: ', text, flags=re.IGNORECASE)
        text = re.sub(r'statements\s*of\s*', 'statements of ', text, flags=re.IGNORECASE)
        text = re.sub(r'cash\s*flows?\s*', 'cash flows ', text, flags=re.IGNORECASE)
        text = re.sub(r'net\s*income\s*', 'net income ', text, flags=re.IGNORECASE)
        text = re.sub(r'and\s*equipment\s*', 'and equipment ', text, flags=re.IGNORECASE)
        text = re.sub(r'long-term\s*', 'long-term ', text, flags=re.IGNORECASE)
        
        # Clean up multiple spaces but don't eliminate all spaces
        text = re.sub(r'\s+', ' ', text).strip()
        
        return text

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
    
    # Step 1: Use final_find_fs.py to get statement pages with confirmation
    print("Step 1: Finding financial statement pages...")
    finder = FinancialStatementFinder()
    lines, toc_pages, statement_pages = finder.extractContent(pdf_path)
    
    if statement_pages:
        print(f"Found statement pages: {statement_pages}")
        
        # Use the confirmation system to get user-verified pages
        confirmed_pages = finder.confirm_statement_pages(pdf_filename)
        
        # Convert to the format expected by extract_text
        statement_pages_dict = {}
        for stmt_type, pages in confirmed_pages.items():
            if pages:  # Only add if we have pages
                statement_pages_dict[stmt_type] = pages
    else:
        print("No statement pages found by final_find_fs.py")
        statement_pages_dict = None
    
    # Step 2: Extract text using the confirmed pages
    print("\nStep 2: Extracting text from confirmed pages...")
    extractor = TextExtractor()
    excel_path, extracted_data = extractor.extract_text(pdf_path, process_numbers=True, statement_pages=statement_pages_dict)
    
    if excel_path:
        print(f"\nExtraction completed successfully!")
        print(f"Excel file saved to: {excel_path}")
    else:
        print("\nExtraction failed!")

if __name__ == "__main__":
    main()