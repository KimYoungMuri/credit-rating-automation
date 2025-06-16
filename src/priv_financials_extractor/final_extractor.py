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

class TextExtractor: 
    def __init__(self):
        self.setup_logging()
        self.finder = FinancialStatementFinder()
        self.output_dir = Path("output_excel")
        self.output_dir.mkdir(exist_ok=True)

    def setup_logging(self):
        """Setup logging configuration"""
        for handler in logging.root.handlers[:]:
            logging.root.removeHandler(handler)
        formatter = logging.Formatter('%(message)s')
        file_handler = logging.FileHandler('text_extraction.log', mode='w', encoding='utf-8')
        file_handler.setFormatter(formatter)
        logging.root.addHandler(file_handler)
        logging.root.setLevel(logging.INFO)
        logging.info("Logging initialized successfully")

    def clean_if_number(self, text: str) -> str:
        original_text = text
        text = text.strip()
        
        text = re.sub(r'(\d)\s+(\d)', r'\1\2', text)
        text = re.sub(r'\s*,\s*', ',', text)
        
        # Handle parenthetical numbers first
        if text.startswith('(') and text.endswith(')'):
            inner_text = text[1:-1].replace('$', '').replace(',', '').strip()
            if inner_text.isdigit() or (inner_text.replace('.', '').isdigit() and inner_text.count('.') == 1):
                return '-' + inner_text
        
        # Handle regular numbers
        if text.replace('$', '').replace(',', '').strip().isdigit():
            return text.replace('$', '').replace(',', '').strip()
        
        if re.fullmatch(r'-?\\d+(\\.\\d+)?', text):
            return text
        
        return original_text

    def convert_to_float(self, text: str) -> float:
        text = text.replace('$', '').replace(' ', '')
        if text.startswith('(') and text.endswith(')'): #parantheses
            text = '-' + text[1:-1]
        return float(text.replace(',', ''))

    def is_number(self, text: str) -> bool:
        text = re.sub(r'(\d)\s+(\d)', r'\1\2', text.strip())
        text = re.sub(r'\s*,\s*', ',', text)
        #return bool(re.match(r'^[\$\s]*-?[\d,]+(\.\d+)?$|^\([\d,]+(\.\d+)?\)$', text))
        return bool(re.match(r'^[\$\s]*[-\(]?[\d,]+(\.\d+)?[\)]?$', text)) #attempted fix
    def find_number_columns(self, words: List[dict]) -> Optional[Tuple[float, float]]:
        x_positions = defaultdict(list)
        
        for word in words:
            word_text = re.sub(r'(\d)\s+(\d)', r'\1\2', word['text'].strip())
            word_text = re.sub(r'\s*,\s*', ',', word_text)
            
            if self.is_number(word_text):
                x_rounded = round(word['x0'] / 5) * 5
                x_positions[x_rounded].append({
                    'x0': word['x0'],
                    'top': word['top'],
                    'text': word_text
                })
        
        number_columns = []
        for x_pos, numbers in x_positions.items():
            # Only consider positions with enough numbers
            if len(numbers) >= 3:
                has_dollar_signs = any('$' in n['text'] for n in numbers)
                try:
                    # Look for large numbers and numbers with commas
                    has_financial_format = any(
                        abs(self.convert_to_float(n['text'])) > 1000 or 
                        ',' in n['text'] or 
                        '$' in n['text']
                        for n in numbers
                    )
                    if has_dollar_signs or has_financial_format:
                        number_columns.append(x_pos)
                except ValueError:
                    continue
        
        if number_columns:
            all_numbers = []
            for x_pos in number_columns:
                all_numbers.extend(x_positions[x_pos])
            
            leftmost_x = min(n['x0'] for n in all_numbers)
            rightmost_x = max(n['x0'] for n in all_numbers)
            return leftmost_x - 20, rightmost_x + 20
        return None

    def process_numbers(self, lines: List[List[dict]], number_zone: Tuple[float, float] = None) -> List[str]:
        processed_lines = []
        
        if number_zone:
            number_zone_start, number_zone_end = number_zone
            
            for line in lines:
                numbers = []
                text = []
                
                current_number = []
                last_x0 = None
                
                words_in_line = sorted(line, key=lambda w: w['x0'])
                i = 0
                while i < len(words_in_line):
                    word = words_in_line[i]
                    word_text = word['text'].strip()
                    
                    # Check for split parenthetical numbers
                    if word_text == '(' and i + 1 < len(words_in_line):
                        next_word = words_in_line[i + 1]
                        next_text = next_word['text'].strip()
                        
                        # If next word ends with ), combine them
                        if next_text.endswith(')'):
                            combined_text = f"({next_text}"
                            if self.is_number(combined_text) and number_zone_start <= word['x0'] <= number_zone_end:
                                numbers.append({
                                    'text': self.clean_if_number(combined_text),
                                    'x0': word['x0']
                                })
                                i += 2  # Skip both the ( and the number)
                                continue
                    
                    # Regular number processing
                    if re.match(r'^[\$\s]*-?[\d,\s]+(\.\d+)?$|^\([\d,]+(\.\d+)?\)$', word_text):
                        if number_zone_start <= word['x0'] <= number_zone_end:
                            if last_x0 is None or word['x0'] - last_x0 <= 15:
                                current_number.append(word)
                            else:
                                if current_number:
                                    combined_text = ' '.join(w['text'] for w in current_number)
                                    if self.is_number(combined_text):
                                        numbers.append({
                                            'text': self.clean_if_number(combined_text),
                                            'x0': current_number[0]['x0']
                                        })
                                    else:
                                        text.extend(current_number)
                                current_number = [word]
                            last_x0 = word['x0']
                        else:
                            text.append(word)
                    else:
                        if current_number:
                            combined_text = ' '.join(w['text'] for w in current_number)
                            if self.is_number(combined_text):
                                numbers.append({
                                    'text': self.clean_if_number(combined_text),
                                    'x0': current_number[0]['x0']
                                })
                            else:
                                text.extend(current_number)
                            current_number = []
                        text.append(word)
                    i += 1
                
                # Don't forget to process the last number group
                if current_number:
                    combined_text = ' '.join(w['text'] for w in current_number)
                    if self.is_number(combined_text):
                        numbers.append({
                            'text': self.clean_if_number(combined_text),
                            'x0': current_number[0]['x0']
                        })
                    else:
                        text.extend(current_number)
                
                # Process text
                text_words = sorted(text, key=lambda w: w['x0'])
                # Handle trailing dash
                text_part = ' '.join(w['text'].replace('$', '').strip() for w in text_words)
                text_part = text_part.strip()
                if text_part.endswith(' -'):
                    text_part = text_part[:-2].strip()  # Remove trailing dash
                    # Add a None placeholder in the numbers list
                    if not numbers:  # Only if there are no numbers yet
                        numbers.append({'text': None, 'x0': float('inf')})
                
                number_part = [n['text'] for n in sorted(numbers, key=lambda n: n['x0'])]
                
                if text_part or number_part:  # Only add non-empty lines
                    processed_lines.append({
                        'description': text_part.strip(),
                        'numbers': number_part
                    })
        else:
            for line in lines:
                line_text = ' '.join(word['text'].replace('$', '').strip() for word in sorted(line, key=lambda w: w['x0']))
                if line_text.strip():  # Only add non-empty lines
                    processed_lines.append({
                        'description': line_text.strip(),
                        'numbers': []
                    })
        
        return processed_lines

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
                    if info['pdf_page'] is not None:
                        print(f"\nExtracting {stmt_type} from page {info['pdf_page']}")
                        
                        # Get the statement page
                        page_num = info['pdf_page']
                        if page_num > len(pdf.pages):
                            print(f"PAGE NUMBER ERROR: {pdf_path} has only {len(pdf.pages)} pages but exceeded")
                            continue
                            
                        page = pdf.pages[page_num - 1]
                        text = page.extract_text()
                        print(f"DEBUG: Page {page_num} text starts with: {text[:100].replace('\n', ' ')}")
                        
                        # Extract words with precise positioning
                        words = page.extract_words(
                            x_tolerance=3,
                            y_tolerance=3,
                            keep_blank_chars=False,
                            use_text_flow=False,
                            horizontal_ltr=True,
                            vertical_ttb=True,
                            extra_attrs=["fontname", "size"]
                        )
                        
                        # Group words into lines
                        lines = {}
                        for word in words:
                            y = round(word['top'])
                            if y not in lines:
                                lines[y] = []
                            lines[y].append(word)
                        
                        # Filter and sort lines
                        filtered_lines = []
                        for y in sorted(lines.keys()):
                            line = sorted(lines[y], key=lambda w: w['x0'])
                            if len(line) > 0:
                                # Filter out single-letter watermarks
                                if len(line[-1]['text'].strip()) == 1 and line[-1]['text'].strip().isupper():
                                    if len(line) > 1:
                                        filtered_lines.append(line[:-1])
                                else:
                                    filtered_lines.append(line)
                        
                        # Find number columns
                        number_zone = self.find_number_columns(words) if process_numbers else None
                        
                        # Process lines into structured data
                        processed_lines = self.process_numbers(filtered_lines, number_zone)
                        print(f"DEBUG: Processed {len(processed_lines)} lines for {stmt_type}")
                        
                        # Check for continuation pages
                        continuation_pages = []
                        next_page = page_num + 1
                        while next_page <= len(pdf.pages):
                            next_text = pdf.pages[next_page - 1].extract_text()
                            is_cont, cont_type = finder.is_continuation_page(next_text, next_page)
                            if is_cont and cont_type == stmt_type:
                                print(f"DEBUG: Found continuation page {next_page} for {stmt_type}")
                                continuation_pages.append(next_page)
                                # Process continuation page
                                cont_page = pdf.pages[next_page - 1]
                                cont_words = cont_page.extract_words(
                                    x_tolerance=3,
                                    y_tolerance=3,
                                    keep_blank_chars=False,
                                    use_text_flow=False,
                                    horizontal_ltr=True,
                                    vertical_ttb=True,
                                    extra_attrs=["fontname", "size"]
                                )
                                
                                # Group and process continuation lines
                                cont_lines = {}
                                for word in cont_words:
                                    y = round(word['top'])
                                    if y not in cont_lines:
                                        cont_lines[y] = []
                                    cont_lines[y].append(word)
                                
                                cont_filtered_lines = []
                                for y in sorted(cont_lines.keys()):
                                    line = sorted(cont_lines[y], key=lambda w: w['x0'])
                                    if len(line) > 0:
                                        if len(line[-1]['text'].strip()) == 1 and line[-1]['text'].strip().isupper():
                                            if len(line) > 1:
                                                cont_filtered_lines.append(line[:-1])
                                        else:
                                            cont_filtered_lines.append(line)
                                
                                # Use same number zone for consistency
                                cont_processed = self.process_numbers(cont_filtered_lines, number_zone)
                                print(f"DEBUG: Processed {len(cont_processed)} lines from continuation page {next_page}")
                                processed_lines.extend(cont_processed)
                                next_page += 1
                            else:
                                break
                        
                        # Store the processed lines
                        extracted_data[stmt_type].extend(processed_lines)
                        print(f"DEBUG: Total {len(processed_lines)} lines stored for {stmt_type}")
                        
                        # Print the table
                        if processed_lines:
                            # Find the maximum width for the description column
                            max_desc_width = max(len(line['description']) for line in processed_lines)
                            max_desc_width = min(max_desc_width, 60)  # Cap at 60 characters
                            
                            # Find the number of columns needed
                            max_numbers = max(len(line['numbers']) for line in processed_lines)
                            
                            # Print header
                            header = f"{'Description':<{max_desc_width}}"
                            for i in range(max_numbers):
                                header += f" | {'Value ' + str(i+1):>12}"
                            print("\n" + "=" * len(header))
                            print(header)
                            print("=" * len(header))
                            
                            # Print rows
                            for line in processed_lines:
                                desc = line['description'][:max_desc_width]
                                row = f"{desc:<{max_desc_width}}"
                                for num in line['numbers']:
                                    if num is None:
                                        row += f" | {'-':>12}"
                                    else:
                                        row += f" | {num:>12}"
                                for _ in range(max_numbers - len(line['numbers'])):
                                    row += f" | {' ':>12}"
                                print(row)
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
    input_pdfs_dir = Path("input_pdfs")
    input_pdfs_dir.mkdir(exist_ok=True) #if doesn't exist, create

    if len(sys.argv) > 1: 
        pdf_filename = sys.argv[1]
    else: 
        pdf_filename = input("Enter PDF filename: ")

    pdf_path = input_pdfs_dir / pdf_filename
    pdf_path = str(pdf_path.resolve())

    if not Path(pdf_path).exists():
        print(f"PDF file path not found: {pdf_path}")
        return
    
    extractor = TextExtractor()
    extractor.extract_text(pdf_path, process_numbers=True)

if __name__ == "__main__":
    main()