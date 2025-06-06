import logging
from pathlib import Path
import pdfplumber
import pandas as pd
from find_fs import FinancialStatementFinder
import os
import re

class TableExtractor:
    def __init__(self):
        self.setup_logging()
        self.finder = FinancialStatementFinder()
        
    def setup_logging(self):
        """Setup logging configuration"""
        logging.basicConfig(
            level=logging.INFO,
            format='%(message)s',
            handlers=[
                logging.FileHandler('table_extraction.log', mode='w'),
                logging.StreamHandler()
            ]
        )
        
        # Create a custom logger for line parsing
        self.line_logger = logging.getLogger('line_parser')
        self.line_logger.setLevel(logging.INFO)
        
        # Remove all handlers
        self.line_logger.handlers = []
        
        # Add handlers with custom format
        formatter = logging.Formatter('%(message)s')
        stream_handler = logging.StreamHandler()
        stream_handler.setFormatter(formatter)
        self.line_logger.addHandler(stream_handler)
        
        file_handler = logging.FileHandler('line_parsing.log', mode='w')
        file_handler.setFormatter(formatter)
        self.line_logger.addHandler(file_handler)
    
    def detect_number_pattern(self, text_lines):
        """Detect how many numbers appear at the end of each line"""
        # Updated pattern to include standalone dash
        number_pattern = r'(?:\$\s*)?(?:\(\s*[\d,]+\s*\)|[\d,]+|\s*-\s*)(?:\s+(?:\$\s*)?(?:\(\s*[\d,]+\s*\)|[\d,]+|\s*-\s*))*$'
        max_numbers = 0
        
        # First pass - look for year headers
        years = []
        for line in text_lines:
            if not line.strip():
                continue
            if any(str(year) in line for year in range(2000, 2050)):
                numbers = re.findall(r'\b(20\d{2})\b', line)
                if numbers:
                    years = numbers
                    max_numbers = len(numbers)
                    break
        
        # If no year header found, count maximum numbers in lines
        if not years:
            for line in text_lines:
                if not line.strip():
                    continue
                    
                match = re.search(number_pattern, line.strip())
                if match:
                    # Count both numbers and dash placeholders
                    numbers = re.findall(r'\$?\s*(\(\s*[\d,]+\s*\)|[\d,]+|\s*-\s*)', match.group(0))
                    max_numbers = max(max_numbers, len(numbers))
        
        return max_numbers

    def parse_financial_line(self, line, num_years):
        """Parse a line into description and numbers"""
        if not line.strip():
            return line.strip(), []
            
        # Find the last occurrence of numbers in the line
        number_pattern = r'(?:\$\s*)?(?:\(\s*[\d,]+\s*\)|[\d,]+|\s*-\s*)(?:\s+(?:\$\s*)?(?:\(\s*[\d,]+\s*\)|[\d,]+|\s*-\s*))*$'
        match = re.search(number_pattern, line.strip())
        
        if not match:
            return line.strip(), []
            
        # Split into description and numbers
        numbers_str = match.group(0)
        description = line[:match.start()].strip()
        
        # Extract individual numbers, preserving parentheses, removing $, and keeping dashes
        numbers = re.findall(r'\$?\s*(\(\s*[\d,]+\s*\)|[\d,]+|\s*-\s*)', numbers_str)
        numbers = [n.strip('$ ') for n in numbers]
        
        # Replace standalone dashes with empty strings for Excel
        numbers = ['' if n.strip() == '-' else n for n in numbers]
        
        # Pad with empty strings if we don't have enough numbers
        while len(numbers) < num_years:
            numbers.append('')
            
        return description, numbers[:num_years]

    def extract_tables(self, pdf_path):
        """Extract tables from PDF file"""
        try:
            # Find financial statements
            finder = FinancialStatementFinder()
            statements = finder.find_statements(pdf_path)
            
            logging.info(f"Found statements: {statements}")
            
            # Get pages to process
            statement_pages = {
                'balance_sheet': [],
                'income_statement': [],
                'cash_flow': []
            }
            
            pages_to_process = set()
            for stmt_type, info in statements.items():
                # Handle both single pages and page ranges
                if info['pdf_page']:
                    pages = [info['pdf_page']]
                    # Add continuation pages if they exist
                    if 'continuation_pages' in info:
                        pages.extend(info['continuation_pages'])
                    statement_pages[stmt_type].extend([p-1 for p in pages])  # Convert to 0-based
                    pages_to_process.update(p-1 for p in pages)
            
            logging.info(f"Pages to process: {pages_to_process}")
            logging.info(f"Statement pages mapping: {statement_pages}")
            
            # Initialize results - store with 1-based page numbers
            tables_by_page = {page_num + 1: [] for page_num in pages_to_process}
            
            # Process each page
            with pdfplumber.open(pdf_path) as pdf:
                for page_num in sorted(pages_to_process):
                    if page_num >= len(pdf.pages):
                        continue
                        
                    page = pdf.pages[page_num]
                    
                    # Extract text and split into lines
                    if page.extract_text():
                        text_lines = [line for line in page.extract_text().split('\n') if line.strip()]
                        # Remove duplicate lines
                        text_lines = list(dict.fromkeys(text_lines))
                        
                        # First detect the number pattern
                        num_years = self.detect_number_pattern(text_lines)
                        logging.info(f"Page {page_num + 1}: Detected {num_years} years")
                        
                        # Log the page header
                        self.line_logger.info(f"\n=== Page {page_num + 1} (Years: {num_years}) ===\n")
                        
                        # Store data for Excel
                        data = []
                        years = []
                        
                        # Process each line
                        for line in text_lines:
                            if line.strip():
                                description, numbers = self.parse_financial_line(line, num_years)
                                
                                # Log the line
                                if numbers:
                                    numbers_str = '  '.join(str(n) for n in numbers)
                                    self.line_logger.info(f"{description:<50} {numbers_str}\n")
                                    
                                    # Try to detect if this is a header row with years
                                    if any(str(year) in description for year in range(2000, 2050)):
                                        years = numbers
                                    else:
                                        # Store the data for Excel
                                        data.append([description] + numbers)
                                else:
                                    self.line_logger.info(f"{line.strip()}\n")
                                    # Store lines without numbers as headers or notes
                                    if not any(str(year) in line for year in range(2000, 2050)):
                                        data.append([line.strip()] + [''] * num_years)
                        
                        # Create DataFrame
                        if data:
                            logging.info(f"Page {page_num + 1}: Creating DataFrame with {len(data)} rows")
                            
                            # Create column headers
                            if years and len(years) == num_years:
                                columns = ['Description'] + years
                            else:
                                columns = ['Description'] + [f'Year_{i+1}' for i in range(num_years)]
                            
                            # Create DataFrame with proper column names
                            df = pd.DataFrame(data, columns=columns)
                            
                            # Determine statement type for this page
                            stmt_type = None
                            for s_type, pages in statement_pages.items():
                                if page_num in pages:
                                    stmt_type = s_type
                                    break
                            
                            logging.info(f"Page {page_num + 1}: Statement type: {stmt_type}")
                            
                            # Store the DataFrame
                            tables_by_page[page_num + 1].append({
                                'type': stmt_type,
                                'data': df,
                                'years': years if years else []
                            })
                            
            return tables_by_page, statement_pages
            
        except Exception as e:
            logging.error(f"Error extracting tables: {str(e)}")
            return {}, {}
    
    def save_tables(self, tables_by_page, statement_pages, output_path):
        """Save extracted tables to Excel with each table on a separate sheet"""
        if not tables_by_page:
            logging.warning("No tables to save")
            return False
            
        try:
            logging.info("Starting to save tables to Excel")
            
            # Create a default DataFrame in case no tables are found
            default_df = pd.DataFrame({'Message': ['No tables were successfully extracted']})
            
            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                sheets_written = False
                
                # Process each statement type
                for stmt_type in ['balance_sheet', 'income_statement', 'cash_flow']:
                    all_dfs = []
                    
                    # Get all pages for this statement type
                    stmt_pages = statement_pages.get(stmt_type, [])
                    if not stmt_pages:
                        continue
                        
                    # Convert to 1-based page numbers
                    pages = [p + 1 for p in stmt_pages]
                    
                    # Collect all DataFrames for this statement type
                    for page_num in pages:
                        if page_num in tables_by_page:
                            for table in tables_by_page[page_num]:
                                if (isinstance(table, dict) and 
                                    table.get('type') == stmt_type and 
                                    isinstance(table.get('data'), pd.DataFrame) and 
                                    not table['data'].empty):
                                    all_dfs.append(table['data'])
                    
                    if all_dfs:
                        # Combine all DataFrames for this statement type
                        combined_df = pd.concat(all_dfs, ignore_index=True)
                        
                        # Clean up the combined DataFrame
                        combined_df = combined_df.dropna(how='all')
                        
                        if not combined_df.empty:
                            # Format sheet name
                            sheet_name = stmt_type.replace('_', ' ').title()
                            if len(sheet_name) > 31:
                                sheet_name = sheet_name[:31]
                            
                            # Write to Excel
                            combined_df.to_excel(writer, sheet_name=sheet_name, index=False)
                            sheets_written = True
                            
                            # Format the worksheet
                            worksheet = writer.sheets[sheet_name]
                            
                            # Set column widths
                            worksheet.column_dimensions['A'].width = 50  # Description column
                            for col in range(1, len(combined_df.columns)):
                                worksheet.column_dimensions[chr(65 + col)].width = 15
                
                # If no sheets were written, write the default sheet
                if not sheets_written:
                    default_df.to_excel(writer, sheet_name='No_Tables_Found', index=False)
            
            logging.info(f"Successfully saved tables to {output_path}")
            return True
            
        except Exception as e:
            logging.error(f"Error saving tables to Excel: {str(e)}")
            return False

def main():
    # Create input_pdfs directory if it doesn't exist
    input_pdfs_dir = Path("input_pdfs")
    input_pdfs_dir.mkdir(exist_ok=True)
    
    # Create output directory
    output_dir = Path("extracted_tables")
    output_dir.mkdir(exist_ok=True)
    
    # Get PDF filename from command line arguments or user input
    import sys
    if len(sys.argv) > 1:
        pdf_filename = sys.argv[1]
    else:
        pdf_filename = input("Enter the PDF filename (e.g. AME_2019.pdf): ")
    
    # Construct full path
    pdf_path = input_pdfs_dir / pdf_filename
    pdf_path = str(pdf_path.resolve())
    
    if not Path(pdf_path).exists():
        print(f"PDF file not found: {pdf_path}")
        print(f"Please make sure the file exists in the {input_pdfs_dir} directory")
        return
    
    # Extract tables
    extractor = TableExtractor()
    tables_by_page, statement_pages = extractor.extract_tables(pdf_path)
    
    # Save tables to Excel
    if tables_by_page:
        output_path = os.path.join(output_dir, f"{os.path.splitext(os.path.basename(pdf_path))[0]}_tables.xlsx")
        extractor.save_tables(tables_by_page, statement_pages, output_path)
        print(f"\nExtraction completed! Check {output_path} for the extracted tables")
    else:
        print("\nNo tables were extracted. Check table_extraction.log for details")

if __name__ == "__main__":
    main() 