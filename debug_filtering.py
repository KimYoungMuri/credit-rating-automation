# Debug script to check filtering logic
from final_extractor import TextExtractor
from final_find_fs import FinancialStatementFinder

extractor = TextExtractor()

confirmed_pages = {
    'income_statement': [9],
}

statement_pages_dict = {}
for stmt_type, pages in confirmed_pages.items():
    if pages:
        statement_pages_dict[stmt_type] = pages

excel_path, extracted_data = extractor.extract_text(
    '../../input_pdfs/US_Venture_2024.pdf', 
    process_numbers=True, 
    statement_pages=statement_pages_dict
)

income_statement_data = extracted_data['income_statement']

print('ACTUAL EXTRACTED DESCRIPTIONS:')
for i, item in enumerate(income_statement_data):
    desc = item.get('description', '').strip()
    if desc:
        print(f'{i:2d}: "{desc}"')
        if 'total' in desc.lower() or 'operating income' in desc.lower() or 'net income' in desc.lower():
            print(f'    --> Should be filtered: {desc.lower()}')

print('\n\nTesting specific problematic line:')
problem_line = "Total operating costs and expenses"
print(f'Problem line: "{problem_line}"')

# Test the calculated rows check
calculated_rows = [
    'total operating costs and expenses',
    'operating income',
    'total other income',
    'income before taxes',
    'net income',
    'net income attributable',
    'comprehensive income',
    'less loss attributable'
]

desc_lower = problem_line.lower().strip()
for calc_row in calculated_rows:
    if calc_row in desc_lower:
        print(f'✅ SHOULD BE FILTERED - matches: "{calc_row}"')
        break
else:
    print('❌ NOT being filtered by calculated_rows check') 