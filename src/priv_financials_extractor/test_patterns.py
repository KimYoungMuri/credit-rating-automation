import re

# Test the patterns
test_cases = [
    'OPERATING ACTIVITIES: Net income $ $',
    'Adjustments to reconcile net income to net ca sh provided',
    'Proceeds on sale of property and equi pment a nd other assets',
    'by operati ng activi ties: Depreci a tion a nd amortization'
]

patterns = [
    (r'operating\s+activities:\s*net\s+income', 'Net Income'),
    (r'adjustments\s+to\s+reconcile\s+net\s+income', 'Net Income'),
    (r'proceeds\s+on\s+sale\s+of\s+property\s+and\s+equipment\s+and\s+other\s+assets', 'Proceeds from asset sales'),
    (r'depreciation\s+and\s+amortization', 'Changes in noncash items')
]

for test_desc in test_cases:
    desc_lower = test_desc.lower()
    print(f'\nTesting: "{test_desc}"')
    print(f'Lower: "{desc_lower}"')
    
    for pattern, field in patterns:
        match = re.search(pattern, desc_lower)
        if match:
            print(f'  ✅ PATTERN MATCHED: "{pattern}" → {field}')
        else:
            print(f'  ❌ NO MATCH: "{pattern}"') 