import re

test_string = 'Total operating costs and expenses'
pattern = r'total\s+operating\s+(?:costs?\s+and\s+)?expenses?'

print(f'Testing: "{test_string}"')
print(f'Pattern: {pattern}')
print(f'Match: {bool(re.search(pattern, test_string.lower()))}')

# Test the actual case sensitivity
print()
print('Testing case sensitivity:')
print(f'Lower: {bool(re.search(pattern, test_string.lower()))}')
print(f'Exact: {bool(re.search(pattern, test_string))}')

# Test a more specific pattern
better_pattern = r'total\s+operating\s+costs\s+and\s+expenses'
print(f'Better pattern: {better_pattern}')
print(f'Better match: {bool(re.search(better_pattern, test_string.lower()))}')

# Test all patterns from the mapper
patterns = [
    r'total\s+operating\s+(?:costs?\s+and\s+)?expenses?',
    r'total\s+other\s+income',
    r'operating\s+income',
    r'income\s+before\s+taxes?',
    r'net\s+income(?:\s+attributable\s+to)?',
    r'comprehensive\s+income',
    r'net\s+income\s+attributable\s+to\s+common',
    r'less\s+loss\s+attributable'
]

print('\nTesting all exclusion patterns:')
for i, pat in enumerate(patterns):
    match = bool(re.search(pat, test_string.lower()))
    print(f'{i+1}. {pat} -> {match}') 