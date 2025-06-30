from final_is_mapper import FinalISMapper

mapper = FinalISMapper()

# Test the problematic line
test_line = "Total operating costs and expenses"

print(f"Testing: '{test_line}'")
print()

# Test 1: is_total_or_net_row
is_filtered = mapper.is_total_or_net_row(test_line)
print(f"1. is_total_or_net_row: {is_filtered}")

# Test 2: apply_enhanced_mapping 
template_field, section, confidence = mapper.apply_enhanced_mapping(test_line)
print(f"2. apply_enhanced_mapping: {template_field}, {section}, {confidence}")

# Test other problematic lines
other_lines = [
    "OPERATING INCOME",
    "NET INCOME", 
    "Total other income (expense)—net"
]

print("\nTesting other calculated lines:")
for line in other_lines:
    is_filtered = mapper.is_total_or_net_row(line)
    template_field, section, confidence = mapper.apply_enhanced_mapping(line)
    print(f"'{line}': filtered={is_filtered}, mapped=({template_field}, {section}, {confidence})") 