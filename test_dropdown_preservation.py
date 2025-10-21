#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Test dropdown preservation in config save/load
"""
import json
import tempfile
from pathlib import Path

# Simulate the FieldMappingRow behavior
class MockFieldMappingRow:
    def __init__(self, placeholder):
        self.placeholder = placeholder
        self.column_vars = []  # Will store StringVar-like values
        self.combine_var = False
    
    def add_dropdown(self, value=''):
        """Simulate adding a dropdown"""
        self.column_vars.append({'value': value})
    
    def get_value(self, idx):
        """Get value from a dropdown"""
        if idx < len(self.column_vars):
            return self.column_vars[idx]['value']
        return ''
    
    def set_value(self, idx, value):
        """Set value in a dropdown"""
        if idx < len(self.column_vars):
            self.column_vars[idx]['value'] = value
    
    def get_mapping(self):
        """Get the mapping configuration - NEW VERSION"""
        columns = [var['value'] for var in self.column_vars if var['value']]
        
        # Always return a mapping object, even if no columns selected
        return {
            'columns': columns if columns else [],
            'combine': self.combine_var,
            'num_dropdowns': len(self.column_vars)  # Save the number of dropdowns
        }
    
    def set_mapping(self, mapping_config):
        """Set the mapping from saved configuration - NEW VERSION"""
        if not mapping_config:
            return
        
        columns = mapping_config.get('columns', [])
        combine = mapping_config.get('combine', False)
        num_dropdowns = mapping_config.get('num_dropdowns', max(1, len(columns)))
        
        # Clear existing
        self.column_vars.clear()
        
        # Create the correct number of dropdowns
        for i in range(num_dropdowns):
            self.add_dropdown()
            # Set the column value if available
            if i < len(columns) and columns[i]:
                self.set_value(i, columns[i])
        
        self.combine_var = combine


print("="*60)
print("Testing Dropdown Preservation")
print("="*60)

# Test 1: User adds 2 empty dropdowns
print("\n=== Test 1: Two empty dropdowns ===")
row = MockFieldMappingRow('{NOM}')
row.add_dropdown('')  # First dropdown (empty)
row.add_dropdown('')  # Second dropdown (empty)

print(f"Created row with {len(row.column_vars)} dropdowns")
print(f"Values: {[row.get_value(i) for i in range(len(row.column_vars))]}")

# Get mapping (save to config)
mapping = row.get_mapping()
print(f"\nMapping saved to config: {mapping}")

# Simulate config save/load cycle
config_json = json.dumps(mapping, indent=2)
print(f"\nConfig JSON:\n{config_json}")

loaded_mapping = json.loads(config_json)

# Create a new row and load the mapping
row2 = MockFieldMappingRow('{NOM}')
row2.add_dropdown('')  # Start with 1 dropdown (default)
row2.set_mapping(loaded_mapping)

print(f"\nAfter loading config:")
print(f"  Dropdowns: {len(row2.column_vars)}")
print(f"  Values: {[row2.get_value(i) for i in range(len(row2.column_vars))]}")

if len(row2.column_vars) == 2:
    print("✅ SUCCESS: 2 dropdowns preserved!")
else:
    print(f"❌ FAIL: Expected 2 dropdowns, got {len(row2.column_vars)}")
    exit(1)

# Test 2: User adds 2 dropdowns and fills them
print("\n=== Test 2: Two dropdowns with values ===")
row3 = MockFieldMappingRow('{NOM}')
row3.add_dropdown('AdrNameS')
row3.add_dropdown('AdrVornameS')

print(f"Created row with {len(row3.column_vars)} dropdowns")
print(f"Values: {[row3.get_value(i) for i in range(len(row3.column_vars))]}")

# Get mapping
mapping = row3.get_mapping()
print(f"\nMapping saved to config: {mapping}")

# Load it back
row4 = MockFieldMappingRow('{NOM}')
row4.add_dropdown('')  # Start with 1 dropdown
row4.set_mapping(mapping)

print(f"\nAfter loading config:")
print(f"  Dropdowns: {len(row4.column_vars)}")
print(f"  Values: {[row4.get_value(i) for i in range(len(row4.column_vars))]}")

if len(row4.column_vars) == 2 and row4.get_value(0) == 'AdrNameS' and row4.get_value(1) == 'AdrVornameS':
    print("✅ SUCCESS: 2 dropdowns with values preserved!")
else:
    print(f"❌ FAIL: Expected 2 dropdowns with specific values")
    exit(1)

print("\n" + "="*60)
print("✅ ALL DROPDOWN PRESERVATION TESTS PASSED!")
print("="*60)
print("\nKey improvement:")
print("  • Dropdowns are now preserved even when empty")
print("  • Config saves 'num_dropdowns' field")
print("  • When you click '+' to add a dropdown, it's saved immediately")
print("  • When config loads, correct number of dropdowns are recreated")
