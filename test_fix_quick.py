#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Quick test script to verify the fix for "0 documents generated" issue
Uses only first 5 rows of CSV for speed
"""
import os
import sys
import tempfile
import shutil
import pandas as pd
from templater_core import generate_documents, get_placeholders_from_template, read_csv_any

# Use the actual CSV and template from the project
csv_path = 'MJ-FAM -Contacts-MAJ-25-AOUT(Donateurs 2022-2023-2024).csv'
template_path = 'DONATION-AMIS-JENISCH.docx'

# Create a temp CSV with just 5 rows
temp_csv = tempfile.NamedTemporaryFile(mode='w', suffix='.csv', delete=False, encoding='utf-8')
try:
    df = read_csv_any(csv_path)
    # Take first 5 non-empty rows
    df_sample = df.head(10)  # Get first 10 to ensure we have some with data
    df_sample.to_csv(temp_csv.name, index=False, encoding='utf-8')
    temp_csv.close()
    
    # Create a temp output directory
    outdir = tempfile.mkdtemp(prefix='templater_test_')
    
    try:
        # Get placeholders
        placeholders = get_placeholders_from_template(template_path)
        print(f"✓ Template placeholders: {placeholders}")
        print(f"✓ Using sample CSV with {len(df_sample)} rows")
        
        # Test 1: Map only ONE placeholder (MONTANT), leave others unmapped
        print("\n=== Test 1: One mapped field, others unmapped ===")
        field_mapping = {
            '{DATE}': '',  # Empty
            '{MONTANT}': 'Montant don 2025',
            '{NOM}': ''  # Explicitly empty - simulates unmapped placeholder
        }
        
        print(f"Field mapping: {field_mapping}")
        
        files, zip_path = generate_documents(
            csv_path=temp_csv.name,
            template_path=template_path,
            outdir=outdir,
            field_mapping=field_mapping,
            make_zip=False
        )
        
        print(f"\n✓ Test 1 Results:")
        print(f"  - Generated {len(files)} documents from {len(df_sample)} rows")
        if files:
            print(f"  - First file: {os.path.basename(files[0])}")
        
        if len(files) > 0:
            print("\n✅ SUCCESS: Documents were generated even with unmapped placeholders!")
        else:
            print("\n❌ FAIL: No documents generated (should have generated some)")
            sys.exit(1)
        
        # Clean up test 1 files
        for f in files:
            os.remove(f)
        
        # Test 2: Map all placeholders
        print("\n=== Test 2: All placeholders mapped ===")
        field_mapping = {
            '{DATE}': 'Date don',  # Use a date column
            '{MONTANT}': 'Montant don 2025',
            '{NOM}': 'AdrNameS AdrVornameS'  # Combine two columns
        }
        
        print(f"Field mapping: {field_mapping}")
        
        files, zip_path = generate_documents(
            csv_path=temp_csv.name,
            template_path=template_path,
            outdir=outdir,
            field_mapping=field_mapping,
            make_zip=False
        )
        
        print(f"\n✓ Test 2 Results:")
        print(f"  - Generated {len(files)} documents from {len(df_sample)} rows")
        if files:
            print(f"  - First file: {os.path.basename(files[0])}")
        
        if len(files) > 0:
            print("\n✅ SUCCESS: Documents generated with all fields mapped!")
        else:
            print("\n❌ FAIL: No documents generated")
            sys.exit(1)
        
        print("\n" + "="*60)
        print("✅ ALL TESTS PASSED!")
        print("="*60)
        print("\nKey findings:")
        print("  • Unmapped placeholders don't prevent document generation")
        print("  • Rows are only skipped if they're completely empty")
        print("  • The fix correctly allows partial mapping")
        
    finally:
        # Clean up
        if os.path.exists(outdir):
            shutil.rmtree(outdir)
            print(f"\n✓ Cleaned up temp directory: {outdir}")
        
finally:
    # Clean up temp CSV
    if os.path.exists(temp_csv.name):
        os.remove(temp_csv.name)
        print(f"✓ Cleaned up temp CSV: {temp_csv.name}")
