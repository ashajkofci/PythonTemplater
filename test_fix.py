#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Test script to verify the fix for "0 documents generated" issue
"""
import os
import sys
import tempfile
import shutil
from templater_core import generate_documents, get_placeholders_from_template

# Use the actual CSV and template from the project
csv_path = 'MJ-FAM -Contacts-MAJ-25-AOUT(Donateurs 2022-2023-2024).csv'
template_path = 'DONATION-AMIS-JENISCH.docx'

# Create a temp output directory
outdir = tempfile.mkdtemp(prefix='templater_test_')

try:
    # Get placeholders
    placeholders = get_placeholders_from_template(template_path)
    print(f"✓ Template placeholders: {placeholders}")
    
    # Test 1: Map only ONE placeholder (MONTANT), leave NOM unmapped
    print("\n=== Test 1: One mapped field, others unmapped ===")
    field_mapping = {
        '{DATE}': '',  # Empty
        '{MONTANT}': 'Montant don 2025',
        '{NOM}': ''  # Explicitly empty - simulates unmapped placeholder
    }
    
    print(f"Field mapping: {field_mapping}")
    
    files, zip_path = generate_documents(
        csv_path=csv_path,
        template_path=template_path,
        outdir=outdir,
        field_mapping=field_mapping,
        make_zip=False
    )
    
    print(f"\n✓ Test 1 Results:")
    print(f"  - Generated {len(files)} documents")
    print(f"  - First few files: {[os.path.basename(f) for f in files[:3]]}")
    
    if len(files) > 0:
        print("\n✅ SUCCESS: Documents were generated even with unmapped placeholder!")
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
        csv_path=csv_path,
        template_path=template_path,
        outdir=outdir,
        field_mapping=field_mapping,
        make_zip=False
    )
    
    print(f"\n✓ Test 2 Results:")
    print(f"  - Generated {len(files)} documents")
    print(f"  - First few files: {[os.path.basename(f) for f in files[:3]]}")
    
    if len(files) > 0:
        print("\n✅ SUCCESS: Documents generated with both fields mapped!")
    else:
        print("\n❌ FAIL: No documents generated")
        sys.exit(1)
    
    print("\n" + "="*60)
    print("✅ ALL TESTS PASSED!")
    print("="*60)
    
finally:
    # Clean up
    if os.path.exists(outdir):
        shutil.rmtree(outdir)
        print(f"\nCleaned up temp directory: {outdir}")
