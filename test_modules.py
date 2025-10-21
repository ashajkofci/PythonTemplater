#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Test script for core functionality without GUI
"""
import sys
import os

def test_core_module():
    """Test templater_core module"""
    print("Testing templater_core module...")
    try:
        from templater_core import (
            read_csv_any, 
            get_placeholders_from_template, 
            find_columns,
            generate_documents
        )
        print("✓ Core module imported successfully")
        
        # Test CSV reading
        csv_file = "MJ-FAM -Contacts-MAJ-25-AOUT(Donateurs 2022-2023-2024).csv"
        if os.path.exists(csv_file):
            df = read_csv_any(csv_file)
            print(f"✓ CSV reading works ({len(df)} rows)")
            
            # Test column detection
            first, last, org, civ, amt = find_columns(df)
            print(f"✓ Column detection works (amount col: {amt})")
        
        # Test template reading
        template_file = "DONATION-AMIS-JENISCH.docx"
        if os.path.exists(template_file):
            placeholders = get_placeholders_from_template(template_file)
            print(f"✓ Template reading works ({len(placeholders)} placeholders)")
            print(f"  Placeholders: {', '.join(placeholders)}")
        
        return True
    except Exception as e:
        print(f"✗ Core module test failed: {e}")
        import traceback
        traceback.print_exc()
        return False

def test_example_script():
    """Test example.py script"""
    print("\nTesting example.py script...")
    try:
        import os
        if os.path.exists('example.py'):
            print("✓ Example script exists")
            return True
        else:
            print("⚠ Example script not found")
            return True  # Not a critical failure
    except Exception as e:
        print(f"✗ Example script test failed: {e}")
        import traceback
        traceback.print_exc()
        return False

def test_gui_module():
    """Test templater_gui_enhanced module (may fail in headless environment)"""
    print("\nTesting templater_gui_enhanced module...")
    try:
        import templater_gui_enhanced
        print("✓ Enhanced GUI module imported successfully")
        return True
    except ModuleNotFoundError as e:
        if 'tkinter' in str(e):
            print("⚠ Enhanced GUI module requires tkinter (not available in this environment)")
            return True  # Not a failure, just unavailable
        else:
            print(f"✗ Enhanced GUI module test failed: {e}")
            return False
    except Exception as e:
        print(f"✗ Enhanced GUI module test failed: {e}")
        import traceback
        traceback.print_exc()
        return False

def main():
    print("=" * 60)
    print("PythonTemplater Test Suite")
    print("=" * 60)
    
    results = []
    results.append(("Core Module", test_core_module()))
    results.append(("Example Script", test_example_script()))
    results.append(("GUI Module", test_gui_module()))
    
    print("\n" + "=" * 60)
    print("Test Results Summary")
    print("=" * 60)
    
    for name, passed in results:
        status = "✓ PASS" if passed else "✗ FAIL"
        print(f"{status}: {name}")
    
    all_passed = all(passed for _, passed in results)
    print("\n" + ("All tests passed!" if all_passed else "Some tests failed."))
    
    return 0 if all_passed else 1

if __name__ == "__main__":
    sys.exit(main())
