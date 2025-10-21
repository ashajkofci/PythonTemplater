#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Simple launcher script for the CSV Templater GUI application
"""
import sys

try:
    from templater_gui import main
    sys.exit(main())
except ImportError as e:
    print("Error: Required dependencies not found.")
    print("\nPlease install dependencies:")
    print("  pip install -r requirements.txt")
    print(f"\nDetails: {e}")
    sys.exit(1)
except Exception as e:
    print(f"Error launching application: {e}")
    import traceback
    traceback.print_exc()
    sys.exit(1)
