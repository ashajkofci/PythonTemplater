#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
example.py

Simple example script showing how to use the templater_core library programmatically.
This demonstrates the basic usage without the GUI.
"""
from templater_core import generate_documents, get_placeholders_from_template

# Example usage
if __name__ == "__main__":
    # Get placeholders from template
    template_file = 'template.docx'  # Your template file
    placeholders = get_placeholders_from_template(template_file)
    print(f"Template placeholders: {placeholders}")
    
    # Create field mapping
    field_mapping = {
        '{NAME}': 'customer_name',
        '{AMOUNT}': 'total_amount',
        '{DATE}': 'invoice_date'
    }
    
    # Generate documents
    files, zip_path = generate_documents(
        csv_path='data.csv',
        template_path=template_file,
        outdir='output/',
        field_mapping=field_mapping,
        filename_field='customer_name',
        filename_prefix='Document_',
        filename_suffix='_2024',
        make_zip=True
    )
    
    print(f"Generated {len(files)} documents")
    if zip_path:
        print(f"ZIP archive created: {zip_path}")
