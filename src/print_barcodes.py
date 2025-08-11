#!/usr/bin/env python3
"""
Zero Export - Python version
Generates a PDF with Code 39 barcodes for products with zero quantity
"""

import pandas as pd
import sys
import os
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch
from reportlab.graphics.barcode import code39
from reportlab.graphics import renderPDF
from reportlab.graphics.shapes import Drawing

def get_zero_quantity_products(filename):
    # Read the Excel file as HTML table (SAP exports are often HTML-based .xls)
    try:
        df = pd.read_html(filename, header=None)
    except Exception as e:
        print(f"Error reading the file: {e}")
        sys.exit(1)

    df = df[0]  # Use the first table found
    df.columns = df.iloc[0]  # Set first row as header
    df = df[1:]  # Remove header row from data

    # Ensure there are enough columns to extract UPCs from column D (index 3)
    if df.shape[1] < 4:
        print("The DataFrame does not have enough columns to extract UPCs from column D.")
        sys.exit(1)

    # Extract UPCs from column D, drop empty values, convert to string
    upcs = df.iloc[:, 3].dropna().astype(str).tolist()
    
    # Extract quantities from column F, drop empty values, convert to float
    quantities = df.iloc[:, 5].dropna().astype(float).tolist()
    # Keep only UPCs where the corresponding quantity is zero
    upcs = [upc for upc, qty in zip(upcs, quantities) if qty == 0]

    # Remove UPCs listed in predefined_upcs.txt (if file exists)
    try:
        with open('predefined_upcs.txt', 'r') as f:
            predefined_upcs = f.read().splitlines()
    except FileNotFoundError:
        predefined_upcs = []
    upcs = [upc for upc in upcs if upc not in predefined_upcs]

    if not upcs:
        print("No valid UPCs found.")
        sys.exit(1)

    # Print extracted UPCs for user verification
    print("Extracted UPCs:")
    print(f"Total UPCs extracted: {len(upcs)}")
    return upcs



def create_barcode_pdf(upcs, output_file="zero_export.pdf"):
    """Create a PDF with barcodes arranged in a 3x10 grid"""
    if not upcs:
        print("No UPCs to export")
        return False
    
    # PDF setup
    c = canvas.Canvas(output_file, pagesize=letter)
    page_width, page_height = letter
    
    # Layout configuration
    cols = 3
    rows = 10
    barcode_width = 2.5 * inch
    barcode_height = 0.8 * inch
    margin_x = 0.5 * inch
    margin_y = 0.5 * inch
    
    # Calculate spacing
    available_width = page_width - (2 * margin_x)
    available_height = page_height - (2 * margin_y)
    
    col_spacing = available_width / cols
    row_spacing = available_height / rows
    
    for i, upc in enumerate(upcs):
        # Start new page if needed
        if i > 0 and i % (cols * rows) == 0:
            c.showPage()
        
        # Calculate position on current page
        index_on_page = i % (cols * rows)
        col = index_on_page % cols
        row = index_on_page // cols
        
        x = margin_x + (col * col_spacing)
        y = page_height - margin_y - ((row + 1) * row_spacing)
        
        try:
            # Create barcode (as Flowable) with thicker bars
            barcode = code39.Standard39(
                str(upc),
                checksum=0,
                barWidth=0.8,  # Increase this value for thicker bars (default is 0.5)
                barHeight=barcode_height * 0.9
            )
            barcode.width = barcode_width * 0.7
            barcode.height = barcode_height * 0.8
            # Draw barcode directly on canvas
            barcode_x = x + (barcode_width - barcode.width) / 2
            barcode_y = y + (barcode_height - barcode.height) / 2
            barcode.drawOn(c, barcode_x, barcode_y)
            
            # Add UPC text below barcode
            c.setFont("Helvetica", 10)
            text_width = c.stringWidth(str(upc), "Helvetica", 10)
            text_x = x + (barcode_width - text_width) / 2
            text_y = y - 10
            c.drawString(text_x, text_y, str(upc))
            
        except Exception as e:
            print(f"Error creating barcode for UPC {upc}: {e}")
            continue
    
    c.save()
    return True

def main():
    """Main function"""
    # Check command line arguments
    if len(sys.argv) != 2:
        print("Usage: python print_barcodes.py <path_to_excel_file>")
        sys.exit(1)
    filename = sys.argv[1]
    
    if not os.path.exists(filename):
        print(f"File not found: {filename}")
        sys.exit(1)
    # Check if the file is an Excel file
    if not filename.lower().endswith(('.xls', '.xlsx')):
        print("The provided file is not an Excel file.")
        sys.exit(1)
    
    # Get products with zero quantity
    print("Fetching products with zero quantity...")
    upcs = get_zero_quantity_products(filename)
    
    if not upcs:
        print("No products with zero quantity found.")
        sys.exit(0)
    
    print(f"Found {len(upcs)} products with zero quantity")
    
    # Generate output filename
    output_file = f"zero_export_{len(upcs)}_items.pdf"
    
    # Create PDF
    print(f"Generating PDF: {output_file}")
    if create_barcode_pdf(upcs, output_file):
        print(f"PDF created successfully: {output_file}")
    else:
        print("Failed to create PDF")
        sys.exit(1)

if __name__ == "__main__":
    main()