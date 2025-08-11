# src/main.py
# Main script for extracting zero-quantity UPCs from SAP-exported .xls files
# and automating their entry into the SAP GUI.

import pandas as pd
import keyboard
import sys, time

def enter_upcs(upc):
    """
    Simulate typing a UPC and pressing Enter twice.
    Used for standard UPC entry in SAP GUI.
    """
    keyboard.write(upc)
    time.sleep(1)
    keyboard.press_and_release('enter')
    time.sleep(1)
    keyboard.press_and_release('enter')

def layout_upc(upc):
    """
    Simulate typing a UPC and pressing Enter three times.
    Used for UPCs that require extra confirmation/layout steps in SAP GUI.
    """
    keyboard.write(upc)
    time.sleep(1)
    keyboard.press_and_release('enter')
    time.sleep(1)
    keyboard.press_and_release('enter')
    time.sleep(1)
    keyboard.press_and_release('enter')

def main(filename):
    """
    Main logic:
    - Reads the SAP-exported .xls file.
    - Extracts UPCs from column D where quantity in column F is zero.
    - Removes UPCs listed in predefined_upcs.txt.
    - Handles special UPCs from multi_layout_upcs.txt.
    - Automates pasting UPCs into SAP GUI.
    """
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

    # Load UPCs that require special layout handling (if file exists)
    try:
        with open('multi_layout_upcs.txt', 'r') as f:
            multi_layout_upcs = f.read().splitlines()
    except FileNotFoundError:
        multi_layout_upcs = []

    # Print extracted UPCs for user verification
    print("Extracted UPCs:")
    for upc in upcs:
        print(upc)
    print(f"Total UPCs extracted: {len(upcs)}")

    # User instructions before automation starts
    print("press enter to start a 5 second cooldown before pasting UPCs into SAP GUI")
    print("Then make sure the SAP GUI is focused and ready to receive input.")
    input()
    time.sleep(5)  # Wait for user to focus SAP GUI

    # Automate pasting UPCs into SAP GUI
    for upc in upcs:
        # Allow user to cancel with ESC
        if keyboard.is_pressed('esc'):
            print("Operation cancelled by user.")
            sys.exit(0)
        # Allow user to pause/resume with Shift
        if keyboard.is_pressed('shift'):
            paused = True
            print("Paused. Press Shift again to continue.")
            time.sleep(1)  # Prevent rapid toggling
            while paused:
                if keyboard.is_pressed('shift'):
                    paused = False
        # Use special layout handling if needed
        if upc in multi_layout_upcs:
            layout_upc(upc)
        else:
            enter_upcs(upc)
        time.sleep(1)  # Delay between UPCs for reliability

if __name__ == "__main__":
    # Check for correct usage
    if len(sys.argv) != 2:
        print("Usage: python main.py <path_to_excel_file>")
        sys.exit(1)
    
    filename = sys.argv[1]
    main(filename)
