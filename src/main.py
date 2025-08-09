# src/main.py
# This is the main/only file of this package.
# This package takes in a xls file exported from SAP and extracts all the UPCs from the D column.
# and stores them in a list to be automaticly pasted into the SAP GUI.
import pandas as pd
import keyboard as keyboard
import sys, time

def main(filename):
    # Read the Excel file
    try:
        df = pd.read_html(filename, header=None)
    except Exception as e:
        print(f"Error reading the file: {e}")
        sys.exit(1)

    df = df[0]
    df.columns = df.iloc[0]
    df = df[1:]

    # Extract UPCs from column D (index 3)
    if df.shape[1] < 4:
        print("The DataFrame does not have enough columns to extract UPCs from column D.")
        sys.exit(1)

    # Drop rows with NaN values in the UPC column and convert to string
    upcs = df.iloc[:, 3].dropna().astype(str).tolist()
    
    # Removes all UPCs with quantity above 0
    # quantity stored in column F
    quantities = df.iloc[:, 5].dropna().astype(float).tolist()
    upcs = [upc for upc, qty in zip(upcs, quantities) if qty == 0]

    # Remove pre-specified UPCs
    try:
        with open('predefined_upcs.txt', 'r') as f:
            predefined_upcs = f.read().splitlines()
    except FileNotFoundError:
        predefined_upcs = []
    upcs = [upc for upc in upcs if upc not in predefined_upcs]

    if not upcs:
        print("No valid UPCs found.")
        sys.exit(1)

    # Print the UPCs for verification
    print("Extracted UPCs:")
    for upc in upcs:
        print(upc)
    print(f"Total UPCs extracted: {len(upcs)}")

    # Give user instructions
    print("press enter to start a 5 second cooldown before pasting UPCs into SAP GUI")
    print("Then make sure the SAP GUI is focused and ready to receive input.")
    input()
    time.sleep(5)  # Wait for 5 seconds before pasting
    # Simulate pasting UPCs into SAP GUI
    for upc in upcs:
        if keyboard.is_pressed('esc'):
            print("Operation cancelled by user.")
            sys.exit(0)
        if keyboard.is_pressed('shift'):
            paused = True
            print("Paused. Press Shift again to continue.")
            time.sleep(1)  # Give a brief pause to avoid rapid toggling
            while paused:
                if keyboard.is_pressed('shift'):
                    paused = False
        keyboard.write(upc)
        keyboard.press_and_release('enter')
        time.sleep(0.5)  # Adjust the delay as needed

if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Usage: python main.py <path_to_excel_file>")
        sys.exit(1)
    
    filename = sys.argv[1]
    main(filename)
