# src/main.py
# This is the main/only file of this package.
# This package takes in a xls file exported from SAP and extracts all the UPCs from the D column.
# and stores them in a list to be automaticly pasted into the SAP GUI.
import packages.pandas as pd
import packages.keyboard as keyboard
import sys, time

def main(filename):
    # Read the Excel file
    try:
        df = pd.read_excel(filename, header=None)
    except Exception as e:
        print(f"Error reading the file: {e}")
        sys.exit(1)

    # Extract UPCs from column D (index 3)
    upcs = df[3].dropna().astype(str).tolist()

    # Print the UPCs for verification
    print("Extracted UPCs:")
    for upc in upcs:
        print(upc)

    # Give user instructions
    print("press enter to start a 5 second cooldown before pasting UPCs into SAP GUI")
    print("Then make sure the SAP GUI is focused and ready to receive input.")
    input()
    # Simulate pasting UPCs into SAP GUI
    for upc in upcs:
        keyboard.write(upc)
        keyboard.press_and_release('enter')
        time.sleep(0.5)  # Adjust the delay as needed

if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Usage: python main.py <path_to_excel_file>")
        sys.exit(1)
    
    filename = sys.argv[1]
    main(filename)
