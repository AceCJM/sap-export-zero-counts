import pandas as pd
import os, sys

# debugging function for saving upc and quantity to text files
def save_upcs_and_quantities(upcs):
    with open('extracted_upcs.txt', 'w') as f:
        for upc in upcs:
            f.write(f"{upc}\n")


def main(filename):
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

    # Extract UPCs and quantities together, drop rows with missing values in either
    upc_qty_df = df.iloc[:, [3, 5]].dropna()
    upcs = upc_qty_df.iloc[:, 0].astype(str).tolist()
    quantities = upc_qty_df.iloc[:, 1].astype(float).tolist()

    # Ensure UPCs and quantities lists are of the same length
    if len(upcs) != len(quantities):
        # quantities len will be greater than upcs len
        print("Warning: Mismatched lengths between UPCs and quantities. Adjusting quantities list.")
        quantities = quantities[:len(upcs)]
        print(len(upcs), len(quantities))
    # Keep only UPCs where the corresponding quantity is zero
    upcs_debug = [upc for upc, qty in zip(upcs, quantities) if qty == 0]
    save_upcs_and_quantities(upcs_debug)
    print("Debugging: Extracted UPCs and quantities saved to extracted_upcs.txt")

    # Check if any UPCs are not 0 quantity in the original list
    for upc, upcs_debug, qty, in zip(upcs, upcs_debug, quantities):
        if upcs_debug == upc and qty != 0:
            print(f"Debugging Warning: UPC {upc} has quantity {qty}, but was included in zero-quantity list.")

if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Usage: python debugging.py <filename>")
        sys.exit(1)

    filename = sys.argv[1]
    if not os.path.isfile(filename):
        print(f"File not found: {filename}")
        sys.exit(1)
    main(filename)