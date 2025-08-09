# SAP Export Zero Counts

This project processes SAP export files in `.xls` format to generate zero count reports or perform related data analysis. It is designed to run in a portable WinPython environment.

## Features

- Reads and processes SAP export `.xls` files
- Outputs zero count reports or other analytics
- Portable: runs with the included WinPython distribution

## Requirements

- Windows OS
- Python 3.12 (bundled in `src/WPy64-31241`)
- Python packages listed in `requirements.txt`

## Usage

1. Place your SAP export `.xls` file in the `src` directory.
2. Open a terminal in the project root.
3. Run the following command:
```.\src\WPy64-31241\python-3.12.4.amd64\python.exe .\src\main.py [file.xls]```
Replace `[file.xls]` with the name of your SAP export file.

## Project Structure

- `src/main.py` — Main script to process the export file
- `src/export_*.xls` — Example SAP export files
- `src/WPy64-31241/` — Portable WinPython environment
- `requirements.txt` — Python dependencies

## License

See `LICENSE` for details.