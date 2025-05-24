## requirements.txt:

subprocess
re
os
shutil
qrcode
pandas
docx
xlsxwriter
time
concurrent.futures

## README.md:

# Vanity Address Generator

## Overview

This project automates the generation of Bitcoin vanity addresses using `vanitygen.exe`. It extracts the generated addresses and private keys, stores them in a Word document, embeds QR codes for public addresses, and exports data to an Excel file.

## Features

- Generates vanity Bitcoin addresses using `vanitygen.exe`
- Multi-threaded execution for faster processing
- Saves private keys in a Word document
- Embeds QR codes for public addresses in a separate Word document
- Saves public addresses in an Excel file

## Installation

Ensure you have Python installed. Install required dependencies using:

```sh
pip install -r requirements.txt
```

## Usage

Run the script with:

```sh
python main.py
```

## Or Run

double click on run.bat

## Output Files

- `private_keys.docx` - Stores private keys
- `public_qr_codes.docx` - Stores QR codes for public addresses
- `public_addresses.xlsx` - Stores public addresses in Excel format

## Dependencies

- Python 3.x
- `vanitygen.exe` (ensure it's accessible in the specified path)
- Required Python modules (see `requirements.txt`)

## Notes

- Ensure `vanitygen.exe` is placed in the correct directory before execution.
- The script uses multi-threading to generate addresses efficiently.
- Temporary QR code images are deleted after document creation.
