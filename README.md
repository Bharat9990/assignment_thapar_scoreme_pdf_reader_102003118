# assignment_thapar_scoreme_pdf_reader_102003118
# PDF Table Extractor to Excel

This Python script extracts tables from a PDF file and saves them into an Excel workbook. It uses PyMuPDF (fitz) for PDF processing and openpyxl for Excel file handling.

## Prerequisites

- Python 3.x
- PyMuPDF (`fitz`)
- openpyxl

You can install the required packages using pip:

```bash
python extract_tables.py <pdf_file_path> <excel_file_path>

Script Overview
extract_tables.py
This script contains functions to extract tables from a PDF and save them into an Excel file.

extract_tables_from_pdf(pdf_file): Opens the PDF file and extracts tables by looking for blocks of text that are formatted as tables.

sanitize_text(text): Removes any characters that cannot be used in Excel sheets.

save_tables_to_excel(tables, excel_file): Saves the extracted tables into an Excel file. Each table is placed into a separate sheet.

Main Function
The script takes command-line arguments for the PDF file path and the Excel file path.
It extracts tables from the specified PDF file and saves them into the specified Excel file.
Error Handling
The script sanitizes text to remove any illegal characters that cannot be used in Excel sheets.
It ensures that the specified directories exist before saving the Excel file.
Notes
If any characters in the extracted text cannot be used in Excel sheets (e.g., special characters), they will be sanitized before saving.
Make sure to provide valid paths to an existing PDF file and specify an output Excel file path where the extracted tables will be saved.
