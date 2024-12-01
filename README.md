# PDF-TO-EXCEL
converts .pdf to .xlsx
# PDF to Excel Converter

A Python tool to convert PDF files into Excel spreadsheets. This program supports single file, multiple files, and folder processing. It is designed with an easy-to-use GUI for selecting input/output directories and includes robust error handling and progress tracking.

## Features

- **Single File, Multiple Files, and Folder Support**: Process one file or an entire folder of PDFs.
- **Folder Structure Replication**: Maintains folder hierarchy for outputs.
- **Live Progress Tracking**: Displays real-time progress for each PDF being processed.
- **Error Handling**: Logs all errors to ensure the program doesn't crash.
- **Data Type Preservation**: Converts numeric columns to numeric types in the output Excel files.

## Requirements

- Python 3.7 or higher
- The following Python libraries:
  - `os`
  - `logging`
  - `tkinter`
  - `pandas`
  - `pdfplumber`

## Installation

1. Clone this repository:
   ```bash
   git clone https://github.com/chidanand111/pdf-to-excel.git
   cd pdf-to-excel
