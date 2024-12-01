import os
import logging
import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import pdfplumber
from typing import Optional, Tuple


class PDFToExcelConverter:
    def __init__(self):
        """
        Initialize the PDF to Excel converter with logging and configuration.
        """
        logging.basicConfig(
            level=logging.INFO,
            format="%(asctime)s - %(levelname)s: %(message)s",
            handlers=[
                logging.FileHandler("pdf_conversion.log"),
                logging.StreamHandler(),
            ],
        )
        self.logger = logging.getLogger(__name__)

    def extract_table_from_page(self, page) -> Optional[list]:
        """
        Extract table from a PDF page with robust error handling.
        """
        try:
            table = page.extract_table()
            return table if table and len(table) > 0 else None
        except Exception as e:
            self.logger.warning(f"Could not extract table from page: {e}")
            return None

    def convert_pdf_to_excel(self, pdf_path: str, excel_path: str) -> bool:
        """
        Convert PDF to Excel with comprehensive error handling and logging.
        """
        try:
            if not os.path.exists(pdf_path):
                self.logger.error(f"PDF file not found: {pdf_path}")
                return False

            with pdfplumber.open(pdf_path) as pdf:
                total_pages = len(pdf.pages)
                print(f"Processing {os.path.basename(pdf_path)} ({total_pages} pages)")

                all_tables = []
                for page_number, page in enumerate(pdf.pages, start=1):
                    table = self.extract_table_from_page(page)

                    if table:
                        if not all_tables:
                            all_tables = table
                        else:
                            all_tables.extend(table[1:])

                    progress = (page_number / total_pages) * 100
                    print(
                        f"\rProcessing page {page_number}/{total_pages} ({progress:.2f}%)",
                        end="",
                        flush=True,
                    )

                print()  # Print a newline after progress updates.

                if not all_tables or len(all_tables) <= 1:
                    self.logger.warning("No valid table data found in PDF")
                    return False

                # Create DataFrame
                df = pd.DataFrame(
                    all_tables[1:],
                    columns=all_tables[0] if len(all_tables[0]) > 1 else None,
                )

                # Attempt to convert columns to appropriate types
                for col in df.columns:
                    try:
                        df[col] = pd.to_numeric(df[col], errors="ignore")
                    except ValueError as e:
                        self.logger.warning(f"Column '{col}' could not be converted: {e}")

                # Save to Excel
                os.makedirs(os.path.dirname(excel_path), exist_ok=True)
                df.to_excel(excel_path, index=False)
                self.logger.info(f"Successfully converted to: {excel_path}")
                return True

        except Exception as e:
            self.logger.error(f"Conversion failed: {e}")
            return False

    def select_folder_and_output(self) -> Tuple[Optional[str], Optional[str]]:
        """
        Interactive GUI for selecting an input folder and destination folder.
        """
        root = tk.Tk()
        root.withdraw()

        try:
            # Select input folder
            input_folder = filedialog.askdirectory(
                title="Select Input Folder Containing PDFs"
            )
            if not input_folder:
                self.logger.info("No input folder selected. Exiting.")
                return None, None

            # Select destination folder
            output_parent_folder = filedialog.askdirectory(
                title="Select Output Parent Folder"
            )
            if not output_parent_folder:
                self.logger.info("No output folder selected. Exiting.")
                return None, None

            return input_folder, output_parent_folder

        except Exception as e:
            self.logger.error(f"Selection error: {e}")
            return None, None

    def run(self):
        """
        Main execution method to convert PDFs in a folder to Excel.
        """
        input_folder, output_parent_folder = self.select_folder_and_output()

        if not input_folder or not output_parent_folder:
            self.logger.info("Conversion cancelled or invalid selection.")
            print("No valid folders selected. Exiting.")
            return

        # Create output folder named the same as the input folder
        output_folder = os.path.join(
            output_parent_folder, os.path.basename(input_folder)
        )
        os.makedirs(output_folder, exist_ok=True)

        for root_dir, _, files in os.walk(input_folder):
            for file_name in files:
                if file_name.lower().endswith(".pdf"):
                    pdf_path = os.path.join(root_dir, file_name)
                    relative_path = os.path.relpath(pdf_path, input_folder)
                    excel_path = os.path.join(
                        output_folder,
                        os.path.splitext(relative_path)[0] + ".xlsx",
                    )

                    print(f"Converting {file_name}...")
                    success = self.convert_pdf_to_excel(pdf_path, excel_path)

                    if success:
                        print(f"Successfully converted {file_name} to Excel!")
                    else:
                        print(f"Failed to convert {file_name}.")

        print(f"All files processed. Output stored in: {output_folder}")


def main():
    """
    Entry point for the PDF to Excel converter.
    """
    converter = PDFToExcelConverter()
    converter.run()


if __name__ == "__main__":
    main()
