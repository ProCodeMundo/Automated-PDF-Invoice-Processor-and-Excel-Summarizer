"""
Project Title: Automated PDF Invoice Processor and Excel Summarizer
Developer: Ruben E.
Date: June 2024
Youtube link: https://www.youtube.com/watch?v=M2nIzosfHTU

This program processes PDF invoices from a specified directory, extracts relevant information, and compiles it into an Excel spreadsheet. The key functionalities include:

PDF Extraction: Extracts invoice details such as invoice number, payment method, date and time, coin amount, price, subtotal, Coinbase fee, and total amount using regex patterns.

File Management: Moves processed PDF files to a designated "Processed_Invoice" folder.

Data Compilation: Aggregates extracted invoice data into an Excel file, updating it with new information each time the program runs.

Excel Formatting: Adjusts column widths, applies styling to headers (dark blue with white text), and sets a gray background for data cells for better readability.

Overall, the program automates the extraction, organization, and storage of invoice data, ensuring a structured and easily accessible summary of all processed invoices.
"""

# Import Libraries 
import os
import re
import shutil
import pandas as pd
from PyPDF2 import PdfReader
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, numbers, PatternFill

class InvoiceProcessor:
    def setup(self):
        # Initialize the invoice data (Empty) list, destination folder, and excel path
        self.invoice_data = []
        #Set the destination folder for processed invoices
        self.destination_folder = "Processed_Invoice"
        # Define the path for the summary Excel file within the destination folder
        self.excel_path = os.path.join(self.destination_folder, "Invoices_Summary.xlsx")
        # Define regex patterns for extracting invoice information
        self.patterns = {
            'invoice_number': r'Reference\s*code\s*([A-Z0-9]+)',
            'payment_method': r'Payment\s*method\s*(.*?)\s*[\n\r]',
            'date_time': r'Date\s+and\s+time\s*(.*?)\s*\n',
            'amount_number': r'Amount\s*([\d]+\.\d{1,12})\s*(\w+)',
            'price_number': r'Price\s*@\s*\$([\d,]+\.\d+)\s*/\s*(\w+)',
            'subtotal_number': r'Subtotal\s*\$([0-9,]+\.\d+)',
            'coinbase_fee_number': r'Coinbase\s*fee\s*\$([0-9,]+\.\d+)',
            'total_number': r'Total\s*\$([0-9,]+\.\d+)'
        }

    def extract_invoice_info(self, pdf_file_path):
         # Extract information from the PDF file
        file_name = os.path.basename(pdf_file_path)
        try:
            with open(pdf_file_path, 'rb') as file:
                pdf_reader = PdfReader(file)
                text = ''.join(page.extract_text() for page in pdf_reader.pages)
                # Apply patterns to extract invoice information
                matches = {key: re.search(pattern, text) for key, pattern in self.patterns.items()}
                invoice_info = {
                    'File Name': file_name,
                    'Invoice Number': matches['invoice_number'].group(1) if matches['invoice_number'] else None,
                    'Payment Method': matches['payment_method'].group(1) if matches['payment_method'] else None,
                    'Date & Time': matches['date_time'].group(1) if matches['date_time'] else None,
                    'Coin Name': matches['amount_number'].group(2) if matches['amount_number'] else None,
                    'Amount of Coin': float(matches['amount_number'].group(1)) if matches['amount_number'] else None,
                    'Coin Price': float(matches['price_number'].group(1).replace(',', '')) if matches['price_number'] else None,
                    'Subtotal': float(matches['subtotal_number'].group(1).replace(',', '')) if matches['subtotal_number'] else 0.0,
                    'CoinBase Fee': float(matches['coinbase_fee_number'].group(1).replace(',', '')) if matches['coinbase_fee_number'] else 0.0,
                    'Total': float(matches['total_number'].group(1).replace(',', '')) if matches['total_number'] else 0.0
                }
                # Append the extracted invoice information to the list
                self.invoice_data.append(invoice_info)
        except Exception as e:
            print(f"Error processing {pdf_file_path}: {e}")

    def move_processed_file(self, file_path):
         # Move the processed file to the destination folder
        os.makedirs(self.destination_folder, exist_ok=True)
        shutil.move(file_path, self.destination_folder)
        print(f'File moved to: {self.destination_folder}')

    def process_invoices(self, directory_path):
         # Load existing invoice data if the excel file exists
        self.setup()
        
        if os.path.exists(self.excel_path):
            existing_df = pd.read_excel(self.excel_path).iloc[:, :10]
        else:
            existing_df = pd.DataFrame()
        
        for filename in os.listdir(directory_path):
            if filename.endswith(".pdf"):
                pdf_file_path = os.path.join(directory_path, filename)
                self.extract_invoice_info(pdf_file_path)
                self.move_processed_file(pdf_file_path)
        
        new_df = pd.DataFrame(self.invoice_data).iloc[:, :10]
        combined_df = pd.concat([existing_df, new_df], ignore_index=True)
        combined_df.to_excel(self.excel_path, index=False)
        print('Invoice summary saved to Processed_Invoice/Invoices_Summary.xlsx')
        self.adjust_excel_file()
        print('Excel columns adjusted and aligned.')

    def adjust_excel_file(self):
         # Load the excel file and get the active worksheet
        wb = load_workbook(self.excel_path)
        ws = wb.active
        
        # Define styles for headers and data cells
        header_font = Font(color="FFFFFF", bold=True)
        header_fill = PatternFill(start_color="00008B", end_color="00008B", fill_type="solid")
        gray_fill = PatternFill(start_color="e0e0e0", end_color="e0e0e0", fill_type="solid")
        
        for cell in ws[1]:
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal='center')

        # Adjust column widths and styles
        for col in ws.columns:
            max_length = max(len(str(cell.value)) for cell in col)
            ws.column_dimensions[col[0].column_letter].width = max_length + 2
            for cell in col:
                if cell.row > 1:  # Skip header row
                    cell.fill = gray_fill
        
        # Adjust alignment and number format
        for col in ['F', 'G', 'H', 'I', 'J']:
            for cell in ws[col]:
                cell.alignment = Alignment(horizontal='right')
                cell.number_format = '0.00' if col != 'F' else '0.000000000' # Some coins need to be this long for better accuracy.
        
        # Save the adjusted excel file
        wb.save(self.excel_path)

# Instantiate and call the processor
processor = InvoiceProcessor()
processor.process_invoices("Invoices/")
