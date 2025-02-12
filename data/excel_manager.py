import os
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment, PatternFill

class ExcelManager:
    """
    This class handles the creation and management of an Excel file for tracking payments.
    """

    def __init__(self, file_path=None):
        """
        Constructor for the ExcelManager class.
        Ensures the correct file path is used.
        """
        # Get the absolute path of the project root directory
        base_dir = os.path.abspath(os.path.dirname(__file__))

        # Ensure 'data' folder is directly under project root
        if "data" in base_dir.split(os.sep):
            base_dir = os.path.dirname(base_dir)  # Move up if already inside 'data'

        data_dir = os.path.join(base_dir, "data")
        if not os.path.exists(data_dir):
            os.makedirs(data_dir)  # Create the 'data' directory if it doesn't exist

        if file_path is None:
            file_path = os.path.join(data_dir, "payment_records.xlsx")  # Set the correct file path

        self.file_path = file_path

        if not os.path.exists(self.file_path):
            self.create_excel_file()
    
    def create_excel_file(self):
        """
        Creates a new Excel file and initializes the headers.
        """
        wb = Workbook()
        ws = wb.active
        ws.title = "Payment Records"

        # Define headers for the Excel file
        headers = [
            "Invoice No", "Task Type", "Tariff Fee", "Gross Fee (TL)", "VAT (%)",
            "VAT Amount (TL)", "Net Fee (TL)", "Case Details", "Submission Date",
            "Invoice Date", "Payment Status"
        ]
        ws.append(headers)

        # Save the Excel file
        wb.save(self.file_path)
        print(f"✅ New Excel file created: {self.file_path}")

    def load_workbook(self):
        """
        Loads the existing Excel file and returns the workbook object.
        """
        return load_workbook(self.file_path)

    def add_payment(self, payment_data):
        """
        Adds a new payment record to the Excel file.
        """
        wb = self.load_workbook()
        ws = wb.active
        ws.append(payment_data)  # Append new row
        wb.save(self.file_path)
        wb.close()
        print("✅ New payment record added.")

    def adjust_excel_formatting(self):
        """
        Adjusts column widths and row heights dynamically.
        Ensures proper text wrapping and applies TL currency formatting.
        """
        wb = self.load_workbook()
        ws = wb.active

        # Task Type column letter (B sütunu)
        task_type_column_letter = "B"

        # Adjust column widths based on content
        for col in ws.columns:
            max_length = 0
            col_letter = col[0].column_letter  # Get the column letter (A, B, C...)

            for cell in col:
                try:
                    if cell.value:
                        max_length = max(max_length, len(str(cell.value)))
                except:
                    pass

            # Limit the column width to 30 characters, force wrapping
            if col_letter == task_type_column_letter:
                ws.column_dimensions[col_letter].width = 30  # Fixed width
            else:
                ws.column_dimensions[col_letter].width = max_length + 2  # Apply precise width

        # Apply TL format and adjust row heights dynamically
        for row_idx, row in enumerate(ws.iter_rows(min_row=2, max_row=ws.max_row), start=2):
            max_height = 15  # Default row height
            for idx, cell in enumerate(row):
                if idx in [2, 3, 5, 6]:  # Columns: Tariff Fee, Gross Fee, VAT Amount, Net Fee
                    if isinstance(cell.value, (int, float)):
                        cell.number_format = '#,##0.00 TL'  # Set currency format
                        cell.value = float(cell.value)  # Ensure numeric format

                # Force text wrapping
                cell.alignment = Alignment(wrap_text=True, vertical="top", horizontal="left")

                # Only for Task Type column (B)
                if cell.column_letter == task_type_column_letter:
                    text_length = len(str(cell.value)) if cell.value else 0
                    lines = (text_length // 30) + 1  # Wrap after 30 characters
                    max_height = max(max_height, lines * 15)  # Adjust row height

            ws.row_dimensions[row_idx].height = max_height  # Apply final row height

        wb.save(self.file_path)
        wb.close()
        print("✅ Excel formatting adjusted: column widths, row heights, and TL format applied!")

    def update_payment_status(self, invoice_no, new_status="Paid"):
        """
        Updates the payment status for a specific invoice number.
        If the invoice is not found, it prints an error message.
        """
        wb = self.load_workbook()
        ws = wb.active

        found = False  # Flag to check if invoice is found

        for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
            if row[0].value == invoice_no:  # Invoice No column (first column)
                print(f"🔄 Updating {invoice_no} from {row[-1].value} to {new_status}")
                row[-1].value = new_status  # Update last column (Payment Status)
                found = True
                break  # Stop searching after finding the first match

        if found:
            wb.save(self.file_path)  # Ensure changes are saved
            wb.close()
            print(f"✅ Payment status updated for Invoice No: {invoice_no}")

            # Verify the update by reloading and printing
            self.list_payments()
        else:
            wb.close()
            print(f"❌ Invoice No {invoice_no} not found.")  # Print error message if invoice not found

    def search_payment(self, invoice_no):
        """
        Searches for a payment record by invoice number.
        If found, prints the payment details.
        """
        wb = self.load_workbook()
        ws = wb.active

        for row in ws.iter_rows(min_row=2, max_row=ws.max_row, values_only=True):
            if row[0] == invoice_no:  # Invoice No column (first column)
                print("\n✅ Payment Record Found:")
                print(f"Invoice No: {row[0]}")
                print(f"Task Type: {row[1]}")
                print(f"Tariff Fee: {row[2]:,.2f} TL")
                print(f"Gross Fee: {row[3]:,.2f} TL")
                print(f"VAT (%): {row[4]}")
                print(f"VAT Amount: {row[5]:,.2f} TL")
                print(f"Net Fee: {row[6]:,.2f} TL")
                print(f"Case Details: {row[7]}")
                print(f"Submission Date: {row[8]}")
                print(f"Invoice Date: {row[9]}")
                print(f"Payment Status: {row[10]}")
                wb.close()
                return  # Stop searching after finding the first match

        wb.close()
        print(f"❌ Invoice No {invoice_no} not found.")

    def list_payments(self):
        """
        Displays all recorded payments in the console.
        """
        wb = self.load_workbook()
        ws = wb.active

        print("\n📌 Recorded Payments:")
        for row in ws.iter_rows(values_only=True):
            print(row)
    
    def analyze_payments(self):
        """
        Analyzes payment records and provides statistical insights.
        """
        wb = self.load_workbook()
        ws = wb.active

        total_payments = 0
        total_paid = 0
        total_pending = 0
        sum_paid_amounts = 0
        sum_pending_amounts = 0

        for row in ws.iter_rows(min_row=2, max_row=ws.max_row, values_only=True):
            if row[-1] == "Paid":
                total_paid += 1
                sum_paid_amounts += row[6]  # Net Fee column (index 6)
            elif row[-1] == "Pending":
                total_pending += 1
                sum_pending_amounts += row[6]  # Net Fee column (index 6)

            total_payments += 1

        avg_paid = sum_paid_amounts / total_paid if total_paid > 0 else 0
        avg_pending = sum_pending_amounts / total_pending if total_pending > 0 else 0

        print("\n📊 Payment Analysis:")
        print(f"Total Payments: {total_payments}")
        print(f"Total Paid: {total_paid} (Total Amount: {sum_paid_amounts:,.2f} TL, Avg: {avg_paid:,.2f} TL)")
        print(f"Total Pending: {total_pending} (Total Amount: {sum_pending_amounts:,.2f} TL, Avg: {avg_pending:,.2f} TL)")

        wb.close()
        

    def highlight_payments(self):
        """
        Highlights 'Paid' payments in green and 'Pending' payments in red in the Excel file.
        """
        wb = self.load_workbook()
        ws = wb.active

        # Define color fills
        green_fill = PatternFill(start_color="C6E0B4", end_color="C6E0B4", fill_type="solid")  # Light green
        red_fill = PatternFill(start_color="F4CCCC", end_color="F4CCCC", fill_type="solid")  # Light red

        for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
            status_cell = row[-1]  # Payment Status column (last column)

            if status_cell.value == "Paid":
                status_cell.fill = green_fill  # Apply green fill
            elif status_cell.value == "Pending":
                status_cell.fill = red_fill  # Apply red fill

        wb.save(self.file_path)
        wb.close()
        print("✅ Payment statuses highlighted in Excel!")        