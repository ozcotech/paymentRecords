import sys
import os
from tkinter import Toplevel, Label, Entry, Button, messagebox

sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), "data"))) 

from excel_manager import ExcelManager
from models import Payment
import tkinter as tk

class PaymentGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Payment Management System")
        self.excel = ExcelManager()

        tk.Label(root, text="Payment Management System", font=("Arial", 14, "bold")).pack(pady=10)

        tk.Button(root, text="Add New Payment", command=self.add_payment, width=30).pack(pady=5)
        tk.Button(root, text="Update Payment Status", command=self.update_payment_status, width=30).pack(pady=5)
        tk.Button(root, text="Search Payment", command=self.search_payment, width=30).pack(pady=5)
        tk.Button(root, text="List All Payments", command=self.list_payments, width=30).pack(pady=5)
        tk.Button(root, text="Analyze Payments", command=self.analyze_payments, width=30).pack(pady=5)
        tk.Button(root, text="Generate Payment Chart", command=self.generate_chart, width=30).pack(pady=5)
        tk.Button(root, text="Exit", command=root.quit, width=30, bg="red", fg="white").pack(pady=5)

    def add_payment(self):
        """
        Opens a new window to add a payment.
        """
        add_window = Toplevel(self.root)
        add_window.title("Add New Payment")
        add_window.geometry("400x400")

        Label(add_window, text="Invoice No:").pack()
        invoice_entry = Entry(add_window)
        invoice_entry.pack()

        Label(add_window, text="Task Type:").pack()
        task_entry = Entry(add_window)
        task_entry.pack()

        Label(add_window, text="Tariff Fee (TL):").pack()
        tariff_entry = Entry(add_window)
        tariff_entry.pack()

        Label(add_window, text="Gross Fee (TL):").pack()
        gross_entry = Entry(add_window)
        gross_entry.pack()

        Label(add_window, text="VAT (%):").pack()
        vat_entry = Entry(add_window)
        vat_entry.pack()

        Label(add_window, text="Case Details:").pack()
        case_entry = Entry(add_window)
        case_entry.pack()

        Label(add_window, text="Submission Date:").pack()
        submission_entry = Entry(add_window)
        submission_entry.pack()

        Label(add_window, text="Invoice Date:").pack()
        invoice_date_entry = Entry(add_window)
        invoice_date_entry.pack()

        def save_payment():
            """
            Saves the entered payment details into the Excel file.
            """
            invoice_no = invoice_entry.get()
            task_type = task_entry.get()
            tariff_fee = float(tariff_entry.get())
            gross_fee = float(gross_entry.get())
            vat_rate = float(vat_entry.get())
            vat_amount = gross_fee * (vat_rate / 100)
            net_fee = gross_fee - vat_amount
            case_details = case_entry.get()
            submission_date = submission_entry.get()
            invoice_date = invoice_date_entry.get()

            new_payment = Payment(invoice_no, task_type, tariff_fee, gross_fee, vat_rate, vat_amount, net_fee, case_details, submission_date, invoice_date)
            self.excel.add_payment(new_payment.to_list())

            messagebox.showinfo("Success", "New payment added successfully!")
            add_window.destroy()

        Button(add_window, text="Save Payment", command=save_payment).pack(pady=10)

    

    def update_payment_status(self):
        """
        Opens a window to update the payment status.
        """
        update_window = Toplevel(self.root)
        update_window.title("Update Payment Status")
        update_window.geometry("350x200")

        Label(update_window, text="Enter Invoice No:").pack()
        invoice_entry = Entry(update_window)
        invoice_entry.pack()

        status_var = tk.StringVar()
        status_var.set("Paid")  # Default selection

        Label(update_window, text="Select New Status:").pack()
        status_dropdown = tk.OptionMenu(update_window, status_var, "Paid", "Pending")
        status_dropdown.pack()

        def save_status():
            invoice_no = invoice_entry.get()
            new_status = status_var.get()

            if not invoice_no:
                messagebox.showerror("Error", "Please enter an Invoice No!")
                return

            success = self.excel.update_payment_status(invoice_no, new_status)

            if success:
                messagebox.showinfo("Success", f"Payment status updated to {new_status}!")
            else:
                messagebox.showerror("Error", "Invoice No not found!")

            update_window.destroy()

        Button(update_window, text="Update Status", command=save_status).pack(pady=10)

    def search_payment(self):
        print("🔹 Search Payment Clicked")

    def list_payments(self):
        print("🔹 List All Payments Clicked")

    def analyze_payments(self):
        print("🔹 Analyze Payments Clicked")

    def generate_chart(self):
        print("🔹 Generate Payment Chart Clicked")

if __name__ == "__main__":
    root = tk.Tk()
    app = PaymentGUI(root)
    root.mainloop()