import sys
import os
from tkinter import Toplevel, Label, Entry, Button, messagebox, ttk
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

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

        tk.Button(root, text="Add New Payment", command=self.add_payment, width=20).pack(pady=5)
        tk.Button(root, text="Update Payment Status", command=self.update_payment_status, width=20).pack(pady=5)
        tk.Button(root, text="Search Payment", command=self.search_payment, width=20).pack(pady=5)
        tk.Button(root, text="List All Payments", command=self.list_payments, width=20).pack(pady=5)
        tk.Button(root, text="Analyze Payments", command=self.analyze_payments_gui, width=20).pack(pady=5)
        tk.Button(root, text="Generate Payment Chart", command=self.generate_chart_gui, width=20).pack(pady=5)
        tk.Button(root, text="Exit", command=root.quit, width=20, bg="red", fg="black").pack(pady=5)

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
        """
        Opens a window to search for a payment by Invoice No.
        """
        search_window = Toplevel(self.root)
        search_window.title("Search Payment")
        search_window.geometry("350x200")

        Label(search_window, text="Enter Invoice No:").pack()
        invoice_entry = Entry(search_window)
        invoice_entry.pack()

        def find_payment():
            invoice_no = invoice_entry.get()

            if not invoice_no:
                messagebox.showerror("Error", "Please enter an Invoice No!")
                return

            payment_data = self.excel.search_payment(invoice_no)

            if payment_data:
                details = f"""
                Invoice No: {payment_data[0]}
                Task Type: {payment_data[1]}
                Tariff Fee: {payment_data[2]:,.3f} TL
                Gross Fee: {payment_data[3]:,.3f} TL
                VAT (%): {payment_data[4]}
                VAT Amount: {payment_data[5]:,.3f} TL
                Net Fee: {payment_data[6]:,.3f} TL
                Case Details: {payment_data[7]}
                Submission Date: {payment_data[8]}
                Invoice Date: {payment_data[9]}
                Payment Status: {payment_data[10]}
                """
                messagebox.showinfo("Payment Found", details)
            else:
                messagebox.showerror("Error", "Invoice No not found!")

            search_window.destroy()

        Button(search_window, text="Search", command=find_payment).pack(pady=10)

    def list_payments(self):
        """
        Opens a window to display all payments in a table format.
        """
        list_window = Toplevel(self.root)
        list_window.title("All Payments")
        list_window.geometry("1000x400")

        columns = ("Invoice No", "Task Type", "Tariff Fee", "Gross Fee", "VAT (%)",
                "VAT Amount", "Net Fee", "Case Details", "Submission Date",
                "Invoice Date", "Payment Status")

        tree = ttk.Treeview(list_window, columns=columns, show="headings")

        for col in columns:
            tree.heading(col, text=col)
            tree.column(col, width=100)  # Set column width

        tree.pack(fill="both", expand=True)

        payments = self.excel.get_all_payments()

        for payment in payments:
            tree.insert("", "end", values=payment)

    def analyze_payments_gui(self):
        """
        Opens a window to display payment analysis statistics.
        Shows total net and gross fee instead of average.
        """
        analysis_window = Toplevel(self.root)
        analysis_window.title("Payment Analysis")
        analysis_window.geometry("400x300")

        # Get analysis results from excel_manager.py
        total_payments, total_paid, total_net_paid, total_gross_paid, total_pending, total_net_pending, total_gross_pending = self.excel.analyze_payments()

        # Display results in the GUI
        Label(analysis_window, text=f"Total Payments: {total_payments}").pack()
        Label(analysis_window, text=f"Total Paid: {total_paid}").pack()
        Label(analysis_window, text=f"Total Pending: {total_pending}").pack()

        # Show total net and gross fee amounts
        Label(analysis_window, text=f"Total Net Paid: {total_net_paid:,.3f} TL").pack()
        Label(analysis_window, text=f"Total Gross Paid: {total_gross_paid:,.3f} TL").pack()
        Label(analysis_window, text=f"Total Net Pending: {total_net_pending:,.3f} TL").pack()
        Label(analysis_window, text=f"Total Gross Pending: {total_gross_pending:,.3f} TL").pack()

    def generate_chart_gui(self):
        """
        Opens a window to display a pie chart of Paid vs Pending payments.
        """
        chart_window = Toplevel(self.root)
        chart_window.title("Payment Chart")
        chart_window.geometry("500x400")

        # Get payment data from Excel
        total_paid, total_pending = self.excel.get_payment_counts()

        if total_paid == 0 and total_pending == 0:
            messagebox.showwarning("No Data", "No payments recorded to generate a chart.")
            chart_window.destroy()
            return

        labels = ["Paid", "Pending"]
        sizes = [total_paid, total_pending]
        colors = ["green", "red"]

        fig, ax = plt.subplots()
        ax.pie(sizes, labels=labels, autopct="%1.1f%%", colors=colors, startangle=90)
        ax.set_title("Paid vs Pending Payments")

        canvas = FigureCanvasTkAgg(fig, master=chart_window)
        canvas.get_tk_widget().pack()
        canvas.draw()

if __name__ == "__main__":
    root = tk.Tk()
    app = PaymentGUI(root)
    root.mainloop()