from excel_manager import ExcelManager
from models import Payment

def main():
    excel = ExcelManager()

    while True:
        print("\n🔹 Payment Management System")
        print("1️⃣ Add New Payment")
        print("2️⃣ Update Payment Status")
        print("3️⃣ Search Payment")
        print("4️⃣ List All Payments")
        print("5️⃣ Analyze Payments")
        print("6️⃣ Generate Payment Chart")
        print("0️⃣ Exit")

        choice = input("Select an option: ")

        if choice == "1":
            invoice_no = input("Enter Invoice No: ")
            task_type = input("Enter Task Type: ")
            tariff_fee = float(input("Enter Tariff Fee (TL): "))
            gross_fee = float(input("Enter Gross Fee (TL): "))
            vat_rate = float(input("Enter VAT Rate (%): "))
            vat_amount = gross_fee * (vat_rate / 100)
            net_fee = gross_fee - vat_amount
            case_details = input("Enter Case Details: ")
            submission_date = input("Enter Submission Date (DD.MM.YYYY): ")
            invoice_date = input("Enter Invoice Date (DD.MM.YYYY): ")
            payment_status = "Pending"

            new_payment = Payment(invoice_no, task_type, tariff_fee, gross_fee, vat_rate,
                                  vat_amount, net_fee, case_details, submission_date, invoice_date, payment_status)

            excel.add_payment(new_payment.to_list())
            print("✅ Payment added successfully.")

        elif choice == "2":
            invoice_no = input("Enter Invoice No to update: ")
            new_status = input("Enter new status (Paid/Pending): ")
            excel.update_payment_status(invoice_no, new_status)

        elif choice == "3":
            invoice_no = input("Enter Invoice No to search: ")
            excel.search_payment(invoice_no)

        elif choice == "4":
            excel.list_payments()

        elif choice == "5":
            excel.analyze_payments()

        elif choice == "6":
            excel.generate_payment_chart()

        elif choice == "0":
            print("🚀 Exiting the system. See you later!")
            break

        else:
            print("❌ Invalid option! Please select a valid option.")

if __name__ == "__main__":
    main()