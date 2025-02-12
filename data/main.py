from excel_manager import ExcelManager
from models import Payment

def main():
    excel = ExcelManager()

    # Add a new payment record
    new_payment = Payment(
        invoice_no="AVV2025001923456789",
        task_type="Compulsory Legal Aid ServiceCompulsory Legal Aid ServiceCompulsory Legal Aid ServiceCompulsory Legal Aid ServiceCompulsory Legal Aid ServiceCompulsory Legal Aid ServiceCompulsory Legal Aid ServiceCompulsory Legal Aid ServiceCompulsory Legal Aid ServiceCompulsory Legal Aid Service",
        tariff_fee=2292.00,
        gross_fee=1910.00,
        vat_rate=20,
        vat_amount=382.00,
        net_fee=1528.00,
        case_details="Giresun CB 2025/1010 Investigation",
        submission_date="10.11.2024",
        invoice_date="15.11.2024"
    )

    excel.add_payment(new_payment.to_list())

    # Adjust column widths and font size
    excel.adjust_excel_formatting()

    # Update payment status test
    invoice_to_update = "AVV2025001923456789"
    excel.update_payment_status(invoice_to_update, "Paid") 

    # Search for a payment record
    invoice_to_search = "AVV2025001923456789"  # Geçerli bir invoice numarası gir
    excel.search_payment(invoice_to_search)

    # Test için yanlış bir invoice numarası da gir
    invalid_invoice = "INVALID2025"
    excel.search_payment(invalid_invoice)

    # Analyze payment data
    excel.analyze_payments()

    # Highlight payment statuses in Excel
    excel.highlight_payments()

    # Display the list of payments
    excel.list_payments()

if __name__ == "__main__":
    main()