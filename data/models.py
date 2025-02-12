class Payment:
    """
    A class representing a payment record.
    """
    
    def __init__(self, invoice_no, task_type, tariff_fee, gross_fee, vat_rate,
                 vat_amount, net_fee, case_details, submission_date, invoice_date, payment_status="Pending"):
        """
        Initializes a Payment object.
        """
        self.invoice_no = invoice_no
        self.task_type = task_type
        self.tariff_fee = tariff_fee
        self.gross_fee = gross_fee
        self.vat_rate = vat_rate
        self.vat_amount = vat_amount
        self.net_fee = net_fee
        self.case_details = case_details
        self.submission_date = submission_date
        self.invoice_date = invoice_date
        self.payment_status = payment_status

    def to_list(self):
        """
        Returns the payment details as a list.
        """
        return [
            self.invoice_no, self.task_type, self.tariff_fee, self.gross_fee,
            self.vat_rate, self.vat_amount, self.net_fee, self.case_details,
            self.submission_date, self.invoice_date, self.payment_status
        ]