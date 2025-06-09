from datetime import datetime

class StatementAccount:
    def __init__(self, hbl_no=None, mbl=None, date=None, debit=None, credit=None, apply_to=None):
        self.hbl_no = hbl_no
        self.mbl = mbl
        self.date = date
        self.debit = debit
        self.credit = credit
        self.apply_to = apply_to

    @staticmethod
    def validate_date(date_str):
        try:
            datetime.strptime(date_str, "%d/%m/%Y")
            return True
        except ValueError:
            return False

    def to_dict(self):
        return {
            "HBL_NO": self.hbl_no,
            "MBL": self.mbl,
            "Date": self.date,
            "DEBIT": self.debit,
            "CREDIT": self.credit,
            "apply_to": self.apply_to
        } 