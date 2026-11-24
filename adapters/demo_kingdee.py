"""Adapter using the exact generic headers in the supplied mock profiles."""

from adapters.base_adapter import BaseAdapter, convert_rows, mapped_rows, source_rows
from core.schema import (
    StandardData, StandardGL, StandardTB, TB_AMOUNT_FIELDS, account_code, account_level,
    amount_value, date_value, direction_value, optional_text, text_value,
)

TB_COLUMNS = {
    "account_code": "科目编码", "account_name": "科目名称", "level": "科目级次", "direction": "余额方向",
    "open_debit": "期初借方", "open_credit": "期初贷方",
    "period_debit": "本期借方", "period_credit": "本期贷方",
    "close_debit": "期末借方", "close_credit": "期末贷方", "aux_name": "往来单位",
}
GL_COLUMNS = {
    "date": "记账日期", "voucher_no": "凭证号", "summary": "摘要", "account_code": "科目编码",
    "debit": "借方金额", "credit": "贷方金额", "aux_name": "往来单位",
}


def auxiliary(value):
    return optional_text(value) or None


def gl_amount(value):
    return amount_value(value, optional=True)


class Adapter(BaseAdapter):
    ar_account = "1122"
    tb_columns = TB_COLUMNS
    # In this synthetic chart, 1231 covers trade receivables (1122) only. For a real
    # chart, use a mapping such as {"1231": ("1122",)} after the data owner confirms the scope.
    cas_allowance_pairs = {}

    def transform(self, tb_path, gl_path):
        tb_types = {
            "account_code": account_code, "account_name": text_value, "level": account_level,
            "direction": direction_value, "aux_name": auxiliary,
            **dict.fromkeys(TB_AMOUNT_FIELDS, amount_value),
        }
        gl_types = {
            "date": date_value, "voucher_no": text_value, "summary": optional_text,
            "account_code": account_code, "debit": gl_amount, "credit": gl_amount, "aux_name": optional_text,
        }
        with source_rows(tb_path, self.tb_source) as rows:
            selected = mapped_rows(rows, self.tb_columns, "TB", self.tb_source.header_row)
            tb = tuple(StandardTB(**row) for row in convert_rows(selected, tb_types, "TB"))
        with source_rows(gl_path, self.gl_source) as rows:
            selected = mapped_rows(rows, GL_COLUMNS, "GL", self.gl_source.header_row)
            gl = tuple(StandardGL(**row) for row in convert_rows(selected, gl_types, "GL"))
        return StandardData(tb, gl)
