"""Raw sheets and their stable literal lookup keys."""

from dataclasses import astuple

from core.schema import (
    GL_FIELDS, GL_SHEET, PRIOR_AMOUNT_FIELDS, PRIOR_TB_FIELDS, PRIOR_TB_SHEET, TB_AMOUNT_FIELDS, TB_FIELDS, TB_SHEET,
)

AJE = "AJE"
AMOUNT_FORMAT = '#,##0.00;(#,##0.00);"-"'
PERCENT_FORMAT = "0.00%"
ACCOUNT_KEY_PREFIX = "acct:"
RAW_TB_FIELDS = (*TB_FIELDS, "account_key")
WORK_TB_FIELDS = (*RAW_TB_FIELDS, "balance_key", "detail_key")
RAW_PRIOR_FIELDS = (*PRIOR_TB_FIELDS, "account_key", "balance_key")
WORK_GL_FIELDS = (*GL_FIELDS, "account_key", "month")
AMOUNT_FIELDS = frozenset((*TB_AMOUNT_FIELDS, *PRIOR_AMOUNT_FIELDS, "debit", "credit"))
TABLE_FORMATS = {**dict.fromkeys(AMOUNT_FIELDS, AMOUNT_FORMAT), "date": "yyyy-mm-dd", "level": "0", "month": "0"}


def account_key(code):
    return ACCOUNT_KEY_PREFIX + code


def balance_keys(index):
    return {line: account_key(code) for code, lines in index.roots.items() for line in lines}


def tb_records(data, ledger, detail_keys):
    keys = balance_keys(ledger.current)
    for line, record in enumerate(data.tb, 2):
        yield (*astuple(record), account_key(record.account_code), keys.get(line), detail_keys.get(line))


def prior_records(data, ledger):
    keys = balance_keys(ledger.prior)
    for line, record in enumerate(data.prior_tb or (), 2):
        yield (*astuple(record), account_key(record.account_code), keys.get(line))


def gl_records(data):
    for record in data.gl:
        yield (*astuple(record), account_key(record.account_code), record.date.month)


def raw_tables(data, ledger, detail_keys=None, *, include_gl=True):
    yield TB_SHEET, WORK_TB_FIELDS, tb_records(data, ledger, detail_keys or {}), len(data.tb)
    yield PRIOR_TB_SHEET, RAW_PRIOR_FIELDS, prior_records(data, ledger), len(data.prior_tb or ())
    if include_gl:
        yield GL_SHEET, WORK_GL_FIELDS, gl_records(data), len(data.gl)


def raw_range(sheet, column, length):
    return f"{sheet}!${column}$2:${column}${max(2, length + 1)}"


def sum_by_key(sheet, column, key_column, key, length, *extra):
    criteria = ",".join((raw_range(sheet, key_column, length), key, *extra))
    return f"SUMIFS({raw_range(sheet, column, length)},{criteria})"


def prior_lookup(code, length):
    key = f'"{ACCOUNT_KEY_PREFIX}"&{code}'
    count = f"COUNTIF({raw_range(PRIOR_TB_SHEET, 'G', length)},{key})"
    debit, credit = (sum_by_key(PRIOR_TB_SHEET, col, "G", key, length) for col in "CD")
    return count, debit, credit
