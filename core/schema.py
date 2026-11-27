"""Field and amount contracts for Standard_TB, Standard_GL, and Standard_Prior_TB."""

from dataclasses import dataclass, fields
from datetime import date, datetime
from decimal import Decimal, InvalidOperation
import re

TB_SHEET = "Raw_TB"
GL_SHEET = "Standard_GL"
PRIOR_TB_SHEET = "Raw_Prior_TB"
ZERO = Decimal("0.00")
CENT = Decimal("0.01")
MAX_AMOUNT = Decimal("9999999999999.99")
TB_AMOUNT_FIELDS = ("open_debit", "open_credit", "period_debit", "period_credit", "close_debit", "close_credit")
PRIOR_AMOUNT_FIELDS = ("prior_audited_debit", "prior_audited_credit")
DIRECTIONS = {"借": "借", "贷": "贷", "平": "平", "借方": "借", "贷方": "贷", "DEBIT": "借", "CREDIT": "贷", "BALANCED": "平"}
# Keep Excel SUMIFS wildcard characters (*, ?, ~) out of normalized account codes.
ACCOUNT_CODE = re.compile(r"[A-Z0-9][A-Z0-9._/-]{0,63}", re.ASCII)
AMOUNT_TEXT = re.compile(r"[+-]?(?:[0-9]+|[0-9]{1,3}(?:,[0-9]{3})+)(?:\.[0-9]+)?", re.ASCII)
ILLEGAL_TEXT = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f\ud800-\udfff\ufffe\uffff]")


class AuditError(Exception):
    """An error message that can be displayed without source cell values."""


@dataclass(frozen=True)
class StandardTB:
    account_code: str
    account_name: str
    level: int
    direction: str
    open_debit: Decimal
    open_credit: Decimal
    period_debit: Decimal
    period_credit: Decimal
    close_debit: Decimal
    close_credit: Decimal
    aux_name: str | None = None


@dataclass(frozen=True)
class StandardGL:
    date: date
    voucher_no: str
    summary: str
    account_code: str
    debit: Decimal
    credit: Decimal
    aux_name: str


@dataclass(frozen=True)
class StandardPriorTB:
    account_code: str
    account_name: str
    prior_audited_debit: Decimal
    prior_audited_credit: Decimal
    aux_name: str | None = None


TB_FIELDS = tuple(field.name for field in fields(StandardTB))
GL_FIELDS = tuple(field.name for field in fields(StandardGL))
PRIOR_TB_FIELDS = tuple(field.name for field in fields(StandardPriorTB))


@dataclass(frozen=True)
class StandardData:
    tb: tuple[StandardTB, ...]
    gl: tuple[StandardGL, ...]
    prior_tb: tuple[StandardPriorTB, ...] | None = None


def text_value(value, *, optional=False) -> str:
    if optional and value is None:
        return ""
    if not isinstance(value, str):
        raise ValueError
    value = value.strip()
    if (not optional and not value) or len(value) > 32767 or ILLEGAL_TEXT.search(value):
        raise ValueError
    return value


def optional_text(value) -> str:
    return text_value(value, optional=True)


def account_code(value) -> str:
    text = text_value(value)
    code = text.upper()
    if not text.isascii() or not ACCOUNT_CODE.fullmatch(code):
        raise ValueError
    return code


def account_level(value) -> int:
    if isinstance(value, str) and re.fullmatch(r"[1-9][0-9]{0,2}", value.strip()):
        value = int(value.strip())
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    if type(value) is not int or not 1 <= value <= 999:
        raise ValueError
    return value


def direction_value(value) -> str:
    try:
        return DIRECTIONS[text_value(value).upper()]
    except KeyError:
        raise ValueError from None


def date_value(value) -> date:
    if isinstance(value, datetime):
        if value.tzinfo is not None or value.time() != datetime.min.time():
            raise ValueError
        value = value.date()
    if isinstance(value, str):
        value = value.strip()
        if not re.fullmatch(r"[0-9]{4}-[0-9]{2}-[0-9]{2}", value):
            raise ValueError
        value = date.fromisoformat(value)
    if type(value) is not date or value < date(1900, 3, 1):
        raise ValueError
    return value


def amount_value(value, *, optional=False) -> Decimal:
    if optional and (value is None or isinstance(value, str) and not value.strip()):
        return ZERO
    if isinstance(value, str):
        value = value.strip()
        if value.startswith("(") and value.endswith(")") and value[1:2] not in {"+", "-"}:
            value = "-" + value[1:-1]
        if not AMOUNT_TEXT.fullmatch(value):
            raise ValueError
        value = value.replace(",", "")
    elif isinstance(value, bool) or not isinstance(value, (int, float, Decimal)):
        raise ValueError
    try:
        amount = Decimal(str(value))
        if not amount.is_finite() or abs(amount) > MAX_AMOUNT or amount != amount.quantize(CENT):
            raise ValueError
    except InvalidOperation:
        raise ValueError from None
    return amount.quantize(CENT)


def validate(data: StandardData) -> None:
    if not data.tb:
        raise AuditError("Standard_TB 没有数据行。")
    for line, row in enumerate(data.tb, 2):
        try:
            if not isinstance(row, StandardTB) or row.account_code != account_code(row.account_code):
                raise ValueError
            if row.account_name != text_value(row.account_name) or type(row.level) is not int:
                raise ValueError
            account_level(row.level)
            if row.direction not in {"借", "贷", "平"}:
                raise ValueError
            if row.aux_name is not None and row.aux_name != text_value(row.aux_name):
                raise ValueError
            for field in TB_AMOUNT_FIELDS:
                value = getattr(row, field)
                if type(value) is not Decimal or value != amount_value(value):
                    raise ValueError
        except ValueError:
            raise AuditError(f"Standard_TB 第 {line} 行不符合字段契约。") from None
    for line, row in enumerate(data.gl, 2):
        try:
            if not isinstance(row, StandardGL) or row.account_code != account_code(row.account_code):
                raise ValueError
            if type(row.date) is not date or row.date != date_value(row.date):
                raise ValueError
            if row.voucher_no != text_value(row.voucher_no):
                raise ValueError
            for field in ("summary", "aux_name"):
                value = getattr(row, field)
                if value != optional_text(value):
                    raise ValueError
            for value in (row.debit, row.credit):
                if type(value) is not Decimal or value != amount_value(value):
                    raise ValueError
        except ValueError:
            raise AuditError(f"Standard_GL 第 {line} 行不符合字段契约。") from None
    if data.prior_tb is None:
        return
    if not data.prior_tb:
        raise AuditError("Standard_Prior_TB 没有数据行。")
    for line, row in enumerate(data.prior_tb, 2):
        try:
            if not isinstance(row, StandardPriorTB) or row.account_code != account_code(row.account_code):
                raise ValueError
            if row.account_name != text_value(row.account_name):
                raise ValueError
            if row.aux_name is not None and row.aux_name != text_value(row.aux_name):
                raise ValueError
            for field in PRIOR_AMOUNT_FIELDS:
                value = getattr(row, field)
                if type(value) is not Decimal or value != amount_value(value):
                    raise ValueError
        except ValueError:
            raise AuditError(f"Standard_Prior_TB 第 {line} 行不符合字段契约。") from None
