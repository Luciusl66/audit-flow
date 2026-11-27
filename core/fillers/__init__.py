"""Resolve a reviewed blueprint archetype to its reusable filler."""

from core.fillers.asset_rollforward import AssetRollforward
from core.fillers.counterparty_tb import CounterpartyTB
from core.fillers.flat_tb import FlatTB
from core.fillers.gl_monthly_pivot import GLMonthlyPivot
from core.fillers.tax_schedule import TaxSchedule

FILLERS = {
    "flat_tb": FlatTB,
    "counterparty_tb": CounterpartyTB,
    "gl_monthly_pivot": GLMonthlyPivot,
    "asset_rollforward": AssetRollforward,
    "tax_schedule": TaxSchedule,
}


def make_filler(writer, ledger, codes):
    blueprint = writer.blueprint
    return FILLERS[blueprint.archetype](writer, ledger, codes, blueprint.normal_balance)
