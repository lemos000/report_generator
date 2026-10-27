from dataclasses import dataclass, field
from typing import Dict

MONTH_MAP: Dict[str, int] = {
    "Janeiro": 1, "Fevereiro": 2, "Março": 3, "Abril": 4, "Maio": 5, "Junho": 6,
    "Julho": 7, "Agosto": 8, "Setembro": 9, "Outubro": 10, "Novembro": 11, "Dezembro": 12
}
DEFAULT_MONTH = 5


@dataclass(frozen=True)
class AuditSchema:
    """Mapeamento das colunas esperadas no CSV de auditoria (nome original -> nome interno)."""
    columns: Dict[str, str] = field(default_factory=lambda: {
        'textBox17': 'Canal',
        'textBox13': 'Categoria',
        'textBox20': 'CheckIn',
        'textBox19': 'CheckOut',
        'textBox22': 'Receita',
        'textBox3': 'Status',
        'textBox11': 'DataReserva'
    })


@dataclass(frozen=True)
class PickupSchema:
    """Mapeamento das colunas das planilhas de evolução e fechamento (Excel de pickup)."""
    evolution_cols: Dict[str, str] = field(default_factory=lambda: {
        'date.present.record_date': 'Data',
        'kpis_occupancy%.past.after.value': 'Ocupação Passado',
        'kpis_occupancy%.present.after.value': 'Ocupação Presente',
        'kpis_adr%.past.after.value': 'ADR Passado',
        'kpis_adr%.present.after.value': 'ADR Presente',
        'kpis_revpar%.past.after.value': 'RevPAR Passado',
        'kpis_revpar%.present.after.value': 'RevPAR Presente'
    })
    closing_cols: Dict[str, str] = field(default_factory=lambda: {
        'date.present.calendar_date': 'DataFechamento',
        'kpis_occupancy%.past.after.value': 'Ocupação Passado',
        'kpis_occupancy%.present.after.value': 'Ocupação Presente',
        'kpis_adr%.past.after.value': 'ADR Passado',
        'kpis_adr%.present.after.value': 'ADR Presente',
        'kpis_revenue.past.after.value': 'Receita Passado',
        'kpis_revenue.present.after.value': 'Receita Presente',
        'kpis_revpar%.past.after.value': 'RevPAR Passado',
        'kpis_revpar%.present.after.value': 'RevPAR Presente'
    })
