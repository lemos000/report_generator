import os
from typing import Any, Dict, Optional, Tuple

import pandas as pd

from src.core.models import AuditSchema, PickupSchema, MONTH_MAP, DEFAULT_MONTH
from src.core.readers import AuditCsvReader, PickupExcelReader
from src.core.utils import clean_currency, parse_date_flex


class DataProcessor:
    """Orquestra a limpeza e o enriquecimento dos dados de auditoria (CSV).

    A leitura do arquivo é delegada a um reader injetável e o schema é configurável.
    """

    REQUIRED_COLS = AuditSchema().columns

    def __init__(self, reader: Optional[AuditCsvReader] = None, schema: Optional[AuditSchema] = None):
        self.reader = reader or AuditCsvReader()
        self.schema = schema or AuditSchema()

    @staticmethod
    def clean_currency(value: Any) -> float:
        return clean_currency(value)

    def process_audit_data(self, file_path: str) -> pd.DataFrame:
        name = os.path.basename(file_path)
        df = self.reader.read(file_path)
        df = self._select_and_rename(df, name)
        df = self._drop_invalid_rows(df, name)
        df = self._parse_values(df, name)
        df = self._add_derived_metrics(df)
        return df

    def _select_and_rename(self, df: pd.DataFrame, name: str) -> pd.DataFrame:
        cols = self.schema.columns
        missing = [col for col in cols.keys() if col not in df.columns]
        if missing:
            raise KeyError(f"Colunas ausentes em {name}: {', '.join(missing)}")
        return df[list(cols.keys())].rename(columns=cols)

    @staticmethod
    def _drop_invalid_rows(df: pd.DataFrame, name: str) -> pd.DataFrame:
        df = df.dropna(subset=['Canal', 'Receita'])
        df = df[~df['Canal'].astype(str).str.contains('Total', case=False, na=False)]
        if df.empty:
            raise ValueError(f"O arquivo {name} está vazio após a limpeza.")
        return df

    @staticmethod
    def _parse_values(df: pd.DataFrame, name: str) -> pd.DataFrame:
        df['Receita'] = df['Receita'].apply(clean_currency)
        df['CheckIn'] = parse_date_flex(df['CheckIn'])
        df['CheckOut'] = parse_date_flex(df['CheckOut'])
        df = df.dropna(subset=['CheckIn', 'CheckOut'])
        if df.empty:
            raise ValueError(f"Nenhuma data válida em {name}.")
        df['DataReserva'] = parse_date_flex(df['DataReserva'])
        return df

    @staticmethod
    def _add_derived_metrics(df: pd.DataFrame) -> pd.DataFrame:
        df['LOS_Calc'] = (df['CheckOut'] - df['CheckIn']).dt.days
        df['LOS_Calc'] = df['LOS_Calc'].apply(lambda x: x if x > 0 else 1)

        df['LeadTime'] = (df['CheckIn'] - df['DataReserva']).dt.days.fillna(0)
        df['LeadTime'] = df['LeadTime'].apply(lambda x: x if x >= 0 else 0)

        df['IsCancelled'] = df['Status'].astype(str).str.contains('cancelada', case=False, na=False).astype(int)
        return df

    def get_summary_metrics(self, df: pd.DataFrame) -> Dict[str, Any]:
        return {
            'total_reservas': len(df),
            'total_receita': df['Receita'].sum(),
            'total_noites': df['LOS_Calc'].sum(),
            'total_canceladas': df['IsCancelled'].sum()
        }


class PickupProcessor:
    """Orquestra o processamento dos dados de Pickup (evolução diária e fechamento mensal)."""

    def __init__(self, reader: Optional[PickupExcelReader] = None, schema: Optional[PickupSchema] = None):
        self.reader = reader or PickupExcelReader()
        self.schema = schema or PickupSchema()

    def process_pickup_evolution(self, file_path: str, selected_month: str) -> Tuple[Optional[pd.DataFrame], Optional[pd.DataFrame]]:
        if not os.path.exists(file_path):
            return None, None

        month_num = MONTH_MAP.get(selected_month, DEFAULT_MONTH)

        try:
            df_evol_raw, df_closing_raw = self.reader.read(file_path)
        except Exception as e:
            print(f"Erro ao ler Excel {file_path}: {e}")
            return None, None

        df_evol = self._build_evolution(df_evol_raw, month_num)
        df_closing = self._build_closing(df_closing_raw, month_num)
        return df_evol, df_closing

    # ---------- helpers compartilhados ----------

    @staticmethod
    def _filter_total_segment(df: pd.DataFrame) -> pd.DataFrame:
        """Mantém apenas linhas com 'segment' vazio (NaN ou string em branco)."""
        if 'segment' in df.columns:
            return df[df['segment'].isna() | (df['segment'].astype(str).str.strip() == '')]
        return df

    @staticmethod
    def _select_and_rename(df: pd.DataFrame, mapping: Dict[str, str]) -> pd.DataFrame:
        existing = {raw: new for raw, new in mapping.items() if raw in df.columns}
        out = df[list(existing.keys())].rename(columns=existing)
        for col in mapping.values():
            if col not in out.columns:
                out[col] = 0.0
        return out

    # ---------- evolução ----------

    def _build_evolution(self, df_raw: pd.DataFrame, month_num: int) -> pd.DataFrame:
        mapping = self.schema.evolution_cols
        df = self._select_and_rename(self._filter_total_segment(df_raw), mapping)

        df['Data'] = pd.to_datetime(df['Data'], errors='coerce')
        df = df.dropna(subset=['Data'])
        df = df[df['Data'].dt.month == month_num]

        if df.empty:
            return pd.DataFrame(columns=mapping.values())

        agg_rules = {col: 'mean' for col in df.columns if col != 'Data'}
        return df.groupby('Data').agg(agg_rules).reset_index().sort_values('Data')

    # ---------- fechamento ----------

    def _build_closing(self, df_raw: pd.DataFrame, month_num: int) -> pd.DataFrame:
        mapping = self.schema.closing_cols
        df = self._select_and_rename(self._filter_total_segment(df_raw), mapping)
        df['DataFechamento'] = pd.to_datetime(df['DataFechamento'], errors='coerce')

        df_month = df[df['DataFechamento'].dt.month == month_num]

        if df_month.empty:
            # Fallback: soma/média de todo o período do closing
            agg_rules = {col: ('sum' if 'Receita' in col else 'mean') for col in df.columns if col != 'DataFechamento'}
            return df.drop(columns=['DataFechamento']).agg(agg_rules).to_frame().T

        result = df_month.iloc[[0]].copy()
        if 'DataFechamento' in result.columns:
            result = result.drop(columns=['DataFechamento'])
        return result
