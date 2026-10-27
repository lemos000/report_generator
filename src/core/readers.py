import io
import os
from typing import Tuple

import pandas as pd

from src.core.utils import read_lines_tolerant


class AuditCsvReader:
    """Responsável apenas por ler o CSV de auditoria e devolver um DataFrame bruto."""

    @staticmethod
    def _sanitize_lines(lines):
        clean_lines = []
        for i, line in enumerate(lines):
            clean_line = line.strip()

            # Remove o delimitador residual no final de cada linha
            if clean_line.endswith(';;;;'):
                clean_line = clean_line[:-4]

            # Remove as aspas globais e normaliza as aspas duplas internas (ex: "2.172,74")
            if i > 0 and clean_line.startswith('"') and clean_line.endswith('"'):
                clean_line = clean_line[1:-1].replace('""', '"')

            clean_lines.append(clean_line)
        return clean_lines

    def read(self, file_path: str) -> pd.DataFrame:
        lines = read_lines_tolerant(file_path)
        clean_lines = self._sanitize_lines(lines)
        try:
            return pd.read_csv(io.StringIO('\n'.join(clean_lines)))
        except Exception as e:
            raise ValueError(f"Erro no parsing estrutural do CSV {os.path.basename(file_path)}: {str(e)}")


class PickupExcelReader:
    """Responsável apenas por localizar e ler as planilhas de evolução e fechamento do Excel."""

    @staticmethod
    def _find_sheet(sheet_names, token: str, fallback: str, exclude_suffix: str = None) -> str:
        for name in sheet_names:
            if token in name and not (exclude_suffix and name.endswith(exclude_suffix)):
                return name
        return fallback

    def read(self, file_path: str) -> Tuple[pd.DataFrame, pd.DataFrame]:
        """Retorna (evolução_bruta, fechamento_bruto). Pode lançar exceção em caso de falha de leitura."""
        xl = pd.ExcelFile(file_path)
        evol_sheet = self._find_sheet(xl.sheet_names, 'evolution_table', xl.sheet_names[0])
        pickup_sheet = self._find_sheet(xl.sheet_names, 'pickup_table', xl.sheet_names[-1], exclude_suffix='_piv')
        return xl.parse(evol_sheet), xl.parse(pickup_sheet)
