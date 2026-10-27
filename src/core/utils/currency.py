from typing import Any

import pandas as pd


def clean_currency(value: Any) -> float:
    """Converte valores monetários no formato BR (ex: '1.250,00 BRL') para float."""
    if pd.isna(value) or not isinstance(value, str):
        return 0.0
    clean_val = value.replace(' BRL', '').replace('.', '').replace(',', '.')
    try:
        return float(clean_val)
    except ValueError:
        return 0.0
