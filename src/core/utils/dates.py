import pandas as pd


def parse_date_flex(series: pd.Series, primary_fmt: str = '%d/%m/%Y') -> pd.Series:
    """Converte datas com formato explícito; fallback para inferência (dayfirst) nas que falharem."""
    parsed = pd.to_datetime(series, format=primary_fmt, errors='coerce')
    mask_nat = parsed.isna()
    if mask_nat.any():
        parsed[mask_nat] = pd.to_datetime(series[mask_nat], dayfirst=True, errors='coerce')
    return parsed
