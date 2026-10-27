"""Exportador PDF do relatório semanal (WeasyPrint).

Reutiliza a identidade visual do relatório mensal (cabeçalho, tabelas, section-header), porém com
capa PRETA seguindo a paleta Revinn (preto + verde #159F92 / #005151).
"""
import os
import sys
from html import escape
from typing import List, Optional, Tuple

import pandas as pd

from src.core.weekly_reader import WeeklyData, _is_vazio, _is_emissao_por_uh, _norm
from src.visualization.weekly_charts import build_benchmark_pages

TEAL = "#159F92"
DARK = "#005151"
MINT = "#7FD1C7"
GRAY_COVER = "#5f6b6d"  # cinza da capa/contracapa (padrão)

# Nomes de exibição das colunas das tabelas dinâmicas
_HEADER_RENAMES = {
    "soma de valor da reserva": "Valor da Reserva",
    "share [%]": "Share (%)",
    "média de los": "LOS Médio",
    "soma de los": "Diárias",
    "média de ads": "ADS Médio",
    "média de adr": "ADR Médio",
    "soma de sequência": "Nº",
    "soma de emissão total": "Emissão Total",
    "média de emissão por uh": "Emissão por UH",
    "média de variação x semana anterior": "Var. x Sem. Anterior",
}


def _resource_path(*parts: str) -> str:
    base = getattr(sys, "_MEIPASS", os.path.abspath("."))
    return os.path.join(base, *parts)


def _fmt_brl(v) -> str:
    return "R$ " + f"{float(v):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def _fmt_num(v, decimals: int = 1) -> str:
    return f"{float(v):,.{decimals}f}".replace(",", "X").replace(".", ",").replace("X", ".")


def _fmt_pct(v) -> str:
    return f"{float(v) * 100:.1f}%".replace(".", ",")


class WeeklyPdfExporter:
    """Gera o PDF do relatório semanal a partir de um WeeklyData."""

    def export(self, data: WeeklyData, output_path: str, property_name: str, reference: str,
               include_benchmark: bool = True, black_cover: bool = False) -> None:
        from weasyprint import HTML  # import tardio: depende das libs GTK

        html = self.build_html(data, property_name, reference, include_benchmark, black_cover)
        tmp_html = output_path + ".temp.html"
        with open(tmp_html, "w", encoding="utf-8") as f:
            f.write(html)
        try:
            HTML(tmp_html).write_pdf(output_path)
        finally:
            if os.path.exists(tmp_html):
                os.remove(tmp_html)

    # ------------------------------------------------------------------ HTML

    def build_html(self, data: WeeklyData, property_name: str, reference: str, include_benchmark: bool,
                   black_cover: bool = False) -> str:
        logo_path = _resource_path("Logo Revinn", "Logo Revinn.png")
        logo_url = "file:///" + logo_path.replace("\\", "/") if os.path.exists(logo_path) else ""
        logo_white_path = _resource_path("Logo Revinn", "Logo Revinn Branco.png")
        logo_white_url = "file:///" + logo_white_path.replace("\\", "/") if os.path.exists(logo_white_path) else logo_url
        # Capa preta usa o logo turquesa original; capa cinza usa o logo branco
        logo_cover_url = logo_url if black_cover else logo_white_url
        year = data.year or ""
        name = escape(property_name)
        ref = escape(reference)

        year_label = str(year) if year else (ref or "2026")

        evolution_df = data.evolution
        if evolution_df is not None and not evolution_df.empty:
            first_col = evolution_df.columns[0]
            evolution_df = evolution_df[~evolution_df[first_col].apply(_is_vazio)].reset_index(drop=True)
            uh_cols = [c for c in evolution_df.columns if _is_emissao_por_uh(c)]
            if uh_cols:
                evolution_df = evolution_df.drop(columns=uh_cols)

        sections = [
            self._table_section("Emissões por Canal", data.by_channel, "Canal"),
            self._table_section("Emissões por Categoria", data.by_category, "Categoria"),
            self._table_section("Evolução Semanal", evolution_df, "Semana"),
            self._future_distribution_section(data.base),
        ]
        benchmark = self._benchmark_html(data) if include_benchmark else ""

        return f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<title>Revinn - Relatório Semanal</title>
<style>
{self._css(black_cover)}
</style>
</head>
<body>

<!-- ===== CAPA (PRETA) ===== -->
<div class="cover-black">
    <div class="cover-frame">
        <div class="cb-top">
            <img src="{logo_cover_url}" style="height: 260px; width: auto;" />
        </div>

        <div class="cb-main">
            <div class="cb-rule"></div>
            <p class="cb-kicker">Relatório Semanal</p>
            <h1 class="cb-hotel">{name}</h1>
            <h2 class="cb-title">Emissão Semanal</h2>
            <p class="cb-ref">{ref}</p>
            <div class="cb-rule"></div>
        </div>

        <div class="cb-footer">
            <p class="cb-privacy-title">TERMO DE CONFIDENCIALIDADE E PRIVACIDADE</p>
            <p class="cb-privacy-text">
                Este documento contém informações comerciais confidenciais e proprietárias da <strong>Revinn Estratégias</strong>
                e do <strong>{name}</strong>. O conteúdo aqui apresentado é destinado exclusivamente para uso interno e não deve
                ser reproduzido, compartilhado ou distribuído a terceiros sem a autorização prévia por escrito.
            </p>
            <p class="cb-date">São Paulo, {year}</p>
        </div>
    </div>
</div>

<div id="running-header">
    <div class="rh-logo"><img src="{logo_url}" /></div>
    <div class="rh-info">
        <h1>{name}</h1>
        <p>Relatório Semanal &bull; {ref}</p>
    </div>
</div>

<div class="container">
    {''.join(sections)}
    {benchmark}
</div>

<!-- ===== CONTRACAPA ===== -->
<div class="contracapa-black">
    <div class="cover-frame" style="justify-content: center; align-items: center; position: relative;">
        <img src="{logo_cover_url}" style="height: 200px; width: auto;" />
        <div class="cb-rule" style="margin-top: 14px;"></div>
        <p class="cb-date" style="position: absolute; bottom: 14mm; left: 0; right: 0; font-size: 8pt;">
            Revinn Estratégias &copy; {year_label} &mdash; Documento confidencial.
        </p>
    </div>
</div>

</body>
</html>"""

    # ------------------------------------------------------------------ tabelas

    def _future_distribution_section(self, base: pd.DataFrame) -> str:
        """Distribuição das reservas emitidas por mês/ano previsto de hospedagem (mesma tabela do mensal)."""
        if base is None or base.empty:
            return ""

        date_col = next((c for c in base.columns if "check-out" in _norm(c) or "checkout" in _norm(c)), None)
        if not date_col:
            date_col = next((c for c in base.columns if "check-in" in _norm(c) or "checkin" in _norm(c)), None)
        if not date_col:
            return ""

        val_col = next((c for c in base.columns if "valor" in _norm(c) or "receita" in _norm(c)), None)
        if not val_col:
            return ""

        los_col = next((c for c in base.columns if "los" in _norm(c) or "noite" in _norm(c) or "diaria" in _norm(c)), None)

        df = base.dropna(subset=[date_col]).copy()
        if df.empty:
            return ""

        df[val_col] = pd.to_numeric(df[val_col], errors="coerce").fillna(0.0)
        df["_los"] = pd.to_numeric(df[los_col], errors="coerce").fillna(0.0) if los_col else pd.Series(0.0, index=df.index)
        df["_mes"] = pd.to_datetime(df[date_col], errors="coerce").dt.to_period("M")
        df = df.dropna(subset=["_mes"])
        if df.empty:
            return ""

        agg = df.groupby("_mes").agg(
            receita=(val_col, "sum"),
            reservas=(val_col, "count"),
            noites=("_los", "sum")
        ).reset_index().sort_values("_mes")

        total_rec = float(agg["receita"].sum())
        total_res = int(agg["reservas"].sum())
        total_noites = float(agg["noites"].sum())
        months = ["Jan", "Fev", "Mar", "Abr", "Mai", "Jun", "Jul", "Ago", "Set", "Out", "Nov", "Dez"]

        rows = []
        for _, r in agg.iterrows():
            share = (r["receita"] / total_rec * 100) if total_rec else 0.0
            adr = (r["receita"] / r["noites"]) if r["noites"] else 0.0
            label = f"{months[r['_mes'].month - 1]}/{r['_mes'].year}"
            share_str = f"{share:.2f}%".replace(".", ",")
            rows.append(f"<tr><td>{label}</td><td>{_fmt_brl(r['receita'])}</td><td>{share_str}</td>"
                        f"<td>{int(r['reservas'])}</td><td>{_fmt_brl(adr)}</td></tr>")

        total_adr = total_rec / total_noites if total_noites else 0.0
        rows.append(f'<tr class="total-row"><td>TOTAL</td><td>{_fmt_brl(total_rec)}</td><td>100,00%</td>'
                    f"<td>{total_res}</td><td>{_fmt_brl(total_adr)}</td></tr>")

        return f"""
    <div class="section-header">Distribuição de Reservas por Previsto de Hospedagem</div>
    <table>
        <thead><tr><th>Mês/Ano</th><th>Receita</th><th>Share de Receita</th><th>Reservas</th><th>Diária Média</th></tr></thead>
        <tbody>{''.join(rows)}</tbody>
    </table>
"""

    def _table_section(self, title: str, df: pd.DataFrame, first_col_name: str) -> str:
        if df is None or df.empty:
            return ""

        # Remove coluna de Emissão por UH caso esteja presente na tabela de Evolução Semanal
        if "evolu" in _norm(title):
            uh_cols = [c for c in df.columns if _is_emissao_por_uh(c)]
            if uh_cols:
                df = df.drop(columns=uh_cols)

        columns = list(df.columns)
        head_cells = []
        for i, col in enumerate(columns):
            key = str(col).strip().lower()
            display = first_col_name if i == 0 else _HEADER_RENAMES.get(key, str(col).strip())
            head_cells.append(f"<th>{escape(display)}</th>")

        body_rows = []
        for _, row in df.iterrows():
            val_0 = row.iloc[0]
            if _is_vazio(val_0):
                continue
            is_total = "total" in str(val_0).lower()
            cells = []
            for i, col in enumerate(columns):
                text, css = self._format_cell(i, col, row[col])
                cells.append(f'<td class="{css}">{escape(text)}</td>' if css else f"<td>{escape(text)}</td>")
            body_rows.append(f'<tr class="total-row">{"".join(cells)}</tr>' if is_total else f"<tr>{''.join(cells)}</tr>")

        return f"""
    <div class="section-header">{escape(title)}</div>
    <table>
        <thead><tr>{''.join(head_cells)}</tr></thead>
        <tbody>
            {''.join(body_rows)}
        </tbody>
    </table>
"""

    @staticmethod
    def _format_cell(idx: int, col: str, value) -> Tuple[str, str]:
        if idx == 0:
            return str(value).strip(), ""
        if value is None or (isinstance(value, float) and pd.isna(value)) or isinstance(value, str):
            return "—", ""  # inclui erros do Excel (#VALUE!, #DIV/0!...)
        key = str(col).lower()
        if "share" in key:
            return _fmt_pct(value), ""
        if "variação" in key or "variacao" in key:
            return _fmt_pct(value), ("positive" if float(value) >= 0 else "negative")
        if any(k in key for k in ("valor", "emissão", "emissao", "adr")):
            return _fmt_brl(value), ""
        if "sequ" in key or "soma de los" in key:
            return f"{int(round(float(value)))}", ""
        return _fmt_num(value, 1), ""

    # ------------------------------------------------------------------ benchmark

    def _benchmark_html(self, data: WeeklyData) -> str:
        pages = build_benchmark_pages(data)
        out = []
        for p_idx, blocks in enumerate(pages):
            header = ""
            if p_idx == 0:
                header = ('<div class="section-header" style="margin-top: 0;">Benchmark de Gráficos</div>'
                          '<p class="bench-note">Visualizações exploratórias para análise e ideias de uso futuro — '
                          'não fazem parte do relatório final.</p>')
            inner = []
            for block in blocks:
                if block[0] == "full":
                    inner.append(self._img(block[1], "100%"))
                else:
                    inner.append(
                        '<table class="chart-grid"><tr>'
                        f'<td>{self._img(block[1], "100%")}</td><td>{self._img(block[2], "100%")}</td>'
                        '</tr></table>')
            out.append(f'<div style="page-break-before: always;">{header}{"".join(inner)}</div>')
        return "".join(out)

    @staticmethod
    def _img(uri: Optional[str], width: str) -> str:
        if not uri:
            return '<div class="chart-missing">Gráfico indisponível para esta base de dados</div>'
        return f'<img class="chart-img" src="{uri}" style="width: {width};" />'

    # ------------------------------------------------------------------ CSS

    @staticmethod
    def _css(black_cover: bool = False) -> str:
        if black_cover:
            c_bg, c_frame, c_rule = "#050707", TEAL, TEAL
            c_kicker, c_title, c_ref = TEAL, "#9fd9d2", "#8e9a9a"
            c_fborder, c_ptitle, c_ptext, c_date = "#1f2a2a", TEAL, "#8e9a9a", "#cfd8d8"
        else:
            c_bg, c_frame, c_rule = GRAY_COVER, "rgba(255,255,255,0.6)", MINT
            c_kicker, c_title, c_ref = MINT, "#d5efec", "#e3e9e9"
            c_fborder, c_ptitle, c_ptext, c_date = "rgba(255,255,255,0.35)", MINT, "#e3e9e9", "#ffffff"
        return f"""
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;600;700&family=Poppins:wght@400;600;700&display=swap');

        @page {{
            size: A4;
            margin: 32mm 12mm 15mm 12mm;
            @top-left {{ content: element(running-header); width: 100%; }}
            @bottom-right {{
                content: "p. " counter(page);
                font-family: 'Inter', sans-serif;
                font-size: 8pt;
                color: #7f8c8d;
            }}
        }}

        /* Capa preta (margem zero + fundo preto em toda a página) */
        @page :first {{
            margin: 0;
            background: {c_bg};
            @top-left {{ content: ""; }}
            @bottom-right {{ content: ""; }}
        }}

        /* Contracapa no final do documento */
        @page contracapa {{
            size: A4;
            margin: 0;
            background: {c_bg};
            @top-left {{ content: ""; }}
            @bottom-right {{ content: ""; }}
        }}

        *, *::before, *::after {{ box-sizing: border-box; }}

        body {{
            font-family: 'Inter', sans-serif;
            color: #2c3e50;
            line-height: 1.4;
            margin: 0;
            padding: 0;
            font-size: 10pt;
        }}

        .cover-black {{
            height: 296mm;
            width: 210mm;
            padding: 10mm;
            page-break-after: always;
            background: {c_bg};
        }}

        .contracapa-black {{
            page: contracapa;
            page-break-before: always;
            break-before: page;
            height: 296mm;
            width: 210mm;
            padding: 10mm;
            background: {c_bg};
        }}

        .cover-frame {{
            height: 100%;
            border: 1px solid {c_frame};
            padding: 22mm 18mm 16mm 18mm;
            display: flex;
            flex-direction: column;
            justify-content: space-between;
            text-align: center;
        }}

        .cb-top {{ text-align: center; }}

        .cb-rule {{
            width: 60px; height: 3px; background: {c_rule};
            margin: 0 auto; border-radius: 2px;
        }}

        .cb-kicker {{
            font-family: 'Inter', sans-serif; font-size: 10pt; font-weight: 600;
            color: {c_kicker}; text-transform: uppercase; letter-spacing: 6px;
            margin: 28px 0 22px 0;
        }}

        .cb-hotel {{
            font-family: 'Poppins', sans-serif; font-size: 28pt; font-weight: 700;
            color: #ffffff; text-transform: uppercase; letter-spacing: 0.5px;
            margin: 0 0 12px 0; line-height: 1.15;
        }}

        .cb-title {{
            font-family: 'Poppins', sans-serif; font-size: 15pt; font-weight: 600;
            color: {c_title}; text-transform: uppercase; letter-spacing: 2px;
            margin: 0 0 6px 0;
        }}

        .cb-ref {{
            font-family: 'Inter', sans-serif; font-size: 12pt; color: {c_ref};
            margin: 0 0 28px 0;
        }}

        .cb-footer {{ border-top: 1px solid {c_fborder}; padding-top: 18px; }}

        .cb-privacy-title {{
            font-family: 'Poppins', sans-serif; font-size: 8pt; font-weight: 700;
            color: {c_ptitle}; margin: 0 0 8px 0; letter-spacing: 0.5px;
        }}

        .cb-privacy-text {{
            font-family: 'Inter', sans-serif; font-size: 7.5pt; color: {c_ptext};
            line-height: 1.5; max-width: 500px; margin: 0 auto 18px auto; text-align: justify;
        }}

        .cb-date {{ font-family: 'Inter', sans-serif; font-size: 9pt; color: {c_date}; margin: 0; }}

        /* Headers de seção */
        .section-header {{
            font-family: 'Poppins', sans-serif; font-size: 10pt; text-transform: uppercase;
            letter-spacing: 0.5px; color: #ffffff; background-color: {DARK};
            padding: 5px 8px; margin-top: 18px; margin-bottom: 10px; font-weight: 700;
            border-radius: 2px;
            page-break-inside: avoid; break-inside: avoid;
            page-break-after: avoid; break-after: avoid;
        }}

        /* Tabelas */
        table {{
            width: 100%; border-collapse: collapse; margin-bottom: 15px;
            font-size: 8.5pt; font-family: 'Inter', sans-serif;
            page-break-inside: auto; break-inside: auto;
        }}
        thead {{
            display: table-header-group;
        }}
        tr {{
            page-break-inside: avoid; break-inside: avoid;
        }}
        th {{
            font-family: 'Poppins', sans-serif; font-weight: 700; text-transform: uppercase;
            font-size: 8pt; color: #ffffff; background-color: {TEAL};
            padding: 6px 8px; text-align: right;
        }}
        th:first-child, td:first-child {{ text-align: left; }}
        td {{
            font-family: 'Inter', sans-serif; padding: 5px 8px;
            border-bottom: 1px solid #e2e8f0; text-align: right; color: #2c3e50;
        }}
        tr:nth-child(even) td {{ background-color: #f8f9fa; }}
        .total-row td {{
            font-weight: 700; background-color: #e8e8e8 !important;
            border-top: 1.5px solid {DARK}; border-bottom: 2px solid {DARK};
        }}
        .negative {{ color: #c0392b; }}
        .positive {{ color: {TEAL}; }}

        /* Benchmark de gráficos */
        .bench-note {{
            font-family: 'Inter', sans-serif; font-size: 8pt; color: #7f8c8d;
            margin: 0 0 10px 0; font-style: italic;
        }}
        .chart-img {{ display: block; margin: 0 auto 8px auto; }}
        .chart-missing {{
            border: 1px dashed #cbd5d5; color: #98a5a5; text-align: center;
            font-size: 8pt; padding: 40px 10px; margin-bottom: 8px;
        }}
        table.chart-grid {{
            margin-bottom: 0;
            page-break-inside: avoid; break-inside: avoid;
        }}
        table.chart-grid td {{
            border: none !important; background: none !important; padding: 0 4px !important;
            text-align: center !important; vertical-align: top !important; width: 50%;
        }}

        /* Running header */
        #running-header {{
            position: running(running-header); display: table; width: 100%;
            border-bottom: 2px solid {TEAL}; padding-bottom: 6px; table-layout: fixed;
        }}
        #running-header .rh-logo {{ display: table-cell; width: 25%; vertical-align: middle; }}
        #running-header .rh-logo img {{ height: 54px; width: auto; max-width: 180px; display: block; }}
        #running-header .rh-info {{ display: table-cell; width: 75%; vertical-align: middle; text-align: right; }}
        #running-header .rh-info h1 {{
            font-family: 'Poppins', sans-serif; font-size: 11pt; font-weight: 700;
            color: {DARK}; text-transform: uppercase; margin: 0;
        }}
        #running-header .rh-info p {{
            font-family: 'Inter', sans-serif; font-size: 8pt; color: #7f8c8d;
            font-weight: 600; margin: 2px 0 0 0;
        }}
        """
