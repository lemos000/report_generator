"""Gráficos 'benchmark' do relatório semanal (ideias de visualização para uso futuro).

Usa apenas a API orientada a objetos do matplotlib (Figure), sem pyplot, para ser seguro em threads.
Todas as funções retornam uma data-URI PNG (base64) pronta para uso em <img src="...">.
"""
import base64
import io
import textwrap
from typing import Callable, List, Optional, Tuple

import numpy as np
import pandas as pd
from matplotlib.figure import Figure
from matplotlib.backends.backend_agg import FigureCanvasAgg

# Paleta Revinn
DARK = "#005151"
TEAL = "#159F92"
MINT = "#7FD1C7"
PALE = "#D5EFEC"
GRAY = "#7f8c8d"
LIGHT_GRAY = "#e2e8f0"
INK = "#2c3e50"
NEG = "#c0392b"
SEQ_COLORS = [DARK, TEAL, MINT, "#9AA5A5", "#C5CDCD", PALE]

FULL = (7.3, 3.5)
HALF = (3.6, 2.8)

_MONTHS_PT = ["Jan", "Fev", "Mar", "Abr", "Mai", "Jun", "Jul", "Ago", "Set", "Out", "Nov", "Dez"]
_WEEKDAYS_PT = ["Seg", "Ter", "Qua", "Qui", "Sex", "Sáb", "Dom"]


# ----------------------------------------------------------------- helpers

def _new_fig(size: Tuple[float, float]) -> Figure:
    # layout 'constrained' mantém dimensões idênticas entre gráficos do mesmo tamanho (sem bbox 'tight')
    fig = Figure(figsize=size, dpi=220, facecolor="white", layout="constrained")
    FigureCanvasAgg(fig)
    return fig


def _style_axes(ax, grid_axis: str = "y") -> None:
    for side in ("top", "right"):
        ax.spines[side].set_visible(False)
    for side in ("left", "bottom"):
        ax.spines[side].set_color(LIGHT_GRAY)
    ax.tick_params(colors=INK, labelsize=6.5, length=0)
    if grid_axis:
        ax.grid(axis=grid_axis, color=LIGHT_GRAY, linewidth=0.6, alpha=0.8)
        ax.set_axisbelow(True)


def _title(ax, title: str, subtitle: str = "") -> None:
    ax.set_title(title, loc="left", fontsize=8.5, fontweight="bold", color=DARK, pad=15 if subtitle else 8)
    if subtitle:
        ax.annotate(subtitle, xy=(0, 1), xycoords="axes fraction", xytext=(0, 3), textcoords="offset points",
                    fontsize=6.3, color=GRAY, va="bottom", ha="left")


def _wrap(text, width: int = 18) -> str:
    return textwrap.fill(str(text), width=width)


def _to_uri(fig: Figure) -> str:
    buf = io.BytesIO()
    fig.savefig(buf, format="png", facecolor="white")
    return "data:image/png;base64," + base64.b64encode(buf.getvalue()).decode()


def _brl_k(v: float) -> str:
    if abs(v) >= 1000:
        return f"R$ {v / 1000:,.1f} mil".replace(",", "X").replace(".", ",").replace("X", ".")
    return f"R$ {v:,.0f}".replace(",", ".")


def _pct(v: float) -> str:
    return f"{v * 100:.1f}%".replace(".", ",")


def _short_week(label: str) -> str:
    label = str(label).strip().replace(" de ", " ")
    parts = label.split(" a ")
    if len(parts) == 2:
        end = parts[1].split(" ")
        if len(end) == 2:
            return f"{parts[0].strip()}–{end[0]} {end[1][:3]}"
    return label


def _is_total(label) -> bool:
    return "total" in str(label).lower()


def _pivot_rows(df: pd.DataFrame, value_col_hint: str = "valor") -> Optional[pd.DataFrame]:
    """Normaliza uma tabela dinâmica (1ª coluna = rótulo, 'valor' = soma) removendo a linha de total."""
    if df is None or df.empty:
        return None
    label_col = df.columns[0]
    val_col = next((c for c in df.columns if value_col_hint in c.lower()), df.columns[1])
    out = df[~df[label_col].apply(_is_total)][[label_col, val_col]].copy()
    out.columns = ["label", "value"]
    out["value"] = pd.to_numeric(out["value"], errors="coerce").fillna(0.0)
    return out.sort_values("value", ascending=False).reset_index(drop=True)


# ----------------------------------------------------------------- charts

def chart_weekly_evolution(history: pd.DataFrame) -> Optional[str]:
    """Barras de emissão semanal + média móvel + variação percentual semana a semana."""
    if history is None or history.empty:
        return None
    seq_col = next((c for c in history.columns if "sequ" in c.lower()), history.columns[0])
    lab_col = next((c for c in history.columns if "interv" in c.lower()), history.columns[1])
    val_col = next((c for c in history.columns if "total" in c.lower()), history.columns[2])
    df = history[[seq_col, lab_col, val_col]].copy()
    df.columns = ["seq", "label", "value"]
    df["value"] = pd.to_numeric(df["value"], errors="coerce")
    df = df.dropna(subset=["value"]).sort_values("seq")
    if df.empty:
        return None

    x = np.arange(len(df))
    values = df["value"].to_numpy()
    avg = values.mean()
    ma = pd.Series(values).rolling(3, min_periods=1).mean().to_numpy()
    var = pd.Series(values).pct_change().to_numpy()

    fig = _new_fig((FULL[0], 3.9))
    gs = fig.add_gridspec(2, 1, height_ratios=[3, 1.1])
    ax = fig.add_subplot(gs[0])
    colors = [TEAL] * len(df)
    colors[-1] = DARK
    ax.bar(x, values, color=colors, width=0.66, zorder=3)
    ax.plot(x, ma, color=GRAY, linewidth=1.3, linestyle="--", marker="o", markersize=2.5, zorder=4, label="Média móvel (3 sem.)")
    ax.axhline(avg, color=NEG, linewidth=0.8, linestyle=":", zorder=2, label=f"Média do período ({_brl_k(avg)})")
    for xi, v in zip(x, values):
        ax.text(xi, v, f"{v / 1000:.0f}k", ha="center", va="bottom", fontsize=5.8, color=INK)
    _style_axes(ax)
    ax.set_xticks(x)
    ax.tick_params(labelbottom=False)
    ax.yaxis.set_major_formatter(lambda v, _: f"{v / 1000:.0f}k")
    ax.legend(loc="upper left", fontsize=6, frameon=False)
    _title(ax, "Evolução da emissão semanal", "Receita das reservas emitidas por semana (R$) e tendência")

    ax2 = fig.add_subplot(gs[1], sharex=ax)
    vcolors = [TEAL if (not np.isnan(v) and v >= 0) else NEG for v in var]
    ax2.bar(x, np.nan_to_num(var, nan=0.0) * 100, color=vcolors, width=0.66, zorder=3)
    ax2.axhline(0, color=GRAY, linewidth=0.6)
    _style_axes(ax2)
    ax2.set_ylabel("Var. %", fontsize=6, color=GRAY)
    ax2.set_xticks(x)
    ax2.set_xticklabels([_short_week(s) for s in df["label"]], rotation=40, ha="right", fontsize=5.8)
    return _to_uri(fig)


def chart_channel_share(by_channel: pd.DataFrame, top: int = 8) -> Optional[str]:
    """Barras horizontais do share de receita por canal (top N + Outros)."""
    rows = _pivot_rows(by_channel)
    if rows is None or rows.empty:
        return None
    total = rows["value"].sum()
    head = rows.head(top)
    other = rows["value"].iloc[top:].sum()
    labels = list(head["label"])
    vals = list(head["value"])
    if other > 0:
        labels.append("Outros")
        vals.append(other)
    labels = [str(l)[:22] for l in labels][::-1]
    vals = vals[::-1]

    fig = _new_fig(HALF)
    ax = fig.add_subplot(111)
    colors = [PALE if l == "Outros" else TEAL for l in labels]
    colors[-1] = DARK if labels[-1] != "Outros" else colors[-1]
    ax.barh(labels, vals, color=colors, height=0.62, zorder=3)
    for i, v in enumerate(vals):
        ax.text(v, i, f" {_pct(v / total)}", va="center", fontsize=6, color=INK)
    _style_axes(ax, "x")
    ax.xaxis.set_major_formatter(lambda v, _: f"{v / 1000:.0f}k")
    ax.set_xlim(0, max(vals) * 1.22)
    _title(ax, "Share de receita por canal", "Top canais da semana")
    return _to_uri(fig)


def chart_category_donut(by_category: pd.DataFrame) -> Optional[str]:
    rows = _pivot_rows(by_category)
    if rows is None or rows.empty:
        return None
    fig = _new_fig(HALF)
    ax = fig.add_subplot(111)
    colors = (SEQ_COLORS * 2)[:len(rows)]
    wedges, _ = ax.pie(rows["value"], colors=colors, startangle=90, counterclock=False,
                       wedgeprops=dict(width=0.38, edgecolor="white", linewidth=1.5))
    total = rows["value"].sum()
    ax.text(0, 0.06, _brl_k(total), ha="center", va="center", fontsize=8, fontweight="bold", color=DARK)
    ax.text(0, -0.14, "emitidos", ha="center", va="center", fontsize=6, color=GRAY)
    legend_labels = [f"{str(l)[:28]} · {_pct(v / total)}" for l, v in zip(rows["label"], rows["value"])]
    ax.legend(wedges, legend_labels, loc="upper center", bbox_to_anchor=(0.5, 0.02), fontsize=5.8,
              frameon=False, ncol=1)
    _title(ax, "Mix por categoria", "Participação na receita emitida")
    return _to_uri(fig)


def chart_weekday(base: pd.DataFrame) -> Optional[str]:
    """Receita (barras) e nº de reservas (linha) por dia da semana da compra."""
    if base is None or base.empty or "Data de Compra" not in base.columns:
        return None
    df = base.dropna(subset=["Data de Compra"]).copy()
    df["dow"] = df["Data de Compra"].dt.dayofweek
    rev = df.groupby("dow")["Valor da Reserva"].sum().reindex(range(7), fill_value=0)
    cnt = df.groupby("dow").size().reindex(range(7), fill_value=0)
    fig = _new_fig(HALF)
    ax = fig.add_subplot(111)
    colors = [DARK if v == rev.max() else TEAL for v in rev]
    ax.bar(_WEEKDAYS_PT, rev.values, color=colors, width=0.62, zorder=3)
    _style_axes(ax)
    ax.yaxis.set_major_formatter(lambda v, _: f"{v / 1000:.0f}k")
    ax2 = ax.twinx()
    ax2.plot(_WEEKDAYS_PT, cnt.values, color=GRAY, marker="o", markersize=3, linewidth=1.2)
    for i, c in enumerate(cnt.values):
        ax2.text(i, c, f"{c}", ha="center", va="bottom", fontsize=5.8, color=GRAY)
    ax2.set_ylim(0, max(cnt.max() * 1.35, 1))
    ax2.spines[["top", "left"]].set_visible(False)
    ax2.spines["right"].set_color(LIGHT_GRAY)
    ax2.tick_params(colors=GRAY, labelsize=6, length=0)
    _title(ax, "Emissão por dia da semana", "Receita (barras) e nº de reservas (linha)")
    return _to_uri(fig)


def chart_lead_time(base: pd.DataFrame) -> Optional[str]:
    """Distribuição da antecedência de compra (ADS) em faixas."""
    if base is None or base.empty or "ADS" not in base.columns:
        return None
    ads = base["ADS"].dropna()
    if ads.empty:
        return None
    bins = [-1, 3, 7, 14, 30, 60, 10_000]
    labels = ["0–3", "4–7", "8–14", "15–30", "31–60", "60+"]
    cut = pd.cut(ads, bins=bins, labels=labels)
    counts = cut.value_counts().reindex(labels, fill_value=0)
    total = counts.sum()
    fig = _new_fig(HALF)
    ax = fig.add_subplot(111)
    colors = [DARK if v == counts.max() else TEAL for v in counts]
    ax.bar(labels, counts.values, color=colors, width=0.66, zorder=3)
    for i, c in enumerate(counts.values):
        ax.text(i, c, f"{_pct(c / total)}", ha="center", va="bottom", fontsize=5.8, color=INK)
    _style_axes(ax)
    ax.axvline(-0.5, color="none")
    ax.set_xlabel("dias de antecedência", fontsize=6, color=GRAY)
    _title(ax, "Antecedência de compra (ADS)", f"Mediana: {ads.median():.0f} dias · Média: {ads.mean():.1f} dias".replace(".", ","))
    return _to_uri(fig)


def chart_los(base: pd.DataFrame) -> Optional[str]:
    if base is None or base.empty or "LOS" not in base.columns:
        return None
    los = base["LOS"].dropna().astype(int).clip(upper=6)
    counts = los.value_counts().sort_index()
    labels = [f"{i}" if i < 6 else "6+" for i in counts.index]
    total = counts.sum()
    fig = _new_fig(HALF)
    ax = fig.add_subplot(111)
    colors = [DARK if v == counts.max() else TEAL for v in counts]
    ax.bar(labels, counts.values, color=colors, width=0.62, zorder=3)
    for i, c in enumerate(counts.values):
        ax.text(i, c, f"{_pct(c / total)}", ha="center", va="bottom", fontsize=5.8, color=INK)
    _style_axes(ax)
    ax.set_xlabel("noites por reserva", fontsize=6, color=GRAY)
    _title(ax, "Permanência (LOS)", "Reservas por número de noites")
    return _to_uri(fig)


def chart_rate_plan(base: pd.DataFrame, top: int = 6) -> Optional[str]:
    if base is None or base.empty or "Tarifário" not in base.columns:
        return None
    rev = base.groupby("Tarifário")["Valor da Reserva"].sum().sort_values(ascending=False)
    if rev.empty:
        return None
    total = rev.sum()
    head = rev.head(top)
    labels = [_wrap(l, 24) for l in head.index][::-1]
    vals = list(head.values)[::-1]
    fig = _new_fig(HALF)
    ax = fig.add_subplot(111)
    ax.barh(labels, vals, color=[TEAL] * (len(vals) - 1) + [DARK], height=0.6, zorder=3)
    for i, v in enumerate(vals):
        ax.text(v, i, f" {_pct(v / total)}", va="center", fontsize=6, color=INK)
    _style_axes(ax, "x")
    ax.xaxis.set_major_formatter(lambda v, _: f"{v / 1000:.0f}k")
    ax.set_xlim(0, max(vals) * 1.22)
    ax.tick_params(axis="y", labelsize=5.6)
    _title(ax, "Mix de tarifários", "Participação na receita emitida")
    return _to_uri(fig)


def chart_channel_category_heatmap(base: pd.DataFrame, top: int = 8) -> Optional[str]:
    """Heatmap Canal x Categoria (receita)."""
    if base is None or base.empty or not {"Canal", "Categoria"}.issubset(base.columns):
        return None
    pv = base.pivot_table(index="Canal", columns="Categoria", values="Valor da Reserva", aggfunc="sum", fill_value=0)
    pv = pv.loc[pv.sum(axis=1).sort_values(ascending=False).index].head(top)
    if pv.empty:
        return None
    data = pv.values
    fig = _new_fig((FULL[0], 3.4))
    ax = fig.add_subplot(111)
    from matplotlib.colors import LinearSegmentedColormap
    cmap = LinearSegmentedColormap.from_list("revinn", ["#ffffff", MINT, TEAL, DARK])
    im = ax.imshow(data, cmap=cmap, aspect="auto")
    ax.set_xticks(range(pv.shape[1]))
    ax.set_xticklabels([_wrap(c, 20) for c in pv.columns], fontsize=6.3)
    ax.set_yticks(range(pv.shape[0]))
    ax.set_yticklabels([str(i)[:24] for i in pv.index], fontsize=6.3)
    ax.tick_params(length=0, colors=INK)
    for s in ax.spines.values():
        s.set_visible(False)
    vmax = data.max() if data.max() > 0 else 1
    for i in range(data.shape[0]):
        for j in range(data.shape[1]):
            v = data[i, j]
            if v > 0:
                ax.text(j, i, f"{v / 1000:.1f}k", ha="center", va="center", fontsize=6,
                        color="white" if v / vmax > 0.45 else INK)
    _title(ax, "Receita por canal × categoria", "Onde cada canal concentra a venda (R$ mil)")
    return _to_uri(fig)


def chart_checkin_months(base: pd.DataFrame) -> Optional[str]:
    """Receita emitida por mês de check-in (para onde a venda da semana está indo)."""
    if base is None or base.empty or "Check-in" not in base.columns:
        return None
    df = base.dropna(subset=["Check-in"]).copy()
    df["ym"] = df["Check-in"].dt.to_period("M")
    rev = df.groupby("ym")["Valor da Reserva"].sum().sort_index()
    if rev.empty:
        return None
    labels = [f"{_MONTHS_PT[p.month - 1]}/{str(p.year)[2:]}" for p in rev.index]
    total = rev.sum()
    fig = _new_fig(HALF)
    ax = fig.add_subplot(111)
    ax.bar(labels, rev.values, color=[DARK if v == rev.max() else TEAL for v in rev], width=0.62, zorder=3)
    for i, v in enumerate(rev.values):
        ax.text(i, v, f"{_pct(v / total)}", ha="center", va="bottom", fontsize=5.8, color=INK)
    _style_axes(ax)
    ax.yaxis.set_major_formatter(lambda v, _: f"{v / 1000:.0f}k")
    _title(ax, "Mês de hospedagem (check-in)", "Para quando a venda da semana foi feita")
    return _to_uri(fig)


def chart_adr_los_bubble(base: pd.DataFrame, top: int = 9) -> Optional[str]:
    """Bolhas por canal: LOS médio (x) × ADR médio (y), tamanho = receita."""
    needed = {"Canal", "Valor da Reserva", "LOS"}
    if base is None or base.empty or not needed.issubset(base.columns):
        return None
    g = base.groupby("Canal").agg(rev=("Valor da Reserva", "sum"), los=("LOS", "sum"), n=("LOS", "size"))
    g["adr"] = g["rev"] / g["los"].replace(0, np.nan)
    g["los_avg"] = g["los"] / g["n"]
    g = g.sort_values("rev", ascending=False).head(top).dropna(subset=["adr"])
    if g.empty:
        return None
    fig = _new_fig(HALF)
    ax = fig.add_subplot(111)
    sizes = 40 + 360 * (g["rev"] / g["rev"].max())
    ax.scatter(g["los_avg"], g["adr"], s=sizes, color=TEAL, alpha=0.75, edgecolor=DARK, linewidth=0.8, zorder=3)
    for name, row in g.iterrows():
        ax.annotate(str(name)[:14], (row["los_avg"], row["adr"]), xytext=(4, 4), textcoords="offset points",
                    fontsize=5.3, color=INK)
    _style_axes(ax, "both")
    ax.set_xlabel("LOS médio (noites)", fontsize=6, color=GRAY)
    ax.set_ylabel("ADR (R$)", fontsize=6, color=GRAY)
    _title(ax, "Eficiência por canal", "ADR × permanência (bolha = receita)")
    return _to_uri(fig)


# ----------------------------------------------------------------- orchestration

ChartFn = Callable[[], Optional[str]]


def build_benchmark_pages(data) -> List[List[Optional[str]]]:
    """Retorna as páginas de gráficos: cada página é uma lista de imagens (None = indisponível).

    Layout de cada página: [largura total] ou [esq, dir] pares — ver weekly_exporter.
    """
    def safe(fn, *args) -> Optional[str]:
        try:
            return fn(*args)
        except Exception as exc:  # um gráfico com problema não deve derrubar o relatório
            print(f"[benchmark] falha em {fn.__name__}: {exc}")
            return None

    return [
        # Página 1
        [("full", safe(chart_weekly_evolution, data.history)),
         ("pair", safe(chart_channel_share, data.by_channel), safe(chart_category_donut, data.by_category))],
        # Página 2
        [("pair", safe(chart_weekday, data.base), safe(chart_lead_time, data.base)),
         ("pair", safe(chart_los, data.base), safe(chart_rate_plan, data.base))],
        # Página 3
        [("full", safe(chart_channel_category_heatmap, data.base)),
         ("pair", safe(chart_checkin_months, data.base), safe(chart_adr_los_bubble, data.base))],
    ]

