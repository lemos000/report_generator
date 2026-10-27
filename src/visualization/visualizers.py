import matplotlib.pyplot as plt
import io
import pandas as pd
from typing import List

class ChartGenerator:
    """Classe responsável por gerar gráficos a partir dos dados processados."""
    
    @staticmethod
    def generate_pickup_charts(df_evol: pd.DataFrame, df_closing: pd.DataFrame) -> List[io.BytesIO]:
        if df_evol is None or df_evol.empty:
            return []
        
        charts = []
        
        # --- Gráfico 1: Evolução Diária ---
        days = [d.strftime('%d/%m') for d in df_evol['Data']]
        fig, ax = plt.subplots(figsize=(12, 5))
        
        # Plot lines
        line_past = ax.plot(days, df_evol['Ocupação Passado'], label='Ano Passado', color='#D3D3D3', marker='o', linestyle='--', linewidth=2)
        line_pres = ax.plot(days, df_evol['Ocupação Presente'], label='Ano Atual', color='#159F92', marker='o', linewidth=2)
        
        # Interval logic for labels and ticks (approx. every 3-5 days depending on total)
        total_points = len(days)
        interval = max(1, total_points // 8) # Show ~8 points
        
        # Adjust Ticks
        ax.set_xticks(range(0, total_points, interval))
        ax.set_xticklabels([days[i] for i in range(0, total_points, interval)], rotation=0)

        # Add Data Labels with interval and offset to avoid overlap
        for i in range(0, total_points, interval):
            # Ano Passado (Grey) - Label below
            ax.annotate(f"{df_evol['Ocupação Passado'].iloc[i]:.1f}%", 
                        (days[i], df_evol['Ocupação Passado'].iloc[i]),
                        textcoords="offset points", xytext=(0,-15), ha='center', 
                        fontsize=8, color='#808080', fontweight='bold')
            
            # Ano Atual (Black) - Label above
            ax.annotate(f"{df_evol['Ocupação Presente'].iloc[i]:.1f}%", 
                        (days[i], df_evol['Ocupação Presente'].iloc[i]),
                        textcoords="offset points", xytext=(0,10), ha='center', 
                        fontsize=9, color='#005151', fontweight='bold')

        ax.set_title('Evolução Diária de Ocupação (%) - Ano vs Ano', fontsize=12, weight='bold')
        ax.legend()
        ax.grid(True, alpha=0.3)
        
        buf1 = io.BytesIO()
        plt.savefig(buf1, format='png', bbox_inches='tight', dpi=120)
        buf1.seek(0)
        charts.append(buf1)
        plt.close()
        
        # --- Gráfico 2: Fechamento Ocupação (%) ---
        metrics_occ = ['Ocupação (%)']
        past_occ = [df_closing['Ocupação Passado'].iloc[0]]
        pres_occ = [df_closing['Ocupação Presente'].iloc[0]]
        
        fig, ax = plt.subplots(figsize=(6, 5))
        x = range(len(metrics_occ))
        width = 0.3
        bar_past = ax.bar([i - width/2 for i in x], past_occ, width, label='Ano Passado', color='#D3D3D3')
        bar_pres = ax.bar([i + width/2 for i in x], pres_occ, width, label='Ano Atual', color='#159F92')
        
        # Add labels to bars
        for bar in bar_past:
            ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 0.5, f"{bar.get_height():.2f}%", ha='center', fontweight='bold', color='#808080')
        for bar in bar_pres:
            ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 0.5, f"{bar.get_height():.2f}%", ha='center', fontweight='bold', color='#159F92')

        ax.set_title('Fechamento: Ocupação (%)', fontsize=12, weight='bold')
        ax.set_xticks(x)
        ax.set_xticklabels(metrics_occ)
        ax.legend()
        
        buf2 = io.BytesIO()
        plt.savefig(buf2, format='png', bbox_inches='tight', dpi=120)
        buf2.seek(0)
        charts.append(buf2)
        plt.close()

        # --- Gráfico 3: Fechamento ADR (R$) ---
        metrics_adr = ['ADR (R$)']
        past_adr = [df_closing['ADR Passado'].iloc[0]]
        pres_adr = [df_closing['ADR Presente'].iloc[0]]
        
        fig, ax = plt.subplots(figsize=(6, 5))
        bar_past_adr = ax.bar([i - width/2 for i in x], past_adr, width, label='Ano Passado', color='#D3D3D3')
        bar_pres_adr = ax.bar([i + width/2 for i in x], pres_adr, width, label='Ano Atual', color='#159F92')
        
        # Add labels to bars (ADR)
        for bar in bar_past_adr:
            ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 5, f"R$ {bar.get_height():.0f}", ha='center', fontweight='bold', color='#808080')
        for bar in bar_pres_adr:
            ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 5, f"R$ {bar.get_height():.0f}", ha='center', fontweight='bold', color='#159F92')

        ax.set_title('Fechamento: ADR (R$)', fontsize=12, weight='bold')
        ax.set_xticks(x)
        ax.set_xticklabels(metrics_adr)
        ax.legend()
        
        buf3 = io.BytesIO()
        plt.savefig(buf3, format='png', bbox_inches='tight', dpi=120)
        buf3.seek(0)
        charts.append(buf3)
        plt.close()
        
        return charts

class TableGenerator:
    """Classe responsável por transformar DataFrames em tabelas formatadas."""
    
    @staticmethod
    def generate_audit_reports(df_co: pd.DataFrame, df_re: pd.DataFrame):
        # Relatório 1: Por Categoria
        rel1 = df_co.groupby('Categoria').agg({
            'Receita': 'sum',
            'LOS_Calc': ['sum', 'mean']
        }).reset_index()
        rel1.columns = ['Categoria', 'Receita', 'LOS_Total', 'LOS_Media']
        rel1['ADR'] = (rel1['Receita'] / rel1['LOS_Total']).fillna(0)
        rel1 = rel1.rename(columns={'LOS_Media': 'LOS', 'LOS_Total': 'RN'})[['Categoria', 'ADR', 'LOS', 'RN', 'Receita']]
        rel1 = rel1.sort_values(by='ADR', ascending=False)

        # Relatório 2 & 3: Por Canal
        def group_by_canal(df):
            total_receita_global = df['Receita'].sum()
            agg = df.groupby('Canal').agg({
                'Receita': 'sum',
                'LOS_Calc': 'sum',
                'IsCancelled': 'sum',
                'Status': 'count',
                'LeadTime': 'sum'
            }).rename(columns={'Status': 'Reservas'}).reset_index()
            
            agg = agg.sort_values(by='Receita', ascending=False)
            
            agg['LOS'] = (agg['LOS_Calc'] / agg['Reservas']).fillna(0)
            agg['ADR'] = (agg['Receita'] / agg['LOS_Calc']).fillna(0)
            agg['ADS'] = (agg['LeadTime'] / agg['Reservas']).fillna(0)
            agg['Share (%)'] = (agg['Receita'] / total_receita_global * 100).fillna(0)

            return agg[['Canal', 'Receita', 'Reservas', 'LOS', 'ADR', 'ADS', 'Share (%)']]

        return rel1, group_by_canal(df_co), group_by_canal(df_re)

    @staticmethod
    def generate_booking_distribution(df: pd.DataFrame, date_col: str = 'DataReserva') -> pd.DataFrame:
        """
        Gera tabela de distribuição de reservas por Mês/Ano de uma coluna de data.
        Colunas: Mês/Ano, Receita, Share de Receita (%), Reservas, Diária Média
        date_col: 'DataReserva' para distribuição de emissão, 'CheckOut' para previsto de hospedagem.
        """
        df_valid = df.dropna(subset=[date_col]).copy()
        df_valid['MesAno'] = df_valid[date_col].dt.to_period('M')

        agg = df_valid.groupby('MesAno').agg(
            Receita=('Receita', 'sum'),
            Reservas=('Receita', 'count'),
            Noites=('LOS_Calc', 'sum')
        ).reset_index()

        total_receita = agg['Receita'].sum()

        agg['Share de Receita (%)'] = (agg['Receita'] / total_receita * 100).fillna(0)
        agg['Diária Média'] = (agg['Receita'] / agg['Noites'].replace(0, 1)).fillna(0)
        agg['Mês/Ano'] = agg['MesAno'].dt.strftime('%b/%Y')
        agg = agg.sort_values('MesAno')

        return agg[['Mês/Ano', 'Receita', 'Share de Receita (%)', 'Reservas', 'Diária Média']]
