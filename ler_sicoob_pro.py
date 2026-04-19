import streamlit as st
import fitz
import pandas as pd
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils import get_column_letter
import io
from datetime import datetime
import warnings
import time

# Ignorar avisos de depreciação do Pandas
warnings.simplefilter(action='ignore', category=FutureWarning)

# Configuração da Página
st.set_page_config(
    page_title="Consolidador SICOOB Pro", 
    page_icon="📊", 
    layout="wide",
    initial_sidebar_state="collapsed"
)

# ==========================================
# CSS CUSTOMIZADO PARA CARDS E INTERFACE
# ==========================================
st.markdown("""
<style>
    /* Estilização Geral do Container */
    .block-container {
        padding-top: 2rem;
        padding-bottom: 2rem;
    }

    /* Estilização dos Cards de Métricas */
    [data-testid="stMetric"] {
        background-color: #ffffff;
        border-radius: 12px;
        padding: 20px;
        box-shadow: 0 4px 12px rgba(0, 0, 0, 0.05);
        border-left: 6px solid #00ae9d; /* Ciano SICOOB */
        transition: transform 0.2s ease;
    }
    
    [data-testid="stMetric"]:hover {
        transform: translateY(-4px);
        box-shadow: 0 8px 20px rgba(0, 0, 0, 0.1);
    }

    /* Ajuste de Fontes das Métricas */
    [data-testid="stMetricLabel"] {
        font-size: 14px !important;
        font-weight: 600 !important;
        color: #64748b !important;
    }

    [data-testid="stMetricValue"] {
        font-size: 28px !important;
        font-weight: 700 !important;
        color: #003641 !important;
    }

    /* Limpeza de elementos padrão */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
    
    /* Botão Primário */
    div.stButton > button:first-child {
        border-radius: 8px;
        font-weight: 600;
        padding: 0.5rem 2rem;
    }
</style>
""", unsafe_allow_html=True)

# ==========================================
# CABEÇALHO COM ÍCONE SVG PROFISSIONAL
# ==========================================
col_logo, col_titulo = st.columns([0.5, 9.5])
with col_logo:
    st.markdown("""
        <div style='margin-top: 5px;'>
            <svg width="45" height="45" viewBox="0 0 24 24" fill="none" stroke="#00ae9d" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path>
                <polyline points="14 2 14 8 20 8"></polyline>
                <line x1="16" y1="13" x2="8" y2="13"></line>
                <line x1="16" y1="17" x2="8" y2="17"></line>
                <polyline points="10 9 9 9 8 9"></polyline>
            </svg>
        </div>
    """, unsafe_allow_html=True)
with col_titulo:
    st.markdown("<h1 style='color: #003641; margin-bottom: 0px;'>Consolidador SICOOB Pro</h1>", unsafe_allow_html=True)
    st.markdown("<p style='color: #64748b;'>Sistema Inteligente de Auditoria e Consolidação de Liquidações</p>", unsafe_allow_html=True)

st.write("")

# Organização por Abas
tab1, tab2 = st.tabs(["📤 Importação e Resumo", "🔎 Auditoria Detalhada"])

with tab1:
    arquivos_pdf = st.file_uploader("Selecione ou arraste os extratos em PDF", type="pdf", accept_multiple_files=True)

    if arquivos_pdf:
        if st.button("Processar Documentos", type="primary"):
            
            progress_bar = st.progress(0)
            status_text = st.empty()
            dados_totais = []
            
            # Processamento com Barra de Progresso
            for i, arquivo in enumerate(arquivos_pdf):
                status_text.text(f"A ler: {arquivo.name}...")
                try:
                    doc = fitz.open(stream=arquivo.read(), filetype="pdf")
                    for page in doc:
                        for t in page.find_tables():
                            df_temp = t.to_pandas()
                            if any('Sacado' in col for col in df_temp.columns):
                                df_temp['Origem'] = arquivo.name
                                dados_totais.append(df_temp)
                    doc.close()
                except Exception:
                    st.error(f"Erro ao processar o ficheiro: {arquivo.name}")
                
                progress_bar.progress((i + 1) / len(arquivos_pdf))
            
            status_text.empty()
            progress_bar.empty()

            if not dados_totais:
                st.warning("Nenhuma tabela de liquidação válida foi encontrada nos PDFs.")
            else:
                # Consolidação e Limpeza
                df = pd.concat(dados_totais, ignore_index=True)
                
                # Mapeamento e Filtro de Colunas
                mapa = {
                    'Sacado': 'Sacado', 'Nosso Número': 'Nosso_Numero', 'Seu Número': 'Seu_Numero',
                    'Dt. Previsão Crédito': 'Prev_Credito', 'Vencimento': 'Vencimento',
                    'Valor (R$)': 'Valor_Original', 'Vlr. Mora': 'Mora', 'Vlr. Cobrado': 'Valor_Cobrado', 
                    'Dt. Liquid.': 'Dt_Liquidacao', 'Origem': 'Origem'
                }
                cols_presentes = [c for c in mapa.keys() if c in df.columns]
                df = df[cols_presentes].copy()
                df.rename(columns=mapa, inplace=True)

                # Tratamento de Dados
                if 'Sacado' in df.columns:
                    df['Sacado'] = df['Sacado'].str.replace('\n', ' ').str.replace(r'\d{11,14}', '', regex=True).str.strip()
                
                # Conversão Numérica
                cols_financeiras = ['Valor_Original', 'Mora', 'Valor_Cobrado']
                for col in cols_financeiras:
                    if col in df.columns:
                        df[col] = df[col].astype(str).str.replace('.', '', regex=False).str.replace(',', '.', regex=False)
                        df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

                # Remoção de Duplicados (Baseado no Nosso Número)
                if 'Nosso_Numero' in df.columns:
                    df.dropna(subset=['Nosso_Numero'], inplace=True)
                    df.drop_duplicates(subset=['Nosso_Numero'], inplace=True)

                # ==========================================
                # EXIBIÇÃO DOS CARDS DE MÉTRICAS
                # ==========================================
                st.write("")
                c1, c2, c3 = st.columns(3)
                
                total_liquidado = df['Valor_Cobrado'].sum()
                mora_total = df['Mora'].sum()
                
                # Formatação Brasileira
                fmt_valor = f"R$ {total_liquidado:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
                fmt_mora = f"R$ {mora_total:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')

                c1.metric("Títulos Auditados", f"{len(df)}")
                c2.metric("Total Liquidado", fmt_valor)
                c3.metric("Total de Mora", fmt_mora)

                # Geração do Excel Profissional
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df.to_excel(writer, sheet_name='Liquidações', index=False, startrow=4)
                    ws = writer.book['Liquidações']
                    
                    # Cabeçalho do Excel
                    ws['A1'] = "CONSOLIDAÇÃO DE LIQUIDAÇÕES - SICOOB"
                    ws['A1'].font = Font(size=14, bold=True, color="003641")
                    ws['A2'] = f"Extraído em: {datetime.now().strftime('%d/%m/%Y %H:%M')}"
                    
                    # Tabela Dinâmica
                    tab_ref = f"A5:{get_column_letter(df.shape[1])}{5 + len(df)}"
                    tab = Table(displayName="DadosSicoob", ref=tab_ref)
                    tab.tableStyleInfo = TableStyleInfo(name="TableStyleMedium9", showRowStripes=True)
                    ws.add_table(tab)

                st.write("")
                st.success("Consolidação concluída com sucesso!")
                
                col_down1, col_down2, col_down3 = st.columns([1, 2, 1])
                with col_down2:
                    st.download_button(
                        label="📥 BAIXAR RELATÓRIO CONSOLIDADO (.XLSX)",
                        data=output.getvalue(),
                        file_name=f"SICOOB_CONSOLIDADO_{datetime.now().strftime('%Y%m%d')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        type="primary",
                        use_container_width=True
                    )

                with tab2:
                    st.dataframe(df, use_container_width=True, hide_index=True)
