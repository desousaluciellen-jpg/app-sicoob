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

warnings.simplefilter(action='ignore', category=FutureWarning)

st.set_page_config(page_title="Consolidador SICOOB", page_icon="📊", layout="wide")

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
    st.markdown("<h1 style='margin-bottom: 0px;'>Consolidador SICOOB Pro</h1>", unsafe_allow_html=True)
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
                            if any('Sacado' in col for col in df_temp.columns) and any('Valor (R$)' in col for col in df_temp.columns):
                                df_temp['Arquivo_Origem'] = arquivo.name
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
                # SEU CÓDIGO ORIGINAL DE PREPARAÇÃO DOS DADOS
                df_raw = pd.concat(dados_totais, ignore_index=True)

                cols_esperadas = ['Sacado','Nosso Número','Seu Número','Dt. Previsão Crédito',
                                  'Vencimento','Dt. Limite\nPgto','Valor (R$)','Vlr. Mora',
                                  'Vlr. Desc.','Vlr. Outros\nAcresc.','Dt. Liquid.','Vlr. Cobrado', 'Arquivo_Origem']
                cols_presentes = [c for c in cols_esperadas if c in df_raw.columns]
                df = df_raw[cols_presentes].copy()

                mapa_nomes = {
                    'Sacado': 'Sacado', 'Nosso Número': 'Nosso_Numero', 'Seu Número': 'Seu_Numero',
                    'Dt. Previsão Crédito': 'Prev_Credito', 'Vencimento': 'Vencimento',
                    'Dt. Limite\nPgto': 'Dt_Limite', 'Valor (R$)': 'Valor_Original', 'Vlr. Mora': 'Mora',
                    'Vlr. Desc.': 'Desconto', 'Vlr. Outros\nAcresc.': 'Outros',
                    'Dt. Liquid.': 'Dt_Liquidacao', 'Vlr. Cobrado': 'Valor_Cobrado', 'Arquivo_Origem': 'Origem'
                }
                df.rename(columns=mapa_nomes, inplace=True)

                if 'Sacado' in df.columns:
                    df['Sacado'] = df['Sacado'].astype(str).str.replace('\n',' ', regex=False).str.strip()
                    df['Sacado'] = df['Sacado'].str.replace(r'\d{11}|\d{14}','', regex=True).str.strip()

                colunas_moeda = ['Valor_Original','Mora','Desconto','Outros','Valor_Cobrado']
                for col in colunas_moeda:
                    if col in df.columns:
                        df[col] = df[col].astype(str).str.replace('.','', regex=False).str.replace(',','.', regex=False)
                        df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0).round(2)

                colunas_data = ['Prev_Credito','Vencimento','Dt_Limite','Dt_Liquidacao']
                for col in colunas_data:
                    if col in df.columns:
                        df[col] = pd.to_datetime(df[col], format='%d/%m/%Y', errors='coerce')

                df = df.replace({pd.NaT: None, float('nan'): None})
                if 'Nosso_Numero' in df.columns:
                    df = df.dropna(subset=['Nosso_Numero'])
                    df = df.drop_duplicates(subset=['Nosso_Numero'])

                # ==========================================
                # EXIBIÇÃO DOS CARDS DE MÉTRICAS (MANTIDO O VISUAL NOVO)
                # ==========================================
                st.write("")
                c1, c2, c3 = st.columns(3)
                
                total_liquidado = df['Valor_Cobrado'].sum() if 'Valor_Cobrado' in df.columns else 0
                mora_total = df['Mora'].sum() if 'Mora' in df.columns else 0
                
                # Formatação Brasileira para a tela
                fmt_valor = f"R$ {total_liquidado:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
                fmt_mora = f"R$ {mora_total:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')

                c1.metric("Títulos Processados", f"{len(df)}")
                c2.metric("Total Liquidado", fmt_valor)
                c3.metric("Total de Mora", fmt_mora)

                # ==========================================
                # SEU CÓDIGO ORIGINAL DE GERAÇÃO DO EXCEL
                # ==========================================
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df.to_excel(writer, sheet_name='Dados', index=False, startrow=4)
                    wb = writer.book
                    ws = wb['Dados']
                    
                    ws.sheet_view.showGridLines = False
                    
                    ws['A1'] = 'RELATÓRIO DE LIQUIDAÇÕES - SICOOB'
                    ws['A1'].font = Font(name='Calibri', size=16, bold=True, color='404040')
                    ws['A2'] = f'Fontes: {len(arquivos_pdf)} ficheiro(s) PDF processado(s)'
                    ws['A2'].font = Font(name='Calibri', size=11, color='595959')
                    ws['A3'] = f'Gerado em {datetime.now().strftime("%d/%m/%Y %H:%M")}'
                    ws['A3'].font = Font(size=9, italic=True, color='808080')
                    
                    header_row = 5
                    last_data_row = header_row + len(df)
                    last_col = get_column_letter(len(df.columns))
                    
                    tabela = Table(displayName="TbSicoob", ref=f"A{header_row}:{last_col}{last_data_row}")
                    tabela.tableStyleInfo = TableStyleInfo(name="TableStyleLight15", showRowStripes=True, showColumnStripes=False)
                    ws.add_table(tabela)
                    ws.freeze_panes = 'A6'
                    
                    formato_moeda = '#,##0.00' 
                    formato_data = 'DD/MM/YYYY'
                    
                    for idx, nome_col in enumerate(df.columns, 1):
                        letra = get_column_letter(idx)
                        if nome_col == 'Sacado': ws.column_dimensions[letra].width = 40
                        elif nome_col == 'Origem': ws.column_dimensions[letra].width = 25
                        elif nome_col in ['Nosso_Numero','Seu_Numero']: ws.column_dimensions[letra].width = 16
                        else: ws.column_dimensions[letra].width = 14
                        
                        for row in range(header_row+1, last_data_row+1):
                            cell = ws[f'{letra}{row}']
                            if nome_col in colunas_moeda: cell.number_format = formato_moeda
                            elif nome_col in colunas_data:
                                if cell.value is not None: cell.number_format = formato_data
                                cell.alignment = Alignment(horizontal='center', vertical='center')
                            else: cell.alignment = Alignment(vertical='center')

                    linha_total = last_data_row + 1
                    ws[f'A{linha_total}'] = "TOTAIS (Visíveis):"
                    ws[f'A{linha_total}'].font = Font(name='Calibri', size=11, bold=True, color='404040')
                    ws[f'A{linha_total}'].alignment = Alignment(horizontal='right', vertical='center')
                    
                    col_inicio = df.columns.get_loc('Valor_Original') if 'Valor_Original' in df.columns else 0
                    if col_inicio > 1:
                        ws.merge_cells(f'A{linha_total}:{get_column_letter(col_inicio)}{linha_total}')

                    for idx, nome_col in enumerate(df.columns, 1):
                        if nome_col in colunas_moeda:
                            letra = get_column_letter(idx)
                            cel = ws[f'{letra}{linha_total}']
                            cel.value = f"=SUBTOTAL(109,{letra}{header_row+1}:{letra}{last_data_row})"
                            cel.number_format = formato_moeda
                            cel.font = Font(name='Calibri', size=11, bold=True, color='1F1F1F')
                            cel.fill = PatternFill(start_color='F2F2F2', end_color='F2F2F2', fill_type='solid')
                            cel.border = Border(top=Side(style='thin', color='BFBFBF'), bottom=Side(style='medium', color='BFBFBF'))

                st.write("")
                st.success("✅ Consolidação concluída com sucesso!")
                
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
                    st.markdown("### Pré-visualização da Tabela")
                    df_view = df.copy()
                    for col in colunas_moeda:
                        if col in df_view.columns:
                            df_view[col] = df_view[col].apply(lambda x: f"R$ {x:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.') if pd.notnull(x) else "")
                    for col in colunas_data:
                        if col in df_view.columns:
                            df_view[col] = pd.to_datetime(df_view[col]).dt.strftime('%d/%m/%Y')
                    st.dataframe(df_view, use_container_width=True, hide_index=True)
