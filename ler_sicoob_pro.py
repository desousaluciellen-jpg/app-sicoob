import streamlit as st
import fitz
import pandas as pd
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils import get_column_letter
import io
from datetime import datetime
import warnings

warnings.simplefilter(action='ignore', category=FutureWarning)

st.set_page_config(page_title="Consolidador SICOOB", page_icon="📊", layout="wide") # Mudei para 'wide' para a tabela caber melhor

st.title("📊 Consolidador de Liquidações SICOOB")
st.markdown("Faça o upload dos extratos em PDF. O sistema irá extrair, consolidar e gerar a planilha profissional automaticamente.")

arquivos_pdf = st.file_uploader("Arraste os seus PDFs para aqui", type="pdf", accept_multiple_files=True)

if arquivos_pdf:
    if st.button("Processar Documentos"):
        with st.spinner("A analisar os PDFs e cruzar os dados..."):
            dados_totais = []
            
            for arquivo in arquivos_pdf:
                doc = fitz.open(stream=arquivo.read(), filetype="pdf")
                for page in doc:
                    for t in page.find_tables():
                        df = t.to_pandas()
                        if any('Sacado' in col for col in df.columns) and any('Valor (R$)' in col for col in df.columns):
                            df['Arquivo_Origem'] = arquivo.name 
                            dados_totais.append(df)
                doc.close()

            if not dados_totais:
                st.error("Nenhuma tabela válida encontrada nos PDFs enviados!")
            else:
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

                # ==========================================
                # NOVA SEÇÃO: PRÉ-VISUALIZAÇÃO NA TELA
                # ==========================================
                st.divider() # Linha divisória visual
                st.subheader("👀 Pré-visualização Rápida")
                
                # 1. Indicadores (Cards de Resumo)
                col1, col2, col3 = st.columns(3)
                total_titulos = len(df)
                valor_total = df['Valor_Cobrado'].sum() if 'Valor_Cobrado' in df.columns else 0
                mora_total = df['Mora'].sum() if 'Mora' in df.columns else 0
                
                # Formatando os totais para os cards (estilo BR)
                str_valor = f"R$ {valor_total:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
                str_mora = f"R$ {mora_total:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
                
                col1.metric(label="Títulos Processados", value=total_titulos)
                col2.metric(label="Total Liquidado", value=str_valor)
                col3.metric(label="Total de Mora", value=str_mora)
                
                # 2. Tabela Interativa (Criando uma cópia apenas para visualização bonita)
                df_view = df.copy()
                for col in colunas_moeda:
                    if col in df_view.columns:
                        # Formata R$ na tela
                        df_view[col] = df_view[col].apply(lambda x: f"R$ {x:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.') if pd.notnull(x) else "")
                
                for col in colunas_data:
                    if col in df_view.columns:
                        # Formata Data na tela
                        df_view[col] = pd.to_datetime(df_view[col]).dt.strftime('%d/%m/%Y')
                
                # Renderiza a tabela na tela do Streamlit
                st.dataframe(df_view, use_container_width=True, hide_index=True)
                
                st.divider()
                # ==========================================

                # Gerar Excel em Memória (O código continua o mesmo a partir daqui)
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

                st.success(f"✅ Planilha pronta para download!")
                
                # Botão de download atualizado
                st.download_button(
                    label="📥 Baixar Excel Completo (com totais dinâmicos)",
                    data=output.getvalue(),
                    file_name=f"RELATORIO_SICOOB_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary"
                )
