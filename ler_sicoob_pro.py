"""
SICOOB PROFISSIONAL v5.1 - Consolidado
Tons cinza, sem linhas de grade (gridlines ocultas) e totais dinâmicos.
"""
import os, sys, glob
from datetime import datetime
import warnings

# Ignorar alertas do pandas
warnings.simplefilter(action='ignore', category=FutureWarning)

print("="*60)
print(" SICOOB - Relatório Profissional (Consolidado)")
print("="*60)

try:
    import fitz
    import pandas as pd
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.worksheet.table import Table, TableStyleInfo
    from openpyxl.utils import get_column_letter
except ImportError:
    print("Erro: Bibliotecas não encontradas. Execute pelo arquivo .BAT.")
    input("Pressione Enter para sair...")
    sys.exit()

# Procurar todos os PDFs na pasta
pdfs = glob.glob("*.pdf")
if not pdfs:
    print("ERRO: Nenhum PDF encontrado na pasta!")
    input("Pressione Enter...")
    sys.exit()

print(f"{len(pdfs)} arquivo(s) PDF encontrado(s). Iniciando leitura...")

dados_totais = []

# Processar PDFs
for pdf_path in pdfs:
    print(f" -> Lendo: {pdf_path}")
    doc = fitz.open(pdf_path)
    for i, page in enumerate(doc, 1):
        for t in page.find_tables():
            df = t.to_pandas()
            if any('Sacado' in col for col in df.columns) and any('Valor (R$)' in col for col in df.columns):
                df['Arquivo_Origem'] = pdf_path 
                dados_totais.append(df)
    doc.close()

if not dados_totais:
    print("ERRO: Nenhuma tabela válida encontrada nos PDFs!")
    input("Pressione Enter...")
    sys.exit()

df_raw = pd.concat(dados_totais, ignore_index=True)

# Limpeza e Padronização
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

# Converter valores para NÚMERO (arredondamento rigoroso)
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

# Criar Excel
output = f"RELATORIO_SICOOB_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
print(f"Criando {output}...")

with pd.ExcelWriter(output, engine='openpyxl') as writer:
    df.to_excel(writer, sheet_name='Dados', index=False, startrow=4)
    wb = writer.book
    ws = wb['Dados']
    
    # -------------------------------------------------------------------
    # OCULTAR LINHAS DE GRADE (GRIDLINES) PARA VISUAL PROFISSIONAL
    # -------------------------------------------------------------------
    ws.sheet_view.showGridLines = False
    
    # Cabeçalho
    ws['A1'] = 'RELATÓRIO DE LIQUIDAÇÕES - SICOOB'
    ws['A1'].font = Font(name='Calibri', size=16, bold=True, color='404040') # Cinza escuro
    ws['A2'] = f'Fontes: {len(pdfs)} arquivo(s) PDF processado(s)'
    ws['A2'].font = Font(name='Calibri', size=11, color='595959')
    ws['A3'] = f'Gerado em {datetime.now().strftime("%d/%m/%Y %H:%M")}'
    ws['A3'].font = Font(size=9, italic=True, color='808080')
    
    header_row = 5
    last_data_row = header_row + len(df)
    last_col = get_column_letter(len(df.columns))
    
    # Tabela dinâmica - Tema Cinza Claro Profissional
    tabela = Table(displayName="TbSicoob", ref=f"A{header_row}:{last_col}{last_data_row}")
    tabela.tableStyleInfo = TableStyleInfo(name="TableStyleLight15", showRowStripes=True, showColumnStripes=False)
    ws.add_table(tabela)
    ws.freeze_panes = 'A6'
    
    # Formato nativo do Excel que se adapta ao Windows (pt-BR = 1.580,00)
    formato_moeda = '#,##0.00' 
    formato_data = 'DD/MM/YYYY'
    
    # Formatação de Colunas e Alinhamentos
    for idx, nome_col in enumerate(df.columns, 1):
        letra = get_column_letter(idx)
        
        # Ajuste inteligente de larguras
        if nome_col == 'Sacado': ws.column_dimensions[letra].width = 40
        elif nome_col == 'Origem': ws.column_dimensions[letra].width = 25
        elif nome_col in ['Nosso_Numero','Seu_Numero']: ws.column_dimensions[letra].width = 16
        else: ws.column_dimensions[letra].width = 14
        
        for row in range(header_row+1, last_data_row+1):
            cell = ws[f'{letra}{row}']
            if nome_col in colunas_moeda:
                cell.number_format = formato_moeda
            elif nome_col in colunas_data:
                if cell.value is not None:
                    cell.number_format = formato_data
                cell.alignment = Alignment(horizontal='center', vertical='center')
            else:
                cell.alignment = Alignment(vertical='center')

    # Adicionar Linha de Totais Dinâmicos (SUBTOTAL)
    linha_total = last_data_row + 1
    ws[f'A{linha_total}'] = "TOTAIS (Visíveis):"
    ws[f'A{linha_total}'].font = Font(name='Calibri', size=11, bold=True, color='404040')
    ws[f'A{linha_total}'].alignment = Alignment(horizontal='right', vertical='center')
    
    # Mesclar até a coluna antes do primeiro valor para ficar visualmente limpo
    col_inicio_valores = df.columns.get_loc('Valor_Original') if 'Valor_Original' in df.columns else 0
    if col_inicio_valores > 1:
        letra_fim_mescla = get_column_letter(col_inicio_valores)
        ws.merge_cells(f'A{linha_total}:{letra_fim_mescla}{linha_total}')

    # Aplicar as fórmulas de soma condicional apenas nas colunas de moeda
    for idx, nome_col in enumerate(df.columns, 1):
        if nome_col in colunas_moeda:
            letra = get_column_letter(idx)
            celula_total = ws[f'{letra}{linha_total}']
            
            # A função 109 soma apenas células visíveis, ignorando as ocultas por filtros
            celula_total.value = f"=SUBTOTAL(109,{letra}{header_row+1}:{letra}{last_data_row})"
            celula_total.number_format = formato_moeda
            celula_total.font = Font(name='Calibri', size=11, bold=True, color='1F1F1F')
            celula_total.fill = PatternFill(start_color='F2F2F2', end_color='F2F2F2', fill_type='solid') # Fundo cinza extra claro
            celula_total.border = Border(top=Side(style='thin', color='BFBFBF'), bottom=Side(style='medium', color='BFBFBF'))

print("\n" + "="*60)
print("CONCLUÍDO COM SUCESSO!")
print(f"Arquivo gerado: {output}")
print(f"Total de registros consolidados: {len(df)}")
print("="*60)