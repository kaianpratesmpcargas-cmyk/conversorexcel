"""
Módulo de exportação para Excel (.xlsx) com formatação profissional via OpenPyXL e Pandas.
"""
from typing import List, Dict, Any
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


def export_to_excel(records: List[Dict[str, Any]], filepath: str) -> None:
    """
    Gera um arquivo Excel profissional com:
    - Aba principal dos dados (cabeçalho estilizado, filtros, congelamento, destaque condicional para SITUAÇÃO, formatação numérica de peso e moeda)
    - Segunda aba 'RESUMO' com totalizadores e agrupamento por destino.
    """
    if not records:
        raise ValueError("Não há registros para exportar.")

    wb = Workbook()

    # ==========================
    # ABA 1: DADOS DAS ENTREGAS
    # ==========================
    ws_data = wb.active
    ws_data.title = "ACOMPANHAMENTO"

    columns = [
        "CTRC",
        "DATA EMISSÃO",
        "N FISCAL",
        "DIAS DE ATRASO",
        "SITUAÇÃO"
    ]

    # Paleta de Cores Profissional
    HEADER_FILL = PatternFill(start_color="1A365D", end_color="1A365D", fill_type="solid") # Azul Marinho Escuro
    HEADER_FONT = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    
    # Preenchimentos para a coluna SITUAÇÃO
    FILL_ATRASADO = PatternFill(start_color="FFD1D1", end_color="FFD1D1", fill_type="solid")   # Vermelho Suave
    FONT_ATRASADO = Font(name="Calibri", size=10, bold=True, color="9C0006")

    FILL_NO_PRAZO = PatternFill(start_color="D1F2D9", end_color="D1F2D9", fill_type="solid")   # Verde Suave
    FONT_NO_PRAZO = Font(name="Calibri", size=10, bold=True, color="006100")

    FILL_SEM_INFO = PatternFill(start_color="EEEEEE", end_color="EEEEEE", fill_type="solid")   # Cinza Claro
    FONT_SEM_INFO = Font(name="Calibri", size=10, color="555555")

    THIN_BORDER = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9')
    )

    # Escrever Cabeçalho
    for col_idx, col_name in enumerate(columns, start=1):
        cell = ws_data.cell(row=1, column=col_idx, value=col_name)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = THIN_BORDER
    
    ws_data.row_dimensions[1].height = 28

    # Formatos de número
    FORMAT_MOEDA = "R$ #,##0.00"
    FORMAT_PESO = "#,##0.00"
    FORMAT_INTEIRO = "#,##0"

    # Preencher Linhas
    for row_idx, rec in enumerate(records, start=2):
        for col_idx, col_name in enumerate(columns, start=1):
            val = rec.get(col_name)
            cell = ws_data.cell(row=row_idx, column=col_idx, value=val)
            cell.border = THIN_BORDER
            cell.font = Font(name="Calibri", size=10)

            # Alinhamentos e Formatações Específicas
            if col_name in ["CTRC", "DATA EMISSÃO", "N FISCAL", "PREVISÃO ENTREGA", "DIAS DE ATRASO", "SITUAÇÃO"]:
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif col_name in ["PESO", "FRETE"]:
                cell.alignment = Alignment(horizontal="right", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center")

            # Formatação numérica
            if col_name == "PESO":
                cell.number_format = FORMAT_PESO
            elif col_name == "FRETE":
                cell.number_format = FORMAT_MOEDA
            elif col_name == "DIAS DE ATRASO" and val is not None:
                cell.number_format = FORMAT_INTEIRO

            # Destaque de SITUAÇÃO
            if col_name == "SITUAÇÃO":
                sit = str(val or "").strip().upper()
                if sit == "ATRASADO":
                    cell.fill = FILL_ATRASADO
                    cell.font = FONT_ATRASADO
                elif sit == "NO PRAZO":
                    cell.fill = FILL_NO_PRAZO
                    cell.font = FONT_NO_PRAZO
                else:
                    cell.fill = FILL_SEM_INFO
                    cell.font = FONT_SEM_INFO

        ws_data.row_dimensions[row_idx].height = 20

    # Auto-filtro na tabela de dados
    last_col_letter = get_column_letter(len(columns))
    last_row = len(records) + 1
    ws_data.auto_filter.ref = f"A1:{last_col_letter}{last_row}"

    # Congelar primeira linha (cabeçalho fixo ao rolar)
    ws_data.freeze_panes = "A2"

    # Ajuste automático de largura das colunas
    for col in ws_data.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            val_str = str(cell.value or '')
            if cell.number_format == FORMAT_MOEDA and isinstance(cell.value, (int, float)):
                val_str = f"R$ {cell.value:,.2f}"
            max_len = max(max_len, len(val_str))
        ws_data.column_dimensions[col_letter].width = max(max_len + 4, 12)

    # ==========================
    # ABA 2: RESUMO E INDICADORES
    # ==========================
    ws_resumo = wb.create_sheet(title="RESUMO")
    ws_resumo.views.sheetView[0].showGridLines = True

    # Cálculos para os indicadores
    df = pd.DataFrame(records)
    total_ctrcs = len(df)
    total_atrasados = int((df['SITUAÇÃO'] == 'ATRASADO').sum()) if 'SITUAÇÃO' in df else 0
    total_no_prazo = int((df['SITUAÇÃO'] == 'NO PRAZO').sum()) if 'SITUAÇÃO' in df else 0
    total_sem_info = int((df['SITUAÇÃO'] == 'SEM INFORMAÇÃO').sum()) if 'SITUAÇÃO' in df else 0
    peso_total = float(df['PESO'].sum()) if 'PESO' in df else 0.0
    frete_total = float(df['FRETE'].sum()) if 'FRETE' in df else 0.0

    # Título do Relatório
    ws_resumo.merge_cells("B2:G2")
    title_cell = ws_resumo["B2"]
    title_cell.value = "MP CARGAS — RESUMO GERENCIAL DE ENTREGAS"
    title_cell.font = Font(name="Calibri", size=14, bold=True, color="1A365D")
    title_cell.alignment = Alignment(horizontal="left", vertical="center")
    ws_resumo.row_dimensions[2].height = 25

    # Tabela de Indicadores Gerais
    resumo_metrics = [
        ("Total de CT-RCs", total_ctrcs, FORMAT_INTEIRO),
        ("Total de Atrasados", total_atrasados, FORMAT_INTEIRO),
        ("Total no Prazo", total_no_prazo, FORMAT_INTEIRO),
        ("Total sem Informação", total_sem_info, FORMAT_INTEIRO),
        ("Peso Total (kg)", peso_total, FORMAT_PESO),
        ("Frete Total (R$)", frete_total, FORMAT_MOEDA),
    ]

    metric_header_fill = PatternFill(start_color="2B6CB0", end_color="2B6CB0", fill_type="solid")
    metric_header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")

    ws_resumo.cell(row=4, column=2, value="Indicador").font = metric_header_font
    ws_resumo.cell(row=4, column=2).fill = metric_header_fill
    ws_resumo.cell(row=4, column=2).alignment = Alignment(horizontal="left", vertical="center")
    ws_resumo.cell(row=4, column=2).border = THIN_BORDER

    ws_resumo.cell(row=4, column=3, value="Valor").font = metric_header_font
    ws_resumo.cell(row=4, column=3).fill = metric_header_fill
    ws_resumo.cell(row=4, column=3).alignment = Alignment(horizontal="right", vertical="center")
    ws_resumo.cell(row=4, column=3).border = THIN_BORDER
    ws_resumo.row_dimensions[4].height = 22

    for idx, (label, val, fmt) in enumerate(resumo_metrics, start=5):
        c_lbl = ws_resumo.cell(row=idx, column=2, value=label)
        c_lbl.font = Font(name="Calibri", size=10, bold=True if "Atrasados" in label else False)
        c_lbl.border = THIN_BORDER
        c_lbl.alignment = Alignment(horizontal="left", vertical="center")

        c_val = ws_resumo.cell(row=idx, column=3, value=val)
        c_val.font = Font(name="Calibri", size=10, bold=True)
        c_val.number_format = fmt
        c_val.border = THIN_BORDER
        c_val.alignment = Alignment(horizontal="right", vertical="center")

        if "Atrasados" in label and total_atrasados > 0:
            c_val.fill = FILL_ATRASADO
            c_val.font = FONT_ATRASADO
        elif "Prazo" in label and total_no_prazo > 0:
            c_val.fill = FILL_NO_PRAZO
            c_val.font = FONT_NO_PRAZO

        ws_resumo.row_dimensions[idx].height = 20

    # Tabela Agrupada por Destino
    dest_start_row = 12
    ws_resumo.merge_cells(f"B{dest_start_row}:E{dest_start_row}")
    dest_title = ws_resumo[f"B{dest_start_row}"]
    dest_title.value = "ACOMPANHAMENTO AGRUPADO POR DESTINO"
    dest_title.font = Font(name="Calibri", size=12, bold=True, color="1A365D")
    dest_title.alignment = Alignment(horizontal="left", vertical="center")
    ws_resumo.row_dimensions[dest_start_row].height = 24

    dest_headers = ["Destino", "Quantidade", "Atrasados", "Peso"]
    header_row = dest_start_row + 1
    for col_idx, h_text in enumerate(dest_headers, start=2):
        c = ws_resumo.cell(row=header_row, column=col_idx, value=h_text)
        c.fill = HEADER_FILL
        c.font = HEADER_FONT
        c.border = THIN_BORDER
        c.alignment = Alignment(horizontal="center" if col_idx > 2 else "left", vertical="center")
    ws_resumo.row_dimensions[header_row].height = 22

    # Agrupar via pandas
    if not df.empty and 'DESTINO' in df:
        grouped = df.groupby('DESTINO').apply(
            lambda g: pd.Series({
                'Quantidade': len(g),
                'Atrasados': int((g['SITUAÇÃO'] == 'ATRASADO').sum()),
                'Peso': float(g['PESO'].sum())
            })
        ).reset_index().sort_values(by='Quantidade', ascending=False)
    else:
        grouped = pd.DataFrame(columns=['DESTINO', 'Quantidade', 'Atrasados', 'Peso'])

    curr_row = header_row + 1
    for _, row in grouped.iterrows():
        c_dest = ws_resumo.cell(row=curr_row, column=2, value=str(row['DESTINO'] or 'NÃO INFORMADO'))
        c_dest.font = Font(name="Calibri", size=10)
        c_dest.border = THIN_BORDER
        c_dest.alignment = Alignment(horizontal="left", vertical="center")

        c_qtd = ws_resumo.cell(row=curr_row, column=3, value=int(row['Quantidade']))
        c_qtd.font = Font(name="Calibri", size=10)
        c_qtd.number_format = FORMAT_INTEIRO
        c_qtd.border = THIN_BORDER
        c_qtd.alignment = Alignment(horizontal="center", vertical="center")

        c_atr = ws_resumo.cell(row=curr_row, column=4, value=int(row['Atrasados']))
        c_atr.font = Font(name="Calibri", size=10)
        c_atr.number_format = FORMAT_INTEIRO
        c_atr.border = THIN_BORDER
        c_atr.alignment = Alignment(horizontal="center", vertical="center")
        if int(row['Atrasados']) > 0:
            c_atr.fill = FILL_ATRASADO
            c_atr.font = FONT_ATRASADO

        c_peso = ws_resumo.cell(row=curr_row, column=5, value=float(row['Peso']))
        c_peso.font = Font(name="Calibri", size=10)
        c_peso.number_format = FORMAT_PESO
        c_peso.border = THIN_BORDER
        c_peso.alignment = Alignment(horizontal="right", vertical="center")

        ws_resumo.row_dimensions[curr_row].height = 20
        curr_row += 1

    # Larguras das colunas do resumo
    ws_resumo.column_dimensions['A'].width = 4
    ws_resumo.column_dimensions['B'].width = 30
    ws_resumo.column_dimensions['C'].width = 18
    ws_resumo.column_dimensions['D'].width = 18
    ws_resumo.column_dimensions['E'].width = 20
    ws_resumo.column_dimensions['F'].width = 15

    # Salva arquivo
    wb.save(filepath)
