from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, Border, Side
from app.core.config import (
    COLUNAS, DASHBOARD_INDICADORES, FILL_BIMESTRES, COLUNAS_SEC, DASHBOARD_SEC_TURMA,
    DASHBOARD_SEC_GERAL, ALINHAMENTO_CENTRALIZADO, DASHBOARD_SEC_APROVACAO, DISCIPLINAS
)

def configurar_largura_colunas(ws, colunas_largura):
    for coluna_nome, largura_cm in colunas_largura.items():
        coluna_idx = (COLUNAS_SEC.index(coluna_nome) if coluna_nome in COLUNAS_SEC else COLUNAS.index(coluna_nome)) + 1
        coluna_letra = get_column_letter(coluna_idx)
        largura_unidades = largura_cm * 3.78
        ws.column_dimensions[coluna_letra].width = largura_unidades

def criar_dashboard_turma(ws, linha_inicio_tabela, linha_inicio_dados):
    border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    
    dashboard_linha = linha_inicio_tabela
    ws[f'N{dashboard_linha}'] = "Resumo da Turma"
    ws[f'N{dashboard_linha}'].font = Font(bold=True)
    ws[f'N{dashboard_linha}'].alignment = ALINHAMENTO_CENTRALIZADO
    ws.merge_cells(f'N{dashboard_linha}:R{dashboard_linha}')

    dashboard_linha += 1
    ws[f'O{dashboard_linha}'] = "1º Bimestre"
    ws[f'P{dashboard_linha}'] = "2º Bimestre"
    ws[f'Q{dashboard_linha}'] = "3º Bimestre"
    ws[f'R{dashboard_linha}'] = "4º Bimestre"
    for col in range(15, 19):
        cell = ws[f'{get_column_letter(col)}{dashboard_linha}']
        cell.font = Font(bold=True)
        cell.alignment = ALINHAMENTO_CENTRALIZADO
        cell.fill = FILL_BIMESTRES

    inicio = linha_inicio_dados
    fim = linha_inicio_dados + 34
    bimestre_cols = ['C', 'D', 'E', 'F']

    for idx, indicador in enumerate(DASHBOARD_INDICADORES):
        dashboard_linha += 1
        ws[f'N{dashboard_linha}'] = indicador["nome"]
        ws[f'N{dashboard_linha}'].alignment = ALINHAMENTO_CENTRALIZADO
        
        if indicador["nome"] == "MATRÍCULAS":
            ws[f'O{dashboard_linha}'] = f'=O{linha_inicio_tabela + 2}+O{linha_inicio_tabela + 3}'
            ws[f'P{dashboard_linha}'] = f'=P{linha_inicio_tabela + 2}+P{linha_inicio_tabela + 3}'
            ws[f'Q{dashboard_linha}'] = f'=Q{linha_inicio_tabela + 2}+Q{linha_inicio_tabela + 3}'
            ws[f'R{dashboard_linha}'] = f'=R{linha_inicio_tabela + 2}+R{linha_inicio_tabela + 3}'
        elif indicador["nome"] == "TAXA DE APROVAÇÃO (%)":
            ws[f'O{dashboard_linha}'] = f'=IF(O{linha_inicio_tabela + 7}=0, 0, O{linha_inicio_tabela + 2}/O{linha_inicio_tabela + 7})'
            ws[f'P{dashboard_linha}'] = f'=IF(P{linha_inicio_tabela + 7}=0, 0, P{linha_inicio_tabela + 2}/P{linha_inicio_tabela + 7})'
            ws[f'Q{dashboard_linha}'] = f'=IF(Q{linha_inicio_tabela + 7}=0, 0, Q{linha_inicio_tabela + 2}/Q{linha_inicio_tabela + 7})'
            ws[f'R{dashboard_linha}'] = f'=IF(R{linha_inicio_tabela + 7}=0, 0, R{linha_inicio_tabela + 2}/R{linha_inicio_tabela + 7})'
        else:
            for col_idx, bimestre_col in enumerate(bimestre_cols):
                ws[f'{get_column_letter(15 + col_idx)}{dashboard_linha}'] = indicador["formula"](bimestre_col, inicio, fim)
        
        for col in range(15, 19):
            ws[f'{get_column_letter(col)}{dashboard_linha}'].alignment = ALINHAMENTO_CENTRALIZADO
            if indicador["formato"]:
                ws[f'{get_column_letter(col)}{dashboard_linha}'].number_format = indicador["formato"]

    for row in range(linha_inicio_tabela, dashboard_linha + 1):
        for col in range(14, 19):
            cell = ws[f'{get_column_letter(col)}{row}']
            cell.border = border

    ws.column_dimensions['N'].width = 25
    ws.column_dimensions['O'].width = 10
    ws.column_dimensions['P'].width = 10
    ws.column_dimensions['Q'].width = 10
    ws.column_dimensions['R'].width = 10

def criar_dashboard_sec_turma(ws, linha_inicio_tabela, linha_inicio_dados, num_alunos):
    border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    
    dashboard_linha = linha_inicio_tabela
    ws[f'G{dashboard_linha}'] = "Resumo Parcial por Turma"
    ws[f'G{dashboard_linha}'].font = Font(bold=True)
    ws[f'G{dashboard_linha}'].alignment = ALINHAMENTO_CENTRALIZADO
    ws.merge_cells(f'G{dashboard_linha}:I{dashboard_linha}')

    inicio = linha_inicio_dados
    fim = linha_inicio_dados + num_alunos - 1

    for idx, indicador in enumerate(DASHBOARD_SEC_TURMA):
        dashboard_linha += 1
        ws[f'G{dashboard_linha}'] = indicador["nome"]
        ws[f'G{dashboard_linha}'].alignment = ALINHAMENTO_CENTRALIZADO
        col_ref = 'C' if indicador["nome"] == "ATIVOS" else 'D' if indicador["nome"] == "TRANSFERIDOS" else 'E' if indicador["nome"] == "DESISTENTES" else 'B'
        ws[f'H{dashboard_linha}'] = indicador["formula"](col_ref, inicio, fim)
        ws[f'H{dashboard_linha}'].alignment = ALINHAMENTO_CENTRALIZADO
        ws[f'I{dashboard_linha}'] = ws[f'H{dashboard_linha}'].value
        ws[f'I{dashboard_linha}'].alignment = ALINHAMENTO_CENTRALIZADO
        
        if indicador["formato"]:
            ws[f'H{dashboard_linha}'].number_format = indicador["formato"]
            ws[f'I{dashboard_linha}'].number_format = indicador["formato"]

    for row in range(linha_inicio_tabela, dashboard_linha + 1):
        for col in range(7, 10):
            cell = ws[f'{get_column_letter(col)}{row}']
            cell.border = border

    ws.column_dimensions['G'].width = 20
    ws.column_dimensions['H'].width = 10
    ws.column_dimensions['I'].width = 10

def criar_dashboard_sec_geral(ws, linhas_inicio_tabelas, num_alunos_por_turma):
    border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    
    dashboard_linha = linhas_inicio_tabelas[0]
    ws[f'J{dashboard_linha}'] = "Resumo Geral da Escola"
    ws[f'J{dashboard_linha}'].font = Font(bold=True)
    ws[f'J{dashboard_linha}'].alignment = ALINHAMENTO_CENTRALIZADO
    ws.merge_cells(f'J{dashboard_linha}:L{dashboard_linha}')

    matriculas_refs = [f'H{linha_inicio + 1}' for linha_inicio in linhas_inicio_tabelas]
    ativos_refs = [f'H{linha_inicio + 2}' for linha_inicio in linhas_inicio_tabelas]
    transferidos_refs = [f'H{linha_inicio + 3}' for linha_inicio in linhas_inicio_tabelas]
    desistentes_refs = [f'H{linha_inicio + 4}' for linha_inicio in linhas_inicio_tabelas]

    refs_por_indicador = {
        "MATRÍCULAS": matriculas_refs,
        "ATIVOS": ativos_refs,
        "TRANSFERIDOS": transferidos_refs,
        "DESISTENTES": desistentes_refs
    }

    for idx, indicador in enumerate(DASHBOARD_SEC_GERAL):
        dashboard_linha += 1
        ws[f'J{dashboard_linha}'] = indicador["nome"]
        ws[f'J{dashboard_linha}'].alignment = ALINHAMENTO_CENTRALIZADO
        
        if indicador["nome"] in refs_por_indicador:
            ws[f'K{dashboard_linha}'] = indicador["formula"](refs_por_indicador[indicador["nome"]])
        else:
            ws[f'K{dashboard_linha}'] = indicador["formula"](dashboard_linha)
        
        ws[f'K{dashboard_linha}'].alignment = ALINHAMENTO_CENTRALIZADO
        if indicador["formato"]:
            ws[f'K{dashboard_linha}'].number_format = indicador["formato"]

    for row in range(linhas_inicio_tabelas[0], dashboard_linha + 1):
        for col in range(10, 13):
            cell = ws[f'{get_column_letter(col)}{row}']
            cell.border = border

    ws.column_dimensions['J'].width = 25
    ws.column_dimensions['K'].width = 10
    ws.column_dimensions['L'].width = 10

def criar_dashboard_sec_aprovacao(ws, turmas, linhas_inicio_tabelas):
    border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    
    dashboard_linha = linhas_inicio_tabelas[0]
    ws[f'M{dashboard_linha}'] = "TAXA DE APROVAÇÃO BIMESTRAL"
    ws[f'M{dashboard_linha}'].font = Font(bold=True)
    ws[f'M{dashboard_linha}'].alignment = ALINHAMENTO_CENTRALIZADO
    ws.merge_cells(f'M{dashboard_linha}:Q{dashboard_linha}')

    dashboard_linha += 1
    ws[f'M{dashboard_linha}'] = "TURMA"
    ws[f'N{dashboard_linha}'] = "B1"
    ws[f'O{dashboard_linha}'] = "B2"
    ws[f'P{dashboard_linha}'] = "B3"
    ws[f'Q{dashboard_linha}'] = "B4"
    for col in range(13, 18):
        cell = ws[f'{get_column_letter(col)}{dashboard_linha}']
        cell.font = Font(bold=True)
        cell.alignment = ALINHAMENTO_CENTRALIZADO
        cell.fill = FILL_BIMESTRES

    # Referências para as notas dos alunos por disciplina e bimestre
    refs_por_turma = {}
    for idx, turma in enumerate(turmas):
        refs_por_turma[turma["nome_turma"]] = []
        for disciplina in DISCIPLINAS:
            # Cada aba de disciplina tem as turmas listadas, com 36 linhas por turma (2 de cabeçalho + 34 de dados)
            linha_inicio_dados = 3 + (idx * 51)  # 51 linhas por turma (36 dados + 15 dashboard)
            refs_por_turma[turma["nome_turma"]].append({
                "B1": f"'{disciplina}'!C{linha_inicio_dados}:C{linha_inicio_dados + 34}",
                "B2": f"'{disciplina}'!D{linha_inicio_dados}:D{linha_inicio_dados + 34}",
                "B3": f"'{disciplina}'!E{linha_inicio_dados}:E{linha_inicio_dados + 34}",
                "B4": f"'{disciplina}'!F{linha_inicio_dados}:F{linha_inicio_dados + 34}"
            })

    # Preenchendo as taxas de aprovação por turma
    taxas_por_bimestre = {"B1": [], "B2": [], "B3": [], "B4": []}
    for idx, turma in enumerate(turmas):
        dashboard_linha += 1
        ws[f'M{dashboard_linha}'] = turma["nome_turma"]
        ws[f'M{dashboard_linha}'].alignment = ALINHAMENTO_CENTRALIZADO

        for bimestre, col in zip(["B1", "B2", "B3", "B4"], ['N', 'O', 'P', 'Q']):
            refs = [ref[bimestre] for ref in refs_por_turma[turma["nome_turma"]]]
            # Construindo a fórmula sem barras invertidas dentro da f-string
            sub_formulas = [f"COUNTIF({ref}, \">=7\")/COUNTA({ref})" for ref in refs]
            formula = f"=AVERAGE({','.join(sub_formulas)})"
            ws[f'{col}{dashboard_linha}'] = formula
            ws[f'{col}{dashboard_linha}'].number_format = '0.00'
            ws[f'{col}{dashboard_linha}'].alignment = ALINHAMENTO_CENTRALIZADO
            taxas_por_bimestre[bimestre].append(f'{col}{dashboard_linha}')

    # Preenchendo as taxas gerais (aprovação e reprovação)
    for indicador in DASHBOARD_SEC_APROVACAO:
        dashboard_linha += 1
        ws[f'M{dashboard_linha}'] = indicador["nome"]
        ws[f'M{dashboard_linha}'].alignment = ALINHAMENTO_CENTRALIZADO

        for bimestre, col in zip(["B1", "B2", "B3", "B4"], ['N', 'O', 'P', 'Q']):
            inicio = taxas_por_bimestre[bimestre][0]
            fim = taxas_por_bimestre[bimestre][-1]
            ws[f'{col}{dashboard_linha}'] = indicador["formula"](col, inicio, fim)
            ws[f'{col}{dashboard_linha}'].alignment = ALINHAMENTO_CENTRALIZADO
            if indicador["formato"]:
                ws[f'{col}{dashboard_linha}'].number_format = indicador["formato"]

    # Aplicando bordas
    for row in range(linhas_inicio_tabelas[0], dashboard_linha + 1):
        for col in range(13, 18):
            cell = ws[f'{get_column_letter(col)}{row}']
            cell.border = border

    # Ajustando larguras das colunas
    ws.column_dimensions['M'].width = 15
    ws.column_dimensions['N'].width = 10
    ws.column_dimensions['O'].width = 10
    ws.column_dimensions['P'].width = 10
    ws.column_dimensions['Q'].width = 10