from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, Border, Side
from openpyxl.chart import BarChart, Reference, Series
from openpyxl.chart.axis import ChartLines
from app.core.config import (
    COLUNAS, DASHBOARD_INDICADORES, FILL_BIMESTRES, COLUNAS_SEC, DASHBOARD_SEC_TURMA,
    DASHBOARD_SEC_GERAL, ALINHAMENTO_CENTRALIZADO, DASHBOARD_SEC_APROVACAO, DISCIPLINAS,
    LARGURAS_COLUNAS, LARGURAS_COLUNAS_ABAS_DISC
)

def configurar_largura_colunas(ws, colunas_largura, colunas_ref):
    for coluna_nome, largura_cm in colunas_largura.items():
        coluna_idx = colunas_ref.index(coluna_nome) + 1
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
    for col in range(15, 19):  # Colunas O, P, Q, R
        cell = ws[f'{get_column_letter(col)}{dashboard_linha}']
        cell.font = Font(bold=True)
        cell.alignment = ALINHAMENTO_CENTRALIZADO
        cell.fill = FILL_BIMESTRES

    inicio = linha_inicio_dados
    fim = linha_inicio_dados + 34  # 35 linhas de dados por turma
    bimestre_cols = ['C', 'D', 'E', 'F']  # Colunas dos bimestres

    # Lista de indicadores que devem ser números inteiros
    indicadores_inteiros = [
        "ALUNOS APROVADOS",
        "ALUNOS REPROVADOS",
        "Nº ALUNOS COM MÉDIA > 8,0",
        "Nº ALUNOS QUE NÃO ATINGIRAM MÉDIA > 8,0",
        "MATRÍCULAS"
    ]

    for idx, indicador in enumerate(DASHBOARD_INDICADORES):
        dashboard_linha += 1
        ws[f'N{dashboard_linha}'] = indicador["nome"]
        ws[f'N{dashboard_linha}'].alignment = ALINHAMENTO_CENTRALIZADO
        
        # Aplica a fórmula para cada bimestre, se existir
        if indicador["formula"]:
            for col_idx, bimestre_col in enumerate(bimestre_cols):
                ws[f'{get_column_letter(15 + col_idx)}{dashboard_linha}'] = indicador["formula"](bimestre_col, inicio, fim)
        
        # Aplica formato: 0 para números inteiros, 0.00 para outros números, ou formato definido (ex.: 0.00% para porcentagens)
        for col in range(15, 19):  # Colunas O, P, Q, R
            cell = ws[f'{get_column_letter(col)}{dashboard_linha}']
            cell.alignment = ALINHAMENTO_CENTRALIZADO
            if indicador["formato"]:
                cell.number_format = indicador["formato"]  # Mantém 0.00% para porcentagens
            elif indicador["nome"] in indicadores_inteiros:
                cell.number_format = '0'  # Formato inteiro para os indicadores especificados
            else:
                cell.number_format = '0.00'  # Formato 0.00 para outros números

    # Aplica bordas
    for row in range(linha_inicio_tabela, dashboard_linha + 1):
        for col in range(14, 19):  # Colunas N, O, P, Q, R
            cell = ws[f'{get_column_letter(col)}{row}']
            cell.border = border

    # Define larguras das colunas
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

    # Lista de indicadores que devem ser números inteiros
    indicadores_inteiros = ["MATRÍCULAS", "ATIVOS", "TRANSFERIDOS", "DESISTENTES"]

    for idx, indicador in enumerate(DASHBOARD_SEC_TURMA):
        dashboard_linha += 1
        ws[f'G{dashboard_linha}'] = indicador["nome"]
        ws[f'G{dashboard_linha}'].alignment = ALINHAMENTO_CENTRALIZADO
        col_ref = 'C' if indicador["nome"] == "ATIVOS" else 'D' if indicador["nome"] == "TRANSFERIDOS" else 'E' if indicador["nome"] == "DESISTENTES" else 'B'
        ws[f'H{dashboard_linha}'] = indicador["formula"](col_ref, inicio, fim)
        ws[f'H{dashboard_linha}'].alignment = ALINHAMENTO_CENTRALIZADO
        ws[f'I{dashboard_linha}'] = ws[f'H{dashboard_linha}'].value
        ws[f'I{dashboard_linha}'].alignment = ALINHAMENTO_CENTRALIZADO
        
        # Aplica formato: 0 para todos os indicadores (MATRÍCULAS, ATIVOS, TRANSFERIDOS, DESISTENTES)
        if indicador["nome"] in indicadores_inteiros:
            ws[f'H{dashboard_linha}'].number_format = '0'
            ws[f'I{dashboard_linha}'].number_format = '0'

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

    # Lista de indicadores que devem ser números inteiros
    indicadores_inteiros = ["MATRÍCULAS", "ATIVOS", "TRANSFERIDOS", "DESISTENTES", "Nº ABANDONO(S)"]

    for idx, indicador in enumerate(DASHBOARD_SEC_GERAL):
        dashboard_linha += 1
        ws[f'J{dashboard_linha}'] = indicador["nome"]
        ws[f'J{dashboard_linha}'].alignment = ALINHAMENTO_CENTRALIZADO
        
        if indicador["nome"] in refs_por_indicador:
            ws[f'K{dashboard_linha}'] = indicador["formula"](refs_por_indicador[indicador["nome"]])
        else:
            ws[f'K{dashboard_linha}'] = indicador["formula"](dashboard_linha)
        
        ws[f'K{dashboard_linha}'].alignment = ALINHAMENTO_CENTRALIZADO
        # Aplica formato: 0 para números inteiros, 0.00 para outros números, ou formato definido (ex.: 0.00% para porcentagens)
        if indicador["formato"]:
            ws[f'K{dashboard_linha}'].number_format = indicador["formato"]  # Mantém 0.00% para "ABANDONO(S) (%)"
        elif indicador["nome"] in indicadores_inteiros:
            ws[f'K{dashboard_linha}'].number_format = '0'  # Formato inteiro para os indicadores especificados
        else:
            ws[f'K{dashboard_linha}'].number_format = '0.00'  # Formato 0.00 para outros números

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
    for col in range(13, 18):  # Colunas M a Q
        cell = ws[f'{get_column_letter(col)}{dashboard_linha}']
        cell.font = Font(bold=True)
        cell.alignment = ALINHAMENTO_CENTRALIZADO
        cell.fill = FILL_BIMESTRES

    # Para cada turma, calcular a média das taxas de aprovação de todas as disciplinas
    for idx, turma in enumerate(turmas):
        dashboard_linha += 1
        ws[f'M{dashboard_linha}'] = turma["nome_turma"]
        ws[f'M{dashboard_linha}'].alignment = ALINHAMENTO_CENTRALIZADO

        # Linha base do dashboard na aba de disciplina
        linha_ref_base = 12  # Ajustado para alinhar com a linha 12 para 1º ANO A
        linha_ref = linha_ref_base + (idx * 52)  # Ajuste para cada turma (52 linhas por turma)

        # Log para depuração
        print(f"Turma: {turma['nome_turma']}, idx: {idx}, linha_ref: {linha_ref}")

        # Fórmulas para cada bimestre (média das taxas de todas as disciplinas com tratamento de erro)
        ws[f'N{dashboard_linha}'] = f'=AVERAGE({",".join([f"IFERROR({disc}!O{linha_ref},0)" for disc in DISCIPLINAS])})'
        ws[f'O{dashboard_linha}'] = f'=AVERAGE({",".join([f"IFERROR({disc}!P{linha_ref},0)" for disc in DISCIPLINAS])})'
        ws[f'P{dashboard_linha}'] = f'=AVERAGE({",".join([f"IFERROR({disc}!Q{linha_ref},0)" for disc in DISCIPLINAS])})'
        ws[f'Q{dashboard_linha}'] = f'=AVERAGE({",".join([f"IFERROR({disc}!R{linha_ref},0)" for disc in DISCIPLINAS])})'

        for col in range(13, 18):  # Colunas M a Q
            cell = ws[f'{get_column_letter(col)}{dashboard_linha}']
            cell.number_format = '0.00%'  # Já está correto para porcentagens
            cell.alignment = ALINHAMENTO_CENTRALIZADO

    linha_inicio_turmas = linhas_inicio_tabelas[0] + 2
    linha_fim_turmas = linha_inicio_turmas + len(turmas) - 1
    
    for indicador in DASHBOARD_SEC_APROVACAO:
        dashboard_linha += 1
        ws[f'M{dashboard_linha}'] = indicador["nome"]
        ws[f'M{dashboard_linha}'].font = Font(size=10)
        ws[f'M{dashboard_linha}'].alignment = ALINHAMENTO_CENTRALIZADO

        for col in ['N', 'O', 'P', 'Q']:
            if indicador["nome"] == "TX APROVAÇÃO %":
                ws[f'{col}{dashboard_linha}'] = f'=AVERAGE({col}{linha_inicio_turmas}:{col}{linha_fim_turmas})'
            else:  # TX REPROVAÇÃO %
                ws[f'{col}{dashboard_linha}'] = f'=IFERROR(1-{col}{dashboard_linha-1},0)'
            
            ws[f'{col}{dashboard_linha}'].font = Font(size=10)
            ws[f'{col}{dashboard_linha}'].number_format = '0.00%'  # Já está correto para porcentagens
            ws[f'{col}{dashboard_linha}'].alignment = ALINHAMENTO_CENTRALIZADO

    for row in range(linhas_inicio_tabelas[0], dashboard_linha + 1):
        for col in range(13, 18):  # Colunas M a Q
            cell = ws[f'{get_column_letter(col)}{row}']
            cell.border = border

    ws.column_dimensions['M'].width = 15
    ws.column_dimensions['N'].width = 10
    ws.column_dimensions['O'].width = 10
    ws.column_dimensions['P'].width = 10
    ws.column_dimensions['Q'].width = 10

    # Criar o gráfico
    chart = BarChart()
    chart.type = "col"
    chart.style = 10
    chart.title = "TAXA DE APROVAÇÃO"
    chart.y_axis.title = "Taxa de Aprovação (%)"
    chart.x_axis.title = "Turma"
    chart.height = 15
    chart.width = 20

    # Dados para todos os bimestres (colunas N, O, P, Q)
    data = Reference(ws, min_col=14, min_row=linhas_inicio_tabelas[0] + 1, max_col=17, max_row=linhas_inicio_tabelas[0] + len(turmas) + 1)
    cats = Reference(ws, min_col=13, min_row=linhas_inicio_tabelas[0] + 2, max_row=linhas_inicio_tabelas[0] + len(turmas) + 1)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)

    # Definir cores diferentes para cada bimestre (B1, B2, B3, B4)
    series_colors = ["DAA520", "CD853F", "F4A460", "DEB887"]  # Dourado, Pêssego escuro, Areia, Bege dourado
    for idx, series in enumerate(chart.series):
        series.graphicalProperties.solidFill = series_colors[idx]

    # Reduzir a largura das barras
    chart.gapWidth = 50  # 50% do espaço entre barras (padrão é 150)

    chart.y_axis.scaling.min = 0
    chart.y_axis.scaling.max = 1
    chart.y_axis.number_format = '0%'

    chart.y_axis.majorGridlines = ChartLines()

    ws.add_chart(chart, f"M{dashboard_linha + 2}")