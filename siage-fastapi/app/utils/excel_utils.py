from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, Alignment, Border, Side
from app.core.config import COLUNAS, DASHBOARD_INDICADORES, FILL_BIMESTRES

def configurar_largura_colunas(ws, colunas_largura):
    """
    Define a largura das colunas especificadas.
    :param ws: A worksheet (aba) onde as colunas serão configuradas.
    :param colunas_largura: Um dicionário onde a chave é o nome da coluna e o valor é a largura em cm.
    """
    for coluna_nome, largura_cm in colunas_largura.items():
        coluna_idx = COLUNAS.index(coluna_nome) + 1  # +1 porque as colunas começam em 1 no Excel
        coluna_letra = get_column_letter(coluna_idx)
        largura_unidades = largura_cm * 3.78
        ws.column_dimensions[coluna_letra].width = largura_unidades

def criar_dashboard_turma(ws, linha_inicio_tabela, linha_inicio_dados):
    """
    Cria um dashboard ao lado da tabela de turma com indicadores por bimestre.
    :param ws: Worksheet onde o dashboard será criado.
    :param linha_inicio_tabela: Linha onde começa o cabeçalho da tabela.
    :param linha_inicio_dados: Linha onde começam os dados dos alunos.
    """
    border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    
    # Título do dashboard
    dashboard_linha = linha_inicio_tabela
    ws[f'N{dashboard_linha}'] = "Resumo da Turma"
    ws[f'N{dashboard_linha}'].font = Font(bold=True)
    ws[f'N{dashboard_linha}'].alignment = Alignment(horizontal='center')
    ws.merge_cells(f'N{dashboard_linha}:R{dashboard_linha}')
    
    # Cabeçalhos dos bimestres
    dashboard_linha += 1
    ws[f'O{dashboard_linha}'] = "1º Bimestre"
    ws[f'P{dashboard_linha}'] = "2º Bimestre"
    ws[f'Q{dashboard_linha}'] = "3º Bimestre"
    ws[f'R{dashboard_linha}'] = "4º Bimestre"
    for col in range(15, 19):  # O a R
        cell = ws[f'{get_column_letter(col)}{dashboard_linha}']
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal='center')
        cell.fill = FILL_BIMESTRES
    
    # Intervalo de dados
    inicio = linha_inicio_dados
    fim = linha_inicio_dados + 34
    bimestre_cols = ['C', 'D', 'E', 'F']
    
    # Aplicação dos indicadores
    for idx, indicador in enumerate(DASHBOARD_INDICADORES):
        dashboard_linha += 1
        ws[f'N{dashboard_linha}'] = indicador["nome"]
        
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
        
        if indicador["formato"]:
            for col in range(15, 19):
                ws[f'{get_column_letter(col)}{dashboard_linha}'].number_format = indicador["formato"]
    
    # Bordas no dashboard
    for row in range(linha_inicio_tabela, dashboard_linha + 1):
        for col in range(14, 19):  # N a R
            cell = ws[f'{get_column_letter(col)}{row}']
            cell.border = border
    
    # Ajusta largura das colunas do dashboard
    ws.column_dimensions['N'].width = 25  # Indicador
    ws.column_dimensions['O'].width = 10  # 1º Bimestre
    ws.column_dimensions['P'].width = 10  # 2º Bimestre
    ws.column_dimensions['Q'].width = 10  # 3º Bimestre
    ws.column_dimensions['R'].width = 10  # 4º Bimestre