from pathlib import Path
from openpyxl.styles import PatternFill, Font, Alignment

DISCIPLINAS = ["BIO", "MAT", "FIS", "QUI", "GEO", "SOC", "HIS", "FIL", "ESP", "POR", "ART", "EDF", "ING"]

COLUNAS = [
    "Nº", "Nome do Aluno", "1º BIM", "2º BIM", "3º BIM", "4º BIM",
    "NF", "MG", "MF", "SITUAÇÃO DO ALUNO", "PF", "SF"
]

COLUNAS_SEC = [
    "Nº", "Nome do Aluno", "ATIVO", "TRANSFERIDO", "DESISTENTE", "SITUAÇÃO DO ALUNO"
]

CAMINHO_IMAGEM = Path("/home/lmbernardo/teste-planilha-siage-interno/siage-fastapi/app/core/static/images/siage_interno.png")

CAMINHO_PADRAO = "/home/lmbernardo/teste-planilha-siage-interno/siage-fastapi"
NOME_ARQUIVO_PADRAO = "planilha_notas_complexa.xlsx"

LARGURAS_COLUNAS = {
    "Nº": 1,
    "Nome do Aluno": 10,
    "SITUAÇÃO DO ALUNO": 10,
    "ATIVO": 4.5,
    "TRANSFERIDO": 4.5,
    "DESISTENTE": 4.5
}

LARGURAS_COLUNAS_ABAS_DISC = {
    "Nº": 1,
    "Nome do Aluno": 15,
    "1º BIM": 3,
    "2º BIM": 3,
    "3º BIM": 3,
    "4º BIM": 3,
    "NF": 3,
    "MG": 3,
    "MF": 3,
    "SITUAÇÃO DO ALUNO": 10,
    "PF": 3,
    "SF": 3
}

COR_ABA = "FFDAB9"
FILL_NOME_ALUNO = PatternFill(start_color="FF6347", end_color="FF6347", fill_type="solid")
FILL_BIMESTRES = PatternFill(start_color="FFA500", end_color="FFA500", fill_type="solid")
FILL_NOTA_FINAL = PatternFill(start_color="FF4500", end_color="FF4500", fill_type="solid")
FILL_SITUACAO = PatternFill(start_color="FFFACD", end_color="FFFACD", fill_type="solid")
FONTE_TITULO_TURMA = Font(name='Arial', size=14, bold=True, color="8B4513")
ALINHAMENTO_CENTRALIZADO = Alignment(horizontal='center', vertical='center')

DASHBOARD_INDICADORES = [
    {"nome": "ALUNOS APROVADOS", "formula": lambda col, inicio, fim: f'=COUNTIF({col}{inicio}:{col}{fim}, ">=7")', "formato": None},
    {"nome": "ALUNOS REPROVADOS", "formula": lambda col, inicio, fim: f'=COUNTIF({col}{inicio}:{col}{fim}, "<7")', "formato": None},
    {"nome": "Nº ALUNOS COM MÉDIA > 8,0", "formula": lambda col, inicio, fim: f'=COUNTIF({col}{inicio}:{col}{fim}, ">=8")', "formato": None},
    {"nome": "Nº ALUNOS QUE NÃO ATINGIRAM MÉDIA > 8,0", "formula": lambda col, inicio, fim: f'=COUNTIF({col}{inicio}:{col}{fim}, "<8")', "formato": None},
    {"nome": "PERCENTUAL DE MÉDIAS > 5,0", "formula": lambda col, inicio, fim: f'=COUNTIF({col}{inicio}:{col}{fim}, ">=5")/COUNTA({col}{inicio}:{col}{fim})', "formato": '0.00%'},
    {"nome": "PERCENTUAL DE MÉDIAS < 5,0", "formula": lambda col, inicio, fim: f'=COUNTIF({col}{inicio}:{col}{fim}, "<5")/COUNTA({col}{inicio}:{col}{fim})', "formato": '0.00%'},
    {"nome": "MATRÍCULAS", "formula": lambda col, inicio, fim: f'=COUNTA({col}{inicio}:{col}{fim})', "formato": None},
    {"nome": "TAXA DE APROVAÇÃO (%)", "formula": lambda col, inicio, fim: f'=IF(COUNTA({col}{inicio}:{col}{fim})=0, 0, COUNTIF({col}{inicio}:{col}{fim}, ">=7")/COUNTA({col}{inicio}:{col}{fim}))', "formato": '0.00%'}
]

DASHBOARD_SEC_TURMA = [
    {"nome": "MATRÍCULAS", "formula": lambda col, inicio, fim: f'=COUNTA({col}{inicio}:{col}{fim})', "formato": None},
    {"nome": "ATIVOS", "formula": lambda col, inicio, fim: f'=COUNTIF({col}{inicio}:{col}{fim}, TRUE)', "formato": None},
    {"nome": "TRANSFERIDOS", "formula": lambda col, inicio, fim: f'=COUNTIF({col}{inicio}:{col}{fim}, TRUE)', "formato": None},
    {"nome": "DESISTENTES", "formula": lambda col, inicio, fim: f'=COUNTIF({col}{inicio}:{col}{fim}, TRUE)', "formato": None}
]

DASHBOARD_SEC_GERAL = [
    {"nome": "MATRÍCULAS", "formula": lambda refs: f'=SUM({",".join(refs)})', "formato": None},
    {"nome": "ATIVOS", "formula": lambda refs: f'=SUM({",".join(refs)})', "formato": None},
    {"nome": "TRANSFERIDOS", "formula": lambda refs: f'=SUM({",".join(refs)})', "formato": None},
    {"nome": "DESISTENTES", "formula": lambda refs: f'=SUM({",".join(refs)})', "formato": None},
    {"nome": "Nº ABANDONO(S)", "formula": lambda linha_atual: f'=K{linha_atual-1}', "formato": None},
    {"nome": "ABANDONO(S) (%)", "formula": lambda linha_atual: f'=K{linha_atual-1}/K{linha_atual-4}', "formato": '0.00%'}
]

DASHBOARD_SEC_APROVACAO = [
    {"nome": "TX APROVAÇÃO %", "formula": lambda col, inicio, fim: f'=AVERAGE({col}{inicio}:{col}{fim})', "formato": '0.00%'},
    {"nome": "TX REPROVAÇÃO %", "formula": lambda col, inicio, fim: f'=1-{col}{inicio-1}', "formato": '0.00%'}
]