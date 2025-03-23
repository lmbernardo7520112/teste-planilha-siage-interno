from pathlib import Path
from openpyxl.styles import PatternFill, Font

# Lista de disciplinas
DISCIPLINAS = ["BIO", "MAT", "FIS", "QUI", "GEO", "SOC", "HIST", "FIL", "ESP", "POR", "ART", "ADF", "ING"]

# Colunas da planilha principal
COLUNAS = [
    "Nº", "Nome do Aluno", "1º BIM", "2º BIM", "3º BIM", "4º BIM",
    "NF", "MG", "MF", "SITUAÇÃO DO ALUNO", "PF", "SF"
]

# Colunas da aba SEC
COLUNAS_SEC = [
    "Nº", "Nome do Aluno", "ATIVO", "TRANSFERIDO", "DESISTENTE", "SITUAÇÃO DO ALUNO"
]

# Caminho da imagem
CAMINHO_IMAGEM = Path(__file__).parent / "static" / "images" / "siage_interno.png"

# Configurações de salvamento
CAMINHO_PADRAO = "/mnt/c/Users/lmbernardo/Downloads"
NOME_ARQUIVO_PADRAO = "planilha_notas_complexa.xlsx"

# Larguras das colunas (em cm)
LARGURAS_COLUNAS = {
    "Nº": 1,
    "Nome do Aluno": 10,
    "SITUAÇÃO DO ALUNO": 4.5,
    "ATIVO": 2.5,
    "TRANSFERIDO": 2.5,
    "DESISTENTE": 2.5
}

# Definições de cores
COR_ABA = "FFDAB9"
FILL_NOME_ALUNO = PatternFill(start_color="FF6347", end_color="FF6347", fill_type="solid")
FILL_BIMESTRES = PatternFill(start_color="FFA500", end_color="FFA500", fill_type="solid")
FILL_NOTA_FINAL = PatternFill(start_color="FF4500", end_color="FF4500", fill_type="solid")
FILL_SITUACAO = PatternFill(start_color="FFFACD", end_color="FFFACD", fill_type="solid")
FONTE_TITULO_TURMA = Font(name='Arial', size=14, bold=True, color="8B4513")

# Definição dos indicadores do dashboard principal
DASHBOARD_INDICADORES = [
    {"nome": "ALUNOS APROVADOS", "formula": lambda col, inicio, fim: f'=COUNTIF({col}{inicio}:{col}{fim}, ">=7")', "formato": None},
    {"nome": "ALUNOS REPROVADOS", "formula": lambda col, inicio, fim: f'=COUNTIF({col}{inicio}:{col}{fim}, "<7")', "formato": None},
    {"nome": "Nº ALUNOS COM MÉDIA > 8,0", "formula": lambda col, inicio, fim: f'=COUNTIF({col}{inicio}:{col}{fim}, ">=8")', "formato": None},
    {"nome": "Nº ALUNOS QUE NÃO ATINGIRAM MÉDIA > 8,0", "formula": lambda col, inicio, fim: f'=COUNTIF({col}{inicio}:{col}{fim}, "<8")', "formato": None},
    {"nome": "PERCENTUAL DE MÉDIAS > 5,0", "formula": lambda col, inicio, fim: f'=COUNTIF({col}{inicio}:{col}{fim}, ">=5")/COUNTA({col}{inicio}:{col}{fim})', "formato": '0.00%'},
    {"nome": "PERCENTUAL DE MÉDIAS < 5,0", "formula": lambda col, inicio, fim: f'=COUNTIF({col}{inicio}:{col}{fim}, "<5")/COUNTA({col}{inicio}:{col}{fim})', "formato": '0.00%'},
    {"nome": "MATRÍCULAS", "formula": None, "formato": None},
    {"nome": "TAXA DE APROVAÇÃO (%)", "formula": None, "formato": '0.00%'}
]

# Definição dos indicadores do dashboard da aba SEC (Resumo Parcial por Turma)
DASHBOARD_SEC_TURMA = [
    {"nome": "MATRÍCULAS", "formula": lambda col, inicio, fim: f'=COUNTA({col}{inicio}:{col}{fim})', "formato": None},
    {"nome": "ATIVOS", "formula": lambda col, inicio, fim: f'=COUNTIF({col}{inicio}:{col}{fim}, TRUE)', "formato": None},
    {"nome": "TRANSFERIDOS", "formula": lambda col, inicio, fim: f'=COUNTIF({col}{inicio}:{col}{fim}, TRUE)', "formato": None},
    {"nome": "DESISTENTES", "formula": lambda col, inicio, fim: f'=COUNTIF({col}{inicio}:{col}{fim}, TRUE)', "formato": None}
]

# Definição do dashboard geral da escola (Resumo Geral da Escola)
DASHBOARD_SEC_GERAL = [
    {"nome": "MATRÍCULAS", "formula": None, "formato": None},
    {"nome": "ATIVOS", "formula": None, "formato": None},
    {"nome": "TRANSFERIDOS", "formula": None, "formato": None},
    {"nome": "DESISTENTES", "formula": None, "formato": None},
    {"nome": "Nº ABANDONO(S)", "formula": None, "formato": None},
    {"nome": "PORCENTAGEM DE ABANDONO(S)", "formula": None, "formato": '0.00%'}
]