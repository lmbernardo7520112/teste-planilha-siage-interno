from pathlib import Path
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side

# --- CAMINHOS ESSENCIAIS ---
BASE_DIR = Path(__file__).resolve().parent.parent.parent
CAMINHO_JSON = BASE_DIR / "turmas_alunos.json"
CAMINHO_IMAGEM = BASE_DIR / "app/core/static/images/siage_interno.png"
CAMINHO_PADRAO = BASE_DIR
NOME_ARQUIVO_PADRAO = "planilha_notas_gerada_final_v5.xlsx"

# Constante para número máximo de linhas de alunos a formatar por turma
MAX_ALUNOS_FORMATAR = 35

# Lista de disciplinas (Códigos)
DISCIPLINAS = [
    "BIO", "MAT", "FIS", "QUI", "GEO", "SOC", "HIS",
    "FIL", "ESP", "POR", "ART", "EDF", "ING"
]

# Mapeamento Código -> Nome Completo
DISCIPLINAS_NOMES = {
    "BIO": "Biologia", "MAT": "Matemática", "FIS": "Física", "QUI": "Química",
    "GEO": "Geografia", "SOC": "Sociologia", "HIS": "História", "FIL": "Filosofia",
    "ESP": "Espanhol", "POR": "Português", "ART": "Artes", "EDF": "Educação Física",
    "ING": "Inglês"
}

# Colunas das abas de disciplinas
COLUNAS = [
    "Nº", "Nome do Aluno", "1º BIM", "2º BIM", "3º BIM", "4º BIM",
    "NF", "MG", "MF", "SITUAÇÃO DO ALUNO", "PF", "SF"
]

# Colunas da aba SEC
COLUNAS_SEC = [
    "Nº", "Nome do Aluno", "ATIVO", "TRANSFERIDO", "DESISTENTE", "SITUAÇÃO DO ALUNO"
]

# --- LARGURAS DAS COLUNAS POR LETRA ---
LARGURAS_COLUNAS_SEC_LETRAS = {
    'A': 5,  # Nº
    'B': 35, # Nome do Aluno
    'C': 10, # ATIVO
    'D': 12, # TRANSFERIDO
    'E': 12, # DESISTENTE
    'F': 15  # SITUAÇÃO DO ALUNO
}

LARGURAS_COLUNAS_ABAS_DISC_LETRAS = {
    'A': 5,  # Nº
    'B': 35, # Nome do Aluno
    'C': 8,  # 1º BIM
    'D': 8,  # 2º BIM
    'E': 8,  # 3º BIM
    'F': 8,  # 4º BIM
    'G': 8,  # NF
    'H': 8,  # MG
    'I': 8,  # MF
    'J': 15, # SITUAÇÃO DO ALUNO
    'K': 8,  # PF
    'L': 8   # SF
}

# Estilos
COR_ABA = "FFDAB9"
FILL_NOME_ALUNO = PatternFill(start_color="ADD8E6", end_color="ADD8E6", fill_type="solid")
FILL_BIMESTRES = PatternFill(start_color="FFFFE0", end_color="FFFFE0", fill_type="solid")
FILL_NOTA_FINAL = PatternFill(start_color="E0FFFF", end_color="E0FFFF", fill_type="solid")
FILL_SITUACAO = PatternFill(start_color="FFFACD", end_color="FFFACD", fill_type="solid")
FONTE_TITULO_TURMA = Font(name='Arial', size=14, bold=True, color="000080")
ALINHAMENTO_CENTRALIZADO = Alignment(horizontal='center', vertical='center', wrap_text=True)
BORDER_THIN = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

# --- Dashboards Config ---
DASHBOARD_INDICADORES = [
    {"nome": "ALUNOS APROVADOS", "formula": lambda col, inicio, fim, ws_title: f'=COUNTIF(INDIRECT("\'{ws_title}\'!{col}{inicio}:{col}{fim}"), ">=7")', "formato": '0'},
    {"nome": "ALUNOS REPROVADOS", "formula": lambda col, inicio, fim, ws_title: f'=COUNTIF(INDIRECT("\'{ws_title}\'!{col}{inicio}:{col}{fim}"), "<7")', "formato": '0'},
    {"nome": "Nº ALUNOS COM MÉDIA > 8,0", "formula": lambda col, inicio, fim, ws_title: f'=COUNTIF(INDIRECT("\'{ws_title}\'!{col}{inicio}:{col}{fim}"), ">=8")', "formato": '0'},
    {"nome": "Nº ALUNOS QUE NÃO ATINGIRAM MÉDIA > 8,0", "formula": lambda col, inicio, fim, ws_title: f'=COUNTIF(INDIRECT("\'{ws_title}\'!{col}{inicio}:{col}{fim}"), "<8")', "formato": '0'},
    {"nome": "PERCENTUAL DE MÉDIAS > 5,0", "formula": lambda col, inicio, fim, ws_title: f'=IFERROR(COUNTIF(INDIRECT("\'{ws_title}\'!{col}{inicio}:{col}{fim}"), ">=5")/MAX(1,COUNTIF(INDIRECT("\'{ws_title}\'!{col}{inicio}:{col}{fim}"), "<>")),0)', "formato": '0.00%'},
    {"nome": "PERCENTUAL DE MÉDIAS < 5,0", "formula": lambda col, inicio, fim, ws_title: f'=IFERROR(COUNTIF(INDIRECT("\'{ws_title}\'!{col}{inicio}:{col}{fim}"), "<5")/MAX(1,COUNTIF(INDIRECT("\'{ws_title}\'!{col}{inicio}:{col}{fim}"), "<>")),0)', "formato": '0.00%'},
    {"nome": "MATRÍCULAS", "formula": lambda col, inicio, fim, ws_title: f'=COUNTA(INDIRECT("\'{ws_title}\'!B{inicio}:B{fim}"))', "formato": '0'},
    {"nome": "TAXA DE APROVAÇÃO (%)", "formula": lambda col, inicio, fim, ws_title: f'=IFERROR(COUNTIF(INDIRECT("\'{ws_title}\'!{col}{inicio}:{col}{fim}"), ">=7")/MAX(1,COUNTIF(INDIRECT("\'{ws_title}\'!{col}{inicio}:{col}{fim}"), "<>")),0)', "formato": '0.00%'}
]
DASHBOARD_SEC_TURMA = [
    {"nome": "MATRÍCULAS", "formula": lambda c, i, f: f'=COUNTA(B{i}:B{f})', "formato": '0'},
    {"nome": "ATIVOS", "formula": lambda c, i, f: f'=COUNTIF(C{i}:C{f}, TRUE)', "formato": '0'},
    {"nome": "TRANSFERIDOS", "formula": lambda c, i, f: f'=COUNTIF(D{i}:D{f}, TRUE)', "formato": '0'},
    {"nome": "DESISTENTES", "formula": lambda c, i, f: f'=COUNTIF(E{i}:E{f}, TRUE)', "formato": '0'},
]
DASHBOARD_SEC_GERAL = [
    {"nome": "MATRÍCULAS", "formula": lambda r: f'=SUM({",".join(r)})', "formato": '0'},
    {"nome": "ATIVOS", "formula": lambda r: f'=SUM({",".join(r)})', "formato": '0'},
    {"nome": "TRANSFERIDOS", "formula": lambda r: f'=SUM({",".join(r)})', "formato": '0'},
    {"nome": "DESISTENTES", "formula": lambda r: f'=SUM({",".join(r)})', "formato": '0'},
    {"nome": "Nº ABANDONO(S)", "formula": lambda l, c: f'={c}{l-1}', "formato": '0'},
    {"nome": "ABANDONO(S) (%)", "formula": lambda l, c: f'=IFERROR({c}{l-1}/MAX(1,{c}{l-4}), 0)', "formato": '0.00%'}
]
DASHBOARD_SEC_APROVACAO = [
    {"nome": "TX APROVAÇÃO %", "formato": '0.00%'},
    {"nome": "TX REPROVAÇÃO %", "formato": '0.00%'}
]

# --- Power Pivot ---
TBL_TURMAS_NAME = "tblTurmas"
TBL_ALUNOS_NAME = "tblAlunos"
TBL_DISCIPLINAS_NAME = "tblDisciplinas"
TBL_NOTAS_NAME = "tblNotas"