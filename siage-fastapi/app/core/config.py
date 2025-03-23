from pathlib import Path
from openpyxl.styles import PatternFill, Font

# Lista de disciplinas
DISCIPLINAS = ["BIO", "MAT", "FIS", "QUI", "GEO", "SOC", "HIST", "FIL", "ESP", "POR", "ART", "ADF", "ING"]

# Colunas da planilha
COLUNAS = [
    "Nº", "Nome do Aluno", "1º BIM", "2º BIM", "3º BIM", "4º BIM",
    "NF", "MG", "MF", "SITUAÇÃO DO ALUNO", "PF", "SF"
]

# Caminho da imagem
CAMINHO_IMAGEM = Path(__file__).parent / "static" / "images" / "siage_interno.png"

# Configurações de salvamento
CAMINHO_PADRAO = "/mnt/c/Users/lmbernardo/Downloads"  # Ajuste para Windows se necessário
NOME_ARQUIVO_PADRAO = "planilha_notas_complexa.xlsx"

# Larguras das colunas (em cm)
LARGURAS_COLUNAS = {
    "Nº": 1,
    "Nome do Aluno": 10,
    "SITUAÇÃO DO ALUNO": 4.5
}

# Definições de cores
COR_ABA = "FFDAB9"  # Laranja claro (Peach Puff)
FILL_NOME_ALUNO = PatternFill(start_color="FF6347", end_color="FF6347", fill_type="solid")  # Vermelho-alaranjado (Tomato)
FILL_BIMESTRES = PatternFill(start_color="FFA500", end_color="FFA500", fill_type="solid")  # Amarelo-alaranjado (Orange)
FILL_NOTA_FINAL = PatternFill(start_color="FF4500", end_color="FF4500", fill_type="solid")  # Laranja escuro (Orange Red)
FILL_SITUACAO = PatternFill(start_color="FFFACD", end_color="FFFACD", fill_type="solid")  # Amarelo claro (Lemon Chiffon)
FONTE_TITULO_TURMA = Font(name='Arial', size=14, bold=True, color="8B4513")  # Marrom escuro (Saddle Brown)

# Definição dos indicadores do dashboard
DASHBOARD_INDICADORES = [
    {
        "nome": "ALUNOS APROVADOS",
        "formula": lambda col, inicio, fim: f'=COUNTIF({col}{inicio}:{col}{fim}, ">=7")',
        "formato": None
    },
    {
        "nome": "ALUNOS REPROVADOS",
        "formula": lambda col, inicio, fim: f'=COUNTIF({col}{inicio}:{col}{fim}, "<7")',
        "formato": None
    },
    {
        "nome": "Nº ALUNOS COM MÉDIA > 8,0",
        "formula": lambda col, inicio, fim: f'=COUNTIF({col}{inicio}:{col}{fim}, ">=8")',
        "formato": None
    },
    {
        "nome": "Nº ALUNOS QUE NÃO ATINGIRAM MÉDIA > 8,0",
        "formula": lambda col, inicio, fim: f'=COUNTIF({col}{inicio}:{col}{fim}, "<8")',
        "formato": None
    },
    {
        "nome": "PERCENTUAL DE MÉDIAS > 5,0",
        "formula": lambda col, inicio, fim: f'=COUNTIF({col}{inicio}:{col}{fim}, ">=5")/COUNTA({col}{inicio}:{col}{fim})',
        "formato": '0.00%'
    },
    {
        "nome": "PERCENTUAL DE MÉDIAS < 5,0",
        "formula": lambda col, inicio, fim: f'=COUNTIF({col}{inicio}:{col}{fim}, "<5")/COUNTA({col}{inicio}:{col}{fim})',
        "formato": '0.00%'
    },
    {
        "nome": "MATRÍCULAS",
        "formula": None,  # Calculado dinamicamente
        "formato": None
    },
    {
        "nome": "TAXA DE APROVAÇÃO (%)",
        "formula": None,  # Calculado dinamicamente
        "formato": '0.00%'
    }
]