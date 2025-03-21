from pathlib import Path

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
CAMINHO_PADRAO = "/mnt/c/Users/lmbernardo/Downloads"
NOME_ARQUIVO_PADRAO = "planilha_notas_complexa.xlsx"

# Larguras das colunas (em cm)
LARGURAS_COLUNAS = {
    "Nome do Aluno": 10,
    "SITUAÇÃO DO ALUNO": 4.5
}