import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.drawing.image import Image
import os

# Lista de disciplinas
disciplinas = ["BIO", "MAT", "FIS", "QUI", "GEO", "SOC", "HIST", "FIL", "ESP", "POR", "ART", "ADF", "ING"]

# Colunas da planilha (apenas para as disciplinas)
colunas = [
    "Nº", "Nome do Aluno", "1º BIM", "2º BIM", "3º BIM", "4º BIM",
    "NF", "MG", "MF", "SITUAÇÃO DO ALUNO", "PF", "SF"
]

# Cria um DataFrame vazio com as colunas definidas
df = pd.DataFrame(columns=colunas)

# Adiciona números de 1 a 35 na coluna "Nº"
df["Nº"] = range(1, 36)

# Cria um arquivo Excel usando openpyxl
wb = Workbook()

# Remove a sheet padrão criada automaticamente
wb.remove(wb.active)

# Caminho da imagem
caminho_imagem = os.path.expanduser("~/teste-planilha-siage-interno/siage_interno.png")

# Verifica se a imagem existe
if not os.path.exists(caminho_imagem):
    raise FileNotFoundError(f"A imagem não foi encontrada no caminho: {caminho_imagem}")

# Carrega a imagem
img = Image(caminho_imagem)

# Reduz a imagem em 50%
img.width = int(img.width * 0.5)  # Nova largura: 540 pixels
img.height = int(img.height * 0.5)  # Nova altura: 106 pixels

# Cria a aba "SEC" (Secretaria Escolar) em branco
ws_sec = wb.create_sheet(title="SEC")

# Adiciona a imagem na célula A1 da aba "SEC"
ws_sec.add_image(img, 'A1')

# Ajusta a altura da linha e a largura da coluna para caber a imagem reduzida
ws_sec.row_dimensions[1].height = img.height * 0.75  # Ajusta a altura da linha (em pontos)
ws_sec.column_dimensions['A'].width = img.width / 7  # Ajusta a largura da coluna (em caracteres)

# Cria uma sheet para cada disciplina
for disciplina in disciplinas:
    ws = wb.create_sheet(title=disciplina)
    
    # Adiciona a imagem na célula A1 da aba da disciplina
    ws.add_image(img, 'A1')
    
    # Ajusta a altura da linha e a largura da coluna para caber a imagem reduzida
    ws.row_dimensions[1].height = img.height * 0.75  # Ajusta a altura da linha (em pontos)
    ws.column_dimensions['A'].width = img.width / 7  # Ajusta a largura da coluna (em caracteres)
    
    # Desloca o cabeçalho 10 células para baixo (linha 11)
    for _ in range(10):  # Adiciona 10 linhas em branco
        ws.append([])
    
    # Adiciona o cabeçalho (nomes das colunas) na linha 11
    ws.append(colunas)
    
    # Adiciona fórmulas para calcular as médias e situações
    for row in range(12, 47):  # 35 alunos (linhas 12 a 46, devido ao deslocamento)
        # Fórmula para Nota Final (NF) - Média dos 4 bimestres
        ws[f'G{row}'] = f'=AVERAGE(C{row}:F{row})'
        
        # Fórmula para Média Geral (MG) - Média de todas as disciplinas
        ws[f'H{row}'] = f'=SUM(C{row}:F{row})/4'
        
        # Fórmula para Média Final (MF)
        ws[f'I{row}'] = f'=IF(H{row}<7, (0.6*H{row}) + (0.4*G{row}), "-")'
        
        # Fórmula para Situação do Aluno (SITUAÇÃO DO ALUNO)
        ws[f'J{row}'] = f'=IF(H{row}<2.5, "REPROVADO", IF(H{row}<7, "FINAL", "APROVADO"))'
        
        # Fórmula para Prova Final (PF)
        ws[f'K{row}'] = f'=IF(H{row}<7, (12.5 - (1.5*H{row})), "-")'
        
        # Fórmula para Situação Final (SF)
        ws[f'L{row}'] = f'=IF(G{row}>=K{row}, "AF", "-")'

# Cria as abas adicionais em branco
abas_adicionais = ["INDIVIDUAL", "BOLETIM", "BOL", "RESULTADO", "FREQUÊNCIA"]
for aba in abas_adicionais:
    ws = wb.create_sheet(title=aba)
    
    # Adiciona a imagem na célula A1 da aba adicional
    ws.add_image(img, 'A1')
    
    # Ajusta a altura da linha e a largura da coluna para caber a imagem reduzida
    ws.row_dimensions[1].height = img.height * 0.75  # Ajusta a altura da linha (em pontos)
    ws.column_dimensions['A'].width = img.width / 7  # Ajusta a largura da coluna (em caracteres)

# Define o caminho onde o arquivo será salvo
caminho_padrao = "/mnt/c/Users/lmbernardo/Downloads"  # Caminho no WSL2 para a pasta Downloads do Windows
nome_arquivo = "planilha_notas_complexa.xlsx"
caminho_completo = os.path.join(caminho_padrao, nome_arquivo)

# Salva o arquivo Excel
wb.save(caminho_completo)

print(f"Arquivo Excel salvo em: {caminho_completo}")