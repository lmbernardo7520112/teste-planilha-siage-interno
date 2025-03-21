import pandas as pd
from openpyxl import Workbook
from openpyxl.drawing.image import Image
from openpyxl.styles import Font, Alignment
import os

# Lista de disciplinas
DISCIPLINAS = ["BIO", "MAT", "FIS", "QUI", "GEO", "SOC", "HIST", "FIL", "ESP", "POR", "ART", "ADF", "ING"]

# Colunas da planilha
COLUNAS = [
    "Nº", "Nome do Aluno", "1º BIM", "2º BIM", "3º BIM", "4º BIM",
    "NF", "MG", "MF", "SITUAÇÃO DO ALUNO", "PF", "SF"
]

# Caminho da imagem
CAMINHO_IMAGEM = os.path.expanduser("~/teste-planilha-siage-interno/siage_interno.png")

# Função para criar uma aba com cabeçalho e imagem
def criar_aba(wb, titulo, img):
    """
    Cria uma nova aba no Workbook com o título especificado,
    adiciona a imagem, o título e o cabeçalho.
    """
    ws = wb.create_sheet(title=titulo)
    
    # Mescla as células da primeira linha de A a J
    ws.merge_cells('A1:J1')
    
    # Ajusta a altura da linha mesclada para caber a imagem
    ws.row_dimensions[1].height = img.height * 0.75  # Ajusta a altura da linha (em pontos)
    
    # Adiciona a imagem na célula mesclada
    ws.add_image(img, 'A1')
    
    # Adiciona o texto "COMPOSITOR LUIS RAMALHO" na célula mesclada
    cell = ws['A1']
    cell.value = "COMPOSITOR LUIS RAMALHO"
    cell.font = Font(name='Arial', size=26, bold=True)
    cell.alignment = Alignment(horizontal='center', vertical='center')
    
    # Desloca o cabeçalho 10 células para baixo (linha 11)
    for _ in range(10):  # Adiciona 10 linhas em branco
        ws.append([])
    
    # Adiciona o cabeçalho (nomes das colunas) na linha 11
    ws.append(COLUNAS)
    
    return ws

# Função para adicionar fórmulas nas células
def adicionar_formulas(ws):
    """
    Adiciona fórmulas para calcular as médias e situações nas células.
    """
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

# Função principal
def criar_planilha():
    """
    Função principal para criar a planilha de notas.
    """
    # Cria um DataFrame vazio com as colunas definidas
    df = pd.DataFrame(columns=COLUNAS)

    # Adiciona números de 1 a 35 na coluna "Nº"
    df["Nº"] = range(1, 36)

    # Cria um arquivo Excel usando openpyxl
    wb = Workbook()

    # Remove a sheet padrão criada automaticamente
    wb.remove(wb.active)

    # Verifica se a imagem existe
    if not os.path.exists(CAMINHO_IMAGEM):
        raise FileNotFoundError(f"A imagem não foi encontrada no caminho: {CAMINHO_IMAGEM}")

    # Carrega a imagem
    img = Image(CAMINHO_IMAGEM)

    # Reduz a imagem em 50%
    img.width = int(img.width * 0.5)  # Nova largura: 540 pixels
    img.height = int(img.height * 0.5)  # Nova altura: 106 pixels

    # Cria a aba "SEC" (Secretaria Escolar)
    ws_sec = criar_aba(wb, "SEC", img)

    # Cria uma sheet para cada disciplina
    for disciplina in DISCIPLINAS:
        ws = criar_aba(wb, disciplina, img)
        adicionar_formulas(ws)

    # Cria as abas adicionais em branco
    abas_adicionais = ["INDIVIDUAL", "BOLETIM", "BOL", "RESULTADO", "FREQUÊNCIA"]
    for aba in abas_adicionais:
        criar_aba(wb, aba, img)

    # Define o caminho onde o arquivo será salvo
    caminho_padrao = "/mnt/c/Users/lmbernardo/Downloads"  # Caminho no WSL2 para a pasta Downloads do Windows
    nome_arquivo = "planilha_notas_complexa.xlsx"
    caminho_completo = os.path.join(caminho_padrao, nome_arquivo)

    # Salva o arquivo Excel
    wb.save(caminho_completo)

    print(f"Arquivo Excel salvo em: {caminho_completo}")

# Executa a função principal
if __name__ == "__main__":
    criar_planilha()