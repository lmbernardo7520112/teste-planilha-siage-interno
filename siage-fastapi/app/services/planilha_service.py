import os
import logging
from pathlib import Path
from openpyxl.styles import Font, Alignment, Border, Side  # Adicionado Border e Side
from openpyxl.utils import get_column_letter
from openpyxl import Workbook
from openpyxl.drawing.image import Image
from app.utils.excel_utils import configurar_largura_colunas
from app.core.config import COLUNAS, DISCIPLINAS, CAMINHO_IMAGEM, CAMINHO_PADRAO, NOME_ARQUIVO_PADRAO, LARGURAS_COLUNAS

def criar_aba_em_branco(wb, titulo, img):
    """
    Cria uma nova aba no Workbook com o título especificado,
    adiciona a imagem e o título, mas sem cabeçalho ou fórmulas.
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
    
    return ws

def criar_aba_disciplina(wb, titulo, caminho_imagem, contador_imagem):
    """
    Cria uma nova aba no Workbook com o título especificado,
    adiciona a imagem, o título, o cabeçalho, as fórmulas e bordas na tabela.
    """
    ws = wb.create_sheet(title=titulo)
    
    # Mescla as células da primeira linha de A a J
    ws.merge_cells('A1:J1')
    
    # Ajusta a altura da linha mesclada para caber a imagem
    ws.row_dimensions[1].height = 80  # Ajusta a altura da linha (em pontos)
    
    # Carrega a imagem a partir do caminho
    img = Image(caminho_imagem)
    
    # Reduz a imagem em 50%
    img.width = int(img.width * 0.5)  # Nova largura: 50% do original
    img.height = int(img.height * 0.5)  # Nova altura: 50% do original
    
    # Adiciona a imagem na célula mesclada
    ws.add_image(img, 'A1')
    
    # Adiciona o texto "COMPOSITOR LUIS RAMALHO" na célula mesclada
    cell = ws['A1']
    cell.value = "COMPOSITOR LUIS RAMALHO"
    cell.font = Font(name='Arial', size=26, bold=True)
    cell.alignment = Alignment(horizontal='center', vertical='center')
    
    # Desloca o cabeçalho para a linha 12
    for _ in range(10):  # Adiciona 10 linhas em branco
        ws.append([])
    
    # Adiciona o cabeçalho (nomes das colunas) na linha 12
    ws.append(COLUNAS)
    
    # Usa LARGURAS_COLUNAS em vez do dicionário hardcoded
    configurar_largura_colunas(ws, LARGURAS_COLUNAS)
    
    # Define o estilo de contorno
    border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    
    # Adiciona os números de 1 a 35 na coluna "Nº" (a partir da linha 13)
    for i in range(1, 36):
        ws[f'A{i + 12}'] = i  # Linha 13 em diante
    
    # Adiciona fórmulas para calcular as médias e situações
    for row in range(13, 48):  # 35 alunos (linhas 13 a 47)
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
    
    # Aplica bordas às células da tabela (linhas 12 a 47, colunas A a L)
    for row in range(12, 48):  # Linha 12 (cabeçalho) até linha 47 (35 alunos)
        for col in range(1, 13):  # Colunas A (1) até L (12)
            cell = ws[f'{get_column_letter(col)}{row}']
            cell.border = border
    
    return ws

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

def criar_planilha():
    logger.info("Iniciando criação da planilha")
    wb = Workbook()
    wb.remove(wb.active)

    if not CAMINHO_IMAGEM.exists():
        raise FileNotFoundError(f"A imagem não foi encontrada no caminho: {CAMINHO_IMAGEM}")

    img = Image(str(CAMINHO_IMAGEM))
    criar_aba_em_branco(wb, "SEC", img)

    contador_imagem = 1
    for disciplina in DISCIPLINAS:
        criar_aba_disciplina(wb, disciplina, str(CAMINHO_IMAGEM), contador_imagem)
        contador_imagem += 1

    abas_adicionais = ["INDIVIDUAL", "BOLETIM", "BOL", "RESULTADO", "FREQUÊNCIA"]
    for aba in abas_adicionais:
        criar_aba_em_branco(wb, aba, img)

    caminho_completo = os.path.join(CAMINHO_PADRAO, NOME_ARQUIVO_PADRAO)
    wb.save(caminho_completo)
    logger.info(f"Planilha salva em: {caminho_completo}")
    return caminho_completo