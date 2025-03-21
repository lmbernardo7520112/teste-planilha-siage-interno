import os
from pathlib import Path
from openpyxl import Workbook
from openpyxl.drawing.image import Image
from openpyxl.styles import Font, Alignment
from openpyxl.utils import get_column_letter
from app.utils.excel_utils import configurar_largura_colunas
from app.core.config import COLUNAS, DISCIPLINAS  # Importa de config


# Caminho da imagem (usando pathlib para garantir o caminho correto)
CAMINHO_IMAGEM = Path(__file__).parent.parent.parent / "app" / "core" / "static" / "images" / "siage_interno.png"

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
    adiciona a imagem, o título, o cabeçalho e as fórmulas.
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
    
    # Configura a largura das colunas específicas
    configurar_largura_colunas(ws, {
        "Nome do Aluno": 10,  # 10 cm
        "SITUAÇÃO DO ALUNO": 4.5  # 4.5 cm
    })
    
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
    
    return ws

def criar_planilha():
    """
    Função principal para criar a planilha de notas.
    """
    # Cria um arquivo Excel usando openpyxl
    wb = Workbook()

    # Remove a sheet padrão criada automaticamente
    wb.remove(wb.active)

    # Verifica se a imagem existe
    if not CAMINHO_IMAGEM.exists():
        raise FileNotFoundError(f"A imagem não foi encontrada no caminho: {CAMINHO_IMAGEM}")

    # Cria a aba "SEC" (Secretaria Escolar) em branco
    img = Image(str(CAMINHO_IMAGEM))  # Converte o caminho para string
    criar_aba_em_branco(wb, "SEC", img)

    # Cria uma sheet para cada disciplina (com cabeçalho e fórmulas)
    contador_imagem = 1  # Contador para garantir nomes únicos para as imagens
    for disciplina in DISCIPLINAS:
        criar_aba_disciplina(wb, disciplina, str(CAMINHO_IMAGEM), contador_imagem)
        contador_imagem += 1

    # Cria as abas adicionais em branco
    abas_adicionais = ["INDIVIDUAL", "BOLETIM", "BOL", "RESULTADO", "FREQUÊNCIA"]
    for aba in abas_adicionais:
        criar_aba_em_branco(wb, aba, img)

    # Define o caminho onde o arquivo será salvo
    caminho_padrao = "/mnt/c/Users/lmbernardo/Downloads"  # Caminho no WSL2 para a pasta Downloads do Windows
    nome_arquivo = "planilha_notas_complexa.xlsx"
    caminho_completo = os.path.join(caminho_padrao, nome_arquivo)

    # Salva o arquivo Excel
    wb.save(caminho_completo)

    return caminho_completo