import os
import json
import logging
from pathlib import Path
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl import Workbook
from openpyxl.drawing.image import Image
from app.utils.excel_utils import configurar_largura_colunas
from app.core.config import COLUNAS, DISCIPLINAS, CAMINHO_IMAGEM, CAMINHO_PADRAO, NOME_ARQUIVO_PADRAO, LARGURAS_COLUNAS

# Caminho do arquivo JSON (assumindo que está no diretório raiz do projeto)
CAMINHO_JSON = Path(__file__).parent.parent.parent / "turmas_alunos.json"

def criar_aba_em_branco(wb, titulo, img):
    """
    Cria uma nova aba no Workbook com o título especificado,
    adiciona a imagem e o título, mas sem cabeçalho ou fórmulas.
    """
    ws = wb.create_sheet(title=titulo)
    
    ws.merge_cells('A1:J1')
    ws.row_dimensions[1].height = img.height * 0.75
    ws.add_image(img, 'A1')
    
    cell = ws['A1']
    cell.value = "COMPOSITOR LUIS RAMALHO"
    cell.font = Font(name='Arial', size=26, bold=True)
    cell.alignment = Alignment(horizontal='center', vertical='center')
    
    return ws

def criar_aba_disciplina(wb, titulo, caminho_imagem, turmas):
    """
    Cria uma aba para a disciplina com uma tabela para cada turma,
    populada com os alunos de cada turma, com 15 linhas de espaço entre tabelas.
    """
    ws = wb.create_sheet(title=titulo)
    
    # Cabeçalho superior com imagem
    ws.merge_cells('A1:J1')
    ws.row_dimensions[1].height = 80
    img = Image(caminho_imagem)
    img.width = int(img.width * 0.5)
    img.height = int(img.height * 0.5)
    ws.add_image(img, 'A1')
    
    cell = ws['A1']
    cell.value = "COMPOSITOR LUIS RAMALHO"
    cell.font = Font(name='Arial', size=26, bold=True)
    cell.alignment = Alignment(horizontal='center', vertical='center')
    
    # Estilo de borda
    border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    
    # Espaçamento inicial
    linha_atual = 2  # Começa após o cabeçalho
    
    # Itera sobre cada turma
    for turma in turmas:
        # Adiciona o nome da turma como título
        ws.merge_cells(f'A{linha_atual}:L{linha_atual}')
        ws[f'A{linha_atual}'] = turma["nome_turma"]
        ws[f'A{linha_atual}'].font = Font(name='Arial', size=14, bold=True)
        ws[f'A{linha_atual}'].alignment = Alignment(horizontal='center')
        ws.row_dimensions[linha_atual].height = 30
        linha_atual += 1
        
        # Adiciona o cabeçalho da tabela
        for col_idx, col_nome in enumerate(COLUNAS, 1):
            cell = ws[f'{get_column_letter(col_idx)}{linha_atual}']
            cell.value = col_nome
            cell.border = border
        
        # Configura largura das colunas
        configurar_largura_colunas(ws, LARGURAS_COLUNAS)
        
        # Linha inicial dos dados
        linha_inicio_dados = linha_atual + 1
        
        # Popula os alunos
        for aluno in turma["alunos"]:
            ws[f'A{linha_atual + int(aluno["numero"])}'] = int(aluno["numero"])  # Número do aluno
            ws[f'B{linha_atual + int(aluno["numero"])}'] = aluno["nome"]        # Nome do aluno
        
        # Adiciona fórmulas para todas as linhas (até 35, mesmo que a turma tenha menos alunos)
        for row in range(linha_inicio_dados, linha_inicio_dados + 35):
            ws[f'G{row}'] = f'=AVERAGE(C{row}:F{row})'  # NF
            ws[f'H{row}'] = f'=SUM(C{row}:F{row})/4'    # MG
            ws[f'I{row}'] = f'=IF(H{row}<7, (0.6*H{row}) + (0.4*G{row}), "-")'  # MF
            ws[f'J{row}'] = f'=IF(H{row}<2.5, "REPROVADO", IF(H{row}<7, "FINAL", "APROVADO"))'  # Situação
            ws[f'K{row}'] = f'=IF(H{row}<7, (12.5 - (1.5*H{row})), "-")'  # PF
            ws[f'L{row}'] = f'=IF(G{row}>=K{row}, "AF", "-")'  # SF
        
        # Aplica bordas à tabela (cabeçalho + 35 linhas)
        for row in range(linha_atual, linha_atual + 36):
            for col in range(1, 13):
                cell = ws[f'{get_column_letter(col)}{row}']
                cell.border = border
        
        # Atualiza a linha atual para a próxima tabela, com 15 linhas de espaço
        linha_atual += 36 + 15  # 1 linha cabeçalho + 35 linhas de dados + 15 linhas de espaço
    
    return ws

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

def criar_planilha():
    logger.info("Iniciando criação da planilha")
    wb = Workbook()
    wb.remove(wb.active)

    if not CAMINHO_IMAGEM.exists():
        raise FileNotFoundError(f"A imagem não foi encontrada no caminho: {CAMINHO_IMAGEM}")

    # Carrega os dados do JSON
    if not CAMINHO_JSON.exists():
        raise FileNotFoundError(f"O arquivo JSON não foi encontrado em: {CAMINHO_JSON}")
    with open(CAMINHO_JSON, 'r', encoding='utf-8') as f:
        dados = json.load(f)
    turmas = dados["turmas"]

    img = Image(str(CAMINHO_IMAGEM))
    criar_aba_em_branco(wb, "SEC", img)

    # Cria uma aba para cada disciplina, com 7 tabelas (uma por turma)
    for disciplina in DISCIPLINAS:
        criar_aba_disciplina(wb, disciplina, str(CAMINHO_IMAGEM), turmas)

    abas_adicionais = ["INDIVIDUAL", "BOLETIM", "BOL", "RESULTADO", "FREQUÊNCIA"]
    for aba in abas_adicionais:
        criar_aba_em_branco(wb, aba, img)

    caminho_completo = os.path.join(CAMINHO_PADRAO, NOME_ARQUIVO_PADRAO)
    wb.save(caminho_completo)
    logger.info(f"Planilha salva em: {caminho_completo}")
    return caminho_completo