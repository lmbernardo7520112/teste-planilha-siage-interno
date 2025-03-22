import os
import json
import logging
from pathlib import Path
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl import Workbook
from openpyxl.drawing.image import Image
from app.utils.excel_utils import configurar_largura_colunas
from app.core.config import (
    COLUNAS, DISCIPLINAS, CAMINHO_IMAGEM, CAMINHO_PADRAO, NOME_ARQUIVO_PADRAO, LARGURAS_COLUNAS,
    COR_ABA, FILL_NOME_ALUNO, FILL_BIMESTRES, FILL_NOTA_FINAL, FILL_SITUACAO, FONTE_TITULO_TURMA
)

# Caminho do arquivo JSON
CAMINHO_JSON = Path(__file__).parent.parent.parent / "turmas_alunos.json"

def criar_aba_em_branco(wb, titulo, img):
    ws = wb.create_sheet(title=titulo)
    ws.merge_cells('A1:J1')
    ws.row_dimensions[1].height = img.height * 0.75
    ws.add_image(img, 'A1')
    
    cell = ws['A1']
    cell.value = "COMPOSITOR LUIS RAMALHO"
    cell.font = Font(name='Arial', size=26, bold=True)
    cell.alignment = Alignment(horizontal='center', vertical='center')
    
    # Cor da aba
    ws.sheet_properties.tabColor = COR_ABA
    
    return ws

def criar_aba_disciplina(wb, titulo, caminho_imagem, turmas):
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
    
    # Cor da aba
    ws.sheet_properties.tabColor = COR_ABA
    
    # Estilo de borda
    border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    
    # Espaçamento inicial
    linha_atual = 2
    
    # Itera sobre cada turma
    for turma in turmas:
        # Título da turma
        ws.merge_cells(f'A{linha_atual}:L{linha_atual}')
        ws[f'A{linha_atual}'] = turma["nome_turma"]
        ws[f'A{linha_atual}'].font = FONTE_TITULO_TURMA
        ws[f'A{linha_atual}'].alignment = Alignment(horizontal='center')
        ws.row_dimensions[linha_atual].height = 30
        linha_atual += 1
        
        # Adiciona o cabeçalho da tabela com cores
        for col_idx, col_nome in enumerate(COLUNAS, 1):
            cell = ws[f'{get_column_letter(col_idx)}{linha_atual}']
            cell.value = col_nome
            cell.border = border
            cell.font = Font(bold=True)
            # Aplica cores específicas
            if col_nome == "Nome do Aluno":
                cell.fill = FILL_NOME_ALUNO
            elif col_nome in ["1º BIM", "2º BIM", "3º BIM", "4º BIM"]:
                cell.fill = FILL_BIMESTRES
            elif col_nome == "NF":
                cell.fill = FILL_NOTA_FINAL
            elif col_nome == "SITUAÇÃO DO ALUNO":
                cell.fill = FILL_SITUACAO
        
        # Configura largura das colunas
        configurar_largura_colunas(ws, LARGURAS_COLUNAS)
        
        # Linha inicial dos dados
        linha_inicio_dados = linha_atual + 1
        
        # Popula os alunos
        for aluno in turma["alunos"]:
            ws[f'A{linha_atual + int(aluno["numero"])}'] = int(aluno["numero"])
            ws[f'B{linha_atual + int(aluno["numero"])}'] = aluno["nome"]
        
        # Adiciona fórmulas e bordas
        for row in range(linha_inicio_dados, linha_inicio_dados + 35):
            ws[f'G{row}'] = f'=AVERAGE(C{row}:F{row})'
            ws[f'H{row}'] = f'=SUM(C{row}:F{row})/4'
            ws[f'I{row}'] = f'=IF(H{row}<7, (0.6*H{row}) + (0.4*G{row}), "-")'
            ws[f'J{row}'] = f'=IF(H{row}<2.5, "REPROVADO", IF(H{row}<7, "FINAL", "APROVADO"))'
            ws[f'K{row}'] = f'=IF(H{row}<7, (12.5 - (1.5*H{row})), "-")'
            ws[f'L{row}'] = f'=IF(G{row}>=K{row}, "AF", "-")'
        
        # Aplica bordas à tabela
        for row in range(linha_atual, linha_atual + 36):
            for col in range(1, 13):
                cell = ws[f'{get_column_letter(col)}{row}']
                cell.border = border
        
        # Espaço de 15 linhas entre tabelas
        linha_atual += 36 + 15
    
    return ws

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

def criar_planilha():
    logger.info("Iniciando criação da planilha")
    wb = Workbook()
    wb.remove(wb.active)

    if not CAMINHO_IMAGEM.exists():
        raise FileNotFoundError(f"A imagem não foi encontrada no caminho: {CAMINHO_IMAGEM}")

    if not CAMINHO_JSON.exists():
        raise FileNotFoundError(f"O arquivo JSON não foi encontrado em: {CAMINHO_JSON}")
    with open(CAMINHO_JSON, 'r', encoding='utf-8') as f:
        dados = json.load(f)
    turmas = dados["turmas"]

    img = Image(str(CAMINHO_IMAGEM))
    criar_aba_em_branco(wb, "SEC", img)

    for disciplina in DISCIPLINAS:
        criar_aba_disciplina(wb, disciplina, str(CAMINHO_IMAGEM), turmas)

    abas_adicionais = ["INDIVIDUAL", "BOLETIM", "BOL", "RESULTADO", "FREQUÊNCIA"]
    for aba in abas_adicionais:
        criar_aba_em_branco(wb, aba, img)

    caminho_completo = os.path.join(CAMINHO_PADRAO, NOME_ARQUIVO_PADRAO)
    wb.save(caminho_completo)
    logger.info(f"Planilha salva em: {caminho_completo}")
    return caminho_completo