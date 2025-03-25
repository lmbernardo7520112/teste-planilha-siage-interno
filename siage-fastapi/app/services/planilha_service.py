import os
import json
import logging
import random  # Necessário para gerar notas aleatórias
from pathlib import Path
from openpyxl.styles import Font, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl import Workbook
from openpyxl.drawing.image import Image
from app.utils.excel_utils import configurar_largura_colunas, criar_dashboard_turma, criar_dashboard_sec_turma, criar_dashboard_sec_geral, criar_dashboard_sec_aprovacao
from app.core.config import (
    COLUNAS, COLUNAS_SEC, DISCIPLINAS, CAMINHO_IMAGEM, CAMINHO_PADRAO, NOME_ARQUIVO_PADRAO, LARGURAS_COLUNAS, LARGURAS_COLUNAS_ABAS_DISC,
    COR_ABA, FILL_NOME_ALUNO, FILL_BIMESTRES, FILL_NOTA_FINAL, FILL_SITUACAO, FONTE_TITULO_TURMA, ALINHAMENTO_CENTRALIZADO
)

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

CAMINHO_JSON = Path(__file__).parent.parent.parent / "turmas_alunos.json"

def criar_aba_em_branco(wb, titulo):
    ws = wb.create_sheet(title=titulo)
    ws.sheet_properties.tabColor = COR_ABA
    img = Image(str(CAMINHO_IMAGEM))
    ws.merge_cells('A1:J1')
    ws.row_dimensions[1].height = img.height * 0.75
    ws.add_image(img, 'A1')
    cell = ws['A1']
    cell.value = "COMPOSITOR LUIS RAMALHO"
    cell.font = Font(name='Arial', size=26, bold=True)
    cell.alignment = ALINHAMENTO_CENTRALIZADO
    return ws

def criar_aba_sec(wb, turmas):
    ws = wb.create_sheet(title="SEC")
    ws.sheet_properties.tabColor = COR_ABA
    img = Image(str(CAMINHO_IMAGEM))
    border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    linha_atual = 1
    linhas_inicio_tabelas = []

    ws.merge_cells('A1:F1')
    ws.row_dimensions[1].height = img.height * 0.75
    ws.add_image(img, 'A1')
    cell = ws['A1']
    cell.value = "COMPOSITOR LUIS RAMALHO"
    cell.font = Font(name='Arial', size=26, bold=True)
    cell.alignment = ALINHAMENTO_CENTRALIZADO
    linha_atual += 1

    for turma in turmas:
        ws.merge_cells(f'A{linha_atual}:F{linha_atual}')
        ws[f'A{linha_atual}'] = turma["nome_turma"]
        ws[f'A{linha_atual}'].font = FONTE_TITULO_TURMA
        ws[f'A{linha_atual}'].alignment = ALINHAMENTO_CENTRALIZADO
        ws.row_dimensions[linha_atual].height = 30
        linha_atual += 1
        linhas_inicio_tabelas.append(linha_atual)

        for col_idx, col_nome in enumerate(COLUNAS_SEC, 1):
            cell = ws[f'{get_column_letter(col_idx)}{linha_atual}']
            cell.value = col_nome
            cell.border = border
            cell.font = Font(bold=True)
            cell.alignment = ALINHAMENTO_CENTRALIZADO
            if col_nome == "Nome do Aluno":
                cell.fill = FILL_NOME_ALUNO
            elif col_nome == "SITUAÇÃO DO ALUNO":
                cell.fill = FILL_SITUACAO
            elif col_nome in ["ATIVO", "TRANSFERIDO", "DESISTENTE"]:
                cell.fill = FILL_BIMESTRES

        configurar_largura_colunas(ws, LARGURAS_COLUNAS, COLUNAS_SEC)
        linha_inicio_dados = linha_atual + 1

        for aluno in turma["alunos"]:
            row = linha_atual + int(aluno["numero"])
            ws[f'A{row}'] = int(aluno["numero"])
            ws[f'A{row}'].alignment = ALINHAMENTO_CENTRALIZADO
            ws[f'B{row}'] = aluno["nome"]
            ws[f'B{row}'].alignment = ALINHAMENTO_CENTRALIZADO
            ws[f'C{row}'] = aluno.get("ativo", True)
            ws[f'C{row}'].alignment = ALINHAMENTO_CENTRALIZADO
            ws[f'D{row}'] = aluno.get("transferido", False)
            ws[f'D{row}'].alignment = ALINHAMENTO_CENTRALIZADO
            ws[f'E{row}'] = aluno.get("desistente", False)
            ws[f'E{row}'].alignment = ALINHAMENTO_CENTRALIZADO
            ws[f'F{row}'] = f'=IF(E{row}, "DESISTENTE", IF(D{row}, "TRANSFERIDO", IF(C{row}, "ATIVO", "INDEFINIDO")))'
            ws[f'F{row}'].alignment = ALINHAMENTO_CENTRALIZADO

        for row in range(linha_atual, linha_atual + len(turma["alunos"]) + 1):
            for col in range(1, 7):
                cell = ws[f'{get_column_letter(col)}{row}']
                cell.border = border

        criar_dashboard_sec_turma(ws, linha_atual, linha_inicio_dados, len(turma["alunos"]))
        linha_atual += len(turma["alunos"]) + 6

    criar_dashboard_sec_geral(ws, linhas_inicio_tabelas, [len(turma["alunos"]) for turma in turmas])
    criar_dashboard_sec_aprovacao(ws, turmas, linhas_inicio_tabelas)

    return ws

def criar_aba_disciplina(wb, titulo, turmas):
    ws = wb.create_sheet(title=titulo)
    img = Image(str(CAMINHO_IMAGEM))
    ws.merge_cells('A1:J1')
    ws.row_dimensions[1].height = img.height * 0.75
    ws.add_image(img, 'A1')
    cell = ws['A1']
    cell.value = "COMPOSITOR LUIS RAMALHO"
    cell.font = Font(name='Arial', size=26, bold=True)
    cell.alignment = ALINHAMENTO_CENTRALIZADO
    ws.sheet_properties.tabColor = COR_ABA

    border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    linha_atual = 2

    for turma in turmas:
        ws.merge_cells(f'A{linha_atual}:L{linha_atual}')
        ws[f'A{linha_atual}'] = turma["nome_turma"]
        ws[f'A{linha_atual}'].font = FONTE_TITULO_TURMA
        ws[f'A{linha_atual}'].alignment = ALINHAMENTO_CENTRALIZADO
        ws.row_dimensions[linha_atual].height = 30
        linha_atual += 1
        
        for col_idx, col_nome in enumerate(COLUNAS, 1):
            cell = ws[f'{get_column_letter(col_idx)}{linha_atual}']
            cell.value = col_nome
            cell.border = border
            cell.font = Font(bold=True)
            cell.alignment = ALINHAMENTO_CENTRALIZADO
            if col_nome == "Nome do Aluno":
                cell.fill = FILL_NOME_ALUNO
            elif col_nome in ["1º BIM", "2º BIM", "3º BIM", "4º BIM"]:
                cell.fill = FILL_BIMESTRES
            elif col_nome == "NF":
                cell.fill = FILL_NOTA_FINAL
            elif col_nome == "SITUAÇÃO DO ALUNO":
                cell.fill = FILL_SITUACAO
        
        configurar_largura_colunas(ws, LARGURAS_COLUNAS_ABAS_DISC, COLUNAS)
        linha_inicio_dados = linha_atual + 1
        
        for aluno in turma["alunos"]:
            linha_dados = linha_atual + int(aluno["numero"])
            ws[f'A{linha_dados}'] = int(aluno["numero"])
            ws[f'A{linha_dados}'].alignment = ALINHAMENTO_CENTRALIZADO
            ws[f'B{linha_dados}'] = aluno["nome"]
            ws[f'B{linha_dados}'].alignment = ALINHAMENTO_CENTRALIZADO

            # Verificar se há aluno (coluna B preenchida) e atribuir notas aleatórias para todos os bimestres
            if ws[f'B{linha_dados}'].value:  # Se a célula não estiver vazia
                ws[f'C{linha_dados}'] = random.uniform(1, 10)  # 1º bimestre
                ws[f'D{linha_dados}'] = random.uniform(1, 10)  # 2º bimestre
                ws[f'E{linha_dados}'] = random.uniform(1, 10)  # 3º bimestre
                ws[f'F{linha_dados}'] = random.uniform(1, 10)  # 4º bimestre
                # Formatar as notas com 2 casas decimais
                ws[f'C{linha_dados}'].number_format = '0.00'
                ws[f'D{linha_dados}'].number_format = '0.00'
                ws[f'E{linha_dados}'].number_format = '0.00'
                ws[f'F{linha_dados}'].number_format = '0.00'
        
        for row in range(linha_inicio_dados, linha_inicio_dados + 35):
            ws[f'G{row}'] = f'=AVERAGE(C{row}:F{row})'
            ws[f'G{row}'].alignment = ALINHAMENTO_CENTRALIZADO
            ws[f'H{row}'] = f'=SUM(C{row}:F{row})/4'
            ws[f'H{row}'].alignment = ALINHAMENTO_CENTRALIZADO
            ws[f'I{row}'] = f'=IF(H{row}<7, (0.6*H{row}) + (0.4*G{row}), "-")'
            ws[f'I{row}'].alignment = ALINHAMENTO_CENTRALIZADO
            ws[f'J{row}'] = f'=IF(H{row}<2.5, "REPROVADO", IF(H{row}<7, "FINAL", "APROVADO"))'
            ws[f'J{row}'].alignment = ALINHAMENTO_CENTRALIZADO
            ws[f'K{row}'] = f'=IF(H{row}<7, (12.5 - (1.5*H{row})), "-")'
            ws[f'K{row}'].alignment = ALINHAMENTO_CENTRALIZADO
            ws[f'L{row}'] = f'=IF(G{row}>=K{row}, "AF", "-")'
            ws[f'L{row}'].alignment = ALINHAMENTO_CENTRALIZADO
        
        for row in range(linha_atual, linha_atual + 36):
            for col in range(1, 13):
                cell = ws[f'{get_column_letter(col)}{row}']
                cell.border = border
        
        criar_dashboard_turma(ws, linha_atual, linha_inicio_dados)
        linha_atual += 36 + 15

    return ws

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

    criar_aba_sec(wb, turmas)
    for disciplina in DISCIPLINAS:
        criar_aba_disciplina(wb, disciplina, turmas)

    abas_adicionais = ["INDIVIDUAL", "BOLETIM", "BOL", "RESULTADO", "FREQUÊNCIA"]
    for aba in abas_adicionais:
        criar_aba_em_branco(wb, aba)

    caminho_completo = os.path.join(CAMINHO_PADRAO, NOME_ARQUIVO_PADRAO)
    wb.save(caminho_completo)
    logger.info(f"Planilha salva em: {caminho_completo}")
    return caminho_completo