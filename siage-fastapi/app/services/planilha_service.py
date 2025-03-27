import os
import json
import logging
import random
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

def criar_aba_boletim(wb, turmas):
    # Criar uma única aba "BOLETIM"
    ws = wb.create_sheet(title="BOLETIM")
    ws.sheet_properties.tabColor = COR_ABA

    # Adicionar o logotipo e o título
    img = Image(str(CAMINHO_IMAGEM))
    ws.merge_cells('A1:AZ1')  # Ajustar o merge para cobrir todas as colunas necessárias
    ws.row_dimensions[1].height = img.height * 0.75
    ws.add_image(img, 'A1')
    cell = ws['A1']
    cell.value = "COMPOSITOR LUIS RAMALHO"
    cell.font = Font(name='Arial', size=26, bold=True)
    cell.alignment = ALINHAMENTO_CENTRALIZADO

    linha_atual = 2
    border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    # Definir os cabeçalhos uma vez para referência de colunas
    headers = ["Nº", "ALUNO"]
    for disciplina in DISCIPLINAS:
        headers.extend([f"{disciplina} B1", f"{disciplina} B2", f"{disciplina} B3", f"{disciplina} B4", f"{disciplina} NF", f"{disciplina} MG"])
    
    # Definir larguras das colunas (aplicadas uma vez para toda a aba)
    ws.column_dimensions['A'].width = 5  # Nº
    ws.column_dimensions['B'].width = 20  # ALUNO
    for col in range(3, len(headers) + 1):
        ws.column_dimensions[get_column_letter(col)].width = 5  # Colunas de notas

    for turma in turmas:
        # Adicionar o nome da turma
        ws.merge_cells(f'A{linha_atual}:AZ{linha_atual}')
        ws[f'A{linha_atual}'] = f"{turma['nome_turma']} - BOLETIM"
        ws[f'A{linha_atual}'].font = FONTE_TITULO_TURMA
        ws[f'A{linha_atual}'].alignment = ALINHAMENTO_CENTRALIZADO
        ws.row_dimensions[linha_atual].height = 30
        linha_atual += 1

        # Adicionar a linha com os nomes das disciplinas (acima do cabeçalho)
        col_idx = 1
        # Preencher as primeiras duas colunas (Nº e ALUNO) como vazias
        for i in range(1, 3):
            cell = ws[f'{get_column_letter(col_idx)}{linha_atual}']
            cell.border = border
            col_idx += 1

        # Para cada disciplina, mesclar 6 colunas e adicionar o nome da disciplina
        for disciplina in DISCIPLINAS:
            # Cada disciplina ocupa 6 colunas (B1, B2, B3, B4, NF, MG)
            inicio_col = get_column_letter(col_idx)
            fim_col = get_column_letter(col_idx + 5)  # 6 colunas no total
            ws.merge_cells(f'{inicio_col}{linha_atual}:{fim_col}{linha_atual}')
            cell = ws[f'{inicio_col}{linha_atual}']
            cell.value = disciplina
            cell.font = Font(bold=True, size=6)  # Ajuste do tamanho da fonte para 6
            cell.alignment = ALINHAMENTO_CENTRALIZADO
            cell.fill = FILL_BIMESTRES  # Usar o mesmo preenchimento das notas para consistência
            cell.border = border
            col_idx += 6

        linha_atual += 1

        # Adicionar os cabeçalhos para a turma (Nº, ALUNO, BIO B1, etc.)
        col_idx = 1
        for header in headers:
            cell = ws[f'{get_column_letter(col_idx)}{linha_atual}']
            cell.value = header
            cell.border = border
            cell.font = Font(bold=True, size=6)  # Ajuste do tamanho da fonte para 6
            cell.alignment = ALINHAMENTO_CENTRALIZADO
            if header == "ALUNO":
                cell.fill = FILL_NOME_ALUNO
            elif "B1" in header or "B2" in header or "B3" in header or "B4" in header:
                cell.fill = FILL_BIMESTRES
            elif "NF" in header or "MG" in header:
                cell.fill = FILL_NOTA_FINAL
            col_idx += 1

        linha_inicio_dados = linha_atual + 1

        # Preencher os dados dos alunos
        for aluno in turma["alunos"]:
            linha_dados = linha_atual + int(aluno["numero"])
            ws[f'A{linha_dados}'] = int(aluno["numero"])
            ws[f'A{linha_dados}'].alignment = ALINHAMENTO_CENTRALIZADO
            ws[f'B{linha_dados}'] = aluno["nome"]
            ws[f'B{linha_dados}'].alignment = ALINHAMENTO_CENTRALIZADO

            # Buscar notas de cada disciplina pelo nome do aluno
            col_idx = 3  # Começar após Nº e ALUNO
            for disciplina in DISCIPLINAS:
                # Acessar a aba da disciplina
                ws_disciplina = wb[disciplina]
                # Determinar o intervalo de linhas para a turma atual
                idx_turma = turmas.index(turma)
                linha_base = 4 + (idx_turma * (36 + 15))  # 36 linhas de dados + 15 de espaço
                linha_fim = linha_base + 35  # Até 35 alunos por turma

                # Procurar o aluno pelo nome na coluna B
                linha_ref = None
                for row in range(linha_base, linha_fim + 1):
                    cell_nome = ws_disciplina[f'B{row}'].value
                    if cell_nome == aluno["nome"]:
                        linha_ref = row
                        break

                if linha_ref is None:
                    logger.warning(f"Aluno {aluno['nome']} não encontrado na aba {disciplina} para a turma {turma['nome_turma']}")
                    # Preencher com 0.00 se o aluno não for encontrado
                    for _ in range(6):  # 4 bimestres + NF + MG
                        cell = ws[f'{get_column_letter(col_idx)}{linha_dados}']
                        cell.value = 0.00
                        cell.number_format = '0.00'
                        cell.alignment = ALINHAMENTO_CENTRALIZADO
                        col_idx += 1
                    continue

                # Copiar as notas (B1, B2, B3, B4, NF, MG)
                for bimestre in ['C', 'D', 'E', 'F']:  # B1, B2, B3, B4
                    cell = ws[f'{get_column_letter(col_idx)}{linha_dados}']
                    cell.value = f"='{disciplina}'!{bimestre}{linha_ref}"
                    cell.number_format = '0.00'
                    cell.alignment = ALINHAMENTO_CENTRALIZADO
                    col_idx += 1
                # NF
                cell = ws[f'{get_column_letter(col_idx)}{linha_dados}']
                cell.value = f"='{disciplina}'!G{linha_ref}"
                cell.number_format = '0.00'
                cell.alignment = ALINHAMENTO_CENTRALIZADO
                col_idx += 1
                # MG
                cell = ws[f'{get_column_letter(col_idx)}{linha_dados}']
                cell.value = f"='{disciplina}'!H{linha_ref}"
                cell.number_format = '0.00'
                cell.alignment = ALINHAMENTO_CENTRALIZADO
                col_idx += 1

        # Aplicar bordas às linhas de dados
        for row in range(linha_inicio_dados, linha_inicio_dados + len(turma["alunos"])):
            for col in range(1, len(headers) + 1):
                cell = ws[f'{get_column_letter(col)}{row}']
                cell.border = border

        # Adicionar espaço entre turmas
        linha_atual += len(turma["alunos"]) + 3  # 2 linhas de espaço após os dados

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

            if ws[f'B{linha_dados}'].value:
                ws[f'C{linha_dados}'] = random.uniform(1, 10)
                ws[f'D{linha_dados}'] = random.uniform(1, 10)
                ws[f'E{linha_dados}'] = random.uniform(1, 10)
                ws[f'F{linha_dados}'] = random.uniform(1, 10)
                ws[f'C{linha_dados}'].number_format = '0.00'
                ws[f'D{linha_dados}'].number_format = '0.00'
                ws[f'E{linha_dados}'].number_format = '0.00'
                ws[f'F{linha_dados}'].number_format = '0.00'
        
        for row in range(linha_inicio_dados, linha_inicio_dados + 35):
            ws[f'G{row}'] = f'=AVERAGE(C{row}:F{row})'
            ws[f'G{row}'].alignment = ALINHAMENTO_CENTRALIZADO
            ws[f'G{row}'].number_format = '0.00'

            ws[f'H{row}'] = f'=SUM(C{row}:F{row})/4'
            ws[f'H{row}'].alignment = ALINHAMENTO_CENTRALIZADO
            ws[f'H{row}'].number_format = '0.00'

            ws[f'I{row}'] = f'=IF(H{row}<7, (0.6*H{row}) + (0.4*G{row}), "-")'
            ws[f'I{row}'].alignment = ALINHAMENTO_CENTRALIZADO
            ws[f'I{row}'].number_format = '0.00'

            ws[f'J{row}'] = f'=IF(H{row}<2.5, "REPROVADO", IF(H{row}<7, "FINAL", "APROVADO"))'
            ws[f'J{row}'].alignment = ALINHAMENTO_CENTRALIZADO

            ws[f'K{row}'] = f'=IF(H{row}<7, (12.5 - (1.5*H{row})), "-")'
            ws[f'K{row}'].alignment = ALINHAMENTO_CENTRALIZADO
            ws[f'K{row}'].number_format = '0.00'

            ws[f'L{row}'] = f'=IF(G{row}>=K{row}, "AF", "-")'
            ws[f'L{row}'].alignment = ALINHAMENTO_CENTRALIZADO
            ws[f'L{row}'].number_format = '0.00'
        
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

    criar_aba_boletim(wb, turmas)

    abas_adicionais = ["INDIVIDUAL", "BOL", "RESULTADO", "FREQUÊNCIA"]
    for aba in abas_adicionais:
        criar_aba_em_branco(wb, aba)

    caminho_completo = os.path.join(CAMINHO_PADRAO, NOME_ARQUIVO_PADRAO)
    wb.save(caminho_completo)
    logger.info(f"Planilha salva em: {caminho_completo}")
    return caminho_completo