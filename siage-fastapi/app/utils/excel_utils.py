from openpyxl.utils import get_column_letter
from app.core.config import COLUNAS  # Importa de config

def configurar_largura_colunas(ws, colunas_largura):
    """
    Define a largura das colunas especificadas.
    :param ws: A worksheet (aba) onde as colunas serão configuradas.
    :param colunas_largura: Um dicionário onde a chave é o nome da coluna e o valor é a largura em cm.
    """
    for coluna_nome, largura_cm in colunas_largura.items():
        # Encontra o índice da coluna com base no nome
        coluna_idx = COLUNAS.index(coluna_nome) + 1  # +1 porque as colunas começam em 1 no Excel
        # Converte o índice para a letra da coluna (A, B, C, etc.)
        coluna_letra = get_column_letter(coluna_idx)
        # Converte a largura de cm para unidades do Excel (1 cm ≈ 3.78 unidades)
        largura_unidades = largura_cm * 3.78
        # Aplica a largura à coluna
        ws.column_dimensions[coluna_letra].width = largura_unidades