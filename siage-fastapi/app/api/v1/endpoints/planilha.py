import logging
from fastapi import APIRouter, HTTPException
from fastapi.responses import FileResponse
from app.services.planilha_service import criar_planilha
from app.core.config import NOME_ARQUIVO_PADRAO

router = APIRouter()
logger = logging.getLogger(__name__)

@router.get("/gerar-planilha")
async def gerar_planilha():
    logger.info("Endpoint /gerar-planilha chamado")
    try:
        caminho_planilha = criar_planilha()
        logger.info(f"Retornando arquivo: {caminho_planilha}")
        return FileResponse(
            caminho_planilha,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            filename=NOME_ARQUIVO_PADRAO
        )
    except Exception as e:
        logger.error(f"Erro ao gerar planilha: {str(e)}")
        raise HTTPException(status_code=500, detail=str(e))