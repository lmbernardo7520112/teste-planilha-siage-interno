from fastapi import APIRouter, HTTPException
from fastapi.responses import FileResponse
from app.services.planilha_service import criar_planilha

router = APIRouter()

@router.get("/gerar-planilha")
async def gerar_planilha():
    try:
        caminho_planilha = criar_planilha()
        return FileResponse(
            caminho_planilha,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            filename="planilha_notas_complexa.xlsx"
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))