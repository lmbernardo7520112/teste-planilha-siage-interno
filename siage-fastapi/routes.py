# routes.py
import os
from fastapi import APIRouter, Request, UploadFile, File
from fastapi.responses import HTMLResponse, RedirectResponse, FileResponse
import logging
from app_init import templates

router = APIRouter()

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

@router.get("/", response_class=HTMLResponse)
async def home(request: Request):
    logger.info("Rota / acessada")
    return templates.TemplateResponse("home.html", {"request": request})

@router.get("/upload", response_class=HTMLResponse)
async def upload(request: Request):
    return templates.TemplateResponse("upload.html", {"request": request})

@router.post("/upload/json")
async def upload_json(file: UploadFile = File(...)):
    try:
        # Ler conteúdo do arquivo
        content = await file.read()
        
        # Salvar o arquivo recebido
        with open("turmas_alunos.json", "wb") as f:
            f.write(content)
        
        # Redirecionar para a página inicial com mensagem de sucesso
        return RedirectResponse(
            url="/?message=Arquivo processado com sucesso. Clique em Download Grades para baixar a planilha.&category=success", 
            status_code=303
        )
    except Exception as e:
        logger.error(f"Erro ao processar arquivo: {str(e)}")
        # Redirecionar para a página de upload com mensagem de erro
        return RedirectResponse(
            url="/upload?message=Erro ao processar arquivo: {}&category=error".format(str(e)), 
            status_code=303
        )

@router.get("/download/grades")
async def download_grades():
    try:
        file_path = "planilha_notas_complexa.xlsx"
        if not os.path.exists(file_path):
            return RedirectResponse(
                url="/?message=Planilha não encontrada. Faça o upload primeiro.&category=error", 
                status_code=303
            )
        
        return FileResponse(
            path=file_path,
            filename="planilha_notas.xlsx",
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        logger.error(f"Erro ao baixar arquivo: {str(e)}")
        return RedirectResponse(
            url="/?message=Erro ao baixar planilha: {}&category=error".format(str(e)), 
            status_code=303
        )