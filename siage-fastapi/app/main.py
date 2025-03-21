from fastapi import FastAPI
from app.api.v1.endpoints import planilha

app = FastAPI()

# Inclui os endpoints da API
app.include_router(planilha.router, prefix="/api/v1", tags=["planilha"])