# main.py (na raiz do projeto)

from fastapi import FastAPI
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from routes import router
from app.database import init_db
import os

app = FastAPI()

# Diretórios
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
templates = Jinja2Templates(directory=os.path.join(BASE_DIR, "templates"))
app.mount("/static", StaticFiles(directory=os.path.join(BASE_DIR, "static")), name="static")

# Inicializa banco
init_db()

# Registra rotas
app.include_router(router)

# DEBUG: Mostra rotas registradas
print("Rotas registradas:")
for route in app.routes:
    print(route.path)