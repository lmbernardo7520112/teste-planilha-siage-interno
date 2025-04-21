import uvicorn
from app_init import app

# Importe e registre as rotas
from routes import router
app.include_router(router)

if __name__ == "__main__":
    uvicorn.run(app, host="0.0.0.0", port=8000)