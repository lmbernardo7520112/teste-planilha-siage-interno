from app.services.planilha_service import criar_planilha

if __name__ == "__main__":
    try:
        caminho_planilha = criar_planilha()
        print(f"Planilha gerada com sucesso em: {caminho_planilha}")
        print(f"Caminho da imagem: {CAMINHO_IMAGEM}")
    except Exception as e:
        print(f"Erro ao gerar a planilha: {e}")