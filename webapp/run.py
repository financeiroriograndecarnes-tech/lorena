from app import create_app

app = create_app()

if __name__ == "__main__":
    # host 0.0.0.0 para o tablet acessar pelo Wi-Fi da loja durante os testes
    app.run(host="0.0.0.0", port=5000, debug=True)
