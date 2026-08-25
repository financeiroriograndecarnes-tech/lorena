import os
from pathlib import Path

from flask import Flask, redirect, url_for

from . import db as db_module

BASE_DIR = Path(__file__).resolve().parent.parent


def _carregar_env_local():
    """Le webapp/.env (se existir) e joga as variaveis em os.environ, sem
    precisar de dependencia extra (python-dotenv). So preenche o que ainda
    nao estiver definido -- variaveis de ambiente reais (Render, etc.)
    sempre tem prioridade."""
    caminho = BASE_DIR / ".env"
    if not caminho.exists():
        return
    for linha in caminho.read_text(encoding="utf-8").splitlines():
        linha = linha.strip()
        if not linha or linha.startswith("#") or "=" not in linha:
            continue
        chave, valor = linha.split("=", 1)
        os.environ.setdefault(chave.strip(), valor.strip().strip('"').strip("'"))


def create_app(test_config=None):
    _carregar_env_local()

    app = Flask(__name__, instance_relative_config=False)
    app.config.from_mapping(
        SECRET_KEY=os.environ.get("SECRET_KEY", "dev"),
        DATABASE_URL=os.environ.get("DATABASE_URL", ""),
    )
    if test_config:
        app.config.update(test_config)

    if not app.config["DATABASE_URL"]:
        raise RuntimeError(
            "DATABASE_URL nao configurada. Crie webapp/.env com "
            "DATABASE_URL=postgresql://... (veja webapp/.env.example) "
            "ou defina a variavel de ambiente."
        )

    db_module.init_app(app)

    from .routes import (
        produtos, clientes, caixa, pdv, config, receber, pagar, dashboard,
        impressao, relatorios,
    )

    app.register_blueprint(produtos.bp)
    app.register_blueprint(clientes.bp)
    app.register_blueprint(caixa.bp)
    app.register_blueprint(pdv.bp)
    app.register_blueprint(config.bp)
    app.register_blueprint(receber.bp)
    app.register_blueprint(pagar.bp)
    app.register_blueprint(dashboard.bp)
    app.register_blueprint(impressao.bp)
    app.register_blueprint(relatorios.bp)

    @app.route("/")
    def index():
        return redirect(url_for("pdv.tela"))

    @app.context_processor
    def inject_globals():
        db = db_module.get_db()
        nome_empresa = db.execute(
            "SELECT valor FROM config WHERE chave = 'NOME_EMPRESA'"
        ).fetchone()
        nome_empresa = (nome_empresa["valor"] if nome_empresa else "") or ""
        return {
            "app_nome": nome_empresa or "Sistema",
            "nome_empresa": nome_empresa,
        }

    return app
