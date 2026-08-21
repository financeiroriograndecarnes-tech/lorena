from pathlib import Path

from flask import Flask, redirect, url_for

from . import db as db_module

BASE_DIR = Path(__file__).resolve().parent.parent


def create_app(test_config=None):
    app = Flask(__name__, instance_relative_config=False)
    app.config.from_mapping(
        SECRET_KEY="dev",
        DATABASE=str(BASE_DIR / "data" / "sistema.db"),
    )
    if test_config:
        app.config.update(test_config)

    (BASE_DIR / "data").mkdir(exist_ok=True)

    db_module.init_app(app)

    from .routes import produtos, clientes, caixa, pdv

    app.register_blueprint(produtos.bp)
    app.register_blueprint(clientes.bp)
    app.register_blueprint(caixa.bp)
    app.register_blueprint(pdv.bp)

    @app.route("/")
    def index():
        return redirect(url_for("pdv.tela"))

    @app.context_processor
    def inject_globals():
        return {"app_nome": "Sistema Lorena"}

    return app
