import os
from datetime import datetime
from pathlib import Path
from zoneinfo import ZoneInfo

import click
import psycopg
from psycopg.rows import dict_row
from flask import current_app, g

BASE_DIR = Path(__file__).resolve().parent.parent
SCHEMA_PATH = BASE_DIR / "schema.sql"

FUSO_HORARIO = "America/Sao_Paulo"


def hoje_brasil():
    """Data de 'hoje' no fuso do Brasil, independente do fuso do servidor.

    O Render roda em UTC. Sem isso, depois das 21h (horario de Brasilia)
    ja e' meia-noite em UTC, e date.today() calculado no servidor passa
    a apontar para o dia seguinte -- enquanto o banco (com a sessao em
    'America/Sao_Paulo') ainda registra o dia de hoje. O caixa aberto
    deixa de ser encontrado, dando a impressao de que "nao abre"."""
    return datetime.now(ZoneInfo(FUSO_HORARIO)).date()


class Conexao:
    """Envelope fino sobre a conexao psycopg: aceita '?' como marcador de
    parametro (igual sqlite3) e devolve linhas como dict (igual sqlite3.Row),
    pra nao precisar reescrever as rotas que ja usam esse estilo."""

    def __init__(self, conn):
        self._conn = conn

    def execute(self, sql, params=()):
        cur = self._conn.cursor(row_factory=dict_row)
        cur.execute(sql.replace("?", "%s"), params)
        return cur

    def executescript(self, sql):
        with self._conn.cursor() as cur:
            cur.execute(sql)

    def commit(self):
        self._conn.commit()

    def close(self):
        self._conn.close()


def get_db():
    if "db" not in g:
        conn = psycopg.connect(current_app.config["DATABASE_URL"], autocommit=False)
        conn.execute(f"SET TIME ZONE '{FUSO_HORARIO}'")
        g.db = Conexao(conn)
    return g.db


def close_db(e=None):
    db = g.pop("db", None)
    if db is not None:
        db.close()


def init_db():
    db = get_db()
    with open(SCHEMA_PATH, "r", encoding="utf-8") as f:
        db.executescript(f.read())
    db.commit()


@click.command("init-db")
def init_db_command():
    """Cria as tabelas do banco (nao apaga dados existentes)."""
    init_db()
    click.echo("Banco de dados inicializado.")


def init_app(app):
    app.teardown_appcontext(close_db)
    app.cli.add_command(init_db_command)
    with app.app_context():
        init_db()
