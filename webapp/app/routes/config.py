from flask import Blueprint, flash, redirect, render_template, request, url_for

from ..db import get_db

bp = Blueprint("config", __name__, url_prefix="/configuracoes")

CAMPOS = [
    ("NOME_EMPRESA", "Nome da empresa"),
    ("CNPJ", "CNPJ / CPF"),
    ("ENDERECO", "Endereco"),
    ("TELEFONE", "Telefone"),
    ("PERC_ATACADO", "% desconto no atacado (sobre o varejo)"),
    ("PERC_CARTAO", "% acrescimo no cartao (sobre o varejo)"),
    ("USUARIO_PADRAO", "Usuario padrao sugerido no PDV"),
]


def get_config(db):
    rows = db.execute("SELECT chave, valor FROM config").fetchall()
    return {r["chave"]: r["valor"] for r in rows}


@bp.route("/", methods=["GET", "POST"])
def tela():
    db = get_db()
    if request.method == "POST":
        for chave, _ in CAMPOS:
            valor = request.form.get(chave, "").strip()
            db.execute(
                "INSERT INTO config (chave, valor) VALUES (?, ?) "
                "ON CONFLICT(chave) DO UPDATE SET valor = excluded.valor",
                (chave, valor),
            )
        db.commit()
        flash("Dados da empresa salvos com sucesso.", "sucesso")
        return redirect(url_for("config.tela"))

    valores = get_config(db)
    return render_template("config/tela.html", active="config", campos=CAMPOS, valores=valores)
