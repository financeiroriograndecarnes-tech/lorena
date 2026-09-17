from flask import Blueprint, flash, redirect, render_template, request, url_for

from ..db import get_db, hoje_brasil, parse_num
from .caixa import registrar_movimento

bp = Blueprint("pagar", __name__, url_prefix="/pagar")


@bp.route("/")
def lista():
    db = get_db()
    hoje = hoje_brasil().isoformat()
    contas = db.execute(
        "SELECT * FROM contas_pagar ORDER BY (status = 'Pago'), vencimento"
    ).fetchall()
    total_aberto = sum(c["valor"] for c in contas if c["status"] != "Pago")
    total_vencido = sum(
        c["valor"] for c in contas if c["status"] != "Pago" and c["vencimento"] < hoje
    )
    return render_template(
        "pagar/lista.html", active="pagar", contas=contas, hoje=hoje,
        total_aberto=total_aberto, total_vencido=total_vencido,
    )


@bp.route("/novo", methods=["POST"])
def novo():
    db = get_db()
    fornecedor = request.form.get("fornecedor", "").strip()
    descricao = request.form.get("descricao", "").strip()
    vencimento = request.form.get("vencimento", "")
    valor = parse_num(request.form.get("valor"))
    if not fornecedor or not vencimento or valor <= 0:
        flash("Preencha fornecedor, vencimento e valor.", "erro")
        return redirect(url_for("pagar.lista"))
    db.execute(
        """INSERT INTO contas_pagar (fornecedor, descricao, vencimento, valor, status)
           VALUES (?,?,?,?,'Em Aberto')""",
        (fornecedor, descricao, vencimento, valor),
    )
    db.commit()
    flash("Conta a pagar lancada.", "sucesso")
    return redirect(url_for("pagar.lista"))


@bp.route("/<int:conta_id>/pagar", methods=["POST"])
def pagar(conta_id):
    db = get_db()
    conta = db.execute("SELECT * FROM contas_pagar WHERE id = ?", (conta_id,)).fetchone()
    if conta is None:
        flash("Conta nao encontrada.", "erro")
        return redirect(url_for("pagar.lista"))
    if conta["status"] == "Pago":
        flash("Conta ja esta paga.", "aviso")
        return redirect(url_for("pagar.lista"))

    forma = request.form.get("forma", "DINHEIRO")
    db.execute(
        "UPDATE contas_pagar SET status = 'Pago', data_pagamento = CURRENT_DATE::text WHERE id = ?",
        (conta_id,),
    )
    registrar_movimento(db, "Pagamento", -conta["valor"], forma, f"Conta #{conta_id} | {conta['fornecedor']}")
    db.commit()
    flash("Conta paga e lancada no caixa.", "sucesso")
    return redirect(url_for("pagar.lista"))
