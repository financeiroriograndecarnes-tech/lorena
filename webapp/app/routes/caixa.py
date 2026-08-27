from datetime import datetime, date

from flask import Blueprint, flash, redirect, render_template, request, url_for

from ..db import get_db, hoje_brasil

bp = Blueprint("caixa", __name__, url_prefix="/caixa")

USUARIO_PADRAO = "OPERADOR"


def registrar_movimento(db, tipo, valor, forma="DINHEIRO", observacao="", usuario=None):
    db.execute(
        """INSERT INTO caixa (data_hora, tipo, valor, forma_pagamento, observacao, usuario)
           VALUES (to_char(now(), 'YYYY-MM-DD HH24:MI:SS'), ?, ?, ?, ?, ?)""",
        (tipo, valor, forma, observacao, usuario or USUARIO_PADRAO),
    )


def caixa_aberto(db, dt=None):
    """Estado = tipo do ULTIMO evento (Abertura/Fechamento) do dia, por horario."""
    dt = dt or hoje_brasil()
    rows = db.execute(
        """SELECT tipo, data_hora FROM caixa
           WHERE tipo IN ('Abertura','Fechamento') AND data_hora::date = ?::date
           ORDER BY data_hora ASC""",
        (dt.isoformat(),),
    ).fetchall()
    if not rows:
        return False
    return rows[-1]["tipo"] == "Abertura"


def total_por_forma(db, forma, dt=None):
    dt = dt or hoje_brasil()
    row = db.execute(
        """SELECT COALESCE(SUM(valor), 0) AS total FROM caixa
           WHERE data_hora::date = ?::date AND forma_pagamento = ? AND tipo != 'Fechamento'""",
        (dt.isoformat(), forma),
    ).fetchone()
    return row["total"] or 0


def saldo_dinheiro(db, dt=None):
    return total_por_forma(db, "DINHEIRO", dt)


def total_geral_dia(db, dt=None):
    dt = dt or hoje_brasil()
    row = db.execute(
        """SELECT COALESCE(SUM(valor), 0) AS total FROM caixa
           WHERE data_hora::date = ?::date AND tipo != 'Fechamento'""",
        (dt.isoformat(),),
    ).fetchone()
    return row["total"] or 0


FORMAS = ["DINHEIRO", "CARTAO CREDITO", "CARTAO DEBITO", "PIX", "BOLETO", "CHEQUE", "A PRAZO", "VALE"]


@bp.route("/")
def tela():
    db = get_db()
    aberto = caixa_aberto(db)
    movimentos = db.execute(
        """SELECT * FROM caixa WHERE data_hora::date = CURRENT_DATE
           ORDER BY data_hora DESC"""
    ).fetchall()
    resumo = {forma: total_por_forma(db, forma) for forma in FORMAS}
    return render_template(
        "caixa/tela.html",
        active="caixa",
        aberto=aberto,
        movimentos=movimentos,
        resumo=resumo,
        total_dia=total_geral_dia(db),
        esperado_dinheiro=saldo_dinheiro(db),
    )


@bp.route("/abrir", methods=["POST"])
def abrir():
    db = get_db()
    if caixa_aberto(db):
        flash("O caixa de hoje ja esta aberto.", "aviso")
        return redirect(url_for("caixa.tela"))
    valor = float(request.form.get("valor") or 0)
    if valor < 0:
        flash("Fundo de troco invalido.", "erro")
        return redirect(url_for("caixa.tela"))
    registrar_movimento(db, "Abertura", valor, "DINHEIRO", "Fundo de troco")
    db.commit()
    flash("Caixa aberto com sucesso.", "sucesso")
    return redirect(url_for("caixa.tela"))


@bp.route("/sangria", methods=["POST"])
def sangria():
    db = get_db()
    if not caixa_aberto(db):
        flash("Abra o caixa antes de fazer sangria.", "erro")
        return redirect(url_for("caixa.tela"))
    valor = float(request.form.get("valor") or 0)
    motivo = request.form.get("motivo", "").strip()
    if valor <= 0:
        flash("Informe um valor valido.", "erro")
        return redirect(url_for("caixa.tela"))
    if not motivo:
        flash("Informe o motivo da sangria.", "erro")
        return redirect(url_for("caixa.tela"))
    registrar_movimento(db, "Sangria", -valor, "DINHEIRO", motivo)
    db.commit()
    flash("Sangria registrada.", "sucesso")
    return redirect(url_for("caixa.tela"))


@bp.route("/suprimento", methods=["POST"])
def suprimento():
    db = get_db()
    if not caixa_aberto(db):
        flash("Abra o caixa antes de lancar suprimento.", "erro")
        return redirect(url_for("caixa.tela"))
    valor = float(request.form.get("valor") or 0)
    motivo = request.form.get("motivo", "").strip()
    if valor <= 0:
        flash("Informe um valor valido.", "erro")
        return redirect(url_for("caixa.tela"))
    if not motivo:
        flash("Informe o motivo do suprimento.", "erro")
        return redirect(url_for("caixa.tela"))
    registrar_movimento(db, "Suprimento", valor, "DINHEIRO", motivo)
    db.commit()
    flash("Suprimento registrado.", "sucesso")
    return redirect(url_for("caixa.tela"))


@bp.route("/fechar", methods=["POST"])
def fechar():
    db = get_db()
    if not caixa_aberto(db):
        flash("Nao existe caixa aberto hoje.", "erro")
        return redirect(url_for("caixa.tela"))
    apurado = float(request.form.get("apurado") or 0)
    esperado = saldo_dinheiro(db)
    diferenca = apurado - esperado
    if abs(diferenca) > 0.009:
        obs = ("QUEBRA (falta) " if diferenca < 0 else "SOBRA ") + f"R$ {abs(diferenca):.2f}"
    else:
        obs = "Caixa conferido sem diferenca"
    registrar_movimento(db, "Fechamento", apurado, "DINHEIRO", f"{obs} | Esperado R$ {esperado:.2f}")
    db.commit()
    if abs(diferenca) < 0.01:
        flash("Caixa fechado sem diferenca.", "sucesso")
    elif diferenca < 0:
        flash(f"Caixa fechado com QUEBRA de R$ {abs(diferenca):.2f}.", "aviso")
    else:
        flash(f"Caixa fechado com SOBRA de R$ {diferenca:.2f}.", "aviso")
    return redirect(url_for("caixa.tela"))
