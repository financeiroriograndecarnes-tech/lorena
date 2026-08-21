from datetime import date

from flask import Blueprint, flash, redirect, render_template, request, url_for

from ..db import get_db
from .caixa import registrar_movimento
from .config import get_config

bp = Blueprint("receber", __name__, url_prefix="/receber")


def calcular_encargos(db, valor_saldo, vencimento_iso, referencia=None):
    """Juros + multa sobre o saldo, se vencido alem da tolerancia. Nao grava nada."""
    cfg = get_config(db)
    tolerancia = int(float(cfg.get("DIAS_TOLERANCIA", 0) or 0))
    juros_dia = float(cfg.get("JUROS_DIA", 0) or 0)
    multa_perc = float(cfg.get("MULTA_PERC", 0) or 0)

    ref = referencia or date.today()
    venc = date.fromisoformat(vencimento_iso)
    dias_atraso = max(0, (ref - venc).days)

    if dias_atraso <= tolerancia or valor_saldo <= 0:
        return 0.0, 0.0, dias_atraso

    multa = round(valor_saldo * multa_perc / 100, 2)
    juros = round(valor_saldo * (juros_dia / 100) * (dias_atraso - tolerancia), 2)
    return juros, multa, dias_atraso


@bp.route("/")
def lista():
    db = get_db()
    filtro_cliente = request.args.get("cliente", "").strip()
    so_vencidas = request.args.get("vencidas") == "1"

    query = """
        SELECT cr.*, c.nome AS cliente_nome, c.responsavel, c.turma
        FROM contas_receber cr
        LEFT JOIN clientes c ON c.id = cr.cliente_id
        WHERE 1=1
    """
    params = []
    if filtro_cliente:
        query += " AND (c.nome LIKE ? OR c.responsavel LIKE ? OR c.turma LIKE ?)"
        like = f"%{filtro_cliente}%"
        params += [like, like, like]
    if so_vencidas:
        query += " AND cr.status != 'Pago' AND cr.vencimento < date('now')"
    query += " ORDER BY cr.vencimento"

    parcelas = db.execute(query, params).fetchall()

    linhas = []
    total_aberto = 0
    total_vencido = 0
    hoje = date.today()
    for p in parcelas:
        saldo = p["valor_parcela"] - p["valor_pago"]
        juros, multa, dias_atraso = calcular_encargos(db, saldo, p["vencimento"])
        devido = round(saldo + juros + multa, 2) if p["status"] != "Pago" else 0
        venc = date.fromisoformat(p["vencimento"])
        if p["status"] == "Pago":
            status_exibicao = "Pago"
        elif venc < hoje:
            status_exibicao = "Vencido"
        elif p["valor_pago"] > 0:
            status_exibicao = "Parcial"
        else:
            status_exibicao = "Em Aberto"

        if p["status"] != "Pago":
            total_aberto += saldo
            if venc < hoje:
                total_vencido += saldo

        linhas.append({
            "row": p,
            "saldo": saldo,
            "juros": juros,
            "multa": multa,
            "devido": devido,
            "status_exibicao": status_exibicao,
            "dias_atraso": dias_atraso,
        })

    return render_template(
        "receber/lista.html",
        active="receber",
        linhas=linhas,
        filtro_cliente=filtro_cliente,
        so_vencidas=so_vencidas,
        total_aberto=total_aberto,
        total_vencido=total_vencido,
    )


@bp.route("/<int:parcela_id>/baixar", methods=["POST"])
def baixar(parcela_id):
    db = get_db()
    parcela = db.execute("SELECT * FROM contas_receber WHERE id = ?", (parcela_id,)).fetchone()
    if parcela is None:
        flash("Parcela nao encontrada.", "erro")
        return redirect(url_for("receber.lista"))
    if parcela["status"] == "Pago":
        flash("Parcela ja esta quitada.", "aviso")
        return redirect(url_for("receber.lista"))

    forma = request.form.get("forma", "DINHEIRO")
    saldo = parcela["valor_parcela"] - parcela["valor_pago"]
    juros, multa, _ = calcular_encargos(db, saldo, parcela["vencimento"])
    encargos = juros + multa

    modo = request.form.get("modo", "total")
    if modo == "total":
        valor_recebido = round(saldo + encargos, 2)
    else:
        valor_recebido = float(request.form.get("valor") or 0)
        if valor_recebido <= 0:
            flash("Informe um valor valido.", "erro")
            return redirect(url_for("receber.lista"))

    if valor_recebido >= encargos:
        aplicado_principal = valor_recebido - encargos
        encargo_cobrado = encargos
    else:
        aplicado_principal = 0
        encargo_cobrado = valor_recebido

    novo_pago = parcela["valor_pago"] + aplicado_principal
    novo_status = "Pago" if novo_pago >= parcela["valor_parcela"] - 0.009 else "Parcial"

    db.execute(
        """UPDATE contas_receber SET valor_pago = ?, juros_multa = juros_multa + ?,
           data_pagamento = date('now'), status = ? WHERE id = ?""",
        (novo_pago, encargo_cobrado, novo_status, parcela_id),
    )
    registrar_movimento(
        db, "Recebimento", valor_recebido, forma,
        f"Parcela #{parcela_id}" + (f" | {parcela['cliente_id']}" if parcela["cliente_id"] else ""),
    )
    db.commit()
    flash(f"Baixa de R$ {valor_recebido:.2f} registrada e lancada no caixa.", "sucesso")
    return redirect(url_for("receber.lista"))
