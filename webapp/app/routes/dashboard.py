from datetime import date

from flask import Blueprint, render_template

from ..db import get_db

bp = Blueprint("dashboard", __name__, url_prefix="/painel")


def _vendas_periodo(db, inicio_sql):
    row = db.execute(
        f"""SELECT COALESCE(SUM(total_liquido), 0) AS total, COUNT(*) AS n
            FROM vendas WHERE status = 'Concluida' AND date(data_hora) >= date('now', ?)""",
        (inicio_sql,),
    ).fetchone()
    return row["total"], row["n"]


@bp.route("/")
def tela():
    db = get_db()
    hoje = date.today().isoformat()

    vendas_dia, n_dia = _vendas_periodo(db, "start of day")
    vendas_mes, n_mes = _vendas_periodo(db, "start of month")
    vendas_ano, n_ano = _vendas_periodo(db, "start of year")

    receber_aberto = db.execute(
        "SELECT COALESCE(SUM(valor_parcela - valor_pago),0) AS t FROM contas_receber WHERE status != 'Pago'"
    ).fetchone()["t"]
    receber_vencido = db.execute(
        "SELECT COALESCE(SUM(valor_parcela - valor_pago),0) AS t FROM contas_receber WHERE status != 'Pago' AND vencimento < ?",
        (hoje,),
    ).fetchone()["t"]
    pagar_aberto = db.execute(
        "SELECT COALESCE(SUM(valor),0) AS t FROM contas_pagar WHERE status != 'Pago'"
    ).fetchone()["t"]
    pagar_vencido = db.execute(
        "SELECT COALESCE(SUM(valor),0) AS t FROM contas_pagar WHERE status != 'Pago' AND vencimento < ?",
        (hoje,),
    ).fetchone()["t"]

    produtos_abaixo_minimo = db.execute(
        "SELECT COUNT(*) AS n FROM produtos WHERE ativo = 1 AND estoque_atual <= estoque_minimo"
    ).fetchone()["n"]
    orcamentos_abertos = db.execute(
        "SELECT COUNT(*) AS n FROM vendas WHERE status = 'Orcamento'"
    ).fetchone()["n"]

    ranking = db.execute(
        """SELECT descricao, SUM(total_item) AS total
           FROM itens_venda iv JOIN vendas v ON v.id = iv.venda_id
           WHERE v.status = 'Concluida' AND date(v.data_hora) >= date('now', 'start of month')
           GROUP BY iv.descricao ORDER BY total DESC LIMIT 10"""
    ).fetchall()

    return render_template(
        "dashboard/tela.html",
        active="dashboard",
        vendas_dia=vendas_dia, vendas_mes=vendas_mes, vendas_ano=vendas_ano,
        n_dia=n_dia, n_mes=n_mes, n_ano=n_ano,
        receber_aberto=receber_aberto, receber_vencido=receber_vencido,
        pagar_aberto=pagar_aberto, pagar_vencido=pagar_vencido,
        produtos_abaixo_minimo=produtos_abaixo_minimo,
        orcamentos_abertos=orcamentos_abertos,
        ranking=ranking,
    )
