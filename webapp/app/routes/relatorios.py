from datetime import date, timedelta

from flask import Blueprint, render_template, request

from ..db import get_db

bp = Blueprint("relatorios", __name__, url_prefix="/relatorios")


def _periodo_padrao():
    """Do dia 1 do mes atual ate hoje."""
    hoje = date.today()
    inicio = hoje.replace(day=1)
    return inicio.isoformat(), hoje.isoformat()


def _ler_periodo():
    inicio_padrao, fim_padrao = _periodo_padrao()
    inicio = request.args.get("inicio", inicio_padrao)
    fim = request.args.get("fim", fim_padrao)
    return inicio, fim


@bp.route("/vendas")
def vendas():
    db = get_db()
    inicio, fim = _ler_periodo()

    vendas_lista = db.execute(
        """SELECT * FROM vendas
           WHERE status = 'Concluida' AND date(data_hora) BETWEEN date(?) AND date(?)
           ORDER BY data_hora DESC""",
        (inicio, fim),
    ).fetchall()

    total = sum(v["total_liquido"] for v in vendas_lista)
    n = len(vendas_lista)
    ticket_medio = (total / n) if n else 0

    por_forma = {}
    for v in vendas_lista:
        if v["valor_pag_1"] and v["forma_pag_1"]:
            por_forma[v["forma_pag_1"]] = por_forma.get(v["forma_pag_1"], 0) + v["valor_pag_1"]
        if v["valor_pag_2"] and v["forma_pag_2"]:
            por_forma[v["forma_pag_2"]] = por_forma.get(v["forma_pag_2"], 0) + v["valor_pag_2"]

    # vendas canceladas e orcamentos no periodo, so pra contexto
    canceladas = db.execute(
        """SELECT COUNT(*) AS n, COALESCE(SUM(total_liquido),0) AS total FROM vendas
           WHERE status = 'Cancelada' AND date(data_hora) BETWEEN date(?) AND date(?)""",
        (inicio, fim),
    ).fetchone()
    orcamentos = db.execute(
        """SELECT COUNT(*) AS n, COALESCE(SUM(total_liquido),0) AS total FROM vendas
           WHERE status = 'Orcamento' AND date(data_hora) BETWEEN date(?) AND date(?)""",
        (inicio, fim),
    ).fetchone()

    return render_template(
        "relatorios/vendas.html", active="relatorios", sub="vendas",
        inicio=inicio, fim=fim, vendas=vendas_lista, total=total, n=n,
        ticket_medio=ticket_medio, por_forma=por_forma,
        canceladas=canceladas, orcamentos=orcamentos,
    )


@bp.route("/produtos")
def produtos():
    db = get_db()
    inicio, fim = _ler_periodo()
    ordenar = request.args.get("ordenar", "faturamento")

    linhas = db.execute(
        """SELECT iv.produto_id, iv.descricao,
                  SUM(iv.qtd) AS qtd_total,
                  SUM(iv.total_item) AS faturamento,
                  COUNT(DISTINCT iv.venda_id) AS n_vendas
           FROM itens_venda iv
           JOIN vendas v ON v.id = iv.venda_id
           WHERE v.status = 'Concluida' AND date(v.data_hora) BETWEEN date(?) AND date(?)
           GROUP BY iv.produto_id, iv.descricao
           ORDER BY """ + ("qtd_total" if ordenar == "qtd" else "faturamento") + " DESC",
        (inicio, fim),
    ).fetchall()

    faturamento_total = sum(l["faturamento"] for l in linhas)

    return render_template(
        "relatorios/produtos.html", active="relatorios", sub="produtos",
        inicio=inicio, fim=fim, linhas=linhas, ordenar=ordenar,
        faturamento_total=faturamento_total,
    )


@bp.route("/margem")
def margem():
    db = get_db()
    inicio, fim = _ler_periodo()

    linhas = db.execute(
        """SELECT iv.produto_id, iv.descricao,
                  SUM(iv.qtd) AS qtd_total,
                  SUM(iv.total_item) AS faturamento,
                  SUM(iv.qtd * COALESCE(p.preco_custo, 0)) AS custo_total
           FROM itens_venda iv
           JOIN vendas v ON v.id = iv.venda_id
           LEFT JOIN produtos p ON p.id = iv.produto_id
           WHERE v.status = 'Concluida' AND date(v.data_hora) BETWEEN date(?) AND date(?)
           GROUP BY iv.produto_id, iv.descricao
           ORDER BY (SUM(iv.total_item) - SUM(iv.qtd * COALESCE(p.preco_custo, 0))) DESC""",
        (inicio, fim),
    ).fetchall()

    resultado = []
    faturamento_total = 0
    custo_total_geral = 0
    for l in linhas:
        lucro = l["faturamento"] - l["custo_total"]
        margem_perc = (lucro / l["faturamento"] * 100) if l["faturamento"] else 0
        faturamento_total += l["faturamento"]
        custo_total_geral += l["custo_total"]
        resultado.append({
            "descricao": l["descricao"],
            "qtd": l["qtd_total"],
            "faturamento": l["faturamento"],
            "custo": l["custo_total"],
            "lucro": lucro,
            "margem_perc": margem_perc,
        })

    lucro_total = faturamento_total - custo_total_geral
    margem_geral = (lucro_total / faturamento_total * 100) if faturamento_total else 0

    return render_template(
        "relatorios/margem.html", active="relatorios", sub="margem",
        inicio=inicio, fim=fim, resultado=resultado,
        faturamento_total=faturamento_total, custo_total_geral=custo_total_geral,
        lucro_total=lucro_total, margem_geral=margem_geral,
    )
