import barcode
from barcode.writer import SVGWriter
from io import BytesIO

from flask import Blueprint, render_template, request

from ..db import get_db, parse_int
from .config import get_config

bp = Blueprint("impressao", __name__, url_prefix="/impressao")


@bp.route("/vendas/<int:venda_id>/cupom")
def cupom(venda_id):
    db = get_db()
    cfg = get_config(db)
    venda = db.execute("SELECT * FROM vendas WHERE id = ?", (venda_id,)).fetchone()
    itens = db.execute(
        "SELECT * FROM itens_venda WHERE venda_id = ?", (venda_id,)
    ).fetchall()
    largura = parse_int(cfg.get("LARGURA_CUPOM"), 80)
    return render_template(
        "impressao/cupom.html", venda=venda, itens=itens, cfg=cfg, largura=largura
    )


@bp.route("/vendas/<int:venda_id>/carne")
def carne(venda_id):
    db = get_db()
    cfg = get_config(db)
    venda = db.execute("SELECT * FROM vendas WHERE id = ?", (venda_id,)).fetchone()
    parcelas = db.execute(
        "SELECT * FROM contas_receber WHERE venda_id = ? ORDER BY vencimento",
        (venda_id,),
    ).fetchall()
    return render_template(
        "impressao/carne.html", venda=venda, parcelas=parcelas, cfg=cfg,
        total=len(parcelas),
    )


def _barcode_svg(codigo):
    """Gera o SVG do codigo de barras (Code128) para um texto/numero."""
    try:
        rv = BytesIO()
        barcode.get("code128", str(codigo), writer=SVGWriter()).write(
            rv, options={"module_height": 10, "write_text": False, "quiet_zone": 1}
        )
        return rv.getvalue().decode("utf-8")
    except Exception:
        return ""


@bp.route("/etiquetas", methods=["GET", "POST"])
def etiquetas():
    db = get_db()
    if request.method == "POST":
        ini = request.form.get("inicial", "").strip()
        fim = request.form.get("final", "").strip()
        copias = parse_int(request.form.get("copias"), 1)

        query = "SELECT * FROM produtos WHERE ativo = 1"
        params = []
        if ini:
            query += " AND id >= ?"
            params.append(parse_int(ini))
        if fim:
            query += " AND id <= ?"
            params.append(parse_int(fim))
        query += " ORDER BY descricao"
        produtos = db.execute(query, params).fetchall()

        etiquetas_lista = []
        for p in produtos:
            codigo = p["codigo_barras"] or str(p["id"])
            for _ in range(max(1, copias)):
                etiquetas_lista.append({
                    "descricao": p["descricao"],
                    "preco": p["preco_varejo"],
                    "unidade": p["unidade"],
                    "codigo": codigo,
                    "svg": _barcode_svg(codigo),
                })
        return render_template("impressao/etiquetas.html", etiquetas=etiquetas_lista)

    return render_template("impressao/etiquetas_form.html", active="produtos")
