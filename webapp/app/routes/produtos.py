from flask import Blueprint, flash, redirect, render_template, request, url_for

from ..db import get_db, parse_num

bp = Blueprint("produtos", __name__, url_prefix="/produtos")


def calcular_precos(custo, venda):
    """A partir do preco de compra (custo) e do preco de venda (varejo)
    digitados, calcula a margem % e os precos de atacado/cartao derivados
    do varejo pelos percentuais configurados."""
    db = get_db()
    cfg = {r["chave"]: r["valor"] for r in db.execute("SELECT chave, valor FROM config")}
    perc_atacado = parse_num(cfg.get("PERC_ATACADO"), 8)
    perc_cartao = parse_num(cfg.get("PERC_CARTAO"), 5)

    margem = ((venda - custo) / custo * 100) if custo > 0 else 0
    atacado = venda * (1 - perc_atacado / 100)
    cartao = venda * (1 + perc_cartao / 100)
    return margem, atacado, cartao


@bp.route("/")
def lista():
    db = get_db()
    busca = request.args.get("q", "").strip()
    if busca:
        like = f"%{busca}%"
        produtos = db.execute(
            """SELECT * FROM produtos WHERE ativo = 1
               AND (descricao LIKE ? OR codigo_barras LIKE ? OR CAST(id AS TEXT) = ?)
               ORDER BY descricao""",
            (like, like, busca),
        ).fetchall()
    else:
        produtos = db.execute(
            "SELECT * FROM produtos WHERE ativo = 1 ORDER BY descricao"
        ).fetchall()
    return render_template("produtos/lista.html", active="produtos", produtos=produtos, busca=busca)


@bp.route("/novo", methods=["GET", "POST"])
def novo():
    if request.method == "POST":
        return _salvar(None)
    return render_template("produtos/form.html", active="produtos", produto=None)


@bp.route("/<int:produto_id>/editar", methods=["GET", "POST"])
def editar(produto_id):
    db = get_db()
    produto = db.execute("SELECT * FROM produtos WHERE id = ?", (produto_id,)).fetchone()
    if produto is None:
        flash("Produto nao encontrado.", "erro")
        return redirect(url_for("produtos.lista"))
    if request.method == "POST":
        return _salvar(produto_id)
    return render_template("produtos/form.html", active="produtos", produto=produto)


def _salvar(produto_id):
    db = get_db()
    descricao = request.form.get("descricao", "").strip()
    if not descricao:
        flash("Informe a descricao do produto.", "erro")
        dest = "produtos.novo" if produto_id is None else "produtos.editar"
        kwargs = {} if produto_id is None else {"produto_id": produto_id}
        return redirect(url_for(dest, **kwargs))

    codigo_barras = request.form.get("codigo_barras", "").strip() or None
    unidade = request.form.get("unidade", "UN")
    custo = parse_num(request.form.get("preco_custo"))
    varejo = parse_num(request.form.get("preco_venda"))
    estoque_atual = parse_num(request.form.get("estoque_atual"))
    estoque_minimo = parse_num(request.form.get("estoque_minimo"))
    fornecedor = request.form.get("fornecedor", "").strip()
    validade = request.form.get("validade") or None

    margem, atacado, cartao = calcular_precos(custo, varejo)

    if produto_id is None:
        db.execute(
            """INSERT INTO produtos
               (codigo_barras, descricao, unidade, preco_custo, margem,
                preco_varejo, preco_atacado, preco_cartao,
                estoque_atual, estoque_minimo, fornecedor, validade)
               VALUES (?,?,?,?,?,?,?,?,?,?,?,?)""",
            (codigo_barras, descricao, unidade, custo, margem,
             varejo, atacado, cartao, estoque_atual, estoque_minimo,
             fornecedor, validade),
        )
        flash("Produto cadastrado com sucesso.", "sucesso")
    else:
        db.execute(
            """UPDATE produtos SET codigo_barras=?, descricao=?, unidade=?,
               preco_custo=?, margem=?, preco_varejo=?, preco_atacado=?,
               preco_cartao=?, estoque_atual=?, estoque_minimo=?,
               fornecedor=?, validade=? WHERE id=?""",
            (codigo_barras, descricao, unidade, custo, margem,
             varejo, atacado, cartao, estoque_atual, estoque_minimo,
             fornecedor, validade, produto_id),
        )
        flash("Produto atualizado com sucesso.", "sucesso")
    db.commit()
    return redirect(url_for("produtos.lista"))


@bp.route("/<int:produto_id>/excluir", methods=["POST"])
def excluir(produto_id):
    db = get_db()
    db.execute("UPDATE produtos SET ativo = 0 WHERE id = ?", (produto_id,))
    db.commit()
    flash("Produto removido.", "sucesso")
    return redirect(url_for("produtos.lista"))
