from flask import Blueprint, jsonify, render_template, request

from ..db import get_db
from .caixa import caixa_aberto, registrar_movimento

bp = Blueprint("pdv", __name__, url_prefix="/pdv")

TABELAS = {"VAREJO": "preco_varejo", "ATACADO": "preco_atacado", "CARTAO": "preco_cartao"}


@bp.route("/")
def tela():
    db = get_db()
    aberto = caixa_aberto(db)
    return render_template("pdv/tela.html", active="pdv", caixa_aberto=aberto)


@bp.route("/produto.json")
def produto_json():
    """Localiza produto por codigo de barras OU codigo interno."""
    db = get_db()
    chave = request.args.get("codigo", "").strip()
    tabela = request.args.get("tabela", "VAREJO").upper()
    if not chave:
        return jsonify({"encontrado": False})

    produto = db.execute(
        "SELECT * FROM produtos WHERE ativo = 1 AND codigo_barras = ?", (chave,)
    ).fetchone()
    if produto is None and chave.isdigit():
        produto = db.execute(
            "SELECT * FROM produtos WHERE ativo = 1 AND id = ?", (int(chave),)
        ).fetchone()
    if produto is None:
        return jsonify({"encontrado": False})

    col = TABELAS.get(tabela, "preco_varejo")
    preco = produto[col] or produto["preco_varejo"]
    return jsonify({
        "encontrado": True,
        "id": produto["id"],
        "descricao": produto["descricao"],
        "unidade": produto["unidade"],
        "preco": preco,
        "estoque_atual": produto["estoque_atual"],
    })


@bp.route("/finalizar", methods=["POST"])
def finalizar():
    db = get_db()
    dados = request.get_json(force=True)

    itens = dados.get("itens", [])
    status = dados.get("status", "Concluida")
    if not itens:
        return jsonify({"ok": False, "erro": "Nenhum item na venda."}), 400

    if status == "Concluida" and not caixa_aberto(db):
        return jsonify({"ok": False, "erro": "O caixa esta fechado. Abra o caixa antes de finalizar."}), 400

    cliente_id = dados.get("cliente_id")
    cliente_nome = dados.get("cliente_nome") or "CONSUMIDOR"
    vendedor = dados.get("vendedor") or "OPERADOR"
    subtotal = float(dados.get("subtotal") or 0)
    desconto = float(dados.get("desconto") or 0)
    total = float(dados.get("total") or 0)
    fp1 = dados.get("forma_pag_1") or ""
    vp1 = float(dados.get("valor_pag_1") or 0)
    fp2 = dados.get("forma_pag_2") or ""
    vp2 = float(dados.get("valor_pag_2") or 0)
    troco = float(dados.get("troco") or 0)
    tabela_preco = dados.get("tabela_preco", "VAREJO")
    n_parcelas = int(dados.get("n_parcelas") or 1)

    if status == "Concluida":
        pago = vp1 + vp2
        if pago < total - 0.009:
            return jsonify({"ok": False, "erro": "Valor pago menor que o total."}), 400
        if fp1.upper() == "A PRAZO" or fp2.upper() == "A PRAZO":
            valor_prazo = (vp1 if fp1.upper() == "A PRAZO" else 0) + (vp2 if fp2.upper() == "A PRAZO" else 0)
            erro = _validar_venda_a_prazo(db, cliente_id, valor_prazo)
            if erro:
                return jsonify({"ok": False, "erro": erro}), 400

    cur = db.execute(
        """INSERT INTO vendas (cliente_id, cliente_nome, vendedor, total_bruto, desconto,
           total_liquido, forma_pag_1, valor_pag_1, forma_pag_2, valor_pag_2, troco,
           status, tabela_preco)
           VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)""",
        (cliente_id, cliente_nome, vendedor, subtotal, desconto, total,
         fp1, vp1, fp2, vp2, troco, status, tabela_preco),
    )
    venda_id = cur.lastrowid

    for item in itens:
        db.execute(
            """INSERT INTO itens_venda (venda_id, produto_id, descricao, qtd, preco_unit, total_item)
               VALUES (?,?,?,?,?,?)""",
            (venda_id, item.get("produto_id"), item["descricao"], item["qtd"],
             item["preco_unit"], item["total"]),
        )
        if status in ("Concluida", "Consignado") and item.get("produto_id"):
            db.execute(
                "UPDATE produtos SET estoque_atual = estoque_atual - ? WHERE id = ?",
                (item["qtd"], item["produto_id"]),
            )

    if status == "Concluida":
        if vp1 > 0:
            valor_caixa = (vp1 - troco) if fp1.upper() == "DINHEIRO" else vp1
            registrar_movimento(db, "Venda", valor_caixa, fp1, f"Venda #{venda_id}", vendedor)
        if vp2 > 0:
            registrar_movimento(db, "Venda", vp2, fp2, f"Venda #{venda_id}", vendedor)

        if fp1.upper() == "A PRAZO" or fp2.upper() == "A PRAZO":
            valor_prazo = (vp1 if fp1.upper() == "A PRAZO" else 0) + (vp2 if fp2.upper() == "A PRAZO" else 0)
            _gerar_parcelas(db, venda_id, cliente_id, valor_prazo, n_parcelas)

    db.commit()
    return jsonify({"ok": True, "venda_id": venda_id})


def _validar_venda_a_prazo(db, cliente_id, valor):
    if not cliente_id:
        return "Venda a prazo exige cliente cadastrado."
    cliente = db.execute("SELECT * FROM clientes WHERE id = ?", (cliente_id,)).fetchone()
    if cliente is None:
        return "Cliente nao encontrado no cadastro."
    if cliente["status"] == "Bloqueado":
        return "Cliente BLOQUEADO no cadastro."
    if cliente["permite_a_prazo"] != "Sim":
        return "Cliente nao esta autorizado a comprar a prazo."
    venc = db.execute(
        """SELECT COUNT(*) AS n FROM contas_receber
           WHERE cliente_id = ? AND status != 'Pago' AND vencimento < date('now')""",
        (cliente_id,),
    ).fetchone()
    if venc["n"] > 0:
        return "Cliente possui parcelas VENCIDAS em aberto."
    saldo = db.execute(
        """SELECT COALESCE(SUM(valor_parcela - valor_pago), 0) AS saldo
           FROM contas_receber WHERE cliente_id = ? AND status != 'Pago'""",
        (cliente_id,),
    ).fetchone()["saldo"]
    disponivel = cliente["limite_credito"] - saldo
    if valor > disponivel:
        return f"Limite insuficiente. Disponivel: R$ {disponivel:.2f} | Necessario: R$ {valor:.2f}"
    return None


def _gerar_parcelas(db, venda_id, cliente_id, valor_total, n_parcelas):
    n_parcelas = max(1, n_parcelas)
    valor_parcela = round(valor_total / n_parcelas, 2)
    soma = 0
    for i in range(n_parcelas):
        if i < n_parcelas - 1:
            valor = valor_parcela
            soma += valor
        else:
            valor = round(valor_total - soma, 2)
        db.execute(
            """INSERT INTO contas_receber (venda_id, cliente_id, vencimento, valor_parcela, status)
               VALUES (?, ?, date('now', ?), ?, 'Em Aberto')""",
            (venda_id, cliente_id, f"+{30 * (i + 1)} days", valor),
        )
