from flask import Blueprint, flash, redirect, render_template, request, url_for

from ..db import get_db, hoje_brasil, parse_num

bp = Blueprint("clientes", __name__, url_prefix="/clientes")


@bp.route("/")
def lista():
    db = get_db()
    busca = request.args.get("q", "").strip()
    if busca:
        like = f"%{busca}%"
        clientes = db.execute(
            """SELECT * FROM clientes
               WHERE nome LIKE ? OR responsavel LIKE ? OR turma LIKE ?
                  OR celular LIKE ? OR CAST(id AS TEXT) = ?
               ORDER BY nome""",
            (like, like, like, like, busca),
        ).fetchall()
    else:
        clientes = db.execute("SELECT * FROM clientes ORDER BY nome").fetchall()
    return render_template("clientes/lista.html", active="clientes", clientes=clientes, busca=busca)


def _saldo_devedor(db, cliente_id):
    row = db.execute(
        """SELECT COALESCE(SUM(valor_parcela - valor_pago), 0) AS saldo
           FROM contas_receber WHERE cliente_id = ? AND status != 'Pago'""",
        (cliente_id,),
    ).fetchone()
    return row["saldo"] or 0


def gasto_hoje(db, cliente_id):
    """Total ja gasto hoje (vendas concluidas) por esse cliente."""
    row = db.execute(
        """SELECT COALESCE(SUM(total_liquido), 0) AS total FROM vendas
           WHERE cliente_id = ? AND status = 'Concluida' AND data_hora::date = ?::date""",
        (cliente_id, hoje_brasil().isoformat()),
    ).fetchone()
    return row["total"] or 0


@bp.route("/novo", methods=["GET", "POST"])
def novo():
    if request.method == "POST":
        return _salvar(None)
    return render_template("clientes/form.html", active="clientes", cliente=None, saldo=0, gasto=0)


@bp.route("/<int:cliente_id>/editar", methods=["GET", "POST"])
def editar(cliente_id):
    db = get_db()
    cliente = db.execute("SELECT * FROM clientes WHERE id = ?", (cliente_id,)).fetchone()
    if cliente is None:
        flash("Cliente nao encontrado.", "erro")
        return redirect(url_for("clientes.lista"))
    if request.method == "POST":
        return _salvar(cliente_id)
    saldo = _saldo_devedor(db, cliente_id)
    gasto = gasto_hoje(db, cliente_id)
    return render_template("clientes/form.html", active="clientes", cliente=cliente, saldo=saldo, gasto=gasto)


def _salvar(cliente_id):
    db = get_db()
    nome = request.form.get("nome", "").strip()
    if not nome:
        flash("Informe o nome.", "erro")
        dest = "clientes.novo" if cliente_id is None else "clientes.editar"
        kwargs = {} if cliente_id is None else {"cliente_id": cliente_id}
        return redirect(url_for(dest, **kwargs))

    responsavel = request.form.get("responsavel", "").strip()
    tutor = request.form.get("tutor", "").strip()
    telefone = request.form.get("telefone", "").strip()
    celular = request.form.get("celular", "").strip()
    limite = parse_num(request.form.get("limite_credito"))
    limite_diario = parse_num(request.form.get("limite_diario"))
    alergia = request.form.get("alergia", "").strip()
    status = request.form.get("status", "Ativo")
    permite_a_prazo = request.form.get("permite_a_prazo", "Nao")
    turma = request.form.get("turma", "").strip()

    if cliente_id is None:
        db.execute(
            """INSERT INTO clientes
               (nome, responsavel, tutor, telefone, celular, limite_credito,
                status, permite_a_prazo, turma, limite_diario, alergia)
               VALUES (?,?,?,?,?,?,?,?,?,?,?)""",
            (nome, responsavel, tutor, telefone, celular, limite,
             status, permite_a_prazo, turma, limite_diario, alergia),
        )
        flash("Cliente cadastrado com sucesso.", "sucesso")
    else:
        db.execute(
            """UPDATE clientes SET nome=?, responsavel=?, tutor=?, telefone=?,
               celular=?, limite_credito=?, status=?, permite_a_prazo=?, turma=?,
               limite_diario=?, alergia=?
               WHERE id=?""",
            (nome, responsavel, tutor, telefone, celular, limite,
             status, permite_a_prazo, turma, limite_diario, alergia, cliente_id),
        )
        flash("Cliente atualizado com sucesso.", "sucesso")
    db.commit()
    return redirect(url_for("clientes.lista"))


@bp.route("/<int:cliente_id>/excluir", methods=["POST"])
def excluir(cliente_id):
    db = get_db()
    saldo = _saldo_devedor(db, cliente_id)
    if saldo > 0:
        flash("Nao e possivel excluir: o cliente possui parcelas em aberto.", "erro")
        return redirect(url_for("clientes.editar", cliente_id=cliente_id))
    db.execute("DELETE FROM clientes WHERE id = ?", (cliente_id,))
    db.commit()
    flash("Cliente excluido.", "sucesso")
    return redirect(url_for("clientes.lista"))


@bp.route("/buscar.json")
def buscar_json():
    """Usado pelo PDV para localizar cliente por numero ou nome/responsavel/turma."""
    db = get_db()
    termo = request.args.get("q", "").strip()
    if not termo:
        return {"resultados": []}
    like = f"%{termo}%"
    rows = db.execute(
        """SELECT id, nome, responsavel, turma, celular, limite_diario, alergia FROM clientes
           WHERE CAST(id AS TEXT) = ? OR nome LIKE ? OR responsavel LIKE ?
              OR turma LIKE ? OR celular LIKE ?
           ORDER BY nome LIMIT 20""",
        (termo, like, like, like, like),
    ).fetchall()
    return {"resultados": [dict(r) for r in rows]}


@bp.route("/<int:cliente_id>/status.json")
def status_json(cliente_id):
    """Usado pelo PDV: alergia + situacao do limite diario do cliente selecionado."""
    db = get_db()
    cliente = db.execute(
        "SELECT id, nome, limite_diario, alergia FROM clientes WHERE id = ?", (cliente_id,)
    ).fetchone()
    if cliente is None:
        return {"encontrado": False}
    gasto = gasto_hoje(db, cliente_id)
    limite = cliente["limite_diario"] or 0
    return {
        "encontrado": True,
        "alergia": cliente["alergia"] or "",
        "limite_diario": limite,
        "gasto_hoje": gasto,
        "disponivel_hoje": (limite - gasto) if limite > 0 else None,
    }
