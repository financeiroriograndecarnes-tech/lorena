(function () {
  "use strict";

  let itens = [];
  let clienteId = null;

  const $ = (id) => document.getElementById(id);
  const moeda = (v) => "R$ " + Number(v || 0).toFixed(2).replace(".", ",");

  function recalcular() {
    const subtotal = itens.reduce((s, i) => s + i.total, 0);
    const desconto = Math.min(parseFloat($("in-desconto").value) || 0, subtotal);
    const total = Math.round((subtotal - desconto) * 100) / 100;
    const vp1 = parseFloat($("in-vp1").value) || 0;
    const vp2 = parseFloat($("in-vp2").value) || 0;
    const pago = vp1 + vp2;
    const troco = Math.max(0, pago - total);

    $("lbl-subtotal").textContent = "SUBTOTAL: " + moeda(subtotal);
    $("lbl-total").textContent = "TOTAL: " + moeda(total);
    $("lbl-troco").textContent = moeda(troco);
    return { subtotal, desconto, total, vp1, vp2, troco };
  }

  function redesenharItens() {
    const tbody = $("tbody-itens");
    tbody.innerHTML = "";
    itens.forEach((item, idx) => {
      const tr = document.createElement("tr");
      tr.innerHTML =
        "<td>" + item.descricao + "</td>" +
        "<td>" + item.qtd + " " + item.unidade + "</td>" +
        "<td>" + moeda(item.preco_unit) + "</td>" +
        "<td>" + moeda(item.total) + "</td>" +
        '<td><button type="button" class="btn btn-cinza" data-idx="' + idx + '">Remover</button></td>';
      tbody.appendChild(tr);
    });
    tbody.querySelectorAll("button[data-idx]").forEach((btn) => {
      btn.addEventListener("click", () => {
        itens.splice(parseInt(btn.dataset.idx, 10), 1);
        redesenharItens();
        recalcular();
      });
    });
  }

  async function adicionarProduto() {
    const codigo = $("in-codigo").value.trim();
    const qtd = parseFloat($("in-qtd").value) || 1;
    const tabela = $("in-tabela").value;
    if (!codigo) return;

    const resp = await fetch("/pdv/produto.json?codigo=" + encodeURIComponent(codigo) + "&tabela=" + tabela);
    const dados = await resp.json();
    if (!dados.encontrado) {
      $("msg-produto").textContent = "Produto nao encontrado: " + codigo;
      $("in-codigo").value = "";
      $("in-codigo").focus();
      return;
    }

    const existente = itens.find((i) => i.produto_id === dados.id && i.preco_unit === dados.preco);
    if (existente) {
      existente.qtd += qtd;
      existente.total = Math.round(existente.qtd * existente.preco_unit * 100) / 100;
    } else {
      itens.push({
        produto_id: dados.id,
        descricao: dados.descricao,
        unidade: dados.unidade,
        qtd: qtd,
        preco_unit: dados.preco,
        total: Math.round(qtd * dados.preco * 100) / 100,
      });
    }
    $("msg-produto").textContent = dados.descricao + " — " + moeda(dados.preco) + " / " + dados.unidade;
    $("in-codigo").value = "";
    $("in-qtd").value = "1";
    $("in-codigo").focus();
    redesenharItens();
    recalcular();
  }

  function limparVenda() {
    itens = [];
    clienteId = null;
    $("in-cliente-cod").value = "";
    $("in-cliente-nome").value = "CONSUMIDOR";
    $("in-desconto").value = "0";
    $("in-vp1").value = "0";
    $("in-vp2").value = "0";
    $("in-parcelas").value = "1";
    redesenharItens();
    recalcular();
    $("in-codigo").focus();
  }

  async function buscarClientePorCodigo() {
    const cod = $("in-cliente-cod").value.trim();
    if (!cod) return;
    const resp = await fetch("/clientes/buscar.json?q=" + encodeURIComponent(cod));
    const dados = await resp.json();
    const achou = dados.resultados.find((r) => String(r.id) === cod);
    if (achou) {
      clienteId = achou.id;
      $("in-cliente-nome").value = achou.nome;
    } else {
      alert("Cliente numero " + cod + " nao encontrado.");
    }
  }

  async function abrirBuscaCliente() {
    const caixa = $("lista-busca-cliente");
    caixa.style.display = "block";
    caixa.innerHTML = "<p>Digite para buscar por nome, responsavel, turma ou celular:</p>" +
      '<input type="text" id="in-filtro-cliente" placeholder="Buscar...">' +
      '<div id="resultado-busca-cliente"></div>';
    $("in-filtro-cliente").addEventListener("input", async (e) => {
      const termo = e.target.value.trim();
      const div = $("resultado-busca-cliente");
      if (!termo) { div.innerHTML = ""; return; }
      const resp = await fetch("/clientes/buscar.json?q=" + encodeURIComponent(termo));
      const dados = await resp.json();
      div.innerHTML = "";
      dados.resultados.forEach((r) => {
        const item = document.createElement("div");
        item.className = "card";
        item.style.cursor = "pointer";
        item.style.margin = "8px 0";
        item.innerHTML = "<strong>#" + r.id + " — " + r.nome + "</strong><br>" +
          (r.responsavel ? "Responsavel: " + r.responsavel + " · " : "") +
          (r.turma ? "Turma: " + r.turma : "");
        item.addEventListener("click", () => {
          clienteId = r.id;
          $("in-cliente-cod").value = r.id;
          $("in-cliente-nome").value = r.nome;
          caixa.style.display = "none";
        });
        div.appendChild(item);
      });
    });
    $("in-filtro-cliente").focus();
  }

  async function finalizar(status) {
    const totais = recalcular();
    if (itens.length === 0) {
      alert("Nenhum item na venda.");
      return;
    }
    const payload = {
      itens: itens,
      status: status,
      cliente_id: clienteId,
      cliente_nome: status === "Orcamento" ? "" : $("in-cliente-nome").value,
      vendedor: "OPERADOR",
      subtotal: totais.subtotal,
      desconto: totais.desconto,
      total: totais.total,
      forma_pag_1: status === "Orcamento" ? "" : $("in-fp1").value,
      valor_pag_1: status === "Orcamento" ? 0 : totais.vp1,
      forma_pag_2: status === "Orcamento" ? "" : $("in-fp2").value,
      valor_pag_2: status === "Orcamento" ? 0 : totais.vp2,
      troco: status === "Orcamento" ? 0 : totais.troco,
      tabela_preco: $("in-tabela").value,
      n_parcelas: parseInt($("in-parcelas").value, 10) || 1,
    };

    const resp = await fetch("/pdv/finalizar", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
    });
    const dados = await resp.json();
    if (!dados.ok) {
      alert("Erro: " + dados.erro);
      return;
    }
    if (confirm("Venda #" + dados.venda_id + " registrada com sucesso.\n\nImprimir comprovante?")) {
      window.open("/impressao/vendas/" + dados.venda_id + "/cupom", "_blank");
    }
    if (dados.a_prazo && confirm("Gerar carnê / promissória desta venda?")) {
      window.open("/impressao/vendas/" + dados.venda_id + "/carne", "_blank");
    }
    limparVenda();
  }

  $("btn-adicionar").addEventListener("click", adicionarProduto);
  $("in-codigo").addEventListener("keydown", (e) => {
    if (e.key === "Enter") { e.preventDefault(); adicionarProduto(); }
  });
  $("in-cliente-cod").addEventListener("change", buscarClientePorCodigo);
  $("btn-buscar-cliente").addEventListener("click", abrirBuscaCliente);
  $("in-desconto").addEventListener("input", recalcular);
  $("in-vp1").addEventListener("input", recalcular);
  $("in-vp2").addEventListener("input", recalcular);
  $("btn-finalizar").addEventListener("click", () => finalizar("Concluida"));
  $("btn-orcamento").addEventListener("click", () => finalizar("Orcamento"));
  $("btn-limpar").addEventListener("click", () => {
    if (itens.length === 0 || confirm("Cancelar a venda em andamento?")) limparVenda();
  });

  recalcular();
  $("in-codigo").focus();
})();
