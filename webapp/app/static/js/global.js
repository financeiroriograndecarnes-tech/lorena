(function () {
  "use strict";

  // Em TODA tela do sistema: ao focar um campo numerico (type="number"
  // ou type="text" com inputmode decimal/numeric -- usado no lugar de
  // type="number" porque no Android o .select() em type="number" as
  // vezes nao seleciona nada), seleciona o conteudo inteiro. Assim
  // digitar substitui em vez de grudar no que ja estava la (o "0200"
  // que aparecia ao tentar digitar "200" em cima do zero).
  function ehCampoNumerico(el) {
    if (!el || el.tagName !== "INPUT") return false;
    if (el.type === "number") return true;
    const modo = el.getAttribute("inputmode");
    return modo === "decimal" || modo === "numeric";
  }

  document.addEventListener(
    "focusin",
    function (e) {
      if (ehCampoNumerico(e.target)) {
        e.target.select();
      }
    },
    true
  );
})();
