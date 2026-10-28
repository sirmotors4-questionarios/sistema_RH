import test from "node:test";
import assert from "node:assert/strict";
import * as LC from "../lancamentos-colunas.js";

const linhas = [
  {Codigo: "0001", Nome: "A", Vencimento: "100", "Subsídios": "50", "Data Pagamento": "46000", Obs: "x", Ano: "2026"},
  {Codigo: "0002", Nome: "B", Vencimento: "90", "Subsídios": "", "Data Pagamento": "45990", Obs: "y", Ano: "2026"}
];

test("catálogo junta a estrutura padrão e as outras colunas do Excel, com o tipo certo", () => {
  const cat = LC.catalogo(linhas), extra = cat.filter(c => c.extra);
  assert.deepEqual(extra.map(c => [c.rotulo, c.tipo, c.total]),
    [["Subsídios", "numero", true], ["Data Pagamento", "data", false], ["Obs", "texto", false], ["Ano", "numero", false]]);
  assert.equal(cat.filter(c => !c.extra).length, LC.PADRAO.length);
});

test("ordenação: número, texto e data; vazios no fim; terceiro clique volta à ordem original", () => {
  const num = {tipo: "numero", valor: l => l.v}, txt = {tipo: "texto", valor: l => l.v}, dt = {tipo: "data", valor: l => l.v};
  assert.deepEqual(LC.ordenar([{v: "10"}, {v: ""}, {v: "9"}], num, "asc").map(x => x.v), ["9", "10", ""]);
  assert.deepEqual(LC.ordenar([{v: "Álvaro"}, {v: "bruno"}, {v: "Zé"}], txt, "desc").map(x => x.v), ["Zé", "bruno", "Álvaro"]);
  assert.deepEqual(LC.ordenar([{v: "02/01/2026"}, {v: "2025-12-31"}], dt, "asc").map(x => x.v), ["2025-12-31", "02/01/2026"]);
  let o = LC.proximaOrdem(null, "x"); assert.deepEqual(o, {id: "x", dir: "asc"});
  o = LC.proximaOrdem(o, "x"); assert.deepEqual(o, {id: "x", dir: "desc"});
  assert.equal(LC.proximaOrdem(o, "x"), null);
});

test("preferências: ids desconhecidos saem, o Nome fica sempre, a estrutura padrão é reconhecida", () => {
  const cat = LC.catalogo(linhas);
  const p = LC.normalizar({colunas: ["liquido", "x:Subsídios", "naoexiste", "codigo"], ordem: {id: "naoexiste", dir: "asc"}}, cat);
  assert.deepEqual(p.colunas, ["liquido", "nome", "x:Subsídios", "codigo"]);
  assert.equal(p.ordem, null);
  assert.ok(LC.eDefault(LC.normalizar(null, cat)));
});
