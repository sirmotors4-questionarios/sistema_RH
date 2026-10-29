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

test("colunas do cadastro: só as que não existem na tbl_lançamento, ligadas por Empresa + Código", () => {
  const lanc = [{Empresa: "SIR", Codigo: "0001", Nome: "A", Profissao: "X", Vencimento: "1"}, {Empresa: "SIR", Codigo: "0002", Nome: "B", Profissao: "Y", Vencimento: "2"}];
  const func = [{Empresa: "SIR", Codigo: "1", Nome: "A", "Profissão": "X", "Data Admissao": "45000", NUIT: "123", Sexo: "F"},
    {Empresa: "SIR", Codigo: "0002", Nome: "B", "Profissão": "Y", "Data Admissao": "44000", NUIT: "456", Sexo: "M"}];
  const cat = LC.catalogo(lanc, func), cad = cat.filter(c => c.origem === "funcionarios");
  assert.deepEqual(cad.map(c => [c.rotulo, c.tipo]), [["Data Admissao", "data"], ["NUIT", "numero"], ["Sexo", "texto"]]);
  const sexo = cad.find(c => c.rotulo === "Sexo"), m = {linha: lanc[0]};
  assert.equal(sexo.valor(m), "F"); // código 0001 ↔ 1
  // preferência com coluna do cadastro ainda por carregar não se perde
  assert.ok(LC.normalizar({colunas: ["nome", "f:Sexo"]}, LC.catalogo(lanc)).colunas.includes("f:Sexo"));
});

test('filtros no cabeçalho: valores distintos e filtragem como no Excel', async () => {
  const LC = await import('../lancamentos-colunas.js');
  const M = await import('../rh-modulos.js');
  const linhas = [
    {Ano: '2026', 'Mês': 'Outubro', Codigo: '1', Nome: 'A', Empresa: 'SIR', 'Escalão': '10º Escalão', 'Salario Liquido': 5000},
    {Ano: '2026', 'Mês': 'Outubro', Codigo: '2', Nome: 'B', Empresa: 'SIR', 'Escalão': '2º Escalão', 'Salario Liquido': 1000},
    {Ano: '2026', 'Mês': 'Outubro', Codigo: '3', Nome: 'C', Empresa: 'OUT', 'Escalão': '2º Escalão', 'Salario Liquido': 300},
    {Ano: '2026', 'Mês': 'Outubro', Codigo: '4', Nome: 'D', Empresa: 'OUT', 'Salario Liquido': 300}];
  const lista = M.lancamentos(linhas, 2026, 10);
  const cat = LC.catalogo(linhas), porId = new Map(cat.map(c => [c.id, c]));
  const esc = cat.find(c => c.rotulo === 'Escalão');
  assert.ok(esc, 'coluna Escalão no catálogo');
  assert.deepEqual(LC.valoresColuna(lista, esc).map(v => v.valor + ':' + v.n), ['2º Escalão:2', '10º Escalão:1', ':1']);
  const f = {[esc.id]: ['2º Escalão']};
  assert.deepEqual(LC.filtrarCabecalho(lista, f, porId).map(l => l.codigo), ['0002', '0003']);
  const emp = cat.find(c => /empresa/i.test(c.rotulo));
  const f2 = {...f, [emp.id]: ['OUT']};
  assert.deepEqual(LC.filtrarCabecalho(lista, f2, porId).map(l => l.codigo), ['0003']);
  // valores de uma coluna ignoram o filtro dessa própria coluna
  assert.deepEqual(LC.valoresColuna(LC.filtrarCabecalho(lista, f2, porId, esc.id), esc).map(v => v.valor), ['2º Escalão', '']);
  assert.equal(LC.filtrosActivos(f2), 2);
  assert.equal(LC.filtrarCabecalho(lista, {}, porId).length, 4);
});
