import test from "node:test";
import assert from "node:assert/strict";
import {relatorioIndemnizacoes} from "../relatorios-dados.js";

const S = iso => String(Math.round((Date.parse(iso) - Date.parse("1899-12-30")) / 86400000));
const ind = [
  {ID_Indemnizacao: "IND-00001", Empresa: "SIR Motors", Codigo: "0001", Nome: "Ana", "Data Admissao": S("2019-11-10"), "Data Referencia": S("2026-10-28"), "Anos de Trabalho": 6, "Vencimento Base": 20000, "Valor Total": 120000, "Numero Prestacoes": 3, Estado: "Activa"},
  {ID_Indemnizacao: "IND-00002", Empresa: "SIR Log", Codigo: "0002", Nome: "Rui", "Valor Total": 50000, "Numero Prestacoes": 1, Estado: "Anulada"}];
const pre = [
  {ID_Prestacao: "a", ID_Indemnizacao: "IND-00001", Nome: "Ana", Numero: 1, "Data Limite": S("2026-09-25"), Valor: 40000, Estado: "Paga", "Data Pagamento": S("2026-09-20")},
  {ID_Prestacao: "b", ID_Indemnizacao: "IND-00001", Nome: "Ana", Numero: 2, "Data Limite": S("2026-10-25"), Valor: 40000, Estado: "Pendente"},
  {ID_Prestacao: "c", ID_Indemnizacao: "IND-00001", Nome: "Ana", Numero: 3, "Data Limite": S("2026-11-25"), Valor: 40000, Estado: "Pendente"},
  {ID_Prestacao: "d", ID_Indemnizacao: "IND-00002", Nome: "Rui", Numero: 1, "Data Limite": S("2026-10-10"), Valor: 50000, Estado: "Pendente"}];

test("relatório de indemnizações: totais, mês, atraso e anulados fora", () => {
  const r = relatorioIndemnizacoes({indemnizacoes: ind, prestacoes: pre, ano: 2026, mes: 10, hoje: new Date("2026-10-28T10:00:00Z")});
  assert.ok(!r.vazio);
  const k = Object.fromEntries(r.kpis.map(x => [x.rotulo, x.valor]));
  assert.match(k["Total comprometido"], /^120.000,00 MT$/);
  assert.match(k["Valor pago"], /^40.000,00/);
  assert.match(k["Em atraso"], /^40.000,00/);
  assert.equal(r.secoes[1].tabela.linhas.length, 1); // só a prestação de Outubro da Ana (Rui anulado)
  assert.ok(r.secoes.some(s => /em atraso/.test(s.titulo)));
  assert.equal(relatorioIndemnizacoes({indemnizacoes: [], prestacoes: [], ano: 2026, mes: 10}).vazio.length > 0, true);
});
