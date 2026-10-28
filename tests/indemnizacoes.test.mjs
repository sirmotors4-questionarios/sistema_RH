import test from "node:test";
import assert from "node:assert/strict";
import * as I from "../indemnizacoes.js";

test("anos completos e total", () => {
  assert.equal(I.anosCompletos("2019-11-10", "2026-10-28"), 6);
  assert.equal(I.anosCompletos("2019-10-28", "2026-10-28"), 7);
  assert.equal(I.anosCompletos("2026-01-01", "2026-10-28"), 0);
  assert.deepEqual(I.calcular({admissao: "2019-11-10", referencia: "2026-10-28", vencimentoBase: "20.000,00"}), {anos: 6, base: 20000, total: 120000});
});

test("plano equitativo: cêntimos certos, dia fixo, fim de mês curto", () => {
  const p = I.gerarPlano(100000, 3, 31, "2027-01");
  assert.deepEqual(p.map(x => x.dataLimite), ["2027-01-31", "2027-02-28", "2027-03-31"]);
  assert.deepEqual(p.map(x => x.valor), [33333.33, 33333.33, 33333.34]);
  assert.deepEqual(I.validarPlano(p, 100000), []);
  assert.match(I.validarPlano([{dataLimite: "2027-02-01", valor: 1}, {dataLimite: "2027-01-01", valor: 1}], 2)[0], /anterior/);
});

test("indicadores do mês, atraso e gasto", () => {
  const ind = [{id: "IND-1", total: 300, estado: "Activa"}, {id: "IND-2", total: 50, estado: "Anulada"}];
  const pr = [
    {id: "a", indemnizacao: "IND-1", numero: 1, dataLimite: "2026-09-25", valor: 100, estado: "Paga", dataPagamento: "2026-09-20"},
    {id: "b", indemnizacao: "IND-1", numero: 2, dataLimite: "2026-10-25", valor: 100, estado: "Pendente"},
    {id: "c", indemnizacao: "IND-1", numero: 3, dataLimite: "2026-11-25", valor: 100, estado: "Pendente"},
    {id: "d", indemnizacao: "IND-2", numero: 1, dataLimite: "2026-10-10", valor: 50, estado: "Pendente"}];
  const k = I.indicadoresMes(ind, pr, 2026, 10, "2026-10-28");
  assert.equal(k.valorMes, 100); assert.equal(k.atrasoMes, 100); assert.equal(k.gastoTotal, 100);
  assert.equal(k.comprometido, 300); assert.equal(k.saldo, 200); assert.equal(k.nAtraso, 1);
  const r = I.resumo(ind[0], pr);
  assert.equal(r.pagas, 1); assert.equal(r.saldo, 200); assert.equal(r.proxima.id, "b");
  const cal = I.calendario(2026, 10, pr);
  assert.equal(cal[0][3].dia, 1); // 1/10/2026 é quinta-feira
  assert.equal(cal.flat().find(d => d && d.dia === 25).prestacoes.length, 1);
});
