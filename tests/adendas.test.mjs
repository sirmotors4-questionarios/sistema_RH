import test from "node:test";
import assert from "node:assert/strict";
import {agrupar, validarAdenda} from "../adendas-ui.js";

test("agrupa por colaborador + número e detecta pendentes e anuladas", () => {
  const l = [
    {id: "AD-2026-0001", empresa: "SIR", codigo: "0225", numero: "1", estado: "Activa", aplicada: "Sim", campo: "Vencimento Base"},
    {id: "AD-2026-0002", empresa: "SIR", codigo: "225", numero: "1", estado: "Activa", aplicada: "Não", campo: "Profissao"},
    {id: "AD-2026-0003", empresa: "SIR", codigo: "0226", numero: "1", estado: "Anulada", aplicada: "Sim", campo: "Departamento"}];
  const g = agrupar(l);
  assert.equal(g.length, 2);
  assert.deepEqual(g[0].ids, ["AD-2026-0001", "AD-2026-0002"]);
  assert.equal(g[0].pendente, true);
  assert.equal(g[1].anulada, true);
});

test("validação da nova adenda", () => {
  const base = {colaborador: {}, motivo: "Alteração Salarial", dataAdenda: "2026-10-29", inicio: "2026-11-01", fim: ""};
  assert.deepEqual(validarAdenda({...base, alteracoes: [{campo: "Vencimento Base", novo: "16500", anterior: "15000"}]}), []);
  assert.match(validarAdenda({...base, alteracoes: [{campo: "Vencimento Base", novo: "15000", anterior: "15000"}]})[0], /igual/);
  assert.match(validarAdenda({...base, alteracoes: [{campo: "Vencimento Base", novo: "abc", anterior: "1"}]})[0], /inválido/);
  assert.match(validarAdenda({...base, fim: "2026-10-01", alteracoes: [{campo: "Profissao", novo: "X"}]})[0], /fim/);
  assert.ok(validarAdenda({...base, alteracoes: []}).some(e => /pelo menos/.test(e)));
});
