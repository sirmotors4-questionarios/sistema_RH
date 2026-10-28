import test from "node:test";
import assert from "node:assert/strict";
import {contarTiposContrato} from "../tipos-contrato.js";

const f = (TipoContrato, Estado) => ({TipoContrato, Estado});

test("conta tipos do cadastro, sem distinguir maiúsculas/acentos, e activos sem tipo", () => {
  const r = contarTiposContrato([f("Efectivo", "Activo"), f("efectivo", "Inactivo"), f("Prazo certo", "Activo"), f("", "Activo"), f("", "Inactivo")]);
  assert.equal(r.tipos.length, 2);
  assert.deepEqual(r.tipos.map(t => [t.nome, t.activos, t.inactivos, t.total]), [["Efectivo", 1, 1, 2], ["Prazo certo", 1, 0, 1]]);
  assert.equal(r.activos, 3);
  assert.equal(r.semTipo, 1);
});

test("tipos propostos aparecem no fim e não duplicam um tipo existente", () => {
  const r = contarTiposContrato([f("Efectivo", "Activo")], [{nome: "EFECTIVO"}, {nome: "Estágio", regime: "Mensal"}]);
  assert.deepEqual(r.tipos.map(t => [t.nome, t.proposto]), [["Efectivo", false], ["Estágio", true]]);
});
