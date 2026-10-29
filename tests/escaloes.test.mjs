import test from 'node:test';
import assert from 'node:assert/strict';
import * as E from '../escaloes.js';

const linhas = [
  {Empresa: 'SIR Motors', Criterio: 'Salario Liquido', Escalao: 'B', Valor_Inicial: 20000, Ordem: 2},
  {Empresa: 'SIR Motors', Criterio: 'Salario Liquido', Escalao: 'A', Valor_Inicial: 0, Ordem: 1},
  {Empresa: 'SIR Motors', Criterio: 'Salario Liquido', Escalao: 'C', Valor_Inicial: 50000, Ordem: 3},
  {Empresa: 'Todas', Criterio: 'Vencimento Base', Escalao: 'Único', Valor_Inicial: 0, Ordem: 1},
  {Empresa: '', Escalao: 'X', Valor_Inicial: 1}
];

test('agrupa por empresa e ordena pelo valor inicial', () => {
  const m = E.agrupar(linhas);
  assert.equal(m.size, 2);
  assert.deepEqual(m.get('SIR MOTORS').faixas.map(f => f.escalao), ['A', 'B', 'C']);
});

test('escalão: início inclusive, último sem limite, abaixo do primeiro vazio', () => {
  const f = E.agrupar(linhas).get('SIR MOTORS').faixas;
  assert.equal(E.escalaoDe(f, 0), 'A');
  assert.equal(E.escalaoDe(f, 19999.99), 'A');
  assert.equal(E.escalaoDe(f, 20000), 'B');
  assert.equal(E.escalaoDe(f, '433854,41'), 'C');
  assert.equal(E.escalaoDe(f, -5), '');
  assert.equal(E.escalaoDe(f, ''), '');
});

test('linha usa o critério da sua empresa e cai para "Todas"', () => {
  const m = E.agrupar(linhas);
  assert.equal(E.escalaoLinha(m, {Empresa: 'sir motors', 'Salario Liquido': 25000, 'Vencimento Base': 1}), 'B');
  assert.equal(E.escalaoLinha(m, {Empresa: 'Outra', 'Vencimento Base': 9000}), 'Único');
  assert.equal(E.escalaoLinha(new Map(), {Empresa: 'Outra'}), '');
});

test('intervalos mostram o limite seguinte', () => {
  const iv = E.intervalos([{escalao: 'B', inicio: 10}, {escalao: 'A', inicio: 0}]);
  assert.deepEqual(iv.map(x => x.ate), [10, null]);
});

test('validação', () => {
  assert.deepEqual(E.validar({empresa: 'X', criterio: 'Salario Liquido', faixas: [{escalao: 'A', inicio: 0}]}), []);
  const e = E.validar({empresa: '', criterio: 'Nada', faixas: [{escalao: 'A', inicio: 0}, {escalao: 'a', inicio: 0}, {escalao: '', inicio: 'x'}]});
  assert.ok(e.some(x => /empresa/.test(x)));
  assert.ok(e.some(x => /Critério/.test(x)));
  assert.ok(e.some(x => /repetido/.test(x)));
  assert.ok(e.some(x => /não tem nome/.test(x)));
  assert.ok(e.some(x => /Valor inicial inválido/.test(x)));
});

test('distribuição pelo cadastro', () => {
  const f = [{escalao: 'A', inicio: 1000}, {escalao: 'B', inicio: 5000}];
  const d = E.distribuicao(f, 'Salario Liquido', [{'Salario Liquido': 1200}, {'Salario Liquido': 7000}, {'Salario Liquido': 10}, {}]);
  assert.deepEqual(d, {porEscalao: [{escalao: 'A', total: 1}, {escalao: 'B', total: 1}], semValor: 1, abaixo: 1});
});
