import test from 'node:test';
import assert from 'node:assert/strict';
import {mesesAte, linhasFuncionarios, linhasLancamentos, MESES_HISTORICO} from '../painel.js';

test('mesesAte devolve os 12 meses até ao período, atravessando o ano', () => {
  const m = mesesAte({ano: 2026, mes: 3});
  assert.equal(m.length, MESES_HISTORICO);
  assert.deepEqual(m[0], {ano: 2025, mes: 4});
  assert.deepEqual(m[11], {ano: 2026, mes: 3});
});

test('só seguem as colunas que o dashboard usa (nada de dados bancários ou documentos)', () => {
  const [f] = linhasFuncionarios([{Empresa: 'SIR', Codigo: '0001', Sexo: 'F', 'Data de Nascimento': '30000', NUIT: '123', 'Conta Bancaria': 'X', 'Estado Reforma': 'Reformado'}]);
  assert.equal(f.Sexo, 'F'); assert.equal(f['Data Nascimento'], '30000'); assert.equal(f['Estado de Reforma'], 'Reformado');
  assert.ok(!('NUIT' in f) && !('Conta Bancaria' in f));
  const [l] = linhasLancamentos([{Empresa: 'SIR', Codigo: '0001', Ano: '2026', Mes: 'Março', IRPS: '10', Observações: 'privado'}]);
  assert.equal(l['Mês'], 'Março'); assert.equal(l.IRPS, '10'); assert.ok(!('Observações' in l));
});
