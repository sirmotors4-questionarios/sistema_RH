import test from 'node:test';
import assert from 'node:assert/strict';
import * as R from '../relatorios-dados.js';

const L = (cod, emp, dep, mes, venc, extra = {}) => ({Codigo: cod, Nome: 'P' + cod, Empresa: emp, Departamento: dep, Ano: '2026', 'Mês': mes,
  'Vencimento Base': String(venc), Vencimento: String(venc), 'Salario Liquido': String(venc * 0.9), 'INSS Colaborador': String(venc * 0.03),
  'INSS Empresa': String(venc * 0.04), IRPS: String(venc * 0.05), 'Total de Descontos': String(venc * 0.1), 'Total de Acréscimos': '0',
  'Estado do Lançamento': 'Pendente', ...extra});
const marco = [L('1', 'SIR', 'Oficina', 'Março', 10000), L('2', 'SIR', 'Oficina', 'Março', 20000), L('3', 'Metrobus', 'Transporte', 'Março', 15000)];
const abril = [L('1', 'SIR', 'Oficina', 'Abril', 10000, {'Faltas Injustificadas': '2', 'Horas Extras': '5'}), L('2', 'SIR', 'Oficina', 'Abril', 22000, {'Estado do Lançamento': 'Finalizado'}),
  L('4', 'Metrobus', 'Transporte', 'Abril', 8000, {'Faltas Justificadas': '1'})];
const func = [
  {Codigo: '1', Nome: 'P1', Empresa: 'SIR', Departamento: 'Oficina', Sexo: 'M', 'Data Nascimento': '1960-01-01', 'Data Admissao': '2010-05-01', Estado: 'Activo', 'Estado de Reforma': 'Não', TipoContrato: 'Efectivo', RegimePagamento: 'Mensal'},
  {Codigo: '2', Nome: 'P2', Empresa: 'SIR', Departamento: 'Oficina', Sexo: 'F', 'Data Nascimento': '1968-03-01', 'Data Admissao': '2026-02-10', Estado: 'Activo', 'Estado de Reforma': 'Reformada'},
  {Codigo: '3', Nome: 'P3', Empresa: 'Metrobus', Departamento: 'Transporte', Sexo: 'Masculino', 'Data Admissao': '2015-01-01', 'Data Saida': '2026-03-31', Estado: 'Inactivo', 'Motivo de saída': 'Fim de contrato'},
  {Codigo: '4', Nome: 'P4', Empresa: 'Metrobus', Departamento: 'Transporte', Sexo: 'F', 'Data Nascimento': '2000-06-01', 'Data Admissao': '2026-04-01', Estado: 'Activo'}];
const hoje = new Date('2026-10-27T12:00:00Z');
const ok = r => { assert.ok(!r.vazio, r.vazio); assert.ok(r.kpis.length >= 4); assert.ok(r.secoes.length >= 2);
  for (const s of r.secoes) for (const tb of [s.tabela, ...(s.subtabelas || [])].filter(Boolean)) for (const l of [...tb.linhas, ...(tb.total ? [tb.total] : [])]) assert.equal(l.length, tb.cols.length, r.id + ' ' + s.titulo); };

test('processamento salarial: totais, variação e entradas/saídas da folha', () => {
  const r = R.relatorioSalarial({lancMes: abril, lancAnterior: marco, ano: 2026, mes: 4});
  ok(r);
  assert.equal(r.kpis[0].valor, '3');
  assert.match(r.kpis[1].valor, /^40\s?000,00 MT$/);
  assert.match(r.secoes[0].paragrafos[1], /diminuiu 5\s?000,00 MT/);
  const es = r.secoes.find(s => /Entradas e saídas/.test(s.titulo));
  assert.equal(es.tabela.linhas.length, 2); // 4 novo, 3 saiu
  assert.ok(R.relatorioSalarial({lancMes: abril, lancAnterior: [], ano: 2026, mes: 5}).vazio);
});
test('encargos: INSS e IRPS por empresa batem com o mapa nominal', () => {
  const r = R.relatorioEncargos({lancMes: abril, ano: 2026, mes: 4}); ok(r);
  const porEmp = r.secoes[1].tabela.total, nominal = r.secoes[2].tabela.total;
  assert.equal(porEmp[3], nominal[4]); assert.equal(porEmp[6], nominal[6]);
});
test('quadro de pessoal: idade de reforma (H60/M55) e reformados', () => {
  const r = R.relatorioEfectivo({funcionarios: func, hoje}); ok(r);
  assert.equal(r.kpis[0].valor, '3');
  assert.equal(r.kpis[4].valor, '2'); // P1 (66, M) e P2 (58, F)
  assert.equal(r.kpis[5].valor, '1'); // P2 reformada
  const lista = r.secoes[r.secoes.length - 1].tabela.linhas;
  assert.equal(lista.length, 2);
});
test('movimentação: admissões e saídas do ano', () => {
  const r = R.relatorioMovimentacao({funcionarios: func, ano: 2026, hoje}); ok(r);
  assert.equal(r.kpis[0].valor, '2'); assert.equal(r.kpis[1].valor, '1');
  assert.equal(r.secoes.find(s => /Saídas/.test(s.titulo)).tabela.linhas[0][6], 'Fim de contrato');
});
test('assiduidade e evolução', () => {
  const a = R.relatorioAssiduidade({lancMes: abril, ano: 2026, mes: 4}); ok(a);
  assert.equal(a.kpis[1].valor, '1 dias'); assert.equal(a.kpis[2].valor, '2 dias'); assert.equal(a.kpis[4].valor, '5 h');
  const e = R.relatorioEvolucao({lancPorMes: [{ano: 2026, mes: 2, linhas: []}, {ano: 2026, mes: 3, linhas: marco}, {ano: 2026, mes: 4, linhas: abril}]}); ok(e);
  assert.equal(e.secoes[1].tabela.linhas.length, 3);
  assert.equal(e.secoes[1].tabela.linhas[0][1], '—');
});
test('filtro por empresa', () => {
  const r = R.relatorioSalarial({lancMes: abril, lancAnterior: marco, ano: 2026, mes: 4, empresa: 'SIR'}); ok(r);
  assert.equal(r.kpis[0].valor, '2');
  assert.ok(!r.secoes.some(s => /Por empresa/.test(s.titulo)));
});
