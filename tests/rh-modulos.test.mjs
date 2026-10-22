import test from 'node:test';
import assert from 'node:assert/strict';
import * as M from '../rh-modulos.js';

// Linhas fictícias com os cabeçalhos reais do Excel (espaços e acentos incluídos).
const funcionarios = [
  {'@odata.etag': '', ItemInternalId: '1', Codigo: 'T001', Nome: 'Teste Um', Empresa: 'A', Departamento: 'Adm', Estado: 'Activo', RegimePagamento: 'Mensal', TipoContrato: 'Efectivo', 'Data Admissao': '45292'},
  {ItemInternalId: '2', Codigo: 'T002', Nome: 'Teste Dois', Empresa: 'B', Departamento: 'Ops', Estado: 'Inactivo', RegimePagamento: 'Diário', TipoContrato: 'Prazo', 'Data Admissao': ''},
  {ItemInternalId: '3', Codigo: '', Nome: '', Empresa: '', Estado: ''}
];

test('colaboradores: lê cabeçalhos com espaços, ignora linhas vazias e converte datas do Excel', () => {
  const l = M.colaboradores(funcionarios);
  assert.equal(l.length, 2);
  assert.equal(l[0].admissao, '2024-01-01');
  assert.equal(l[1].regime, 'Diário');
  assert.equal(M.filtrarColaboradores(l, 'dois').length, 1);
  assert.equal(M.filtrarColaboradores(l, '', 'activo').length, 1);
  assert.equal(M.filtrarColaboradores(l, 'adm', 'Inactivo').length, 0);
});

test('campo tolera "_" em vez de espaço e diferenças de acentos', () => {
  assert.equal(M.campo({Data_Admissao: 'x'}, 'Data Admissao'), 'x');
  assert.equal(M.campo({'Estado do Lancamento': 'Pendente'}, 'Estado do Lançamento'), 'Pendente');
  assert.equal(M.campo({}, 'Nada'), '');
});

test('números e datas em formatos do Excel', () => {
  assert.equal(M.numero('17,000.50'), 17000.5);
  assert.equal(M.numero('17.000,50'), 17000.5);
  assert.equal(M.numero('1 250 MT'), 1250);
  assert.equal(M.numero(''), null);
  assert.equal(M.numero('abc'), null);
  assert.equal(M.data('2026-10-22T00:00:00Z'), '2026-10-22');
  assert.equal(M.data('5/3/2026'), '2026-03-05');
  assert.equal(M.formatarData('45292'), '01/01/2024');
  assert.equal(M.mes('Março'), 3);
  assert.equal(M.mes('10'), 10);
  assert.equal(M.mes('Out'), 10);
  assert.equal(M.mes(''), null);
});

test('lançamentos filtrados por ano e mês, totais e períodos', () => {
  const linhas = [
    {Ano: '2026', 'Mês': '9', Codigo: 'T001', 'Salario Liquido': '1000', 'Total de Acréscimos': '100', 'Estado do Lançamento': 'Finalizado'},
    {Ano: '2026', 'Mês': 'Outubro', Codigo: 'T001', 'Salario Liquido': '1,200.00', 'Estado do Lançamento': 'Pendente'},
    {Ano: '2026', 'Mês': '10', Codigo: 'T002', 'Salario Liquido': '800', 'Estado do Lançamento': 'Pendente'}
  ];
  const out = M.lancamentos(linhas, 2026, 10);
  assert.equal(out.length, 2);
  assert.equal(M.totais(out, 'liquido'), 2000);
  assert.deepEqual(M.periodos(linhas)[0], {ano: 2026, mes: 10});
  assert.deepEqual(M.contagem(out, l => l.estado), [['Pendente', 2]]);
});

test('parâmetros vigentes respeitam datas e inactivos; adendas ordenadas pela vigência', () => {
  const p = [
    {RegimePagamento: 'Diário', Valor_Unitario: '500', Vigencia_Inicio: '2026-01-01', Vigencia_Fim: '', Activo: 'Sim'},
    {RegimePagamento: 'Diário', Valor_Unitario: '450', Vigencia_Inicio: '2025-01-01', Vigencia_Fim: '2025-12-31', Activo: 'Sim'},
    {RegimePagamento: 'Hora', Valor_Unitario: '80', Activo: 'Não'}
  ];
  const v = M.parametrosVigentes(p, '2026-10-22');
  assert.equal(v.length, 1);
  assert.equal(v[0].valor, '500');
  const a = M.adendas([{Numero_Adenda: '1', Data_Inicio_Vigencia: '2025-01-01'}, {Numero_Adenda: '2', Data_Inicio_Vigencia: '2026-05-01'}]);
  assert.equal(a[0].numero, '2');
});

test('férias: colunas sem campos técnicos; feriados ordenados', () => {
  assert.deepEqual(M.colunas([{ItemInternalId: 1, '@odata.etag': 'x', Codigo: 'T', Inicio: '1'}]), ['Codigo', 'Inicio']);
  assert.deepEqual(M.colunas([]), []);
  const f = M.feriados([{Dia: '25', Mes: '6', Feriado: 'B'}, {Dia: '1', Mes: '1', Feriado: 'A'}, {Dia: '', Mes: '', Feriado: ''}]);
  assert.deepEqual(f.map(x => x.nome), ['A', 'B']);
});
