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

test('chave Empresa + Código igual ao VBA', () => {
  assert.equal(M.normalizarCodigo(' 12 '), '0012');
  assert.equal(M.normalizarCodigo('00123'), '0123');
  assert.equal(M.normalizarCodigo('12345'), '12345');
  assert.equal(M.normalizarCodigo('RH-7'), 'RH-7');
  assert.equal(M.chaveRH(' Sir Motors ', 12), 'SIR MOTORS|0012');
  const linhas = [{Codigo: '12', Empresa: 'A', Nome: 'Um'}, {Codigo: '0012', Empresa: 'B', Nome: 'Dois'}];
  assert.equal(M.localizarColaborador(linhas, 'b', '12').Nome, 'Dois');
  assert.equal(M.localizarColaborador(linhas, 'C', '12'), null);
});

test('filtros em cascata Empresa -> Departamento', () => {
  const l = M.colaboradores(funcionarios);
  assert.deepEqual(M.opcoesFiltro(l), {empresas: ['A', 'B'], departamentos: ['Adm', 'Ops']});
  assert.deepEqual(M.opcoesFiltro(l, 'b').departamentos, ['Ops']);
  assert.equal(M.filtrarColaboradores(l, '', '', {empresa: 'A'}).length, 1);
  assert.equal(M.filtrarColaboradores(l, '', '', {empresa: 'A', departamento: 'Ops'}).length, 0);
});

test('ficha: 6 separadores, colunas reais e listas distintas', () => {
  assert.equal(M.SEPARADORES.length, 6);
  assert.deepEqual(M.definicao('Sexo').opcoes, ['Masculino', 'Feminino']);
  assert.ok(M.CAMPOS_COLABORADOR.every(c => M.SEPARADORES.some(s => s.id === c.sep)));
  const cab = ['Codigo', 'Empresa', 'Nome', 'Data de Admissão', 'Profissão', 'NIB_Conta'];
  assert.equal(M.colunaReal(cab, M.definicao('Data Admissao')), 'Data de Admissão');
  assert.equal(M.colunaReal(cab, M.definicao('Profissao')), 'Profissão');
  assert.equal(M.colunaReal(cab, M.definicao('NIB/Conta')), 'NIB_Conta');
  assert.equal(M.colunaReal(cab, M.definicao('Distrito')), '');
  assert.deepEqual(M.valoresDistintos([{Departamento: 'ops'}, {Departamento: 'Adm'}, {Departamento: 'Ops'}, {}], M.definicao('Departamento')), ['Adm', 'ops']);
});

test('validação da ficha: obrigatórios, duplicado na mesma empresa, regime e datas', () => {
  const existentes = [{Codigo: '0009', Empresa: 'A'}];
  let r = M.validarColaborador({}, 'criar');
  assert.ok(['Código', 'Empresa', 'Nome completo'].every(n => r.erros.some(e => e.startsWith(n))));
  r = M.validarColaborador({Codigo: '9', Empresa: 'a', Nome: 'X'}, 'criar', {existentes});
  assert.ok(r.erros.some(e => /Já existe o código 0009/.test(e)));
  assert.equal(r.dados.Codigo, '0009');
  r = M.validarColaborador({Codigo: '9', Empresa: 'B', Nome: 'X', RegimePagamento: 'Hora', 'Vencimento Base': '5000', Sexo: 'M'}, 'criar', {existentes});
  assert.equal(r.dados['Vencimento Base'], '');
  assert.ok(r.erros.some(e => /Sexo/.test(e)));
  r = M.validarColaborador({Codigo: 'T1', Empresa: 'B', Nome: 'X', RegimePagamento: 'Mensal'}, 'criar');
  assert.ok(r.erros.some(e => /Vencimento base é obrigatório/.test(e)));
  r = M.validarColaborador({Codigo: 'T1', Empresa: 'B', Nome: 'X', 'Data Admissao': '2026-05-01', 'Data de Saída': '2026-01-01'}, 'criar');
  assert.ok(r.erros.some(e => /saída/.test(e)));
  // Em edição, um NUIT antigo fora do formato não impede gravar outro campo.
  const original = {Codigo: 'T1', Empresa: 'B', Nome: 'X', NUIT: '123'};
  r = M.validarColaborador({...original, Departamento: 'Ops'}, 'editar', {original});
  assert.deepEqual(r.erros, []);
  r = M.validarColaborador({...original, NUIT: '1234'}, 'editar', {original});
  assert.ok(r.erros.some(e => /NUIT/.test(e)));
});

test('prepararGravacao: alvo fixo, só alterações, cabeçalhos reais, datas em série e auditoria', () => {
  const cab = ['Codigo', 'Empresa', 'Nome', 'Departamento', 'Data de Admissão', 'Vencimento Base', 'Data Ultima Alteracao', 'Utilizador Alteracao'];
  const original = {Codigo: '0012', Empresa: 'A', Nome: 'X', Departamento: 'Adm', 'Data Admissao': '2024-01-01'};
  const dados = {...original, Empresa: 'Outra', Departamento: 'Ops', 'Data Admissao': '2024-01-02', Distrito: 'Maputo'};
  const agora = new Date('2026-10-26T10:00:00Z');
  const r = M.prepararGravacao('editar', dados, {original, cabecalhos: cab, utilizador: 'rh@sir.co', agora});
  assert.deepEqual(r.alvo, {Empresa: 'A', Codigo: '0012'});
  assert.equal(r.chave, 'A|0012');
  assert.equal(r.dados.Empresa, undefined);
  assert.equal(r.dados.Departamento, 'Ops');
  assert.equal(r.dados['Data de Admissão'], '45293');
  assert.equal(r.dados['Utilizador Alteracao'], 'rh@sir.co');
  assert.match(r.dados['Data Ultima Alteracao'], /^4\d{4}\.\d+$/);
  assert.deepEqual(r.ignorados, ['Distrito']);
  assert.equal(M.prepararGravacao('editar', original, {original, cabecalhos: cab}).alteracoes, 0);
  const c = M.prepararGravacao('criar', {Codigo: '0001', Empresa: 'B', Nome: 'Y', Departamento: ''}, {cabecalhos: cab});
  assert.deepEqual(c.dados, {Codigo: '0001', Empresa: 'B', Nome: 'Y', 'Data Ultima Alteracao': c.dados['Data Ultima Alteracao'], 'Utilizador Alteracao': ''});
});

test('indicador de preenchimento usa a escala do VBA', () => {
  const v = {Nome: 'X', Sexo: 'Masculino', 'Data Nascimento': '2000-01-01', EstadoCivil: 'Solteiro', Nacionalidade: 'Moçambicana', Naturalidade: '', TotalDependentes: '0', 'Nome Abreviado': ''};
  const p = M.preenchimento(v, 'pessoais');
  assert.equal(p.total, 8); assert.equal(p.pct, 75); assert.equal(p.nivel, 'Médio');
  assert.equal(M.preenchimento({}, 'bancarios').nivel, 'Baixo');
});
