import test from 'node:test';
import assert from 'node:assert/strict';
import * as F from '../rh-folha.js';

const S = iso => F.serie(iso); // série Excel, como o conector devolve as datas

test('período 21–20 e dias base 30', () => {
  assert.deepEqual(F.periodoFolha(2026, 4), {inicio: '2026-03-21', fim: '2026-04-20', referencia: '2026-04-01'});
  assert.deepEqual(F.periodoFolha(2026, 1), {inicio: '2025-12-21', fim: '2026-01-20', referencia: '2026-01-01'});
  assert.equal(F.diasSalario30('2026-03-21', '2026-04-20'), 30);              // 31 dias de calendário -> 30
  assert.equal(F.diasSalario30('2026-01-21', '2026-02-20', '2026-02-11'), 10); // admissão a 11/02
  assert.equal(F.diasSalario30('2026-03-21', '2026-04-20', '', '2026-03-25'), 5); // saída a 25/03
  assert.equal(F.diasSalario30('2026-03-21', '2026-04-20', '2026-05-01'), 0);
});

test('adenda activa e vigente prevalece; anulada e futura não', () => {
  const ad = [
    {Codigo: '12', Empresa: 'A', Campo_Alterado: 'Vencimento Base', Estado_Adenda: 'Activa', Data_Inicio_Vigencia: S('2026-01-01'), Valor_Novo: '20000'},
    {Codigo: '0012', Empresa: 'A', Campo_Alterado: 'Vencimento Base', Estado_Adenda: 'Activa', Data_Inicio_Vigencia: S('2026-04-15'), Valor_Novo: '25000'},
    {Codigo: '0012', Empresa: 'A', Campo_Alterado: 'Vencimento Base', Estado_Adenda: 'Anulada', Data_Inicio_Vigencia: S('2026-04-16'), Valor_Novo: '99999'},
    {Codigo: '0012', Empresa: 'B', Campo_Alterado: 'Vencimento Base', Estado_Adenda: 'Activa', Data_Inicio_Vigencia: S('2026-04-17'), Valor_Novo: '1'},
    {Codigo: '0012', Empresa: 'A', Campo_Alterado: 'Vencimento Base', Estado_Adenda: 'Activa', Data_Inicio_Vigencia: S('2026-05-01'), Valor_Novo: '30000'}
  ];
  assert.equal(F.adendaVigente(ad, 'A', '0012', 'vencimento base', '2026-04-20').valor, '25000');
  assert.equal(F.adendaVigente(ad, 'A', '0012', 'Vencimento Base', '2026-03-20').valor, '20000');
  assert.equal(F.adendaVigente(ad, 'A', '0012', 'Estado', '2026-04-20'), null);
});

test('contrato no período: admissão futura, saída, prazo e inactivo', () => {
  const per = F.periodoFolha(2026, 4);
  const base = {Codigo: '1', Empresa: 'A', TipoContrato: 'Efectivo', Estado: 'Activo', 'Data Admissao': S('2020-01-01')};
  assert.equal(F.contratoNoPeriodo(base, [], per).vigente, true);
  assert.match(F.contratoNoPeriodo({...base, 'Data Admissao': S('2026-04-21')}, [], per).motivo, /Admissão posterior/);
  assert.match(F.contratoNoPeriodo({...base, 'Data de Saída': S('2026-03-20')}, [], per).motivo, /cessado/);
  const saiu = F.contratoNoPeriodo({...base, Estado: 'Inactivo', 'Data de Saída': S('2026-03-25')}, [], per);
  assert.equal(saiu.vigente, true); assert.equal(saiu.fimContrato, '2026-03-25');
  assert.match(F.contratoNoPeriodo({...base, Estado: 'Inactivo'}, [], per).motivo, /inactivo/);
  const prazo = [{Codigo: '1', Empresa: 'A', Campo_Alterado: 'Data Fim Contrato', Estado_Adenda: 'Activa', Data_Inicio_Vigencia: S('2025-01-01'), Valor_Novo: S('2026-03-01')}];
  assert.match(F.contratoNoPeriodo({...base, TipoContrato: 'Prazo certo'}, prazo, per).motivo, /Prazo/);
  assert.equal(F.contratoNoPeriodo(base, prazo, per).vigente, true); // indeterminado ignora data fim
});

test('tarifas das Settings: vigência, inactivas e duplicadas', () => {
  const p = [
    {RegimePagamento: 'Hora', Categoria: 'Motorista', Valor_Unitario: '85', Unidade: 'MT/hora', Vigencia_Inicio: S('2026-01-01'), Activo: 'Sim'},
    {RegimePagamento: 'Hora', Categoria: 'Motorista', Valor_Unitario: '70', Vigencia_Inicio: S('2025-01-01'), Vigencia_Fim: S('2025-12-31'), Activo: 'Sim'},
    {RegimePagamento: 'Diario', Categoria: 'Guarda', Valor_Unitario: '500', Vigencia_Inicio: S('2026-01-01'), Activo: 'Não'}
  ];
  assert.equal(F.tarifaVigente(p, 'Hora', 'motorista', '2026-04-20').valor, 85);
  assert.equal(F.tarifaVigente(p, 'Diário', 'Guarda', '2026-04-20').valor, 0);
  assert.match(F.tarifaVigente([...p, p[0]], 'Hora', 'Motorista', '2026-04-20').erro, /Várias/);
});

test('novo mês: critérios, cópia do mês anterior e bloqueios', () => {
  const func = [
    {Codigo: '1', Nome: 'Mensal Antigo', Empresa: 'A', Departamento: 'Adm', Profissao: 'Técnico', Estado: 'Activo', RegimePagamento: 'Mensal', 'Vencimento Base': '17000', TipoContrato: 'Efectivo', 'Data Admissao': S('2020-01-01'), TotalDependentes: '2'},
    {Codigo: '2', Nome: 'Hora Novo', Empresa: 'A', Departamento: 'Ops', Profissao: 'Motorista', Estado: 'Activo', RegimePagamento: 'Hora', TipoContrato: 'Prazo', 'Data Admissao': S('2026-04-01'), TotalDependentes: '1'},
    {Codigo: '3', Nome: 'Saiu', Empresa: 'A', Estado: 'Inactivo', RegimePagamento: 'Mensal', 'Vencimento Base': '9000'},
    {Codigo: '1', Nome: 'Mesmo código outra empresa', Empresa: 'B', Profissao: 'Técnico', Estado: 'Activo', RegimePagamento: 'Mensal', 'Vencimento Base': '12000', 'Data Admissao': S('2026-04-11')}
  ];
  const parametros = [{RegimePagamento: 'Hora', Categoria: 'Motorista', Valor_Unitario: '85', Vigencia_Inicio: S('2026-01-01'), Activo: 'Sim'}];
  const anterior = [{Codigo: '0001', Empresa: 'A', Ano: '2026', 'Mês': 'Março', 'Subsídios': '500', 'Adiantamentos': '1000', Canalizar: 'Sim', TotalDependentes: '3', 'Observações': 'x'}];
  const r = F.prepararNovoMes({ano: 2026, mes: 4, funcionarios: func, parametros, lancamentosAnterior: anterior, utilizador: 'rh@sir.co', agora: new Date('2026-04-22T08:00:00Z')});
  assert.deepEqual(r.erros, []);
  assert.equal(r.linhas.length, 3);
  assert.deepEqual(r.excluidos.map(e => e.codigo), ['0003']);
  const a1 = r.linhas.find(l => l.Empresa === 'A' && l.Codigo === '0001');
  assert.equal(a1['Origem_Lancamento'], 'Cópia Mês Anterior');
  assert.equal(a1['Subsídios'], '500'); assert.equal(a1.TotalDependentes, '3'); assert.equal(a1.Canalizar, 'Sim');
  assert.equal(a1['Dias Trabalhados'], '30'); assert.equal(a1['Valor Unitário'], '17000'); assert.equal(a1['Horas Trabalhadas'], '');
  assert.equal(a1['Mês'], 'Abril'); assert.equal(a1.ID_Lancamento, '2026-04-A-0001'); assert.equal(a1['Estado do Lançamento'], 'Pendente');
  const h = r.linhas.find(l => l.Codigo === '0002');
  assert.equal(h['Valor Unitário'], '85'); assert.equal(h['Horas Trabalhadas'], '0'); assert.equal(h['Dias Trabalhados'], undefined);
  assert.equal(h['Origem_Lancamento'], 'Novo Colaborador'); assert.equal(h.TotalDependentes, '1'); assert.equal(h['Bónus'], '0');
  const b1 = r.linhas.find(l => l.Empresa === 'B');
  assert.equal(b1['Dias Trabalhados'], '10'); // admitido a 11/04
  assert.ok(r.linhas.every(l => F.FORMULAS_LANCAMENTO.every(k => !(k in l))));
  // Bloqueios: mês já existe, tarifa em falta, código repetido.
  const sem = F.prepararNovoMes({ano: 2026, mes: 4, funcionarios: [...func, {...func[0]}], parametros: [], lancamentosMes: [{Ano: '2026', 'Mês': 'Abril'}]});
  assert.ok(sem.erros.some(e => /Já existem/.test(e)));
  assert.ok(sem.erros.some(e => /repetido/.test(e)));
  assert.ok(sem.erros.some(e => /Valor unitário não encontrado para A \| 0002/.test(e)));
});

test('folha: totais efectivos e mapeamento do VBA', () => {
  const l = [{Codigo: '1', Nome: 'X', Ano: '2026', 'Mês': 'Abril', Empresa: 'A', Departamento: 'Adm', 'Vencimento Base': '10000', Vencimento: '9000', 'Subsídios': '300', 'Bónus': '200', 'Total de Acréscimos': '400',
    'INSS Colaborador': '300', IRPS: '100', Adiantamentos: '50', 'Total de Descontos': '20', 'Salario Liquido': '9550'},
  {Codigo: '2', Ano: '2026', 'Mês': 'Março', Empresa: 'A'}];
  const f = F.linhasFolha(l, 2026, 4, {empresa: 'TODOS'});
  assert.equal(f.length, 1);
  assert.equal(f[0].acrescimos, 500); assert.equal(f[0].outrasRem, 200); assert.equal(f[0].totalRem, 10500);
  assert.equal(f[0].totalDesc, 450); assert.equal(f[0].outrosDesc, 50); assert.equal(f[0].vencimento, 9000); assert.equal(f[0].valorFalta, null);
  assert.equal(F.linhasFolha(l, 2026, 4, {departamento: 'Ops'}).length, 0);
  assert.equal(F.linhasFolha(l, 2026, 4, {codigo: '0001'}).length, 1);
});

test('relatório comparativo: grupos, variação, peso e comentários', () => {
  const l = [
    {Codigo: '1', Nome: 'Ana', Empresa: 'A', Departamento: 'Adm', Ano: '2026', 'Mês': 'Março', 'Salario Liquido': '1000'},
    {Codigo: '2', Nome: 'Rui', Empresa: 'A', Departamento: 'Adm', Ano: '2026', 'Mês': 'Março', 'Salario Liquido': '500'},
    {Codigo: '1', Nome: 'Ana', Empresa: 'A', Departamento: 'Ops', Ano: '2026', 'Mês': 'Abril', 'Salario Liquido': '1200', 'Faltas Injustificadas': '1', 'Total de Descontos': '50'},
    {Codigo: '3', Nome: 'Eva', Empresa: 'A', Departamento: 'Adm', Ano: '2026', 'Mês': 'Abril', 'Salario Liquido': '800', 'Horas Trabalhadas': '10', 'Valor Unitário': '80'}
  ];
  const r = F.relatorioComparativo(l, 2026, 4);
  assert.deepEqual(r.anterior, {ano: 2026, mes: 3});
  const adm = r.grupos.find(g => g.departamento === 'Adm'), ops = r.grupos.find(g => g.departamento === 'Ops');
  assert.equal(adm.valorAnterior, 1500); assert.equal(adm.valorActual, 800); assert.equal(adm.variacao, -700);
  assert.equal(r.totalActual, 2000); assert.equal(ops.peso, 0.6);
  assert.match(ops.comentario, /Ana transferido de Adm para Ops/);
  assert.match(ops.comentario, /Ana: 1 falta\(s\) \(-50,00 MT\)/);
  assert.match(adm.comentario, /Entrada de Eva/); assert.match(adm.comentario, /Saída de Rui/);
  assert.match(adm.comentario, /Aumento global de 10 horas \(\+800,00 MT\)/);
});

test('novo mês: contrato e adendas no período 21–20 (admissão, saída, renovação, vencimento)', () => {
  const S = d => F.serie(d);
  const ad = (Codigo, Campo_Alterado, Valor_Novo, ini, extra = {}) => ({Codigo, Empresa: 'A', Campo_Alterado, Valor_Novo, Estado_Adenda: 'Activa', Data_Inicio_Vigencia: S(ini), ...extra});
  const base = {Empresa: 'A', Estado: 'Activo', RegimePagamento: 'Mensal', 'Vencimento Base': '20000', TipoContrato: 'Efectivo', 'Data Admissao': S('2020-01-01'), Profissao: 'Técnico'};
  const funcionarios = [
    {...base, Codigo: '10', Nome: 'Admitido a meio', 'Data Admissao': S('2026-04-06')},            // 06/04 a 20/04 = 15 dias
    {...base, Codigo: '11', Nome: 'Saída por adenda'},                                              // saída 31/03 por adenda = 11 dias
    {...base, Codigo: '12', Nome: 'Saiu antes do período', Estado: 'Inactivo', 'Data de Saída': S('2026-03-20')},
    {...base, Codigo: '13', Nome: 'Admitido depois', 'Data Admissao': S('2026-04-21')},
    {...base, Codigo: '14', Nome: 'Prazo renovado por adenda', TipoContrato: 'Prazo', 'Data Fim Contrato': S('2026-03-31')},
    {...base, Codigo: '15', Nome: 'Prazo terminado sem renovação', TipoContrato: 'Prazo', 'Data Fim Contrato': S('2026-03-31')},
    {...base, Codigo: '16', Nome: 'Aumento por adenda'},
    {...base, Codigo: '17', Nome: 'Adenda anulada ignorada'}
  ];
  const adendas = [
    ad('11', 'Data de Saída', S('2026-03-31'), '2026-03-25'),
    ad('14', 'Data Fim Contrato', S('2026-12-31'), '2026-03-15'),
    ad('16', 'Vencimento Base', '26000', '2026-04-01'),
    ad('17', 'Vencimento Base', '99999', '2026-04-01', {Estado_Adenda: 'Anulada'})
  ];
  const r = F.prepararNovoMes({ano: 2026, mes: 4, funcionarios, adendas});
  assert.deepEqual(r.erros, []);
  const L = c => r.linhas.find(l => l.Codigo === c);
  assert.equal(L('0010')['Dias Trabalhados'], '15');
  assert.equal(L('0011')['Dias Trabalhados'], '11');
  assert.equal(L('0011')['Data Fim Contrato'], S('2026-03-31'));
  assert.equal(L('0012'), undefined);
  assert.equal(L('0013'), undefined);
  assert.equal(L('0014')['Dias Trabalhados'], '30');
  assert.equal(L('0015')['Dias Trabalhados'], '11');            // prazo 31/03 dentro do período: 21/03 a 31/03
  assert.equal(L('0016')['Valor Unitário'], '26000');
  assert.equal(L('0017')['Valor Unitário'], '20000');
  const exc = Object.fromEntries(r.excluidos.map(e => [e.codigo, e.motivo]));
  assert.match(exc['0012'], /cessado/);
  assert.match(exc['0013'], /Admissão posterior/);
});

test('adendas antigas (cabeçalho + detalhes) contam como no VBA', () => {
  const S = d => F.serie(d);
  const cab = [{ID_Adenda: 'AD-1', Codigo: '20', Empresa: 'A', Estado_Adenda: 'Aplicada', Data_Inicio_Vigencia: S('2026-03-10')},
               {ID_Adenda: 'AD-2', Codigo: '20', Empresa: 'A', Estado_Adenda: 'Anulada', Data_Inicio_Vigencia: S('2026-03-12')}];
  const det = [{ID_Adenda: 'AD-1', Campo_Alterado: 'Data de Saída', Valor_Novo: S('2026-04-05')},
               {ID_Adenda: 'AD-2', Campo_Alterado: 'Data de Saída', Valor_Novo: S('2026-03-22')}];
  const todas = F.unirAdendas([], cab, det);
  assert.equal(todas.length, 2);
  const f = {Codigo: '20', Empresa: 'A', Nome: 'X', Estado: 'Activo', RegimePagamento: 'Mensal', 'Vencimento Base': '15000', TipoContrato: 'Efectivo', 'Data Admissao': S('2019-01-01')};
  const r = F.prepararNovoMes({ano: 2026, mes: 4, funcionarios: [f], adendas: todas});
  assert.equal(r.linhas[0]['Dias Trabalhados'], '16');   // 21/03 a 05/04; a adenda anulada é ignorada
});
