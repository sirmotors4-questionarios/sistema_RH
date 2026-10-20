import test from 'node:test';
import assert from 'node:assert/strict';
import {evaluateFormula, resolveConfig, validateScales} from '../rh-engine.js';
import {readTables, resolveTables, simulateRecord} from '../rh-data.js';

const common={Estado:'Aprovado',Nivel_Aplicacao:'Geral',Versao:1,Vigencia_Inicio:'2026-01-01'};
const config=(scope,value,fields={})=>({...common,Codigo_Parametro:'TAXA',Tipo_Valor:'Decimal',Valor:value,Nivel_Aplicacao:scope,...fields});
const context={empresa:'Teste',categoria:'Teste',codigo:'T001'};
const at='2026-10-20';
const tables=()=>readTables({tbl_configuracoes_RH:[config('Geral','0.03')],tbl_formulas_RH:[{...common,Codigo_Formula:'TESTE',Expressao:'ROUND(BASE * TAXA, 2)',Validada:true}],tbl_escaloes_RH:[],tbl_auditoria_RH:[]});

test('quatro tabelas são obrigatórias; não aceita resposta parcial',()=>{
 assert.throws(()=>readTables({tbl_configuracoes_RH:[]}),/tbl_escaloes_RH/);
 assert.throws(()=>readTables(null),/inválida/);
});
test('selecção respeita prioridade, vigência e estado',()=>{
 const rows=[config('Geral',1),config('Empresa',2,{Empresa:'Teste'}),config('Categoria',3,{Empresa:'Teste',Categoria:'Teste'}),config('Colaborador',4,{Empresa:'Teste',Codigo_Colaborador:'T001'}),config('Colaborador',5,{Empresa:'Teste',Codigo_Colaborador:'T001',Vigencia_Inicio:'2027-01-01'}),config('Colaborador',6,{Estado:'Rascunho',Empresa:'Teste',Codigo_Colaborador:'T001'})];
 assert.equal(resolveConfig(rows,'TAXA',context,at).Valor,4);
 assert.equal(resolveConfig(rows,'TAXA',{empresa:'Outra'},at).Valor,1);
 assert.throws(()=>resolveConfig([config('Geral',1),config('Geral',2)],'TAXA',context,at),/ambíguas/);
 assert.throws(()=>resolveConfig([config('Geral',1,{Vigencia_Inicio:'errada'})],'TAXA',context,at),/Vigência/);
});
test('simulação usa fórmula da tabela e parâmetros resolvidos sem converter taxa outra vez',()=>{
 assert.equal(simulateRecord(tables(),'TESTE',context,at,{BASE:1000,TAXA:99}).result,30);
 assert.throws(()=>simulateRecord(tables(),'TESTE',context,at,{}),/Variável desconhecida/);
 const data=tables();data.tbl_formulas_RH[0].Validada=false;
 assert.equal(resolveTables(data,context,at).unavailable.length,1);
 assert.throws(()=>simulateRecord(data,'TESTE',context,at,{BASE:1000}),/não validada/);
});
test('IF analisa ambos os ramos mas executa apenas o seleccionado',()=>{
 assert.equal(evaluateFormula('IF(BASE == 0, 0, 10 / BASE)',{BASE:0}),0);
 assert.equal(evaluateFormula('IF(TRUE, 3, AUSENTE)',{}),3);
 assert.throws(()=>evaluateFormula('IF(TRUE, 3, exec(1))'),/Função desconhecida/);
 assert.equal(evaluateFormula('IF(REGIME == "Mensal", ROUND(BASE / DIAS, 2), 0)',{REGIME:'Mensal',BASE:17000,DIAS:30}),566.67);
 assert.equal(evaluateFormula('IF(CANALIZAR == TRUE, 1, 0)',{CANALIZAR:true}),1);
});
test('arredondamentos, calendário e erros numéricos',()=>{
 assert.equal(evaluateFormula(' ROUND(-1.005, 2) '),-1.01);
 assert.equal(evaluateFormula('ROUNDUP(-1.001, 2)'),-1.01);
 assert.equal(evaluateFormula('ROUNDDOWN(-1.009, 2)'),-1);
 assert.equal(evaluateFormula('NETWORKDAYS(INICIO, FIM, FERIADOS)',{INICIO:'2026-10-19',FIM:'2026-10-23',FERIADOS:['2026-10-20']}),4);
 assert.throws(()=>evaluateFormula('1 / 0'),/Divisão por zero/);
 assert.throws(()=>evaluateFormula('BASE * 2',{BASE:null}),/não numérico/);
 assert.throws(()=>evaluateFormula('globalThis.alert(1)'),/Símbolo não permitido/);
});
test('escalão aberto seguido por outro é rejeitado',()=>{
 assert.throws(()=>validateScales([{...common,ID_Tabela_Escaloes:'TESTE',Limite_Inferior:0,Limite_Superior:''},{...common,ID_Tabela_Escaloes:'TESTE',Limite_Inferior:10,Limite_Superior:20}],'TESTE',{},at),/sobrepostos/);
});
