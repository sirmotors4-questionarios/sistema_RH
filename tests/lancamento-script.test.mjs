import test from 'node:test';
import assert from 'node:assert/strict';
import {readFileSync} from 'node:fs';
import {stripTypeScriptTypes} from 'node:module';
import vm from 'node:vm';
// Executa o próprio Office Script com uma tabela em memória: nenhuma chamada ao Excel real.
const source=stripTypeScriptTypes(readFileSync(new URL('../office-scripts/LancamentoNovoMes.ts',import.meta.url),'utf8'));
const ctx=vm.createContext({ExcelScript:{CalculationType:{full:'full'}}});vm.runInContext(source,ctx);
const cab=['Empresa','Codigo','Ano','Mês','Estado do Lançamento','Revisto','RegimePagamento','Valor Unitário','Dias Trabalhados','Data Inicio Contrato','Data Fim Contrato','Subsídios','Faltas Injustificadas','Horas Trabalhadas','Observações','Utilizador Alteracao','Data Ultima Alteracao','Vencimento'];
const obj=(code,state='Pendente')=>({Empresa:'A',Codigo:code,Ano:2026,'Mês':'Abril','Estado do Lançamento':state,Revisto:'Não',RegimePagamento:'Mensal','Valor Unitário':10000,'Dias Trabalhados':30,'Subsídios':500,'Faltas Injustificadas':2,'Observações':'Inicial',Vencimento:10000});
function workbook(objects){
 const tables=new Map();let calculations=0;
 function table(headers,rows,name){
  const range=(r=0,c=0,nr=rows.length,nc=headers.length)=>({getValues:()=>rows.slice(r,r+nr).map(l=>l.slice(c,c+nc)),
   getFormulasR1C1:()=>rows.slice(r,r+nr).map(l=>l.slice(c,c+nc).map(String)),setFormulaR1C1:v=>rows[r][c]=v,
   setValues:vs=>vs.forEach((l,i)=>l.forEach((v,j)=>rows[r+i][c+j]=v)),setValue:v=>rows[r][c]=v,
   getRow:i=>range(r+i,c,1,nc),getColumn:j=>range(r,c+j,nr,1),getCell:(i,j)=>range(r+i,c+j,1,1)});
  const t={getHeaderRowRange:()=>({getValues:()=>[headers]}),getRangeBetweenHeaderAndTotal:()=>range(),getRowCount:()=>rows.length,
   addRows:(i,vs)=>rows.push(...vs.map(l=>[...l])),setName:n=>tables.set(n,t),rows,headers};if(name)tables.set(name,t);return t;
 }
 const main=table(cab,objects.map(o=>cab.map(h=>o[h]??'')),'tbl_lançamento');
 return {tables,main,getTable:n=>tables.get(n),getApplication:()=>({calculate:()=>{calculations++;for(const l of main.rows)l[cab.indexOf('Vencimento')]=l[cab.indexOf('Valor Unitário')]*l[cab.indexOf('Dias Trabalhados')]/30}}),
  addWorksheet:()=>{let headers;return {getRange:()=>({setValues:v=>headers=v[0]}),addTable:()=>table(headers,[]) }},
  get calculations(){return calculations},read:i=>Object.fromEntries(cab.map((h,j)=>[h,main.rows[i][j]]))};
}
const run=(w,rows)=>JSON.parse(JSON.stringify(ctx.main(w,2026,'Abril',JSON.stringify(rows),'editar')));
const patch=code=>({Empresa:'A',Codigo:code,RegimePagamento:'Mensal','Valor Unitário':18000,'Dias Trabalhados':10,'Data Inicio Contrato':46123,'Data Fim Contrato':'','Utilizador Alteracao':'teste','Data Ultima Alteracao':46200});
const batch=(w,codes)=>run(w,[{__rh_operacao:'recalcular_pendentes',atualizacoes:codes.map(patch)}]);
const restore=(w,Codigo)=>run(w,[{__rh_operacao:'repor_inicial',Empresa:'A',Codigo,'Utilizador Alteracao':'teste','Data Ultima Alteracao':46201}]);
test('script revalida estados, só recalcula Pendentes e conserva entradas',()=>{
 const w=workbook([obj('1'),obj('2','Alterado'),obj('3','Finalizado')]);const before=w.main.rows.map(l=>[...l]);
 const r=batch(w,['1','2','3']);assert.equal(r.ok,true);assert.equal(r.alteradas,1);assert.equal(JSON.parse(r.linha).ignoradas,2);
 assert.equal(w.read(0)['Dias Trabalhados'],10);assert.equal(w.read(0)['Subsídios'],500);assert.equal(w.read(0)['Estado do Lançamento'],'Pendente');assert.equal(w.read(0).Vencimento,6000);
 assert.deepEqual(w.main.rows[1],before[1]);assert.deepEqual(w.main.rows[2],before[2]);
});
test('referência persiste após edições e reposição não usa valores enviados pelo navegador',()=>{
 const w=workbook([obj('1')]);batch(w,['1']);
 assert.equal(run(w,[{Empresa:'A',Codigo:'1','Subsídios':900,'Dias Trabalhados':15}]).ok,true);
 assert.equal(run(w,[{Empresa:'A',Codigo:'1','Subsídios':1200}]).ok,true);
 const r=restore(w,'1');assert.equal(r.ok,true);assert.equal(w.read(0)['Subsídios'],500);assert.equal(w.read(0)['Dias Trabalhados'],10);assert.equal(w.read(0)['Estado do Lançamento'],'Pendente');assert.equal(w.read(0).Vencimento,6000);
 assert.equal(w.tables.get('tbl_lancamentos_iniciais_RH').getRowCount(),1);
});
test('edição sem recálculo prévio também captura referência de Pendente',()=>{
 const w=workbook([obj('1')]);run(w,[{Empresa:'A',Codigo:'1','Subsídios':900}]);assert.equal(restore(w,'1').ok,true);assert.equal(w.read(0)['Subsídios'],500);
});
test('Finalizados e Alterados antigos sem referência não podem ser repostos',()=>{
 const w=workbook([obj('1','Alterado'),obj('2','Finalizado')]);const before=JSON.stringify(w.main.rows);
 assert.match(restore(w,'1').erros.join(),/Não existe referência/);assert.match(restore(w,'2').erros.join(),/Finalizado/);assert.equal(JSON.stringify(w.main.rows),before);
});
test('lote inválido não grava parcialmente; fórmula e duplicados são recusados',()=>{
 for(const rows of [[patch('1'),{...patch('2'),Vencimento:1}],[patch('1'),patch('1')]]){
  const w=workbook([obj('1'),obj('2')]),before=JSON.stringify(w.main.rows);
  const r=run(w,[{__rh_operacao:'recalcular_pendentes',atualizacoes:rows}]);assert.equal(r.ok,false);assert.equal(JSON.stringify(w.main.rows),before);assert.equal(w.tables.size,1);
 }
});
test('novo carregamento preserva Alterado e a sua referência original',()=>{
 const w=workbook([obj('1')]);batch(w,['1']);run(w,[{Empresa:'A',Codigo:'1','Subsídios':999}]);
 assert.equal(batch(w,['1']).alteradas,0);assert.equal(w.read(0)['Subsídios'],999);restore(w,'1');assert.equal(w.read(0)['Subsídios'],500);
});

test('reposição mantém fórmulas existentes nas colunas de entradas',()=>{
 const w=workbook([{...obj('1'),'Horas Trabalhadas':'=SUM(RC[-1])'}]);
 run(w,[{Empresa:'A',Codigo:'1','Horas Trabalhadas':10}]);
 assert.equal(restore(w,'1').ok,true);assert.equal(w.read(0)['Horas Trabalhadas'],'=SUM(RC[-1])');
});
