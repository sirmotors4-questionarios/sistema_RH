import test from 'node:test';
import assert from 'node:assert/strict';
import {readFileSync} from 'node:fs';
import {stripTypeScriptTypes} from 'node:module';
import vm from 'node:vm';
// Executa o próprio Office Script com uma tabela em memória: nenhuma chamada ao Excel real.
const source=stripTypeScriptTypes(readFileSync(new URL('../office-scripts/LancamentoNovoMes.ts',import.meta.url),'utf8'));
const ctx=vm.createContext({ExcelScript:{CalculationType:{full:'full',recalculate:'recalculate'}}});vm.runInContext(source,ctx);
const cab=['Empresa','Codigo','Ano','Mês','Estado do Lançamento','Revisto','RegimePagamento','Valor Unitário','Dias Trabalhados','Data Inicio Contrato','Data Fim Contrato','Subsídios','Faltas Injustificadas','Horas Trabalhadas','Observações','Utilizador Alteracao','Data Ultima Alteracao','Vencimento','Escalão'];
const obj=(code,state='Pendente')=>({Empresa:'A',Codigo:code,Ano:2026,'Mês':'Abril','Estado do Lançamento':state,Revisto:'Não',RegimePagamento:'Mensal','Valor Unitário':10000,'Dias Trabalhados':30,'Subsídios':500,'Faltas Injustificadas':2,'Observações':'Inicial',Vencimento:10000});
function workbook(objects){
 const tables=new Map();let calculations=0,leituras=0;
 function table(headers,rows,name){
  const range=(r=0,c=0,nr=rows.length,nc=headers.length)=>({getValues:()=>{leituras++;return rows.slice(r,r+nr).map(l=>l.slice(c,c+nc))},
   getFormulasR1C1:()=>{leituras++;return rows.slice(r,r+nr).map(l=>l.slice(c,c+nc).map(String))},setFormulaR1C1:v=>rows[r][c]=v,
   getFormulas:()=>{leituras++;return rows.slice(r,r+nr).map(l=>l.slice(c,c+nc).map(String))},
   getWorksheet:()=>({getRangeByIndexes:(ri,ci,a,b)=>range(ri,ci,a,b)}),getRowIndex:()=>r,getColumnIndex:()=>c,
   setValues:vs=>vs.forEach((l,i)=>l.forEach((v,j)=>rows[r+i][c+j]=v)),setValue:v=>rows[r][c]=v,
   getRow:i=>range(r+i,c,1,nc),getColumn:j=>range(r,c+j,nr,1),getCell:(i,j)=>range(r+i,c+j,1,1)});
  const t={getHeaderRowRange:()=>({getValues:()=>[headers]}),getRangeBetweenHeaderAndTotal:()=>range(),getRowCount:()=>rows.length,
   addRows:(i,vs)=>rows.push(...vs.map(l=>[...l])),setName:n=>tables.set(n,t),rows,headers};if(name)tables.set(name,t);return t;
 }
 const main=table(cab,objects.map(o=>cab.map(h=>o[h]??'')),'tbl_lançamento');
 return {tables,main,getTable:n=>tables.get(n),getApplication:()=>({calculate:()=>{calculations++;for(const l of main.rows)l[cab.indexOf('Vencimento')]=l[cab.indexOf('Valor Unitário')]*l[cab.indexOf('Dias Trabalhados')]/30}}),
  addWorksheet:()=>{let headers;return {getRange:()=>({setValues:v=>headers=v[0]}),addTable:()=>table(headers,[]) }},
  get calculations(){return calculations},get leituras(){return leituras},read:i=>Object.fromEntries(cab.map((h,j)=>[h,main.rows[i][j]]))};
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

// ---------- v3: retirar e introduzir (Carregar mês) ----------
test('retirar apaga só Pendentes indicados e revalida o estado no Excel',()=>{
 const w=workbook([obj('1'),obj('2','Alterado'),obj('3'),obj('4')]);
 w.main.deleteRowsAt=(i,n=1)=>w.main.rows.splice(i,n);w.main.getAutoFilter=()=>null;
 const r=run(w,[{__rh_operacao:'retirar_pendentes',retirar:[{Empresa:'A',Codigo:'1'},{Empresa:'A',Codigo:'2'},{Empresa:'A',Codigo:'3'}]}]);
 assert.equal(r.ok,true);assert.equal(r.alteradas,2);
 const l=JSON.parse(r.linha);assert.equal(l.versao,3);assert.equal(l.operacao,'retirar_pendentes');assert.equal(l.ignoradas,1);
 assert.deepEqual(w.main.rows.map(x=>x[1]),['2','4']);
});
test('retirar com lançamento inexistente não apaga nada',()=>{
 const w=workbook([obj('1'),obj('2')]);w.main.deleteRowsAt=(i,n=1)=>w.main.rows.splice(i,n);w.main.getAutoFilter=()=>null;
 const r=run(w,[{__rh_operacao:'retirar_pendentes',retirar:[{Empresa:'A',Codigo:'1'},{Empresa:'A',Codigo:'9'}]}]);
 assert.equal(r.ok,false);assert.equal(w.main.rows.length,2);
});
test('introduzir insere num mês existente com as regras do gravar e recusa chaves já existentes',()=>{
 const PROT=['Estado','Estado de Reforma','Vencimento Base','INSS Colaborador','INSS Empresa','Total INSS','Salario Liquido','Total de Acréscimos','Total de Descontos','IRPS','Vencimento'];
 const headers=['Empresa','Codigo','Ano','Mês','Nome','Estado do Lançamento',...PROT];
 const linha=(c)=>['A',c,2026,'Abril','N'+c,'Pendente',...PROT.map(p=>'=F_'+p)];
 const rows=[linha('0001')];
 const range=(r,c,nr,nc)=>({getValues:()=>rows.slice(r,r+nr).map(l=>l.slice(c,c+nc).map(v=>String(v).startsWith('=')?0:v)),
  getFormulasR1C1:()=>rows.slice(r,r+nr).map(l=>l.slice(c,c+nc).map(String)),
  setFormulasR1C1:vs=>vs.forEach((l,i)=>l.forEach((v,j)=>rows[r+i][c+j]=v)),setValues:vs=>vs.forEach((l,i)=>l.forEach((v,j)=>rows[r+i][c+j]=v)),
  setNumberFormat(){},getColumn:j=>range(r,c+j,nr,1),getRow:i=>range(r+i,c,1,nc),getCell:(i,j)=>range(r+i,c+j,1,1),
  getRowIndex:()=>r-1,getColumnIndex:()=>c,getWorksheet:()=>({getRangeByIndexes:(ri,ci,nr2,nc2)=>range(ri+1,ci,nr2,nc2)})});
 const corpo={...range(0,0,1e6,headers.length),getRowIndex:()=>0,getColumnIndex:()=>0,getWorksheet:()=>({getRangeByIndexes:(ri,ci,nr2,nc2)=>range(ri,ci,nr2,nc2)})};
 corpo.getValues=()=>rows.map(l=>l.map(v=>String(v).startsWith('=')?0:v));corpo.getFormulasR1C1=()=>rows.map(l=>l.map(String));
 corpo.getColumn=j=>({getValues:()=>rows.map(l=>[String(l[j]).startsWith('=')?0:l[j]])});
 const t={getHeaderRowRange:()=>({getValues:()=>[headers]}),getRangeBetweenHeaderAndTotal:()=>corpo,getRowCount:()=>rows.length,
  addRows:(i,vs)=>rows.push(...vs.map(l=>[...l])),getAutoFilter:()=>null};
 const w={getTable:n=>n==='tbl_lançamento'?t:null,getApplication:()=>({calculate(){}})};
 const novo=c=>({Empresa:'A',Codigo:c,Nome:'N'+c,'Estado do Lançamento':'Pendente'});
 const r=JSON.parse(JSON.stringify(ctx.main(w,2026,'Abril',JSON.stringify([{__rh_operacao:'introduzir_lancamentos',linhas:[novo('0226')]}]),'editar')));
 assert.equal(r.ok,true,r.erros.join());assert.equal(r.inseridas,1);assert.equal(JSON.parse(r.linha).operacao,'introduzir_lancamentos');
 assert.equal(rows.length,2);assert.equal(rows[1][1],'0226');assert.equal(rows[1][headers.indexOf('Vencimento')],'=F_Vencimento');
 const dup=JSON.parse(JSON.stringify(ctx.main(w,2026,'Abril',JSON.stringify([{__rh_operacao:'introduzir_lancamentos',linhas:[novo('0001')]}]),'editar')));
 assert.equal(dup.ok,false);assert.equal(rows.length,2);
});

// ---------- desempenho: leituras ao Excel não dependem do número de colaboradores ----------
test('recalcular 200 Pendentes faz um número fixo de leituras ao Excel',()=>{
 const contar=n=>{const w=workbook(Array.from({length:n},(_,i)=>obj(String(i+1))));const antes=w.leituras;
  const r=batch(w,Array.from({length:n},(_,i)=>String(i+1)));assert.equal(r.ok,true);assert.equal(r.alteradas,n);return w.leituras-antes;};
 const poucas=contar(2),muitas=contar(200);
 assert.equal(muitas,poucas,'leituras: '+poucas+' com 2 linhas, '+muitas+' com 200');
 assert.ok(muitas<=15,'leituras a mais: '+muitas);
});
test('escalões preenchidos numa só passagem, também em linhas não seguidas',()=>{
 const w=workbook([obj('1'),obj('2','Alterado'),obj('3')]);
 const esc=[['A','Vencimento','E1',0],['A','Vencimento','E2',5000],['A','Vencimento','E3',8000]];
 w.tables.set('tbl_escaloes_lancamento',{getRowCount:()=>esc.length,getHeaderRowRange:()=>({getValues:()=>[['Empresa','Criterio','Escalao','Valor_Inicial']]}),getRangeBetweenHeaderAndTotal:()=>({getValues:()=>esc})});
 const r=batch(w,['1','2','3']);assert.equal(r.ok,true);assert.equal(r.alteradas,2);
 assert.equal(w.read(0).Vencimento,6000);assert.equal(w.read(0)['Escalão'],'E2');assert.equal(w.read(2)['Escalão'],'E2');assert.equal(w.read(1)['Escalão'],'');
 // Edição de uma linha recalcula o escalão dessa linha.
 run(w,[{Empresa:'A',Codigo:'3','Dias Trabalhados':15}]);assert.equal(w.read(2).Vencimento,9000);assert.equal(w.read(2)['Escalão'],'E3');
});
test('retirar linhas seguidas usa um só pedido de remoção',()=>{
 const w=workbook([obj('1'),obj('2'),obj('3'),obj('4'),obj('5')]);const chamadas=[];
 w.main.deleteRowsAt=(i,n=1)=>{chamadas.push([i,n]);w.main.rows.splice(i,n)};w.main.getAutoFilter=()=>null;
 const r=run(w,[{__rh_operacao:'retirar_pendentes',retirar:['2','3','4','1'].map(c=>({Empresa:'A',Codigo:c}))}]);
 assert.equal(r.ok,true);assert.deepEqual(chamadas,[[0,4]]);assert.deepEqual(w.main.rows.map(x=>x[1]),['5']);
});
