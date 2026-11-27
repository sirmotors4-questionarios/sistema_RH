import test from 'node:test';
import assert from 'node:assert/strict';
import {prepararPendentes} from '../rh-pendentes.js';
import {operarLancamentos} from '../rh-api.js';
const row=(Codigo,estado)=>({Codigo,Empresa:'A',Ano:2026,'Mês':'Abril','Estado do Lançamento':estado,'Dias Trabalhados':30});
const f=Codigo=>({Codigo,Empresa:'A',Estado:'Activo',TipoContrato:'Efectivo',RegimePagamento:'Mensal','Vencimento Base':20000,'Data Admissao':'2026-04-11'});
test('lote só prepara pendentes, mantém entradas e estados dos restantes',()=>{
 const lancamentos=[row('1','Pendente'),row('2','Alterado'),row('3','Finalizado'),row('4','')],antes=structuredClone(lancamentos);
 const r=prepararPendentes({lancamentos,funcionarios:[f('1')],ano:2026,mes:4});
 assert.deepEqual(r.erros,[]);assert.equal(r.linhas.length,1);assert.equal(r.preservados,3);
 assert.equal(r.linhas[0]['Dias Trabalhados'],10);assert.ok(!('Estado do Lançamento' in r.linhas[0]));assert.deepEqual(lancamentos,antes);
});
test('bloqueios ficam de fora sem impedir os restantes Pendentes',()=>{
 for(const funcionarios of [[f('2')],[f('1'),f('1'),f('2')]]){
  const r=prepararPendentes({lancamentos:[row('1','Pendente'),row('2','Pendente')],funcionarios,ano:2026,mes:4});
  assert.equal(r.erros.length,1);assert.match(r.erros[0],/^A\|0001: /);assert.equal(r.linhas.length,1);assert.equal(r.linhas[0].Codigo,'0002');
 }
 // Lançamento repetido: nenhuma das cópias é recalculada.
 const rep=prepararPendentes({lancamentos:[row('1','Pendente'),row('1','Pendente'),row('2','Pendente')],funcionarios:[f('1'),f('2')],ano:2026,mes:4});
 assert.equal(rep.erros.length,1);assert.match(rep.erros[0],/repetido/);assert.deepEqual(rep.linhas.map(l=>l.Codigo),['0002']);
 // Admissão posterior ao período (caso real SIR|0226): fica de fora, os outros seguem.
 const tarde={...f('3'),'Data Admissao':'2026-05-02'};
 const adm=prepararPendentes({lancamentos:[row('2','Pendente'),row('3','Pendente')],funcionarios:[f('2'),tarde],ano:2026,mes:4});
 assert.equal(adm.linhas.length,1);assert.match(adm.erros[0],/^A\|0003: Admissão posterior/);
});
test('API usa envelope no ramo editar e exige confirmação v2',async()=>{
 let pedido;
 const deps={getToken:async()=> 'teste',url:'https://example.test',fetchImpl:async(u,p)=>{pedido=JSON.parse(p.body);return {ok:true,status:200,json:async()=>({ok:true,linha:JSON.stringify({versao:2,operacao:'recalcular_pendentes',linhas:[],alteradas:0,ignoradas:1})})}}};
 const r=await operarLancamentos('recalcular_pendentes',{ano:2026,mes:4,mesNome:'Abril',atualizacoes:[{Empresa:'A',Codigo:'1'}]},deps);
 assert.equal(r.alteradas,0);assert.equal(pedido.action,'lancamento_editar');assert.equal(pedido.linhas.length,1);assert.equal(pedido.linhas[0].__rh_operacao,'recalcular_pendentes');
 await assert.rejects(operarLancamentos('repor_inicial',{ano:2026,mes:4,Empresa:'A',Codigo:'1'},deps),/Office Script/);
});
import {prepararSincronizacao} from '../rh-pendentes.js';
test('sincronização: retira Pendentes sem contrato, introduz em falta e preserva Alterados',()=>{
 const tarde=c=>({...f(c),'Data Admissao':'2026-05-02'});
 const r=prepararSincronizacao({ano:2026,mes:4,
  lancamentos:[row('1','Pendente'),row('2','Pendente'),row('3','Alterado'),row('4','Retirado')],
  funcionarios:[f('1'),tarde('2'),tarde('3'),tarde('4'),f('5')]});
 assert.deepEqual(r.retirar.map(x=>x.Codigo),['0002']);assert.match(r.retirar[0].motivo,/Admissão posterior/);
 assert.deepEqual(r.aRever.map(x=>x.Codigo),['0003']);
 assert.deepEqual(r.atualizacoes.map(l=>l.Codigo),['0001']);assert.deepEqual(r.erros,[]);
 assert.deepEqual(r.introduzir.map(l=>l.Codigo),['0005']);assert.equal(r.introduzir[0]['Estado do Lançamento'],'Pendente');
});
test('sincronização não retira com cadastro ambíguo; introduz só quem tem dados completos',()=>{
 const r=prepararSincronizacao({ano:2026,mes:4,lancamentos:[row('1','Pendente')],funcionarios:[f('1'),f('1'),f('7'),f('7'),f('8'),{...f('9'),'Vencimento Base':''}]});
 assert.deepEqual(r.retirar,[]);assert.equal(r.erros.length,1);
 assert.deepEqual(r.introduzir.map(l=>l.Codigo),['0008']);
 assert.ok(r.avisosIntroducao.some(e=>/repetido/.test(e)));assert.ok(r.avisosIntroducao.some(e=>/0009/.test(e)));
});
test('API v3 exige a nova versão do script',async()=>{
 const resp=l=>({getToken:async()=>'t',url:'https://example.test',fetchImpl:async(u,p)=>({ok:true,status:200,json:async()=>({ok:true,linha:JSON.stringify(l)})})});
 await assert.rejects(operarLancamentos('retirar_pendentes',{ano:2026,mes:4,mesNome:'Abril',retirar:[{Empresa:'A',Codigo:'1'}]},resp({versao:2,operacao:'retirar_pendentes',linhas:[]})),/ainda não sabe/);
 const r=await operarLancamentos('introduzir_lancamentos',{ano:2026,mes:4,mesNome:'Abril',linhas:[{Empresa:'A',Codigo:'1'}]},resp({versao:3,operacao:'introduzir_lancamentos',alteradas:1,linhas:[]}));
 assert.equal(r.alteradas,1);
});
import {recalcularEmLotes} from '../rh-pendentes.js';
test('lotes adaptativos: divide o lote que falha e isola a linha problemática',async()=>{
 const linhas=Array.from({length:100},(_,i)=>({Empresa:'SIR',Codigo:String(i).padStart(4,'0')}));
 const gravados=new Set(),erro=Object.assign(Error('Não foi possível executar o Office Script. (500)'),{status:500});
 const operar=async g=>{if(g.some(l=>l.Codigo==='0087'))throw erro;g.forEach(l=>gravados.add(l.Codigo));return {alteradas:g.length};};
 const r=await recalcularEmLotes({linhas,lote:40,operar});
 assert.equal(r.alteradas,99);assert.deepEqual(r.falhas.map(f=>f.chave),['SIR|0087']);assert.match(r.falhas[0].motivo,/500/);
 assert.equal(gravados.size,99);assert.ok(r.execucoes<=20,'execuções: '+r.execucoes);
});
test('lotes adaptativos: lote demorado passa ao ser dividido; erros 4xx não são repetidos',async()=>{
 const linhas=Array.from({length:40},(_,i)=>({Empresa:'A',Codigo:String(i)}));
 const r=await recalcularEmLotes({linhas,lote:40,operar:async g=>{if(g.length>20)throw Object.assign(Error('Timeout'),{status:504});return {alteradas:g.length};}});
 assert.equal(r.alteradas,40);assert.equal(r.falhas.length,0);assert.equal(r.execucoes,3);
 await assert.rejects(recalcularEmLotes({linhas,operar:async()=>{throw Object.assign(Error('Sem permissão'),{status:403});}}),/Sem permissão/);
});
