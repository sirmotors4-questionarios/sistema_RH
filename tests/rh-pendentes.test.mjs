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
