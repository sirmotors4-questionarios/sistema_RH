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
test('bloqueios e duplicados impedem preparar um lote parcial',()=>{
 for(const funcionarios of [[],[f('1'),f('1')]]){
  const r=prepararPendentes({lancamentos:[row('1','Pendente')],funcionarios,ano:2026,mes:4});
  assert.ok(r.erros.length);assert.equal(r.linhas.length,0);
 }
 const r=prepararPendentes({lancamentos:[row('1','Pendente'),row('1','Pendente')],funcionarios:[f('1')],ano:2026,mes:4});assert.ok(r.erros.length);assert.equal(r.linhas.length,0);
});
test('API usa envelope no ramo editar e exige confirmação v2',async()=>{
 let pedido;
 const deps={getToken:async()=> 'teste',url:'https://example.test',fetchImpl:async(u,p)=>{pedido=JSON.parse(p.body);return {ok:true,status:200,json:async()=>({ok:true,linha:JSON.stringify({versao:2,operacao:'recalcular_pendentes',linhas:[],alteradas:0,ignoradas:1})})}}};
 const r=await operarLancamentos('recalcular_pendentes',{ano:2026,mes:4,mesNome:'Abril',atualizacoes:[{Empresa:'A',Codigo:'1'}]},deps);
 assert.equal(r.alteradas,0);assert.equal(pedido.action,'lancamento_editar');assert.equal(pedido.linhas.length,1);assert.equal(pedido.linhas[0].__rh_operacao,'recalcular_pendentes');
 await assert.rejects(operarLancamentos('repor_inicial',{ano:2026,mes:4,Empresa:'A',Codigo:'1'},deps),/Office Script/);
});
