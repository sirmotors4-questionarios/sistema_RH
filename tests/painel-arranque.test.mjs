import test from 'node:test';
import assert from 'node:assert/strict';
import {iniciarPainel} from '../painel.js';

const tick = () => new Promise(ok => setImmediate(ok));
function ambiente({falharMes = false} = {}) {
  const anteriores = {window: globalThis.window, document: globalThis.document, location: globalThis.location};
  const elementos = new Map();
  let iframe, enviado, leituras = 0;
  const elemento = () => ({hidden:false, disabled:false, textContent:'', replaceChildren(){}, addEventListener(){}, remove(){}});
  globalThis.window = new EventTarget();
  globalThis.location = {origin:'https://rh.example'};
  globalThis.document = {
    getElementById(id) { if (!elementos.has(id)) elementos.set(id, elemento()); return elementos.get(id); },
    createElement() { iframe = {...elemento(), contentWindow:{postMessage(m){enviado=m;}}}; return iframe; }
  };
  const painel = iniciarPainel({ligado:()=>true, funcionarios:async()=>[], periodoMaisRecente:async()=>({ano:2026,mes:11}),
    carregarMes:async()=>{leituras++; if(falharMes)throw Error('Rede indisponível');return [];},
    indemnizacoes:async()=>({indemnizacoes:[],prestacoes:[]})});
  return {painel, get enviado(){return enviado;}, get leituras(){return leituras;},
    mensagem(tipo, extras={}, origin=location.origin) {const ev=new Event('message');Object.assign(ev,{origin,source:iframe.contentWindow,data:{tipo,...extras}});window.dispatchEvent(ev);},
    limpar(){painel.limpar();Object.assign(globalThis,anteriores);}};
}
test('dashboard aguarda a versão apresentada, ignora origem/pedido incorrectos e reutiliza dados', async()=>{
  const a=ambiente();
  try {
    let concluido=false;
    const pedido=a.painel.abrir(false,{estrito:true}).then(()=>concluido=true);
    await tick();a.mensagem('rh-dashboard-ola');await tick();
    assert.equal(a.leituras,12);assert.equal(concluido,false);
    a.mensagem('rh-dashboard-pronto',{pedido:a.enviado.pedido},'https://outra.example');
    a.mensagem('rh-dashboard-pronto',{pedido:-1});await tick();assert.equal(concluido,false);
    a.mensagem('rh-dashboard-pronto',{pedido:a.enviado.pedido});await pedido;
    assert.equal(concluido,true);await a.painel.abrir();assert.equal(a.leituras,12);
  } finally {a.limpar();}
});
test('histórico incompleto bloqueia a entrada estrita', async()=>{
  const a=ambiente({falharMes:true});
  try {await assert.rejects(a.painel.abrir(false,{estrito:true}),/12 mês/);assert.equal(a.enviado,undefined);}
  finally {a.limpar();}
});
test('erro do Streamlit não é tratado como dashboard pronto', async()=>{
  const a=ambiente();
  try {
    const pedido=a.painel.abrir(false,{estrito:true});
    const rejeitado=assert.rejects(pedido,/Erro nos gráficos/);
    await tick();a.mensagem('rh-dashboard-ola');await tick();
    a.mensagem('rh-dashboard-erro',{pedido:a.enviado.pedido,erro:'Erro nos gráficos'});await rejeitado;
  } finally {a.limpar();}
});
