import test from 'node:test';
import assert from 'node:assert/strict';
import {callApi, loadConfiguracoes, permissoes, ApiError} from '../rh-api.js';

const tabelas = {tbl_configuracoes_RH: [{Codigo_Parametro: 'X'}], tbl_escaloes_RH: [], tbl_formulas_RH: [], tbl_auditoria_RH: []};
const resposta = (status, body) => async () => ({status, ok: status >= 200 && status < 300, json: async () => body});
const getToken = async () => 'token-teste';

test('envia POST autenticado com a acção pedida', async () => {
  let pedido;
  const fetchImpl = async (url, init) => { pedido = {url, init}; return resposta(200, {ok: true, ...tabelas, utilizador: {}})(); };
  await callApi('configuracoes', {getToken, fetchImpl, url: 'https://exemplo.test/fluxo'});
  assert.equal(pedido.url, 'https://exemplo.test/fluxo');
  assert.equal(pedido.init.method, 'POST');
  assert.equal(pedido.init.headers.Authorization, 'Bearer token-teste');
  assert.deepEqual(JSON.parse(pedido.init.body), {action: 'configuracoes'});
});

test('carrega as quatro tabelas e o perfil do utilizador', async () => {
  const body = {ok: true, ...tabelas, utilizador: {Nome: 'NB', Email: 'nb@sir.test', Perfil: 'Tecnico RH', Empresa: 'Todas'}};
  const r = await loadConfiguracoes({getToken, fetchImpl: resposta(200, body), url: 'u'});
  assert.equal(r.tables.tbl_configuracoes_RH.length, 1);
  assert.equal(r.utilizador.perfil, 'Técnico RH');
  assert.equal(r.utilizador.propor, false);
  assert.equal(r.utilizador.consultar, true);
});

test('403 identifica a conta não autorizada', async () => {
  await assert.rejects(callApi('configuracoes', {getToken, fetchImpl: resposta(403, {ok: false, email: 'x@sir.test'}), url: 'u'}),
    e => e instanceof ApiError && e.code === 'RH_FORBIDDEN' && /x@sir\.test/.test(e.message));
});

test('401 pede nova sessão', async () => {
  await assert.rejects(callApi('configuracoes', {getToken, fetchImpl: resposta(401, null), url: 'u'}), e => e.code === 'RH_AUTH');
});

test('falha de rede e resposta incompleta são erros claros', async () => {
  await assert.rejects(callApi('configuracoes', {getToken, fetchImpl: async () => { throw new TypeError('offline'); }, url: 'u'}), e => e.code === 'RH_NETWORK');
  await assert.rejects(callApi('configuracoes', {getToken, fetchImpl: resposta(200, {ok: true}), url: 'u'}).then(b => loadConfiguracoes({getToken, fetchImpl: resposta(200, b), url: 'u'})), /tbl_configuracoes_RH/);
  await assert.rejects(callApi('configuracoes', {getToken, fetchImpl: resposta(500, {ok: false}), url: 'u'}), e => e.code === 'RH_RESPONSE');
});

test('sem autenticação não chama o fluxo', async () => {
  let chamado = false;
  await assert.rejects(callApi('configuracoes', {fetchImpl: async () => { chamado = true; }}), e => e.code === 'RH_AUTH');
  assert.equal(chamado, false);
});

test('perfis: apenas Superadministrador aprova; perfis desconhecidos não têm acesso', () => {
  assert.deepEqual(permissoes('Superadministrador'), {perfil: 'Superadministrador', consultar: true, propor: true, aprovar: true});
  assert.equal(permissoes('administrador rh').propor, true);
  assert.equal(permissoes('Administrador RH').aprovar, false);
  assert.equal(permissoes('Gestor').propor, false);
  assert.equal(permissoes('Desconhecido').consultar, false);
});

test('loadModulo envia a acção ao fluxo de consulta e devolve as tabelas por nome', async () => {
  const {loadModulo} = await import('../rh-api.js');
  let pedido;
  const body = {ok: true, action: 'ferias', utilizador: {Perfil: 'Gestor'}, tabelas: [
    {tabela: 'tbl_ferias', linhas: [{Codigo: 'T1'}]}, {tabela: 'tbl_feriados', linhas: []}]};
  const fetchImpl = async (url, init) => { pedido = {url, init}; return resposta(200, body)(); };
  const r = await loadModulo('ferias', {getToken, fetchImpl, url: 'https://exemplo.test/consultar'});
  assert.deepEqual(JSON.parse(pedido.init.body), {action: 'ferias'});
  assert.equal(pedido.url, 'https://exemplo.test/consultar');
  assert.equal(r.tabelas.tbl_ferias.length, 1);
  assert.deepEqual(r.tabelas.tbl_feriados, []);
  assert.equal(r.utilizador.perfil, 'Gestor');
});

test('loadModulo rejeita módulo desconhecido, tabelas em falta e 400', async () => {
  const {loadModulo} = await import('../rh-api.js');
  await assert.rejects(loadModulo('salarios', {getToken, fetchImpl: resposta(200, {}), url: 'u'}), e => e.code === 'RH_ACTION');
  await assert.rejects(loadModulo('contratos', {getToken, fetchImpl: resposta(200, {ok: true, tabelas: [{tabela: 'tbl_adendas', linhas: []}]}), url: 'u'}), /tbl_parametros_pagamento/);
  await assert.rejects(loadModulo('ferias', {getToken, fetchImpl: resposta(400, {ok: false, erro: 'Accao invalida.'}), url: 'u'}), e => e.code === 'RH_ACTION' && /invalida/.test(e.message));
});
