import test from 'node:test';
import assert from 'node:assert/strict';
import {createHash} from 'node:crypto';

function memoria() { const m = new Map(); return {getItem: k => m.has(k) ? m.get(k) : null, setItem: (k, v) => m.set(k, String(v)), removeItem: k => m.delete(k)}; }
globalThis.localStorage = memoria();
globalThis.sessionStorage = memoria();
const A = await import('../rh-acesso.js');
const {callApi, loadModulo} = await import('../rh-api.js');
const URL_OK = 'https://default00.c9.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/abc/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=TESTE';

test('chave = SHA-256 de email normalizado + ":" + código (igual ao que o administrador regista)', async () => {
  const esperado = createHash('sha256').update('pessoa@sir.test:codigo-123').digest('hex');
  assert.equal(await A.chaveAcesso('  Pessoa@SIR.test ', 'codigo-123'), esperado);
  assert.notEqual(await A.chaveAcesso('pessoa@sir.test', 'codigo-124'), esperado);
});

test('só aceita URLs de gatilho Power Automate/Logic Apps com assinatura', () => {
  assert.equal(A.validarUrl(URL_OK), true);
  assert.equal(A.validarUrl('https://prod-12.westeurope.logic.azure.com/workflows/x/triggers/manual/paths/invoke?sig=y'), true);
  assert.equal(A.validarUrl(URL_OK.replace('&sig=TESTE', '')), false);
  assert.equal(A.validarUrl(URL_OK.replace('https:', 'http:')), false);
  assert.equal(A.validarUrl('https://evil.example.com/?sig=1'), false);
  assert.equal(A.validarUrl('não é url'), false);
});

test('entrar valida dados, guarda URL no navegador e sessão sem o código', async () => {
  await assert.rejects(A.entrarComCodigo({email: 'x', codigo: '12345678', url: URL_OK}), /e-mail/);
  await assert.rejects(A.entrarComCodigo({email: 'a@b.co', codigo: 'curto', url: URL_OK}), /8 caracteres/);
  await assert.rejects(A.entrarComCodigo({email: 'a@b.co', codigo: '12345678'}), /endereço/);
  const s = await A.entrarComCodigo({email: 'A@b.co', codigo: '12345678', url: URL_OK});
  assert.equal(s.email, 'a@b.co');
  assert.match(s.chave, /^[0-9a-f]{64}$/);
  assert.equal(A.urlGuardada(), URL_OK);
  assert.deepEqual(A.sessaoAcesso(), s);
  assert.ok(!sessionStorage.getItem('rh.acesso.sessao').includes('12345678'));
  A.sairAcesso();
  assert.equal(A.sessaoAcesso(), null);
  assert.equal(A.urlGuardada(), URL_OK);
});

test('callApi com credenciais: sem Authorization, envia email+chave; 403 com mensagem genérica', async () => {
  let pedido;
  const ok = async (url, init) => { pedido = {url, init}; return {status: 200, ok: true, json: async () => ({ok: true, action: 'ferias', utilizador: {Perfil: 'Gestor'}, tabelas: [{tabela: 'tbl_ferias', linhas: []}, {tabela: 'tbl_feriados', linhas: []}]})}; };
  const credenciais = {email: 'a@b.co', chave: 'f'.repeat(64)};
  const r = await loadModulo('ferias', {credenciais, url: URL_OK, fetchImpl: ok});
  assert.equal(pedido.url, URL_OK);
  assert.equal(pedido.init.headers.Authorization, undefined);
  assert.deepEqual(JSON.parse(pedido.init.body), {action: 'ferias', email: 'a@b.co', chave: 'f'.repeat(64)});
  assert.equal(r.utilizador.perfil, 'Gestor');
  const proibido = async () => ({status: 403, ok: false, json: async () => ({ok: false, email: 'a@b.co'})});
  await assert.rejects(callApi('ferias', {credenciais, url: URL_OK, fetchImpl: proibido}), e => e.code === 'RH_FORBIDDEN' && /código de acesso/.test(e.message) && !/a@b\.co/.test(e.message));
  await assert.rejects(callApi('ferias', {credenciais: {email: 'a@b.co'}, url: URL_OK, fetchImpl: ok}), e => e.code === 'RH_AUTH');
});

test('params vão no corpo sem substituir acção; configuracoes aceita formato {tabelas:[…]}', async () => {
  const {loadConfiguracoes} = await import('../rh-api.js');
  let corpo;
  const tabelas = ['tbl_configuracoes_RH', 'tbl_escaloes_RH', 'tbl_formulas_RH', 'tbl_auditoria_RH'].map(t => ({tabela: t, linhas: []}));
  const f = async (url, init) => { corpo = JSON.parse(init.body); return {status: 200, ok: true, json: async () => ({ok: true, utilizador: {Perfil: 'Superadministrador'}, tabelas})}; };
  const credenciais = {email: 'a@b.co', chave: 'f'.repeat(64)};
  await callApi('lancamentos', {credenciais, url: URL_OK, fetchImpl: f, params: {ano: 2026, mes: 10, action: 'hack', chave: 'x'}});
  assert.deepEqual(corpo, {ano: 2026, mes: 10, action: 'lancamentos', email: 'a@b.co', chave: 'f'.repeat(64)});
  const r = await loadConfiguracoes({credenciais, url: URL_OK, fetchImpl: f});
  assert.equal(corpo.action, 'configuracoes');
  assert.deepEqual(Object.keys(r.tables).sort(), ['tbl_auditoria_RH', 'tbl_configuracoes_RH', 'tbl_escaloes_RH', 'tbl_formulas_RH']);
  assert.equal(r.utilizador.aprovar, true);
});
