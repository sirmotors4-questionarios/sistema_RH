import test from 'node:test';
import assert from 'node:assert/strict';
import * as A from '../acessos-ui.js';

const mem = () => { const m = new Map(); return {getItem: k => m.has(k) ? m.get(k) : null, setItem: (k, v) => m.set(k, String(v)), removeItem: k => m.delete(k)}; };
globalThis.localStorage = mem(); globalThis.sessionStorage = mem();
const R = await import('../rh-acesso.js');

test('utilizadores e filtro', () => {
  const l = A.utilizadores([{Nome: 'Zé', Email: 'Ze@Sir.com', Perfil: 'Gestor', Estado: 'Inactivo', Tem_Codigo: true, Trocar_Codigo: 'Sim'},
    {Nome: 'Ana', Email: 'ana@sir.com', Perfil: 'Superadministrador', Estado: 'Activo', Tem_Codigo: false}, {Nome: 'sem email'}]);
  assert.equal(l.length, 2);
  assert.equal(l[0].email, 'ze@sir.com'); assert.equal(l[0].trocar, true);
  assert.deepEqual(A.filtrar(l).map(u => u.nome), ['Ana', 'Zé']);
  assert.deepEqual(A.filtrar(l, {procurar: 'ZE'}).map(u => u.nome), ['Zé']);
  assert.deepEqual(A.filtrar(l, {estado: 'Activo'}).map(u => u.nome), ['Ana']);
});

test('validação de novo utilizador', () => {
  const lista = [{email: 'ana@sir.com'}];
  assert.deepEqual(A.validar({nome: 'B', email: 'b@sir.com', perfil: 'Gestor'}, {novo: true, lista, codigo: 'abcdefgh'}), []);
  const e = A.validar({nome: '', email: 'ana@sir.com', perfil: 'Chefe'}, {novo: true, lista, codigo: 'x'});
  assert.equal(e.length, 4);
});

test('código temporário e link de primeiro acesso', () => {
  const c = R.gerarCodigoTemporario(12);
  assert.equal(c.length, 12); assert.ok(!/[0O1lI]/.test(c));
  const a = 'https://x.powerplatform.com/a?sig=1', i = 'https://x.powerplatform.com/i?sig=2', g = 'https://x.powerplatform.com/g?sig=3';
  R.guardarUrl(a); R.guardarUrlIndemnizacoes(i); R.guardarUrlGravacao(g);
  const sem = R.linkConfiguracao('https://site/x/index.html#y', {email: 'Ana@Sir.com'});
  const com = R.linkConfiguracao('https://site/x/', {gravacao: true});
  assert.ok(sem.startsWith('https://site/x/index.html#config='));
  localStorage.removeItem('rh.acesso.url'); localStorage.removeItem('rh.acesso.url.indemnizacoes'); localStorage.removeItem('rh.acesso.url.gravar');
  const r = R.lerConfiguracaoDoLink('#' + sem.split('#')[1]);
  assert.deepEqual(r, {guardados: ['acesso', 'gestão'], email: 'ana@sir.com'});
  assert.equal(R.urlGuardada(), a); assert.equal(R.urlIndemnizacoes(), i); assert.equal(R.urlGravacao(), '');
  R.lerConfiguracaoDoLink('#' + com.split('#')[1]);
  assert.equal(R.urlGravacao(), g);
  assert.equal(R.lerConfiguracaoDoLink('#entrar'), null);
  assert.throws(() => R.lerConfiguracaoDoLink('#config=@@@'.replace('@@@', 'eyJh')), /link/);
});

test('troca da chave na sessão', async () => {
  const s = await R.entrarComCodigo({email: 'a@sir.com', codigo: '12345678'});
  assert.equal(R.actualizarChaveSessao('zz'), false);
  assert.equal(R.actualizarChaveSessao('f'.repeat(64)), true);
  assert.equal(R.sessaoAcesso().chave, 'f'.repeat(64)); assert.equal(R.sessaoAcesso().email, s.email);
});
