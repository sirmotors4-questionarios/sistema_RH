// Acesso por código pessoal (independente do login Microsoft).
// - O endereço do fluxo RH_API_Acesso contém uma assinatura secreta: é colado pelo utilizador e fica
//   apenas neste navegador (localStorage). Nunca é escrito no código publicado no GitHub.
// - O código pessoal nunca sai do navegador: é enviado apenas SHA-256(email + ":" + código),
//   que o fluxo compara com a coluna Codigo_Hash de tbl_utilizadores_RH.
// - A sessão (e-mail + chave) fica em sessionStorage e termina ao fechar o separador.

const KEY_URL = 'rh.acesso.url';
const KEY_SESSAO = 'rh.acesso.sessao';
export const CODIGO_MINIMO = 8;

function store(tipo) {
  try { return globalThis[tipo] || null; } catch { return null; }
}
function ler(tipo, chave) {
  try { return store(tipo)?.getItem(chave) ?? null; } catch { return null; }
}
function escrever(tipo, chave, valor) {
  try { valor == null ? store(tipo)?.removeItem(chave) : store(tipo)?.setItem(chave, valor); return true; } catch { return false; }
}

export const normalizarEmail = email => String(email ?? '').trim().toLowerCase();

export function validarEmail(email) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(normalizarEmail(email));
}

// Aceita apenas endereços de gatilho HTTP do Power Automate / Logic Apps com assinatura.
export function validarUrl(valor) {
  let url;
  try { url = new URL(String(valor ?? '').trim()); } catch { return false; }
  const host = url.hostname.toLowerCase();
  return url.protocol === 'https:' &&
    (host.endsWith('.powerplatform.com') || host.endsWith('.logic.azure.com')) &&
    url.searchParams.has('sig');
}

export async function chaveAcesso(email, codigo, subtle = globalThis.crypto?.subtle) {
  if (!subtle) throw Error('Este navegador não suporta o cálculo seguro da chave (Web Crypto).');
  const dados = new TextEncoder().encode(normalizarEmail(email) + ':' + String(codigo ?? ''));
  const hash = new Uint8Array(await subtle.digest('SHA-256', dados));
  return [...hash].map(b => b.toString(16).padStart(2, '0')).join('');
}

export function validarCodigo(codigo) {
  return String(codigo ?? '').length >= CODIGO_MINIMO;
}

export function urlGuardada() {
  const url = ler('localStorage', KEY_URL);
  return validarUrl(url) ? url : '';
}

export function guardarUrl(url) {
  if (!validarUrl(url)) throw Error('Endereço do fluxo inválido. Cole o URL completo do gatilho HTTP (contém "sig=").');
  escrever('localStorage', KEY_URL, String(url).trim());
}

export function esquecerUrl() {
  escrever('localStorage', KEY_URL, null);
}

export function sessaoAcesso() {
  try {
    const s = JSON.parse(ler('sessionStorage', KEY_SESSAO) || 'null');
    return s && validarEmail(s.email) && /^[0-9a-f]{64}$/.test(s.chave) ? s : null;
  } catch { return null; }
}

export async function entrarComCodigo({email, codigo, url}, deps = {}) {
  if (!validarEmail(email)) throw Error('Indique um e-mail válido.');
  if (!validarCodigo(codigo)) throw Error('O código de acesso tem pelo menos ' + CODIGO_MINIMO + ' caracteres.');
  if (url) guardarUrl(url);
  if (!urlGuardada()) throw Error('Falta o endereço do fluxo RH_API_Acesso.');
  const sessao = {email: normalizarEmail(email), chave: await chaveAcesso(email, codigo, deps.subtle)};
  escrever('sessionStorage', KEY_SESSAO, JSON.stringify(sessao));
  return sessao;
}

export function sairAcesso() {
  escrever('sessionStorage', KEY_SESSAO, null);
}

// Endereço do fluxo de gravação (RH_API_Actualizar): também secreto, também só neste navegador.
const KEY_URL_GRAVAR = 'rh.acesso.url.gravar';
export function urlGravacao() {
  const url = ler('localStorage', KEY_URL_GRAVAR);
  return validarUrl(url) ? url : '';
}
export function guardarUrlGravacao(url) {
  if (!validarUrl(url)) throw Error('Endereço do fluxo de gravação inválido. Cole o URL completo do gatilho HTTP (contém "sig=").');
  if (String(url).trim() === urlGuardada()) throw Error('Este é o endereço do fluxo de consulta; cole o do RH_API_Actualizar.');
  escrever('localStorage', KEY_URL_GRAVAR, String(url).trim());
}

// Endereço do fluxo RH_API_Indemnizacoes: secreto, só neste navegador.
const KEY_URL_INDEMN = 'rh.acesso.url.indemnizacoes';
export function urlIndemnizacoes() {
  const url = ler('localStorage', KEY_URL_INDEMN);
  return validarUrl(url) ? url : '';
}
export function guardarUrlIndemnizacoes(url) {
  if (!validarUrl(url)) throw Error('Endereço do fluxo inválido. Cole o URL completo do gatilho HTTP do RH_API_Indemnizacoes (contém "sig=").');
  const u = String(url).trim();
  if (u === urlGuardada() || u === urlGravacao()) throw Error('Este endereço é de outro fluxo; cole o do RH_API_Indemnizacoes.');
  escrever('localStorage', KEY_URL_INDEMN, u);
}
