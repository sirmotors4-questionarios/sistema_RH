import {RH_CONFIG} from './rh-config.js';
import {readTables} from './rh-data.js';

// Perfis definidos em tbl_utilizadores_RH. A interface usa-os apenas para orientar o utilizador;
// a autorização efectiva é verificada no fluxo Power Automate.
const PERFIS = {
  'superadministrador': {perfil: 'Superadministrador', consultar: true, propor: true, aprovar: true},
  'administrador rh': {perfil: 'Administrador RH', consultar: true, propor: true, aprovar: false},
  'tecnico rh': {perfil: 'Técnico RH', consultar: true, propor: false, aprovar: false},
  'gestor': {perfil: 'Gestor', consultar: true, propor: false, aprovar: false}
};

const key = value => String(value ?? '').normalize('NFD').replace(/[̀-ͯ]/g, '').trim().toLowerCase();

export function permissoes(perfil) {
  return PERFIS[key(perfil)] || {perfil: String(perfil ?? ''), consultar: false, propor: false, aprovar: false};
}

export class ApiError extends Error {
  constructor(message, status, code) {
    super(message);
    this.status = status;
    this.code = code;
  }
}

export async function callApi(action, {getToken, fetchImpl = globalThis.fetch, url = RH_CONFIG.apiUrl} = {}) {
  if (typeof getToken !== 'function') throw new ApiError('Autenticação indisponível.', 0, 'RH_AUTH');
  const token = await getToken();
  let response;
  try {
    response = await fetchImpl(url, {
      method: 'POST',
      headers: {'Content-Type': 'application/json', Authorization: 'Bearer ' + token},
      body: JSON.stringify({action})
    });
  } catch {
    throw new ApiError('Não foi possível contactar o Power Automate. Verifique a ligação.', 0, 'RH_NETWORK');
  }
  let body = null;
  try { body = await response.json(); } catch { body = null; }
  if (response.status === 401) throw new ApiError('Sessão Microsoft inválida ou expirada. Inicie sessão novamente.', 401, 'RH_AUTH');
  if (response.status === 403) {
    const quem = body?.email ? ' (' + body.email + ')' : '';
    throw new ApiError('A sua conta' + quem + ' não está autorizada no Sistema RH. Peça a um Superadministrador para a registar em tbl_utilizadores_RH.', 403, 'RH_FORBIDDEN');
  }
  if (!response.ok || !body || body.ok !== true) {
    throw new ApiError('O Power Automate devolveu uma resposta inválida (' + response.status + ').', response.status, 'RH_RESPONSE');
  }
  return body;
}

export async function loadConfiguracoes(deps) {
  const body = await callApi('configuracoes', deps);
  const user = body.utilizador || {};
  return {
    utilizador: {nome: user.Nome || '', email: user.Email || '', empresa: user.Empresa || '', ...permissoes(user.Perfil)},
    tables: readTables(body)
  };
}
