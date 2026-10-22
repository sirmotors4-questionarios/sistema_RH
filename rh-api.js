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
  if (response.status === 400) throw new ApiError(body?.erro || 'Pedido inválido.', 400, 'RH_ACTION');
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

// Módulos servidos pelo fluxo RH_API_Consultar e as tabelas que cada um devolve.
// A lista oficial está no próprio fluxo; aqui serve para validar a resposta.
export const MODULOS = Object.freeze({
  funcionarios: ['tbl_funcionarios'],
  contratos: ['tbl_adendas', 'tbl_parametros_pagamento'],
  lancamentos: ['tbl_lançamento', 'Table1', 'tbl_arredondamento'],
  ferias: ['tbl_ferias', 'tbl_feriados']
});

export async function loadModulo(action, deps = {}) {
  const esperadas = MODULOS[action];
  if (!esperadas) throw new ApiError('Módulo desconhecido: ' + action, 0, 'RH_ACTION');
  const body = await callApi(action, {url: RH_CONFIG.apiConsultarUrl, ...deps});
  if (!Array.isArray(body.tabelas)) throw new ApiError('Resposta sem tabelas.', 200, 'RH_RESPONSE');
  const tabelas = {};
  for (const item of body.tabelas) {
    if (!item || typeof item.tabela !== 'string' || !Array.isArray(item.linhas)) throw new ApiError('Tabela com formato inválido.', 200, 'RH_RESPONSE');
    tabelas[item.tabela] = item.linhas.map(linha => ({...linha}));
  }
  const faltam = esperadas.filter(nome => !(nome in tabelas));
  if (faltam.length) throw new ApiError('O fluxo não devolveu: ' + faltam.join(', ') + '.', 200, 'RH_RESPONSE');
  const user = body.utilizador || {};
  return {
    utilizador: {nome: user.Nome || '', email: user.Email || '', empresa: user.Empresa || '', ...permissoes(user.Perfil)},
    action,
    tabelas
  };
}
