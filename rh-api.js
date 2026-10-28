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

// Dois modos de autenticação:
//  - getToken: token Microsoft Entra (fluxos "Any user in my tenant");
//  - credenciais {email, chave}: código pessoal (fluxo RH_API_Acesso), sem login Microsoft.
export async function callApi(action, {getToken, credenciais, params, fetchImpl = globalThis.fetch, url = RH_CONFIG.apiUrl} = {}) {
  const headers = {'Content-Type': 'application/json'};
  // params: filtros opcionais (ex.: {ano, mes} para lançamentos); não podem substituir a acção nem as credenciais.
  const pedido = {...(params || {}), action};
  if (credenciais) {
    if (!credenciais.email || !credenciais.chave) throw new ApiError('Inicie sessão com o seu e-mail e código de acesso.', 0, 'RH_AUTH');
    if (!url) throw new ApiError('Falta o endereço do fluxo de acesso.', 0, 'RH_AUTH');
    pedido.email = credenciais.email;
    pedido.chave = credenciais.chave;
  } else {
    if (typeof getToken !== 'function') throw new ApiError('Autenticação indisponível.', 0, 'RH_AUTH');
    headers.Authorization = 'Bearer ' + await getToken();
  }
  let response;
  try {
    response = await fetchImpl(url, {method: 'POST', headers, body: JSON.stringify(pedido)});
  } catch {
    throw new ApiError('Não foi possível contactar o Power Automate. Verifique a ligação.', 0, 'RH_NETWORK');
  }
  let body = null;
  try { body = await response.json(); } catch { body = null; }
  if (response.status === 401) throw new ApiError('Sessão Microsoft inválida ou expirada. Inicie sessão novamente.', 401, 'RH_AUTH');
  if (response.status === 403) {
    if (credenciais && body?.erro) throw new ApiError(body.erro, 403, 'RH_FORBIDDEN');
    if (credenciais) throw new ApiError('E-mail ou código de acesso inválidos, ou conta não activa em tbl_utilizadores_RH.', 403, 'RH_FORBIDDEN');
    const quem = body?.email ? ' (' + body.email + ')' : '';
    throw new ApiError('A sua conta' + quem + ' não está autorizada no Sistema RH. Peça a um Superadministrador para a registar em tbl_utilizadores_RH.', 403, 'RH_FORBIDDEN');
  }
  if (response.status === 400) throw new ApiError(body?.erro || 'Pedido inválido.', 400, 'RH_ACTION');
  if (!response.ok && body?.erro) throw new ApiError(body.erro, response.status, response.status === 409 ? 'RH_CONFLITO' : 'RH_NAO_ENCONTRADO');
  if (!response.ok || !body || body.ok !== true) {
    throw new ApiError('O Power Automate devolveu uma resposta inválida (' + response.status + ').', response.status, 'RH_RESPONSE');
  }
  return body;
}

export async function loadConfiguracoes(deps = {}) {
  const body = await callApi('configuracoes', deps);
  // O fluxo de acesso por código devolve {tabelas:[{tabela, linhas}]}; o fluxo Entra devolve as tabelas no topo.
  if (Array.isArray(body.tabelas)) {
    for (const item of body.tabelas) if (item && typeof item.tabela === 'string' && Array.isArray(item.linhas)) body[item.tabela] = item.linhas;
  }
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

// Gravação directa de um colaborador (fluxo RH_API_Actualizar). modo: 'criar' | 'editar'.
// O colaborador é identificado por Empresa + Código (alvo); dados usa os cabeçalhos reais da tabela.
// O fluxo só aceita perfis Superadministrador ou Administrador RH.
export async function gravarColaborador(modo, {alvo, dados} = {}, deps = {}) {
  if (!['criar', 'editar'].includes(modo)) throw new ApiError('Modo de gravação inválido.', 0, 'RH_ACTION');
  if (!alvo?.Empresa || !alvo?.Codigo) throw new ApiError('Falta a Empresa ou o Código do colaborador.', 0, 'RH_ACTION');
  const body = await callApi('colaborador_gravar', {...deps, params: {modo, alvo, dados}});
  return {modo: body.modo || modo, codigo: body.codigo || alvo.Codigo, empresa: body.empresa || alvo.Empresa};
}

// Novo mês (fluxo RH_API_Actualizar, acção lancamento_novo_mes): o site prepara as linhas com as regras do VBA
// (rh-folha.js) e o fluxo responde 202 logo que aceita o pedido; a gravação continua no Power Automate.
export async function gravarNovoMes({ano, mes, mesNome, linhas} = {}, deps = {}) {
  if (!Array.isArray(linhas) || !linhas.length) throw new ApiError('Não há lançamentos para gravar.', 0, 'RH_ACTION');
  const body = await callApi('lancamento_novo_mes', {...deps, params: {ano: Number(ano), mes: Number(mes), mesNome, linhas}});
  return {total: Number(body.total) || linhas.length, modo: body.modo};
}

// Edição de um lançamento do mês (Office Script, modo "editar"): só as entradas mensais e as condições
// sincronizadas com o contrato. O fluxo responde com o resultado do script ({ok, alteradas}).
export async function gravarLancamento({ano, mes, mesNome, linha} = {}, deps = {}) {
  if (!linha || !linha.Empresa || !linha.Codigo) throw new ApiError('Lançamento sem Empresa/Código.', 0, 'RH_ACTION');
  const body = await callApi('lancamento_editar', {...deps, params: {ano: Number(ano), mes: Number(mes), mesNome, linhas: [linha]}});
  // O script devolve a linha recalculada (body.linha); sem ela o site relê o mês.
  let linhaExcel = body.linha;
  if (typeof linhaExcel === 'string') { try { linhaExcel = JSON.parse(linhaExcel); } catch { linhaExcel = null; } }
  if (!linhaExcel || typeof linhaExcel !== 'object' || Array.isArray(linhaExcel)) linhaExcel = null;
  return {alteradas: Number(body.alteradas) || 1, linha: linhaExcel};
}

// Finalizar mês (modo "finalizar"): todos os lançamentos do mês passam a Finalizado / Revisto = Sim.
export async function finalizarMes({ano, mes, mesNome, utilizador = '', agoraSerie = ''} = {}, deps = {}) {
  if (!mesNome) throw new ApiError('Mês inválido.', 0, 'RH_ACTION');
  const body = await callApi('lancamento_finalizar', {...deps, params: {ano: Number(ano), mes: Number(mes), mesNome,
    linhas: [{'Utilizador Alteracao': utilizador, 'Data Ultima Alteracao': agoraSerie}]}});
  return {alteradas: Number(body.alteradas) || 0};
}

// ---------- Indemnizações (fluxo RH_API_Indemnizacoes → Office Script RHIndemnizacoes) ----------
// O script devolve sempre texto JSON; o fluxo responde {ok, resultado: "<json>"} ou já com o objecto.
function resultadoScript(body) {
  let r = body?.resultado ?? body;
  if (typeof r === 'string') { try { r = JSON.parse(r); } catch { throw new ApiError('Resposta do script inválida.', 200, 'RH_RESPONSE'); } }
  if (!r || r.ok !== true) throw new ApiError(r?.erro || 'O Excel recusou o pedido.', 409, 'RH_CONFLITO');
  return r;
}
const listaJson = v => { if (Array.isArray(v)) return v; try { const x = JSON.parse(v || '[]'); return Array.isArray(x) ? x : []; } catch { return []; } };
export async function indemnizacoesListar(deps = {}) {
  const r = resultadoScript(await callApi('indemnizacoes_listar', deps));
  return {indemnizacoes: listaJson(r.indemnizacoes), prestacoes: listaJson(r.prestacoes)};
}
export async function indemnizacaoGravar(pedido, deps = {}) {
  return resultadoScript(await callApi('indemnizacao_gravar', {...deps, params: {dados: JSON.stringify(pedido)}}));
}
export async function prestacaoPagar({id, paga = true, data = ''}, deps = {}) {
  return resultadoScript(await callApi('prestacao_pagar', {...deps, params: {dados: JSON.stringify({id, paga, data})}}));
}
export async function indemnizacaoAnular({id, observacoes = ''}, deps = {}) {
  return resultadoScript(await callApi('indemnizacao_anular', {...deps, params: {dados: JSON.stringify({id, observacoes})}}));
}
