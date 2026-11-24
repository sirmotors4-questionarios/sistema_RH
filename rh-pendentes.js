import {campo, chaveRH, normalizarCodigo} from './rh-modulos.js';
import {prepararRecalculoLancamento, prepararNovoMes, contratoNoPeriodo, periodoFolha} from './rh-folha.js';

// Só prepara condições de pagamento; não copia entradas manuais nem muda estados.
// Um lançamento bloqueado (cadastro em falta/duplicado, contrato fora do período…) fica de fora e é indicado em
// `erros`; os restantes Pendentes seguem no lote. Os bloqueados não são escritos e continuam Pendentes.
export function prepararPendentes({lancamentos, funcionarios, adendas, parametros, ano, mes, utilizador}) {
  const porChave = new Map(), porLancamento = new Map(), erros = [], linhas = [];
  let preservados = 0;
  for (const f of funcionarios) {
    const k = chaveRH(campo(f, 'Empresa'), campo(f, 'Codigo'));
    porChave.set(k, [...(porChave.get(k) || []), f]);
  }
  for (const original of lancamentos) {
    if (String(campo(original, 'Estado do Lançamento')).trim().toUpperCase() !== 'PENDENTE') { preservados++; continue; }
    const k = chaveRH(campo(original, 'Empresa'), campo(original, 'Codigo'));
    porLancamento.set(k, [...(porLancamento.get(k) || []), original]);
  }
  for (const [k, originais] of porLancamento) {
    const fs = porChave.get(k) || [];
    if (originais.length !== 1) { erros.push(k + ': lançamento repetido neste mês.'); continue; }
    if (fs.length !== 1) { erros.push(k + ': ' + (fs.length ? 'colaborador duplicado no cadastro.' : 'colaborador inexistente no cadastro.')); continue; }
    const r = prepararRecalculoLancamento({original: originais[0], funcionario: fs[0], adendas, parametros, ano, mes, utilizador});
    if (r.erros.length) erros.push(k + ': ' + r.erros.join(' '));
    else linhas.push(r.linha);
  }
  return {linhas, erros, preservados};
}

// Sincronização do mês ao carregar: retira os Pendentes sem contrato no período, recalcula os restantes
// Pendentes e introduz os colaboradores com contrato vigente que ainda não têm lançamento no mês.
// Alterados e Finalizados nunca são retirados nem recalculados; se já não tiverem contrato, ficam em `aRever`.
export function prepararSincronizacao({lancamentos, lancamentosAnterior = [], funcionarios, adendas = [], parametros = [], ano, mes, utilizador = "", agora}) {
  const per = periodoFolha(Number(ano), Number(mes));
  const porChave = new Map();
  for (const f of funcionarios) {
    const k = chaveRH(campo(f, 'Empresa'), campo(f, 'Codigo'));
    porChave.set(k, [...(porChave.get(k) || []), f]);
  }
  const contagem = new Map();
  for (const l of lancamentos) { const k = chaveRH(campo(l, 'Empresa'), campo(l, 'Codigo')); contagem.set(k, (contagem.get(k) || 0) + 1); }
  const retirar = [], aRever = [], ficam = [];
  for (const l of lancamentos) {
    const k = chaveRH(campo(l, 'Empresa'), campo(l, 'Codigo')), fs = porChave.get(k) || [];
    // Só se retira quando o cadastro é inequívoco e diz que não há contrato no período.
    const c = contagem.get(k) === 1 && fs.length === 1 ? contratoNoPeriodo(fs[0], adendas, per) : null;
    if (!c || c.vigente) { ficam.push(l); continue; }
    const pendente = String(campo(l, 'Estado do Lançamento')).trim().toUpperCase() === 'PENDENTE';
    const item = {Empresa: String(campo(l, 'Empresa')).trim(), Codigo: normalizarCodigo(campo(l, 'Codigo')), chave: k,
      nome: String(campo(l, 'Nome')).trim(), estado: String(campo(l, 'Estado do Lançamento')).trim(), motivo: c.motivo};
    if (pendente) retirar.push(item); else { aRever.push(item); ficam.push(l); }
  }
  const recalculo = prepararPendentes({lancamentos: ficam, funcionarios, adendas, parametros, ano, mes, utilizador});
  // Colaboradores a introduzir: as mesmas regras do Novo mês, só para quem não tem lançamento no mês.
  const alvo = prepararNovoMes({ano, mes, funcionarios, adendas, parametros, lancamentosMes: [], lancamentosAnterior, utilizador, ...(agora ? {agora} : {})});
  // Quem tem dados em falta não entra (fica nos avisos); um código repetido no cadastro nunca é introduzido.
  const existentes = new Set(contagem.keys());
  const introduzir = alvo.linhas.filter(l => { const k = chaveRH(l.Empresa, l.Codigo); return !existentes.has(k) && (porChave.get(k) || []).length === 1; });
  const avisosIntroducao = alvo.erros;
  return {retirar, aRever, atualizacoes: recalculo.linhas, erros: recalculo.erros, preservados: recalculo.preservados,
    introduzir, avisosIntroducao, periodo: per};
}
