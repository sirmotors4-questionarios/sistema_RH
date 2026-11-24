import {campo, chaveRH} from './rh-modulos.js';
import {prepararRecalculoLancamento} from './rh-folha.js';

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
