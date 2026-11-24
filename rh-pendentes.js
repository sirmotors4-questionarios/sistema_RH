import {campo, chaveRH} from './rh-modulos.js';
import {prepararRecalculoLancamento} from './rh-folha.js';

// Só prepara condições de pagamento; não copia entradas manuais nem muda estados.
export function prepararPendentes({lancamentos, funcionarios, adendas, parametros, ano, mes, utilizador}) {
  const porChave = new Map(), vistos = new Set(), linhas = [], erros = [];
  let preservados = 0;
  for (const f of funcionarios) {
    const k = chaveRH(campo(f, 'Empresa'), campo(f, 'Codigo'));
    porChave.set(k, [...(porChave.get(k) || []), f]);
  }
  for (const original of lancamentos) {
    if (String(campo(original, 'Estado do Lançamento')).trim().toUpperCase() !== 'PENDENTE') { preservados++; continue; }
    const k = chaveRH(campo(original, 'Empresa'), campo(original, 'Codigo'));
    const fs = porChave.get(k) || [];
    if (vistos.has(k) || fs.length !== 1) { erros.push(k + ': lançamento/cadastro inexistente ou duplicado.'); continue; }
    vistos.add(k);
    const r = prepararRecalculoLancamento({original, funcionario: fs[0], adendas, parametros, ano, mes, utilizador});
    if (r.erros.length) erros.push(k + ': ' + r.erros.join(' '));
    else linhas.push(r.linha);
  }
  return {linhas: erros.length ? [] : linhas, erros, preservados};
}
