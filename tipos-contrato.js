// Tipos de contrato: contagem a partir do cadastro (tbl_funcionarios) mais os tipos propostos na sessão.
import * as M from "./rh-modulos.js";

export const SEM_TIPO = "Sem tipo de contrato";

export function contarTiposContrato(funcionarios = [], propostos = []) {
  const mapa = new Map();
  const chave = n => n.toLocaleLowerCase("pt").normalize("NFD").replace(/[̀-ͯ]/g, "");
  let activos = 0, semTipo = 0;
  for (const l of funcionarios) {
    const nome = String(M.campo(l, "TipoContrato", "Tipo Contrato", "Tipo de Contrato") ?? "").trim();
    const activo = M.activo(M.campo(l, "Estado"));
    if (activo) activos += 1;
    if (!nome) { if (activo) semTipo += 1; continue; }
    const k = chave(nome);
    if (!mapa.has(k)) mapa.set(k, {nome, activos: 0, inactivos: 0, total: 0, proposto: false});
    const t = mapa.get(k);
    t.total += 1; if (activo) t.activos += 1; else t.inactivos += 1;
  }
  for (const p of propostos) {
    const k = chave(p.nome);
    if (mapa.has(k)) continue;
    mapa.set(k, {...p, activos: 0, inactivos: 0, total: 0, proposto: true});
  }
  const tipos = [...mapa.values()].sort((a, b) => (a.proposto - b.proposto) || (b.total - a.total) || a.nome.localeCompare(b.nome, "pt"));
  return {tipos, activos, semTipo};
}
