// Colunas da tabela de lançamentos: estrutura padrão, colunas extra vindas de tbl_lançamento,
// ordenação por cabeçalho (conforme o tipo) e preferências guardadas neste navegador.
import * as M from "./rh-modulos.js";

const texto = v => String(v ?? "").trim();

// Estrutura padrão (a que o sistema mostra de origem). «fonte» é o cabeçalho do Excel que a coluna já mostra,
// para não aparecer repetido na lista de colunas a adicionar.
export const PADRAO = Object.freeze([
  {id: "codigo", rotulo: "Código", tipo: "texto", fonte: ["Codigo", "Código"], valor: l => l.codigo},
  {id: "nome", rotulo: "Nome", tipo: "texto", fonte: ["Nome"], valor: l => l.nome, fixa: true},
  {id: "empresa", rotulo: "Empresa", tipo: "texto", fonte: ["Empresa"], valor: l => l.empresa},
  {id: "departamento", rotulo: "Departamento", tipo: "texto", fonte: ["Departamento"], valor: l => l.departamento},
  {id: "regime", rotulo: "Regime", tipo: "texto", fonte: ["RegimePagamento"], valor: l => l.regime},
  {id: "dias", rotulo: "Dias", tipo: "numero", inteiro: true, fonte: ["Dias Trabalhados"], valor: l => l.dias},
  {id: "base", rotulo: "Vencimento", tipo: "numero", total: true, fonte: ["Vencimento"], valor: l => l.base},
  {id: "acrescimos", rotulo: "Acréscimos", tipo: "numero", total: true, fonte: ["Total de Acréscimos"], valor: l => l.acrescimos},
  {id: "descontos", rotulo: "Descontos", tipo: "numero", total: true, fonte: ["Total de Descontos"], valor: l => l.descontos},
  {id: "inssColab", rotulo: "INSS colab.", tipo: "numero", total: true, fonte: ["INSS Colaborador"], valor: l => l.inssColab},
  {id: "irps", rotulo: "IRPS", tipo: "numero", total: true, fonte: ["IRPS"], valor: l => l.irps},
  {id: "liquido", rotulo: "Líquido", tipo: "numero", total: true, destaque: true, fonte: ["Salario Liquido", "Salário Líquido"], valor: l => l.liquido},
  {id: "estado", rotulo: "Estado", tipo: "estado", fonte: ["Estado do Lançamento"], valor: l => l.estado}
]);
export const PADRAO_IDS = Object.freeze(PADRAO.map(c => c.id));

const SEM_TOTAL = /^(ano|m[eê]s|c[oó]digo|n[º.o]|id\b)|taxa|%|percent|escal[aã]o|categoria/i;
const chave = s => String(s).toLocaleLowerCase("pt").normalize("NFD").replace(/[̀-ͯ]/g, "").replace(/[\s_]/g, "");

// Tipo de uma coluna do Excel, pelo nome e pelos valores: data, número ou texto.
export function tipoColuna(cabecalho, linhas) {
  const vals = linhas.map(l => l[cabecalho]).filter(v => v !== "" && v != null).slice(0, 200);
  if (/data|^dt[\s_]/i.test(cabecalho)) return "data";
  if (!vals.length) return "texto";
  const nums = vals.filter(v => typeof v === "number" || /^-?[\d\s.,]+$/.test(String(v).trim()) && M.numero(v) !== null).length;
  return nums / vals.length >= 0.8 ? "numero" : "texto";
}

// Todas as colunas possíveis: as da estrutura padrão + as restantes colunas de tbl_lançamento.
export function catalogo(linhasExcel = []) {
  const usadas = new Set(PADRAO.flatMap(c => c.fonte.map(chave)));
  const extra = M.colunas(linhasExcel)
    .filter(h => !usadas.has(chave(h)))
    .map(h => {
      const tipo = tipoColuna(h, linhasExcel);
      return {id: "x:" + h, rotulo: h, tipo, extra: true, total: tipo === "numero" && !SEM_TOTAL.test(h), valor: l => l.linha?.[h]};
    });
  return [...PADRAO, ...extra];
}

// Preferências: {colunas: [ids por ordem], ordem: {id, dir: "asc"|"desc"} | null}. Ids desconhecidos são ignorados;
// o Nome nunca sai (identifica a linha e mostra o estado da gravação).
export function normalizar(pref, cat) {
  const ids = new Set(cat.map(c => c.id));
  let cols = Array.isArray(pref?.colunas) ? pref.colunas.filter((id, i, a) => ids.has(id) && a.indexOf(id) === i) : [...PADRAO_IDS];
  if (!cols.includes("nome")) cols.splice(Math.min(1, cols.length), 0, "nome");
  const o = pref?.ordem;
  const ordem = o && cols.includes(o.id) && (o.dir === "asc" || o.dir === "desc") ? {id: o.id, dir: o.dir} : null;
  return {colunas: cols, ordem};
}

// Clique no cabeçalho: crescente → decrescente → sem ordenação (ordem original).
export function proximaOrdem(ordem, id) {
  if (!ordem || ordem.id !== id) return {id, dir: "asc"};
  return ordem.dir === "asc" ? {id, dir: "desc"} : null;
}

function chaveOrdenacao(col, l) {
  const v = col.valor(l);
  if (col.tipo === "numero") return M.numero(v);
  if (col.tipo === "data") return M.data(v) || null;
  const s = texto(v);
  return s === "" ? null : s;
}

// Ordena sem alterar a lista original. Vazios ficam sempre no fim; empates mantêm a ordem original.
export function ordenar(lista, col, dir) {
  if (!col || !dir) return lista;
  const f = dir === "desc" ? -1 : 1;
  const coll = new Intl.Collator("pt", {numeric: true, sensitivity: "base"});
  return lista.map((l, i) => ({l, i, k: chaveOrdenacao(col, l)})).sort((a, b) => {
    if (a.k === null || b.k === null) return a.k === b.k ? a.i - b.i : a.k === null ? 1 : -1;
    const c = typeof a.k === "number" ? a.k - b.k : coll.compare(String(a.k), String(b.k));
    return c ? c * f : a.i - b.i;
  }).map(x => x.l);
}

export function formatar(col, l) {
  const v = col.valor(l);
  if (col.tipo === "numero") { const s = M.formatarValor(v); return col.inteiro ? s.replace(/,00$/, "") : s; }
  if (col.tipo === "data") return M.formatarData(v);
  return texto(v);
}

export function somar(lista, col) {
  return lista.reduce((s, l) => s + (M.numero(col.valor(l)) ?? 0), 0);
}

// ---------- preferências neste navegador ----------
const CHAVE = "sirRH.lancamentos.colunas.v1";
export function lerPreferencias(storage = globalThis.localStorage) {
  try { return JSON.parse(storage.getItem(CHAVE) || "null"); } catch { return null; }
}
export function guardarPreferencias(pref, storage = globalThis.localStorage) {
  try { storage.setItem(CHAVE, JSON.stringify(pref)); } catch { /* navegador sem armazenamento: fica só nesta visita */ }
}
export function eDefault(pref) {
  return !pref.ordem && pref.colunas.length === PADRAO_IDS.length && pref.colunas.every((id, i) => id === PADRAO_IDS[i]);
}
