/**
 * Office Script "RHIndemnizacoes" — chamado pelo fluxo RH_API_Indemnizacoes.
 * Parâmetros: accao, dados (texto JSON), utilizador (e-mail de quem pediu).
 * accao "preparar" — cria a folha "Indemnizacoes" e as tabelas tbl_indemnizacoes e tbl_indemnizacoes_prestacoes
 *                    se não existirem (não altera nada se já existirem).
 * accao "listar"   — devolve as duas tabelas.
 * accao "gravar"   — acrescenta uma indemnização e o seu plano de prestações (tudo validado antes de escrever).
 * accao "pagar"    — marca/desmarca uma prestação como paga (dados: {id, paga, data}).
 * accao "anular"   — marca uma indemnização como Anulada (não apaga linhas).
 * Nunca apaga linhas nem toca noutras tabelas. Devolve sempre texto JSON.
 */
const FOLHA = "Indemnizacoes";
const T_IND = "tbl_indemnizacoes";
const T_PRE = "tbl_indemnizacoes_prestacoes";
const COL_IND = ["ID_Indemnizacao", "Empresa", "Codigo", "Nome", "Departamento", "Data Admissao", "Data Referencia",
  "Anos de Trabalho", "Vencimento Base", "Valor Total", "Numero Prestacoes", "Dia Limite", "Primeiro Mes", "Estado",
  "Observacoes", "Criado Por", "Data Criacao", "Utilizador Alteracao", "Data Ultima Alteracao"];
const COL_PRE = ["ID_Prestacao", "ID_Indemnizacao", "Empresa", "Codigo", "Nome", "Numero", "Data Limite", "Valor",
  "Estado", "Data Pagamento", "Pago Por", "Observacoes"];
const DATAS = ["Data Admissao", "Data Referencia", "Primeiro Mes", "Data Criacao", "Data Ultima Alteracao", "Data Limite", "Data Pagamento"];
const MOEDA = ["Vencimento Base", "Valor Total", "Valor"];
const TEXTO = ["ID_Indemnizacao", "ID_Prestacao", "Empresa", "Codigo", "Nome", "Departamento", "Estado", "Observacoes",
  "Criado Por", "Utilizador Alteracao", "Pago Por"];
const LIMITE_PRESTACOES = 120;

type Celula = string | number | boolean;
interface Prestacao { numero: number; dataLimite: string; valor: number; }
interface Pedido {
  empresa?: string; codigo?: string; nome?: string; departamento?: string;
  dataAdmissao?: string; dataReferencia?: string; vencimentoBase?: number;
  diaLimite?: number; primeiroMes?: string; observacoes?: string;
  prestacoes?: Prestacao[];
  id?: string; paga?: boolean; data?: string;
}

function main(workbook: ExcelScript.Workbook, accao: string, dados: string, utilizador: string): string {
  try {
    const a = String(accao || "").trim().toLowerCase();
    let p: Pedido = {};
    if (dados && String(dados).trim()) p = JSON.parse(dados) as Pedido;
    if (a === "preparar") return JSON.stringify(preparar(workbook));
    const ind = workbook.getTable(T_IND), pre = workbook.getTable(T_PRE);
    if (!ind || !pre) return erro("As tabelas das indemnizações ainda não existem. Corra a acção «preparar».");
    if (a === "listar") return JSON.stringify({ok: true, accao: a, indemnizacoes: ler(ind), prestacoes: ler(pre)});
    if (a === "gravar") return gravar(ind, pre, p, utilizador);
    if (a === "pagar") return pagar(pre, p, utilizador);
    if (a === "anular") return anular(ind, p, utilizador);
    return erro("Acção desconhecida: " + a);
  } catch (e) {
    return erro("Erro no script: " + String(e));
  }
}

function erro(m: string): string { return JSON.stringify({ok: false, erro: m}); }

function preparar(workbook: ExcelScript.Workbook): {ok: boolean, accao: string, criadas: string[]} {
  const criadas: string[] = [];
  let folha = workbook.getWorksheet(FOLHA);
  if (!folha) { folha = workbook.addWorksheet(FOLHA); }
  if (!workbook.getTable(T_IND)) {
    const r = folha.getRangeByIndexes(0, 0, 1, COL_IND.length);
    r.setValues([COL_IND]);
    const t = workbook.addTable(r, true); t.setName(T_IND);
    criadas.push(T_IND);
  }
  if (!workbook.getTable(T_PRE)) {
    const inicio = COL_IND.length + 2;
    const r = folha.getRangeByIndexes(0, inicio, 1, COL_PRE.length);
    r.setValues([COL_PRE]);
    const t = workbook.addTable(r, true); t.setName(T_PRE);
    criadas.push(T_PRE);
  }
  formatar(workbook.getTable(T_IND)); formatar(workbook.getTable(T_PRE));
  folha.getUsedRange()?.getFormat().autofitColumns();
  return {ok: true, accao: "preparar", criadas};
}

function formatar(t: ExcelScript.Table) {
  const cab = t.getHeaderRowRange().getValues()[0].map(h => String(h));
  cab.forEach((h, j) => {
    const col = t.getColumn(h).getRangeBetweenHeaderAndTotal();
    if (DATAS.indexOf(h) >= 0) col.setNumberFormat("dd/mm/yyyy");
    else if (MOEDA.indexOf(h) >= 0) col.setNumberFormat("#,##0.00");
    else if (TEXTO.indexOf(h) >= 0) col.setNumberFormat("@");
  });
}

// Uma tabela acabada de criar tem uma linha de dados vazia: é reaproveitada na primeira gravação e ignorada na leitura.
function vazia(v: Celula[]): boolean { return v.every(x => x === "" || x === null); }
function inserir(t: ExcelScript.Table, linhas: Celula[][]) {
  if (!linhas.length) return;
  if (t.getRowCount() === 1 && vazia(t.getRangeBetweenHeaderAndTotal().getValues()[0] as Celula[])) {
    t.getRangeBetweenHeaderAndTotal().getRow(0).setValues([linhas[0]]);
    linhas = linhas.slice(1);
  }
  if (linhas.length) t.addRows(-1, linhas);
}
function dadosTabela(t: ExcelScript.Table): Celula[][] {
  if (t.getRowCount() === 0) return [];
  return (t.getRangeBetweenHeaderAndTotal().getValues() as Celula[][]);
}

function ler(t: ExcelScript.Table): string {
  const cab = t.getHeaderRowRange().getValues()[0].map(h => String(h));
  const n = t.getRowCount();
  const linhas: { [k: string]: Celula }[] = [];
  if (n > 0) {
    const vals = dadosTabela(t).filter(v => !vazia(v));
    vals.forEach(v => { const o: { [k: string]: Celula } = {}; cab.forEach((h, j) => { o[h] = v[j] as Celula; }); linhas.push(o); });
  }
  return JSON.stringify(linhas);
}

// Data ISO (AAAA-MM-DD) -> número de série do Excel.
function serie(iso: string): number {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(iso || "").trim());
  if (!m) return NaN;
  const d = Date.UTC(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  const check = new Date(d);
  if (check.getUTCMonth() !== Number(m[2]) - 1) return NaN;
  return Math.round((d - Date.UTC(1899, 11, 30)) / 86400000);
}
function hojeSerie(): number {
  const d = new Date();
  return Math.round((Date.UTC(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate()) - Date.UTC(1899, 11, 30)) / 86400000);
}
function normCodigo(v: unknown): string {
  const s = String(v ?? "").trim();
  return /^\d+$/.test(s) ? String(Number(s)).padStart(4, "0") : s.toUpperCase();
}
const cent = (x: number) => Math.round(x * 100) / 100;

// Anos completos entre a admissão e a data de referência.
function anosCompletos(adm: string, ref: string): number {
  const a = /^(\d{4})-(\d{2})-(\d{2})$/.exec(adm), r = /^(\d{4})-(\d{2})-(\d{2})$/.exec(ref);
  if (!a || !r) return NaN;
  let anos = Number(r[1]) - Number(a[1]);
  if (Number(r[2]) < Number(a[2]) || (Number(r[2]) === Number(a[2]) && Number(r[3]) < Number(a[3]))) anos -= 1;
  return Math.max(0, anos);
}

function linhaPorCabecalho(t: ExcelScript.Table, o: { [k: string]: Celula }): Celula[] {
  const cab = t.getHeaderRowRange().getValues()[0].map(h => String(h));
  return cab.map(h => (h in o ? o[h] : ""));
}

function gravar(ind: ExcelScript.Table, pre: ExcelScript.Table, p: Pedido, utilizador: string): string {
  const erros: string[] = [];
  const empresa = String(p.empresa || "").trim(), codigo = normCodigo(p.codigo);
  if (!empresa || !codigo) erros.push("Falta a Empresa ou o Código do colaborador.");
  const sAdm = serie(String(p.dataAdmissao || "")), sRef = serie(String(p.dataReferencia || ""));
  if (!isFinite(sAdm)) erros.push("Data de admissão inválida.");
  if (!isFinite(sRef)) erros.push("Data de referência inválida.");
  if (isFinite(sAdm) && isFinite(sRef) && sRef < sAdm) erros.push("A data de referência é anterior à admissão.");
  const base = Number(p.vencimentoBase);
  if (!isFinite(base) || base <= 0) erros.push("Vencimento base inválido.");
  const anos = anosCompletos(String(p.dataAdmissao), String(p.dataReferencia));
  const total = cent(anos * base);
  if (isFinite(total) && total <= 0) erros.push("O colaborador não tem anos completos de trabalho: indemnização = 0.");
  const prest = Array.isArray(p.prestacoes) ? p.prestacoes : [];
  if (!prest.length || prest.length > LIMITE_PRESTACOES) erros.push("O plano deve ter entre 1 e " + LIMITE_PRESTACOES + " prestações.");
  let soma = 0;
  const series: number[] = [];
  prest.forEach((x, i) => {
    const s = serie(String(x.dataLimite || ""));
    if (!isFinite(s)) erros.push("Data limite inválida na prestação " + (i + 1) + ".");
    if (series.length && s < series[series.length - 1]) erros.push("As datas limite devem estar por ordem (prestação " + (i + 1) + ").");
    series.push(s);
    const v = Number(x.valor);
    if (!isFinite(v) || v <= 0) erros.push("Valor inválido na prestação " + (i + 1) + ".");
    soma += v;
  });
  if (prest.length && Math.abs(cent(soma) - total) > 0.01) erros.push("A soma das prestações (" + cent(soma) + ") é diferente do total (" + total + ").");
  // Uma indemnização activa por colaborador.
  const cab = ind.getHeaderRowRange().getValues()[0].map(h => String(h));
  const iE = cab.indexOf("Empresa"), iC = cab.indexOf("Codigo"), iS = cab.indexOf("Estado"), iId = cab.indexOf("ID_Indemnizacao");
  const existentes = dadosTabela(ind).filter(v => !vazia(v));
  if (existentes.some(v => String(v[iE]).trim().toUpperCase() === empresa.toUpperCase() && normCodigo(v[iC]) === codigo && String(v[iS]) !== "Anulada"))
    erros.push("Já existe uma indemnização activa para " + empresa + " / " + codigo + ". Anule-a antes de criar outra.");
  if (erros.length) return JSON.stringify({ok: false, erro: erros.join(" "), erros});

  let maior = 0;
  existentes.forEach(v => { const m = /^IND-(\d+)$/.exec(String(v[iId])); if (m) maior = Math.max(maior, Number(m[1])); });
  const id = "IND-" + String(maior + 1).padStart(5, "0");
  const agora = hojeSerie();
  const dia = Math.min(31, Math.max(1, Math.round(Number(p.diaLimite) || 0))) || "";
  const linhaInd: { [k: string]: Celula } = {
    "ID_Indemnizacao": id, "Empresa": empresa, "Codigo": codigo, "Nome": String(p.nome || ""), "Departamento": String(p.departamento || ""),
    "Data Admissao": sAdm, "Data Referencia": sRef, "Anos de Trabalho": anos, "Vencimento Base": cent(base), "Valor Total": total,
    "Numero Prestacoes": prest.length, "Dia Limite": dia, "Primeiro Mes": series[0], "Estado": "Activa",
    "Observacoes": String(p.observacoes || ""), "Criado Por": String(utilizador || ""), "Data Criacao": agora,
    "Utilizador Alteracao": String(utilizador || ""), "Data Ultima Alteracao": agora
  };
  const linhasPre = prest.map((x, i) => linhaPorCabecalho(pre, {
    "ID_Prestacao": id + "-" + String(i + 1).padStart(3, "0"), "ID_Indemnizacao": id, "Empresa": empresa, "Codigo": codigo,
    "Nome": String(p.nome || ""), "Numero": i + 1, "Data Limite": series[i], "Valor": cent(Number(x.valor)),
    "Estado": "Pendente", "Data Pagamento": "", "Pago Por": "", "Observacoes": ""
  }));
  limparFiltro(ind); limparFiltro(pre);
  inserir(ind, [linhaPorCabecalho(ind, linhaInd)]);
  inserir(pre, linhasPre);
  return JSON.stringify({ok: true, accao: "gravar", id, anos, total, prestacoes: linhasPre.length});
}

function limparFiltro(t: ExcelScript.Table) {
  const f = t.getAutoFilter();
  if (f && f.getIsDataFiltered()) f.clearCriteria();
}

function pagar(pre: ExcelScript.Table, p: Pedido, utilizador: string): string {
  const id = String(p.id || "").trim();
  if (!id) return erro("Falta o identificador da prestação.");
  const cab = pre.getHeaderRowRange().getValues()[0].map(h => String(h));
  const n = pre.getRowCount();
  if (!n) return erro("Prestação não encontrada.");
  const vals = pre.getRangeBetweenHeaderAndTotal().getValues();
  const i = vals.findIndex(v => String(v[cab.indexOf("ID_Prestacao")]) === id);
  if (i < 0) return erro("Prestação não encontrada: " + id);
  const paga = p.paga !== false;
  let sData: number | string = "";
  if (paga) { sData = p.data ? serie(String(p.data)) : hojeSerie(); if (!isFinite(Number(sData))) return erro("Data de pagamento inválida."); }
  const linha = pre.getRangeBetweenHeaderAndTotal().getRow(i);
  const set = (h: string, v: Celula) => { const j = cab.indexOf(h); if (j >= 0) linha.getCell(0, j).setValue(v); };
  set("Estado", paga ? "Paga" : "Pendente");
  set("Data Pagamento", paga ? sData : "");
  set("Pago Por", paga ? String(utilizador || "") : "");
  return JSON.stringify({ok: true, accao: "pagar", id, paga});
}

function anular(ind: ExcelScript.Table, p: Pedido, utilizador: string): string {
  const id = String(p.id || "").trim();
  const cab = ind.getHeaderRowRange().getValues()[0].map(h => String(h));
  if (!id || !ind.getRowCount()) return erro("Indemnização não encontrada.");
  const vals = ind.getRangeBetweenHeaderAndTotal().getValues();
  const i = vals.findIndex(v => String(v[cab.indexOf("ID_Indemnizacao")]) === id);
  if (i < 0) return erro("Indemnização não encontrada: " + id);
  const linha = ind.getRangeBetweenHeaderAndTotal().getRow(i);
  const set = (h: string, v: Celula) => { const j = cab.indexOf(h); if (j >= 0) linha.getCell(0, j).setValue(v); };
  set("Estado", "Anulada");
  set("Utilizador Alteracao", String(utilizador || ""));
  set("Data Ultima Alteracao", hojeSerie());
  if (p.observacoes) set("Observacoes", String(p.observacoes));
  return JSON.stringify({ok: true, accao: "anular", id});
}
