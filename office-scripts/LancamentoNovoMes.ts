/**
 * Office Script "LancamentoNovoMes" — escrita em tbl_lançamento chamada pelo RH_API_Actualizar.
 * Parâmetros: ano (número), mes (nome do mês, ex.: "Março"), linhas (JSON), modo.
 * modo "validar"   — valida um novo mês sem escrever.
 * modo "gravar"    — insere o mês num só bloco; repõe as fórmulas das colunas calculadas.
 * modo "editar"    — altera UMA linha existente (Empresa + Código + Ano + Mês) só nas colunas editáveis;
 *                    recusa se estiver Finalizado; marca Estado do Lançamento = "Alterado".
 * modo "finalizar" — marca todas as linhas do mês como Finalizado / Revisto = Sim.
 * Nunca escreve nas 11 colunas calculadas (PROTEGIDAS). Valida tudo antes de escrever.
 */
interface Resultado {
  ok: boolean;
  modo: string;
  recebidas: number;
  inseridas: number;
  alteradas: number;
  duplicadasNoPedido: string[];
  jaExistentes: string[];
  erros: string[];
  colunasFormula: string[];
  ms: number;
  filtrosLimpos?: boolean;
  linha?: string; // JSON da linha recalculada (o Excel não aceita objectos genéricos no retorno)
}
const TABELA = "tbl_lançamento";
const LIMITE_LINHAS = 2000;
const PROTEGIDAS = ["Estado", "Estado de Reforma", "Vencimento Base", "INSS Colaborador", "INSS Empresa", "Total INSS",
  "Salario Liquido", "Total de Acréscimos", "Total de Descontos", "IRPS", "Vencimento"];
const EDITAVEIS = ["Canalizar", "TotalDependentes", "Horas Trabalhadas", "Faltas Justificadas", "Faltas Injustificadas",
  "Horas Extras", "Subsídios", "Bónus", "Outros Acréscimos", "Adiantamentos", "Empréstimos", "Outros Descontos", "Observações",
  "RegimePagamento", "Valor Unitário", "Data Inicio Contrato", "Data Fim Contrato", "Dias Trabalhados",
  "Data Ultima Alteracao", "Utilizador Alteracao"];
const CHAVE = ["Empresa", "Codigo", "Ano", "Mês"];
const TEXTO = ["Codigo", "Nome", "Empresa", "Departamento", "Profissao", "Mês", "Observações", "Canalizar", "Estado do Lançamento",
  "Origem_Lancamento", "Revisto", "Utilizador Alteracao", "RegimePagamento", "ID_Lancamento"];

function norm(v: unknown): string {
  return String(v ?? "").trim().toUpperCase();
}
function normCodigo(v: unknown): string {
  const s = String(v ?? "").trim();
  return /^\d+$/.test(s) ? String(Number(s)).padStart(4, "0") : s.toUpperCase();
}
function chave(empresa: unknown, codigo: unknown, ano: unknown, mes: unknown): string {
  return [norm(empresa), normCodigo(codigo), String(Number(ano)), norm(mes)].join("|");
}
function vazio(v: unknown): boolean {
  return v === null || v === undefined || v === "";
}
// Colunas de texto ficam texto (Código "0012" nunca vira 12); nas restantes um número em texto passa a número.
function valorCelula(coluna: string, v: unknown): string | number | boolean {
  if (vazio(v)) return "";
  if (coluna === "Codigo") return normCodigo(v);
  if (TEXTO.indexOf(coluna) >= 0) return String(v);
  if (typeof v === "string" && /^-?\d+(\.\d+)?$/.test(v.trim())) return Number(v.trim());
  return v as string | number | boolean;
}

function main(workbook: ExcelScript.Workbook, ano: number, mes: string, linhas: string, modo: string): Resultado {
  const t0 = Date.now();
  const modos = ["validar", "gravar", "editar", "finalizar"];
  const r: Resultado = {
    ok: false, modo: modos.indexOf(modo) >= 0 ? modo : "validar", recebidas: 0, inseridas: 0, alteradas: 0,
    duplicadasNoPedido: [], jaExistentes: [], erros: [], colunasFormula: [], ms: 0
  };
  const fim = () => { r.ms = Date.now() - t0; return r; };

  let pedido: Record<string, unknown>[] = [];
  try { pedido = JSON.parse(linhas || "[]"); } catch { r.erros.push("linhas não é JSON válido"); return fim(); }
  if (!Array.isArray(pedido)) { r.erros.push("linhas tem de ser uma lista"); return fim(); }
  if (r.modo !== "finalizar" && pedido.length === 0) { r.erros.push("linhas vazio"); return fim(); }
  if (pedido.length > LIMITE_LINHAS) { r.erros.push(`máximo ${LIMITE_LINHAS} linhas por chamada`); return fim(); }
  if (!Number.isInteger(ano) || ano < 1900 || ano > 9999) { r.erros.push("ano inválido"); return fim(); }
  if (!String(mes ?? "").trim()) { r.erros.push("mês vazio"); return fim(); }
  r.recebidas = pedido.length;

  const tabela = workbook.getTable(TABELA);
  if (!tabela) { r.erros.push(`tabela ${TABELA} não encontrada`); return fim(); }
  const cab = tabela.getHeaderRowRange().getValues()[0].map(v => String(v));
  const idx = (n: string) => cab.indexOf(n);
  for (const c of CHAVE) if (idx(c) < 0) r.erros.push(`coluna ${c} em falta`);
  if (r.erros.length) return fim();

  const corpo = tabela.getRangeBetweenHeaderAndTotal();
  const nLinhas = tabela.getRowCount();
  const col = (n: string) => corpo.getColumn(idx(n)).getValues().map(v => v[0]);
  const [e, c, a, m] = CHAVE.map(n => col(n));
  const linhasPorChave = new Map<string, number[]>();
  for (let i = 0; i < nLinhas; i++) {
    const k = chave(e[i], c[i], a[i], m[i]);
    const lista = linhasPorChave.get(k) ?? [];
    lista.push(i);
    linhasPorChave.set(k, lista);
  }
  const estadoCol = idx("Estado do Lançamento") >= 0 ? col("Estado do Lançamento") : [];
  const finalizado = (i: number) => norm(estadoCol[i]) === "FINALIZADO";
  const definir = (i: number, nome: string, v: unknown) => {
    if (idx(nome) >= 0) corpo.getCell(i, idx(nome)).setValue(valorCelula(nome, v));
  };

  if (r.modo === "finalizar") {
    const alvo = norm(mes);
    const meta = pedido[0] ?? {};
    const linhasMes: number[] = [];
    for (let i = 0; i < nLinhas; i++) if (Number(a[i]) === ano && norm(m[i]) === alvo) linhasMes.push(i);
    if (!linhasMes.length) { r.erros.push("não há lançamentos para este mês"); return fim(); }
    for (const i of linhasMes) {
      if (finalizado(i)) continue;
      definir(i, "Estado do Lançamento", "Finalizado");
      definir(i, "Revisto", "Sim");
      if (!vazio(meta["Data Ultima Alteracao"])) definir(i, "Data Ultima Alteracao", meta["Data Ultima Alteracao"] as number);
      if (!vazio(meta["Utilizador Alteracao"])) definir(i, "Utilizador Alteracao", String(meta["Utilizador Alteracao"]));
      r.alteradas++;
    }
    r.ok = true;
    return fim();
  }

  if (r.modo === "editar") {
    if (pedido.length !== 1) { r.erros.push("editar recebe exactamente uma linha"); return fim(); }
    const o = pedido[0];
    if (!o || typeof o !== "object") { r.erros.push("linha inválida"); return fim(); }
    const k = chave(o["Empresa"], o["Codigo"], ano, mes);
    const encontradas = linhasPorChave.get(k) ?? [];
    if (encontradas.length === 0) { r.erros.push(`lançamento ${k} não existe`); return fim(); }
    if (encontradas.length > 1) { r.erros.push(`lançamento ${k} está repetido na tabela`); return fim(); }
    const i = encontradas[0];
    if (finalizado(i)) { r.erros.push("este lançamento está Finalizado e não pode ser alterado"); return fim(); }
    const campos = Object.keys(o).filter(p => CHAVE.indexOf(p) < 0);
    for (const p of campos) {
      if (EDITAVEIS.indexOf(p) < 0) r.erros.push(`coluna não editável: ${p}`);
      else if (idx(p) < 0) r.erros.push(`coluna desconhecida: ${p}`);
    }
    if (r.erros.length) return fim();
    for (const p of campos) {
      const v = o[p];
      definir(i, p, v);
    }
    definir(i, "Estado do Lançamento", "Alterado");
    // Devolve a linha já recalculada (Vencimento, INSS, IRPS, Líquido…) para o site actualizar só esta linha.
    const valores = corpo.getRow(i).getValues()[0];
    const obj: { [k: string]: string | number | boolean } = {};
    cab.forEach((h, j) => { obj[h] = valores[j]; });
    r.linha = JSON.stringify(obj);
    r.alteradas = 1;
    r.ok = true;
    return fim();
  }

  // Só o novo mês precisa das fórmulas de referência: ler as fórmulas de toda a tabela é o passo mais
  // pesado do script, por isso editar e finalizar não o fazem.
  if (nLinhas === 0) { r.erros.push("tabela sem linhas: não é possível obter as fórmulas de referência"); return fim(); }
  const todas = corpo.getFormulasR1C1();
  const formulasRef: string[] = cab.map((_, j) => {
    const conta = new Map<string, number>();
    for (const l of todas) { const f = String(l[j]); if (f.startsWith("=")) conta.set(f, (conta.get(f) ?? 0) + 1); }
    let melhor = "", n = 0;
    conta.forEach((v, f) => { if (v > n) { melhor = f; n = v; } });
    return melhor;
  });
  const ehFormula = formulasRef.map(f => f !== "");
  r.colunasFormula = cab.filter((_, i) => ehFormula[i]);
  for (const c of PROTEGIDAS) {
    if (idx(c) < 0) r.erros.push(`coluna calculada ${c} em falta`);
    else if (!ehFormula[idx(c)]) r.erros.push(`coluna ${c} devia ter fórmula e não tem`);
  }
  if (r.erros.length) return fim();

  const vistas = new Set<string>();
  const desconhecidas = new Set<string>();
  const matriz: (string | number | boolean)[][] = [];
  const comFormula: boolean[][] = [];
  pedido.forEach((o, n) => {
    if (!o || typeof o !== "object") { r.erros.push(`linha ${n + 1}: não é objecto`); return; }
    if (!norm(o["Empresa"]) || !normCodigo(o["Codigo"])) { r.erros.push(`linha ${n + 1}: Empresa/Codigo em falta`); return; }
    const k = chave(o["Empresa"], o["Codigo"], ano, mes);
    if (vistas.has(k)) { r.duplicadasNoPedido.push(k); return; }
    vistas.add(k);
    if (linhasPorChave.has(k)) { r.jaExistentes.push(k); return; }
    for (const p of Object.keys(o)) {
      if (idx(p) < 0) desconhecidas.add(p);
      else if (PROTEGIDAS.indexOf(p) >= 0 && !vazio(o[p])) r.erros.push(`linha ${n + 1}: tentativa de escrever na coluna calculada "${p}"`);
    }
    const usa = cab.map((h, i) => ehFormula[i] && (PROTEGIDAS.indexOf(h) >= 0 || vazio(o[h])));
    comFormula.push(usa);
    matriz.push(cab.map((h, i) => {
      if (usa[i]) return "";
      if (h === "Ano") return ano;
      if (h === "Mês") return mes;
      return valorCelula(h, o[h]);
    }));
  });
  if (desconhecidas.size) r.erros.push(`colunas desconhecidas: ${Array.from(desconhecidas).join(", ")}`);
  if (r.duplicadasNoPedido.length) r.erros.push(`${r.duplicadasNoPedido.length} duplicado(s) no pedido`);
  if (r.jaExistentes.length) r.erros.push(`${r.jaExistentes.length} chave(s) já existem na tabela`);
  if (r.erros.length) return fim();
  if (r.modo === "validar") { r.ok = true; return fim(); }

  // Excel recusa addRows numa tabela filtrada (InsertDeleteConflict): limpa o filtro (só a vista, não os dados).
  const filtro = tabela.getAutoFilter();
  if (filtro && filtro.getIsDataFiltered()) { filtro.clearCriteria(); r.filtrosLimpos = true; }
  tabela.addRows(undefined, matriz);
  const novo = corpo.getWorksheet().getRangeByIndexes(
    corpo.getRowIndex() + nLinhas, corpo.getColumnIndex(), matriz.length, cab.length);
  const jc = idx("Codigo");
  novo.getColumn(jc).setNumberFormat("@");
  novo.getColumn(jc).setValues(matriz.map(l => [String(l[jc])]));
  cab.forEach((_, j) => {
    if (!ehFormula[j]) return;
    const actual = novo.getColumn(j).getFormulasR1C1();
    novo.getColumn(j).setFormulasR1C1(actual.map((l, i) => [comFormula[i][j] ? formulasRef[j] : l[0]]));
  });
  r.inseridas = matriz.length;
  r.ok = true;
  return fim();
}
