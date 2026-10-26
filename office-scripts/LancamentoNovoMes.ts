/**
 * Office Script — gravação em lote de um novo mês em tbl_lançamento.
 *
 * ESTADO: rascunho para validação numa CÓPIA do ficheiro. Não ligar ao fluxo
 * de produção antes dos testes descritos em office-scripts/README.md.
 *
 * Chamado pelo Power Automate (Excel Online (Business) → "Run script from
 * SharePoint library") com:
 *   ano    número (ex.: 2026)
 *   mes    nome do mês tal como na coluna "Mês" (ex.: "Março")
 *   linhas JSON (texto) com um array de objectos {<cabeçalho>: valor}
 *   modo   "validar" (não escreve nada) | "gravar"
 *
 * Garantias:
 *  - Nunca escreve valores nas 11 colunas calculadas (PROTEGIDAS): repõe nas
 *    linhas novas a fórmula R1C1 mais frequente de cada coluna. Outras colunas
 *    com fórmula (ex.: Dias Trabalhados) aceitam um valor enviado, como no VBA.
 *  - Chave única Empresa + Código + Ano + Mês (normalizada). Rejeita duplicados
 *    no pedido e linhas que já existam na tabela.
 *  - Valida tudo antes de escrever; se houver erros, não escreve nenhuma linha.
 *  - Escrita num único bloco (addRows + setFormulasR1C1), em vez de uma chamada
 *    por linha.
 */

interface Resultado {
  ok: boolean;
  modo: string;
  recebidas: number;
  inseridas: number;
  duplicadasNoPedido: string[];
  jaExistentes: string[];
  erros: string[];
  colunasFormula: string[];
  ms: number;
}

const TABELA = "tbl_lançamento";
const LIMITE_LINHAS = 2000;
// Colunas calculadas que nunca recebem valores (iguais a FORMULAS_LANCAMENTO no site).
const PROTEGIDAS = ["Estado", "Estado de Reforma", "Vencimento Base", "INSS Colaborador", "INSS Empresa", "Total INSS",
  "Salario Liquido", "Total de Acréscimos", "Total de Descontos", "IRPS", "Vencimento"];

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

function main(workbook: ExcelScript.Workbook, ano: number, mes: string, linhas: string, modo: string): Resultado {
  const t0 = Date.now();
  const r: Resultado = {
    ok: false, modo: modo === "gravar" ? "gravar" : "validar", recebidas: 0, inseridas: 0,
    duplicadasNoPedido: [], jaExistentes: [], erros: [], colunasFormula: [], ms: 0
  };
  const fim = () => { r.ms = Date.now() - t0; return r; };

  // 1. Pedido
  let pedido: Record<string, unknown>[];
  try { pedido = JSON.parse(linhas); } catch { r.erros.push("linhas não é JSON válido"); return fim(); }
  if (!Array.isArray(pedido) || pedido.length === 0) { r.erros.push("linhas vazio"); return fim(); }
  if (pedido.length > LIMITE_LINHAS) { r.erros.push(`máximo ${LIMITE_LINHAS} linhas por chamada`); return fim(); }
  if (!Number.isInteger(ano) || ano < 1900 || ano > 9999) { r.erros.push("ano inválido"); return fim(); }
  if (!String(mes ?? "").trim()) { r.erros.push("mês vazio"); return fim(); }
  r.recebidas = pedido.length;

  // 2. Estrutura da tabela
  const tabela = workbook.getTable(TABELA);
  if (!tabela) { r.erros.push(`tabela ${TABELA} não encontrada`); return fim(); }
  const cab = tabela.getHeaderRowRange().getValues()[0].map(String);
  const idx = (n: string) => cab.indexOf(n);
  for (const c of ["Empresa", "Codigo", "Ano", "Mês"]) if (idx(c) < 0) r.erros.push(`coluna ${c} em falta`);
  if (r.erros.length) return fim();

  const corpo = tabela.getRangeBetweenHeaderAndTotal();
  const nLinhas = tabela.getRowCount();
  if (nLinhas === 0) { r.erros.push("tabela sem linhas: não é possível obter as fórmulas de referência"); return fim(); }
  // Fórmula de referência por coluna = a fórmula R1C1 mais frequente nessa coluna
  // (algumas linhas antigas têm valores no lugar da fórmula, ex.: Dias Trabalhados).
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

  // 3. Chaves existentes (lê só as 4 colunas da chave)
  const existentes = new Set<string>();
  if (nLinhas > 0) {
    const col = (n: string) => corpo.getColumn(idx(n)).getValues().map(v => v[0]);
    const [e, c, a, m] = ["Empresa", "Codigo", "Ano", "Mês"].map(col);
    for (let i = 0; i < nLinhas; i++) existentes.add(chave(e[i], c[i], a[i], m[i]));
  }

  // 4. Validação do pedido (nada é escrito se houver erros)
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
    if (existentes.has(k)) { r.jaExistentes.push(k); return; }
    for (const p of Object.keys(o)) {
      if (idx(p) < 0) desconhecidas.add(p);
      else if (PROTEGIDAS.includes(p) && o[p] !== null && o[p] !== undefined && o[p] !== "")
        r.erros.push(`linha ${n + 1}: tentativa de escrever na coluna calculada "${p}"`);
    }
    const vazio = (v: unknown) => v === null || v === undefined || v === "";
    // Coluna com fórmula: mantém a fórmula, salvo se o pedido trouxer um valor e a coluna não for protegida.
    const usa = cab.map((h, i) => ehFormula[i] && (PROTEGIDAS.includes(h) || vazio(o[h])));
    comFormula.push(usa);
    matriz.push(cab.map((h, i) => {
      if (usa[i]) return "";                     // reposto abaixo com a fórmula
      if (h === "Ano") return ano;
      if (h === "Mês") return mes;
      const v = o[h];
      return vazio(v) ? "" : (v as string | number | boolean);
    }));
  });
  if (desconhecidas.size) r.erros.push(`colunas desconhecidas: ${[...desconhecidas].join(", ")}`);
  if (r.duplicadasNoPedido.length) r.erros.push(`${r.duplicadasNoPedido.length} duplicado(s) no pedido`);
  if (r.jaExistentes.length) r.erros.push(`${r.jaExistentes.length} chave(s) já existem na tabela`);
  if (r.erros.length) return fim();
  if (r.modo === "validar") { r.ok = true; return fim(); }

  // 5. Escrita em bloco + reposição das fórmulas
  const primeira = nLinhas;
  tabela.addRows(undefined, matriz);
  const novo = corpo.getWorksheet().getRangeByIndexes(
    corpo.getRowIndex() + primeira, corpo.getColumnIndex(), matriz.length, cab.length);
  // Só as colunas com fórmula são reescritas, coluna a coluna; as restantes mantêm os valores inseridos.
  cab.forEach((_, j) => {
    if (!ehFormula[j]) return;
    const actual = novo.getColumn(j).getFormulasR1C1();
    novo.getColumn(j).setFormulasR1C1(actual.map((l, i) => [comFormula[i][j] ? formulasRef[j] : l[0]]));
  });
  r.inseridas = matriz.length;
  r.ok = true;
  return fim();
}
