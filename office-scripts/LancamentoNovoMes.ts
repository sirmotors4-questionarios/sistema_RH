/**
 * Office Script "LancamentoNovoMes" — escrita em tbl_lançamento chamada pelo RH_API_Actualizar.
 * Parâmetros: ano (número), mes (nome do mês, ex.: "Março"), linhas (JSON), modo.
 * modo "validar"   — valida um novo mês sem escrever.
 * modo "gravar"    — insere o mês num só bloco; repõe as fórmulas das colunas calculadas.
 * modo "editar"    — altera UMA linha existente (Empresa + Código + Ano + Mês) só nas colunas editáveis;
 *                    recusa se estiver Finalizado; marca Estado do Lançamento = "Alterado".
 * modo "finalizar" — marca todas as linhas do mês como Finalizado / Revisto = Sim.
 * Nunca escreve nas 11 colunas calculadas (PROTEGIDAS). Valida tudo antes de escrever.
 * Escalão: se existir tbl_escaloes_lancamento e a coluna "Escalão" (sem fórmula), o gravar preenche-a nas linhas
 * novas e o editar recalcula-a, com o valor do critério da empresa (por omissão Salario Liquido) já recalculado.
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

// Referências persistentes: nunca guardadas no navegador nem reconstruídas a partir de Alterados.
const TABELA_INICIAIS = "tbl_lancamentos_iniciais_RH";
const CONDICOES = ["RegimePagamento", "Valor Unitário", "Data Inicio Contrato", "Data Fim Contrato", "Dias Trabalhados"];
const AUDITORIA = ["Data Ultima Alteracao", "Utilizador Alteracao"];
const CAMPOS_INICIAIS = EDITAVEIS.filter(c => AUDITORIA.indexOf(c) < 0);
interface ReferenciaInicial { chave: string; valores: Record<string, string | number | boolean>; formulas?: Record<string, string>; }
function tabelaIniciais(workbook: ExcelScript.Workbook): ExcelScript.Table {
  const existente = workbook.getTable(TABELA_INICIAIS);
  if (existente) return existente;
  // addWorksheet sem nome evita colidir com folhas existentes do utilizador.
  const folha = workbook.addWorksheet();
  const cab = folha.getRange("A1:C1");
  cab.setValues([["Chave", "Referencia_JSON", "Guardado_Em"]]);
  const t = folha.addTable(cab, true);
  t.setName(TABELA_INICIAIS);
  return t;
}
function referenciasIniciais(workbook: ExcelScript.Workbook): Map<string, {indice: number; json: string}> {
  const t = workbook.getTable(TABELA_INICIAIS), mapa = new Map<string, {indice: number; json: string}>();
  if (!t || !t.getRowCount()) return mapa;
  const cab = t.getHeaderRowRange().getValues()[0].map(v => String(v));
  if (cab.join("|") !== "Chave|Referencia_JSON|Guardado_Em") throw Error("Estrutura inválida em " + TABELA_INICIAIS);
  t.getRangeBetweenHeaderAndTotal().getValues().forEach((l, i) => {
    const k = String(l[0]);
    if (mapa.has(k)) throw Error("Referência inicial duplicada: " + k);
    mapa.set(k, {indice: i, json: String(l[1])});
  });
  return mapa;
}
function guardarInicial(t: ExcelScript.Table, referencias: Map<string, {indice: number; json: string}>, ref: ReferenciaInicial): void {
  const json = JSON.stringify(ref), existente = referencias.get(ref.chave);
  if (json.length > 32000) throw Error("Referência inicial demasiado longa: " + ref.chave);
  const valores = [[ref.chave, json, new Date().toISOString()]];
  if (existente) t.getRangeBetweenHeaderAndTotal().getRow(existente.indice).setValues(valores);
  else { t.addRows(-1, valores); referencias.set(ref.chave, {indice: t.getRowCount() - 1, json}); }
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

  const objectoLinha = (i: number): Record<string, string | number | boolean> => {
    const valores = corpo.getRow(i).getValues()[0], obj: Record<string, string | number | boolean> = {};
    cab.forEach((h, j) => { obj[h] = valores[j]; });
    return obj;
  };
  const referencia = (i: number, patch: Record<string, unknown> = {}): ReferenciaInicial => {
    const linha = objectoLinha(i), valores: Record<string, string | number | boolean> = {};
    for (const h of CAMPOS_INICIAIS) if (idx(h) >= 0) valores[h] = Object.prototype.hasOwnProperty.call(patch, h) ? valorCelula(h, patch[h]) : linha[h];
    const formulas: Record<string, string> = {}, originais = corpo.getRow(i).getFormulasR1C1()[0];
    for (const h of CAMPOS_INICIAIS) if (idx(h) >= 0 && !Object.prototype.hasOwnProperty.call(patch, h) && String(originais[idx(h)]).startsWith("=")) formulas[h] = String(originais[idx(h)]);
    return {chave: chave(linha.Empresa, linha.Codigo, ano, mes), valores, formulas};
  };
  // Um envelope é transportado pela acção já existente lancamento_editar.
  // Scripts antigos rejeitam-no sem escrever (coluna desconhecida/chave ausente).
  const envelope = pedido[0] ?? {};
  const operacao = String(envelope["__rh_operacao"] ?? "");
  if (r.modo === "editar" && operacao) {
    if (pedido.length !== 1 || ["recalcular_pendentes", "repor_inicial"].indexOf(operacao) < 0) { r.erros.push("operação inválida"); return fim(); }
    if (idx("Estado do Lançamento") < 0) { r.erros.push("coluna Estado do Lançamento em falta"); return fim(); }
    const referencias = referenciasIniciais(workbook);
    const planos: {i: number; patch: Record<string, unknown>; ref?: ReferenciaInicial; formulas?: Record<string, string>}[] = [];
    let ignoradas = 0;
    if (operacao === "recalcular_pendentes") {
      const atualizacoes = envelope["atualizacoes"] as Record<string, unknown>[];
      if (!Array.isArray(atualizacoes) || !atualizacoes.length || atualizacoes.length > LIMITE_LINHAS) { r.erros.push("lista de actualizações inválida"); return fim(); }
      const vistos = new Set<string>();
      for (const patch of atualizacoes) {
        if (!patch || typeof patch !== "object" || Array.isArray(patch)) { r.erros.push("actualização inválida"); continue; }
        const k = chave(patch.Empresa, patch.Codigo, ano, mes), ids = linhasPorChave.get(k) ?? [];
        if (vistos.has(k) || ids.length !== 1) { r.erros.push("lançamento inexistente ou duplicado: " + k); continue; }
        vistos.add(k);
        const i = ids[0];
        // Revalidar o estado no Excel, mesmo que o browser tenha lido Pendente antes.
        if (norm(estadoCol[i]) !== "PENDENTE") { ignoradas++; continue; }
        for (const h of Object.keys(patch)) {
          if (["Empresa", "Codigo"].indexOf(h) >= 0) continue;
          if (CONDICOES.indexOf(h) < 0 && AUDITORIA.indexOf(h) < 0) r.erros.push("campo proibido no recálculo: " + h);
          else if (idx(h) < 0) r.erros.push("coluna desconhecida: " + h);
        }
        const valor = Number(patch["Valor Unitário"]), dias = patch["Dias Trabalhados"];
        if (!isFinite(valor) || valor <= 0 || ["MENSAL", "HORA", "DIÁRIO", "DIARIO"].indexOf(norm(patch.RegimePagamento)) < 0)
          r.erros.push("condições inválidas: " + k);
        if (norm(patch.RegimePagamento) === "MENSAL" && (vazio(dias) || !isFinite(Number(dias)) || Number(dias) < 0 || Number(dias) > 30))
          r.erros.push("dias inválidos: " + k);
        const ref = referencia(i, patch);
        if (JSON.stringify(ref).length > 32000) r.erros.push("referência inicial demasiado longa: " + k);
        planos.push({i, patch, ref});
      }
    } else {
      const k = chave(envelope.Empresa, envelope.Codigo, ano, mes), ids = linhasPorChave.get(k) ?? [];
      if (ids.length !== 1) { r.erros.push("lançamento inexistente ou duplicado: " + k); return fim(); }
      const i = ids[0];
      if (finalizado(i)) { r.erros.push("este lançamento está Finalizado e não pode ser reposto"); return fim(); }
      const guardada = referencias.get(k);
      if (!guardada) { r.erros.push("Não existe referência inicial guardada para este lançamento. Alterações antigas não podem ser reconstruídas."); return fim(); }
      let ref: ReferenciaInicial;
      try { ref = JSON.parse(guardada.json) as ReferenciaInicial; } catch { r.erros.push("referência inicial inválida"); return fim(); }
      if (ref.chave !== k || !ref.valores || typeof ref.valores !== "object" || Array.isArray(ref.valores)) { r.erros.push("referência inicial inválida"); return fim(); }
      const patch: Record<string, unknown> = {};
      for (const h of Object.keys(ref.valores)) {
        if (CAMPOS_INICIAIS.indexOf(h) < 0 || idx(h) < 0) r.erros.push("coluna inválida na referência: " + h);
        else patch[h] = ref.valores[h];
      }
      for (const h of CAMPOS_INICIAIS) if (idx(h) >= 0 && !Object.prototype.hasOwnProperty.call(ref.valores, h)) r.erros.push("referência incompleta: " + h);
      for (const h of AUDITORIA) if (idx(h) >= 0) patch[h] = envelope[h] ?? "";
      const formulas = ref.formulas ?? {};
      for (const h of Object.keys(formulas)) {
        if (CAMPOS_INICIAIS.indexOf(h) < 0 || idx(h) < 0 || typeof formulas[h] !== "string" || !formulas[h].startsWith("=")) r.erros.push("fórmula inválida na referência: " + h);
      }
      planos.push({i, patch, formulas});
    }
    if (r.erros.length) return fim();
    // Guardar a referência antes de qualquer alteração salarial. Recalcular mantém Pendente.
    if (planos.some(p => Boolean(p.ref))) {
      const t = tabelaIniciais(workbook);
      for (const p of planos) if (p.ref) guardarInicial(t, referencias, p.ref);
    }
    for (const p of planos) {
      for (const h of Object.keys(p.patch)) if (CHAVE.indexOf(h) < 0 && !p.formulas?.[h]) definir(p.i, h, p.patch[h]);
      for (const h of Object.keys(p.formulas ?? {})) corpo.getCell(p.i, idx(h)).setFormulaR1C1(p.formulas[h]);
      if (operacao === "repor_inicial") { definir(p.i, "Estado do Lançamento", "Pendente"); definir(p.i, "Revisto", "Não"); }
    }
    workbook.getApplication().calculate(ExcelScript.CalculationType.full);
    for (const p of planos) preencherEscaloes(workbook, corpo, cab, p.i, 1);
    r.alteradas = planos.length;
    r.linha = JSON.stringify({versao: 2, operacao, alteradas: planos.length, ignoradas, linhas: operacao === "repor_inicial" ? planos.map(p => objectoLinha(p.i)) : []});
    r.ok = true;
    return fim();
  }

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
    // Captura apenas Pendentes: não inventa uma origem para lançamentos já Alterados.
    if (norm(estadoCol[i]) === "PENDENTE") {
      const ref = referencia(i);
      if (JSON.stringify(ref).length > 32000) { r.erros.push("referência inicial demasiado longa"); return fim(); }
      guardarInicial(tabelaIniciais(workbook), referenciasIniciais(workbook), ref);
    }
    for (const p of campos) {
      const v = o[p];
      definir(i, p, v);
    }
    definir(i, "Estado do Lançamento", "Alterado");
    workbook.getApplication().calculate(ExcelScript.CalculationType.full);
    preencherEscaloes(workbook, corpo, cab, i, 1);
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
  preencherEscaloes(workbook, corpo, cab, nLinhas, matriz.length);
  r.ok = true;
  return fim();
}

// ---------- Escalões (mesma regra que o site e o script RHIndemnizacoes) ----------
// Um escalão começa no seu Valor_Inicial (inclusive) e vai até ao início do seguinte; o último não tem limite.
// A empresa "Todas" vale para as empresas sem escalões próprios. Sem tabela ou sem coluna, não faz nada.
interface DefEscalao { criterio: string; faixas: { escalao: string; inicio: number }[]; }
function preencherEscaloes(workbook: ExcelScript.Workbook, corpo: ExcelScript.Range, cab: string[], inicio: number, n: number) {
  const j = cab.indexOf("Escalão");
  const tEsc = workbook.getTable("tbl_escaloes_lancamento");
  if (j < 0 || !tEsc || n <= 0 || tEsc.getRowCount() === 0) return;
  const bloco = corpo.getWorksheet().getRangeByIndexes(corpo.getRowIndex() + inicio, corpo.getColumnIndex(), n, cab.length);
  const colEsc = bloco.getColumn(j);
  if (colEsc.getFormulas().some(f => String(f[0]).startsWith("="))) return;
  const ce = tEsc.getHeaderRowRange().getValues()[0].map(v => String(v));
  const iE = ce.indexOf("Empresa"), iC = ce.indexOf("Criterio"), iN = ce.indexOf("Escalao"), iV = ce.indexOf("Valor_Inicial");
  const mapa = new Map<string, DefEscalao>();
  for (const v of tEsc.getRangeBetweenHeaderAndTotal().getValues()) {
    const e = String(v[iE] ?? "").trim().toUpperCase(), nome = String(v[iN] ?? "").trim(), x = Number(v[iV]);
    if (!e || !nome || String(v[iV] ?? "").trim() === "" || !isFinite(x)) continue;
    let d = mapa.get(e);
    if (!d) { d = {criterio: String(v[iC] ?? "").trim() || "Salario Liquido", faixas: []}; mapa.set(e, d); }
    d.faixas.push({escalao: nome, inicio: x});
  }
  if (!mapa.size) return;
  mapa.forEach(d => d.faixas.sort((a, b) => a.inicio - b.inicio));
  const valores = bloco.getValues();
  const iEmp = cab.indexOf("Empresa");
  colEsc.setValues(valores.map(l => {
    const d = mapa.get(String(l[iEmp] ?? "").trim().toUpperCase()) || mapa.get("TODAS");
    if (!d) return [""];
    const k = cab.indexOf(d.criterio), bruto = k >= 0 ? l[k] : "";
    const v = Number(bruto);
    if (bruto === "" || bruto === null || !isFinite(v)) return [""];
    let r = "";
    for (const f of d.faixas) if (v >= f.inicio) r = f.escalao;
    return [r];
  }));
}
