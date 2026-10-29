/**
 * Office Script "RHIndemnizacoes" — chamado pelo fluxo RH_API_Indemnizacoes.
 * Parâmetros: accao, dados (texto JSON), utilizador (e-mail de quem pediu).
 * accao "preparar" — cria a folha "Indemnizacoes" e as tabelas tbl_indemnizacoes e tbl_indemnizacoes_prestacoes
 *                    se não existirem (não altera nada se já existirem).
 * accao "listar"   — devolve as duas tabelas.
 * accao "gravar"   — acrescenta uma indemnização e o seu plano de prestações (tudo validado antes de escrever).
 * accao "pagar"    — marca/desmarca uma prestação como paga (dados: {id, paga, data}).
 * accao "anular"   — marca uma indemnização como Anulada (não apaga linhas).
 * Cadastro e adendas (mesmo fluxo):
 * accao "chaves"   — preenche Chave_RH (EMPRESA|CÓDIGO) na tbl_funcionarios onde estiver vazia (necessária para editar fichas).
 * accao "registar" — regista uma adenda (uma linha por campo em tbl_adendas) e, se a vigência já começou e «aplicar»
 *                    for verdadeiro, aplica os novos valores à tbl_funcionarios (nunca a células com fórmula).
 * accao "aplicar"  — aplica ao cadastro adendas pendentes (dados: {ids}).
 * accao "revogar"  — marca adendas como Anuladas (dados: {ids, motivo}); o cadastro não é revertido.
 * Escalões salariais (tbl_escaloes_lancamento, folha "Escaloes"):
 * accao "escaloes" — dados.op "ler": devolve as definições; "gravar": substitui os escalões de UMA empresa
 *                    (cria a tabela se não existir); "aplicar": preenche a coluna Escalão da tbl_lançamento
 *                    num mês (dados: {ano, mesNome}), só nas linhas não finalizadas.
 * Acessos (tbl_utilizadores_RH):
 * accao "trocar" ou dados.modulo "trocar" — troca o código do próprio utilizador (o fluxo só deixa passar
 *                    os perfis com permissão de gravação):
 *                    dados {chaveActual, chaveNova} ou {op:"estado"} para saber se tem de trocar.
 * dados.modulo "acessos" — só Superadministrador: op "listar" (sem as chaves), "gravar" (criar/alterar
 *                    utilizador) e "codigo" (definir código temporário: Trocar_Codigo = Sim).
 * Só apaga linhas da tbl_escaloes_lancamento (ao substituir os escalões da empresa gravada). Devolve um objecto simples (sem índices genéricos).
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
interface Resultado {
  ok: boolean; accao: string; erro?: string; erros?: string[]; criadas?: string[];
  id?: string; anos?: number; total?: number; prestacoes?: number; paga?: boolean;
  listaIndemnizacoes?: string; listaPrestacoes?: string; // JSON das tabelas (accao "listar")
  ids?: string[]; numero?: number; aplicada?: boolean; preenchidas?: number; semCodigo?: number;
  listaEscaloes?: string; alteradas?: number; semEscalao?: number;
  listaUtilizadores?: string; trocar?: boolean;
}
interface Alteracao { campo: string; novo: string; }
interface Pedido {
  empresa?: string; codigo?: string; nome?: string; departamento?: string;
  dataAdmissao?: string; dataReferencia?: string; vencimentoBase?: number;
  diaLimite?: number; primeiroMes?: string; observacoes?: string;
  prestacoes?: Prestacao[];
  id?: string; paga?: boolean; data?: string;
  // adendas
  motivo?: string; dataAdenda?: string; inicio?: string; fim?: string; documento?: string;
  alteracoes?: Alteracao[]; aplicar?: boolean; ids?: string[];
  // escalões
  // acessos
  email?: string; perfil?: string; estado?: string; chave?: string; chaveActual?: string; chaveNova?: string; idUtilizador?: string;
  modulo?: string; op?: string; criterio?: string; faixas?: Faixa[]; ano?: number; mesNome?: string;
}
interface Faixa { escalao: string; inicio: number; }

function main(workbook: ExcelScript.Workbook, accao: string, dados: string, utilizador: string): Resultado {
  try {
    const a = String(accao || "").trim().toLowerCase();
    let p: Pedido = {};
    if (dados && String(dados).trim()) p = JSON.parse(dados) as Pedido;
    // "trocar" (acção acesso_trocar) é a única aceite pelo fluxo para qualquer perfil: só troca o código do próprio
    // utilizador e ignora tudo o resto do pedido (nunca chega aos outros módulos).
    if (a === "trocar" || p.modulo === "trocar") return trocarCodigo(workbook, p, utilizador);
    if (a === "preparar") return preparar(workbook);
    if (a === "chaves") return preencherChaves(workbook);
    if (a === "registar") return registarAdenda(workbook, p, utilizador);
    if (a === "aplicar") return aplicarAdendas(workbook, p, utilizador);
    if (a === "revogar") return revogarAdendas(workbook, p, utilizador);
    // Os escalões também chegam por qualquer acção já aceite pelo fluxo, com dados.modulo = "escaloes".
    if (a === "escaloes" || p.modulo === "escaloes") return escaloes(workbook, p, utilizador);
    if (p.modulo === "acessos") return acessos(workbook, p, utilizador);
    const ind = workbook.getTable(T_IND), pre = workbook.getTable(T_PRE);
    if (!ind || !pre) return erro("As tabelas das indemnizações ainda não existem. Corra a acção «preparar».");
    if (a === "listar") return {ok: true, accao: a, listaIndemnizacoes: ler(ind), listaPrestacoes: ler(pre)};
    if (a === "gravar") return gravar(ind, pre, p, utilizador);
    if (a === "pagar") return pagar(pre, p, utilizador);
    if (a === "anular") return anular(ind, p, utilizador);
    return erro("Acção desconhecida: " + a);
  } catch (e) {
    return erro("Erro no script: " + String(e));
  }
}

function erro(m: string): Resultado { return {ok: false, accao: "erro", erro: m, erros: [m]}; }

function preparar(workbook: ExcelScript.Workbook): Resultado {
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

function gravar(ind: ExcelScript.Table, pre: ExcelScript.Table, p: Pedido, utilizador: string): Resultado {
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
  if (erros.length) return {ok: false, accao: "gravar", erro: erros.join(" "), erros};

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
  return {ok: true, accao: "gravar", id, anos, total, prestacoes: linhasPre.length};
}

function limparFiltro(t: ExcelScript.Table) {
  const f = t.getAutoFilter();
  if (f && f.getIsDataFiltered()) f.clearCriteria();
}

function pagar(pre: ExcelScript.Table, p: Pedido, utilizador: string): Resultado {
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
  return {ok: true, accao: "pagar", id, paga};
}

function anular(ind: ExcelScript.Table, p: Pedido, utilizador: string): Resultado {
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
  return {ok: true, accao: "anular", id};
}


// ======================================================================== cadastro e adendas
const T_FUNC = "tbl_funcionarios";
const T_AD = "tbl_adendas";
// Campos que uma adenda pode alterar (nomes iguais aos cabeçalhos da tbl_funcionarios). Empresa e Código não:
// mudar de empresa muda a chave do colaborador.
const CAMPOS_ADENDA = ["Vencimento Base", "RegimePagamento", "TipoContrato", "Departamento", "Profissao", "Categoria", "SituacaoProfissional"];
const REGIMES = ["Mensal", "Diário", "Hora"];

function agoraSerie(): number {
  const d = new Date();
  return (d.getTime() - d.getTimezoneOffset() * 60000) / 86400000 + 25569;
}
function textoCelula(v: Celula): string { return String(v ?? "").trim(); }
function cabecalho(t: ExcelScript.Table): string[] { return t.getHeaderRowRange().getValues()[0].map(h => String(h)); }

// Linha (índice no corpo) de um colaborador pela Empresa + Código.
function linhaColaborador(func: ExcelScript.Table, empresa: string, codigo: string): number {
  const cab = cabecalho(func), iE = cab.indexOf("Empresa"), iC = cab.indexOf("Codigo");
  const vals = dadosTabela(func);
  return vals.findIndex(v => textoCelula(v[iE]).toUpperCase() === empresa.toUpperCase() && normCodigo(v[iC]) === codigo);
}

function preencherChaves(workbook: ExcelScript.Workbook): Resultado {
  const func = workbook.getTable(T_FUNC);
  if (!func) return erro("Tabela tbl_funcionarios não encontrada.");
  const cab = cabecalho(func), iK = cab.indexOf("Chave_RH"), iE = cab.indexOf("Empresa"), iC = cab.indexOf("Codigo");
  if (iK < 0) return erro("A tbl_funcionarios não tem a coluna Chave_RH.");
  const col = func.getColumnByName("Chave_RH").getRangeBetweenHeaderAndTotal();
  const formulas = col.getFormulas();
  if (formulas.some(f => String(f[0]).startsWith("="))) return erro("A coluna Chave_RH tem fórmulas; não foi alterada.");
  const vals = dadosTabela(func), novos: Celula[][] = [];
  let preenchidas = 0, semCodigo = 0;
  vals.forEach(v => {
    const actual = textoCelula(v[iK]), emp = textoCelula(v[iE]), cod = textoCelula(v[iC]);
    if (actual) { novos.push([v[iK]]); return; }
    if (!emp || !cod) { if (emp || cod || textoCelula(v[cab.indexOf("Nome")])) semCodigo++; novos.push([""]); return; }
    novos.push([emp.toUpperCase() + "|" + normCodigo(cod)]); preenchidas++;
  });
  // chaves repetidas impediriam a edição: verifica antes de escrever
  const vistas = new Map<string, number>();
  novos.forEach(n => { const k = textoCelula(n[0]).toUpperCase(); if (k) vistas.set(k, (vistas.get(k) || 0) + 1); });
  const repetidas = Array.from(vistas.entries()).filter(e => e[1] > 1).map(e => e[0]);
  if (repetidas.length) return {ok: false, accao: "chaves", erro: "Há colaboradores repetidos (mesma Empresa + Código): " + repetidas.slice(0, 10).join(", ") + ". Nada foi alterado.", erros: repetidas};
  if (preenchidas) col.setValues(novos);
  return {ok: true, accao: "chaves", preenchidas, semCodigo};
}

function validarValor(campo: string, novo: string): string {
  if (!novo) return "Indique o novo valor de " + campo + ".";
  if (campo === "Vencimento Base") { const n = Number(novo); if (!isFinite(n) || n <= 0) return "Vencimento base inválido."; }
  if (campo === "RegimePagamento" && REGIMES.indexOf(novo) < 0) return "Regime de pagamento inválido (Mensal, Diário ou Hora).";
  if (novo.length > 120) return campo + ": máximo 120 caracteres.";
  return "";
}

// Escreve no cadastro os valores de uma adenda. Devolve erro (texto) ou "".
function aplicarNoCadastro(func: ExcelScript.Table, iLinha: number, alteracoes: Alteracao[], utilizador: string): string {
  const cab = cabecalho(func), linha = func.getRangeBetweenHeaderAndTotal().getRow(iLinha);
  const formulas = linha.getFormulas()[0];
  for (const a of alteracoes) {
    const j = cab.indexOf(a.campo);
    if (j < 0) return "A tbl_funcionarios não tem a coluna " + a.campo + ".";
    if (String(formulas[j]).startsWith("=")) return "A coluna " + a.campo + " tem fórmula no cadastro; não foi alterada.";
  }
  for (const a of alteracoes) {
    const j = cab.indexOf(a.campo);
    linha.getCell(0, j).setValue(a.campo === "Vencimento Base" ? Number(a.novo) : a.novo);
  }
  const jD = cab.indexOf("Data Ultima Alteracao"), jU = cab.indexOf("Utilizador Alteracao"), jK = cab.indexOf("Chave_RH");
  if (jD >= 0 && !String(formulas[jD]).startsWith("=")) linha.getCell(0, jD).setValue(agoraSerie());
  if (jU >= 0 && !String(formulas[jU]).startsWith("=")) linha.getCell(0, jU).setValue(utilizador);
  if (jK >= 0 && !textoCelula(linha.getValues()[0][jK] as Celula)) {
    const v = linha.getValues()[0] as Celula[];
    linha.getCell(0, jK).setValue(textoCelula(v[cab.indexOf("Empresa")]).toUpperCase() + "|" + normCodigo(v[cab.indexOf("Codigo")]));
  }
  return "";
}

function registarAdenda(workbook: ExcelScript.Workbook, p: Pedido, utilizador: string): Resultado {
  const func = workbook.getTable(T_FUNC), ad = workbook.getTable(T_AD);
  if (!func || !ad) return erro("Faltam as tabelas tbl_funcionarios ou tbl_adendas.");
  const erros: string[] = [];
  const empresa = String(p.empresa || "").trim(), codigo = normCodigo(p.codigo);
  const iLinha = linhaColaborador(func, empresa, codigo);
  if (iLinha < 0) return erro("Colaborador não encontrado: " + empresa + " / " + codigo + ".");
  const motivo = String(p.motivo || "").trim();
  if (!motivo) erros.push("Indique o motivo da adenda.");
  const sAd = serie(String(p.dataAdenda || "")), sIni = serie(String(p.inicio || "")), sFim = p.fim ? serie(String(p.fim)) : NaN;
  if (!isFinite(sAd)) erros.push("Data da adenda inválida.");
  if (!isFinite(sIni)) erros.push("Data de início de vigência inválida.");
  if (p.fim && (!isFinite(sFim) || sFim < sIni)) erros.push("A data de fim de vigência é inválida ou anterior ao início.");
  const alts = Array.isArray(p.alteracoes) ? p.alteracoes.map(a => ({campo: String(a.campo || "").trim(), novo: String(a.novo ?? "").trim()})) : [];
  if (!alts.length) erros.push("Indique pelo menos uma alteração.");
  const vistos: string[] = [];
  const cabF = cabecalho(func), linhaF = func.getRangeBetweenHeaderAndTotal().getRow(iLinha).getValues()[0] as Celula[];
  alts.forEach(a => {
    if (CAMPOS_ADENDA.indexOf(a.campo) < 0) erros.push("Campo não permitido numa adenda: " + a.campo + ".");
    if (vistos.indexOf(a.campo) >= 0) erros.push("Campo repetido: " + a.campo + ".");
    vistos.push(a.campo);
    const e = validarValor(a.campo, a.novo); if (e) erros.push(e);
    const j = cabF.indexOf(a.campo);
    if (j >= 0 && textoCelula(linhaF[j]) === a.novo) erros.push(a.campo + ": o novo valor é igual ao actual.");
  });
  if (erros.length) return {ok: false, accao: "registar", erro: erros.join(" "), erros};

  const cabA = cabecalho(ad), existentes = dadosTabela(ad).filter(v => !vazia(v));
  const iId = cabA.indexOf("ID_Adenda"), iNum = cabA.indexOf("Numero_Adenda"), iCod = cabA.indexOf("Codigo"), iEmp = cabA.indexOf("Empresa");
  const ano = new Date().getFullYear();
  let maiorId = 0, maiorNum = 0;
  existentes.forEach(v => {
    const m = new RegExp("^AD-" + ano + "-(\\d+)$").exec(textoCelula(v[iId])); if (m) maiorId = Math.max(maiorId, Number(m[1]));
    if (normCodigo(v[iCod]) === codigo && textoCelula(v[iEmp]).toUpperCase() === empresa.toUpperCase()) maiorNum = Math.max(maiorNum, Number(v[iNum]) || 0);
  });
  const numero = maiorNum + 1, hoje = hojeSerie(), agora = agoraSerie();
  const aplicarJa = p.aplicar !== false && sIni <= hoje;
  const ids: string[] = [];
  const nome = textoCelula(linhaF[cabF.indexOf("Nome")]), dep = textoCelula(linhaF[cabF.indexOf("Departamento")]);
  const linhas = alts.map((a, k) => {
    const id = "AD-" + ano + "-" + String(maiorId + 1 + k).padStart(4, "0"); ids.push(id);
    const j = cabF.indexOf(a.campo), anterior = j >= 0 ? linhaF[j] : "";
    const dif = a.campo === "Vencimento Base" && isFinite(Number(anterior)) ? cent(Number(a.novo) - Number(anterior)) : "";
    const o: { [k: string]: Celula } = {
      "ID_Adenda": id, "Codigo": codigo, "Nome": nome, "Empresa": empresa, "Departamento": dep, "Numero_Adenda": numero,
      "Data_Adenda": sAd, "Data_Inicio_Vigencia": sIni, "Data_Fim_Vigencia": isFinite(sFim) ? sFim : "", "Motivo": motivo,
      "Campo_Alterado": a.campo, "Valor_Anterior": anterior, "Valor_Novo": a.campo === "Vencimento Base" ? Number(a.novo) : a.novo, "Diferenca": dif,
      "Observacoes": String(p.observacoes || ""), "Observacoes_Adicionais": "", "Documento_Referencia": String(p.documento || ""),
      "Estado_Adenda": "Activa", "Data_Aplicacao": aplicarJa ? hoje : "", "Aplicada": aplicarJa ? "Sim" : "Não",
      "Origem_Alteracao": "Sistema RH (web)", "Data_Criacao": agora, "Utilizador_Criacao": utilizador,
      "Data_Ultima_Alteracao": agora, "Utilizador_Alteracao": utilizador
    };
    return linhaPorCabecalho(ad, o);
  });
  if (aplicarJa) { const e = aplicarNoCadastro(func, iLinha, alts, utilizador); if (e) return erro(e + " A adenda não foi registada."); }
  limparFiltro(ad);
  inserir(ad, linhas);
  return {ok: true, accao: "registar", ids, numero, aplicada: aplicarJa};
}

// Adendas seleccionadas por ID (vários campos da mesma adenda têm IDs diferentes).
function adendasPorId(ad: ExcelScript.Table, ids: string[]): number[] {
  const cab = cabecalho(ad), iId = cab.indexOf("ID_Adenda"), alvo = ids.map(x => String(x).trim().toUpperCase());
  const out: number[] = [];
  dadosTabela(ad).forEach((v, i) => { if (alvo.indexOf(textoCelula(v[iId]).toUpperCase()) >= 0) out.push(i); });
  return out;
}

function aplicarAdendas(workbook: ExcelScript.Workbook, p: Pedido, utilizador: string): Resultado {
  const func = workbook.getTable(T_FUNC), ad = workbook.getTable(T_AD);
  if (!func || !ad) return erro("Faltam as tabelas tbl_funcionarios ou tbl_adendas.");
  const ids = Array.isArray(p.ids) ? p.ids : [];
  const linhas = adendasPorId(ad, ids);
  if (!linhas.length) return erro("Adenda não encontrada.");
  const cab = cabecalho(ad), corpo = ad.getRangeBetweenHeaderAndTotal(), vals = dadosTabela(ad), hoje = hojeSerie();
  const c = (k: string) => cab.indexOf(k);
  for (const i of linhas) {
    const v = vals[i];
    if (textoCelula(v[c("Estado_Adenda")]).toLowerCase().startsWith("anulad")) return erro("A adenda " + textoCelula(v[c("ID_Adenda")]) + " está anulada.");
    if (textoCelula(v[c("Aplicada")]).toLowerCase() === "sim") continue;
    const iLinha = linhaColaborador(func, textoCelula(v[c("Empresa")]), normCodigo(v[c("Codigo")]));
    if (iLinha < 0) return erro("Colaborador da adenda " + textoCelula(v[c("ID_Adenda")]) + " não encontrado.");
    const e = aplicarNoCadastro(func, iLinha, [{campo: textoCelula(v[c("Campo_Alterado")]), novo: textoCelula(v[c("Valor_Novo")])}], utilizador);
    if (e) return erro(e);
    const r = corpo.getRow(i);
    r.getCell(0, c("Aplicada")).setValue("Sim");
    r.getCell(0, c("Data_Aplicacao")).setValue(hoje);
    if (c("Data_Ultima_Alteracao") >= 0) r.getCell(0, c("Data_Ultima_Alteracao")).setValue(agoraSerie());
    if (c("Utilizador_Alteracao") >= 0) r.getCell(0, c("Utilizador_Alteracao")).setValue(utilizador);
  }
  return {ok: true, accao: "aplicar", ids, aplicada: true};
}

function revogarAdendas(workbook: ExcelScript.Workbook, p: Pedido, utilizador: string): Resultado {
  const ad = workbook.getTable(T_AD);
  if (!ad) return erro("Tabela tbl_adendas não encontrada.");
  const ids = Array.isArray(p.ids) ? p.ids : [];
  const linhas = adendasPorId(ad, ids);
  if (!linhas.length) return erro("Adenda não encontrada.");
  const motivo = String(p.motivo || "").trim();
  if (!motivo) return erro("Indique o motivo da anulação.");
  const cab = cabecalho(ad), corpo = ad.getRangeBetweenHeaderAndTotal(), vals = dadosTabela(ad);
  const c = (k: string) => cab.indexOf(k);
  for (const i of linhas) {
    const r = corpo.getRow(i), obs = textoCelula(vals[i][c("Observacoes_Adicionais")]);
    r.getCell(0, c("Estado_Adenda")).setValue("Anulada");
    r.getCell(0, c("Observacoes_Adicionais")).setValue((obs ? obs + " | " : "") + "Anulada por " + utilizador + ": " + motivo);
    if (c("Data_Ultima_Alteracao") >= 0) r.getCell(0, c("Data_Ultima_Alteracao")).setValue(agoraSerie());
    if (c("Utilizador_Alteracao") >= 0) r.getCell(0, c("Utilizador_Alteracao")).setValue(utilizador);
  }
  return {ok: true, accao: "revogar", ids};
}

// ---------- Escalões salariais ----------
const FOLHA_ESC = "Escaloes";
const T_ESC = "tbl_escaloes_lancamento";
const T_LANC = "tbl_lançamento";
const COL_ESC = ["Empresa", "Criterio", "Escalao", "Valor_Inicial", "Ordem", "Data Ultima Alteracao", "Utilizador Alteracao"];
const CRITERIOS = ["Salario Liquido", "Vencimento Base", "Vencimento"];
const COL_ESCALAO = "Escalão";

function escaloes(workbook: ExcelScript.Workbook, p: Pedido, utilizador: string): Resultado {
  const op = String(p.op || "ler").trim().toLowerCase();
  if (op === "ler") {
    const t = workbook.getTable(T_ESC);
    return {ok: true, accao: "escaloes", listaEscaloes: t ? ler(t) : "[]"};
  }
  if (op === "gravar") return gravarEscaloes(workbook, p, utilizador);
  if (op === "aplicar") return aplicarEscaloesMes(workbook, p);
  return erro("Operação de escalões desconhecida: " + op);
}

function tabelaEscaloes(workbook: ExcelScript.Workbook): ExcelScript.Table {
  let t = workbook.getTable(T_ESC);
  if (t) return t;
  let folha = workbook.getWorksheet(FOLHA_ESC);
  if (!folha) folha = workbook.addWorksheet(FOLHA_ESC);
  const r = folha.getRangeByIndexes(0, 0, 1, COL_ESC.length);
  r.setValues([COL_ESC]);
  t = workbook.addTable(r, true); t.setName(T_ESC);
  t.getColumnByName("Valor_Inicial").getRangeBetweenHeaderAndTotal().setNumberFormat("#,##0.00");
  t.getColumnByName("Data Ultima Alteracao").getRangeBetweenHeaderAndTotal().setNumberFormat("dd/mm/yyyy hh:mm");
  return t;
}

function gravarEscaloes(workbook: ExcelScript.Workbook, p: Pedido, utilizador: string): Resultado {
  const erros: string[] = [];
  const empresa = String(p.empresa || "").trim(), criterio = String(p.criterio || "").trim();
  if (!empresa) erros.push("Escolha a empresa.");
  if (CRITERIOS.indexOf(criterio) < 0) erros.push("Critério inválido.");
  const faixas = Array.isArray(p.faixas) ? p.faixas : [];
  if (!faixas.length || faixas.length > 50) erros.push("Defina entre 1 e 50 escalões.");
  const nomes: string[] = [], inicios: number[] = [];
  faixas.forEach((f, i) => {
    const n = String(f.escalao ?? "").trim(), v = Number(f.inicio);
    if (!n) erros.push("O escalão " + (i + 1) + " não tem nome.");
    else if (nomes.indexOf(n.toUpperCase()) >= 0) erros.push("O nome «" + n + "» está repetido.");
    nomes.push(n.toUpperCase());
    if (!isFinite(v) || v < 0) erros.push("Valor inicial inválido no escalão " + (i + 1) + ".");
    else if (inicios.indexOf(v) >= 0) erros.push("Dois escalões começam no mesmo valor (" + v + ").");
    inicios.push(v);
  });
  if (erros.length) return {ok: false, accao: "escaloes", erro: erros.join(" "), erros};

  const t = tabelaEscaloes(workbook);
  limparFiltro(t);
  const cab = cabecalho(t), iE = cab.indexOf("Empresa");
  // Remove (de baixo para cima) as linhas desta empresa e acrescenta a nova definição.
  const vals = dadosTabela(t);
  for (let i = vals.length - 1; i >= 0; i--) {
    if (textoCelula(vals[i][iE]).toUpperCase() === empresa.toUpperCase()) t.deleteRowsAt(i, 1);
  }
  const agora = agoraSerie();
  const ordenadas = faixas.map(f => ({escalao: String(f.escalao).trim(), inicio: cent(Number(f.inicio))})).sort((a, b) => a.inicio - b.inicio);
  const linhas = ordenadas.map((f, i) => linhaPorCabecalho(t, {
    "Empresa": empresa, "Criterio": criterio, "Escalao": f.escalao, "Valor_Inicial": f.inicio, "Ordem": i + 1,
    "Data Ultima Alteracao": agora, "Utilizador Alteracao": String(utilizador || "")
  }));
  inserir(t, linhas);
  return {ok: true, accao: "escaloes", alteradas: linhas.length, listaEscaloes: ler(t)};
}

interface DefEscalao { criterio: string; faixas: Faixa[]; }
function definicoesEscaloes(workbook: ExcelScript.Workbook): Map<string, DefEscalao> {
  const mapa = new Map<string, DefEscalao>();
  const t = workbook.getTable(T_ESC);
  if (!t) return mapa;
  const cab = cabecalho(t);
  const iE = cab.indexOf("Empresa"), iC = cab.indexOf("Criterio"), iN = cab.indexOf("Escalao"), iV = cab.indexOf("Valor_Inicial");
  dadosTabela(t).forEach(v => {
    const e = textoCelula(v[iE]).toUpperCase(), n = textoCelula(v[iN]), x = Number(v[iV]);
    if (!e || !n || textoCelula(v[iV]) === "" || !isFinite(x)) return;
    let d = mapa.get(e);
    if (!d) { d = {criterio: textoCelula(v[iC]) || "Salario Liquido", faixas: []}; mapa.set(e, d); }
    d.faixas.push({escalao: n, inicio: x});
  });
  mapa.forEach(d => d.faixas.sort((a, b) => a.inicio - b.inicio));
  return mapa;
}
function escalaoDe(mapa: Map<string, DefEscalao>, empresa: string, valorDe: (coluna: string) => Celula): string {
  const d = mapa.get(empresa.toUpperCase()) || mapa.get("TODAS");
  if (!d) return "";
  const bruto = valorDe(d.criterio);
  if (bruto === "" || bruto === null || bruto === undefined) return "";
  const v = Number(bruto);
  if (!isFinite(v)) return "";
  let r = "";
  d.faixas.forEach(f => { if (v >= f.inicio) r = f.escalao; });
  return r;
}

function aplicarEscaloesMes(workbook: ExcelScript.Workbook, p: Pedido): Resultado {
  const ano = Number(p.ano), mes = String(p.mesNome || "").trim().toUpperCase();
  if (!Number.isInteger(ano) || !mes) return erro("Indique o ano e o mês.");
  const mapa = definicoesEscaloes(workbook);
  if (!mapa.size) return erro("Ainda não há escalões definidos.");
  const t = workbook.getTable(T_LANC);
  if (!t) return erro("Tabela tbl_lançamento não encontrada.");
  const cab = cabecalho(t), j = cab.indexOf(COL_ESCALAO);
  if (j < 0) return erro("A tbl_lançamento não tem a coluna «" + COL_ESCALAO + "».");
  const col = t.getColumnByName(COL_ESCALAO).getRangeBetweenHeaderAndTotal();
  if (col.getFormulas().some(f => String(f[0]).startsWith("="))) return erro("A coluna «" + COL_ESCALAO + "» tem fórmulas; não foi alterada.");
  const vals = dadosTabela(t);
  const iA = cab.indexOf("Ano"), iM = cab.indexOf("Mês"), iE = cab.indexOf("Empresa"), iS = cab.indexOf("Estado do Lançamento");
  let alteradas = 0, semEscalao = 0, doMes = 0;
  const novos: Celula[][] = vals.map(v => {
    if (Number(v[iA]) !== ano || textoCelula(v[iM]).toUpperCase() !== mes) return [v[j]];
    doMes++;
    if (iS >= 0 && textoCelula(v[iS]).toUpperCase() === "FINALIZADO") return [v[j]];
    const e = escalaoDe(mapa, textoCelula(v[iE]), c => { const k = cab.indexOf(c); return k >= 0 ? v[k] : ""; });
    if (!e) semEscalao++;
    if (textoCelula(v[j]) !== e) alteradas++;
    return [e];
  });
  if (!doMes) return erro("Não há lançamentos para esse mês.");
  if (alteradas) col.setValues(novos);
  return {ok: true, accao: "escaloes", alteradas, semEscalao};
}

// ---------- Acessos (tbl_utilizadores_RH) ----------
const T_UTIL = "tbl_utilizadores_RH";
const PERFIS_RH = ["Superadministrador", "Administrador RH", "Técnico RH", "Gestor"];
const COL_UTIL_EXTRA = ["Trocar_Codigo", "Data Ultima Alteracao", "Utilizador Alteracao"];
const HASH = /^[0-9a-f]{64}$/;

function semAcento(v: string): string { return v.normalize("NFD").replace(/[\u0300-\u036f]/g, "").trim().toLowerCase(); }
function perfilValido(v: string): string { const k = semAcento(v); for (const p of PERFIS_RH) if (semAcento(p) === k) return p; return ""; }
function emailNorm(v: unknown): string { return String(v ?? "").trim().toLowerCase(); }

// Acrescenta no fim da tabela as colunas de apoio que faltarem (não mexe nas existentes).
function colunasUtilizadores(t: ExcelScript.Table) {
  const cab = cabecalho(t);
  for (const c of COL_UTIL_EXTRA) if (cab.indexOf(c) < 0) t.addColumn(-1, undefined, c);
}
interface Conta { i: number; email: string; perfil: string; activo: boolean; hash: string; }
function contas(t: ExcelScript.Table): Conta[] {
  const cab = cabecalho(t), iE = cab.indexOf("Email"), iP = cab.indexOf("Perfil"), iS = cab.indexOf("Estado"), iH = cab.indexOf("Codigo_Hash");
  return dadosTabela(t).map((v, i) => ({i, email: emailNorm(v[iE]), perfil: perfilValido(textoCelula(v[iP])),
    activo: semAcento(textoCelula(v[iS])) === "activo", hash: textoCelula(v[iH]).toLowerCase()})).filter(c => c.email !== "");
}
function definirCelulas(t: ExcelScript.Table, i: number, valores: { [k: string]: Celula }) {
  const cab = cabecalho(t), linha = t.getRangeBetweenHeaderAndTotal().getRow(i);
  Object.keys(valores).forEach(k => { const j = cab.indexOf(k); if (j >= 0) linha.getCell(0, j).setValue(valores[k]); });
}

function trocarCodigo(workbook: ExcelScript.Workbook, p: Pedido, utilizador: string): Resultado {
  const t = workbook.getTable(T_UTIL);
  if (!t) return erro("Tabela tbl_utilizadores_RH não encontrada.");
  const eu = emailNorm(utilizador), c = contas(t).filter(x => x.email === eu);
  if (c.length !== 1 || !c[0].activo) return erro("Utilizador não encontrado ou inactivo.");
  const cab = cabecalho(t), iT = cab.indexOf("Trocar_Codigo");
  const trocar = iT >= 0 && semAcento(textoCelula(dadosTabela(t)[c[0].i][iT])) === "sim";
  if (String(p.op || "") === "estado") return {ok: true, accao: "trocar", trocar};
  const actual = String(p.chaveActual || "").toLowerCase(), nova = String(p.chaveNova || "").toLowerCase();
  if (!HASH.test(actual) || actual !== c[0].hash) return erro("O código actual não está certo.");
  if (!HASH.test(nova)) return erro("Novo código inválido.");
  if (nova === actual) return erro("O novo código tem de ser diferente do actual.");
  colunasUtilizadores(t);
  definirCelulas(t, c[0].i, {"Codigo_Hash": nova, "Trocar_Codigo": "Não", "Data Ultima Alteracao": agoraSerie(), "Utilizador Alteracao": eu});
  return {ok: true, accao: "trocar", trocar: false};
}

function acessos(workbook: ExcelScript.Workbook, p: Pedido, utilizador: string): Resultado {
  const t = workbook.getTable(T_UTIL);
  if (!t) return erro("Tabela tbl_utilizadores_RH não encontrada.");
  const eu = emailNorm(utilizador), lista = contas(t);
  const quem = lista.filter(x => x.email === eu);
  if (quem.length !== 1 || !quem[0].activo || quem[0].perfil !== "Superadministrador") return erro("Só um Superadministrador activo pode gerir acessos.");
  const op = String(p.op || "listar");
  if (op === "listar") return {ok: true, accao: "acessos", listaUtilizadores: listarUtilizadores(t)};
  colunasUtilizadores(t);
  if (op === "codigo") {
    const alvo = lista.filter(x => x.email === emailNorm(p.email));
    if (alvo.length !== 1) return erro("Utilizador não encontrado.");
    const chave = String(p.chave || "").toLowerCase();
    if (!HASH.test(chave)) return erro("Chave do código temporário inválida.");
    const proprio = alvo[0].email === eu;
    definirCelulas(t, alvo[0].i, {"Codigo_Hash": chave, "Trocar_Codigo": proprio ? "Não" : "Sim", "Data Ultima Alteracao": agoraSerie(), "Utilizador Alteracao": eu});
    return {ok: true, accao: "acessos", listaUtilizadores: listarUtilizadores(t)};
  }
  if (op === "gravar") return gravarUtilizador(t, lista, p, eu);
  return erro("Operação de acessos desconhecida: " + op);
}

function listarUtilizadores(t: ExcelScript.Table): string {
  const cab = cabecalho(t);
  const out: { [k: string]: Celula }[] = [];
  dadosTabela(t).filter(v => !vazia(v)).forEach(v => {
    const o: { [k: string]: Celula } = {};
    cab.forEach((h, j) => { if (h !== "Codigo_Hash") o[h] = v[j]; });
    o["Tem_Codigo"] = HASH.test(textoCelula(v[cab.indexOf("Codigo_Hash")]).toLowerCase());
    out.push(o);
  });
  return JSON.stringify(out);
}

function gravarUtilizador(t: ExcelScript.Table, lista: Conta[], p: Pedido, eu: string): Resultado {
  const erros: string[] = [];
  const email = emailNorm(p.email), nome = String(p.nome || "").trim(), perfil = perfilValido(String(p.perfil || ""));
  const estado = semAcento(String(p.estado || "Activo")) === "inactivo" ? "Inactivo" : "Activo";
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) erros.push("E-mail inválido.");
  if (!nome) erros.push("Indique o nome.");
  if (!perfil) erros.push("Perfil inválido.");
  const existente = lista.filter(x => x.email === email);
  if (existente.length > 1) erros.push("Este e-mail está repetido na tbl_utilizadores_RH.");
  const novo = existente.length === 0;
  const chave = String(p.chave || "").toLowerCase();
  if (novo && !HASH.test(chave)) erros.push("Defina o código temporário do novo utilizador.");
  if (!novo && existente[0].email === eu && (perfil !== "Superadministrador" || estado !== "Activo"))
    erros.push("Não pode retirar o seu próprio acesso de Superadministrador.");
  // Tem de ficar pelo menos um Superadministrador activo.
  const restantes = lista.filter(x => x.email !== email && x.activo && x.perfil === "Superadministrador").length;
  if (!novo && restantes === 0 && (perfil !== "Superadministrador" || estado !== "Activo")) erros.push("Tem de existir pelo menos um Superadministrador activo.");
  if (erros.length) return {ok: false, accao: "acessos", erro: erros.join(" "), erros};
  const agora = agoraSerie();
  const valores: { [k: string]: Celula } = {"Nome": nome, "Email": email, "Perfil": perfil, "Empresa": String(p.empresa || "").trim(),
    "Estado": estado, "Data Ultima Alteracao": agora, "Utilizador Alteracao": eu};
  if (novo) {
    const cab = cabecalho(t), iId = cab.indexOf("ID_Utilizador");
    let maior = 0;
    dadosTabela(t).forEach(v => { const m = /(\d+)\s*$/.exec(textoCelula(v[iId])); if (m) maior = Math.max(maior, Number(m[1])); });
    const prefixo = (() => { const v = dadosTabela(t).map(x => textoCelula(x[iId])).find(x => x !== ""); const m = v ? /^(.*?)(\d+)$/.exec(v) : null; return m ? m[1] : "U"; })();
    valores["ID_Utilizador"] = prefixo + String(maior + 1).padStart(3, "0");
    valores["Codigo_Hash"] = chave;
    valores["Trocar_Codigo"] = "Sim";
    limparFiltro(t);
    inserir(t, [linhaPorCabecalho(t, valores)]);
  } else {
    definirCelulas(t, existente[0].i, valores);
  }
  return {ok: true, accao: "acessos", listaUtilizadores: listarUtilizadores(t)};
}
