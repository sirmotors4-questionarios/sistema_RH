// Escalões salariais por empresa (tbl_escaloes_lancamento): lógica pura, partilhada pela página de
// Configurações e pelos testes. A mesma regra está nos Office Scripts LancamentoNovoMes e RHIndemnizacoes.
//
// Cada empresa tem um critério (coluna do lançamento usada como valor) e uma lista de escalões, cada um
// com o valor a partir do qual começa. Um escalão vai do seu valor inicial (inclusive) até ao valor inicial
// do seguinte (exclusive); o último não tem limite. Assim não há sobreposições nem buracos entre escalões.
// A linha com Empresa = "Todas" vale para as empresas que não tenham escalões próprios.

export const TODAS = "Todas";
export const CRITERIOS = Object.freeze([
  {coluna: "Salario Liquido", rotulo: "Salário líquido"},
  {coluna: "Vencimento Base", rotulo: "Vencimento base"},
  {coluna: "Vencimento", rotulo: "Vencimento (bruto do mês)"}
]);
export const CRITERIO_PADRAO = "Salario Liquido";
export const COLUNAS = Object.freeze(["Empresa", "Criterio", "Escalao", "Valor_Inicial", "Ordem", "Data Ultima Alteracao", "Utilizador Alteracao"]);

const texto = v => String(v ?? "").trim();
const chave = v => texto(v).toUpperCase();
export function numero(v) {
  if (typeof v === "number") return Number.isFinite(v) ? v : null;
  let s = texto(v).replace(/\s/g, "");
  if (!s) return null;
  // 1.234,56 ou 1234,56 (pt) e 1234.56
  if (/,\d{1,2}$/.test(s)) s = s.replace(/\./g, "").replace(",", ".");
  else s = s.replace(/,/g, "");
  const n = Number(s);
  return Number.isFinite(n) ? n : null;
}

// Linhas do Excel -> {empresa: {empresa, criterio, faixas:[{escalao, inicio}]}}
export function agrupar(linhas = []) {
  const mapa = new Map();
  for (const l of linhas) {
    const empresa = texto(l.Empresa), nome = texto(l.Escalao), inicio = numero(l.Valor_Inicial);
    if (!empresa || !nome || inicio === null) continue;
    const k = chave(empresa);
    if (!mapa.has(k)) mapa.set(k, {empresa, criterio: texto(l.Criterio) || CRITERIO_PADRAO, faixas: []});
    mapa.get(k).faixas.push({escalao: nome, inicio, ordem: numero(l.Ordem) ?? 0});
  }
  for (const d of mapa.values()) d.faixas.sort((a, b) => a.inicio - b.inicio || a.ordem - b.ordem).forEach(f => delete f.ordem);
  return mapa;
}

// Definição que vale para uma empresa: a própria ou, na falta dela, a de "Todas".
export function definicaoDe(mapa, empresa) {
  return mapa.get(chave(empresa)) || mapa.get(chave(TODAS)) || null;
}

// Escalão de um valor: o último cujo valor inicial é <= valor. Abaixo do primeiro: "".
export function escalaoDe(faixas = [], valor) {
  const v = numero(valor);
  if (v === null) return "";
  let r = "";
  for (const f of faixas) if (v >= f.inicio) r = f.escalao;
  return r;
}

// Escalão de uma linha (lançamento ou colaborador) segundo a definição da sua empresa.
export function escalaoLinha(mapa, linha) {
  const d = definicaoDe(mapa, linha?.Empresa);
  if (!d) return "";
  return escalaoDe(d.faixas, linha?.[d.criterio]);
}

// Limite superior apresentado ao utilizador (valor inicial do escalão seguinte).
export function intervalos(faixas = []) {
  const ord = [...faixas].sort((a, b) => a.inicio - b.inicio);
  return ord.map((f, i) => ({...f, ate: i + 1 < ord.length ? ord[i + 1].inicio : null}));
}

export function validar({empresa, criterio, faixas} = {}) {
  const erros = [];
  if (!texto(empresa)) erros.push("Escolha a empresa.");
  if (!CRITERIOS.some(c => c.coluna === criterio)) erros.push("Critério inválido.");
  const lista = Array.isArray(faixas) ? faixas : [];
  if (!lista.length) erros.push("Defina pelo menos um escalão.");
  const nomes = new Set(), inicios = new Set();
  lista.forEach((f, i) => {
    const n = texto(f.escalao), v = numero(f.inicio);
    if (!n) erros.push("O escalão " + (i + 1) + " não tem nome.");
    else if (nomes.has(chave(n))) erros.push("O nome «" + n + "» está repetido.");
    else nomes.add(chave(n));
    if (v === null || v < 0) erros.push("Valor inicial inválido no escalão " + (i + 1) + ".");
    else if (inicios.has(v)) erros.push("Dois escalões começam no mesmo valor (" + v + ").");
    else inicios.add(v);
  });
  return erros;
}

// Pedido de gravação já normalizado (ordenado pelo valor inicial).
export function pedidoGravacao({empresa, criterio, faixas}) {
  const ord = (faixas || []).map(f => ({escalao: texto(f.escalao), inicio: numero(f.inicio)})).sort((a, b) => a.inicio - b.inicio);
  return {empresa: texto(empresa), criterio, faixas: ord};
}

// Simulação: quantos colaboradores ficam em cada escalão (com os valores do cadastro).
export function distribuicao(faixas, criterio, linhas = []) {
  const conta = new Map(faixas.map(f => [f.escalao, 0]));
  let semValor = 0, abaixo = 0;
  for (const l of linhas) {
    const v = numero(l?.[criterio]);
    if (v === null) { semValor++; continue; }
    const e = escalaoDe(faixas, v);
    if (!e) abaixo++; else conta.set(e, (conta.get(e) || 0) + 1);
  }
  return {porEscalao: [...conta].map(([escalao, total]) => ({escalao, total})), semValor, abaixo};
}
