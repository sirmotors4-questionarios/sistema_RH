// Leitura das linhas devolvidas pelo fluxo RH_API_Consultar.
// O conector Excel devolve os cabeçalhos tal como estão no ficheiro (com espaços e acentos),
// números como texto e datas como número de série do Excel. Este módulo normaliza esses valores
// para apresentação; não altera nem calcula salários.

const TECNICAS = new Set(['ItemInternalId', '@odata.etag']);
const MESES = ['janeiro', 'fevereiro', 'março', 'abril', 'maio', 'junho', 'julho', 'agosto', 'setembro', 'outubro', 'novembro', 'dezembro'];

const texto = v => String(v ?? '').trim();
const chave = v => texto(v).normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase();

// Procura o valor de uma coluna tolerando diferenças de espaços, "_" e acentos no cabeçalho.
export function campo(linha, ...nomes) {
  if (!linha) return '';
  for (const nome of nomes) if (Object.hasOwn(linha, nome)) return linha[nome];
  const alvo = new Set(nomes.map(n => chave(n).replace(/[\s_]/g, '')));
  for (const [k, v] of Object.entries(linha)) if (alvo.has(chave(k).replace(/[\s_]/g, ''))) return v;
  return '';
}

export function numero(v) {
  if (typeof v === 'number') return Number.isFinite(v) ? v : null;
  let s = texto(v).replace(/\s|MT|MZN/gi, '');
  if (!s) return null;
  // "17.000,50" (pt) ou "17,000.50" (en)
  if (/,\d{1,2}$/.test(s)) s = s.replace(/\./g, '').replace(',', '.');
  else s = s.replace(/,/g, '');
  const n = Number(s);
  return Number.isFinite(n) ? n : null;
}

// Devolve AAAA-MM-DD ou '' (aceita série Excel, ISO e dd/mm/aaaa).
export function data(v) {
  const s = texto(v);
  if (!s) return '';
  if (/^\d+(\.\d+)?$/.test(s)) {
    const n = Number(s);
    if (n < 1 || n > 200000) return '';
    return new Date(Date.UTC(1899, 11, 30) + Math.floor(n) * 86400000).toISOString().slice(0, 10);
  }
  let m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m) return m[1] + '-' + m[2] + '-' + m[3];
  m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (m) return m[3] + '-' + m[2].padStart(2, '0') + '-' + m[1].padStart(2, '0');
  return '';
}

export function formatarData(v) {
  const d = data(v);
  return d ? d.slice(8, 10) + '/' + d.slice(5, 7) + '/' + d.slice(0, 4) : texto(v);
}

export function formatarValor(v) {
  const n = numero(v);
  return n === null ? texto(v) : new Intl.NumberFormat('pt-MZ', {minimumFractionDigits: 2, maximumFractionDigits: 2}).format(n);
}

export function mes(v) {
  const n = numero(v);
  if (n !== null && n >= 1 && n <= 12) return Math.trunc(n);
  const i = MESES.indexOf(chave(v).replace('marco', 'março'));
  if (i >= 0) return i + 1;
  const curto = MESES.findIndex(m => chave(m).slice(0, 3) === chave(v).slice(0, 3));
  return curto >= 0 && chave(v).length >= 3 ? curto + 1 : null;
}

export function activo(estado) {
  return ['activo', 'ativo', 'active'].includes(chave(estado));
}

export function colunas(linhas) {
  const vistas = new Set();
  for (const linha of linhas) for (const k of Object.keys(linha)) if (!TECNICAS.has(k)) vistas.add(k);
  return [...vistas];
}

export function colaboradores(linhas) {
  return linhas.map(l => ({
    codigo: texto(campo(l, 'Codigo')),
    nome: texto(campo(l, 'Nome')),
    empresa: texto(campo(l, 'Empresa')),
    departamento: texto(campo(l, 'Departamento')),
    categoria: texto(campo(l, 'Categoria')),
    profissao: texto(campo(l, 'Profissao')),
    regime: texto(campo(l, 'RegimePagamento')),
    contrato: texto(campo(l, 'TipoContrato')),
    admissao: data(campo(l, 'Data Admissao')),
    estado: texto(campo(l, 'Estado'))
  })).filter(c => c.codigo || c.nome);
}

export function filtrarColaboradores(lista, termo = '', estado = '') {
  const t = chave(termo);
  return lista.filter(c => (!estado || chave(c.estado) === chave(estado)) &&
    (!t || [c.codigo, c.nome, c.empresa, c.departamento].some(x => chave(x).includes(t))));
}

export function adendas(linhas) {
  return linhas.map(l => ({
    id: texto(campo(l, 'ID_Adenda')),
    numero: texto(campo(l, 'Numero_Adenda')),
    codigo: texto(campo(l, 'Codigo')),
    nome: texto(campo(l, 'Nome')),
    empresa: texto(campo(l, 'Empresa')),
    campo: texto(campo(l, 'Campo_Alterado')),
    anterior: texto(campo(l, 'Valor_Anterior')),
    novo: texto(campo(l, 'Valor_Novo')),
    inicio: data(campo(l, 'Data_Inicio_Vigencia')),
    fim: data(campo(l, 'Data_Fim_Vigencia')),
    estado: texto(campo(l, 'Estado_Adenda')),
    aplicada: texto(campo(l, 'Aplicada'))
  })).sort((a, b) => (b.inicio || '').localeCompare(a.inicio || ''));
}

// Parâmetros vigentes numa data (AAAA-MM-DD). Linhas sem datas contam como vigentes.
export function parametrosVigentes(linhas, em) {
  return linhas.map(l => ({
    regime: texto(campo(l, 'RegimePagamento')),
    categoria: texto(campo(l, 'Categoria')),
    valor: campo(l, 'Valor_Unitario'),
    unidade: texto(campo(l, 'Unidade')),
    inicio: data(campo(l, 'Vigencia_Inicio')),
    fim: data(campo(l, 'Vigencia_Fim')),
    activo: texto(campo(l, 'Activo'))
  })).filter(p => (!p.inicio || p.inicio <= em) && (!p.fim || p.fim >= em) && !['nao', 'não', 'false', '0'].includes(chave(p.activo)));
}

export function contagem(lista, chaveFn) {
  const mapa = new Map();
  for (const item of lista) {
    const k = texto(chaveFn(item)) || '(sem valor)';
    mapa.set(k, (mapa.get(k) || 0) + 1);
  }
  return [...mapa.entries()].sort((a, b) => b[1] - a[1]);
}

export function lancamentos(linhas, ano, mesRef) {
  return linhas.map(l => ({
    ano: numero(campo(l, 'Ano')),
    mes: mes(campo(l, 'Mês', 'Mes')),
    codigo: texto(campo(l, 'Codigo')),
    nome: texto(campo(l, 'Nome')),
    empresa: texto(campo(l, 'Empresa')),
    regime: texto(campo(l, 'RegimePagamento')),
    dias: campo(l, 'Dias Trabalhados'),
    horas: campo(l, 'Horas Trabalhadas'),
    acrescimos: campo(l, 'Total de Acréscimos'),
    descontos: campo(l, 'Total de Descontos'),
    irps: campo(l, 'IRPS'),
    liquido: campo(l, 'Salario Liquido'),
    estado: texto(campo(l, 'Estado do Lançamento')),
    revisto: texto(campo(l, 'Revisto'))
  })).filter(l => (ano == null || l.ano === Number(ano)) && (mesRef == null || l.mes === Number(mesRef)));
}

export function periodos(linhas) {
  const vistos = new Map();
  for (const l of lancamentos(linhas)) if (l.ano && l.mes) vistos.set(l.ano * 100 + l.mes, {ano: l.ano, mes: l.mes});
  return [...vistos.entries()].sort((a, b) => b[0] - a[0]).map(([, p]) => p);
}

export function totais(lista, campoNome) {
  return lista.reduce((s, l) => s + (numero(l[campoNome]) ?? 0), 0);
}

export function feriados(linhas) {
  return linhas.map(l => ({dia: numero(campo(l, 'Dia')), mes: mes(campo(l, 'Mes', 'Mês')), nome: texto(campo(l, 'Feriado'))}))
    .filter(f => f.dia && f.mes).sort((a, b) => a.mes - b.mes || a.dia - b.dia);
}

export const NOME_MES = MESES.map(m => m[0].toUpperCase() + m.slice(1));
