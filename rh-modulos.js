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
    codigo: normalizarCodigo(campo(l, 'Codigo')),
    nome: texto(campo(l, 'Nome', 'Nome Completo')),
    empresa: texto(campo(l, 'Empresa')),
    departamento: texto(campo(l, 'Departamento')),
    categoria: texto(campo(l, 'Categoria')),
    profissao: texto(campo(l, 'Profissao')),
    regime: texto(campo(l, 'RegimePagamento')),
    contrato: texto(campo(l, 'TipoContrato')),
    admissao: data(campo(l, 'Data Admissao')),
    estado: texto(campo(l, 'Estado')),
    linha: l
  })).filter(c => c.codigo || c.nome);
}

// Filtros como no VBA: Empresa -> Departamento (em cascata), estado e texto livre.
export function filtrarColaboradores(lista, termo = '', estado = '', {empresa = '', departamento = ''} = {}) {
  const t = chave(termo);
  return lista.filter(c => (!estado || chave(c.estado) === chave(estado)) &&
    (!empresa || chave(c.empresa) === chave(empresa)) &&
    (!departamento || chave(c.departamento) === chave(departamento)) &&
    (!t || [c.codigo, c.nome, c.empresa, c.departamento, c.profissao].some(x => chave(x).includes(t))));
}

// Opções em cascata: empresas; departamentos da empresa escolhida.
export function opcoesFiltro(lista, empresa = '') {
  const unico = arr => [...new Map(arr.filter(Boolean).map(v => [chave(v), v])).values()].sort((a, b) => a.localeCompare(b, 'pt'));
  return {
    empresas: unico(lista.map(c => c.empresa)),
    departamentos: unico(lista.filter(c => !empresa || chave(c.empresa) === chave(empresa)).map(c => c.departamento))
  };
}

export function adendas(linhas) {
  return linhas.map(l => ({
    id: texto(campo(l, 'ID_Adenda')),
    numero: texto(campo(l, 'Numero_Adenda')),
    data: data(campo(l, 'Data_Adenda')),
    codigo: texto(campo(l, 'Codigo')),
    nome: texto(campo(l, 'Nome')),
    empresa: texto(campo(l, 'Empresa')),
    campo: texto(campo(l, 'Campo_Alterado')),
    anterior: texto(campo(l, 'Valor_Anterior')),
    novo: texto(campo(l, 'Valor_Novo')),
    inicio: data(campo(l, 'Data_Inicio_Vigencia')),
    fim: data(campo(l, 'Data_Fim_Vigencia')),
    estado: texto(campo(l, 'Estado_Adenda')),
    aplicada: texto(campo(l, 'Aplicada')),
    motivo: texto(campo(l, 'Motivo')),
    documento: texto(campo(l, 'Documento_Referencia')),
    departamento: texto(campo(l, 'Departamento'))
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
    codigo: normalizarCodigo(campo(l, 'Codigo')),
    nome: texto(campo(l, 'Nome')),
    empresa: texto(campo(l, 'Empresa')),
    departamento: texto(campo(l, 'Departamento')),
    profissao: texto(campo(l, 'Profissao')),
    regime: texto(campo(l, 'RegimePagamento')),
    base: campo(l, 'Vencimento', 'Vencimento Base', 'Valor Unitário'),
    inssColab: campo(l, 'INSS Colaborador', 'INSS'),
    inssEmpresa: campo(l, 'INSS Empresa'),
    dias: campo(l, 'Dias Trabalhados'),
    horas: campo(l, 'Horas Trabalhadas'),
    acrescimos: campo(l, 'Total de Acréscimos'),
    descontos: campo(l, 'Total de Descontos'),
    irps: campo(l, 'IRPS'),
    liquido: campo(l, 'Salario Liquido'),
    estado: texto(campo(l, 'Estado do Lançamento')),
    revisto: texto(campo(l, 'Revisto')),
    linha: l
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


// ---------- chave do colaborador (igual ao VBA: modRH_ChaveEmpresaCodigo) ----------
// O mesmo Código pode existir em empresas diferentes; dentro da mesma empresa é único.
// Códigos numéricos são guardados com 4 dígitos ("12" -> "0012").
export function normalizarCodigo(v) {
  const s = texto(v);
  return /^\d+$/.test(s) ? String(Number(s)).padStart(4, '0') : s;
}
export function chaveRH(empresa, codigo) {
  return texto(empresa).toUpperCase() + '|' + normalizarCodigo(codigo).toUpperCase();
}

// ---------- ficha do colaborador ----------
// Separadores e campos iguais ao formulário G_Colaboradores do VBA. "alt" são os outros nomes de
// coluna que o VBA aceita; o site usa o cabeçalho que existir em tbl_funcionarios.
// tipo: texto | data | numero | email | escolha (lista fixa) | lista (valores existentes na coluna).
export const SEPARADORES = Object.freeze([
  {id: 'pessoais', nome: 'Dados pessoais'},
  {id: 'documentos', nome: 'Documentos'},
  {id: 'contactos', nome: 'Contactos'},
  {id: 'bancarios', nome: 'Dados bancários'},
  {id: 'profissionais', nome: 'Dados profissionais'},
  {id: 'contrato', nome: 'Contrato e adendas'}
]);

export const CAMPOS_COLABORADOR = Object.freeze([
  {k: 'Nome', alt: ['Nome Completo'], rotulo: 'Nome completo', sep: 'pessoais', obrig: true, max: 120, largo: true},
  {k: 'Nome Abreviado', rotulo: 'Nome abreviado', sep: 'pessoais', max: 60},
  {k: 'Sexo', rotulo: 'Sexo', sep: 'pessoais', tipo: 'escolha', opcoes: ['Masculino', 'Feminino']},
  {k: 'Data Nascimento', alt: ['Data de Nascimento', 'DataNascimento'], rotulo: 'Data de nascimento', sep: 'pessoais', tipo: 'data'},
  {k: 'EstadoCivil', alt: ['Estado Civil'], rotulo: 'Estado civil', sep: 'pessoais', tipo: 'lista'},
  {k: 'Nacionalidade', rotulo: 'Nacionalidade', sep: 'pessoais', tipo: 'lista'},
  {k: 'Naturalidade', rotulo: 'Naturalidade', sep: 'pessoais', max: 80},
  {k: 'TotalDependentes', rotulo: 'Dependentes', sep: 'pessoais', tipo: 'numero', inteiro: true},

  {k: 'T. Documento', alt: ['Tipo Documento', 'TipoDocumento'], rotulo: 'Tipo de documento', sep: 'documentos', tipo: 'lista'},
  {k: 'NumBI', alt: ['Num Documento', 'NumDocumento'], rotulo: 'N.º do documento', sep: 'documentos', max: 30},
  {k: 'Data Emissao', alt: ['Data Emissão', 'Data de Emissão', 'Data EmissaoBI'], rotulo: 'Data de emissão', sep: 'documentos', tipo: 'data'},
  {k: 'Validade do Documento', alt: ['Data Validade', 'Validade', 'ValidadeBI'], rotulo: 'Validade', sep: 'documentos', tipo: 'data'},
  {k: 'Local Emissao', alt: ['Local Emissão', 'Local de Emissão'], rotulo: 'Local de emissão', sep: 'documentos', max: 80},
  {k: 'NUIT', rotulo: 'NUIT', sep: 'documentos', padrao: /^\d{9}$/, ajuda: '9 dígitos'},
  {k: 'NumBeneficiario', alt: ['Num INSS', 'NumINSS'], rotulo: 'N.º de beneficiário INSS', sep: 'documentos', max: 30},

  {k: 'Telemovel', alt: ['Telemóvel'], rotulo: 'Telemóvel', sep: 'contactos', padrao: /^\+?[\d\s]{7,16}$/, ajuda: 'só dígitos, opcional +'},
  {k: 'email', alt: ['E-mail', 'Email'], rotulo: 'E-mail', sep: 'contactos', tipo: 'email'},
  {k: 'Localidade', rotulo: 'Localidade / residência', sep: 'contactos', max: 120, largo: true},
  {k: 'Distrito', rotulo: 'Distrito', sep: 'contactos', tipo: 'lista'},
  {k: 'CodPostal', alt: ['Codigo Postal', 'Código Postal'], rotulo: 'Código postal', sep: 'contactos', max: 20},

  {k: 'BancoFuncionario', alt: ['Banco Funcionario', 'Banco'], rotulo: 'Banco', sep: 'bancarios', tipo: 'lista'},
  {k: 'CodBanco', alt: ['Codigo Banco', 'Código Banco'], rotulo: 'Código do banco', sep: 'bancarios', max: 20},
  {k: 'NIB/Conta', alt: ['NIB', 'Conta', 'NIB Conta'], rotulo: 'NIB / conta', sep: 'bancarios', max: 40, largo: true},
  {k: 'ContaEmpresa', alt: ['Conta Empresa'], rotulo: 'Conta da empresa', sep: 'bancarios', max: 40},
  {k: 'CodEstabelecimento', alt: ['Cod Estabelecimento', 'Código Estabelecimento'], rotulo: 'Código do estabelecimento', sep: 'bancarios', max: 20},

  {k: 'Codigo', rotulo: 'Código', sep: 'profissionais', obrig: true, chave: true, max: 20},
  {k: 'Empresa', rotulo: 'Empresa', sep: 'profissionais', obrig: true, chave: true, tipo: 'lista', max: 80},
  {k: 'Departamento', rotulo: 'Departamento', sep: 'profissionais', tipo: 'lista'},
  {k: 'Profissao', alt: ['Profissão'], rotulo: 'Profissão', sep: 'profissionais', tipo: 'lista'},
  {k: 'Categoria', rotulo: 'Categoria', sep: 'profissionais', tipo: 'lista'},
  {k: 'TipoContrato', alt: ['Tipo Contrato'], rotulo: 'Tipo de contrato', sep: 'profissionais', tipo: 'lista'},
  {k: 'SituacaoProfissional', alt: ['Situação Profissional'], rotulo: 'Situação profissional', sep: 'profissionais', tipo: 'lista'},
  {k: 'Estado', rotulo: 'Estado', sep: 'profissionais', tipo: 'escolha', opcoes: ['Activo', 'Inactivo']},
  {k: 'Estado de Reforma', alt: ['Estado Reforma'], rotulo: 'Estado de reforma', sep: 'profissionais', tipo: 'lista'},
  {k: 'Data Admissao', alt: ['Data de Admissão', 'Data Admissão'], rotulo: 'Data de admissão', sep: 'profissionais', tipo: 'data'},
  {k: 'Data de Saída', alt: ['Data Saida', 'Data de Saida'], rotulo: 'Data de saída', sep: 'profissionais', tipo: 'data'},
  {k: 'Motivo', alt: ['Motivo Saida', 'Motivo de Saída'], rotulo: 'Motivo de saída', sep: 'profissionais', max: 120},
  {k: 'RegimePagamento', alt: ['Regime Pagamento'], rotulo: 'Regime de pagamento', sep: 'profissionais', tipo: 'escolha', opcoes: ['Mensal', 'Hora', 'Diário']},
  {k: 'Vencimento Base', rotulo: 'Vencimento base (MT/mês)', sep: 'profissionais', tipo: 'numero'},
  {k: 'Canalizar', rotulo: 'Ainda a canalizar?', sep: 'profissionais', tipo: 'escolha', opcoes: ['Sim', 'Não']}
]);

// Colunas de auditoria escritas em cada gravação (se existirem na tabela), como no VBA.
export const AUDITORIA = Object.freeze({data: 'Data Ultima Alteracao', utilizador: 'Utilizador Alteracao', criacao: 'Data Criacao'});
// Colunas calculadas por fórmulas no Excel: o site nunca as escreve.
export const CALCULADAS = Object.freeze(['Idade', 'Anos de Trabalho', 'INSS Colaborador', 'INSS Empresa', 'Total INSS', 'IRPS', 'Salario Liquido', 'Chave_RH']);

const semEspacos = s => chave(s).replace(/[\s_.]/g, '');
export const definicao = k => CAMPOS_COLABORADOR.find(c => c.k === k);

// Cabeçalho real de tbl_funcionarios para um campo (ou '' se a coluna não existir).
export function colunaReal(cabecalhos, def) {
  const nomes = [def.k, ...(def.alt || [])];
  for (const n of nomes) if (cabecalhos.includes(n)) return n;
  const alvo = new Set(nomes.map(semEspacos));
  return cabecalhos.find(h => alvo.has(semEspacos(h))) || '';
}

// Valores distintos de uma coluna, para as listas (como CarregarDistinct no VBA).
export function valoresDistintos(linhas, def, filtro) {
  const vistos = new Map();
  for (const l of linhas) {
    if (filtro && !filtro(l)) continue;
    const v = texto(campo(l, def.k, ...(def.alt || [])));
    if (v && !vistos.has(chave(v))) vistos.set(chave(v), v);
  }
  return [...vistos.values()].sort((a, b) => a.localeCompare(b, 'pt'));
}

export function serialExcel(iso) {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(texto(iso));
  if (!m) return '';
  return String(Math.round((Date.UTC(+m[1], +m[2] - 1, +m[3]) - Date.UTC(1899, 11, 30)) / 86400000));
}

export const regimeSemBase = regime => ['hora', 'diario'].includes(chave(regime));

// Valores de uma linha do Excel prontos para a ficha (datas AAAA-MM-DD, números sem formatação).
export function linhaParaFormulario(linha) {
  const out = {};
  for (const c of CAMPOS_COLABORADOR) {
    const v = campo(linha, c.k, ...(c.alt || []));
    out[c.k] = c.tipo === 'data' ? data(v) : c.tipo === 'numero' ? (numero(v) ?? '') : c.k === 'Codigo' ? normalizarCodigo(v) : texto(v);
  }
  return out;
}

// Encontra a linha de um colaborador pelo par Empresa + Código.
export function localizarColaborador(linhas, empresa, codigo) {
  const alvo = chaveRH(empresa, codigo);
  return linhas.find(l => chaveRH(campo(l, 'Empresa'), campo(l, 'Codigo')) === alvo) || null;
}

// Normaliza e valida a ficha. Em 'editar', só valida o formato dos campos que mudaram
// (dados antigos fora do formato não impedem gravar outras alterações).
export function validarColaborador(entrada, modo, {original = {}, existentes = []} = {}) {
  const dados = {}, erros = [];
  for (const c of CAMPOS_COLABORADOR) {
    let v = texto(entrada?.[c.k]);
    if (c.k === 'Codigo') v = normalizarCodigo(v);
    const mudou = modo === 'criar' || texto(original[c.k]) !== v;
    if (!v) {
      if (c.obrig) erros.push(c.rotulo + ' é obrigatório.');
      dados[c.k] = '';
      continue;
    }
    if (mudou) {
      if (c.max && v.length > c.max) erros.push(c.rotulo + ': máximo ' + c.max + ' caracteres.');
      if (c.tipo === 'escolha' && !c.opcoes.includes(v)) erros.push(c.rotulo + ': escolha um valor da lista.');
      if (c.padrao && !c.padrao.test(v)) erros.push(c.rotulo + ': formato inválido (' + c.ajuda + ').');
      if (c.tipo === 'email' && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(v)) erros.push(c.rotulo + ': formato inválido.');
      if (c.tipo === 'data' && !/^\d{4}-\d{2}-\d{2}$/.test(v)) erros.push(c.rotulo + ': data inválida.');
    }
    if (c.tipo === 'numero') {
      const n = numero(v);
      if (n === null || n < 0 || (c.inteiro && !Number.isInteger(n))) { if (mudou) erros.push(c.rotulo + ': número inválido.'); dados[c.k] = v; continue; }
      v = String(n);
    }
    dados[c.k] = v;
  }
  if (/['"<>|]/.test(dados.Codigo)) erros.push('Código: não use aspas, < > nem |.');
  if (/['"<>|]/.test(dados.Empresa)) erros.push('Empresa: não use aspas, < > nem |.');
  if (regimeSemBase(dados.RegimePagamento)) dados['Vencimento Base'] = '';
  else if (texto(dados.RegimePagamento) && !dados['Vencimento Base'] && modo === 'criar') erros.push('Vencimento base é obrigatório no regime Mensal.');
  if (dados['Data de Saída'] && dados['Data Admissao'] && dados['Data de Saída'] < dados['Data Admissao']) erros.push('A data de saída não pode ser anterior à admissão.');
  if (modo === 'criar' && dados.Codigo && dados.Empresa) {
    const k = chaveRH(dados.Empresa, dados.Codigo);
    if (existentes.some(l => chaveRH(campo(l, 'Empresa'), campo(l, 'Codigo')) === k))
      erros.push('Já existe o código ' + dados.Codigo + ' na empresa ' + dados.Empresa + '. O mesmo código só pode repetir-se noutra empresa.');
  }
  return {dados, erros};
}

// Corpo do pedido ao fluxo RH_API_Actualizar:
// - alvo: Empresa + Código (chave do colaborador; em 'editar' é o registo carregado e não muda);
// - dados: cabeçalhos reais de tbl_funcionarios -> valor (datas como número de série do Excel);
// - em 'editar' só segue o que mudou; Código e Empresa nunca mudam por edição (usar adenda).
export function prepararGravacao(modo, dados, {original = {}, cabecalhos = [], utilizador = '', agora = new Date()} = {}) {
  const alvo = modo === 'editar' ? {Empresa: original.Empresa, Codigo: normalizarCodigo(original.Codigo)} : {Empresa: dados.Empresa, Codigo: dados.Codigo};
  const saida = {}, ignorados = [];
  for (const c of CAMPOS_COLABORADOR) {
    if (modo === 'editar' && (c.chave || texto(original[c.k]) === texto(dados[c.k]))) continue;
    if (modo === 'criar' && !texto(dados[c.k])) continue;
    const col = colunaReal(cabecalhos, c);
    if (!col) { if (texto(dados[c.k])) ignorados.push(c.rotulo); continue; }
    saida[col] = c.tipo === 'data' ? serialExcel(dados[c.k]) : texto(dados[c.k]);
  }
  const alteracoes = Object.keys(saida).length;
  if (alteracoes) {
    const colData = colunaReal(cabecalhos, {k: AUDITORIA.data}), colUser = colunaReal(cabecalhos, {k: AUDITORIA.utilizador});
    const serie = String(Math.round(((agora.getTime() - agora.getTimezoneOffset() * 60000) / 86400000 + 25569) * 100000) / 100000);
    if (colData) saida[colData] = serie;
    if (colUser) saida[colUser] = utilizador;
    const colCriacao = modo === 'criar' && colunaReal(cabecalhos, {k: AUDITORIA.criacao});
    if (colCriacao) saida[colCriacao] = serie;
  }
  return {alvo, chave: chaveRH(alvo.Empresa, alvo.Codigo), dados: saida, alteracoes, ignorados};
}

// Indicador de preenchimento (VBA: 0–49% Baixo, 50–79% Médio, 80–100% Bom).
export function preenchimento(valores, sep) {
  const campos = CAMPOS_COLABORADOR.filter(c => !sep || c.sep === sep);
  const total = campos.length, cheios = campos.filter(c => texto(valores[c.k])).length;
  const pct = total ? Math.round(cheios / total * 100) : 0;
  return {preenchidos: cheios, total, pct, nivel: !total ? 'Sem campos' : pct < 50 ? 'Baixo' : pct < 80 ? 'Médio' : 'Bom'};
}
