// Regras do lançamento mensal, folha salarial e relatório comparativo — réplica do VBA:
// clsLancamentoMensal.CriarNovoMes, modRH_CondPag_Indet_v7, modRH_PeriodoFolha20_v3,
// modRH_Settings_FINAL (tarifas), modRH_CalculosResumo, modRH_PrevisualizarFolha,
// modRH_ExportarFolha e modRH_RelatorioComparativo.
// Funções puras: recebem linhas do Excel (cabeçalhos reais) e devolvem dados prontos.
import {campo, numero, data, mes as mesNumero, normalizarCodigo, chaveRH, NOME_MES} from "./rh-modulos.js";

const texto = v => String(v ?? "").trim();
const semAcento = s => texto(s).normalize("NFD").replace(/[̀-ͯ]/g, "").toUpperCase();
const igual = (a, b) => semAcento(a) === semAcento(b);
const n = v => numero(v) ?? 0;
const iso = (a, m, d) => new Date(Date.UTC(a, m - 1, d)).toISOString().slice(0, 10);
const diasEntre = (a, b) => Math.round((Date.parse(b) - Date.parse(a)) / 86400000);
export const serie = isoData => isoData ? String(diasEntre("1899-12-30", isoData)) : "";

// ---------- período 21–20 ----------
// Abril/2026 -> 21/03/2026 a 20/04/2026. O período pertence ao mês nominal.
export function periodoFolha(ano, mes) {
  return {inicio: iso(ano, mes - 1, 21), fim: iso(ano, mes, 20), referencia: iso(ano, mes, 1)};
}

// Mensal: 30 dias por período completo; admissão/saída dentro do período = dias de calendário, máx. 30.
export function diasSalario30(inicioPeriodo, fimPeriodo, dataInicio = "", dataFim = "") {
  const ini = dataInicio && dataInicio > inicioPeriodo ? dataInicio : inicioPeriodo;
  const fim = dataFim && dataFim < fimPeriodo ? dataFim : fimPeriodo;
  if (ini > fim) return 0;
  if (ini <= inicioPeriodo && fim >= fimPeriodo) return 30;
  return Math.min(diasEntre(ini, fim) + 1, 30);
}

// ---------- adendas e contrato ----------
const ESTADOS_ADENDA = ["ACTIVA", "ATIVA", "ACTIVO", "ATIVO", "APLICADA", "APLICADO"];
const INDETERMINADO = ["INDETERMINADO", "TEMPO INDETERMINADO", "POR TEMPO INDETERMINADO", "CONTRATO INDETERMINADO", "SEM TERMO", "EFECTIVO", "EFETIVO"];
const ENCERRADO = ["INACTIVO", "INATIVO", "CESSADO", "TERMINADO", "RESCINDIDO", "DESLIGADO", "ENCERRADO"];
const CAMPOS_FIM_CONTRATO = ["Data Fim Contrato", "Data Fim do Contrato", "Data_Fim_Contrato", "Fim Contrato", "Data Termino Contrato", "Data Término Contrato", "Validade Contrato"];
export const tipoIndeterminado = t => INDETERMINADO.includes(semAcento(t));
export const estadoEncerrado = e => ENCERRADO.includes(semAcento(e));

// Valor de uma adenda ACTIVA e vigente na data (a mais recente prevalece). Empresa vazia na adenda = registo antigo, aplica-se.
export function adendaVigente(adendas, empresa, codigo, campoAlterado, dataRef) {
  let melhor = null;
  const cod = normalizarCodigo(codigo).toUpperCase();
  for (const a of adendas) {
    if (normalizarCodigo(campo(a, "Codigo")).toUpperCase() !== cod) continue;
    const emp = texto(campo(a, "Empresa"));
    if (emp && empresa && !igual(emp, empresa)) continue;
    if (!igual(campo(a, "Campo_Alterado"), campoAlterado)) continue;
    if (!ESTADOS_ADENDA.includes(semAcento(campo(a, "Estado_Adenda")))) continue;
    const efectiva = data(campo(a, "Data_Inicio_Vigencia")) || data(campo(a, "Data_Adenda"));
    if (!efectiva || efectiva > dataRef) continue;
    const fim = data(campo(a, "Data_Fim_Vigencia"));
    if (fim && dataRef > fim) continue;
    if (!melhor || efectiva >= melhor.data) melhor = {data: efectiva, valor: campo(a, "Valor_Novo")};
  }
  return melhor;
}
// Adendas no formato antigo (tbl_adendas_colaborador + tbl_adendas_detalhes), como RH_ProcurarAdendaLegacy:
// cada detalhe herda Código, Empresa, estado e vigência do cabeçalho com o mesmo ID_Adenda.
export function unirAdendas(unificadas = [], cabecalhos = [], detalhes = []) {
  const porId = new Map();
  for (const c of cabecalhos) { const id = texto(campo(c, "ID_Adenda")); if (id) porId.set(id.toUpperCase(), c); }
  const legadas = [];
  for (const d of detalhes) {
    const c = porId.get(texto(campo(d, "ID_Adenda")).toUpperCase());
    if (!c) continue;
    legadas.push({Codigo: campo(c, "Codigo"), Empresa: campo(c, "Empresa"), Estado_Adenda: campo(c, "Estado_Adenda"),
      Data_Adenda: campo(c, "Data_Adenda"), Data_Inicio_Vigencia: campo(c, "Data_Inicio_Vigencia"), Data_Fim_Vigencia: campo(c, "Data_Fim_Vigencia"),
      Campo_Alterado: campo(d, "Campo_Alterado"), Valor_Novo: campo(d, "Valor_Novo"), ID_Adenda: campo(c, "ID_Adenda")});
  }
  return [...unificadas, ...legadas];
}
const vigente = (ad, f, campoAlterado, dataRef, valorCadastro) => {
  const a = adendaVigente(ad, campo(f, "Empresa"), campo(f, "Codigo"), campoAlterado, dataRef);
  return a ? a.valor : valorCadastro;
};

function fimContratoVigente(f, adendas, dataRef) {
  let melhor = null;
  for (const c of CAMPOS_FIM_CONTRATO) {
    const a = adendaVigente(adendas, campo(f, "Empresa"), campo(f, "Codigo"), c, dataRef);
    if (a && (!melhor || a.data >= melhor.data)) melhor = a;
  }
  if (melhor) return data(melhor.valor);
  for (const c of CAMPOS_FIM_CONTRATO) { const v = data(campo(f, c)); if (v) return v; }
  return "";
}

// RH_ContratoVigenteNoPeriodo + RH_LimitesContratoNoPeriodo.
export function contratoNoPeriodo(f, adendas, {inicio, fim}) {
  const tipo = texto(vigente(adendas, f, "TipoContrato", fim, campo(f, "TipoContrato")));
  const estado = texto(vigente(adendas, f, "Estado", fim, campo(f, "Estado")));
  const admissao = data(vigente(adendas, f, "Data Admissao", fim, campo(f, "Data Admissao", "Data Admissão", "Data de Admissao", "Data de Admissão")));
  const saida = data(vigente(adendas, f, "Data de Saída", fim, campo(f, "Data de Saída", "Data de Saida", "Data Saída", "Data Saida", "Data Demissão", "Data Demissao")));
  const prazo = fimContratoVigente(f, adendas, fim);
  const indet = tipoIndeterminado(tipo);
  const r = {vigente: false, motivo: "", inicioContrato: admissao, fimContrato: ""};
  if (admissao && admissao > fim) return {...r, motivo: "Admissão posterior ao fim do período 21–20."};
  if (saida && saida < inicio) return {...r, motivo: "Contrato cessado antes deste período."};
  if (!indet && prazo && prazo < inicio) return {...r, motivo: "Prazo do contrato terminou antes deste período."};
  if (estadoEncerrado(estado) && !saida) return {...r, motivo: "Colaborador inactivo (sem data de saída)."};
  let fimEf = saida;
  if (!indet && prazo && (!fimEf || prazo < fimEf)) fimEf = prazo;
  return {vigente: true, motivo: indet ? "Contrato indeterminado vigente." : prazo ? "Contrato com prazo vigente." : "Contrato vigente; sem data final cadastrada.",
    inicioContrato: admissao, fimContrato: fimEf};
}

// Tarifa das Settings (tbl_parametros_pagamento): Regime + Categoria (= Profissão) + vigência; "Não" em Activo exclui.
export function tarifaVigente(parametros, regime, categoria, dataRef) {
  const encontradas = parametros.filter(p => igual(campo(p, "RegimePagamento"), regime) && igual(campo(p, "Categoria"), categoria) &&
    !igual(campo(p, "Activo"), "Não")).filter(p => {
    const ini = data(campo(p, "Vigencia_Inicio")), fim = data(campo(p, "Vigencia_Fim"));
    return ini && dataRef >= ini && (!fim || dataRef <= fim);
  });
  if (encontradas.length > 1) return {erro: "Várias tarifas vigentes para " + regime + " / " + categoria + "."};
  const p = encontradas[0];
  return p ? {valor: n(campo(p, "Valor_Unitario")), unidade: texto(campo(p, "Unidade"))} : {valor: 0};
}

// ---------- Novo mês ----------
export const ENTRADAS_COPIADAS = ["Canalizar", "TotalDependentes", "Faltas Justificadas", "Faltas Injustificadas", "Horas Extras", "Subsídios", "Bónus", "Outros Acréscimos", "Adiantamentos", "Empréstimos", "Outros Descontos", "Observações"];
const ENTRADAS_ZERO = ["Faltas Justificadas", "Faltas Injustificadas", "Horas Extras", "Subsídios", "Bónus", "Outros Acréscimos", "Adiantamentos", "Empréstimos", "Outros Descontos"];
// Colunas calculadas por fórmulas na tbl_lançamento: nunca são escritas pelo site.
export const FORMULAS_LANCAMENTO = ["Estado", "Estado de Reforma", "Vencimento Base", "INSS Colaborador", "INSS Empresa", "Total INSS", "Salario Liquido", "Total de Acréscimos", "Total de Descontos", "IRPS", "Vencimento"];

export function idLancamento(ano, mes, empresa, codigo) {
  const emp = texto(empresa).toUpperCase().replace(/[\s\-/\\.]/g, "");
  return ano + "-" + String(mes).padStart(2, "0") + "-" + emp + "-" + normalizarCodigo(codigo);
}

const periodoIgual = (l, ano, mes) => n(campo(l, "Ano")) === Number(ano) && mesNumero(campo(l, "Mês", "Mes")) === Number(mes);

// Condição de pagamento vigente no período (RH_ActualizarPagamentoLinha / CriarNovoMes):
// regime e vencimento com adendas; Hora/Diário pelas Settings; limites do contrato e dias (Mensal).
export function condicoesPagamento(f, adendas, parametros, per, contrato = contratoNoPeriodo(f, adendas, per)) {
  const codigo = normalizarCodigo(campo(f, "Codigo")), empresa = texto(campo(f, "Empresa")), nome = texto(campo(f, "Nome"));
  const profissao = texto(campo(f, "Profissao", "Profissão"));
  const regime = texto(vigente(adendas, f, "RegimePagamento", per.fim, campo(f, "RegimePagamento")));
  if (!regime) return {erro: "Regime de pagamento vazio para " + empresa + " | " + codigo + " (" + nome + ")."};
  const baseCadastro = n(campo(f, "Vencimento Base"));
  const adendaBase = adendaVigente(adendas, empresa, codigo, "Vencimento Base", per.fim);
  let valorUnitario = 0, unidade = "", origem = "";
  if (igual(regime, "Hora") || igual(regime, "Diário")) {
    const t = tarifaVigente(parametros, regime, profissao, per.fim);
    if (t.erro) return {erro: t.erro + " (" + empresa + " | " + codigo + ")"};
    valorUnitario = t.valor; unidade = igual(regime, "Hora") ? "MT/hora" : "MT/dia"; origem = "Settings " + regime + " / " + profissao;
  } else if (igual(regime, "Mensal")) {
    valorUnitario = adendaBase && numero(adendaBase.valor) !== null ? n(adendaBase.valor) : baseCadastro;
    unidade = "MT/mês"; origem = adendaBase && numero(adendaBase.valor) !== null ? "Adenda de vencimento" : "Vencimento base do cadastro";
  } else return {erro: "Regime de pagamento inválido para " + empresa + " | " + codigo + ": «" + regime + "»."};
  if (!(valorUnitario > 0)) return {erro: "Valor unitário não encontrado para " + empresa + " | " + codigo + " (" + regime + " / " + (profissao || "sem profissão") + "). " +
    (igual(regime, "Mensal") ? "Verifique o vencimento base do contrato/adenda." : "Verifique as Settings: regime, categoria/profissão e vigência.")};
  const dias = igual(regime, "Mensal") ? diasSalario30(per.inicio, per.fim, contrato.inicioContrato, contrato.fimContrato) : null;
  return {profissao, regime, valorUnitario, unidade, origem, inicioContrato: contrato.inicioContrato, fimContrato: contrato.fimContrato, dias, contrato};
}

// ---------- edição de um lançamento (btnGravarAlteracaoMes_Click) ----------
export const CAMPOS_EDITAVEIS = [
  {k: "Canalizar", tipo: "simnao"}, {k: "TotalDependentes", tipo: "inteiro", rotulo: "N.º de dependentes"},
  {k: "Horas Trabalhadas", tipo: "numero", soHora: true}, {k: "Faltas Justificadas", tipo: "numero"}, {k: "Faltas Injustificadas", tipo: "numero"},
  {k: "Horas Extras", tipo: "numero"}, {k: "Subsídios", tipo: "numero", grupo: "acrescimo"}, {k: "Bónus", tipo: "numero", grupo: "acrescimo"},
  {k: "Outros Acréscimos", tipo: "numero", grupo: "acrescimo"}, {k: "Adiantamentos", tipo: "numero", grupo: "desconto"},
  {k: "Empréstimos", tipo: "numero", grupo: "desconto"}, {k: "Outros Descontos", tipo: "numero", grupo: "desconto"}, {k: "Observações", tipo: "texto"}];
const decimal = v => { const t = texto(v).replace(/\s/g, ""); if (!t) return 0; const x = Number(t.includes(",") ? t.replace(/\./g, "").replace(",", ".") : t); return Number.isFinite(x) ? x : NaN; };
export const lancamentoFinalizado = l => igual(campo(l, "Estado do Lançamento"), "Finalizado");

// Devolve {erros, linha}: linha = pedido para o script (modo "editar") com as entradas mensais
// e as condições de pagamento sincronizadas com o contrato/adendas, como o VBA faz ao gravar.
export function prepararEdicaoLancamento({original, valores, funcionario, adendas = [], parametros = [], ano, mes, utilizador = "", agora = new Date()}) {
  const erros = [];
  if (!original) return {erros: ["Lançamento não encontrado."]};
  if (lancamentoFinalizado(original)) return {erros: ["Este lançamento está Finalizado e não pode ser alterado."]};
  const empresa = texto(campo(original, "Empresa")), codigo = normalizarCodigo(campo(original, "Codigo"));
  if (!empresa || !codigo) return {erros: ["Lançamento sem Empresa/Código."]};
  if (!funcionario) return {erros: ["Colaborador não encontrado: " + empresa + " | " + codigo + "."]};
  const per = periodoFolha(Number(ano), Number(mes));
  const cp = condicoesPagamento(funcionario, adendas, parametros, per);
  if (cp.erro) return {erros: [cp.erro]};
  const hora = igual(cp.regime, "Hora");
  const linha = {Empresa: empresa, Codigo: codigo};
  for (const c of CAMPOS_EDITAVEIS) {
    const v = valores[c.k];
    if (c.tipo === "simnao") {
      if (!["Sim", "Não"].includes(texto(v))) erros.push("Seleccione Sim ou Não no campo Canalizar.");
      linha[c.k] = texto(v);
    } else if (c.tipo === "texto") linha[c.k] = texto(v);
    else if (c.soHora && !hora) linha[c.k] = "";
    else {
      const x = decimal(v);
      if (!Number.isFinite(x) || x < 0) erros.push("Valor inválido em «" + (c.rotulo || c.k) + "»: use um número igual ou superior a zero.");
      else if (c.tipo === "inteiro" && !Number.isInteger(x)) erros.push("«" + (c.rotulo || c.k) + "» tem de ser um número inteiro.");
      linha[c.k] = x;
    }
  }
  linha.RegimePagamento = cp.regime;
  linha["Valor Unitário"] = cp.valorUnitario;
  linha["Data Inicio Contrato"] = cp.inicioContrato ? Number(serie(cp.inicioContrato)) : "";
  linha["Data Fim Contrato"] = cp.fimContrato ? Number(serie(cp.fimContrato)) : "";
  if (cp.dias !== null) linha["Dias Trabalhados"] = cp.dias;
  linha["Data Ultima Alteracao"] = Math.round(((agora.getTime() - agora.getTimezoneOffset() * 60000) / 86400000 + 25569) * 100000) / 100000;
  linha["Utilizador Alteracao"] = utilizador;
  return {erros, linha, condicoes: cp};
}

// Prepara as linhas do novo mês exactamente como CriarNovoMes; não grava nada.
// Devolve {linhas, incluidos, excluidos, erros, avisos}. Com erros, o mês não pode ser criado.
export function prepararNovoMes({ano, mes, funcionarios, adendas = [], parametros = [], lancamentosMes = [], lancamentosAnterior = [], utilizador = "", agora = new Date()}) {
  ano = Number(ano); mes = Number(mes);
  const erros = [], avisos = [], incluidos = [], excluidos = [], linhas = [];
  const mesNome = NOME_MES[mes - 1];
  if (!mesNome) return {linhas, incluidos, excluidos, erros: ["Mês inválido."], avisos};
  if (lancamentosMes.some(l => periodoIgual(l, ano, mes))) erros.push("Já existem lançamentos para " + mesNome + "/" + ano + ".");

  // ValidarCodigosActivosParaLancamento
  const vistos = new Set();
  for (const f of funcionarios) {
    if (!igual(campo(f, "Estado"), "Activo")) continue;
    const cod = normalizarCodigo(campo(f, "Codigo")), emp = texto(campo(f, "Empresa"));
    if (!cod) { erros.push("Existe um colaborador activo sem Código (" + (texto(campo(f, "Nome")) || "sem nome") + ")."); continue; }
    if (!emp) { erros.push("O colaborador activo " + cod + " não tem Empresa definida."); continue; }
    const k = chaveRH(emp, cod);
    if (vistos.has(k)) erros.push("O código " + cod + " está repetido na empresa " + emp + ".");
    vistos.add(k);
  }

  // Índice do mês anterior (Empresa + Código); duplicados bloqueiam, como no VBA.
  const anterior = mes === 1 ? {ano: ano - 1, mes: 12} : {ano, mes: mes - 1};
  const indice = new Map();
  for (const l of lancamentosAnterior) {
    if (!periodoIgual(l, anterior.ano, anterior.mes)) continue;
    const k = chaveRH(campo(l, "Empresa"), campo(l, "Codigo"));
    if (indice.has(k)) erros.push("Existe mais de um lançamento para " + k.replace("|", " | ") + " em " + NOME_MES[anterior.mes - 1] + "/" + anterior.ano + ".");
    indice.set(k, l);
  }
  if (!indice.size) avisos.push("Não há lançamentos de " + NOME_MES[anterior.mes - 1] + "/" + anterior.ano + ": todos entram como «Novo Colaborador» (entradas a zero).");

  const per = periodoFolha(ano, mes);
  const agoraSerie = String(Math.round(((agora.getTime() - agora.getTimezoneOffset() * 60000) / 86400000 + 25569) * 100000) / 100000);
  for (const f of funcionarios) {
    const codigo = normalizarCodigo(campo(f, "Codigo")), empresa = texto(campo(f, "Empresa")), nome = texto(campo(f, "Nome"));
    if (!codigo || !empresa) continue;
    const id = {codigo, empresa, nome, departamento: texto(campo(f, "Departamento"))};
    const c = contratoNoPeriodo(f, adendas, per);
    if (!c.vigente) { excluidos.push({...id, motivo: c.motivo}); continue; }

    const cp = condicoesPagamento(f, adendas, parametros, per, c);
    if (cp.erro) { erros.push(cp.erro); continue; }
    const {profissao, regime, valorUnitario, unidade, origem} = cp;

    const ant = indice.get(chaveRH(empresa, codigo));
    const linha = {
      "Codigo": codigo, "Ano": String(ano), "Mês": mesNome, "Data Referencia": serie(per.referencia),
      "Nome": nome, "Departamento": texto(campo(f, "Departamento")), "Empresa": empresa, "Profissao": profissao,
      "RegimePagamento": regime, "Valor Unitário": String(valorUnitario),
      "Data Inicio Contrato": serie(c.inicioContrato), "Data Fim Contrato": serie(c.fimContrato),
      "ID_Lancamento": idLancamento(ano, mes, empresa, codigo),
      "Estado do Lançamento": "Pendente", "Origem_Lancamento": ant ? "Cópia Mês Anterior" : "Novo Colaborador", "Revisto": "Não",
      "Data Criacao": agoraSerie, "Data Ultima Alteracao": agoraSerie, "Utilizador Alteracao": utilizador
    };
    if (ant) for (const k of ENTRADAS_COPIADAS) linha[k] = texto(campo(ant, k));
    else { for (const k of ENTRADAS_ZERO) linha[k] = "0"; linha.TotalDependentes = texto(campo(f, "TotalDependentes")); linha.Canalizar = ""; linha["Observações"] = ""; }
    linha["Horas Trabalhadas"] = igual(regime, "Hora") ? (ant ? texto(campo(ant, "Horas Trabalhadas")) : "0") : "";
    if (igual(regime, "Mensal")) linha["Dias Trabalhados"] = String(diasSalario30(per.inicio, per.fim, c.inicioContrato, c.fimContrato));
    linhas.push(linha);
    incluidos.push({...id, regime, valorUnitario, unidade, origemValor: origem, origem: linha.Origem_Lancamento, dias: linha["Dias Trabalhados"] ?? "", motivo: c.motivo});
  }
  if (!linhas.length && !erros.length) erros.push("Nenhum colaborador com contrato vigente no período " + per.inicio + " a " + per.fim + ".");
  return {linhas, incluidos, excluidos, erros, avisos, periodo: per, mesNome};
}

// ---------- totais efectivos (modRH_CalculosResumo) ----------
export function totaisEfectivos(l) {
  const acrComp = n(campo(l, "Subsídios")) + n(campo(l, "Bónus")) + n(campo(l, "Outros Acréscimos"));
  const valorFalta = Object.hasOwn(l, "Valor de Falta") ? n(l["Valor de Falta"]) : 0;
  const outros = n(campo(l, "Adiantamentos")) + n(campo(l, "Empréstimos")) + n(campo(l, "Outros Descontos")) + valorFalta;
  const descComp = n(campo(l, "INSS Colaborador")) + n(campo(l, "IRPS")) + outros;
  return {acrescimos: Math.max(n(campo(l, "Total de Acréscimos")), acrComp), descontos: Math.max(n(campo(l, "Total de Descontos")), descComp), outrosDescontos: outros};
}

// ---------- folha (pré-visualização e exportação) ----------
// Filtros como no VBA: Empresa, Departamento e Colaborador (vazio ou "TODOS" = todos).
// Filtro de escalão: "" = todos; SEM_ESCALAO = linhas sem escalão; outro valor = esse escalão.
export const SEM_ESCALAO = "__sem__";
export function escalaoIgual(l, escalao) {
  if (!texto(escalao)) return true;
  const e = texto(campo(l, "Escalão"));
  return escalao === SEM_ESCALAO ? !e : igual(e, escalao);
}
// Escalões presentes num conjunto de lançamentos, por ordem natural (1.º, 2.º, …, 10.º).
export function escaloesPresentes(lancamentos) {
  const vistos = new Map();
  for (const l of lancamentos) { const e = texto(campo(l, "Escalão")); if (e && !vistos.has(e.toUpperCase())) vistos.set(e.toUpperCase(), e); }
  return [...vistos.values()].sort((a, b) => a.localeCompare(b, "pt", {numeric: true, sensitivity: "base"}));
}
export function filtrarFolha(lancamentos, ano, mes, {empresa = "", departamento = "", codigo = "", escalao = ""} = {}) {
  const todos = v => !texto(v) || igual(v, "TODOS");
  return lancamentos.filter(l => periodoIgual(l, ano, mes) &&
    (todos(empresa) || igual(campo(l, "Empresa"), empresa)) &&
    (todos(departamento) || igual(campo(l, "Departamento"), departamento)) &&
    (!texto(codigo) || normalizarCodigo(campo(l, "Codigo")) === normalizarCodigo(codigo)) &&
    escalaoIgual(l, escalao));
}
export function linhasFolha(lancamentos, ano, mes, filtros = {}) {
  return filtrarFolha(lancamentos, ano, mes, filtros)
    .map(l => {
      const t = totaisEfectivos(l), base = n(campo(l, "Vencimento Base")), sub = n(campo(l, "Subsídios"));
      const temVenc = Object.hasOwn(l, "Vencimento");
      return {
        codigo: normalizarCodigo(campo(l, "Codigo")), nome: texto(campo(l, "Nome")), profissao: texto(campo(l, "Profissao")),
        empresa: texto(campo(l, "Empresa")), departamento: texto(campo(l, "Departamento")),
        base, diasFalta: n(campo(l, "Faltas Injustificadas")), valorFalta: Object.hasOwn(l, "Valor de Falta") ? n(l["Valor de Falta"]) : null,
        vencimento: temVenc ? n(l.Vencimento) : base, subsidios: sub, outrasRem: Math.max(0, t.acrescimos - sub), totalRem: base + t.acrescimos,
        acrescimos: t.acrescimos, irps: n(campo(l, "IRPS")), inss: n(campo(l, "INSS Colaborador")), outrosDesc: t.outrosDescontos,
        totalDesc: t.descontos, liquido: n(campo(l, "Salario Liquido")), estado: texto(campo(l, "Estado do Lançamento"))
      };
    });
}
export function somar(linhas, k) { return linhas.reduce((s, l) => s + (Number(l[k]) || 0), 0); }

// ---------- relatório comparativo (mês seleccionado vs anterior, base Salário Líquido) ----------
const fmt2 = v => new Intl.NumberFormat("pt-PT", {minimumFractionDigits: 2, maximumFractionDigits: 2, useGrouping: "always"}).format(v);
export function relatorioComparativo(lancamentos, ano, mes) {
  const act = {ano: Number(ano), mes: Number(mes)}, ant = act.mes === 1 ? {ano: act.ano - 1, mes: 12} : {ano: act.ano, mes: act.mes - 1};
  const grupo = l => texto(campo(l, "Empresa")) + "|" + texto(campo(l, "Departamento"));
  const chave = l => chaveRH(campo(l, "Empresa"), campo(l, "Codigo"));
  const rowsAnt = lancamentos.filter(l => periodoIgual(l, ant.ano, ant.mes)), rowsAct = lancamentos.filter(l => periodoIgual(l, act.ano, act.mes));
  const porChaveAnt = new Map(rowsAnt.map(l => [chave(l), l])), porChaveAct = new Map(rowsAct.map(l => [chave(l), l]));
  const liq = l => n(campo(l, "Salario Liquido"));
  const grupos = new Map();
  const g = k => { if (!grupos.has(k)) grupos.set(k, {empresa: k.split("|")[0], departamento: k.split("|")[1], qtdAnterior: 0, valorAnterior: 0, qtdActual: 0, valorActual: 0, comentarios: []}); return grupos.get(k); };
  for (const l of rowsAnt) { const x = g(grupo(l)); x.qtdAnterior++; x.valorAnterior += liq(l); }
  for (const l of rowsAct) { const x = g(grupo(l)); x.qtdActual++; x.valorActual += liq(l); }
  const totalActual = rowsAct.reduce((s, l) => s + liq(l), 0);
  for (const l of rowsAct) {
    const x = g(grupo(l)), nome = texto(campo(l, "Nome")), a = porChaveAnt.get(chave(l));
    if (!a) x.comentarios.push("Entrada de " + nome + " (+" + fmt2(liq(l)) + " MT).");
    else if (!igual(campo(a, "Departamento"), campo(l, "Departamento"))) x.comentarios.push(nome + " transferido de " + texto(campo(a, "Departamento")) + " para " + texto(campo(l, "Departamento")) + ".");
    const faltas = n(campo(l, "Faltas Injustificadas"));
    if (faltas > 0) {
      const vf = Object.hasOwn(l, "Valor de Falta") ? n(l["Valor de Falta"]) : n(campo(l, "Total de Descontos"));
      x.comentarios.push(nome + ": " + String(Math.round(faltas * 100) / 100).replace(".", ",") + " falta(s) (-" + fmt2(vf) + " MT).");
    }
  }
  for (const l of rowsAnt) if (!porChaveAct.has(chave(l))) g(grupo(l)).comentarios.push("Saída de " + texto(campo(l, "Nome")) + " (-" + fmt2(liq(l)) + " MT).");
  const horas = (rows, k) => rows.filter(l => grupo(l) === k).reduce((s, l) => s + n(campo(l, "Horas Trabalhadas")), 0);
  const custo = (rows, k) => rows.filter(l => grupo(l) === k).reduce((s, l) => s + n(campo(l, "Horas Trabalhadas")) * n(campo(l, "Valor Unitário")), 0);
  const lista = [...grupos.entries()].sort(([a], [b]) => a.localeCompare(b, "pt")).map(([k, x]) => {
    const hA = horas(rowsAnt, k), hB = horas(rowsAct, k);
    if (hB !== hA) {
      const dc = custo(rowsAct, k) - custo(rowsAnt, k);
      x.comentarios.push(hB > hA ? "Aumento global de " + String(Math.round((hB - hA) * 100) / 100).replace(".", ",") + " horas (+" + fmt2(dc) + " MT)."
        : "Redução global de " + String(Math.round((hA - hB) * 100) / 100).replace(".", ",") + " horas (" + fmt2(dc) + " MT).");
    }
    return {...x, unidade: x.empresa + " - " + x.departamento, variacao: x.valorActual - x.valorAnterior, peso: totalActual ? x.valorActual / totalActual : null,
      comentario: x.comentarios.length ? x.comentarios.join(" ") : "Sem alterações relevantes identificadas."};
  });
  return {anterior: ant, actual: act, grupos: lista, totalAnterior: rowsAnt.reduce((s, l) => s + liq(l), 0), totalActual,
    qtdAnterior: rowsAnt.length, qtdActual: rowsAct.length};
}
