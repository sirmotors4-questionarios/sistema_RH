// Relatórios formais do Sistema RH: cada relatório é um "documento" estruturado (sumário, secções com texto e
// tabelas, notas), igual para a pré-visualização, o PDF e o Word. Sem DOM: testável em Node.
import * as M from "./rh-modulos.js";
import {linhasFolha, somar} from "./rh-folha.js";
import * as I from "./indemnizacoes.js";

const n = v => M.numero(v) ?? 0;
const t = v => String(v ?? "").trim();
export const mt = v => new Intl.NumberFormat("pt-PT", {minimumFractionDigits: 2, maximumFractionDigits: 2, useGrouping: "always"}).format(v || 0);
const int = v => new Intl.NumberFormat("pt-PT", {useGrouping: "always"}).format(Math.round(v || 0));
const pct = (a, b) => b ? (a / b * 100).toFixed(1).replace(".", ",") + "%" : "—";
const varPct = (a, b) => b ? ((a - b) / b * 100 >= 0 ? "+" : "") + ((a - b) / b * 100).toFixed(1).replace(".", ",") + "%" : "—";
const sinal = v => (v > 0 ? "+" : "") + mt(v);
const nomeMes = (ano, mes) => M.NOME_MES[mes - 1] + " de " + ano;
const anterior = (ano, mes) => mes === 1 ? {ano: ano - 1, mes: 12} : {ano, mes: mes - 1};
const IDADE_REFORMA = {M: 60, F: 55}; // como modRH_Dashboard.bas

export const CATALOGO = Object.freeze([
  {id: "salarial", titulo: "Relatório Mensal de Processamento Salarial", icone: "▦", precisa: ["mes", "anterior"],
   descricao: "Síntese da folha do mês: massa salarial, descontos, líquido e custo para a empresa, por empresa e departamento, com a comparação ao mês anterior."},
  {id: "encargos", titulo: "Mapa de Encargos Sociais e Fiscais", icone: "§", precisa: ["mes"],
   descricao: "INSS (colaborador e entidade empregadora) e IRPS retido no mês, por empresa e por colaborador — base para entrega e contabilização."},
  {id: "efectivo", titulo: "Quadro de Pessoal", icone: "♙", precisa: ["cadastro"],
   descricao: "Efectivo por empresa, sexo, idade, antiguidade, contrato e regime, com a lista nominal de colaboradores em idade de reforma e reformados."},
  {id: "movimentacao", titulo: "Relatório de Movimentação de Pessoal", icone: "⇄", precisa: ["cadastro", "ano"],
   descricao: "Admissões e saídas no ano, com lista nominal, motivos, saldo e taxa de rotatividade."},
  {id: "assiduidade", titulo: "Relatório de Assiduidade e Horas Extra", icone: "◷", precisa: ["mes"],
   descricao: "Faltas justificadas e injustificadas e horas extra do mês, por departamento e por colaborador."},
  {id: "evolucao", titulo: "Relatório de Custos com Pessoal — 12 meses", icone: "↗", precisa: ["doze", "indemnizacoes_opcional"],
   descricao: "Evolução mensal do efectivo processado, massa salarial, encargos, indemnizações pagas e custo total nos últimos 12 meses, com totais por empresa."},
  {id: "indemnizacoes", titulo: "Relatório de Indemnizações", icone: "◈", precisa: ["indemnizacoes"],
   descricao: "Responsabilidades com indemnizações: planos em curso, prestações do mês, valores pagos, saldo, atrasos e calendário dos próximos 12 meses, por empresa."}
]);

function cabecalho(id, extra = {}) {
  const c = CATALOGO.find(r => r.id === id);
  return {id, titulo: c.titulo, ...extra, kpis: [], secoes: [], notas: []};
}
const filtroEmpresa = (rows, empresa) => !empresa ? rows : rows.filter(l => t(M.campo(l, "Empresa")).toUpperCase() === empresa.toUpperCase());
const textoEmpresa = empresa => empresa ? "Empresa: " + empresa : "Todas as empresas";

function agrupar(lista, chaveFn, valores) {
  const m = new Map();
  for (const x of lista) {
    const k = chaveFn(x) || "(Sem informação)";
    if (!m.has(k)) m.set(k, {chave: k, n: 0, ...Object.fromEntries(Object.keys(valores).map(c => [c, 0]))});
    const g = m.get(k); g.n += 1;
    for (const [c, f] of Object.entries(valores)) g[c] += f(x);
  }
  return [...m.values()];
}

// Linha da folha + INSS empresa (que a folha não traz).
function folha(rows, ano, mes, empresa) {
  const base = filtroEmpresa(rows, empresa);
  const inssEmp = new Map(base.map(l => [M.chaveRH(M.campo(l, "Empresa"), M.campo(l, "Codigo")) + "|" + M.mes(M.campo(l, "Mês", "Mes")), n(M.campo(l, "INSS Empresa"))]));
  return linhasFolha(base, ano, mes).map(l => ({...l, inssEmp: inssEmp.get(M.chaveRH(l.empresa, l.codigo) + "|" + mes) || 0}))
    .map(l => ({...l, custo: l.vencimento + l.inssEmp}));
}

// ---------------------------------------------------------------- 1. processamento salarial
export function relatorioSalarial({lancMes, lancAnterior, ano, mes, empresa = ""}) {
  const r = cabecalho("salarial", {periodo: nomeMes(ano, mes), ambito: textoEmpresa(empresa)});
  const L = folha(lancMes, ano, mes, empresa);
  const ant = anterior(ano, mes);
  const A = folha(lancAnterior, ant.ano, ant.mes, empresa);
  if (!L.length) { r.vazio = "Não há lançamentos para " + nomeMes(ano, mes) + (empresa ? " em " + empresa : "") + "."; return r; }
  const tot = k => somar(L, k), totA = k => somar(A, k);
  r.kpis = [
    {rotulo: "Colaboradores processados", valor: int(L.length), nota: A.length ? "mês anterior: " + int(A.length) : ""},
    {rotulo: "Massa salarial (vencimento)", valor: mt(tot("vencimento")) + " MT", nota: A.length ? "variação " + varPct(tot("vencimento"), totA("vencimento")) : ""},
    {rotulo: "Total de descontos", valor: mt(tot("totalDesc")) + " MT", nota: "IRPS + INSS + outros"},
    {rotulo: "Salário líquido", valor: mt(tot("liquido")) + " MT", nota: "a transferir"},
    {rotulo: "INSS entidade empregadora", valor: mt(tot("inssEmp")) + " MT", nota: "encargo da empresa"},
    {rotulo: "Custo total para a empresa", valor: mt(tot("custo")) + " MT", nota: "vencimento + INSS empresa"}
  ];
  const dv = tot("vencimento") - totA("vencimento");
  r.secoes.push({titulo: "1. Síntese", paragrafos: [
    `Em ${nomeMes(ano, mes)} foram processados ${int(L.length)} colaboradores (${textoEmpresa(empresa).toLowerCase()}), com uma massa salarial de ${mt(tot("vencimento"))} MT e um salário líquido a pagar de ${mt(tot("liquido"))} MT.`,
    A.length ? `Face a ${nomeMes(ant.ano, ant.mes)}, a massa salarial ${dv >= 0 ? "aumentou" : "diminuiu"} ${mt(Math.abs(dv))} MT (${varPct(tot("vencimento"), totA("vencimento"))}) e o número de colaboradores processados passou de ${int(A.length)} para ${int(L.length)}.`
             : `Não há lançamentos de ${nomeMes(ant.ano, ant.mes)} para comparação.`,
    `Os descontos somam ${mt(tot("totalDesc"))} MT, dos quais IRPS ${mt(tot("irps"))} MT e INSS do colaborador ${mt(tot("inss"))} MT. O custo total para a empresa, incluindo o INSS a cargo da entidade empregadora (${mt(tot("inssEmp"))} MT), é de ${mt(tot("custo"))} MT.`
  ]});
  const colsValores = [{t: "Colab.", num: true}, {t: "Vencimento", num: true}, {t: "Acréscimos", num: true}, {t: "Descontos", num: true}, {t: "Líquido", num: true}, {t: "Custo empresa", num: true}];
  const linhaGrupo = g => [g.chave, int(g.n), mt(g.venc), mt(g.acr), mt(g.desc), mt(g.liq), mt(g.custo)];
  const vals = {venc: l => l.vencimento, acr: l => l.acrescimos, desc: l => l.totalDesc, liq: l => l.liquido, custo: l => l.custo};
  const totalLinha = ["Total", int(L.length), mt(tot("vencimento")), mt(tot("acrescimos")), mt(tot("totalDesc")), mt(tot("liquido")), mt(tot("custo"))];
  if (!empresa) r.secoes.push({titulo: "2. Por empresa", tabela: {cols: [{t: "Empresa"}, ...colsValores],
    linhas: agrupar(L, l => l.empresa, vals).sort((a, b) => b.custo - a.custo).map(linhaGrupo), total: totalLinha}});
  r.secoes.push({titulo: (empresa ? "2" : "3") + ". Por departamento", tabela: {cols: [{t: "Departamento"}, ...colsValores],
    linhas: agrupar(L, l => l.departamento, vals).sort((a, b) => b.custo - a.custo).map(linhaGrupo), total: totalLinha}});

  // Variações individuais face ao mês anterior
  if (A.length) {
    const mapaA = new Map(A.map(l => [M.chaveRH(l.empresa, l.codigo), l]));
    const mapaL = new Map(L.map(l => [M.chaveRH(l.empresa, l.codigo), l]));
    const novos = L.filter(l => !mapaA.has(M.chaveRH(l.empresa, l.codigo)));
    const sairam = A.filter(l => !mapaL.has(M.chaveRH(l.empresa, l.codigo)));
    const variacoes = L.filter(l => mapaA.has(M.chaveRH(l.empresa, l.codigo)))
      .map(l => ({l, d: l.liquido - mapaA.get(M.chaveRH(l.empresa, l.codigo)).liquido})).filter(x => Math.abs(x.d) >= 0.01)
      .sort((a, b) => Math.abs(b.d) - Math.abs(a.d)).slice(0, 15);
    const sec = empresa ? 3 : 4;
    r.secoes.push({titulo: `${sec}. Alterações face a ${nomeMes(ant.ano, ant.mes)}`, paragrafos: [
      `${int(novos.length)} colaborador(es) processado(s) pela primeira vez e ${int(sairam.length)} que deixaram de ser processados. ` +
      (variacoes.length ? `As ${int(variacoes.length)} maiores variações de salário líquido estão abaixo.` : "Não há variações de salário líquido entre os colaboradores comuns aos dois meses.")],
      tabela: variacoes.length ? {cols: [{t: "Código"}, {t: "Nome"}, {t: "Departamento"}, {t: "Líquido anterior", num: true}, {t: "Líquido actual", num: true}, {t: "Variação", num: true}],
        linhas: variacoes.map(({l, d}) => [l.codigo, l.nome, l.departamento, mt(l.liquido - d), mt(l.liquido), sinal(d)])} : null});
    if (novos.length || sairam.length) r.secoes.push({titulo: `${sec + 1}. Entradas e saídas na folha`, tabela: {
      cols: [{t: "Situação"}, {t: "Código"}, {t: "Nome"}, {t: "Empresa"}, {t: "Departamento"}, {t: "Líquido", num: true}],
      linhas: [...novos.map(l => ["Novo na folha", l.codigo, l.nome, l.empresa, l.departamento, mt(l.liquido)]),
               ...sairam.map(l => ["Saiu da folha", l.codigo, l.nome, l.empresa, l.departamento, mt(l.liquido)])]}});
  }
  // Estado do processamento
  const estados = agrupar(L, l => l.estado || "Pendente", {});
  const pendentes = L.filter(l => !/finaliz/i.test(l.estado)).length;
  r.secoes.push({titulo: `${r.secoes.length + 1}. Estado do processamento`, paragrafos: [pendentes
    ? `${int(pendentes)} lançamento(s) ainda não finalizado(s). A folha só deve ser paga depois de todos os lançamentos estarem revistos e finalizados.`
    : "Todos os lançamentos do mês estão finalizados."],
    tabela: {cols: [{t: "Estado do lançamento"}, {t: "Lançamentos", num: true}, {t: "% do total", num: true}],
      linhas: estados.sort((a, b) => b.n - a.n).map(g => [g.chave, int(g.n), pct(g.n, L.length)])}});
  r.notas.push("Valores em meticais (MT). Vencimento, INSS, IRPS e líquido são os calculados pelo sistema.");
  return r;
}

// ---------------------------------------------------------------- 2. encargos sociais e fiscais
export function relatorioEncargos({lancMes, ano, mes, empresa = ""}) {
  const r = cabecalho("encargos", {periodo: nomeMes(ano, mes), ambito: textoEmpresa(empresa)});
  const L = folha(lancMes, ano, mes, empresa);
  if (!L.length) { r.vazio = "Não há lançamentos para " + nomeMes(ano, mes) + "."; return r; }
  const tot = k => somar(L, k);
  r.kpis = [
    {rotulo: "Base de incidência", valor: mt(tot("vencimento")) + " MT", nota: int(L.length) + " colaboradores"},
    {rotulo: "INSS colaborador", valor: mt(tot("inss")) + " MT", nota: pct(tot("inss"), tot("vencimento")) + " da base"},
    {rotulo: "INSS entidade empregadora", valor: mt(tot("inssEmp")) + " MT", nota: pct(tot("inssEmp"), tot("vencimento")) + " da base"},
    {rotulo: "Total INSS a entregar", valor: mt(tot("inss") + tot("inssEmp")) + " MT", nota: "colaborador + empresa"},
    {rotulo: "IRPS retido", valor: mt(tot("irps")) + " MT", nota: "taxa efectiva " + pct(tot("irps"), tot("vencimento"))},
    {rotulo: "Total de encargos", valor: mt(tot("inss") + tot("inssEmp") + tot("irps")) + " MT", nota: "INSS + IRPS"}
  ];
  r.secoes.push({titulo: "1. Resumo", paragrafos: [
    `Em ${nomeMes(ano, mes)}, sobre uma base de incidência de ${mt(tot("vencimento"))} MT, há ${mt(tot("inss") + tot("inssEmp"))} MT de INSS a entregar (${mt(tot("inss"))} MT retidos aos colaboradores e ${mt(tot("inssEmp"))} MT a cargo da entidade empregadora) e ${mt(tot("irps"))} MT de IRPS retido na fonte.`]});
  const vals = {base: l => l.vencimento, ic: l => l.inss, ie: l => l.inssEmp, irps: l => l.irps};
  r.secoes.push({titulo: "2. Por empresa", tabela: {
    cols: [{t: "Empresa"}, {t: "Colab.", num: true}, {t: "Base", num: true}, {t: "INSS colab.", num: true}, {t: "INSS empresa", num: true}, {t: "Total INSS", num: true}, {t: "IRPS", num: true}],
    linhas: agrupar(L, l => l.empresa, vals).sort((a, b) => b.base - a.base).map(g => [g.chave, int(g.n), mt(g.base), mt(g.ic), mt(g.ie), mt(g.ic + g.ie), mt(g.irps)]),
    total: ["Total", int(L.length), mt(tot("vencimento")), mt(tot("inss")), mt(tot("inssEmp")), mt(tot("inss") + tot("inssEmp")), mt(tot("irps"))]}});
  r.secoes.push({titulo: "3. Mapa nominal", tabela: {
    cols: [{t: "Código"}, {t: "Nome"}, {t: "Empresa"}, {t: "Base", num: true}, {t: "INSS colab.", num: true}, {t: "INSS empresa", num: true}, {t: "IRPS", num: true}],
    linhas: [...L].sort((a, b) => a.empresa.localeCompare(b.empresa) || a.codigo.localeCompare(b.codigo))
      .map(l => [l.codigo, l.nome, l.empresa, mt(l.vencimento), mt(l.inss), mt(l.inssEmp), mt(l.irps)]),
    total: ["", "Total", "", mt(tot("vencimento")), mt(tot("inss")), mt(tot("inssEmp")), mt(tot("irps"))]}});
  r.notas.push("Os valores de INSS e IRPS são os calculados pelo sistema; confirme as taxas em vigor antes da entrega.");
  return r;
}

// ---------------------------------------------------------------- dados do cadastro
function hojeISO(hoje) { return (hoje || new Date()).toISOString().slice(0, 10); }
function idade(iso, hoje) {
  if (!iso) return null;
  const [a, m, d] = iso.split("-").map(Number), [ha, hm, hd] = hojeISO(hoje).split("-").map(Number);
  return ha - a - ((hm < m || (hm === m && hd < d)) ? 1 : 0);
}
function antiguidade(iso, hoje) {
  if (!iso || iso > hojeISO(hoje)) return null;
  const [a, m, d] = iso.split("-").map(Number), [ha, hm, hd] = hojeISO(hoje).split("-").map(Number);
  return ((ha - a) * 12 + hm - m - (hd < d ? 1 : 0)) / 12;
}
const sexoCurto = v => { const s = t(v).toUpperCase(); return s === "M" || s === "MASCULINO" ? "M" : s === "F" || s === "FEMININO" ? "F" : ""; };
const reformadoV = v => ["SIM", "REFORMADO", "REFORMADA"].includes(t(v).toUpperCase());

export function pessoas(funcionarios, hoje) {
  return funcionarios.map(f => {
    const nasc = M.data(M.campo(f, "Data Nascimento", "Data de Nascimento", "DataNascimento"));
    const adm = M.data(M.campo(f, "Data Admissao", "Data de Admissão", "Data Admissão"));
    const s = sexoCurto(M.campo(f, "Sexo"));
    const id = idade(nasc, hoje);
    return {
      codigo: M.normalizarCodigo(M.campo(f, "Codigo")), nome: t(M.campo(f, "Nome")), empresa: t(M.campo(f, "Empresa")),
      departamento: t(M.campo(f, "Departamento")), profissao: t(M.campo(f, "Profissao")), sexo: s,
      activo: M.activo(M.campo(f, "Estado")), reformado: reformadoV(M.campo(f, "Estado de Reforma")),
      idade: id, antiguidade: antiguidade(adm, hoje), admissao: adm,
      saida: M.data(M.campo(f, "Data Saida", "Data de Saída", "Data Saída")),
      motivo: t(M.campo(f, "Motivo de saída", "Motivo de Saída", "Motivo Saida")),
      contrato: t(M.campo(f, "TipoContrato", "Tipo de Contrato")), regime: t(M.campo(f, "RegimePagamento")),
      emIdadeReforma: id != null && IDADE_REFORMA[s] != null && id >= IDADE_REFORMA[s]
    };
  });
}

function tabelaContagem(titulo, lista, chaveFn, total, ordem) {
  const g = agrupar(lista, chaveFn, {});
  if (ordem) g.sort((a, b) => ordem.indexOf(a.chave) - ordem.indexOf(b.chave)); else g.sort((a, b) => b.n - a.n);
  return {cols: [{t: titulo}, {t: "Colaboradores", num: true}, {t: "%", num: true}], linhas: g.map(x => [x.chave, int(x.n), pct(x.n, total)]),
    total: ["Total", int(total), "100,0%"]};
}

// ---------------------------------------------------------------- 3. quadro de pessoal
export function relatorioEfectivo({funcionarios, empresa = "", hoje}) {
  const data = (hoje || new Date()).toLocaleDateString("pt-PT");
  const r = cabecalho("efectivo", {periodo: "Situação em " + data, ambito: textoEmpresa(empresa)});
  const P = pessoas(filtroEmpresa(funcionarios, empresa), hoje);
  const A = P.filter(p => p.activo);
  if (!A.length) { r.vazio = "Não há colaboradores activos no cadastro" + (empresa ? " de " + empresa : "") + "."; return r; }
  const h = A.filter(p => p.sexo === "M").length, m = A.filter(p => p.sexo === "F").length;
  const comIdade = A.filter(p => p.idade != null), comAnt = A.filter(p => p.antiguidade != null);
  const idadeMedia = comIdade.length ? comIdade.reduce((s, p) => s + p.idade, 0) / comIdade.length : null;
  const antMedia = comAnt.length ? comAnt.reduce((s, p) => s + p.antiguidade, 0) / comAnt.length : null;
  const emIdade = A.filter(p => p.emIdadeReforma), reformados = A.filter(p => p.reformado);
  const um = v => v == null ? "—" : v.toFixed(1).replace(".", ",");
  const anos = v => v == null ? "—" : um(v) + " anos";
  r.kpis = [
    {rotulo: "Colaboradores activos", valor: int(A.length), nota: int(P.length - A.length) + " inactivos no cadastro"},
    {rotulo: "Homens / Mulheres", valor: int(h) + " / " + int(m), nota: pct(m, A.length) + " de mulheres"},
    {rotulo: "Idade média", valor: anos(idadeMedia), nota: int(comIdade.length) + " com data de nascimento"},
    {rotulo: "Antiguidade média", valor: anos(antMedia), nota: "desde a admissão"},
    {rotulo: "Em idade de reforma", valor: int(emIdade.length), nota: "homens ≥ 60, mulheres ≥ 55"},
    {rotulo: "Reformados no activo", valor: int(reformados.length), nota: "«Estado de Reforma» = Sim"}
  ];
  r.secoes.push({titulo: "1. Síntese", paragrafos: [
    `À data de ${data}, o efectivo activo (${textoEmpresa(empresa).toLowerCase()}) é de ${int(A.length)} colaboradores: ${int(h)} homens e ${int(m)} mulheres` +
    (A.length - h - m ? ` (${int(A.length - h - m)} sem sexo indicado)` : "") + `. ` + (idadeMedia == null ? "As datas de nascimento não estão registadas" : `A idade média é de ${um(idadeMedia)} anos`) +
    (antMedia == null ? "." : ` e a antiguidade média de ${um(antMedia)} anos.`),
    `${int(emIdade.length)} colaborador(es) activo(s) atingiram a idade legal de reforma e ${int(reformados.length)} estão registados como reformados. A lista nominal está na secção 6.`]});
  if (!empresa) {
    const g = agrupar(A, p => p.empresa, {h: p => p.sexo === "M" ? 1 : 0, m: p => p.sexo === "F" ? 1 : 0});
    r.secoes.push({titulo: "2. Efectivo por empresa e sexo", tabela: {cols: [{t: "Empresa"}, {t: "Homens", num: true}, {t: "Mulheres", num: true}, {t: "Total", num: true}, {t: "%", num: true}],
      linhas: g.sort((a, b) => b.n - a.n).map(x => [x.chave, int(x.h), int(x.m), int(x.n), pct(x.n, A.length)]), total: ["Total", int(h), int(m), int(A.length), "100,0%"]}});
  }
  r.secoes.push({titulo: (empresa ? "2" : "3") + ". Por departamento", tabela: tabelaContagem("Departamento", A, p => p.departamento, A.length)});
  const faixa = i => i == null ? "(Sem data)" : i < 25 ? "< 25" : i < 35 ? "25 - 34" : i < 45 ? "35 - 44" : i < 55 ? "45 - 54" : i < 60 ? "55 - 59" : "60+";
  const faixaAnt = a => a == null ? "(Sem data)" : a < 1 ? "< 1 ano" : a < 4 ? "1 - 3 anos" : a < 6 ? "4 - 5 anos" : a < 11 ? "6 - 10 anos" : "> 10 anos";
  r.secoes.push({titulo: (empresa ? "3" : "4") + ". Estrutura etária e antiguidade", subtabelas: [
    tabelaContagem("Faixa etária", A, p => faixa(p.idade), A.length, ["< 25", "25 - 34", "35 - 44", "45 - 54", "55 - 59", "60+", "(Sem data)"]),
    tabelaContagem("Antiguidade", A, p => faixaAnt(p.antiguidade), A.length, ["< 1 ano", "1 - 3 anos", "4 - 5 anos", "6 - 10 anos", "> 10 anos", "(Sem data)"])]});
  r.secoes.push({titulo: (empresa ? "4" : "5") + ". Vínculo e regime", subtabelas: [
    tabelaContagem("Tipo de contrato", A, p => p.contrato, A.length), tabelaContagem("Regime de pagamento", A, p => p.regime, A.length)]});
  const lista = [...emIdade.map(p => ({...p, sit: p.reformado ? "Reformado" : "Em idade de reforma"})),
                 ...reformados.filter(p => !p.emIdadeReforma).map(p => ({...p, sit: "Reformado"}))]
    .sort((a, b) => (b.idade ?? 0) - (a.idade ?? 0));
  r.secoes.push({titulo: (empresa ? "5" : "6") + ". Colaboradores em idade de reforma e reformados",
    paragrafos: lista.length ? [] : ["Não há colaboradores activos em idade de reforma nem reformados."],
    tabela: lista.length ? {cols: [{t: "Código"}, {t: "Nome"}, {t: "Empresa"}, {t: "Departamento"}, {t: "Sexo"}, {t: "Idade", num: true}, {t: "Situação"}],
      linhas: lista.map(p => [p.codigo, p.nome, p.empresa, p.departamento, p.sexo || "—", p.idade ?? "—", p.sit])} : null});
  r.notas.push("Idade de reforma considerada: 60 anos (homens) e 55 anos (mulheres).");
  return r;
}

// ---------------------------------------------------------------- 4. movimentação
export function relatorioMovimentacao({funcionarios, ano, empresa = "", efectivoMedio = 0, hoje}) {
  const r = cabecalho("movimentacao", {periodo: "Ano de " + ano, ambito: textoEmpresa(empresa)});
  const P = pessoas(filtroEmpresa(funcionarios, empresa), hoje);
  const ent = P.filter(p => p.admissao.startsWith(String(ano))).sort((a, b) => a.admissao.localeCompare(b.admissao));
  const sai = P.filter(p => p.saida.startsWith(String(ano))).sort((a, b) => a.saida.localeCompare(b.saida));
  const activos = P.filter(p => p.activo).length;
  const media = efectivoMedio || activos;
  r.kpis = [
    {rotulo: "Admissões", valor: int(ent.length), nota: "em " + ano},
    {rotulo: "Saídas", valor: int(sai.length), nota: "em " + ano},
    {rotulo: "Saldo", valor: (ent.length - sai.length > 0 ? "+" : "") + int(ent.length - sai.length), nota: "admissões − saídas"},
    {rotulo: "Taxa de rotatividade", valor: pct(sai.length, media), nota: "saídas / efectivo médio (" + int(media) + ")"},
    {rotulo: "Efectivo activo actual", valor: int(activos), nota: "cadastro"},
    {rotulo: "Taxa de entrada", valor: pct(ent.length, media), nota: "admissões / efectivo médio"}
  ];
  r.secoes.push({titulo: "1. Síntese", paragrafos: [
    `Em ${ano} registaram-se ${int(ent.length)} admissões e ${int(sai.length)} saídas (${textoEmpresa(empresa).toLowerCase()}), um saldo de ${ent.length - sai.length >= 0 ? "+" : ""}${int(ent.length - sai.length)} colaboradores. ` +
    `A taxa de rotatividade, face a um efectivo médio de ${int(media)}, é de ${pct(sai.length, media)}.`]});
  const porMes = Array.from({length: 12}, (_, i) => {
    const mm = String(i + 1).padStart(2, "0");
    const e = ent.filter(p => p.admissao.slice(5, 7) === mm).length, s = sai.filter(p => p.saida.slice(5, 7) === mm).length;
    return [M.NOME_MES[i], int(e), int(s), (e - s > 0 ? "+" : "") + int(e - s)];
  });
  r.secoes.push({titulo: "2. Por mês", tabela: {cols: [{t: "Mês"}, {t: "Admissões", num: true}, {t: "Saídas", num: true}, {t: "Saldo", num: true}], linhas: porMes,
    total: ["Total", int(ent.length), int(sai.length), (ent.length - sai.length > 0 ? "+" : "") + int(ent.length - sai.length)]}});
  const fd = d => d ? d.slice(8, 10) + "/" + d.slice(5, 7) + "/" + d.slice(0, 4) : "";
  r.secoes.push({titulo: "3. Admissões", paragrafos: ent.length ? [] : ["Sem admissões no ano."],
    tabela: ent.length ? {cols: [{t: "Data"}, {t: "Código"}, {t: "Nome"}, {t: "Empresa"}, {t: "Departamento"}, {t: "Profissão"}, {t: "Contrato"}],
      linhas: ent.map(p => [fd(p.admissao), p.codigo, p.nome, p.empresa, p.departamento, p.profissao, p.contrato])} : null});
  r.secoes.push({titulo: "4. Saídas", paragrafos: sai.length ? [] : ["Sem saídas no ano."],
    tabela: sai.length ? {cols: [{t: "Data"}, {t: "Código"}, {t: "Nome"}, {t: "Empresa"}, {t: "Departamento"}, {t: "Antiguidade", num: true}, {t: "Motivo"}],
      linhas: sai.map(p => [fd(p.saida), p.codigo, p.nome, p.empresa, p.departamento,
        p.admissao ? (((new Date(p.saida) - new Date(p.admissao)) / 31557600000).toFixed(1).replace(".", ",") + " anos") : "—", p.motivo || "(Sem informação)"])} : null});
  if (sai.length) r.secoes.push({titulo: "5. Motivos de saída", tabela: tabelaContagem("Motivo", sai, p => p.motivo, sai.length)});
  r.notas.push("Admissões e saídas segundo as datas de admissão e de saída registadas no cadastro.");
  return r;
}

// ---------------------------------------------------------------- 5. assiduidade
export function relatorioAssiduidade({lancMes, ano, mes, empresa = ""}) {
  const r = cabecalho("assiduidade", {periodo: nomeMes(ano, mes), ambito: textoEmpresa(empresa)});
  const rows = filtroEmpresa(lancMes, empresa).filter(l => M.numero(M.campo(l, "Ano")) === ano && M.mes(M.campo(l, "Mês", "Mes")) === mes);
  if (!rows.length) { r.vazio = "Não há lançamentos para " + nomeMes(ano, mes) + "."; return r; }
  const X = rows.map(l => ({codigo: M.normalizarCodigo(M.campo(l, "Codigo")), nome: t(M.campo(l, "Nome")), empresa: t(M.campo(l, "Empresa")),
    departamento: t(M.campo(l, "Departamento")), fj: n(M.campo(l, "Faltas Justificadas")), fi: n(M.campo(l, "Faltas Injustificadas")),
    he: n(M.campo(l, "Horas Extras")), dias: n(M.campo(l, "Dias Trabalhados"))}));
  const s = k => X.reduce((a, x) => a + x[k], 0);
  const comFaltas = X.filter(x => x.fj + x.fi > 0), comHE = X.filter(x => x.he > 0);
  const diasPrevistos = X.length * 22;
  r.kpis = [
    {rotulo: "Colaboradores", valor: int(X.length), nota: "lançamentos do mês"},
    {rotulo: "Faltas justificadas", valor: int(s("fj")) + " dias", nota: int(X.filter(x => x.fj > 0).length) + " colaboradores"},
    {rotulo: "Faltas injustificadas", valor: int(s("fi")) + " dias", nota: int(X.filter(x => x.fi > 0).length) + " colaboradores"},
    {rotulo: "Taxa de absentismo", valor: pct(s("fj") + s("fi"), diasPrevistos), nota: "faltas / (colab. × 22 dias úteis)"},
    {rotulo: "Horas extra", valor: int(s("he")) + " h", nota: int(comHE.length) + " colaboradores"},
    {rotulo: "Sem faltas", valor: int(X.length - comFaltas.length), nota: pct(X.length - comFaltas.length, X.length) + " do efectivo"}
  ];
  r.secoes.push({titulo: "1. Síntese", paragrafos: [
    `Em ${nomeMes(ano, mes)} registaram-se ${int(s("fj"))} dias de faltas justificadas e ${int(s("fi"))} de faltas injustificadas, distribuídos por ${int(comFaltas.length)} colaboradores ` +
    `(taxa de absentismo estimada de ${pct(s("fj") + s("fi"), diasPrevistos)}). Foram registadas ${int(s("he"))} horas extra por ${int(comHE.length)} colaboradores.`]});
  const g = agrupar(X, x => x.departamento, {fj: x => x.fj, fi: x => x.fi, he: x => x.he});
  r.secoes.push({titulo: "2. Por departamento", tabela: {cols: [{t: "Departamento"}, {t: "Colab.", num: true}, {t: "Faltas just.", num: true}, {t: "Faltas injust.", num: true}, {t: "Horas extra", num: true}],
    linhas: g.sort((a, b) => (b.fi + b.fj) - (a.fi + a.fj) || b.he - a.he).map(x => [x.chave, int(x.n), int(x.fj), int(x.fi), int(x.he)]),
    total: ["Total", int(X.length), int(s("fj")), int(s("fi")), int(s("he"))]}});
  r.secoes.push({titulo: "3. Colaboradores com faltas", paragrafos: comFaltas.length ? [] : ["Sem faltas registadas no mês."],
    tabela: comFaltas.length ? {cols: [{t: "Código"}, {t: "Nome"}, {t: "Departamento"}, {t: "Faltas just.", num: true}, {t: "Faltas injust.", num: true}],
      linhas: comFaltas.sort((a, b) => b.fi - a.fi || b.fj - a.fj).map(x => [x.codigo, x.nome, x.departamento, int(x.fj), int(x.fi)])} : null});
  r.secoes.push({titulo: "4. Colaboradores com horas extra", paragrafos: comHE.length ? [] : ["Sem horas extra registadas no mês."],
    tabela: comHE.length ? {cols: [{t: "Código"}, {t: "Nome"}, {t: "Departamento"}, {t: "Horas extra", num: true}],
      linhas: comHE.sort((a, b) => b.he - a.he).map(x => [x.codigo, x.nome, x.departamento, int(x.he)])} : null});
  r.notas.push("Taxa de absentismo estimada com 22 dias úteis por colaborador; as faltas e horas extra são as registadas nos lançamentos do mês.");
  return r;
}

// ---------------------------------------------------------------- 6. evolução 12 meses
export function relatorioEvolucao({lancPorMes, empresa = "", indemnizacoes = null, prestacoes = []}) {
  // lancPorMes: [{ano, mes, linhas}] do mais antigo para o mais recente.
  // Indemnizações pagas entram no custo do mês do pagamento (Data Pagamento; sem ela, a data limite).
  const comInd = Array.isArray(indemnizacoes);
  const planos = comInd ? I.indemnizacoes(indemnizacoes).filter(i => !/anulad/i.test(i.estado) && (!empresa || i.empresa.toUpperCase() === empresa.toUpperCase())) : [];
  const porId = new Map(planos.map(i => [i.id, i]));
  const pagas = comInd ? I.prestacoes(prestacoes).filter(p => I.paga(p) && porId.has(p.indemnizacao)) : [];
  const indMes = (ano, mes) => { const k = ano + "-" + String(mes).padStart(2, "0"); return pagas.filter(p => (p.dataPagamento || p.dataLimite || "").slice(0, 7) === k); };
  const meses = lancPorMes.map(({ano, mes, linhas}) => {
    const L = folha(linhas, ano, mes, empresa), P = indMes(ano, mes);
    const salarial = somar(L, "custo"), ind = P.reduce((s, p) => s + p.valor, 0);
    return {ano, mes, L, P, n: L.length, venc: somar(L, "vencimento"), inssEmp: somar(L, "inssEmp"), salarial, ind, custo: salarial + ind,
      liq: somar(L, "liquido"), irps: somar(L, "irps"), inss: somar(L, "inss")};
  });
  const com = meses.filter(m => m.n || m.ind);
  const prim = lancPorMes[0], ult = lancPorMes[lancPorMes.length - 1];
  const r = cabecalho("evolucao", {periodo: prim ? `${M.NOME_MES[prim.mes - 1]} ${prim.ano} a ${M.NOME_MES[ult.mes - 1]} ${ult.ano}` : "", ambito: textoEmpresa(empresa)});
  if (!com.length) { r.vazio = "Não há lançamentos nos últimos 12 meses."; return r; }
  const tot = k => com.reduce((s, m) => s + m[k], 0);
  const comL = com.filter(m => m.n), a = comL[0] || com[0], b = comL[comL.length - 1] || com[com.length - 1];
  r.kpis = [
    {rotulo: "Custo total no período", valor: mt(tot("custo")) + " MT", nota: tot("ind") ? "inclui " + mt(tot("ind")) + " MT de indemnizações" : int(com.length) + " meses com lançamentos"},
    {rotulo: "Custo médio mensal", valor: mt(tot("custo") / com.length) + " MT", nota: "salarial + indemnizações"},
    {rotulo: "Massa salarial no período", valor: mt(tot("venc")) + " MT", nota: "vencimento"},
    {rotulo: "Encargos (INSS + IRPS)", valor: mt(tot("inss") + tot("inssEmp") + tot("irps")) + " MT", nota: "no período"},
    {rotulo: "Indemnizações pagas", valor: mt(tot("ind")) + " MT", nota: comInd ? pct(tot("ind"), tot("custo")) + " do custo total" : "não lidas (página Indemnizações por ligar)"},
    {rotulo: "Variação do custo", valor: varPct(b.custo, a.custo), nota: `${M.NOME_MES[b.mes - 1]} face a ${M.NOME_MES[a.mes - 1]}`}
  ];
  r.secoes.push({titulo: "1. Síntese", paragrafos: [
    `No período de ${r.periodo}, o custo com pessoal (${textoEmpresa(empresa).toLowerCase()}) totalizou ${mt(tot("custo"))} MT, uma média de ${mt(tot("custo") / com.length)} MT por mês` +
      (tot("ind") ? `, dos quais ${mt(tot("salarial"))} MT de custo salarial (vencimento + INSS da empresa) e ${mt(tot("ind"))} MT de indemnizações pagas (${pct(tot("ind"), tot("custo"))}).` : "."),
    `Entre ${M.NOME_MES[a.mes - 1]} ${a.ano} e ${M.NOME_MES[b.mes - 1]} ${b.ano}, o custo mensal passou de ${mt(a.custo)} MT para ${mt(b.custo)} MT (${varPct(b.custo, a.custo)}) e o efectivo processado de ${int(a.n)} para ${int(b.n)} colaboradores.`]
    .concat(comInd ? [] : ["As indemnizações não foram incluídas porque a página Indemnizações ainda não está ligada neste navegador."])});
  let antC = null;
  r.secoes.push({titulo: "2. Evolução mensal", tabela: {
    cols: [{t: "Mês"}, {t: "Colab.", num: true}, {t: "Vencimento", num: true}, {t: "INSS empresa", num: true}, {t: "Custo salarial", num: true}, {t: "Indemnizações", num: true}, {t: "Custo total", num: true}, {t: "Var. custo", num: true}],
    linhas: meses.map(m => { const tem = m.n || m.ind; const lin = [M.NOME_MES[m.mes - 1] + " " + m.ano, m.n ? int(m.n) : "—", m.n ? mt(m.venc) : "—", m.n ? mt(m.inssEmp) : "—", m.n ? mt(m.salarial) : "—", m.ind ? mt(m.ind) : "—", tem ? mt(m.custo) : "—", tem && antC ? varPct(m.custo, antC) : "—"]; if (tem) antC = m.custo; return lin; }),
    total: ["Total", "", mt(tot("venc")), mt(tot("inssEmp")), mt(tot("salarial")), mt(tot("ind")), mt(tot("custo")), ""]}});
  if (!empresa) {
    const g = agrupar(com.flatMap(m => m.L), l => l.empresa, {venc: l => l.vencimento, ie: l => l.inssEmp, custo: l => l.custo});
    const indEmp = new Map();
    for (const p of com.flatMap(m => m.P)) { const e = porId.get(p.indemnizacao)?.empresa || "(Sem informação)"; indEmp.set(e, (indEmp.get(e) || 0) + p.valor); }
    for (const [e] of indEmp) if (!g.some(x => x.chave === e)) g.push({chave: e, n: 0, venc: 0, ie: 0, custo: 0});
    r.secoes.push({titulo: "3. Por empresa (período)", tabela: {cols: [{t: "Empresa"}, {t: "Vencimento", num: true}, {t: "INSS empresa", num: true}, {t: "Custo salarial", num: true}, {t: "Indemnizações", num: true}, {t: "Custo total", num: true}, {t: "% do custo", num: true}],
      linhas: g.map(x => ({...x, ind: indEmp.get(x.chave) || 0})).sort((x, y) => (y.custo + y.ind) - (x.custo + x.ind))
        .map(x => [x.chave, mt(x.venc), mt(x.ie), mt(x.custo), x.ind ? mt(x.ind) : "—", mt(x.custo + x.ind), pct(x.custo + x.ind, tot("custo"))]),
      total: ["Total", mt(tot("venc")), mt(tot("inssEmp")), mt(tot("salarial")), mt(tot("ind")), mt(tot("custo")), "100,0%"]}});
  }
  r.notas.push("Custo total = vencimento + INSS a cargo da empresa + indemnizações pagas no mês (pela data de pagamento).");
  r.notas.push("Meses sem lançamentos nem pagamentos aparecem com «—» e não entram nas médias.");
  return r;
}

// ---------- Indemnizações ----------
// indemnizacoes/prestacoes: linhas de tbl_indemnizacoes e tbl_indemnizacoes_prestacoes. Planos anulados ficam de fora.
const dataPt = iso => iso ? iso.slice(8, 10) + "/" + iso.slice(5, 7) + "/" + iso.slice(0, 4) : "—";
export function relatorioIndemnizacoes({indemnizacoes = [], prestacoes = [], ano, mes, empresa = "", hoje}) {
  const hj = hojeISO(hoje);
  const todos = I.indemnizacoes(indemnizacoes).filter(i => !/anulad/i.test(i.estado));
  const lista = empresa ? todos.filter(i => i.empresa.toUpperCase() === empresa.toUpperCase()) : todos;
  const ids = new Set(lista.map(i => i.id));
  const pre = I.prestacoes(prestacoes).filter(p => ids.has(p.indemnizacao));
  const r = cabecalho("indemnizacoes", {periodo: nomeMes(ano, mes), ambito: textoEmpresa(empresa)});
  if (!lista.length) { r.vazio = "Não há indemnizações registadas" + (empresa ? " para " + empresa : "") + "."; return r; }
  const soma = arr => arr.reduce((s, p) => s + p.valor, 0);
  const chaveMes = ano + "-" + String(mes).padStart(2, "0");
  const doMes = pre.filter(p => (p.dataLimite || "").slice(0, 7) === chaveMes).sort((a, b) => a.dataLimite.localeCompare(b.dataLimite));
  const atraso = pre.filter(p => I.emAtraso(p, hj)).sort((a, b) => a.dataLimite.localeCompare(b.dataLimite));
  const total = lista.reduce((s, i) => s + i.total, 0), pago = soma(pre.filter(I.paga));
  const resumos = lista.map(i => ({i, r: I.resumo(i, pre)}));
  const emCurso = resumos.filter(x => !x.r.concluida);
  r.kpis = [
    {rotulo: "Total comprometido", valor: mt(total) + " MT", nota: int(lista.length) + " plano(s) · " + int(emCurso.length) + " em curso"},
    {rotulo: "Valor pago", valor: mt(pago) + " MT", nota: pct(pago, total) + " do total"},
    {rotulo: "Saldo por pagar", valor: mt(total - pago) + " MT", nota: "responsabilidade em aberto"},
    {rotulo: "Previsto no mês", valor: mt(soma(doMes)) + " MT", nota: int(doMes.length) + " prestação(ões) em " + M.NOME_MES[mes - 1].toLowerCase()},
    {rotulo: "Pago no mês", valor: mt(soma(doMes.filter(I.paga))) + " MT", nota: "das prestações do mês"},
    {rotulo: "Em atraso", valor: mt(soma(atraso)) + " MT", nota: int(atraso.length) + " prestação(ões) vencida(s)"}
  ];
  const sec = [];
  sec.push({titulo: "Síntese", paragrafos: [
    `A ${dataPt(hj)}, existem ${int(lista.length)} plano(s) de indemnização (${textoEmpresa(empresa).toLowerCase()}), num total de ${mt(total)} MT, dos quais foram pagos ${mt(pago)} MT (${pct(pago, total)}). O saldo por pagar é de ${mt(total - pago)} MT.`,
    `Em ${nomeMes(ano, mes)} vencem ${int(doMes.length)} prestação(ões), no valor de ${mt(soma(doMes))} MT, das quais ${mt(soma(doMes.filter(I.paga)))} MT já pagos.` +
      (atraso.length ? ` Há ${int(atraso.length)} prestação(ões) em atraso, no valor de ${mt(soma(atraso))} MT, que requerem regularização.` : " Não há prestações em atraso.")]});
  sec.push({titulo: "Prestações de " + nomeMes(ano, mes), tabela: doMes.length ? {
    cols: [{t: "Colaborador"}, {t: "Empresa"}, {t: "Prest.", num: true}, {t: "Data limite"}, {t: "Valor", num: true}, {t: "Situação"}, {t: "Pago em"}],
    linhas: doMes.map(p => { const i = lista.find(x => x.id === p.indemnizacao); return [p.nome || i?.nome || p.codigo, i?.empresa || p.empresa, p.numero + "/" + (i?.n || "—"), dataPt(p.dataLimite), mt(p.valor), I.situacao(p, hj), dataPt(p.dataPagamento)]; }),
    total: ["Total", "", "", "", mt(soma(doMes)), "", ""]} : null,
    paragrafos: doMes.length ? [] : ["Não há prestações com data limite neste mês."]});
  sec.push({titulo: "Planos de indemnização", tabela: {
    cols: [{t: "Colaborador"}, {t: "Empresa"}, {t: "Saída"}, {t: "Anos", num: true}, {t: "Venc. base", num: true}, {t: "Total", num: true}, {t: "Pago", num: true}, {t: "Saldo", num: true}, {t: "Pagas", num: true}, {t: "Próxima"}],
    linhas: resumos.sort((a, b) => (a.r.proxima?.dataLimite || "9999").localeCompare(b.r.proxima?.dataLimite || "9999")).map(({i, r: x}) =>
      [i.nome, i.empresa, dataPt(i.referencia), int(i.anos), mt(i.base), mt(i.total), mt(x.pago), mt(x.saldo), x.pagas + "/" + x.prestacoes.length, x.proxima ? dataPt(x.proxima.dataLimite) : "Concluído"]),
    total: ["Total", "", "", "", "", mt(total), mt(pago), mt(total - pago), "", ""]}});
  if (!empresa) {
    const g = agrupar(lista, i => i.empresa, {total: i => i.total, pago: i => I.resumo(i, pre).pago});
    sec.push({titulo: "Por empresa", tabela: {cols: [{t: "Empresa"}, {t: "Planos", num: true}, {t: "Total", num: true}, {t: "Pago", num: true}, {t: "Saldo", num: true}, {t: "% do total", num: true}],
      linhas: g.sort((a, b) => b.total - a.total).map(x => [x.chave, int(x.n), mt(x.total), mt(x.pago), mt(x.total - x.pago), pct(x.total, total)]),
      total: ["Total", int(lista.length), mt(total), mt(pago), mt(total - pago), "100,0%"]}});
  }
  const meses = I.totaisMensais(lista, pre, ano, mes, 0, 11);
  sec.push({titulo: "Calendário de pagamentos — próximos 12 meses", tabela: {
    cols: [{t: "Mês"}, {t: "Prestações", num: true}, {t: "Previsto", num: true}, {t: "Pago", num: true}, {t: "Por pagar", num: true}],
    linhas: meses.map(m => [M.NOME_MES[m.mes - 1] + " " + m.ano, m.n ? int(m.n) : "—", m.n ? mt(m.previsto) : "—", m.n ? mt(m.pago) : "—", m.n ? mt(m.previsto - m.pago) : "—"]),
    total: ["Total", int(meses.reduce((s, m) => s + m.n, 0)), mt(meses.reduce((s, m) => s + m.previsto, 0)), mt(meses.reduce((s, m) => s + m.pago, 0)), mt(meses.reduce((s, m) => s + m.previsto - m.pago, 0))]}});
  if (atraso.length) sec.push({titulo: "Prestações em atraso", tabela: {
    cols: [{t: "Colaborador"}, {t: "Data limite"}, {t: "Dias", num: true}, {t: "Valor", num: true}],
    linhas: atraso.map(p => [p.nome || p.codigo, dataPt(p.dataLimite), int((Date.parse(hj) - Date.parse(p.dataLimite)) / 86400000), mt(p.valor)]),
    total: ["Total", "", "", mt(soma(atraso))]}});
  r.secoes = sec.map((x, k) => ({...x, titulo: (k + 1) + ". " + x.titulo}));
  r.notas.push("Indemnização = anos completos de trabalho (da admissão à data de saída) × vencimento base. Planos anulados não são considerados.");
  r.notas.push("Situação «Em atraso»: prestação não paga cuja data limite já passou à data do relatório.");
  return r;
}
