// Indemnizações: cálculo (anos completos × vencimento base), plano de prestações e indicadores por mês.
// Lógica pura (sem DOM), partilhada pela página e pelos testes. Datas em ISO (AAAA-MM-DD).
import * as M from "./rh-modulos.js";

const texto = v => String(v ?? "").trim();
const cent = x => Math.round((Number(x) + Number.EPSILON) * 100) / 100;
const pad = n => String(n).padStart(2, "0");
export const iso = (a, m, d) => a + "-" + pad(m) + "-" + pad(d);
const diasNoMes = (a, m) => new Date(Date.UTC(a, m, 0)).getUTCDate();

// Anos completos entre a admissão e a data de referência (aniversário ainda não atingido não conta).
export function anosCompletos(admissao, referencia) {
  const a = M.data(admissao), r = M.data(referencia);
  if (!a || !r || r < a) return 0;
  const [ay, am, ad] = a.split("-").map(Number), [ry, rm, rd] = r.split("-").map(Number);
  let anos = ry - ay;
  if (rm < am || (rm === am && rd < ad)) anos -= 1;
  return Math.max(0, anos);
}

export function calcular({admissao, referencia, vencimentoBase}) {
  const anos = anosCompletos(admissao, referencia), base = M.numero(vencimentoBase) ?? 0;
  return {anos, base, total: cent(anos * base)};
}

// Data limite de um mês com um dia fixo; se o mês não tiver esse dia, usa o último dia do mês.
export function dataLimite(ano, mes, dia) {
  return iso(ano, mes, Math.min(Math.max(1, Number(dia) || 1), diasNoMes(ano, mes)));
}

// Divisão equitativa em n meses: prestações iguais em cêntimos; a última acerta a diferença.
export function gerarPlano(total, n, dia, primeiroMes) {
  const t = cent(total), k = Math.trunc(Number(n));
  if (!(t > 0) || !(k >= 1)) return [];
  const [a0, m0] = String(primeiroMes || "").split("-").map(Number);
  if (!a0 || !m0) return [];
  const parcela = Math.floor((t / k) * 100) / 100;
  return Array.from({length: k}, (_, i) => {
    const d = new Date(Date.UTC(a0, m0 - 1 + i, 1));
    const valor = i === k - 1 ? cent(t - parcela * (k - 1)) : parcela;
    return {numero: i + 1, dataLimite: dataLimite(d.getUTCFullYear(), d.getUTCMonth() + 1, dia), valor};
  });
}

export function validarPlano(plano, total) {
  const erros = [];
  if (!plano.length) erros.push("O plano não tem prestações.");
  let anterior = "";
  plano.forEach((p, i) => {
    if (!M.data(p.dataLimite)) erros.push("Data limite inválida na prestação " + (i + 1) + ".");
    else if (anterior && p.dataLimite < anterior) erros.push("A prestação " + (i + 1) + " tem data anterior à prestação " + i + ".");
    anterior = p.dataLimite || anterior;
    if (!(Number(p.valor) > 0)) erros.push("Valor inválido na prestação " + (i + 1) + ".");
  });
  const soma = cent(plano.reduce((s, p) => s + Number(p.valor || 0), 0));
  if (plano.length && Math.abs(soma - cent(total)) > 0.01) erros.push("A soma das prestações (" + soma + ") é diferente do total (" + cent(total) + ").");
  return erros;
}

// Linhas do Excel -> objectos do site.
export function indemnizacoes(linhas = []) {
  return linhas.map(l => ({
    id: texto(M.campo(l, "ID_Indemnizacao")), empresa: texto(M.campo(l, "Empresa")), codigo: texto(M.campo(l, "Codigo")),
    nome: texto(M.campo(l, "Nome")), departamento: texto(M.campo(l, "Departamento")),
    admissao: M.data(M.campo(l, "Data Admissao")), referencia: M.data(M.campo(l, "Data Referencia")),
    anos: M.numero(M.campo(l, "Anos de Trabalho")) ?? 0, base: M.numero(M.campo(l, "Vencimento Base")) ?? 0,
    total: M.numero(M.campo(l, "Valor Total")) ?? 0, n: M.numero(M.campo(l, "Numero Prestacoes")) ?? 0,
    estado: texto(M.campo(l, "Estado")) || "Activa", observacoes: texto(M.campo(l, "Observacoes")),
    criadoPor: texto(M.campo(l, "Criado Por")), criadoEm: M.data(M.campo(l, "Data Criacao"))
  })).filter(x => x.id);
}
export function prestacoes(linhas = []) {
  return linhas.map(l => ({
    id: texto(M.campo(l, "ID_Prestacao")), indemnizacao: texto(M.campo(l, "ID_Indemnizacao")),
    empresa: texto(M.campo(l, "Empresa")), codigo: texto(M.campo(l, "Codigo")), nome: texto(M.campo(l, "Nome")),
    numero: M.numero(M.campo(l, "Numero")) ?? 0, dataLimite: M.data(M.campo(l, "Data Limite")),
    valor: M.numero(M.campo(l, "Valor")) ?? 0, estado: texto(M.campo(l, "Estado")) || "Pendente",
    dataPagamento: M.data(M.campo(l, "Data Pagamento")), pagoPor: texto(M.campo(l, "Pago Por"))
  })).filter(x => x.id);
}

export const paga = p => /^paga$/i.test(p.estado);
export const emAtraso = (p, hoje) => !paga(p) && p.dataLimite && p.dataLimite < hoje;
export function situacao(p, hoje) { return paga(p) ? "Paga" : emAtraso(p, hoje) ? "Em atraso" : "Pendente"; }

// Só contam as prestações de indemnizações não anuladas.
export function activas(lista, prest) {
  const anuladas = new Set(lista.filter(i => /anulad/i.test(i.estado)).map(i => i.id));
  return prest.filter(p => !anuladas.has(p.indemnizacao));
}

// Indicadores do mês (prestações com data limite nesse mês) e acumulados até ao fim do mês.
export function indicadoresMes(lista, prest, ano, mes, hoje) {
  const validas = activas(lista, prest), chave = ano + "-" + pad(mes), fim = iso(ano, mes, diasNoMes(ano, mes));
  const doMes = validas.filter(p => (p.dataLimite || "").slice(0, 7) === chave);
  const soma = arr => cent(arr.reduce((s, p) => s + p.valor, 0));
  const pagasAte = validas.filter(p => paga(p) && (p.dataPagamento || p.dataLimite) <= fim);
  const comprometido = cent(lista.filter(i => !/anulad/i.test(i.estado)).reduce((s, i) => s + i.total, 0));
  const gastoTotal = soma(validas.filter(paga));
  return {
    doMes, valorMes: soma(doMes), pagoMes: soma(doMes.filter(paga)), porPagarMes: soma(doMes.filter(p => !paga(p))),
    atrasoMes: soma(doMes.filter(p => emAtraso(p, hoje))), nMes: doMes.length,
    gastoAteMes: soma(pagasAte), gastoTotal, comprometido, saldo: cent(comprometido - gastoTotal),
    atrasoTotal: soma(validas.filter(p => emAtraso(p, hoje))), nAtraso: validas.filter(p => emAtraso(p, hoje)).length
  };
}

// Totais por mês (para a faixa de meses): previsto e pago, n meses à volta de {ano, mes}.
export function totaisMensais(lista, prest, ano, mes, antes = 5, depois = 6) {
  const validas = activas(lista, prest), out = [];
  for (let i = -antes; i <= depois; i++) {
    const d = new Date(Date.UTC(ano, mes - 1 + i, 1)), a = d.getUTCFullYear(), m = d.getUTCMonth() + 1, k = a + "-" + pad(m);
    const doMes = validas.filter(p => (p.dataLimite || "").slice(0, 7) === k);
    out.push({ano: a, mes: m, previsto: cent(doMes.reduce((s, p) => s + p.valor, 0)), pago: cent(doMes.filter(paga).reduce((s, p) => s + p.valor, 0)), n: doMes.length});
  }
  return out;
}

// Resumo por indemnização: pagas, valor pago e saldo.
export function resumo(ind, prest) {
  const minhas = prest.filter(p => p.indemnizacao === ind.id).sort((a, b) => a.numero - b.numero);
  const pago = cent(minhas.filter(paga).reduce((s, p) => s + p.valor, 0));
  const proxima = minhas.find(p => !paga(p));
  return {prestacoes: minhas, pagas: minhas.filter(paga).length, pago, saldo: cent(ind.total - pago), proxima,
    concluida: minhas.length > 0 && minhas.every(paga)};
}

// Semanas do calendário (segunda a domingo) com as prestações de cada dia.
export function calendario(ano, mes, prest) {
  const n = diasNoMes(ano, mes), primeiro = (new Date(Date.UTC(ano, mes - 1, 1)).getUTCDay() + 6) % 7;
  const dias = [];
  for (let i = 0; i < primeiro; i++) dias.push(null);
  for (let d = 1; d <= n; d++) { const k = iso(ano, mes, d); dias.push({dia: d, data: k, prestacoes: prest.filter(p => p.dataLimite === k)}); }
  while (dias.length % 7) dias.push(null);
  const semanas = [];
  for (let i = 0; i < dias.length; i += 7) semanas.push(dias.slice(i, i + 7));
  return semanas;
}
