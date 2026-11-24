// Dashboard RH: junta o cadastro e os últimos 12 meses de lançamentos e entrega-os ao dashboard
// Streamlit (painel.html, que corre com stlite dentro de um iframe da mesma origem).
// Só seguem as colunas que o dashboard usa; os dados nunca saem do navegador.
import * as M from "./rh-modulos.js";

export const MESES_HISTORICO = 12;

const COLUNAS_FUNCIONARIOS = {
  "Empresa": ["Empresa"], "Codigo": ["Codigo", "Código"], "Nome": ["Nome"], "Departamento": ["Departamento"],
  "Profissao": ["Profissao", "Profissão"], "Sexo": ["Sexo", "Genero", "Género"],
  "Data Nascimento": ["Data Nascimento", "Data de Nascimento", "DataNascimento"],
  "Data Admissao": ["Data Admissao", "Data de Admissão", "Data Admissão"], "Data Saida": ["Data Saida", "Data de Saída", "Data Saída"],
  "Estado": ["Estado"], "Estado de Reforma": ["Estado de Reforma", "Estado Reforma"], "TipoContrato": ["TipoContrato", "Tipo de Contrato"],
  "RegimePagamento": ["RegimePagamento", "Regime de Pagamento"], "Nacionalidade": ["Nacionalidade"], "Estado Civil": ["Estado Civil"],
  "Motivo de saída": ["Motivo de saída", "Motivo de Saída", "Motivo Saida"]
};
const COLUNAS_LANCAMENTOS = ["Empresa", "Codigo", "Nome", "Departamento", "Profissao", "Ano", "Mês", "RegimePagamento", "Dias Trabalhados",
  "Faltas Justificadas", "Faltas Injustificadas", "Horas Extras", "Subsídios", "Bónus", "Outros Acréscimos", "Adiantamentos", "Empréstimos",
  "Outros Descontos", "Vencimento Base", "Total de Acréscimos", "Total de Descontos", "INSS Colaborador", "INSS Empresa", "Total INSS",
  "IRPS", "Vencimento", "Salario Liquido", "Estado do Lançamento"];

const ALT_LANC = {"Codigo": ["Código"], "Mês": ["Mes"], "Profissao": ["Profissão"], "Salario Liquido": ["Salário Líquido"]};

export function linhasFuncionarios(rows) {
  return rows.map(l => Object.fromEntries(Object.entries(COLUNAS_FUNCIONARIOS).map(([k, alt]) => [k, M.campo(l, ...alt)])));
}
const COLUNAS_IND = ["ID_Indemnizacao", "Empresa", "Codigo", "Nome", "Departamento", "Data Admissao", "Data Referencia",
  "Anos de Trabalho", "Vencimento Base", "Valor Total", "Numero Prestacoes", "Estado"];
const COLUNAS_PRE = ["ID_Prestacao", "ID_Indemnizacao", "Empresa", "Codigo", "Nome", "Numero", "Data Limite", "Valor", "Estado", "Data Pagamento"];
const so = cols => rows => (rows || []).map(l => Object.fromEntries(cols.map(k => [k, l[k] ?? ""])));
export const linhasIndemnizacoes = so(COLUNAS_IND);
export const linhasPrestacoes = so(COLUNAS_PRE);

export function linhasLancamentos(rows) {
  return rows.map(l => Object.fromEntries(COLUNAS_LANCAMENTOS.map(k => [k, M.campo(l, k, ...(ALT_LANC[k] || []))])));
}

// Últimos n meses terminando em {ano, mes} (inclusive), do mais antigo para o mais recente.
export function mesesAte({ano, mes}, n = MESES_HISTORICO) {
  const out = [];
  for (let i = n - 1; i >= 0; i--) {
    const d = new Date(ano, mes - 1 - i, 1);
    out.push({ano: d.getFullYear(), mes: d.getMonth() + 1});
  }
  return out;
}

export function iniciarPainel(ctx) {
  const $ = id => document.getElementById(id);
  let iframe = null, pronto = null, activo = null, carregado = false, geracao = 0;
  let removerOuvinte = null, resposta = null, pedido = 0, rejeitarPronto = null;
  const estado = txt => { $("painel-estado").textContent = txt; };
  function garantirIframe() {
    if (iframe) return pronto;
    iframe = document.createElement("iframe");
    iframe.id = "painel-frame"; iframe.title = "Dashboard RH"; iframe.src = "painel.html";
    pronto = new Promise((ok, falha) => {
      const timer = setTimeout(() => falha(Error("O dashboard não respondeu. Tente novamente.")), 180000);
      rejeitarPronto = () => { clearTimeout(timer); falha(Error("Dashboard interrompido.")); };
      const ouvir = ev => {
        if (ev.origin !== location.origin || ev.source !== iframe?.contentWindow) return;
        if (ev.data?.tipo === "rh-dashboard-ola") { clearTimeout(timer); ok(); }
        if (ev.data?.tipo === "rh-dashboard-pronto" && ev.data.pedido === resposta?.pedido) resposta.ok();
        if (ev.data?.tipo === "rh-dashboard-erro" && ev.data.pedido === resposta?.pedido) resposta.falha(Error(ev.data.erro));
        if (ev.data?.tipo === "rh-dashboard-reiniciar") { limpar(); abrir(true); }
      };
      window.addEventListener("message", ouvir);
      removerOuvinte = () => window.removeEventListener("message", ouvir);
    });
    pronto.catch(() => {});
    $("painel-quadro").replaceChildren(iframe);
    return pronto;
  }
  async function lerMeses(lista, forcar, progresso, g) {
    const rows = [], falhas = [], fila = [...lista]; let feitos = 0;
    async function trabalhador() {
      while (fila.length && g === geracao) {
        const p = fila.shift();
        try { rows.push(...await ctx.carregarMes(p.ano, p.mes, forcar)); } catch { falhas.push(p); }
        if (g !== geracao) return;
        feitos++;
        estado(`A ler lançamentos… ${feitos} de ${lista.length} meses`);
        progresso?.(feitos, lista.length);
      }
    }
    await Promise.all([trabalhador(), trabalhador(), trabalhador()]);
    return {rows, falhas};
  }
  function abrir(forcar = false, {estrito = false, progresso} = {}) {
    if (!ctx.ligado()) { $("painel-vazio").hidden = false; $("painel-conteudo").hidden = true; return Promise.resolve(); }
    $("painel-vazio").hidden = true; $("painel-conteudo").hidden = false;
    if (activo) return activo;
    if (carregado && !forcar) return Promise.resolve();
    const g = geracao;
    const verificar = () => { if (g !== geracao) throw Error("Dashboard interrompido."); };
    $("painel-actualizar").disabled = true;
    activo = (async () => {
      try {
        const espera = garantirIframe();
        estado("A ler o cadastro de colaboradores…");
        const funcionarios = await ctx.funcionarios(forcar); verificar();
        const ultimo = (await ctx.periodoMaisRecente()) || {ano: new Date().getFullYear(), mes: new Date().getMonth() + 1}; verificar();
        const lista = mesesAte(ultimo);
        const {rows, falhas} = await lerMeses(lista, forcar, progresso, g); verificar();
        if (estrito && falhas.length) throw Error("Dashboard: não foi possível ler " + falhas.length + " mês(es) do histórico. Tente novamente.");
        let ind = null;
        if (ctx.indemnizacoes) ind = await ctx.indemnizacoes(forcar).catch(e => { if (estrito) throw e; return null; });
        verificar();
        const nome = p => M.NOME_MES[p.mes - 1].slice(0, 3) + " " + p.ano;
        const meta = {gerado: new Date().toLocaleString("pt-MZ", {dateStyle: "short", timeStyle: "short"}), indemnizacoes: Boolean(ind),
          meses: `${nome(lista[0])} – ${nome(lista[lista.length - 1])} (${MESES_HISTORICO} meses)` + (falhas.length ? ` · ${falhas.length} mês(es) não lidos` : "")};
        await espera; verificar();
        const id = ++pedido;
        estado("A apresentar os gráficos do Dashboard RH…");
        progresso?.(12, 12, "A apresentar os gráficos do Dashboard RH…");
        await new Promise((ok, falha) => {
          const timer = setTimeout(() => falha(Error("O dashboard não terminou a apresentação. Tente novamente.")), 180000);
          resposta = {pedido: id, ok: () => { clearTimeout(timer); ok(); }, falha: e => { clearTimeout(timer); falha(e); }};
          iframe.contentWindow.postMessage({tipo: "rh-dashboard-dados", pedido: id, dados: {
            funcionarios: linhasFuncionarios(funcionarios), lancamentos: linhasLancamentos(rows),
            indemnizacoes: linhasIndemnizacoes(ind?.indemnizacoes), prestacoes: linhasPrestacoes(ind?.prestacoes), meta}}, location.origin);
        });
        verificar(); carregado = true;
        estado("Dashboard actualizado: " + new Date().toLocaleTimeString("pt-MZ", {hour: "2-digit", minute: "2-digit"}));
      } catch (e) {
        if (g === geracao) estado("Não foi possível preparar o dashboard: " + (e.message || e));
        if (estrito) throw e;
      } finally {
        if (g === geracao) { activo = null; resposta = null; $("painel-actualizar").disabled = false; }
      }
    })();
    return activo;
  }
  function limpar() {
    geracao++; rejeitarPronto?.(); resposta?.falha(Error("Dashboard interrompido.")); removerOuvinte?.();
    if (iframe) iframe.remove();
    iframe = null; pronto = null; activo = null; resposta = null; carregado = false;
    removerOuvinte = rejeitarPronto = null;
    $("painel-quadro").replaceChildren(); estado("");
    $("painel-vazio").hidden = false; $("painel-conteudo").hidden = true; $("painel-actualizar").disabled = false;
  }
  $("painel-actualizar").addEventListener("click", () => { limpar(); abrir(true); });
  return {abrir, limpar};
}
