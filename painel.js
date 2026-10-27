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
  let iframe = null, pronto = null, aCarregar = false;

  function estado(txt) { $("painel-estado").textContent = txt; }

  function garantirIframe() {
    if (iframe) return pronto;
    iframe = document.createElement("iframe");
    iframe.id = "painel-frame";
    iframe.title = "Dashboard RH";
    iframe.src = "painel.html";
    pronto = new Promise(ok => {
      const ouvir = ev => {
        if (ev.origin !== location.origin || ev.source !== iframe?.contentWindow) return;
        if (ev.data?.tipo === "rh-dashboard-ola") ok();
        if (ev.data?.tipo === "rh-dashboard-pronto") estado("Dashboard actualizado: " + new Date().toLocaleTimeString("pt-MZ", {hour: "2-digit", minute: "2-digit"}));
        if (ev.data?.tipo === "rh-dashboard-erro") estado("Erro no dashboard: " + ev.data.erro);
        if (ev.data?.tipo === "rh-dashboard-reiniciar") reiniciar();
      };
      window.addEventListener("message", ouvir);
    });
    $("painel-quadro").replaceChildren(iframe);
    return pronto;
  }

  async function lerMeses(lista, forcar) {
    const rows = [], falhas = [];
    let feitos = 0;
    const fila = [...lista];
    async function trabalhador() {
      while (fila.length) {
        const p = fila.shift();
        try { rows.push(...await ctx.carregarMes(p.ano, p.mes, forcar)); } catch { falhas.push(p); }
        feitos += 1;
        estado(`A ler lançamentos do Excel… ${feitos} de ${lista.length} meses`);
      }
    }
    await Promise.all([trabalhador(), trabalhador(), trabalhador()]); // 3 leituras em paralelo
    return {rows, falhas};
  }

  async function abrir(forcar = false) {
    if (!ctx.ligado()) { $("painel-vazio").hidden = false; $("painel-conteudo").hidden = true; return; }
    $("painel-vazio").hidden = true; $("painel-conteudo").hidden = false;
    if (iframe && !forcar) return;
    if (aCarregar) return;
    aCarregar = true; $("painel-actualizar").disabled = true;
    const t0 = Date.now();
    try {
      const espera = garantirIframe();
      estado("A ler o cadastro de colaboradores…");
      const funcionarios = await ctx.funcionarios(forcar);
      // Último mês com lançamentos (o sistema procura-o ao entrar); sem ele, o mês actual.
      const ultimo = (await ctx.periodoMaisRecente()) || {ano: new Date().getFullYear(), mes: new Date().getMonth() + 1};
      const lista = mesesAte(ultimo);
      const {rows, falhas} = await lerMeses(lista, forcar);
      const nome = p => M.NOME_MES[p.mes - 1].slice(0, 3) + " " + p.ano;
      const meta = {
        gerado: new Date().toLocaleString("pt-MZ", {dateStyle: "short", timeStyle: "short"}),
        meses: `${nome(lista[0])} – ${nome(lista[lista.length - 1])} (${MESES_HISTORICO} meses)` + (falhas.length ? ` · ${falhas.length} mês(es) não lidos` : "")
      };
      await espera;
      iframe.contentWindow.postMessage({tipo: "rh-dashboard-dados", dados: {
        funcionarios: linhasFuncionarios(funcionarios), lancamentos: linhasLancamentos(rows), meta}}, location.origin);
      estado(`Dados enviados (${Math.round((Date.now() - t0) / 1000)} s): ${funcionarios.length} colaboradores, ${rows.length} lançamentos` +
        (falhas.length ? ` · não foi possível ler ${falhas.map(nome).join(", ")}` : "") + ". A abrir o dashboard…");
    } catch (e) {
      estado("Não foi possível preparar o dashboard: " + (e.message || e));
    } finally {
      aCarregar = false; $("painel-actualizar").disabled = false;
    }
  }

  // «Tentar de novo» (pedido pelo próprio dashboard quando não arranca): novo iframe e nova leitura.
  function reiniciar() {
    if (iframe) iframe.remove();
    iframe = null; pronto = null;
    abrir(true);
  }

  // Ao terminar sessão, o dashboard e os dados que tem são descartados.
  function limpar() {
    if (iframe) iframe.remove();
    iframe = null; pronto = null;
    $("painel-quadro").replaceChildren();
    estado("");
    $("painel-vazio").hidden = false; $("painel-conteudo").hidden = true;
  }

  $("painel-actualizar").addEventListener("click", () => abrir(true));
  return {abrir, limpar};
}
