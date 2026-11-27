import {criarProgresso} from "./rh-progresso.js";
// Lançamento salarial: pré-visualização da folha, exportação PDF/Excel, relatório comparativo e Novo mês.
import * as M from "./rh-modulos.js";
import * as F from "./rh-folha.js";
import * as X from "./rh-export.js";
import * as LC from "./lancamentos-colunas.js";

const $ = id => document.getElementById(id);
const el = (tag, props = {}, ...filhos) => {
  const n = document.createElement(tag);
  for (const [k, v] of Object.entries(props)) {
    if (k === "class") n.className = v; else if (k === "text") n.textContent = v;
    else if (k.startsWith("data-") || k === "colspan") n.setAttribute(k, v); else n[k] = v;
  }
  n.append(...filhos.filter(f => f != null && f !== false));
  return n;
};
const fmt = v => new Intl.NumberFormat("pt-PT", {minimumFractionDigits: 2, maximumFractionDigits: 2, useGrouping: "always"}).format(v || 0);
const pct = v => v == null ? "—" : (v * 100).toFixed(2).replace(".", ",") + "%";
const tabela = (cabecalho, linhas, {num = [], rodape = null, classe = ""} = {}) => el("div", {class: "table-scroll"},
  el("table", {class: "rel-table " + classe},
    el("thead", {}, el("tr", {}, ...cabecalho.map((t, i) => el("th", {class: num.includes(i) ? "num" : "", text: t})))),
    el("tbody", {}, ...linhas.map(l => el("tr", {}, ...l.map((v, i) => el("td", {class: num.includes(i) ? "num" : "", text: v ?? ""}))))),
    rodape && el("tfoot", {}, el("tr", {}, ...rodape.map((v, i) => el("td", {class: num.includes(i) ? "num" : "", text: v ?? ""}))))));
const cartao = (rotulo, valor, sub = "") => el("article", {class: "metric mini"}, el("span", {text: rotulo}), el("strong", {text: valor}), sub && el("small", {text: sub}));

export function iniciarLancamentos(ctx) {
  const dlg = $("relatorio");
  let criandoMes = false;
  let tarefa = 0; // invalida pedidos antigos quando o diálogo é reaberto

  function abrir({eyebrow, titulo, sub = ""}) {
    $("rel-eyebrow").textContent = eyebrow; $("rel-titulo").textContent = titulo; $("rel-sub").textContent = sub;
    $("rel-body").replaceChildren(el("p", {class: "loading-line", text: "A preparar…"}));
    $("rel-accoes").replaceChildren(); $("rel-erro").textContent = "";
    if (!dlg.open) { if (typeof dlg.showModal === "function") dlg.showModal(); else dlg.setAttribute("open", ""); }
    return ++tarefa;
  }
  const fechar = () => { if (criandoMes) return; tarefa++; if (typeof dlg.close === "function") dlg.close(); else dlg.removeAttribute("open"); };
  $("rel-fechar").addEventListener("click", fechar);
  dlg.addEventListener("cancel", e => { if (criandoMes) { e.preventDefault(); return; } tarefa++; });
  function accoes(lista) {
    $("rel-accoes").replaceChildren(el("span"), ...lista.map(a => {
      const b = el("button", {type: "button", class: a.classe || "outline", text: a.texto, disabled: Boolean(a.desactivado)});
      if (a.titulo) b.title = a.titulo;
      b.addEventListener("click", async () => {
        if (!a.click || b.disabled) return;
        const t = b.textContent, ocupado = a.ocupado || "A gerar…", t0 = Date.now();
        b.disabled = true; b.textContent = ocupado; $("rel-erro").textContent = "";
        const relogio = setInterval(() => { if (b.isConnected && b.disabled) b.textContent = ocupado + " " + Math.round((Date.now() - t0) / 1000) + " s"; }, 1000);
        try { await a.click(); } catch (e) { $("rel-erro").textContent = e.message + " (" + Math.round((Date.now() - t0) / 1000) + " s)"; }
        finally { clearInterval(relogio); if (b.isConnected) { b.disabled = false; b.textContent = t; } }
      });
      return b;
    }));
  }
  const erro = e => { $("rel-body").replaceChildren(el("p", {class: "form-error", text: e.message || String(e)})); };

  // ---------- folha ----------
  // Folha = lançamentos do mês com os filtros (Empresa, Departamento, Colaborador), pela ordenação escolhida na tabela.
  // Estrutura padrão da tabela → layout oficial da folha (igual ao VBA). Estrutura alterada → as colunas escolhidas, pela mesma ordem.
  function dadosFolha() {
    const {ano, mes} = ctx.periodo();
    const filtros = ctx.filtros(), vista = ctx.vistaTabela();
    let brutas = F.filtrarFolha(ctx.lancamentosMes(), ano, mes, filtros);
    if (vista.filtrar) brutas = vista.filtrar(brutas);   // filtros do cabeçalho da tabela
    if (!brutas.length) throw Error("Não existem lançamentos para os filtros seleccionados.");
    if (vista.ordem) brutas = LC.ordenar(M.lancamentos(brutas), vista.ordem.col, vista.ordem.dir).map(m => m.linha);
    const ordemTexto = [vista.ordem ? "Ordenado por " + vista.ordem.col.rotulo + " (" + (vista.ordem.dir === "asc" ? "crescente" : "decrescente") + ")" : "",
      vista.filtrosTexto ? "Filtro: " + vista.filtrosTexto : ""].filter(Boolean).join(" · ");
    const d = {ano, mes, filtros, ordemTexto, linhas: F.linhasFolha(brutas, ano, mes)};
    if (vista.personalizada) {
      const modelos = M.lancamentos(brutas), cols = vista.colunas;
      d.personalizada = {
        colunas: cols.map(c => ({rotulo: c.rotulo, num: c.tipo === "numero", total: c.tipo === "numero" && Boolean(c.total), inteiro: Boolean(c.inteiro)})),
        valores: modelos.map(m => cols.map(c => c.tipo === "numero" ? M.numero(c.valor(m)) : LC.formatar(c, m))),
        totais: cols.map(c => c.tipo === "numero" && c.total ? LC.somar(modelos, c) : null)
      };
    }
    return d;
  }
  const textoFiltro = f => "Empresa: " + (f.empresa || "TODOS") + "   |   Departamento: " + (f.departamento || "TODOS") + "   |   " + (f.codigo ? "Código: " + f.codigo : "Colaborador: TODOS");
  async function exportarFolha(tipo) {
    const d = dadosFolha(), opt = {ano: d.ano, mes: d.mes, filtros: d.filtros, utilizador: ctx.utilizador()};
    opt.subtitulo = d.ordemTexto;
    const blob = d.personalizada
      ? (tipo === "pdf" ? await X.folhaPersonalizadaPDF(d.personalizada, opt) : await X.folhaPersonalizadaExcel(d.personalizada, opt))
      : (tipo === "pdf" ? await X.folhaPDF(d.linhas, opt) : await X.folhaExcel(d.linhas, opt));
    X.descarregar(blob, X.nomeFicheiroFolha(d.ano, d.mes, d.filtros, tipo));
    ctx.aviso("Folha exportada: " + X.nomeFicheiroFolha(d.ano, d.mes, d.filtros, tipo));
  }
  function preVisualizar() {
    let d;
    try { d = dadosFolha(); } catch (e) { ctx.aviso(e.message); return; }
    abrir({eyebrow: "PRÉ-VISUALIZAÇÃO", titulo: "Folha de Salários - " + M.NOME_MES[d.mes - 1] + " de " + d.ano, sub: textoFiltro(d.filtros) + (d.ordemTexto ? "   |   " + d.ordemTexto : "")});
    const l = d.linhas, s = k => F.somar(l, k);
    const fv = (c, v) => v == null || v === "" ? "" : c.num ? (c.inteiro ? fmt(v).replace(/,00$/, "") : fmt(v)) : v;
    const P = d.personalizada;
    $("rel-body").replaceChildren(
      ...(P ? [el("p", {class: "rel-nota", text: "Folha com as colunas escolhidas na tabela (" + P.colunas.length + "). Para o layout oficial, reponha a estrutura padrão da tabela."})] : []),
      el("div", {class: "stats compact-stats rel-stats"}, cartao("Colaboradores", String(l.length)), cartao("Vencimento base", fmt(s("base")), "MT"),
        cartao("Acréscimos", fmt(s("acrescimos")), "MT"), cartao("Descontos", fmt(s("totalDesc")), "MT"), cartao("Salário líquido", fmt(s("liquido")), "MT")),
      P ? tabela(P.colunas.map(c => c.rotulo), P.valores.map(r => r.map((v, i) => fv(P.colunas[i], v))),
        {num: P.colunas.map((c, i) => c.num ? i : -1).filter(i => i >= 0), rodape: P.totais.map((t, i) => i === X.colunaRotuloTotais(P.totais) ? "Totais" : t == null ? "" : fmt(t))}) :
      tabela(["Código", "Nome", "Profissão", "Venc. base", "Acréscimos", "IRPS", "INSS", "Outros desc.", "Total desc.", "Líquido"],
        l.map(x => [x.codigo, x.nome, x.profissao, fmt(x.base), fmt(x.acrescimos), fmt(x.irps), fmt(x.inss), fmt(x.outrosDesc), fmt(x.totalDesc), fmt(x.liquido)]),
        {num: [3, 4, 5, 6, 7, 8, 9], rodape: ["Totais", l.length + " colab.", "", fmt(s("base")), fmt(s("acrescimos")), fmt(s("irps")), fmt(s("inss")), fmt(s("outrosDesc")), fmt(s("totalDesc")), fmt(s("liquido"))]}));
    accoes([{texto: "Folha em PDF", click: () => exportarFolha("pdf")}, {texto: "Folha em Excel", click: () => exportarFolha("xlsx")}, {texto: "Fechar", click: fechar}]);
  }

  // ---------- relatório comparativo ----------
  async function comparativo() {
    const {ano, mes} = ctx.periodo();
    const ant = mes === 1 ? {ano: ano - 1, mes: 12} : {ano, mes: mes - 1};
    const t = abrir({eyebrow: "RELATÓRIO COMPARATIVO", titulo: M.NOME_MES[ant.mes - 1] + " " + ant.ano + " vs " + M.NOME_MES[mes - 1] + " " + ano,
      sub: "Base: Salário Líquido · agrupado por Empresa e Departamento"});
    $("rel-body").replaceChildren(el("p", {class: "loading-line", text: "A carregar " + M.NOME_MES[ant.mes - 1] + " " + ant.ano + " do Excel…"}));
    try {
      const anteriores = await ctx.carregarMes(ant.ano, ant.mes);
      if (t !== tarefa) return;
      const r = F.relatorioComparativo([...anteriores, ...ctx.lancamentosMes()], ano, mes);
      if (!r.qtdActual) throw Error("Não existem lançamentos em " + M.NOME_MES[mes - 1] + "/" + ano + ".");
      const variacao = r.totalActual - r.totalAnterior;
      $("rel-body").replaceChildren(
        el("div", {class: "stats compact-stats rel-stats"},
          cartao(M.NOME_MES[ant.mes - 1] + " " + ant.ano, fmt(r.totalAnterior), r.qtdAnterior + " trabalhadores"),
          cartao(M.NOME_MES[mes - 1] + " " + ano, fmt(r.totalActual), r.qtdActual + " trabalhadores"),
          cartao("Variação", (variacao > 0 ? "+" : "") + fmt(variacao), r.totalAnterior ? pct(variacao / r.totalAnterior) : "sem mês anterior")),
        !r.qtdAnterior && el("p", {class: "notice-inline", text: "Não há lançamentos em " + M.NOME_MES[ant.mes - 1] + "/" + ant.ano + ": todos os colaboradores aparecem como entradas."}),
        tabela(["Unidade", "Trab. ant.", M.NOME_MES[ant.mes - 1].slice(0, 3) + " " + ant.ano, "Variação", M.NOME_MES[mes - 1].slice(0, 3) + " " + ano, "Trab. act.", "Peso %", "Comentários"],
          r.grupos.map(g => [g.unidade, g.qtdAnterior, fmt(g.valorAnterior), fmt(g.variacao), fmt(g.valorActual), g.qtdActual, pct(g.peso), g.comentario]),
          {num: [1, 2, 3, 4, 5, 6], classe: "comparativo", rodape: ["Total", r.qtdAnterior, fmt(r.totalAnterior), fmt(variacao), fmt(r.totalActual), r.qtdActual, r.totalActual ? "100,00%" : "", ""]}));
      accoes([
        {texto: "Exportar Excel", click: async () => { X.descarregar(await X.comparativoExcel(r), X.nomeFicheiroComparativo(r, "xlsx")); }},
        {texto: "Exportar PDF", click: async () => { X.descarregar(await X.comparativoPDF(r, {utilizador: ctx.utilizador()}), X.nomeFicheiroComparativo(r, "pdf")); }},
        {texto: "Fechar", click: fechar}]);
    } catch (e) { if (t === tarefa) erro(e); }
  }

  // ---------- Novo mês ----------
  function proximoMes() {
    const p = ctx.periodoMaisRecente() || ctx.periodo();
    return p.mes === 12 ? {ano: p.ano + 1, mes: 1} : {ano: p.ano, mes: p.mes + 1};
  }
  function novoMes() {
    const sug = proximoMes();
    const t = abrir({eyebrow: "NOVO MÊS", titulo: "Criar lançamentos do mês", sub: "Os critérios são os mesmos do sistema VBA (CriarNovoMes)."});
    const selMes = el("select", {id: "nm-mes"}, ...M.NOME_MES.map((m, i) => el("option", {value: String(i + 1), text: m})));
    const selAno = el("select", {id: "nm-ano"}, ...[sug.ano - 1, sug.ano, sug.ano + 1].map(a => el("option", {value: String(a), text: String(a)})));
    selMes.value = String(sug.mes); selAno.value = String(sug.ano);
    const resultado = el("div", {id: "nm-resultado"});
    const progressoHost = el("div", {id: "nm-progresso"});
    $("rel-body").replaceChildren(
      progressoHost,
      el("div", {class: "nm-escolha"}, el("label", {}, "Mês", selMes), el("label", {}, "Ano", selAno),
        el("p", {class: "help", text: "Período de folha 21–20. Entram os colaboradores com contrato vigente no período; as entradas do mês anterior (faltas, horas extra, subsídios, bónus, adiantamentos, empréstimos, outros descontos, dependentes, canalizar e observações) são copiadas; o valor unitário vem do contrato/adenda (Mensal) ou das Settings (Hora/Diário)."})),
      el("ul", {class: "criterios"},
        ...["O mês ainda não pode existir na tbl_lançamento.", "Colaboradores activos com Código e Empresa, sem pares repetidos.", "Contrato vigente no período (admissão, saída, prazo, estado e adendas activas).",
          "Regime de pagamento e valor unitário válidos para todos.", "Mensal: dias trabalhados base 30 (proporcionais na admissão ou saída).", "Estado «Pendente», Revisto «Não», origem «Cópia Mês Anterior» ou «Novo Colaborador»."].map(x => el("li", {text: x}))),
      resultado);
    const verificar = async () => {
      const ano = Number(selAno.value), mes = Number(selMes.value), ant = mes === 1 ? {ano: ano - 1, mes: 12} : {ano, mes: mes - 1};
      resultado.replaceChildren(el("p", {class: "loading-line", text: "A ler colaboradores, adendas, Settings e os lançamentos de " + M.NOME_MES[mes - 1] + " e " + M.NOME_MES[ant.mes - 1] + "… (pode demorar 1 a 2 minutos)"}));
      const progresso = criarProgresso(progressoHost, "Verificar mês", 5);
      selAno.disabled = selMes.disabled = true;
      let feitas = 0;
      try {
      const [funcionarios, contratos, mesAlvo, mesAnt] = await Promise.all([ctx.funcionarios(), ctx.contratos(), ctx.carregarMes(ano, mes), ctx.carregarMes(ant.ano, ant.mes)].map(p => p.then(v => { progresso.actualizar(++feitas, "A verificar os dados do mês…"); return v; })));
      if (t !== tarefa) return;
      const r = F.prepararNovoMes({ano, mes, funcionarios, adendas: contratos.adendas, parametros: contratos.parametros, lancamentosMes: mesAlvo, lancamentosAnterior: mesAnt, utilizador: ctx.utilizador()});
      progresso.concluir(r.erros.length ? "Verificação concluída com bloqueios; corrija antes de criar." : "Verificação concluída. Pronto para criar o mês.");
      mostrarVerificacao(r, ano, mes, t, () => { ctx.esquecerMes(ano, mes); ctx.esquecerMes(ant.ano, ant.mes); return verificar(); });
      } catch (e) { progresso.falhar(e.message); throw e; }
      finally { selAno.disabled = selMes.disabled = false; }
    };
    accoes([{texto: "Verificar critérios", classe: "primary", ocupado: "A verificar…", click: verificar}, {texto: "Fechar", click: fechar}]);
  }

  function mostrarVerificacao(r, ano, mes, t, reverificar) {
    const copia = r.incluidos.filter(i => i.origem === "Cópia Mês Anterior").length;
    const bloco = [
      el("div", {class: "stats compact-stats rel-stats"},
        cartao("A lançar", String(r.linhas.length), copia + " cópia · " + (r.linhas.length - copia) + " novos"),
        cartao("Excluídos", String(r.excluidos.length), "sem contrato vigente"),
        cartao("Bloqueios", String(r.erros.length), r.erros.length ? "corrigir antes de criar" : "nenhum"),
        cartao("Período", r.periodo ? r.periodo.inicio.split("-").reverse().join("/") : "—", r.periodo ? "a " + r.periodo.fim.split("-").reverse().join("/") : ""))
    ];
    if (r.erros.length) bloco.push(el("div", {class: "nm-erros"}, el("h3", {text: "Bloqueios (" + r.erros.length + ")"}), el("ul", {}, ...r.erros.map(e => el("li", {text: e})))));
    if (r.avisos.length) bloco.push(el("div", {class: "notice-inline"}, ...r.avisos.map(a => el("p", {text: a}))));
    if (r.incluidos.length) bloco.push(el("details", {open: !r.erros.length}, el("summary", {text: "Colaboradores a lançar (" + r.incluidos.length + ")"}),
      tabela(["Código", "Nome", "Empresa", "Departamento", "Regime", "Valor unitário", "Dias", "Origem"],
        r.incluidos.map(i => [i.codigo, i.nome, i.empresa, i.departamento, i.regime, fmt(i.valorUnitario) + " " + i.unidade, i.dias, i.origem]), {num: [5, 6]})));
    if (r.excluidos.length) bloco.push(el("details", {}, el("summary", {text: "Excluídos (" + r.excluidos.length + ")"}),
      tabela(["Código", "Nome", "Empresa", "Motivo"], r.excluidos.map(e => [e.codigo, e.nome, e.empresa, e.motivo]))));
    const pode = ctx.podeGravar(), semUrl = !ctx.urlGravacao();
    const url = el("label", {class: "nm-url", hidden: !semUrl}, "Endereço do fluxo RH_API_Actualizar", el("input", {type: "url", id: "nm-url", autocomplete: "off", placeholder: "https://…powerplatform.com/…&sig=…"}));
    bloco.push(url);
    $("nm-resultado").replaceChildren(...bloco);
    const nome = M.NOME_MES[mes - 1] + " " + ano;
    accoes([
      {texto: "Verificar de novo", ocupado: "A verificar…", click: reverificar},
      {texto: "Criar " + nome + " (" + r.linhas.length + ")", classe: "primary", ocupado: "A enviar…", desactivado: r.erros.length > 0 || !r.linhas.length || !pode,
        titulo: !pode ? "Só Superadministrador ou Administrador RH" : r.erros.length ? "Corrija os bloqueios primeiro" : "",
        click: () => ctx.operarMes(() => criar(r, ano, mes, t))},
      {texto: "Fechar", click: fechar}]);
    if (!pode) $("rel-erro").textContent = "O seu perfil só permite consultar. Criar o mês exige Superadministrador ou Administrador RH.";
  }
  async function criar(r, ano, mes, t) {
    const campoUrl = $("nm-url");
    if (campoUrl && !campoUrl.closest("label").hidden && campoUrl.value.trim()) ctx.guardarUrlGravacao(campoUrl.value);
    if (!ctx.urlGravacao()) throw Error("Falta o endereço do fluxo RH_API_Actualizar (peça-o ao administrador).");
    const progresso = criarProgresso($("nm-progresso"), "Criar mês", 3);
    const controles = [...$("rel-accoes").querySelectorAll("button"), $("nm-ano"), $("nm-mes"), $("rel-fechar")].filter(Boolean);
    const anteriores = controles.map(b => b.disabled);
    criandoMes = true;
    controles.forEach(b => b.disabled = true);
    $("rel-body").scrollTop = 0;
    try {
      // Verifica outra vez que o mês continua vazio antes de enviar.
      ctx.esquecerMes(ano, mes);
      const agora = await ctx.carregarMes(ano, mes);
      if (t !== tarefa) return;
      if (agora.length) throw Error("Entretanto já foram criados lançamentos para " + M.NOME_MES[mes - 1] + "/" + ano + ".");
      progresso.actualizar(1, "Mês disponível. A aguardar a gravação no Excel…");
      const res = await ctx.gravarMes({ano, mes, mesNome: r.mesNome, linhas: r.linhas});
      progresso.actualizar(2, "Pedido aceite. A confirmar os lançamentos gravados…");
      if (t !== tarefa) return;
      await acompanhar(res.total, ano, mes, t, progresso);
    } catch (e) { progresso.falhar(e.message); throw e; }
    finally { criandoMes = false; controles.forEach((b, i) => b.disabled = anteriores[i]); }
  }

  async function acompanhar(total, ano, mes, t, progresso) {
    const inicio = Date.now();
    let primeira = true;
    while (Date.now() - inicio < 20 * 60000) {
      if (!primeira) await new Promise(ok => setTimeout(ok, 30000));
      primeira = false;
      if (t !== tarefa) { ctx.esquecerMes(ano, mes); return; }
      ctx.esquecerMes(ano, mes);
      let linhas;
      try { linhas = (await ctx.carregarMes(ano, mes)).filter(l => F.linhasFolha([l], ano, mes).length); }
      catch { progresso.actualizar(2, "Não foi possível confirmar os dados; a tentar novamente…"); continue; }
      if (t !== tarefa) return;
      progresso.actualizar(2, linhas.length + " de " + total + " linhas confirmadas. A aguardar conclusão…");
      if (linhas.length >= total) {
        const semFormula = linhas.filter(l => M.campo(l, "Salario Liquido") === "" || M.campo(l, "Vencimento") === "").length;
        if (semFormula) {
          progresso.falhar("Gravadas " + linhas.length + " linhas, mas " + semFormula + " sem Vencimento/Salário líquido calculado. Confirme as fórmulas no Excel.");
        } else progresso.concluir(linhas.length + " lançamentos criados com estado Pendente.");
        accoes([{texto: "Fechar", click: fechar}]);
        ctx.mostrarMes(ano, mes);
        return;
      }
    }
    throw Error("A confirmação está a demorar mais do que o previsto. Consulte o histórico do RH_API_Actualizar antes de repetir a criação.");
  }


  // ---------- editar um lançamento e finalizar o mês (G_Colaboradores: btnGravarAlteracaoMes / btnFinalizarMes) ----------
  const campoUrl = () => el("label", {class: "nm-url", hidden: Boolean(ctx.urlGravacao())}, "Endereço do fluxo RH_API_Actualizar",
    el("input", {type: "url", id: "nm-url", autocomplete: "off", placeholder: "https://…powerplatform.com/…&sig=…"}));
  function guardarUrlSePreciso() {
    const c = $("nm-url");
    if (c && !c.closest("label").hidden && c.value.trim()) ctx.guardarUrlGravacao(c.value);
    if (!ctx.urlGravacao()) throw Error("Falta o endereço do fluxo RH_API_Actualizar (peça-o ao administrador).");
  }
  const agoraSerie = () => { const a = new Date(); return Math.round(((a.getTime() - a.getTimezoneOffset() * 60000) / 86400000 + 25569) * 100000) / 100000; };

  function editar(original) {
    const {ano, mes} = ctx.periodo();
    const v = k => M.campo(original, k);
    const regime = String(v("RegimePagamento") || "");
    const hora = /^hora$/i.test(regime.trim());
    const fechado = F.lancamentoFinalizado(original) || F.lancamentoRetirado(original), pode = ctx.podeGravar();
    const t = abrir({eyebrow: "LANÇAMENTO · " + M.NOME_MES[mes - 1].toUpperCase() + " " + ano, titulo: M.normalizarCodigo(v("Codigo")) + " · " + (v("Nome") || ""),
      sub: [v("Empresa"), v("Departamento"), regime, v("Estado do Lançamento") || "—"].filter(Boolean).join("  ·  ")});
    const num = x => { const n = M.numero(x); return n == null ? "—" : fmt(n); };
    const entrada = c => {
      const atual = v(c.k);
      if (c.tipo === "simnao") return el("label", {}, "Canalizar", el("select", {id: "le-" + c.k, disabled: fechado || !pode},
        el("option", {value: "", text: "—"}), el("option", {value: "Sim", text: "Sim"}), el("option", {value: "Não", text: "Não"})));
      if (c.tipo === "texto") return el("label", {class: "span-2"}, c.k, el("textarea", {id: "le-" + c.k, rows: 2, value: atual ?? "", disabled: fechado || !pode}));
      const off = fechado || !pode || (c.soHora && !hora);
      return el("label", {}, c.rotulo || c.k, el("input", {id: "le-" + c.k, type: "text", inputMode: "decimal", value: off && c.soHora && !hora ? "" : (atual ?? ""), disabled: off,
        placeholder: c.soHora && !hora ? "Só regime Hora" : "0"}));
    };
    const grupo = (titulo, filtro) => el("fieldset", {class: "le-grupo"}, el("legend", {text: titulo}), ...F.CAMPOS_EDITAVEIS.filter(filtro).map(entrada));
    $("rel-body").replaceChildren(
      el("div", {class: "stats compact-stats rel-stats"},
        cartao("Dias trabalhados", num(v("Dias Trabalhados")), hora ? "regime Hora" : "automático (21–20)"), cartao("Dias de férias", num(v("Dias de Férias")), "automático"),
        cartao("Valor unitário", num(v("Valor Unitário")), "contrato/adenda ou Settings"), cartao("INSS colab.", num(v("INSS Colaborador")), "fórmula"),
        cartao("IRPS", num(v("IRPS")), "fórmula"), cartao("Líquido", num(v("Salario Liquido")), "fórmula")),
      fechado ? el("p", {class: "notice-inline", text: "Este lançamento está Finalizado: só consulta."}) :
        !pode ? el("p", {class: "notice-inline", text: "O seu perfil só permite consultar. Alterar exige Superadministrador ou Administrador RH."}) :
        el("p", {class: "help", text: "Ao gravar, o regime, o valor unitário, as datas do contrato e os dias (Mensal) são recalculados com o contrato e as adendas vigentes; o estado passa a «Alterado». Vencimento, INSS, IRPS e líquido continuam a ser calculados pelas fórmulas do Excel."}),
      el("div", {class: "le-form"},
        grupo("Presença", c => !c.grupo && c.tipo !== "texto"), grupo("Acréscimos", c => c.grupo === "acrescimo"), grupo("Descontos", c => c.grupo === "desconto"),
        grupo("Observações", c => c.tipo === "texto")),
      el("div", {id: "le-recalculo", role: "status", "aria-live": "polite"}),
      campoUrl());
    const canal = $("le-Canalizar"); if (canal) canal.value = ["Sim", "Não"].includes(String(v("Canalizar"))) ? String(v("Canalizar")) : "";
    let emCurso = false;
    const consultar = async () => {
      const [funcionarios, contratos] = await Promise.all([ctx.funcionarios(true), ctx.contratos(true)]);
      if (t !== tarefa) return null;
      if (!ctx.podeGravar()) throw Error("O seu perfil já não permite alterar lançamentos.");
      const chave = M.chaveRH(v("Empresa"), v("Codigo"));
      const encontrados = funcionarios.filter(f => M.chaveRH(M.campo(f, "Empresa"), M.campo(f, "Codigo")) === chave);
      if (encontrados.length !== 1) throw Error("É necessário um único colaborador para " + chave + ". Encontrados: " + encontrados.length + ".");
      return {original, funcionario: encontrados[0], adendas: contratos.adendas, parametros: contratos.parametros, ano, mes, utilizador: ctx.utilizador()};
    };
    const exclusivo = acao => async () => {
      if (emCurso) return;
      emCurso = true;
      try { await acao(); } catch (e) { if (t === tarefa) throw e; } finally { emCurso = false; }
    };
    const lista = [{texto: "Fechar", click: fechar}];
    if (!fechado && pode) lista.unshift({texto: "Recalcular colaborador", ocupado: "A recalcular…", click: exclusivo(async () => {
      $("le-recalculo").replaceChildren();
      const opcoes = await consultar();
      if (!opcoes) return;
      const r = F.prepararRecalculoLancamento(opcoes);
      if (r.erros.length) { $("le-recalculo").replaceChildren(); throw Error(r.erros.join(" ")); }
      const cp = r.condicoes;
      const horas = $("le-Horas Trabalhadas");
      horas.disabled = !/^hora$/i.test(cp.regime.trim());
      horas.placeholder = horas.disabled ? "Só regime Hora" : "0";
      const dt = d => d ? d.split("-").reverse().join("/") : "Sem data final";
      $("le-recalculo").replaceChildren(
        el("h3", {text: "Recálculo do colaborador — por gravar"}),
        tabela(["Condição", "No lançamento", "Actualizada"], [
          ["Regime", regime, cp.regime],
          ["Início do contrato", M.data(v("Data Inicio Contrato")) || "—", cp.inicioContrato ? dt(cp.inicioContrato) : "Não indicada"],
          ["Fim do contrato", M.data(v("Data Fim Contrato")) || "—", dt(cp.fimContrato)],
          ["Dias trabalhados", num(v("Dias Trabalhados")), cp.dias === null ? "Mantidos; dependem da presença registada" : num(cp.dias)],
          ["Valor unitário", num(v("Valor Unitário")), fmt(cp.valorUnitario) + " " + cp.unidade]]),
        el("p", {class: "help", text: "Período: " + dt(r.periodo.inicio) + " a " + dt(r.periodo.fim) + ". Origem do valor: " + cp.origem +
          ". Clique em «Gravar alteração» para aplicar a este colaborador, incluindo as entradas editadas acima. O vencimento, INSS, IRPS e líquido serão actualizados após a gravação. Os totais no topo ainda são os valores gravados."}));
    })});
    if (!fechado && pode) lista.unshift({texto: "Gravar alteração", classe: "primary", ocupado: "A gravar…", click: exclusivo(async () => {
      const opcoes = await consultar();
      if (!opcoes) return;
      const valores = Object.fromEntries(F.CAMPOS_EDITAVEIS.map(c => [c.k, $("le-" + c.k)?.value ?? ""]));
      const chave = M.chaveRH(v("Empresa"), v("Codigo"));
      const r = F.prepararEdicaoLancamento({...opcoes, valores});
      if (r.erros.length) throw Error(r.erros.join(" "));
      guardarUrlSePreciso();
      // O Excel Online demora ~15–20 s: a janela fecha já e a gravação continua em segundo plano,
      // com os segundos a contar na linha da tabela (ver gravarEmFundo em app.js).
      fechar();
      ctx.gravarEmFundo({ano, mes, chave, pedido: {ano, mes, mesNome: M.NOME_MES[mes - 1], linha: r.linha}});
    })});
    if (!fechado && pode) lista.splice(2, 0, {texto: "Voltar ao estado inicial", click: exclusivo(async () => {
      const entradasAntes = Object.fromEntries(F.CAMPOS_EDITAVEIS.map(c => [c.k, $("le-" + c.k)?.value ?? ""]));
      const reposicao = abrir({eyebrow: "REPOR LANÇAMENTO", titulo: "Voltar ao estado inicial", sub: v("Nome") + " · " + M.NOME_MES[mes - 1] + " " + ano});
      $("rel-body").replaceChildren(el("p", {text: "Esta operação substitui as entradas e condições de pagamento deste colaborador pelos valores iniciais guardados antes das alterações e repõe o estado Pendente. Inclui alterações já gravadas. As fórmulas serão recalculadas. Lançamentos Finalizados não podem ser repostos."}),
        el("p", {class: "help", text: "Se não existir uma referência inicial guardada, a reposição será recusada. Edições ainda não gravadas neste formulário serão descartadas."}), campoUrl());
      accoes([{texto: "Confirmar reposição", classe: "primary", ocupado: "A repor…", click: async () => {
        guardarUrlSePreciso();
        await ctx.reporInicial({ano, mes, Empresa: v("Empresa"), Codigo: M.normalizarCodigo(v("Codigo"))});
        if (reposicao !== tarefa) return;
        fechar(); ctx.aviso("Valores iniciais repostos. O lançamento voltou a Pendente.");
      }}, {texto: "Cancelar", click: () => { editar(original); for (const [k, valor] of Object.entries(entradasAntes)) if ($("le-" + k)) $("le-" + k).value = valor; }}]);
    })});
    accoes(lista);
  }

  function finalizar() {
    const {ano, mes} = ctx.periodo();
    const nome = M.NOME_MES[mes - 1] + "/" + ano;
    const linhas = ctx.lancamentosMes();
    const abertas = linhas.filter(l => !F.lancamentoFinalizado(l) && !F.lancamentoRetirado(l)).length;
    abrir({eyebrow: "FINALIZAR MÊS", titulo: "Finalizar " + nome, sub: linhas.length + " lançamento(s) · " + abertas + " por finalizar"});
    $("rel-body").replaceChildren(
      el("p", {text: abertas ? "Todos os lançamentos de " + nome + " passam a «Finalizado» e «Revisto = Sim». Depois de finalizados ficam apenas para consulta." : "Este período já está Finalizado."}),
      campoUrl());
    const pode = ctx.podeGravar();
    accoes([{texto: "Finalizar " + nome, classe: "primary", ocupado: "A finalizar…", desactivado: !abertas || !pode, titulo: !pode ? "Só Superadministrador ou Administrador RH" : "",
      click: async () => {
        guardarUrlSePreciso();
        const t0 = Date.now();
        const r = await ctx.finalizarMes({ano, mes, mesNome: M.NOME_MES[mes - 1], utilizador: ctx.utilizador(), agoraSerie: agoraSerie()});
        fechar(); ctx.aviso(r.alteradas + " lançamentos finalizados em " + nome + " (" + Math.round((Date.now() - t0) / 1000) + " s)."); ctx.recarregarMes(ano, mes);
      }}, {texto: "Fechar", click: fechar}]);
  }

  // ---------- Carregar mês: confirmar o que sai e o que entra ----------
  // Devolve "aplicar" (retirar + recalcular + introduzir), "recalcular" (só recalcular Pendentes) ou null (cancelar).
  function confirmarSincronizacao({ano, mes, retirar, introduzir, aRever = [], recalcular = 0}) {
    return new Promise(resolve => {
      let decidido = false;
      const decidir = v => { if (decidido) return; decidido = true; dlg.removeEventListener("close", aoFechar); fechar(); resolve(v); };
      const aoFechar = () => { if (!decidido) { decidido = true; resolve(null); } };
      const nome = M.NOME_MES[mes - 1] + "/" + ano;
      abrir({eyebrow: "CARREGAR MÊS", titulo: "Actualizar " + nome, sub: retirar.length + " a retirar · " + introduzir.length + " a introduzir · " + recalcular + " Pendente(s) a recalcular"});
      dlg.addEventListener("close", aoFechar);
      const bloco = [];
      bloco.push(el("p", {text: "O cadastro e os contratos mudaram desde que o mês foi criado. Confirme as alterações abaixo antes de as gravar."}));
      if (retirar.length) bloco.push(el("h3", {text: "A retirar (" + retirar.length + ")"}),
        el("p", {class: "help", text: "Lançamentos Pendentes de colaboradores sem contrato neste período. As linhas são apagadas do mês."}),
        tabela(["Empresa", "Código", "Nome", "Motivo"], retirar.map(x => [x.Empresa, x.Codigo, x.nome, x.motivo])));
      if (introduzir.length) bloco.push(el("h3", {text: "A introduzir (" + introduzir.length + ")"}),
        el("p", {class: "help", text: "Colaboradores com contrato vigente no período que ainda não têm lançamento. Entram como Pendentes, com as regras do Novo mês."}),
        tabela(["Empresa", "Código", "Nome", "Regime", "Origem"], introduzir.map(l => [l.Empresa, l.Codigo, l.Nome, l.RegimePagamento, l.Origem_Lancamento])));
      if (aRever.length) bloco.push(el("h3", {text: "A retirar — Alterados/Finalizados (" + aRever.length + ")"}),
        el("p", {class: "muted", text: "Lançamentos já Alterados ou Finalizados sem contrato neste período. As linhas são apagadas do mês, com as alterações feitas nelas."}),
        tabela(["Empresa", "Código", "Nome", "Estado actual", "Motivo"], aRever.map(x => [x.Empresa, x.Codigo, x.nome, x.estado, x.motivo])));
      $("rel-body").replaceChildren(...bloco);
      accoes([
        {texto: "Aplicar alterações", classe: "primary", click: () => decidir("aplicar")},
        {texto: "Só recalcular Pendentes", click: () => decidir("recalcular")},
        {texto: "Cancelar", click: () => decidir(null)}]);
    });
  }

  return {preVisualizar, exportarFolha, comparativo, novoMes, editar, finalizar, confirmarSincronizacao};
}
