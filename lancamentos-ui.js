// Lançamento salarial: pré-visualização da folha, exportação PDF/Excel, relatório comparativo e Novo mês.
import * as M from "./rh-modulos.js";
import * as F from "./rh-folha.js";
import * as X from "./rh-export.js";

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
  let tarefa = 0; // invalida pedidos antigos quando o diálogo é reaberto

  function abrir({eyebrow, titulo, sub = ""}) {
    $("rel-eyebrow").textContent = eyebrow; $("rel-titulo").textContent = titulo; $("rel-sub").textContent = sub;
    $("rel-body").replaceChildren(el("p", {class: "loading-line", text: "A preparar…"}));
    $("rel-accoes").replaceChildren(); $("rel-erro").textContent = "";
    if (!dlg.open) { if (typeof dlg.showModal === "function") dlg.showModal(); else dlg.setAttribute("open", ""); }
    return ++tarefa;
  }
  const fechar = () => { tarefa++; if (typeof dlg.close === "function") dlg.close(); else dlg.removeAttribute("open"); };
  $("rel-fechar").addEventListener("click", fechar);
  dlg.addEventListener("cancel", () => { tarefa++; });
  function accoes(lista) {
    $("rel-accoes").replaceChildren(el("span"), ...lista.map(a => {
      const b = el("button", {type: "button", class: a.classe || "outline", text: a.texto, disabled: Boolean(a.desactivado)});
      if (a.titulo) b.title = a.titulo;
      b.addEventListener("click", async () => {
        if (!a.click) return;
        const t = b.textContent; b.disabled = true; b.textContent = a.ocupado || "A gerar…"; $("rel-erro").textContent = "";
        try { await a.click(); } catch (e) { $("rel-erro").textContent = e.message; } finally { if (b.isConnected) { b.disabled = false; b.textContent = t; } }
      });
      return b;
    }));
  }
  const erro = e => { $("rel-body").replaceChildren(el("p", {class: "form-error", text: e.message || String(e)})); };

  // ---------- folha ----------
  function dadosFolha() {
    const {ano, mes} = ctx.periodo();
    const filtros = ctx.filtros();
    const linhas = F.linhasFolha(ctx.lancamentosMes(), ano, mes, filtros);
    if (!linhas.length) throw Error("Não existem lançamentos para os filtros seleccionados.");
    return {ano, mes, filtros, linhas};
  }
  const textoFiltro = f => "Empresa: " + (f.empresa || "TODOS") + "   |   Departamento: " + (f.departamento || "TODOS") + "   |   " + (f.codigo ? "Código: " + f.codigo : "Colaborador: TODOS");
  async function exportarFolha(tipo) {
    const d = dadosFolha(), opt = {ano: d.ano, mes: d.mes, filtros: d.filtros, utilizador: ctx.utilizador()};
    const blob = tipo === "pdf" ? await X.folhaPDF(d.linhas, opt) : await X.folhaExcel(d.linhas, opt);
    X.descarregar(blob, X.nomeFicheiroFolha(d.ano, d.mes, d.filtros, tipo));
    ctx.aviso("Folha exportada: " + X.nomeFicheiroFolha(d.ano, d.mes, d.filtros, tipo));
  }
  function preVisualizar() {
    let d;
    try { d = dadosFolha(); } catch (e) { ctx.aviso(e.message); return; }
    abrir({eyebrow: "PRÉ-VISUALIZAÇÃO", titulo: "Folha de Salários - " + M.NOME_MES[d.mes - 1] + " de " + d.ano, sub: textoFiltro(d.filtros)});
    const l = d.linhas, s = k => F.somar(l, k);
    $("rel-body").replaceChildren(
      el("div", {class: "stats compact-stats rel-stats"}, cartao("Colaboradores", String(l.length)), cartao("Vencimento base", fmt(s("base")), "MT"),
        cartao("Acréscimos", fmt(s("acrescimos")), "MT"), cartao("Descontos", fmt(s("totalDesc")), "MT"), cartao("Salário líquido", fmt(s("liquido")), "MT")),
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
    $("rel-body").replaceChildren(
      el("div", {class: "nm-escolha"}, el("label", {}, "Mês", selMes), el("label", {}, "Ano", selAno),
        el("p", {class: "help", text: "Período de folha 21–20. Entram os colaboradores com contrato vigente no período; as entradas do mês anterior (faltas, horas extra, subsídios, bónus, adiantamentos, empréstimos, outros descontos, dependentes, canalizar e observações) são copiadas; o valor unitário vem do contrato/adenda (Mensal) ou das Settings (Hora/Diário)."})),
      el("ul", {class: "criterios"},
        ...["O mês ainda não pode existir na tbl_lançamento.", "Colaboradores activos com Código e Empresa, sem pares repetidos.", "Contrato vigente no período (admissão, saída, prazo, estado e adendas activas).",
          "Regime de pagamento e valor unitário válidos para todos.", "Mensal: dias trabalhados base 30 (proporcionais na admissão ou saída).", "Estado «Pendente», Revisto «Não», origem «Cópia Mês Anterior» ou «Novo Colaborador»."].map(x => el("li", {text: x}))),
      resultado);
    const verificar = async () => {
      const ano = Number(selAno.value), mes = Number(selMes.value), ant = mes === 1 ? {ano: ano - 1, mes: 12} : {ano, mes: mes - 1};
      resultado.replaceChildren(el("p", {class: "loading-line", text: "A ler colaboradores, adendas, Settings e os lançamentos de " + M.NOME_MES[mes - 1] + " e " + M.NOME_MES[ant.mes - 1] + "… (pode demorar 1 a 2 minutos)"}));
      const [funcionarios, contratos, mesAlvo, mesAnt] = await Promise.all([ctx.funcionarios(), ctx.contratos(), ctx.carregarMes(ano, mes), ctx.carregarMes(ant.ano, ant.mes)]);
      if (t !== tarefa) return;
      const r = F.prepararNovoMes({ano, mes, funcionarios, adendas: contratos.adendas, parametros: contratos.parametros, lancamentosMes: mesAlvo, lancamentosAnterior: mesAnt, utilizador: ctx.utilizador()});
      mostrarVerificacao(r, ano, mes, t, () => { ctx.esquecerMes(ano, mes); ctx.esquecerMes(ant.ano, ant.mes); return verificar(); });
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
        click: () => criar(r, ano, mes, t)},
      {texto: "Fechar", click: fechar}]);
    if (!pode) $("rel-erro").textContent = "O seu perfil só permite consultar. Criar o mês exige Superadministrador ou Administrador RH.";
  }
  async function criar(r, ano, mes, t) {
    const campoUrl = $("nm-url");
    if (campoUrl && !campoUrl.closest("label").hidden && campoUrl.value.trim()) ctx.guardarUrlGravacao(campoUrl.value);
    if (!ctx.urlGravacao()) throw Error("Falta o endereço do fluxo RH_API_Actualizar (peça-o ao administrador).");
    // Confirmação final: verifica outra vez que o mês continua vazio.
    ctx.esquecerMes(ano, mes);
    const agora = await ctx.carregarMes(ano, mes);
    if (agora.length) throw Error("Entretanto já foram criados lançamentos para " + M.NOME_MES[mes - 1] + "/" + ano + ".");
    const res = await ctx.gravarMes({ano, mes, mesNome: r.mesNome, linhas: r.linhas});
    if (t !== tarefa) return;
    acompanhar(res.total, ano, mes, t);
  }

  async function acompanhar(total, ano, mes, t) {
    const barra = el("progress", {max: total, value: 0}), estado = el("p", {class: "loading-line", text: "Pedido aceite. O Power Automate está a gravar " + total + " linhas na tbl_lançamento…"});
    $("rel-body").replaceChildren(el("div", {class: "nm-progresso"}, el("h3", {text: "A criar " + M.NOME_MES[mes - 1] + " " + ano}), barra, estado,
      el("p", {class: "help", text: "Pode fechar esta janela: a gravação continua no Power Automate. Cada linha demora cerca de 1 segundo."})));
    accoes([{texto: "Fechar", click: fechar}]);
    const inicio = Date.now();
    while (Date.now() - inicio < 20 * 60000) {
      await new Promise(ok => setTimeout(ok, 30000));
      if (t !== tarefa) { ctx.esquecerMes(ano, mes); return; }
      ctx.esquecerMes(ano, mes);
      let linhas;
      try { linhas = (await ctx.carregarMes(ano, mes)).filter(l => F.linhasFolha([l], ano, mes).length); } catch { continue; }
      if (t !== tarefa) return;
      barra.value = Math.min(linhas.length, total);
      estado.textContent = linhas.length + " de " + total + " linhas gravadas…";
      if (linhas.length >= total) {
        const semFormula = linhas.filter(l => M.campo(l, "Salario Liquido") === "" || M.campo(l, "Vencimento") === "").length;
        estado.textContent = semFormula
          ? "Gravadas " + linhas.length + " linhas, mas " + semFormula + " ainda sem Vencimento/Salário líquido calculado. Abra o Excel e confirme que as colunas de fórmulas se estenderam às novas linhas."
          : "Concluído: " + linhas.length + " lançamentos criados com estado Pendente.";
        estado.className = semFormula ? "form-error" : "ok-line";
        ctx.mostrarMes(ano, mes);
        return;
      }
    }
    estado.textContent = "A gravação está a demorar mais do que o previsto. Consulte o histórico de execuções do RH_API_Actualizar.";
  }

  return {preVisualizar, exportarFolha, comparativo, novoMes};
}
