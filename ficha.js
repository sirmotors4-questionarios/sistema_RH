// Ficha do colaborador: reproduz o formulário G_Colaboradores do VBA.
// Modos: ver (tudo bloqueado) -> editar (Código e Empresa bloqueados) | novo. Guardar / Cancelar.
import * as M from "./rh-modulos.js";

const $ = id => document.getElementById(id);
const el = (tag, props = {}, ...filhos) => {
  const n = document.createElement(tag);
  for (const [k, v] of Object.entries(props)) {
    if (k === "class") n.className = v; else if (k === "text") n.textContent = v;
    else if (k.startsWith("data-")) n.setAttribute(k, v); else n[k] = v;
  }
  n.append(...filhos.filter(f => f != null));
  return n;
};
const iniciais = nome => (nome || "?").split(/\s+/).filter(Boolean).map(p => p[0]).slice(0, 2).join("").toUpperCase();
const NIVEL = {Baixo: "low", "Médio": "mid", Bom: "good"};

export function criarFicha(ctx) {
  const dlg = $("ficha");
  let modo = "ver", original = {}, separador = "pessoais", avisoFechar = false;

  // ---------- construção ----------
  const tabs = $("ficha-tabs"), corpo = $("ficha-body");
  M.SEPARADORES.forEach(s => {
    tabs.append(el("button", {type: "button", role: "tab", class: "ficha-tab", "data-sep": s.id, id: "tab-" + s.id},
      el("span", {text: s.nome}), el("small", {class: "pct", "data-pct": s.id})));
    const painel = el("section", {class: "ficha-panel", role: "tabpanel", "data-panel": s.id});
    painel.setAttribute("aria-labelledby", "tab-" + s.id);
    if (s.id === "contrato") painel.append(blocoContrato());
    else {
      const grelha = el("div", {class: "ficha-grid"});
      M.CAMPOS_COLABORADOR.filter(c => c.sep === s.id).forEach(c => grelha.append(campoHtml(c)));
      painel.append(grelha);
      if (s.id === "profissionais") painel.append(el("p", {class: "pay-note", id: "pay-note"}));
    }
    corpo.append(painel);
  });
  tabs.addEventListener("click", ev => { const b = ev.target.closest("[data-sep]"); if (b) mostrar(b.dataset.sep); });
  tabs.addEventListener("keydown", ev => {
    if (!["ArrowRight", "ArrowLeft"].includes(ev.key)) return;
    const ids = M.SEPARADORES.map(s => s.id), i = ids.indexOf(separador);
    const novo = ids[(i + (ev.key === "ArrowRight" ? 1 : ids.length - 1)) % ids.length];
    mostrar(novo); tabs.querySelector('[data-sep="' + novo + '"]').focus();
  });

  function campoHtml(c) {
    const id = "f-" + c.k.replace(/[^a-z0-9]/gi, "_");
    const rot = el("span", {text: c.rotulo, class: c.obrig ? "req" : ""});
    let inp;
    if (c.tipo === "escolha") {
      inp = el("select", {id});
      inp.append(el("option", {value: "", text: "—"}), ...c.opcoes.map(o => el("option", {value: o, text: o})));
    } else {
      inp = el("input", {id, type: c.tipo === "data" ? "date" : c.tipo === "email" ? "email" : "text", autocomplete: "off"});
      if (c.tipo === "numero") inp.inputMode = "decimal";
      if (c.max) inp.maxLength = c.max;
      if (c.tipo === "lista") { inp.setAttribute("list", id + "-l"); }
      if (c.ajuda) inp.placeholder = c.ajuda;
    }
    inp.dataset.campo = c.k;
    const label = el("label", {class: "fld" + (c.largo ? " wide" : "")}, rot, inp);
    if (c.tipo === "lista") label.append(el("datalist", {id: id + "-l"}));
    if (c.chave) label.append(el("small", {class: "hint", "data-hint-chave": "1", text: "Identifica o colaborador (Empresa + Código)."}));
    return label;
  }

  function blocoContrato() {
    const frag = el("div", {class: "contract"});
    frag.append(
      el("div", {class: "contract-card"}, el("h3", {text: "Contrato actual"}), el("dl", {id: "contrato-actual"})),
      el("div", {class: "contract-card"}, el("h3", {text: "Histórico de adendas"}),
        el("p", {class: "help", text: "As alterações contratuais (vencimento, departamento, tipo de contrato, saída…) ficam registadas como adendas. Uma adenda nunca é apagada: é anulada."}),
        el("div", {class: "table-scroll"}, el("table", {class: "mini"},
          el("thead", {}, el("tr", {}, ...["N.º", "Data", "Início", "Campo", "Anterior", "Novo", "Estado"].map(t => el("th", {text: t})))),
          el("tbody", {id: "ficha-adendas"})))));
    return frag;
  }

  // ---------- estado ----------
  const inputs = () => [...dlg.querySelectorAll("[data-campo]")];
  const valores = () => Object.fromEntries(inputs().map(i => [i.dataset.campo, i.value]));
  const sujo = () => modo !== "ver" && inputs().some(i => String(original[i.dataset.campo] ?? "") !== i.value);

  function mostrar(sep) {
    separador = sep;
    tabs.querySelectorAll("[data-sep]").forEach(b => { const on = b.dataset.sep === sep; b.classList.toggle("active", on); b.setAttribute("aria-selected", String(on)); b.tabIndex = on ? 0 : -1; });
    corpo.querySelectorAll("[data-panel]").forEach(p => p.hidden = p.dataset.panel !== sep);
  }

  function preencherListas() {
    const linhas = ctx.funcionarios();
    M.CAMPOS_COLABORADOR.filter(c => c.tipo === "lista").forEach(c => {
      const dl = $("f-" + c.k.replace(/[^a-z0-9]/gi, "_") + "-l");
      dl.replaceChildren(...M.valoresDistintos(linhas, c).map(v => el("option", {value: v})));
    });
    // Em "escolha", um valor antigo fora da lista (ex.: "M") aparece para não se perder.
    M.CAMPOS_COLABORADOR.filter(c => c.tipo === "escolha").forEach(c => {
      const s = $("f-" + c.k.replace(/[^a-z0-9]/gi, "_"));
      s.querySelectorAll("[data-legado]").forEach(o => o.remove());
      const v = original[c.k];
      if (v && !c.opcoes.includes(v)) s.append(el("option", {value: v, text: v + " (valor antigo)", "data-legado": "1"}));
    });
  }

  function aplicarModo(novo) {
    modo = novo; avisoFechar = false;
    const ver = modo === "ver";
    inputs().forEach(i => {
      const c = M.definicao(i.dataset.campo);
      i.disabled = ver || (modo === "editar" && c.chave);
    });
    dlg.dataset.modo = modo;
    $("ficha-modo").textContent = {ver: "Visualização", editar: "Em edição", novo: "Novo colaborador"}[modo];
    $("ficha-editar").hidden = !ver || !ctx.podeGravar();
    $("ficha-guardar").hidden = ver;
    $("ficha-cancelar").textContent = ver ? "Fechar" : "Cancelar";
    $("ficha-erro").textContent = "";
    $("ficha-url").hidden = ver || Boolean(ctx.urlGravacao());
    dlg.querySelectorAll("[data-hint-chave]").forEach(h => h.hidden = modo !== "editar");
    regime();
    indicadores();
  }

  function cabecalho() {
    const v = valores();
    $("ficha-avatar").textContent = iniciais(v.Nome);
    $("ficha-nome").textContent = v.Nome || (modo === "novo" ? "Novo colaborador" : "—");
    $("ficha-id").textContent = modo === "novo" ? "Preencha o Código, a Empresa e o Nome (obrigatórios)." : "Código " + (original.Codigo || "") + " · " + (original.Empresa || "");
    const est = $("ficha-estado");
    est.hidden = !v.Estado; est.textContent = v.Estado; est.classList.toggle("off", !M.activo(v.Estado));
  }

  function indicadores() {
    const v = valores();
    M.SEPARADORES.forEach(s => {
      const chip = tabs.querySelector('[data-pct="' + s.id + '"]');
      if (s.id === "contrato") { chip.textContent = ""; return; }
      const p = M.preenchimento(v, s.id);
      chip.textContent = p.pct + "%"; chip.className = "pct " + NIVEL[p.nivel];
      chip.title = p.preenchidos + " de " + p.total + " campos · " + p.nivel;
    });
    const g = M.preenchimento(v);
    const anel = $("ficha-geral");
    anel.style.setProperty("--p", g.pct);
    anel.className = "ring " + NIVEL[g.nivel];
    anel.querySelector("b").textContent = g.pct + "%";
    anel.title = "Preenchimento geral: " + g.preenchidos + " de " + g.total + " campos (" + g.nivel + "). 0–49% Baixo · 50–79% Médio · 80–100% Bom";
    cabecalho();
  }

  // Regime Hora/Diário: sem vencimento base; mostra o valor unitário vigente das Settings.
  function regime() {
    const v = valores(), base = $("f-Vencimento_Base"), nota = $("pay-note");
    const semBase = M.regimeSemBase(v.RegimePagamento);
    if (semBase) { base.value = ""; base.disabled = true; base.placeholder = "N/A"; }
    else { base.disabled = modo === "ver"; base.placeholder = ""; }
    const unidade = {mensal: "MT/mês", hora: "MT/hora", diario: "MT/dia"}[(v.RegimePagamento || "").normalize("NFD").replace(/[̀-ͯ]/g, "").toLowerCase()];
    if (!v.RegimePagamento) { nota.textContent = ""; return; }
    if (!semBase) { nota.textContent = "Regime Mensal: o vencimento base (" + unidade + ") é o valor contratual; uma adenda activa e vigente prevalece."; return; }
    const hoje = new Date().toLocaleDateString("en-CA", {timeZone: "Africa/Maputo"});
    const params = ctx.parametros();
    if (!params) { nota.textContent = "Regime " + v.RegimePagamento + ": o valor unitário vem das Settings de pagamento (" + unidade + ")."; return; }
    const prof = (v.Profissao || "").toLowerCase();
    const vig = M.parametrosVigentes(params, hoje).filter(p => M.regimeSemBase(p.regime) && p.regime.normalize("NFD").replace(/[̀-ͯ]/g, "").toLowerCase() === v.RegimePagamento.normalize("NFD").replace(/[̀-ͯ]/g, "").toLowerCase());
    const p = vig.find(x => x.categoria.toLowerCase() === prof) || vig.find(x => !x.categoria);
    nota.textContent = p ? "Valor unitário vigente (Settings): " + M.formatarValor(p.valor) + " " + (p.unidade || unidade) + (p.categoria ? " · " + p.categoria : "") + "."
      : "Sem valor unitário vigente nas Settings para " + v.RegimePagamento + (v.Profissao ? " / " + v.Profissao : "") + ". Sem ele o mês não pode ser lançado.";
  }

  function contrato() {
    const v = original, dl = $("contrato-actual");
    const linhas = [["Tipo de contrato", v.TipoContrato], ["Empresa", v.Empresa], ["Departamento", v.Departamento], ["Profissão", v.Profissao],
      [M.regimeSemBase(v.RegimePagamento) ? "Regime" : "Vencimento base", M.regimeSemBase(v.RegimePagamento) ? v.RegimePagamento : (v["Vencimento Base"] !== "" && v["Vencimento Base"] != null ? M.formatarValor(v["Vencimento Base"]) + " MT/mês" : "")],
      ["Data de admissão", M.formatarData(v["Data Admissao"])], ["Estado", v.Estado]];
    dl.replaceChildren(...linhas.flatMap(([k, x]) => [el("dt", {text: k}), el("dd", {text: x || "—"})]));
    const body = $("ficha-adendas"); body.replaceChildren();
    const linha = txt => { const tr = el("tr"); tr.append(el("td", {colSpan: 7, class: "no-data", text: txt})); body.append(tr); };
    if (modo === "novo") return linha("As adendas ficam disponíveis depois de criar o colaborador.");
    const ad = ctx.adendas();
    if (!ad) { linha("A carregar adendas…"); return; }
    const k = M.chaveRH(v.Empresa, v.Codigo);
    const lista = M.adendas(ad).filter(a => M.chaveRH(a.empresa || v.Empresa, a.codigo) === k);
    if (!lista.length) return linha("Sem adendas registadas para este colaborador.");
    lista.forEach(a => body.append(el("tr", {}, ...[a.numero || a.id, M.formatarData(a.data), M.formatarData(a.inicio), a.campo, a.anterior, a.novo, a.estado].map(t => el("td", {text: t || "—"})))));
  }

  // ---------- abrir / fechar ----------
  function abrir(novoModo, linha) {
    original = linha ? M.linhaParaFormulario(linha) : {};
    preencherListas();
    inputs().forEach(i => { i.value = original[i.dataset.campo] ?? ""; });
    if (novoModo === "novo") { $("f-Estado").value = "Activo"; $("f-RegimePagamento").value = "Mensal"; }
    aplicarModo(novoModo);
    mostrar(novoModo === "novo" ? "profissionais" : "pessoais");
    contrato();
    if (!ctx.adendas() || !ctx.parametros()) ctx.carregarContratos().then(() => { if (dlg.open) { contrato(); regime(); } }).catch(() => {});
    if (typeof dlg.showModal === "function") dlg.showModal(); else dlg.setAttribute("open", "");
    (novoModo === "novo" ? $("f-Codigo") : $("ficha-cancelar")).focus();
  }
  function fechar(forcar = false) {
    if (!forcar && sujo() && !avisoFechar) {
      avisoFechar = true;
      $("ficha-erro").textContent = "Tem alterações por gravar. Carregue de novo para as descartar.";
      return false;
    }
    if (typeof dlg.close === "function") dlg.close(); else dlg.removeAttribute("open");
    return true;
  }

  dlg.addEventListener("input", ev => { if (!ev.target.dataset.campo) return; avisoFechar = false; if (["RegimePagamento", "Profissao"].includes(ev.target.dataset.campo)) regime(); indicadores(); });
  dlg.addEventListener("change", ev => { if (ev.target.dataset.campo === "RegimePagamento") { regime(); indicadores(); } });
  dlg.addEventListener("cancel", ev => { if (!fechar()) ev.preventDefault(); });
  $("ficha-fechar").addEventListener("click", () => fechar());
  $("ficha-cancelar").addEventListener("click", () => {
    if (modo === "editar") {
      if (sujo() && !avisoFechar) { avisoFechar = true; $("ficha-erro").textContent = "Tem alterações por gravar. Carregue de novo em Cancelar para as descartar."; return; }
      inputs().forEach(i => { i.value = original[i.dataset.campo] ?? ""; }); aplicarModo("ver"); return;
    }
    fechar();
  });
  $("ficha-editar").addEventListener("click", () => { aplicarModo("editar"); mostrar(separador === "contrato" ? "pessoais" : separador); });
  $("ficha-mudar-url").addEventListener("click", () => { $("ficha-url").hidden = false; $("ficha-url-input").focus(); });

  $("ficha-form").addEventListener("submit", async ev => {
    ev.preventDefault();
    if (modo === "ver") return;
    const erro = $("ficha-erro"); erro.textContent = "";
    const entrada = valores();
    if (modo === "editar") { entrada.Codigo = original.Codigo; entrada.Empresa = original.Empresa; }
    const {dados, erros} = M.validarColaborador(entrada, modo === "novo" ? "criar" : "editar", {original, existentes: ctx.funcionarios()});
    if (erros.length) {
      erro.textContent = erros.join(" ");
      const primeiro = M.CAMPOS_COLABORADOR.find(c => erros.some(e => e.startsWith(c.rotulo)));
      if (primeiro) { mostrar(primeiro.sep); $("f-" + primeiro.k.replace(/[^a-z0-9]/gi, "_")).focus(); }
      return;
    }
    const pedido = M.prepararGravacao(modo === "novo" ? "criar" : "editar", dados, {original, cabecalhos: M.colunas(ctx.funcionarios()), utilizador: ctx.utilizador()});
    if (modo === "editar" && !pedido.alteracoes) { erro.textContent = pedido.ignorados.length ? "Os campos alterados não existem na tbl_funcionarios: " + pedido.ignorados.join(", ") + "." : "Não alterou nenhum campo."; return; }
    const btn = $("ficha-guardar"); btn.disabled = true; btn.textContent = "A gravar…";
    try {
      if (!$("ficha-url").hidden && $("ficha-url-input").value.trim()) ctx.guardarUrlGravacao($("ficha-url-input").value);
      if (!ctx.urlGravacao()) throw Error("Falta o endereço do fluxo RH_API_Actualizar (peça-o ao administrador).");
      const r = await ctx.gravar(modo === "novo" ? "criar" : "editar", pedido);
      fechar(true);
      ctx.aviso((modo === "novo" ? "Colaborador criado: " : "Alterações gravadas: ") + r.codigo + " · " + r.empresa +
        (pedido.ignorados.length ? ". Não gravados (coluna inexistente): " + pedido.ignorados.join(", ") : "") + ". A actualizar a lista…");
      await ctx.recarregar();
      const linha = M.localizarColaborador(ctx.funcionarios(), r.empresa, r.codigo);
      if (linha) abrir("ver", linha);
    } catch (e) {
      erro.textContent = e.code === "RH_CONFLITO" ? e.message + " (o mesmo código só pode repetir-se noutra empresa)" : e.message;
    } finally { btn.disabled = false; btn.textContent = "Guardar"; }
  });

  return {
    ver: linha => abrir("ver", linha),
    novo: () => abrir("novo", null),
    aberta: () => dlg.open,
    actualizar: () => { if (dlg.open) { contrato(); regime(); } }
  };
}
