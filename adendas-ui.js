// Adendas: registar uma adenda (uma linha por campo em tbl_adendas), aplicá-la ao cadastro e anulá-la.
// A escrita é feita pelo Office Script (accões registar / aplicar / revogar), chamado pelo fluxo RH_API_Indemnizacoes.
import * as M from "./rh-modulos.js";

const $ = id => document.getElementById(id);
const el = (tag, props = {}, ...filhos) => {
  const n = document.createElement(tag);
  for (const [k, v] of Object.entries(props)) {
    if (v == null || v === false) continue;
    if (k === "class") n.className = v; else if (k === "text") n.textContent = v;
    else if (k.startsWith("data-") || k.startsWith("aria-") || k === "list") n.setAttribute(k, v); else n[k] = v;
  }
  for (const f of filhos.flat(Infinity)) if (f != null && f !== false) n.append(f);
  return n;
};
const fmt = v => M.formatarValor(v);
const dataPt = iso => iso ? iso.slice(8, 10) + "/" + iso.slice(5, 7) + "/" + iso.slice(0, 4) : "—";
const hojeIso = () => new Date().toLocaleDateString("en-CA", {timeZone: "Africa/Maputo"});

// Campos que uma adenda pode alterar (cabeçalhos reais da tbl_funcionarios).
export const CAMPOS = [
  {k: "Vencimento Base", rotulo: "Vencimento base", tipo: "numero"},
  {k: "RegimePagamento", rotulo: "Regime de pagamento", tipo: "escolha", opcoes: ["Mensal", "Diário", "Hora"]},
  {k: "TipoContrato", rotulo: "Tipo de contrato", tipo: "lista"},
  {k: "Departamento", rotulo: "Departamento", tipo: "lista"},
  {k: "Profissao", rotulo: "Profissão", tipo: "lista"},
  {k: "Categoria", rotulo: "Categoria", tipo: "lista"},
  {k: "SituacaoProfissional", rotulo: "Situação profissional", tipo: "lista"}
];
export const MOTIVOS = ["Alteração Salarial", "Actualização Salário Mínimo", "Mudança de Departamento", "Mudança de Profissão",
  "Mudança de Regime", "Renovação / alteração de contrato", "Promoção", "Outro"];

// Agrupa as linhas da tbl_adendas por adenda (Empresa + Código + N.º): cada campo alterado é uma linha.
export function agrupar(lista) {
  const m = new Map();
  for (const a of lista) {
    const k = [a.empresa.toUpperCase(), M.normalizarCodigo(a.codigo), a.numero || a.id].join("|");
    if (!m.has(k)) m.set(k, {chave: k, empresa: a.empresa, codigo: a.codigo, nome: a.nome, numero: a.numero, inicio: a.inicio, fim: a.fim,
      data: a.data, motivo: a.motivo, estado: a.estado, linhas: []});
    m.get(k).linhas.push(a);
  }
  return [...m.values()].map(g => ({...g, anulada: g.linhas.every(l => /anulad/i.test(l.estado)),
    pendente: g.linhas.some(l => !/anulad/i.test(l.estado) && !/^sim$/i.test(l.aplicada)), ids: g.linhas.map(l => l.id).filter(Boolean)}));
}

export function validarAdenda({colaborador, motivo, dataAdenda, inicio, fim, alteracoes}) {
  const erros = [];
  if (!colaborador) erros.push("Escolha o colaborador.");
  if (!motivo) erros.push("Indique o motivo da adenda.");
  if (!M.data(dataAdenda)) erros.push("Indique a data da adenda.");
  if (!M.data(inicio)) erros.push("Indique a data de início de vigência.");
  if (fim && (!M.data(fim) || fim < inicio)) erros.push("A data de fim é inválida ou anterior ao início.");
  if (!alteracoes.length) erros.push("Acrescente pelo menos uma alteração.");
  const vistos = new Set();
  for (const a of alteracoes) {
    const c = CAMPOS.find(x => x.k === a.campo);
    if (!c) { erros.push("Escolha o campo a alterar."); continue; }
    if (vistos.has(a.campo)) erros.push(c.rotulo + " aparece duas vezes."); vistos.add(a.campo);
    if (!String(a.novo ?? "").trim()) erros.push(c.rotulo + ": indique o novo valor.");
    else if (c.tipo === "numero" && !(M.numero(a.novo) > 0)) erros.push(c.rotulo + ": valor inválido.");
    else if (String(a.novo).trim() === String(a.anterior ?? "").trim()) erros.push(c.rotulo + ": o novo valor é igual ao actual.");
  }
  return erros;
}

export function iniciarAdendas(ctx) {
  const dlg = $("ad-dlg");
  function abrir({eyebrow, titulo, sub = ""}) {
    $("ad-dlg-eyebrow").textContent = eyebrow; $("ad-dlg-titulo").textContent = titulo; $("ad-dlg-sub").textContent = sub;
    $("ad-dlg-erro").textContent = ""; $("ad-dlg-accoes").replaceChildren();
    if (!dlg.open) { if (typeof dlg.showModal === "function") dlg.showModal(); else dlg.setAttribute("open", ""); }
  }
  const fechar = () => { if (typeof dlg.close === "function") dlg.close(); else dlg.removeAttribute("open"); };
  $("ad-dlg-fechar").addEventListener("click", fechar);
  function accoes(lista) {
    $("ad-dlg-accoes").replaceChildren(el("span"), ...lista.map(a => {
      const b = el("button", {type: "button", class: a.classe || "outline", text: a.texto});
      b.addEventListener("click", async () => {
        const t = b.textContent, t0 = Date.now();
        b.disabled = true; b.textContent = a.ocupado || "A gravar…"; $("ad-dlg-erro").textContent = "";
        const relogio = setInterval(() => { if (b.isConnected) b.textContent = (a.ocupado || "A gravar…") + " " + Math.round((Date.now() - t0) / 1000) + " s"; }, 1000);
        try { await a.click(); } catch (e) { $("ad-dlg-erro").textContent = e.message || String(e); }
        finally { clearInterval(relogio); if (b.isConnected) { b.disabled = false; b.textContent = t; } }
      });
      return b;
    }));
  }
  function pronto() {
    if (!ctx.ligado()) { ctx.aviso("Inicie sessão para gerir adendas."); return false; }
    if (!ctx.podeGravar()) { ctx.aviso("O seu perfil permite consultar, mas não gravar adendas."); return false; }
    if (!ctx.temUrl()) { ctx.aviso("Para gravar adendas, abra uma vez a página Indemnizações e cole o endereço do fluxo RH_API_Indemnizacoes."); return false; }
    return true;
  }
  const campo = (rot, input, ajuda) => el("label", {class: "ind-campo"}, rot, input, ajuda ? el("small", {text: ajuda}) : null);

  function colaboradores() {
    return (ctx.funcionarios() || []).map(l => ({linha: l, empresa: String(M.campo(l, "Empresa") || "").trim(), codigo: M.normalizarCodigo(M.campo(l, "Codigo")),
      nome: String(M.campo(l, "Nome") || "").trim(), estado: String(M.campo(l, "Estado") || "").trim()})).filter(c => c.nome && c.codigo);
  }
  function valoresDistintos(k) {
    return [...new Set((ctx.funcionarios() || []).map(l => String(M.campo(l, k) ?? "").trim()).filter(Boolean))].sort((a, b) => a.localeCompare(b, "pt"));
  }

  // ---------- nova adenda ----------
  function nova(pre) {
    if (!pronto()) return;
    const lista = colaboradores(), hoje = hojeIso();
    const st = {c: null, alts: []};
    abrir({eyebrow: "NOVA ADENDA", titulo: "Nova adenda contratual", sub: "As alterações ficam registadas na tbl_adendas e, se a vigência já começou, são aplicadas ao cadastro."});
    const dl = el("datalist", {id: "ad-colabs"}, ...lista.map(c => el("option", {value: c.codigo + " · " + c.nome + " · " + c.empresa, label: c.estado})));
    const f = {
      colab: el("input", {id: "ad-f-colab", list: "ad-colabs", placeholder: "Código ou nome…", autocomplete: "off"}),
      motivo: el("select", {id: "ad-f-motivo"}, el("option", {value: "", text: "Escolha o motivo…"}), ...MOTIVOS.map(m => el("option", {value: m, text: m}))),
      outro: el("input", {id: "ad-f-outro", maxLength: 100, placeholder: "Descreva o motivo"}),
      data: el("input", {id: "ad-f-data", type: "date", value: hoje}),
      inicio: el("input", {id: "ad-f-inicio", type: "date", value: hoje}),
      fim: el("input", {id: "ad-f-fim", type: "date"}),
      doc: el("input", {id: "ad-f-doc", maxLength: 120, placeholder: "Ex.: Adenda n.º 2 assinada em …"}),
      obs: el("input", {id: "ad-f-obs", maxLength: 250, placeholder: "Opcional"}),
      aplicar: el("input", {id: "ad-f-aplicar", type: "checkbox", checked: true})
    };
    const lOutro = campo("Outro motivo", f.outro); lOutro.hidden = true;
    f.motivo.addEventListener("change", () => { lOutro.hidden = f.motivo.value !== "Outro"; });
    const resumo = el("div", {class: "ad-resumo"}), altsBox = el("div", {class: "ad-alts"});
    const notaAplicar = el("small", {class: "ad-nota"});
    const actualizarAplicar = () => {
      const futuro = f.inicio.value && f.inicio.value > hojeIso();
      f.aplicar.disabled = futuro; if (futuro) f.aplicar.checked = false;
      notaAplicar.textContent = futuro ? "A vigência começa em " + dataPt(f.inicio.value) + ": a adenda fica pendente e aplica-se depois (botão «Aplicar ao cadastro»)." : "";
    };
    f.inicio.addEventListener("input", actualizarAplicar);
    const actual = k => st.c ? M.campo(st.c.linha, k) : "";
    const desenharAlts = () => {
      altsBox.replaceChildren(el("h3", {text: "Alterações"}), ...st.alts.map((a, i) => {
        const def = CAMPOS.find(x => x.k === a.campo);
        const sel = el("select", {"aria-label": "Campo a alterar"}, el("option", {value: "", text: "Campo…"}),
          ...CAMPOS.map(c => el("option", {value: c.k, text: c.rotulo, selected: c.k === a.campo})));
        sel.addEventListener("change", () => { a.campo = sel.value; a.novo = ""; a.anterior = String(actual(a.campo) ?? ""); desenharAlts(); });
        let inp;
        if (def?.tipo === "escolha") { inp = el("select", {}, el("option", {value: "", text: "Novo valor…"}), ...def.opcoes.map(o => el("option", {value: o, text: o, selected: o === a.novo}))); }
        else {
          const lid = "ad-lista-" + i;
          inp = el("input", {value: a.novo || "", placeholder: "Novo valor", type: def?.tipo === "numero" ? "number" : "text", step: "0.01", min: "0", list: def?.tipo === "lista" ? lid : null});
          if (def?.tipo === "lista") inp = el("span", {class: "ad-inp"}, inp, el("datalist", {id: lid}, ...valoresDistintos(def.k).map(v => el("option", {value: v}))));
        }
        const input = inp.tagName === "SPAN" ? inp.querySelector("input") : inp;
        const dif = el("small", {class: "ad-dif"});
        const calc = () => { a.novo = input.value; dif.textContent = def?.tipo === "numero" && M.numero(a.anterior) !== null && M.numero(a.novo) !== null ? "Diferença: " + (M.numero(a.novo) - M.numero(a.anterior) >= 0 ? "+" : "") + fmt(M.numero(a.novo) - M.numero(a.anterior)) : ""; };
        input.addEventListener("input", calc); input.addEventListener("change", calc); calc();
        const rm = el("button", {type: "button", class: "icon-btn", text: "✕", title: "Retirar alteração", "aria-label": "Retirar alteração"});
        rm.addEventListener("click", () => { st.alts.splice(i, 1); desenharAlts(); });
        return el("div", {class: "ad-alt"}, sel,
          el("div", {class: "ad-ant"}, el("span", {text: "Actual"}), el("b", {text: def ? (def.tipo === "numero" ? fmt(a.anterior) : (a.anterior || "—")) : "—"})),
          el("div", {class: "ad-novo"}, inp, dif), rm);
      }), (() => { const b = el("button", {type: "button", class: "link-btn", text: "+ Acrescentar alteração"}); b.addEventListener("click", () => { st.alts.push({campo: "", novo: "", anterior: ""}); desenharAlts(); }); return b; })());
    };
    const escolher = () => {
      const [cod, , emp] = f.colab.value.split("·").map(x => (x || "").trim());
      const c = lista.find(x => x.codigo === M.normalizarCodigo(cod) && (!emp || x.empresa === emp)) || null;
      if (c === st.c && st.desenhado) return; // mesmo colaborador: não redesenha (evita perder o clique seguinte)
      st.c = c; st.desenhado = true;
      st.alts.forEach(a => { a.anterior = String(actual(a.campo) ?? ""); });
      resumo.replaceChildren(st.c ? el("p", {class: "rel-nota", text: st.c.nome + " · " + st.c.codigo + " · " + st.c.empresa + " · " + (st.c.estado || "—") +
        " · Regime: " + (M.campo(st.c.linha, "RegimePagamento") || "—") + " · Vencimento base: " + fmt(M.campo(st.c.linha, "Vencimento Base"))}) : "");
      desenharAlts();
    };
    f.colab.addEventListener("change", escolher);
    $("ad-dlg-body").replaceChildren(dl,
      el("div", {class: "ind-form"}, campo("Colaborador", f.colab), campo("Motivo", f.motivo), lOutro, campo("Data da adenda", f.data),
        campo("Início de vigência", f.inicio), campo("Fim de vigência", f.fim, "Opcional."), campo("Documento de referência", f.doc), campo("Observações", f.obs)),
      resumo, altsBox,
      el("label", {class: "ind-confirma"}, f.aplicar, " Aplicar já ao cadastro (tbl_funcionarios)"), notaAplicar);
    st.alts.push({campo: "", novo: "", anterior: ""});
    if (pre) { f.colab.value = pre.codigo + " · " + pre.nome + " · " + pre.empresa; escolher(); }
    desenharAlts(); actualizarAplicar();
    accoes([{texto: "Cancelar", click: async () => fechar()}, {texto: "Registar adenda", classe: "primary", click: async () => {
      const motivo = f.motivo.value === "Outro" ? f.outro.value.trim() : f.motivo.value;
      const alts = st.alts.filter(a => a.campo || a.novo);
      const erros = validarAdenda({colaborador: st.c, motivo, dataAdenda: f.data.value, inicio: f.inicio.value, fim: f.fim.value, alteracoes: alts});
      if (erros.length) throw Error(erros.join(" "));
      const r = await ctx.registar({empresa: st.c.empresa, codigo: st.c.codigo, motivo, dataAdenda: f.data.value, inicio: f.inicio.value, fim: f.fim.value || "",
        documento: f.doc.value.trim(), observacoes: f.obs.value.trim(), aplicar: f.aplicar.checked,
        alteracoes: alts.map(a => ({campo: a.campo, novo: CAMPOS.find(c => c.k === a.campo)?.tipo === "numero" ? String(M.numero(a.novo)) : a.novo.trim()}))});
      fechar();
      ctx.aviso("Adenda n.º " + r.numero + " registada para " + st.c.nome + (r.aplicada ? " e aplicada ao cadastro." : " (pendente: aplica-se quando a vigência começar)."));
      await ctx.recarregar();
    }}]);
    f.colab.focus();
  }

  // ---------- detalhe de uma adenda ----------
  function detalhe(g) {
    abrir({eyebrow: "ADENDA N.º " + (g.numero || "—"), titulo: g.nome + " · " + g.empresa,
      sub: (g.motivo || "Sem motivo") + " · vigência " + dataPt(g.inicio) + (g.fim ? " a " + dataPt(g.fim) : "") + " · " + (g.anulada ? "Anulada" : g.pendente ? "Pendente de aplicação" : "Aplicada")});
    $("ad-dlg-body").replaceChildren(el("div", {class: "table-scroll"}, el("table", {class: "rel-table"},
      el("thead", {}, el("tr", {}, ...["ID", "Campo", "Anterior", "Novo", "Estado", "Aplicada"].map(t => el("th", {text: t})))),
      el("tbody", {}, ...g.linhas.map(l => el("tr", {}, el("td", {text: l.id}), el("td", {text: l.campo}), el("td", {text: l.anterior || "—"}), el("td", {text: l.novo || "—"}),
        el("td", {text: l.estado || "—"}), el("td", {text: l.aplicada || "—"})))))),
      g.linhas[0]?.documento ? el("p", {class: "rel-nota", text: "Documento: " + g.linhas[0].documento}) : null,
      g.anulada ? null : el("div", {class: "ind-form ind-form-1"}, campo("Motivo da anulação (se for anular)", el("input", {id: "ad-f-anular", maxLength: 150, placeholder: "Ex.: registada por engano"}))),
      g.anulada ? null : el("p", {class: "ad-nota", text: "Anular não repõe o cadastro: para voltar ao valor anterior, registe uma nova adenda."}));
    const lista = [];
    if (!g.anulada) lista.push({texto: "Anular adenda", ocupado: "A anular…", click: async () => {
      if (!pronto()) return;
      const motivo = $("ad-f-anular").value.trim();
      if (!motivo) throw Error("Indique o motivo da anulação.");
      await ctx.revogar({ids: g.ids, motivo}); fechar(); ctx.aviso("Adenda n.º " + g.numero + " anulada."); await ctx.recarregar();
    }});
    if (!g.anulada && g.pendente) lista.push({texto: "Aplicar ao cadastro", classe: "primary", ocupado: "A aplicar…", click: async () => {
      if (!pronto()) return;
      await ctx.aplicar({ids: g.ids.filter((id, i) => !/^sim$/i.test(g.linhas[i].aplicada))}); fechar(); ctx.aviso("Adenda aplicada ao cadastro."); await ctx.recarregar();
    }});
    lista.push({texto: "Nova adenda para este colaborador", click: async () => nova({codigo: M.normalizarCodigo(g.codigo), nome: g.nome, empresa: g.empresa})});
    lista.push({texto: "Fechar", classe: lista.length === 1 ? "primary" : "outline", click: async () => fechar()});
    accoes(lista);
  }

  return {nova, detalhe};
}
