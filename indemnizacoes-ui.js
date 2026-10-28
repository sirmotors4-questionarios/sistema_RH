// Página Indemnizações: indicadores do mês, faixa de meses, calendário de pagamentos, planos e diálogos.
import * as M from "./rh-modulos.js";
import * as I from "./indemnizacoes.js";
import {indemnizacoesListar, indemnizacaoGravar, prestacaoPagar, indemnizacaoAnular} from "./rh-api.js";

const $ = id => document.getElementById(id);
const el = (tag, props = {}, ...filhos) => {
  const n = document.createElement(tag);
  for (const [k, v] of Object.entries(props)) {
    if (v == null || v === false) continue;
    if (k === "class") n.className = v; else if (k === "text") n.textContent = v;
    else if (k.startsWith("data-") || k.startsWith("aria-") || k === "for" || k === "list") n.setAttribute(k, v); else n[k] = v;
  }
  for (const f of filhos.flat(Infinity)) if (f != null && f !== false) n.append(f);
  return n;
};
const fmt = v => new Intl.NumberFormat("pt-MZ", {minimumFractionDigits: 2, maximumFractionDigits: 2}).format(v || 0);
const curto = v => { const n = Number(v) || 0; return Math.abs(n) >= 1e6 ? (n / 1e6).toFixed(1).replace(".", ",") + " M" : Math.abs(n) >= 1e3 ? Math.round(n / 1e3) + " mil" : fmt(n); };
const dataPt = iso => iso ? iso.slice(8, 10) + "/" + iso.slice(5, 7) + "/" + iso.slice(0, 4) : "—";
const hojeIso = () => new Date().toLocaleDateString("en-CA", {timeZone: "Africa/Maputo"});
const CLASSE = {"Paga": "p-paga", "Pendente": "p-pendente", "Em atraso": "p-atraso"};

export function iniciarIndemnizacoes(ctx) {
  let dados = null, aCarregar = null;
  const hoje0 = hojeIso();
  let vista = {ano: Number(hoje0.slice(0, 4)), mes: Number(hoje0.slice(5, 7))};
  const dlg = $("ind-dlg");

  // ---------- carregar ----------
  function deps() { return ctx.transporte(); }
  async function carregar(forcar = false) {
    if (!ctx.ligado()) { mostrarVazio(); return; }
    if (!ctx.temUrl()) { mostrarVazio(true); return; }
    if (dados && !forcar) { render(); return; }
    if (aCarregar) return aCarregar;
    $("ind-badge").textContent = "A ler o Excel…";
    const t0 = Date.now();
    aCarregar = indemnizacoesListar(deps()).then(r => {
      dados = {lista: I.indemnizacoes(r.indemnizacoes), prest: I.prestacoes(r.prestacoes), lido: new Date()};
      $("ind-badge").textContent = "Lido em " + Math.round((Date.now() - t0) / 1000) + " s · " + dados.lista.length + " plano(s)";
      render();
    }).catch(e => {
      ctx.aviso("Indemnizações: " + e.message);
      $("ind-badge").textContent = "Erro: " + e.message;
      if (!dados) mostrarVazio(false, "Não foi possível ler as indemnizações: " + e.message);
    }).finally(() => { aCarregar = null; });
    return aCarregar;
  }
  function mostrarVazio(pedirUrl = false, txt = "") {
    document.querySelectorAll('[data-empty="indemnizacoes"]').forEach(x => { x.hidden = false; });
    document.querySelectorAll('[data-loaded="indemnizacoes"]').forEach(x => { x.hidden = true; });
    $("ind-url-box").hidden = !(ctx.ligado() && pedirUrl);
    $("ind-vazio-txt").textContent = txt || (!ctx.ligado() ? "Inicie sessão para consultar as indemnizações e os planos de pagamento."
      : pedirUrl ? "Para ligar esta página ao Excel, cole uma vez o endereço do fluxo RH_API_Indemnizacoes." : "A carregar…");
  }

  // ---------- render ----------
  function render() {
    if (!dados) return;
    document.querySelectorAll('[data-empty="indemnizacoes"]').forEach(x => { x.hidden = true; });
    document.querySelectorAll('[data-loaded="indemnizacoes"]').forEach(x => { x.hidden = false; });
    const hoje = hojeIso(), {ano, mes} = vista, actual = ano === Number(hoje.slice(0, 4)) && mes === Number(hoje.slice(5, 7));
    $("ind-mes-titulo").textContent = M.NOME_MES[mes - 1] + " " + ano + (actual ? " · mês actual" : "");
    $("ind-hoje").hidden = actual;
    const k = I.indicadoresMes(dados.lista, dados.prest, ano, mes, hoje);
    const card = (rot, val, sub, cls = "") => el("article", {class: "metric " + cls}, el("span", {text: rot}), el("strong", {text: val}), el("small", {text: sub}));
    $("ind-kpis").replaceChildren(
      card("Valor do mês", fmt(k.valorMes), k.nMes + " prestação(ões) · MT"),
      card("Pago no mês", fmt(k.pagoMes), "MT"),
      card("Por pagar no mês", fmt(k.porPagarMes), k.atrasoMes ? fmt(k.atrasoMes) + " MT em atraso" : "MT", k.atrasoMes ? "alerta" : ""),
      card("Valor gasto total", fmt(k.gastoTotal), "pago até hoje · MT"),
      card("Total comprometido", fmt(k.comprometido), "saldo por pagar " + fmt(k.saldo) + " MT"),
      card("Em atraso (todos)", fmt(k.atrasoTotal), k.nAtraso + " prestação(ões) · MT", k.nAtraso ? "alerta" : ""));
    renderFaixa(); renderCalendario(k.doMes, hoje); renderLista();
  }

  function renderFaixa() {
    const meses = I.totaisMensais(dados.lista, dados.prest, vista.ano, vista.mes), max = Math.max(1, ...meses.map(m => m.previsto));
    $("ind-faixa").replaceChildren(...meses.map(m => {
      const sel = m.ano === vista.ano && m.mes === vista.mes;
      const b = el("button", {type: "button", class: "ind-mesbtn" + (sel ? " sel" : ""), title: M.NOME_MES[m.mes - 1] + " " + m.ano + ": previsto " + fmt(m.previsto) + " MT, pago " + fmt(m.pago) + " MT"},
        el("div", {class: "ind-barra"}, el("i", {class: "prev", style: "height:" + (m.previsto / max * 100) + "%"}), el("i", {class: "pago", style: "height:" + (m.pago / max * 100) + "%"})),
        el("b", {text: m.previsto ? curto(m.previsto) : "—"}), el("small", {text: M.NOME_MES[m.mes - 1].slice(0, 3) + " " + String(m.ano).slice(2)}));
      b.addEventListener("click", () => { vista = {ano: m.ano, mes: m.mes}; render(); });
      return b;
    }));
  }

  function renderCalendario(doMes, hoje) {
    const sem = I.calendario(vista.ano, vista.mes, doMes);
    const cab = el("div", {class: "ind-cal-cab"}, ...["Seg", "Ter", "Qua", "Qui", "Sex", "Sáb", "Dom"].map(d => el("span", {text: d})));
    const grelha = el("div", {class: "ind-cal-grelha"}, ...sem.flat().map(d => {
      if (!d) return el("div", {class: "ind-dia vazio"});
      const total = d.prestacoes.reduce((s, p) => s + p.valor, 0);
      return el("div", {class: "ind-dia" + (d.data === hoje ? " hoje" : "") + (d.prestacoes.length ? " tem" : "")},
        el("div", {class: "ind-dia-n"}, el("span", {text: String(d.dia)}), total ? el("em", {text: curto(total)}) : null),
        ...d.prestacoes.map(p => {
          const s = I.situacao(p, hoje);
          const b = el("button", {type: "button", class: "ind-chip " + CLASSE[s], title: p.nome + " · prestação " + p.numero + " · " + fmt(p.valor) + " MT · " + s + (p.dataPagamento ? " em " + dataPt(p.dataPagamento) : "")},
            el("span", {text: p.nome || p.codigo}), el("b", {text: fmt(p.valor)}));
          b.addEventListener("click", () => abrirPrestacao(p));
          return b;
        }));
    }));
    const lista = doMes.length ? null : el("p", {class: "ind-sem", text: "Sem prestações com data limite em " + M.NOME_MES[vista.mes - 1].toLowerCase() + " de " + vista.ano + "."});
    $("ind-cal").replaceChildren(cab, grelha, lista || "");
  }

  function renderLista() {
    const filtro = $("ind-estado").value, hoje = hojeIso(), body = $("ind-body");
    body.replaceChildren();
    const linhas = dados.lista.map(i => ({i, r: I.resumo(i, dados.prest)})).map(x => ({...x, estado: /anulad/i.test(x.i.estado) ? "Anulada" : x.r.concluida ? "Concluída" : "Activa"}))
      .filter(x => filtro ? x.estado === filtro : x.estado !== "Anulada")
      .sort((a, b) => (a.r.proxima?.dataLimite || "9999").localeCompare(b.r.proxima?.dataLimite || "9999") || a.i.nome.localeCompare(b.i.nome, "pt"));
    if (!linhas.length) body.append(el("tr", {}, el("td", {colSpan: 11, class: "no-data", text: dados.lista.length ? "Nenhum plano com este estado." : "Ainda não há indemnizações registadas."})));
    linhas.forEach(({i, r, estado}) => {
      const atraso = r.prestacoes.some(p => I.emAtraso(p, hoje));
      const tr = el("tr", {tabIndex: 0, title: "Ver o plano de " + i.nome},
        el("td", {}, el("strong", {text: i.nome}), el("small", {class: "ind-cod", text: " " + i.codigo})), el("td", {text: i.empresa}),
        el("td", {text: dataPt(i.admissao)}), el("td", {class: "num", text: String(i.anos)}), el("td", {class: "num", text: fmt(i.base)}),
        el("td", {class: "num", text: fmt(i.total)}), el("td", {class: "num", text: r.pagas + " / " + r.prestacoes.length}),
        el("td", {class: "num", text: fmt(r.pago)}), el("td", {class: "num", text: fmt(r.saldo)}),
        el("td", {text: r.proxima ? dataPt(r.proxima.dataLimite) + " · " + fmt(r.proxima.valor) : "—"}),
        el("td", {}, el("span", {class: "state" + (estado === "Activa" && !atraso ? "" : " off") + (atraso ? " atraso" : ""), text: atraso ? "Em atraso" : estado})));
      const abre = () => abrirPlano(i);
      tr.addEventListener("click", abre); tr.addEventListener("keydown", e => { if (e.key === "Enter") abre(); });
      body.append(tr);
    });
  }

  // ---------- diálogo ----------
  function abrirDlg({eyebrow, titulo, sub = ""}) {
    $("ind-dlg-eyebrow").textContent = eyebrow; $("ind-dlg-titulo").textContent = titulo; $("ind-dlg-sub").textContent = sub;
    $("ind-dlg-erro").textContent = ""; $("ind-dlg-accoes").replaceChildren();
    if (!dlg.open) { if (typeof dlg.showModal === "function") dlg.showModal(); else dlg.setAttribute("open", ""); }
  }
  const fechar = () => { if (typeof dlg.close === "function") dlg.close(); else dlg.removeAttribute("open"); };
  $("ind-dlg-fechar").addEventListener("click", fechar);
  function accoes(lista) {
    $("ind-dlg-accoes").replaceChildren(el("span"), ...lista.map(a => {
      const b = el("button", {type: "button", class: a.classe || "outline", text: a.texto, disabled: Boolean(a.desactivado)});
      b.addEventListener("click", async () => {
        if (!a.click) return;
        const t = b.textContent, t0 = Date.now();
        b.disabled = true; b.textContent = (a.ocupado || "A gravar…"); $("ind-dlg-erro").textContent = "";
        const relogio = setInterval(() => { if (b.isConnected) b.textContent = (a.ocupado || "A gravar…") + " " + Math.round((Date.now() - t0) / 1000) + " s"; }, 1000);
        try { await a.click(); } catch (e) { $("ind-dlg-erro").textContent = e.message; }
        finally { clearInterval(relogio); if (b.isConnected) { b.disabled = false; b.textContent = t; } }
      });
      return b;
    }));
  }
  function podeGravar() {
    if (!ctx.podeGravar()) { ctx.aviso("O seu perfil permite consultar, mas não gravar indemnizações."); return false; }
    return true;
  }

  // Prestação: marcar como paga (com data) ou voltar a pendente.
  function abrirPrestacao(p) {
    const hoje = hojeIso(), s = I.situacao(p, hoje), ind = dados.lista.find(i => i.id === p.indemnizacao);
    abrirDlg({eyebrow: "PRESTAÇÃO " + p.numero + (ind ? " DE " + ind.n : ""), titulo: p.nome + " · " + fmt(p.valor) + " MT", sub: "Data limite " + dataPt(p.dataLimite) + " · " + s});
    const data = el("input", {type: "date", id: "ind-pag-data", value: p.dataPagamento || (hoje < p.dataLimite ? hoje : hoje)});
    $("ind-dlg-body").replaceChildren(el("div", {class: "ind-prest"},
      el("p", {}, "Estado: ", el("span", {class: "ind-chip " + CLASSE[s], text: s}), p.dataPagamento ? " · paga em " + dataPt(p.dataPagamento) + (p.pagoPor ? " por " + p.pagoPor : "") : ""),
      I.paga(p) ? null : el("label", {class: "ind-campo"}, "Data do pagamento", data),
      ind ? el("button", {type: "button", class: "link-btn", text: "Ver o plano completo", onclick: () => abrirPlano(ind)}) : null));
    accoes(I.paga(p)
      ? [{texto: "Voltar a pendente", click: () => gravarPag(p, false)}, {texto: "Fechar", click: fechar}]
      : [{texto: "Fechar", click: fechar}, {texto: "Marcar como paga", classe: "primary", click: () => gravarPag(p, true, $("ind-pag-data").value)}]);
  }
  async function gravarPag(p, paga, data = "") {
    if (!podeGravar()) return;
    if (paga && !M.data(data)) throw Error("Indique a data do pagamento.");
    await prestacaoPagar({id: p.id, paga, data}, deps());
    // Actualiza já no ecrã; a próxima leitura confirma com o Excel.
    Object.assign(p, paga ? {estado: "Paga", dataPagamento: data, pagoPor: ctx.utilizador()} : {estado: "Pendente", dataPagamento: "", pagoPor: ""});
    fechar(); render(); ctx.aviso(paga ? "Prestação marcada como paga." : "Prestação voltou a pendente.");
  }

  // Plano completo de uma indemnização.
  function abrirPlano(i) {
    const r = I.resumo(i, dados.prest), hoje = hojeIso();
    abrirDlg({eyebrow: "PLANO " + i.id, titulo: i.nome + " · " + i.empresa,
      sub: i.anos + " ano(s) completos × " + fmt(i.base) + " MT = " + fmt(i.total) + " MT · admissão " + dataPt(i.admissao) + " · referência " + dataPt(i.referencia)});
    $("ind-dlg-body").replaceChildren(
      el("div", {class: "stats compact-stats rel-stats"},
        ...[["Total", fmt(i.total)], ["Pago", fmt(r.pago)], ["Saldo", fmt(r.saldo)], ["Prestações pagas", r.pagas + " / " + r.prestacoes.length]]
          .map(([a, b]) => el("article", {class: "metric mini"}, el("span", {text: a}), el("strong", {text: b})))),
      el("div", {class: "table-scroll"}, el("table", {class: "rel-table"},
        el("thead", {}, el("tr", {}, ...["Nº", "Data limite", "Valor", "Estado", "Pago em", ""].map((t, j) => el("th", {class: j === 2 ? "num" : "", text: t})))),
        el("tbody", {}, ...r.prestacoes.map(p => {
          const s = I.situacao(p, hoje);
          return el("tr", {}, el("td", {text: String(p.numero)}), el("td", {text: dataPt(p.dataLimite)}), el("td", {class: "num", text: fmt(p.valor)}),
            el("td", {}, el("span", {class: "ind-chip " + CLASSE[s], text: s})), el("td", {text: dataPt(p.dataPagamento)}),
            el("td", {}, /anulad/i.test(i.estado) ? "" : el("button", {type: "button", class: "inline-btn", text: I.paga(p) ? "Desmarcar" : "Marcar paga", onclick: () => abrirPrestacao(p)})));
        })))),
      i.observacoes ? el("p", {class: "rel-nota", text: "Observações: " + i.observacoes}) : null);
    accoes([
      ...(/anulad/i.test(i.estado) || r.pagas ? [] : [{texto: "Anular plano", click: () => anular(i)}]),
      {texto: "Fechar", classe: "primary", click: fechar}]);
  }
  async function anular(i) {
    if (!podeGravar()) return;
    if (!confirmarAnular(i)) return;
    await indemnizacaoAnular({id: i.id, observacoes: "Anulada por " + ctx.utilizador()}, deps());
    i.estado = "Anulada"; fechar(); render(); ctx.aviso("Plano " + i.id + " anulado (as linhas ficam no Excel como histórico).");
  }
  // Confirmação no próprio diálogo (sem janelas do navegador).
  let confirmacao = null;
  function confirmarAnular(i) {
    if (confirmacao === i.id) { confirmacao = null; return true; }
    confirmacao = i.id; $("ind-dlg-erro").textContent = "Carregue de novo em «Anular plano» para confirmar. O plano fica marcado como Anulado; nada é apagado.";
    return false;
  }

  // ---------- nova indemnização ----------
  function colaboradores() {
    return (ctx.funcionarios() || []).map(l => ({
      empresa: String(M.campo(l, "Empresa") || "").trim(), codigo: M.normalizarCodigo ? M.normalizarCodigo(M.campo(l, "Codigo")) : String(M.campo(l, "Codigo")),
      nome: String(M.campo(l, "Nome") || "").trim(), departamento: String(M.campo(l, "Departamento") || "").trim(),
      admissao: M.data(M.campo(l, "Data Admissao", "Data de Admissão", "Data Admissão")),
      saida: M.data(M.campo(l, "Data Saida", "Data de Saída", "Data Saída")),
      base: M.numero(M.campo(l, "Vencimento Base", "Vencimento", "Salario Base")) ?? null
    })).filter(c => c.nome);
  }
  function novo() {
    if (!ctx.ligado()) { ctx.aviso("Inicie sessão para registar indemnizações."); return; }
    if (!dados) { ctx.aviso("Aguarde a leitura das indemnizações."); carregar(); return; }
    if (!podeGravar()) return;
    const lista = colaboradores();
    const hoje = hojeIso(), prox = new Date(Date.UTC(Number(hoje.slice(0, 4)), Number(hoje.slice(5, 7)), 1));
    const st = {c: null, plano: []};
    abrirDlg({eyebrow: "NOVA INDEMNIZAÇÃO", titulo: "Nova indemnização", sub: "Escolha o colaborador; o sistema calcula os anos completos e o valor (anos × vencimento base)."});
    const dl = el("datalist", {id: "ind-colabs"}, ...lista.map(c => el("option", {value: c.codigo + " · " + c.nome + " · " + c.empresa})));
    const f = {
      colab: el("input", {id: "ind-f-colab", list: "ind-colabs", placeholder: lista.length ? "Código ou nome…" : "A carregar colaboradores…", autocomplete: "off"}),
      adm: el("input", {id: "ind-f-adm", type: "date"}), ref: el("input", {id: "ind-f-ref", type: "date", value: hoje}),
      base: el("input", {id: "ind-f-base", type: "number", min: "0", step: "0.01"}),
      n: el("input", {id: "ind-f-n", type: "number", min: "1", max: "120", step: "1", value: "6"}),
      dia: el("input", {id: "ind-f-dia", type: "number", min: "1", max: "31", step: "1", value: "25"}),
      mes: el("input", {id: "ind-f-mes", type: "month", value: prox.toISOString().slice(0, 7)}),
      obs: el("input", {id: "ind-f-obs", maxLength: 200, placeholder: "Opcional"})
    };
    const calc = el("div", {class: "ind-calc"}), planoBox = el("div", {class: "ind-plano"});
    const campo = (rot, input, ajuda) => el("label", {class: "ind-campo"}, rot, input, ajuda ? el("small", {text: ajuda}) : null);
    $("ind-dlg-body").replaceChildren(dl,
      el("div", {class: "ind-form"},
        campo("Colaborador", f.colab), campo("Data de admissão", f.adm, "Do cadastro; pode corrigir."),
        campo("Data de referência", f.ref, "Data de saída ou da decisão."), campo("Vencimento base (MT)", f.base, "Do cadastro; pode corrigir."),
        campo("Número de meses", f.n), campo("Dia limite de cada mês", f.dia, "Se o mês não tiver esse dia, usa o último."),
        campo("Primeiro mês", f.mes), campo("Observações", f.obs)),
      calc, planoBox);
    const escolher = () => {
      const v = f.colab.value.trim(), cod = v.split("·")[0].trim(), emp = (v.split("·")[2] || "").trim();
      const c = lista.find(x => x.codigo === (M.normalizarCodigo ? M.normalizarCodigo(cod) : cod) && (!emp || x.empresa === emp));
      st.c = c || null;
      if (c) { f.adm.value = c.admissao || ""; if (c.base != null) f.base.value = String(c.base); if (c.saida) f.ref.value = c.saida; }
      refazer();
    };
    const refazer = () => {
      const r = I.calcular({admissao: f.adm.value, referencia: f.ref.value, vencimentoBase: f.base.value});
      st.total = r.total; st.anos = r.anos;
      calc.replaceChildren(
        el("div", {}, el("span", {text: "Anos completos"}), el("b", {text: String(r.anos)})),
        el("div", {}, el("span", {text: "Vencimento base"}), el("b", {text: fmt(r.base)})),
        el("div", {class: "total"}, el("span", {text: "Indemnização"}), el("b", {text: fmt(r.total) + " MT"})));
      st.plano = I.gerarPlano(r.total, f.n.value, f.dia.value, f.mes.value);
      desenharPlano();
    };
    const desenharPlano = () => {
      if (!st.plano.length) { planoBox.replaceChildren(el("p", {class: "ind-sem", text: st.c ? "Sem valor a pagar: verifique as datas e o vencimento base." : "Escolha o colaborador para ver o plano."})); return; }
      const soma = st.plano.reduce((s, p) => s + Number(p.valor), 0);
      planoBox.replaceChildren(el("h3", {text: "Plano de pagamento · " + st.plano.length + " prestação(ões)"}),
        el("div", {class: "table-scroll"}, el("table", {class: "rel-table"},
          el("thead", {}, el("tr", {}, el("th", {text: "Nº"}), el("th", {text: "Data limite (editável)"}), el("th", {class: "num", text: "Valor"}))),
          el("tbody", {}, ...st.plano.map((p, j) => {
            const d = el("input", {type: "date", value: p.dataLimite, "aria-label": "Data limite da prestação " + p.numero});
            d.addEventListener("change", () => { st.plano[j].dataLimite = d.value; });
            return el("tr", {}, el("td", {text: String(p.numero)}), el("td", {}, d), el("td", {class: "num", text: fmt(p.valor)}));
          })),
          el("tfoot", {}, el("tr", {}, el("td", {text: "Total"}), el("td"), el("td", {class: "num", text: fmt(soma)}))))));
    };
    f.colab.addEventListener("change", escolher);
    [f.adm, f.ref, f.base, f.n, f.dia, f.mes].forEach(x => x.addEventListener("input", refazer));
    refazer();
    accoes([{texto: "Cancelar", click: fechar}, {texto: "Gravar indemnização e plano", classe: "primary", click: async () => {
      if (!st.c) throw Error("Escolha um colaborador da lista.");
      if (!(st.total > 0)) throw Error("A indemnização é 0: o colaborador não tem anos completos ou falta o vencimento base.");
      const erros = I.validarPlano(st.plano, st.total);
      if (erros.length) throw Error(erros.join(" "));
      const activa = dados.lista.find(x => x.empresa === st.c.empresa && M.normalizarCodigo(x.codigo) === st.c.codigo && !/anulad/i.test(x.estado));
      if (activa) throw Error("Já existe o plano " + activa.id + " para este colaborador. Anule-o antes de criar outro.");
      const r = await indemnizacaoGravar({empresa: st.c.empresa, codigo: st.c.codigo, nome: st.c.nome, departamento: st.c.departamento,
        dataAdmissao: f.adm.value, dataReferencia: f.ref.value, vencimentoBase: Number(f.base.value), diaLimite: Number(f.dia.value),
        primeiroMes: f.mes.value, observacoes: f.obs.value.trim(), prestacoes: st.plano}, deps());
      fechar(); ctx.aviso("Indemnização " + r.id + " gravada: " + fmt(r.total) + " MT em " + r.prestacoes + " prestação(ões).");
      const [a, m] = f.mes.value.split("-").map(Number); vista = {ano: a, mes: m};
      await carregar(true);
    }}]);
  }

  // ---------- eventos ----------
  const mudar = d => { const x = new Date(Date.UTC(vista.ano, vista.mes - 1 + d, 1)); vista = {ano: x.getUTCFullYear(), mes: x.getUTCMonth() + 1}; render(); };
  $("ind-ant").addEventListener("click", () => mudar(-1));
  $("ind-seg").addEventListener("click", () => mudar(1));
  $("ind-hoje").addEventListener("click", () => { const h = hojeIso(); vista = {ano: Number(h.slice(0, 4)), mes: Number(h.slice(5, 7))}; render(); });
  $("ind-actualizar").addEventListener("click", () => carregar(true));
  $("ind-estado").addEventListener("change", renderLista);
  $("ind-nova").addEventListener("click", novo);
  $("ind-url-guardar").addEventListener("click", () => {
    try { ctx.guardarUrl($("ind-url").value); $("ind-url").value = ""; carregar(true); } catch (e) { ctx.aviso(e.message); }
  });

  return {
    abrir: () => { const h = hojeIso(); if (!dados) vista = {ano: Number(h.slice(0, 4)), mes: Number(h.slice(5, 7))}; carregar(); },
    limpar: () => { dados = null; fechar(); mostrarVazio(); },
    dados: () => dados
  };
}
