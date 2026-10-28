// Página «Relatórios»: catálogo de relatórios formais, pré-visualização em folha A4 e exportação para PDF e Word.
// O conteúdo vem de relatorios-dados.js; os três formatos são desenhados a partir do mesmo documento.
import * as M from "./rh-modulos.js";
import * as R from "./relatorios-dados.js";
import {descarregar} from "./rh-export.js";
import {mesesAte} from "./painel.js";

const LIBS = {
  jspdf: "https://cdn.jsdelivr.net/npm/jspdf@2.5.2/dist/jspdf.umd.min.js",
  autotable: "https://cdn.jsdelivr.net/npm/jspdf-autotable@3.8.4/dist/jspdf.plugin.autotable.min.js",
  docx: "https://cdn.jsdelivr.net/npm/docx@9.5.1/dist/index.iife.js"
};
const carregados = {};
function script(url) {
  if (!carregados[url]) carregados[url] = new Promise((ok, falha) => {
    const s = document.createElement("script");
    s.src = url; s.crossOrigin = "anonymous"; s.onload = ok;
    s.onerror = () => { delete carregados[url]; falha(Error("Não foi possível carregar a biblioteca de exportação. Verifique a ligação à internet.")); };
    document.head.append(s);
  });
  return carregados[url];
}

const EMPRESA_GRUPO = "SIR GRUPO";
const VERMELHO = [190, 37, 55], TINTA = [24, 35, 48], CINZA = [115, 128, 146];
const nomeSeguro = s => String(s).normalize("NFD").replace(/[̀-ͯ]/g, "").replace(/[\\/:*?"<>|]/g, "_").replace(/\s+/g, "_");
const agora = () => new Date().toLocaleString("pt-PT", {timeZone: "Africa/Maputo"}).replace(",", "");
const el = (tag, props = {}, ...filhos) => {
  const n = document.createElement(tag);
  for (const [k, v] of Object.entries(props)) {
    if (k === "class") n.className = v; else if (k === "text") n.textContent = v; else n[k] = v;
  }
  n.append(...filhos.flat(Infinity).filter(f => f != null && f !== false));
  return n;
};
const nomeFicheiro = (doc, ext) => nomeSeguro(doc.titulo.replace(/ — /g, " ") + " " + (doc.periodo || "") + (doc.ambito && !/Todas/.test(doc.ambito) ? " " + doc.ambito.replace("Empresa: ", "") : "")) + "." + ext;

// ---------------------------------------------------------------- pré-visualização (folha A4)
function tabelaHTML(tb) {
  const th = tb.cols.map(c => el("th", {class: c.num ? "num" : "", text: c.t}));
  return el("table", {class: "rel-doc-tabela"},
    el("thead", {}, el("tr", {}, th)),
    el("tbody", {}, tb.linhas.map(l => el("tr", {}, l.map((v, i) => el("td", {class: tb.cols[i].num ? "num" : "", text: String(v)}))))),
    tb.total ? el("tfoot", {}, el("tr", {}, tb.total.map((v, i) => el("td", {class: tb.cols[i].num ? "num" : "", text: String(v)})))) : null);
}
export function documentoHTML(doc, {utilizador = ""} = {}) {
  return el("article", {class: "rel-doc"},
    el("header", {class: "rel-doc-cab"},
      el("div", {class: "rel-doc-marca"}, el("span", {class: "rel-doc-logo", text: "S"}), el("div", {}, el("b", {text: EMPRESA_GRUPO}), el("small", {text: "Departamento de Recursos Humanos"}))),
      el("div", {class: "rel-doc-meta"}, el("span", {text: "Emitido em " + agora()}), utilizador ? el("span", {text: "por " + utilizador}) : null)),
    el("h1", {text: doc.titulo}),
    el("p", {class: "rel-doc-sub", text: [doc.periodo, doc.ambito].filter(Boolean).join("  ·  ")}),
    doc.vazio ? el("p", {class: "rel-doc-vazio", text: doc.vazio}) : [
      el("div", {class: "rel-doc-kpis"}, doc.kpis.map(k => el("div", {class: "rel-doc-kpi"}, el("span", {text: k.rotulo}), el("b", {text: k.valor}), el("small", {text: k.nota || ""})))),
      doc.secoes.map(s => el("section", {class: "rel-doc-sec"},
        el("h2", {text: s.titulo}),
        (s.paragrafos || []).map(p => el("p", {text: p})),
        s.tabela ? tabelaHTML(s.tabela) : null,
        s.subtabelas ? el("div", {class: "rel-doc-duas"}, s.subtabelas.map(tabelaHTML)) : null)),
      doc.notas.length ? el("div", {class: "rel-doc-notas"}, el("b", {text: "Notas"}), doc.notas.map(n => el("p", {text: n}))) : null,
      el("div", {class: "rel-doc-assin"},
        el("div", {}, el("i"), el("span", {text: "Elaborado por (Dep. Recursos Humanos)"})),
        el("div", {}, el("i"), el("span", {text: "Aprovado por"})))
    ],
    el("footer", {class: "rel-doc-rod", text: "Fonte: Gestão de Colaboradores.xlsx · Sistema RH SIR Grupo"}));
}

// ---------------------------------------------------------------- PDF
export async function documentoPDF(doc, {utilizador = ""} = {}) {
  await script(LIBS.jspdf); await script(LIBS.autotable);
  const pdf = new globalThis.jspdf.jsPDF({orientation: "portrait", unit: "mm", format: "a4"});
  const W = pdf.internal.pageSize.getWidth(), H = pdf.internal.pageSize.getHeight(), ML = 16, MR = 16, LW = W - ML - MR;
  let y = 18;
  const espaco = h => { if (y + h > H - 22) { pdf.addPage(); y = 22; } };
  // Cabeçalho
  pdf.setFillColor(...VERMELHO); pdf.roundedRect(ML, y - 6, 9, 9, 1.5, 1.5, "F");
  pdf.setTextColor(255); pdf.setFont("helvetica", "bold"); pdf.setFontSize(11); pdf.text("S", ML + 4.5, y + 0.6, {align: "center"});
  pdf.setTextColor(...TINTA); pdf.setFontSize(10); pdf.text(EMPRESA_GRUPO, ML + 12, y - 1.5);
  pdf.setFont("helvetica", "normal"); pdf.setFontSize(8); pdf.setTextColor(...CINZA); pdf.text("Departamento de Recursos Humanos", ML + 12, y + 2.5);
  pdf.text("Emitido em " + agora() + (utilizador ? " por " + utilizador : ""), W - MR, y - 1.5, {align: "right"});
  y += 7; pdf.setDrawColor(...VERMELHO); pdf.setLineWidth(0.6); pdf.line(ML, y, W - MR, y); y += 10;
  pdf.setTextColor(...TINTA); pdf.setFont("helvetica", "bold"); pdf.setFontSize(16);
  const tl = pdf.splitTextToSize(doc.titulo, LW); pdf.text(tl, ML, y); y += tl.length * 6.5;
  pdf.setFont("helvetica", "normal"); pdf.setFontSize(9.5); pdf.setTextColor(...CINZA);
  pdf.text([doc.periodo, doc.ambito].filter(Boolean).join("  ·  "), ML, y); y += 8;
  if (doc.vazio) { pdf.setTextColor(...TINTA); pdf.text(doc.vazio, ML, y); }
  else {
    // Indicadores: grelha 3 × 2
    const cw = (LW - 2 * 4) / 3, ch = 19;
    doc.kpis.forEach((k, i) => {
      const cx = ML + (i % 3) * (cw + 4), cy = y + Math.floor(i / 3) * (ch + 4);
      pdf.setDrawColor(232, 237, 242); pdf.setFillColor(252, 247, 248); pdf.setLineWidth(0.3); pdf.roundedRect(cx, cy, cw, ch, 2, 2, "FD");
      pdf.setFillColor(...VERMELHO); pdf.rect(cx, cy + 2, 0.9, ch - 4, "F");
      pdf.setFont("helvetica", "normal"); pdf.setFontSize(7.5); pdf.setTextColor(...CINZA); pdf.text(pdf.splitTextToSize(k.rotulo, cw - 6)[0], cx + 4, cy + 5.5);
      pdf.setFont("helvetica", "bold"); pdf.setFontSize(11.5); pdf.setTextColor(...TINTA); pdf.text(pdf.splitTextToSize(k.valor, cw - 6)[0], cx + 4, cy + 11.5);
      pdf.setFont("helvetica", "normal"); pdf.setFontSize(7); pdf.setTextColor(...CINZA); pdf.text(pdf.splitTextToSize(k.nota || "", cw - 6)[0] || "", cx + 4, cy + 16);
    });
    y += Math.ceil(doc.kpis.length / 3) * (ch + 4) + 4;
    const tabela = (tb, x = ML, largura = LW) => {
      pdf.autoTable({
        startY: y, margin: {left: x, right: W - x - largura, top: 22, bottom: 20}, tableWidth: largura, theme: "plain",
        head: [tb.cols.map(c => c.t)], body: tb.linhas.map(l => l.map(String)), foot: tb.total ? [tb.total.map(String)] : undefined, showFoot: "lastPage",
        styles: {font: "helvetica", fontSize: 8, cellPadding: {top: 1.6, bottom: 1.6, left: 2, right: 2}, textColor: TINTA, overflow: "linebreak"},
        headStyles: {fillColor: [248, 236, 238], textColor: [122, 18, 32], fontStyle: "bold", fontSize: 7.5},
        footStyles: {fontStyle: "bold", lineWidth: {top: 0.4}, lineColor: [122, 18, 32]},
        bodyStyles: {lineWidth: {bottom: 0.1}, lineColor: [228, 232, 236]},
        columnStyles: Object.fromEntries(tb.cols.map((c, i) => [i, c.num ? {halign: "right"} : {}])),
        didParseCell: d => { if (tb.cols[d.column.index]?.num && d.section !== "body") d.cell.styles.halign = "right"; }
      });
      return pdf.lastAutoTable.finalY;
    };
    for (const s of doc.secoes) {
      espaco(s.tabela || s.subtabelas ? 38 : 22); // o título nunca fica sozinho no fundo da página
      pdf.setFont("helvetica", "bold"); pdf.setFontSize(11.5); pdf.setTextColor(122, 18, 32); pdf.text(s.titulo, ML, y); y += 2;
      pdf.setDrawColor(239, 176, 182); pdf.setLineWidth(0.3); pdf.line(ML, y, W - MR, y); y += 5;
      pdf.setFont("helvetica", "normal"); pdf.setFontSize(9.5); pdf.setTextColor(...TINTA);
      for (const p of s.paragrafos || []) {
        const linhas = pdf.splitTextToSize(p, LW);
        espaco(linhas.length * 4.6);
        pdf.text(linhas, ML, y, {lineHeightFactor: 1.35}); y += linhas.length * 4.6 + 2;
      }
      if (s.tabela) { espaco(18); y = tabela(s.tabela) + 7; }
      if (s.subtabelas) {
        espaco(18);
        const meia = (LW - 6) / 2, y0 = y;
        const fins = s.subtabelas.map((tb, i) => { y = y0; const f = tabela(tb, ML + i * (meia + 6), meia); return {f, p: pdf.internal.getCurrentPageInfo().pageNumber}; });
        const ultima = fins.reduce((a, b) => (b.p > a.p || (b.p === a.p && b.f > a.f) ? b : a));
        pdf.setPage(ultima.p); y = ultima.f + 7;
      }
    }
    if (doc.notas.length) {
      const txt = doc.notas.flatMap(nn => pdf.splitTextToSize(nn, LW - 8));
      espaco(txt.length * 4 + 12);
      pdf.setFillColor(250, 250, 251); pdf.setDrawColor(232, 237, 242); pdf.roundedRect(ML, y, LW, txt.length * 4 + 9, 2, 2, "FD");
      pdf.setFont("helvetica", "bold"); pdf.setFontSize(8.5); pdf.setTextColor(...TINTA); pdf.text("Notas", ML + 4, y + 5);
      pdf.setFont("helvetica", "normal"); pdf.setFontSize(8); pdf.setTextColor(...CINZA); pdf.text(txt, ML + 4, y + 9.5); y += txt.length * 4 + 15;
    }
    espaco(30); y += 14;
    pdf.setDrawColor(...TINTA); pdf.setLineWidth(0.3);
    pdf.line(ML, y, ML + 70, y); pdf.line(W - MR - 70, y, W - MR, y);
    pdf.setFontSize(8); pdf.setTextColor(...CINZA);
    pdf.text("Elaborado por (Dep. Recursos Humanos)", ML + 35, y + 4.5, {align: "center"});
    pdf.text("Aprovado por", W - MR - 35, y + 4.5, {align: "center"});
  }
  const total = pdf.getNumberOfPages();
  for (let p = 1; p <= total; p++) {
    pdf.setPage(p); pdf.setFont("helvetica", "normal"); pdf.setFontSize(7.5); pdf.setTextColor(...CINZA);
    pdf.setDrawColor(232, 237, 242); pdf.setLineWidth(0.2); pdf.line(ML, H - 13, W - MR, H - 13);
    pdf.text(doc.titulo + (doc.periodo ? " · " + doc.periodo : ""), ML, H - 8.5);
    pdf.text("Página " + p + " de " + total, W - MR, H - 8.5, {align: "right"});
  }
  return pdf.output("blob");
}

// ---------------------------------------------------------------- Word (.docx)
export async function documentoWord(doc, {utilizador = ""} = {}) {
  await script(LIBS.docx);
  const D = globalThis.docx;
  const VERM = "BE2537", ESC = "7A1220", TIN = "182330", CIN = "738092";
  const larguraUtil = 9638; // A4 com margens de 2 cm, em twips
  const run = (text, o = {}) => new D.TextRun({text: String(text), font: "Calibri", size: o.size || 20, bold: o.bold, color: o.color || TIN, italics: o.italics});
  const par = (filhos, o = {}) => new D.Paragraph({children: [].concat(filhos), spacing: {after: o.after ?? 120, before: o.before ?? 0, line: o.line}, alignment: o.align, keepNext: o.keepNext,
    border: o.border, heading: o.heading});
  const celula = (texto, o = {}) => new D.TableCell({
    children: [par(run(texto, {bold: o.bold, size: o.size || 17, color: o.color}), {after: 0, align: o.num ? D.AlignmentType.RIGHT : D.AlignmentType.LEFT})],
    shading: o.fill ? {type: D.ShadingType.CLEAR, color: "auto", fill: o.fill} : undefined,
    margins: {top: 50, bottom: 50, left: 90, right: 90}, verticalAlign: D.VerticalAlign.CENTER, width: o.width ? {size: o.width, type: D.WidthType.DXA} : undefined,
    borders: o.borders
  });
  const semBordas = {top: {style: D.BorderStyle.NONE}, bottom: {style: D.BorderStyle.NONE}, left: {style: D.BorderStyle.NONE}, right: {style: D.BorderStyle.NONE}};
  const linhaFina = {top: {style: D.BorderStyle.NONE}, left: {style: D.BorderStyle.NONE}, right: {style: D.BorderStyle.NONE},
    bottom: {style: D.BorderStyle.SINGLE, size: 2, color: "E4E8EC"}};
  const tabela = (tb, largura = larguraUtil) => {
    const w = Math.floor(largura / tb.cols.length);
    return new D.Table({
      width: {size: largura, type: D.WidthType.DXA}, columnWidths: tb.cols.map(() => w),
      borders: {...semBordas, insideHorizontal: {style: D.BorderStyle.NONE}, insideVertical: {style: D.BorderStyle.NONE}},
      rows: [
        new D.TableRow({tableHeader: true, children: tb.cols.map(c => celula(c.t, {bold: true, fill: "F8ECEE", color: ESC, num: c.num, width: w, size: 16, borders: semBordas}))}),
        ...tb.linhas.map(l => new D.TableRow({cantSplit: true, children: l.map((v, i) => celula(v, {num: tb.cols[i].num, width: w, borders: linhaFina}))})),
        ...(tb.total ? [new D.TableRow({children: tb.total.map((v, i) => celula(v, {bold: true, num: tb.cols[i].num, width: w,
          borders: {...semBordas, top: {style: D.BorderStyle.SINGLE, size: 8, color: ESC}}}))})] : [])
      ]
    });
  };
  const corpo = [
    par([run("S ", {bold: true, color: VERM, size: 28}), run(EMPRESA_GRUPO, {bold: true, size: 22}), run("   Departamento de Recursos Humanos", {color: CIN, size: 17})],
      {after: 60, border: {bottom: {style: D.BorderStyle.SINGLE, size: 12, color: VERM, space: 6}}}),
    par(run("Emitido em " + agora() + (utilizador ? " por " + utilizador : ""), {color: CIN, size: 16}), {after: 240, align: D.AlignmentType.RIGHT}),
    par(run(doc.titulo, {bold: true, size: 36}), {after: 80, heading: D.HeadingLevel.TITLE}),
    par(run([doc.periodo, doc.ambito].filter(Boolean).join("  ·  "), {color: CIN, size: 20}), {after: 240})
  ];
  if (doc.vazio) corpo.push(par(run(doc.vazio)));
  else {
    // Indicadores numa grelha de 3 colunas
    const kw = Math.floor(larguraUtil / 3), linhasK = [];
    for (let i = 0; i < doc.kpis.length; i += 3) {
      linhasK.push(new D.TableRow({children: doc.kpis.slice(i, i + 3).concat(Array(Math.max(0, 3 - doc.kpis.slice(i, i + 3).length)).fill(null)).map(k =>
        new D.TableCell({width: {size: kw, type: D.WidthType.DXA}, shading: k ? {type: D.ShadingType.CLEAR, color: "auto", fill: "FCF7F8"} : undefined,
          margins: {top: 100, bottom: 100, left: 140, right: 100},
          borders: {top: {style: D.BorderStyle.SINGLE, size: 4, color: "FFFFFF"}, bottom: {style: D.BorderStyle.SINGLE, size: 4, color: "FFFFFF"},
            right: {style: D.BorderStyle.SINGLE, size: 4, color: "FFFFFF"}, left: {style: D.BorderStyle.SINGLE, size: 18, color: k ? VERM : "FFFFFF"}},
          children: k ? [par(run(k.rotulo, {size: 15, color: CIN}), {after: 40}), par(run(k.valor, {bold: true, size: 24}), {after: 40}), par(run(k.nota || "", {size: 14, color: CIN}), {after: 0})] : [par(run(""))]
        }))}));
    }
    corpo.push(new D.Table({width: {size: larguraUtil, type: D.WidthType.DXA}, columnWidths: [kw, kw, kw], rows: linhasK,
      borders: {...semBordas, insideHorizontal: {style: D.BorderStyle.NONE}, insideVertical: {style: D.BorderStyle.NONE}}}));
    corpo.push(par(run(""), {after: 120}));
    for (const s of doc.secoes) {
      corpo.push(par(run(s.titulo, {bold: true, size: 24, color: ESC}), {before: 200, after: 120, keepNext: true, heading: D.HeadingLevel.HEADING_2,
        border: {bottom: {style: D.BorderStyle.SINGLE, size: 4, color: "EFB0B6", space: 4}}}));
      for (const p of s.paragrafos || []) corpo.push(par(run(p, {size: 20}), {after: 140, line: 300}));
      if (s.tabela) { corpo.push(tabela(s.tabela)); corpo.push(par(run(""), {after: 80})); }
      if (s.subtabelas) for (const tb of s.subtabelas) { corpo.push(tabela(tb, Math.floor(larguraUtil * 0.62))); corpo.push(par(run(""), {after: 80})); }
    }
    if (doc.notas.length) {
      corpo.push(par(run("Notas", {bold: true, size: 18}), {before: 200, after: 60}));
      for (const nn of doc.notas) corpo.push(par(run(nn, {size: 16, color: CIN}), {after: 60}));
    }
    const aw = Math.floor(larguraUtil / 2);
    corpo.push(par(run(""), {after: 600}));
    corpo.push(new D.Table({width: {size: larguraUtil, type: D.WidthType.DXA}, columnWidths: [aw, aw],
      borders: {...semBordas, insideHorizontal: {style: D.BorderStyle.NONE}, insideVertical: {style: D.BorderStyle.NONE}},
      rows: [new D.TableRow({children: ["Elaborado por (Dep. Recursos Humanos)", "Aprovado por"].map(txt => new D.TableCell({width: {size: aw, type: D.WidthType.DXA},
        margins: {left: 300, right: 300}, borders: semBordas,
        children: [par(run("_______________________________", {color: CIN}), {after: 40, align: D.AlignmentType.CENTER}), par(run(txt, {size: 16, color: CIN}), {after: 0, align: D.AlignmentType.CENTER})]}))})]}));
  }
  const documento = new D.Document({
    creator: "Sistema RH SIR Grupo", title: doc.titulo, description: [doc.periodo, doc.ambito].filter(Boolean).join(" · "),
    styles: {default: {document: {run: {font: "Calibri", size: 20, color: TIN}}}},
    sections: [{
      properties: {page: {size: {width: 11906, height: 16838}, margin: {top: 1134, bottom: 1134, left: 1134, right: 1134}}},
      footers: {default: new D.Footer({children: [par([run(doc.titulo + (doc.periodo ? " · " + doc.periodo : "") + "    Página ", {size: 15, color: CIN}),
        new D.TextRun({children: [D.PageNumber.CURRENT], size: 15, color: CIN, font: "Calibri"}), run(" de ", {size: 15, color: CIN}),
        new D.TextRun({children: [D.PageNumber.TOTAL_PAGES], size: 15, color: CIN, font: "Calibri"})], {align: D.AlignmentType.RIGHT, after: 0})]})},
      children: corpo
    }]
  });
  return D.Packer.toBlob(documento);
}

// ---------------------------------------------------------------- página
export function iniciarRelatorios(ctx) {
  const $ = id => document.getElementById(id);
  let escolhido = "salarial", actual = null, aGerar = false;

  function catalogo() {
    $("rel-catalogo").replaceChildren(...R.CATALOGO.map(c => {
      const b = el("button", {type: "button", class: "rel-cartao" + (c.id === escolhido ? " activo" : "")},
        el("span", {class: "rel-cartao-ic", text: c.icone}), el("b", {text: c.titulo}), el("small", {text: c.descricao}));
      b.addEventListener("click", () => { escolhido = c.id; catalogo(); parametros(); gerar(); });
      return b;
    }));
  }
  function parametros() {
    const c = R.CATALOGO.find(x => x.id === escolhido);
    $("relp-mes-l").hidden = !c.precisa.includes("mes");
    $("relp-ano-l").hidden = !(c.precisa.includes("mes") || c.precisa.includes("ano"));
  }
  async function preencherFiltros() {
    const p = (await ctx.periodoMaisRecente()) || {ano: new Date().getFullYear(), mes: new Date().getMonth() + 1};
    const anoSel = $("relp-ano"), mesSel = $("relp-mes");
    if (!mesSel.options.length) mesSel.replaceChildren(...M.NOME_MES.map((m, i) => new Option(m, String(i + 1))));
    if (!anoSel.options.length) for (let a = p.ano; a >= p.ano - 5; a--) anoSel.append(new Option(String(a), String(a)));
    if (!$("relp-mes").dataset.iniciado) { mesSel.value = String(p.mes); anoSel.value = String(p.ano); $("relp-mes").dataset.iniciado = "1"; }
    try {
      const emp = [...new Set((await ctx.funcionarios()).map(f => String(M.campo(f, "Empresa") || "").trim()).filter(Boolean))].sort();
      const sel = $("relp-empresa"), v = sel.value;
      sel.replaceChildren(new Option("Todas as empresas", ""), ...emp.map(e => new Option(e, e)));
      if (emp.includes(v)) sel.value = v;
    } catch { /* sem cadastro: fica «Todas» */ }
  }

  async function dados() {
    const c = R.CATALOGO.find(x => x.id === escolhido);
    const ano = Number($("relp-ano").value), mes = Number($("relp-mes").value), empresa = $("relp-empresa").value;
    const base = {ano, mes, empresa, hoje: new Date()};
    const estado = txt => { $("rel-estado").textContent = txt; };
    if (c.precisa.includes("mes")) { estado("A ler os lançamentos de " + M.NOME_MES[mes - 1] + " " + ano + "…"); base.lancMes = await ctx.carregarMes(ano, mes); }
    if (c.precisa.includes("anterior")) { const a = mes === 1 ? {ano: ano - 1, mes: 12} : {ano, mes: mes - 1}; base.lancAnterior = await ctx.carregarMes(a.ano, a.mes).catch(() => []); }
    if (c.precisa.includes("cadastro")) { estado("A ler o cadastro…"); base.funcionarios = await ctx.funcionarios(); }
    if (c.precisa.includes("indemnizacoes")) {
      estado("A ler as indemnizações…");
      const d = ctx.indemnizacoes ? await ctx.indemnizacoes() : null;
      if (!d) throw Error("Para este relatório, abra primeiro a página Indemnizações e cole o endereço do fluxo RH_API_Indemnizacoes.");
      base.indemnizacoes = d.indemnizacoes; base.prestacoes = d.prestacoes;
    }
    if (c.precisa.includes("doze")) {
      const ultimo = (await ctx.periodoMaisRecente()) || {ano, mes};
      const lista = mesesAte(ultimo);
      base.lancPorMes = [];
      let feitos = 0;
      for (const p of lista) { base.lancPorMes.push({...p, linhas: await ctx.carregarMes(p.ano, p.mes).catch(() => [])}); estado(`A ler lançamentos… ${++feitos} de ${lista.length} meses`); }
    }
    return base;
  }
  const GERADORES = {salarial: R.relatorioSalarial, encargos: R.relatorioEncargos, efectivo: R.relatorioEfectivo,
    movimentacao: R.relatorioMovimentacao, assiduidade: R.relatorioAssiduidade, evolucao: R.relatorioEvolucao,
    indemnizacoes: R.relatorioIndemnizacoes};

  async function gerar() {
    if (!ctx.ligado() || aGerar) return;
    aGerar = true; $("rel-gerar").disabled = true; botoes(false);
    const t0 = Date.now();
    $("rel-folha").replaceChildren(el("p", {class: "loading-line", text: "A preparar o relatório…"}));
    try {
      const doc = GERADORES[escolhido](await dados());
      actual = doc;
      $("rel-folha").replaceChildren(documentoHTML(doc, {utilizador: ctx.utilizador()}));
      $("rel-estado").textContent = `Relatório gerado em ${Math.round((Date.now() - t0) / 1000)} s.`;
      botoes(!doc.vazio);
    } catch (e) {
      $("rel-folha").replaceChildren(el("p", {class: "form-error", text: "Não foi possível gerar o relatório: " + (e.message || e)}));
      $("rel-estado").textContent = "";
    } finally { aGerar = false; $("rel-gerar").disabled = false; }
  }
  function botoes(activos) { ["rel-pdf", "rel-word"].forEach(id => { $(id).disabled = !activos; }); }
  async function exportar(tipo, botao) {
    if (!actual) return;
    const texto = botao.textContent; botao.disabled = true; botao.textContent = "A exportar…";
    try {
      const blob = tipo === "pdf" ? await documentoPDF(actual, {utilizador: ctx.utilizador()}) : await documentoWord(actual, {utilizador: ctx.utilizador()});
      descarregar(blob, nomeFicheiro(actual, tipo === "pdf" ? "pdf" : "docx"));
    } catch (e) { ctx.aviso(e.message || String(e)); }
    finally { botao.disabled = false; botao.textContent = texto; }
  }

  $("rel-gerar").addEventListener("click", gerar);
  $("rel-pdf").addEventListener("click", e => exportar("pdf", e.currentTarget));
  $("rel-word").addEventListener("click", e => exportar("word", e.currentTarget));

  async function abrir() {
    const ligado = ctx.ligado();
    $("rel-vazio").hidden = ligado; $("rel-conteudo").hidden = !ligado;
    if (!ligado) return;
    catalogo(); parametros(); botoes(Boolean(actual && !actual.vazio));
    await preencherFiltros();
    if (!actual) gerar();
  }
  function limpar() { actual = null; $("rel-folha").replaceChildren(); $("rel-estado").textContent = ""; $("rel-vazio").hidden = false; $("rel-conteudo").hidden = true; }
  return {abrir, limpar};
}
