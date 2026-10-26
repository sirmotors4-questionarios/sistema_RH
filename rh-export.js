// Exportação da folha salarial (PDF e Excel) e do relatório comparativo (Excel e PDF).
// Gerado no navegador a partir de tbl_lançamento: nada sai do computador do utilizador.
// Layout igual ao modRH_ExportarFolha / modRH_RelatorioComparativo do VBA.
import {NOME_MES} from "./rh-modulos.js";
import {somar} from "./rh-folha.js";

const LIBS = {
  exceljs: "https://cdn.jsdelivr.net/npm/exceljs@4.4.0/dist/exceljs.min.js",
  jspdf: "https://cdn.jsdelivr.net/npm/jspdf@2.5.2/dist/jspdf.umd.min.js",
  autotable: "https://cdn.jsdelivr.net/npm/jspdf-autotable@3.8.4/dist/jspdf.plugin.autotable.min.js"
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
async function excel() { await script(LIBS.exceljs); return globalThis.ExcelJS; }
async function pdf() { await script(LIBS.jspdf); await script(LIBS.autotable); return globalThis.jspdf.jsPDF; }

const nomeSeguro = s => String(s).replace(/[\\/:*?"<>|]/g, "_").replace(/\s+/g, "_");
const agoraTexto = () => new Date().toLocaleString("pt-PT", {timeZone: "Africa/Maputo"}).replace(",", "");
const hojeTexto = () => new Date().toLocaleDateString("pt-PT", {timeZone: "Africa/Maputo"});
const fmt = v => v == null || v === "" ? "" : new Intl.NumberFormat("pt-PT", {minimumFractionDigits: 2, maximumFractionDigits: 2, useGrouping: "always"}).format(v);
const mesCurto = m => NOME_MES[m - 1].slice(0, 3).toLowerCase();
const mesLongo = m => NOME_MES[m - 1].toLowerCase();

export function descarregar(blob, nome) {
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob); a.download = nome;
  document.body.append(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(a.href), 30000);
}

// ---------- folha salarial ----------
export const COLUNAS_FOLHA = [
  ["codigo", "Func."], ["nome", "Nome"], ["profissao", "Categoria"], ["base", "Vencimento\nBase"], ["diasFalta", "Dias de\nFalta"],
  ["valorFalta", "Valor de\nFalta"], ["vencimento", "Vencimento"], ["subsidios", "Sub. Férias"], ["outrasRem", "Outras\nRem."],
  ["totalRem", "Total Rem."], ["irps", "IRPS"], ["inss", "INSS"], ["outrosDesc", "Outros desc"], ["totalDesc", "Total Desc."], ["liquido", "Líq. MT"]
];
const tituloFolha = (ano, mes) => "Folha de Salarios do mês de " + NOME_MES[mes - 1] + " de " + ano;
const filtroTexto = ({empresa, departamento}) => departamento && departamento !== "TODOS" ? "Departamento: " + departamento : "";
const empresaTitulo = ({empresa}) => empresa && empresa !== "TODOS" ? empresa : "FOLHA SALARIAL";
export const nomeFicheiroFolha = (ano, mes, filtros, ext) =>
  nomeSeguro("Folha_Salarios_" + NOME_MES[mes - 1] + "_" + ano + (filtros.empresa && filtros.empresa !== "TODOS" ? "_" + filtros.empresa : "")) + "." + ext;

export async function folhaPDF(linhas, {ano, mes, filtros = {}, utilizador = ""}) {
  const JsPDF = await pdf();
  const doc = new JsPDF({orientation: "landscape", unit: "mm", format: "a4"});
  const W = doc.internal.pageSize.getWidth();
  doc.setFont("helvetica", "normal"); doc.setFontSize(9); doc.text(empresaTitulo(filtros), 5, 10);
  doc.setFont("helvetica", "bold"); doc.setFontSize(15); doc.text(tituloFolha(ano, mes), 5, 18);
  doc.setFontSize(9); doc.text(filtroTexto(filtros), 5, 24);
  const num = k => !["codigo", "nome", "profissao"].includes(k);
  const corpo = linhas.map(l => COLUNAS_FOLHA.map(([k]) => k === "diasFalta" ? (l[k] ? String(l[k]).replace(".", ",") : "") : num(k) ? fmt(l[k]) : l[k]));
  const totais = COLUNAS_FOLHA.map(([k], i) => i === 0 ? "Totais:" : num(k) && k !== "diasFalta" ? fmt(somar(linhas, k)) : "");
  doc.autoTable({
    startY: 28, margin: {left: 5, right: 5, top: 14, bottom: 12}, theme: "plain",
    head: [[{content: "Remunerações", colSpan: 10, styles: {halign: "center"}}, {content: "Descontos", colSpan: 4, styles: {halign: "center"}}, ""],
      COLUNAS_FOLHA.map(([, t]) => t)],
    body: corpo, foot: [totais], showFoot: "lastPage",
    styles: {font: "helvetica", fontSize: 7, cellPadding: 1.2, overflow: "linebreak"},
    headStyles: {fontStyle: "bold", halign: "center", valign: "bottom", lineWidth: {top: 0.4, bottom: 0.2}, lineColor: 40},
    footStyles: {fontStyle: "bold", lineWidth: {top: 0.4}, lineColor: 40},
    bodyStyles: {lineWidth: {bottom: 0.1}, lineColor: 210},
    columnStyles: Object.fromEntries(COLUNAS_FOLHA.map(([k], i) => [i, num(k) ? {halign: "right"} : {}]).concat([[1, {cellWidth: 44}], [2, {cellWidth: 34}]])),
    didParseCell: d => { if (d.section === "head" && d.row.index === 1 && d.column.index < 3) d.cell.styles.halign = "left"; }
  });
  let y = doc.lastAutoTable.finalY + 8;
  if (y > 175) { doc.addPage(); y = 20; }
  doc.setFont("helvetica", "normal"); doc.setFontSize(8);
  doc.text("Impresso por: " + utilizador + "  " + agoraTexto(), 5, y);
  doc.line(95, y + 12, 160, y + 12); doc.text("(O Dep Rec. Humanos)", 127.5, y + 16, {align: "center"});
  doc.line(185, y + 12, 250, y + 12); doc.text("(Aprovado por)", 217.5, y + 16, {align: "center"});
  const total = doc.getNumberOfPages();
  for (let p = 1; p <= total; p++) {
    doc.setPage(p); doc.setFontSize(7); doc.setFont("helvetica", "normal");
    doc.text(hojeTexto() + " | Pág. " + p + "/" + total, W - 5, 6, {align: "right"});
    doc.text("Gerado pelo Sistema de Gestão de Colaboradores", 5, doc.internal.pageSize.getHeight() - 5);
  }
  return doc.output("blob");
}

export async function folhaExcel(linhas, {ano, mes, filtros = {}, utilizador = ""}) {
  const ExcelJS = await excel();
  const wb = new ExcelJS.Workbook(); wb.creator = "Sistema RH SIR";
  const ws = wb.addWorksheet("Folha " + NOME_MES[mes - 1].slice(0, 3) + " " + ano, {
    pageSetup: {orientation: "landscape", paperSize: 9, fitToPage: true, fitToWidth: 1, fitToHeight: 0, horizontalCentered: true,
      margins: {left: 0.2, right: 0.2, top: 0.28, bottom: 0.28, header: 0.1, footer: 0.1}, printTitlesRow: "1:6"},
    headerFooter: {oddHeader: "&R" + hojeTexto() + " | Pág. &P/&N", oddFooter: "&LGerado pelo Sistema de Gestão de Colaboradores"}
  });
  const larguras = [8, 28, 23, 12, 8, 10, 12, 10, 11, 12, 9, 9, 11, 11, 12];
  larguras.forEach((w, i) => { ws.getColumn(i + 1).width = w; });
  const fonte = (c, extra = {}) => { c.font = {name: "Arial", size: 8, ...extra}; };
  ws.mergeCells("A1:O1"); ws.getCell("A1").value = empresaTitulo(filtros); fonte(ws.getCell("A1"), {size: 9});
  ws.mergeCells("A2:O2"); ws.getCell("A2").value = tituloFolha(ano, mes); fonte(ws.getCell("A2"), {size: 15, bold: true});
  ws.mergeCells("A3:O3"); ws.getCell("A3").value = filtroTexto(filtros); fonte(ws.getCell("A3"), {size: 9, bold: true});
  [1, 2, 3].forEach(r => { ws.getRow(r).height = 18; });
  ws.mergeCells("A5:J5"); ws.getCell("A5").value = "Remunerações";
  ws.mergeCells("K5:N5"); ws.getCell("K5").value = "Descontos";
  ["A5", "K5"].forEach(a => { fonte(ws.getCell(a), {bold: true}); ws.getCell(a).alignment = {horizontal: "center", vertical: "middle"}; });
  const cab = ws.getRow(6); cab.height = 30;
  COLUNAS_FOLHA.forEach(([, t], i) => {
    const c = cab.getCell(i + 1); c.value = t; fonte(c, {bold: true});
    c.alignment = {wrapText: true, vertical: "bottom", horizontal: i < 3 ? "left" : "center"};
    c.border = {top: {style: "medium"}, bottom: {style: "thin"}};
  });
  const primeira = 9;
  linhas.forEach((l, i) => {
    const r = ws.getRow(primeira + i);
    COLUNAS_FOLHA.forEach(([k], j) => {
      const c = r.getCell(j + 1);
      c.value = k === "valorFalta" && l[k] == null ? null : l[k];
      fonte(c);
      if (j >= 3) c.numFmt = j === 4 ? "0.##" : "#,##0.00";
      c.border = {bottom: {style: "thin", color: {argb: "FFD2D2D2"}}};
    });
  });
  const ultima = primeira + linhas.length - 1, tot = ultima + 1;
  const rt = ws.getRow(tot);
  rt.getCell(1).value = "Totais:";
  for (let c = 4; c <= 15; c++) if (c !== 5) {
    const col = ws.getColumn(c).letter;
    rt.getCell(c).value = {formula: "SUM(" + col + primeira + ":" + col + ultima + ")", result: somar(linhas, COLUNAS_FOLHA[c - 1][0])};
    rt.getCell(c).numFmt = "#,##0.00";
  }
  for (let c = 1; c <= 15; c++) { fonte(rt.getCell(c), {bold: true}); rt.getCell(c).border = {top: {style: "medium"}}; }
  const rod = tot + 2;
  ws.mergeCells(rod, 1, rod, 5); ws.getCell(rod, 1).value = "Impresso por: " + utilizador + "  " + agoraTexto(); fonte(ws.getCell(rod, 1));
  ws.mergeCells(rod + 2, 6, rod + 2, 9); ws.getCell(rod + 2, 6).value = "(O Dep Rec. Humanos)";
  ws.mergeCells(rod + 2, 11, rod + 2, 14); ws.getCell(rod + 2, 11).value = "(Aprovado por)";
  [6, 11].forEach(c => { const x = ws.getCell(rod + 2, c); fonte(x); x.alignment = {horizontal: "center"}; for (let k = c; k < c + 4; k++) ws.getCell(rod + 1, k).border = {bottom: {style: "thin"}}; });
  ws.pageSetup.printArea = "A1:O" + (rod + 6);
  return new Blob([await wb.xlsx.writeBuffer()], {type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"});
}

// ---------- relatório comparativo ----------
const cabecalhoComparativo = r => ["Unidade", "Trab. " + mesCurto(r.anterior.mes), mesCurto(r.anterior.mes) + " " + r.anterior.ano, "Variação",
  mesCurto(r.actual.mes) + " " + r.actual.ano, "Trab. " + mesCurto(r.actual.mes), "Peso %", "Comentários"];
const tituloComparativo = r => "Relatório Comparativo - " + mesLongo(r.anterior.mes) + " " + r.anterior.ano + " vs " + mesLongo(r.actual.mes) + " " + r.actual.ano;
export const nomeFicheiroComparativo = (r, ext) => "Relatorio_Comparativo_" + r.actual.ano + "_" + String(r.actual.mes).padStart(2, "0") + "." + ext;

export async function comparativoExcel(r) {
  const ExcelJS = await excel();
  const wb = new ExcelJS.Workbook(); wb.creator = "Sistema RH SIR";
  const ws = wb.addWorksheet("RELATORIO", {pageSetup: {orientation: "landscape", paperSize: 9, fitToPage: true, fitToWidth: 1, fitToHeight: 0}});
  ws.mergeCells("A1:H1"); ws.getCell("A1").value = tituloComparativo(r); ws.getCell("A1").font = {bold: true, size: 14};
  const cab = ws.getRow(3); cabecalhoComparativo(r).forEach((t, i) => { cab.getCell(i + 1).value = t; cab.getCell(i + 1).font = {bold: true}; });
  r.grupos.forEach((g, i) => {
    ws.getRow(4 + i).values = [g.unidade, g.qtdAnterior, g.valorAnterior, g.variacao, g.valorActual, g.qtdActual, g.peso, g.comentario];
  });
  const t = ws.getRow(4 + r.grupos.length);
  t.values = ["Total", r.qtdAnterior, r.totalAnterior, r.totalActual - r.totalAnterior, r.totalActual, r.qtdActual, r.totalActual ? 1 : null, ""];
  t.font = {bold: true}; t.eachCell(c => { c.border = {top: {style: "thin"}}; });
  [["A", 34], ["B", 10], ["C", 15], ["D", 14], ["E", 15], ["F", 10], ["G", 9], ["H", 70]].forEach(([c, w]) => { ws.getColumn(c).width = w; });
  ["C", "D", "E"].forEach(c => { ws.getColumn(c).numFmt = "#,##0.00"; });
  ws.getColumn("G").numFmt = "0.00%";
  ws.getColumn("H").alignment = {wrapText: true, vertical: "top"};
  return new Blob([await wb.xlsx.writeBuffer()], {type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"});
}

export async function comparativoPDF(r, {utilizador = ""} = {}) {
  const JsPDF = await pdf();
  const doc = new JsPDF({orientation: "landscape", unit: "mm", format: "a4"});
  doc.setFont("helvetica", "bold"); doc.setFontSize(14); doc.text(tituloComparativo(r), 8, 14);
  const pct = v => v == null ? "" : (v * 100).toFixed(2).replace(".", ",") + "%";
  doc.autoTable({
    startY: 20, margin: {left: 8, right: 8}, theme: "grid",
    head: [cabecalhoComparativo(r)],
    body: r.grupos.map(g => [g.unidade, g.qtdAnterior, fmt(g.valorAnterior), fmt(g.variacao), fmt(g.valorActual), g.qtdActual, pct(g.peso), g.comentario]),
    foot: [["Total", r.qtdAnterior, fmt(r.totalAnterior), fmt(r.totalActual - r.totalAnterior), fmt(r.totalActual), r.qtdActual, r.totalActual ? "100,00%" : "", ""]],
    styles: {fontSize: 7, cellPadding: 1.5, valign: "top"}, headStyles: {fillColor: [190, 37, 55]}, footStyles: {fillColor: [240, 242, 245], textColor: 20, fontStyle: "bold"},
    columnStyles: {0: {cellWidth: 45}, 1: {halign: "right", cellWidth: 14}, 2: {halign: "right", cellWidth: 24}, 3: {halign: "right", cellWidth: 22}, 4: {halign: "right", cellWidth: 24}, 5: {halign: "right", cellWidth: 14}, 6: {halign: "right", cellWidth: 16}}
  });
  doc.setFont("helvetica", "normal"); doc.setFontSize(8);
  doc.text("Gerado por " + utilizador + " em " + agoraTexto() + " · Base: Salário Líquido (tbl_lançamento)", 8, doc.internal.pageSize.getHeight() - 6);
  return doc.output("blob");
}
