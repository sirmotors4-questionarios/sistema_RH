import {authConfigured, initAuth, currentAccount, signIn, signOut, getToken} from "./rh-auth.js";
import {loadModulo} from "./rh-api.js";
import * as M from "./rh-modulos.js";
"use strict";

// Sem sessão Microsoft: dados fictícios. Com sessão: dados do Excel via RH_API_Consultar (só leitura).
const mockEmployees = [
 {codigo:"RH-001",nome:"Colaborador Exemplo A",empresa:"Empresa A",departamento:"Administração",regime:"Mensal",estado:"Activo"},
 {codigo:"RH-002",nome:"Colaborador Exemplo B",empresa:"Empresa B",departamento:"Operações",regime:"Diário",estado:"Activo"},
 {codigo:"RH-003",nome:"Colaborador Exemplo C",empresa:"Empresa A",departamento:"Manutenção",regime:"Hora",estado:"Inactivo"},
 {codigo:"RH-004",nome:"Colaborador Exemplo D",empresa:"Empresa C",departamento:"Financeiro",regime:"Mensal",estado:"Activo"}
];
const MAX_LINHAS = 300;
const $ = id => document.getElementById(id);
const labels={dashboard:"Visão geral",colaboradores:"Colaboradores",contratos:"Contratos",adendas:"Adendas",salarios:"Lançamento salarial",ferias:"Férias",relatorios:"Relatórios",configuracoes:"Configurações"};
// Página -> módulos do fluxo de que precisa.
const precisa={dashboard:["funcionarios","contratos","lancamentos","ferias"],colaboradores:["funcionarios"],contratos:["contratos","funcionarios"],adendas:["contratos"],salarios:["lancamentos"],ferias:["ferias"]};

const dados={};          // módulo -> {tabelas, carregado}
const pedidos={};        // módulo -> Promise em curso
let ligado=false, paginaActual="dashboard";

function go(page){
 if(page==="configuracoes"){window.location.href="configuracoes.html";return;}
 if(!Object.hasOwn(labels,page))return;
 paginaActual=page;
 document.querySelectorAll(".page").forEach(el=>el.classList.toggle("active",el.id===page));
 document.querySelectorAll("[data-page]").forEach(el=>el.classList.toggle("active",el.dataset.page===page));
 $("crumb").textContent=labels[page];
 $("sidebar").classList.remove("open");
 window.scrollTo({top:0,behavior:"instant"});
 if(ligado)(precisa[page]||[]).forEach(m=>carregar(m));
}
document.querySelectorAll("[data-page]").forEach(el=>el.addEventListener("click",()=>go(el.dataset.page)));
document.querySelectorAll("[data-go]").forEach(el=>el.addEventListener("click",()=>go(el.dataset.go)));
$("menu-toggle").addEventListener("click",()=>$("sidebar").classList.toggle("open"));
function showToast(message){const t=$("toast");t.textContent=message;t.classList.add("show");clearTimeout(showToast.timer);showToast.timer=setTimeout(()=>t.classList.remove("show"),4500)}
document.querySelectorAll("[data-action]").forEach(el=>el.addEventListener("click",()=>showToast("Registo de alterações ainda não activado: a gravação com aprovação (RH_API_Actualizar) é a próxima fase.")));

// ---------- utilitários de tabela ----------
function cell(row,value,opts={}){const td=document.createElement("td");if(opts.num)td.className="num";const node=document.createElement(opts.strong?"strong":"span");node.textContent=value??"";td.append(node);row.append(td);return td}
function stateCell(row,estado){const td=document.createElement("td"),b=document.createElement("span");b.className="state"+(M.activo(estado)?"":" off");b.textContent=estado||"—";td.append(b);row.append(td)}
function emptyRow(body,cols,text,cls="no-data"){const tr=document.createElement("tr"),td=document.createElement("td");td.colSpan=cols;td.className=cls;td.textContent=text;tr.append(td);body.append(tr)}
function fillSelect(select,valores,primeiro){const atual=select.value;select.replaceChildren();const o=document.createElement("option");o.value="";o.textContent=primeiro;select.append(o);valores.forEach(v=>{const op=document.createElement("option");op.value=v;op.textContent=v;select.append(op)});if(valores.includes(atual))select.value=atual}
function setSource(modulo,text){document.querySelectorAll('[data-source="'+modulo+'"]').forEach(el=>el.textContent=text)}
function showLoaded(page,yes){document.querySelectorAll('[data-empty="'+page+'"]').forEach(el=>el.hidden=yes);document.querySelectorAll('[data-loaded="'+page+'"]').forEach(el=>el.hidden=!yes)}
const hoje=()=>new Date().toLocaleDateString("en-CA",{timeZone:"Africa/Maputo"});
const fmt=n=>new Intl.NumberFormat("pt-MZ",{minimumFractionDigits:2,maximumFractionDigits:2}).format(n);

// ---------- carregamento ----------
function carregar(modulo,forcar=false){
 if(!forcar&&dados[modulo])return Promise.resolve(dados[modulo]);
 if(pedidos[modulo])return pedidos[modulo];
 setSource(modulo,"A carregar do Excel…");
 pedidos[modulo]=loadModulo(modulo,{getToken}).then(r=>{
  dados[modulo]={tabelas:r.tabelas,carregado:new Date()};
  if(r.utilizador)mostrarUtilizador(r.utilizador);
  setSource(modulo,"Fonte: Gestão de Colaboradores.xlsx via Power Automate · "+dados[modulo].carregado.toLocaleTimeString("pt-MZ",{hour:"2-digit",minute:"2-digit"}));
  render();
  return dados[modulo];
 }).catch(e=>{
  setSource(modulo,"Erro: "+e.message);
  estadoLigacao("error","Não foi possível carregar",e.message);
  throw e;
 }).finally(()=>{delete pedidos[modulo]});
 pedidos[modulo].catch(()=>{});
 return pedidos[modulo];
}
function linhas(modulo,tabela){return dados[modulo]?.tabelas[tabela]||[]}

// ---------- páginas ----------
function renderEmployees(){
 const lista=dados.funcionarios?M.colaboradores(linhas("funcionarios","tbl_funcionarios")):mockEmployees;
 const filtro=$("employee-filter");
 if(dados.funcionarios&&!filtro.dataset.real){fillSelect(filtro,M.contagem(lista,c=>c.estado).map(([k])=>k),"Todos os estados");filtro.dataset.real="1"}
 const result=M.filtrarColaboradores(lista,$("employee-search").value,filtro.value);
 const tbody=$("employees-body");tbody.replaceChildren();
 if(!result.length)emptyRow(tbody,7,"Nenhum colaborador corresponde aos filtros.");
 result.slice(0,MAX_LINHAS).forEach(e=>{const tr=document.createElement("tr");cell(tr,e.codigo);cell(tr,e.nome,{strong:true});cell(tr,e.empresa);cell(tr,e.departamento);cell(tr,e.regime);stateCell(tr,e.estado);
  const td=document.createElement("td"),btn=document.createElement("button");btn.className="inline-btn";btn.textContent="Consultar";btn.addEventListener("click",()=>showToast(e.nome+" · "+[e.categoria||e.profissao,e.contrato,e.admissao?"admitido em "+M.formatarData(e.admissao):""].filter(Boolean).join(" · ")));td.append(btn);tr.append(td);tbody.append(tr)});
 $("employee-count").textContent=dados.funcionarios
  ?result.length+" de "+lista.length+" colaborador(es)"+(result.length>MAX_LINHAS?" · a mostrar "+MAX_LINHAS+", refine a pesquisa":"")
  :result.length+" registo(s) demonstrativo(s)";
}
function renderContratos(){
 if(!dados.contratos)return;
 showLoaded("contratos",true);
 const body=$("params-body");body.replaceChildren();
 const params=M.parametrosVigentes(linhas("contratos","tbl_parametros_pagamento"),hoje());
 if(!params.length)emptyRow(body,6,"Sem parâmetros vigentes hoje.");
 params.forEach(p=>{const tr=document.createElement("tr");cell(tr,p.regime,{strong:true});cell(tr,p.categoria||"Todas");cell(tr,M.formatarValor(p.valor),{num:true});cell(tr,p.unidade);cell(tr,M.formatarData(p.inicio));cell(tr,p.fim?M.formatarData(p.fim):"—");body.append(tr)});
 const activos=M.colaboradores(linhas("funcionarios","tbl_funcionarios")).filter(c=>M.activo(c.estado));
 [["contract-types",c=>c.contrato],["regimes",c=>c.regime]].forEach(([id,fn])=>{const b=$(id);b.replaceChildren();
  if(!dados.funcionarios){emptyRow(b,2,"A carregar colaboradores…","loading");return}
  M.contagem(activos,fn).forEach(([k,n])=>{const tr=document.createElement("tr");cell(tr,k);cell(tr,String(n),{num:true});b.append(tr)})});
}
function renderAdendas(){
 if(!dados.contratos)return;
 showLoaded("adendas",true);
 const lista=M.adendas(linhas("contratos","tbl_adendas"));
 const filtro=$("addenda-filter");
 if(!filtro.dataset.real){fillSelect(filtro,M.contagem(lista,a=>a.estado).map(([k])=>k),"Todos os estados");filtro.dataset.real="1"}
 const termo=$("addenda-search").value.trim().toLocaleLowerCase("pt"),estado=filtro.value;
 const result=lista.filter(a=>(!estado||a.estado===estado)&&(!termo||[a.codigo,a.nome,a.campo,a.numero].some(x=>x.toLocaleLowerCase("pt").includes(termo))));
 const body=$("addenda-body");body.replaceChildren();
 if(!result.length)emptyRow(body,9,"Nenhuma adenda corresponde aos filtros.");
 result.slice(0,MAX_LINHAS).forEach(a=>{const tr=document.createElement("tr");cell(tr,a.numero||a.id);cell(tr,[a.codigo,a.nome].filter(Boolean).join(" · "),{strong:true});cell(tr,a.campo);cell(tr,a.anterior);cell(tr,a.novo);cell(tr,M.formatarData(a.inicio));cell(tr,a.fim?M.formatarData(a.fim):"—");cell(tr,a.estado);cell(tr,a.aplicada);body.append(tr)});
 $("addenda-count").textContent=result.length+" de "+lista.length+" adenda(s)"+(result.length>MAX_LINHAS?" · a mostrar "+MAX_LINHAS:"");
}
let periodosPreenchidos=false;
function renderSalarios(){
 if(!dados.lancamentos)return;
 const todas=linhas("lancamentos","tbl_lançamento");
 if(!periodosPreenchidos){const ps=M.periodos(todas);if(ps.length){const anos=[...new Set(ps.map(p=>p.ano))];year.replaceChildren();anos.forEach(a=>{const op=document.createElement("option");op.value=String(a);op.textContent=String(a);year.append(op)});year.value=String(ps[0].ano);month.value=String(ps[0].mes)}periodosPreenchidos=true}
 showLoaded("lancamentos",true);
 const lista=M.lancamentos(todas,Number(year.value),Number(month.value));
 const termo=$("pay-search").value.trim().toLocaleLowerCase("pt");
 const result=lista.filter(l=>!termo||[l.codigo,l.nome,l.empresa].some(x=>x.toLocaleLowerCase("pt").includes(termo)));
 $("pay-badge").textContent=M.NOME_MES[Number(month.value)-1]+" "+year.value+" · "+lista.length+" lançamento(s)";
 $("p-count").textContent=String(lista.length);
 $("p-states").textContent=M.contagem(lista,l=>l.estado).map(([k,n])=>k+": "+n).join(" · ")||"Sem lançamentos neste mês";
 $("p-acr").textContent=fmt(M.totais(lista,"acrescimos"));$("p-desc").textContent=fmt(M.totais(lista,"descontos"));$("p-liq").textContent=fmt(M.totais(lista,"liquido"));
 const body=$("pay-body");body.replaceChildren();
 if(!result.length)emptyRow(body,10,lista.length?"Nenhum lançamento corresponde à pesquisa.":"Sem lançamentos para este mês.");
 result.slice(0,MAX_LINHAS).forEach(l=>{const tr=document.createElement("tr");cell(tr,l.codigo);cell(tr,l.nome,{strong:true});cell(tr,l.empresa);cell(tr,l.regime);cell(tr,M.formatarValor(l.dias).replace(/,00$/,""),{num:true});cell(tr,M.formatarValor(l.acrescimos),{num:true});cell(tr,M.formatarValor(l.descontos),{num:true});cell(tr,M.formatarValor(l.irps),{num:true});cell(tr,M.formatarValor(l.liquido),{num:true,strong:true});cell(tr,l.estado);body.append(tr)});
 $("pay-count").textContent=result.length+" lançamento(s)"+(result.length>MAX_LINHAS?" · a mostrar "+MAX_LINHAS:"");
}
function renderFerias(){
 if(!dados.ferias)return;
 showLoaded("ferias",true);
 const reg=linhas("ferias","tbl_ferias"),cols=M.colunas(reg).slice(0,8),head=$("leave-head"),body=$("leave-body");
 head.replaceChildren();body.replaceChildren();
 (cols.length?cols:["Registos"]).forEach(c=>{const th=document.createElement("th");th.textContent=c;head.append(th)});
 if(!reg.length)emptyRow(body,1,"Ainda não existem registos de férias na tbl_ferias.");
 reg.slice(0,MAX_LINHAS).forEach(r=>{const tr=document.createElement("tr");cols.forEach(c=>cell(tr,/data|inicio|fim/i.test(c)?M.formatarData(r[c]):r[c]));body.append(tr)});
 const fb=$("holiday-body");fb.replaceChildren();
 const fs=M.feriados(linhas("ferias","tbl_feriados"));
 if(!fs.length)emptyRow(fb,2,"Sem feriados registados.");
 fs.forEach(f=>{const tr=document.createElement("tr");cell(tr,String(f.dia).padStart(2,"0")+"/"+String(f.mes).padStart(2,"0"));cell(tr,f.nome);fb.append(tr)});
}
function renderDashboard(){
 if(dados.funcionarios){const l=M.colaboradores(linhas("funcionarios","tbl_funcionarios"));$("m-activos").textContent=String(l.filter(c=>M.activo(c.estado)).length);$("m-activos-sub").textContent="de "+l.length+" registados"}
 if(dados.contratos){const a=M.adendas(linhas("contratos","tbl_adendas"));$("m-adendas").textContent=String(a.length);$("m-adendas-sub").textContent=a.filter(x=>M.activo(x.estado)||/vigor/i.test(x.estado)).length+" em vigor"}
 if(dados.ferias){const n=linhas("ferias","tbl_ferias").length;$("m-ferias").textContent=String(n);$("m-ferias-sub").textContent=n?"na tbl_ferias":"Sem registos"}
 if(dados.lancamentos){const p=M.periodos(linhas("lancamentos","tbl_lançamento"))[0];if(p){const l=M.lancamentos(linhas("lancamentos","tbl_lançamento"),p.ano,p.mes);$("m-mes").textContent=M.NOME_MES[p.mes-1].slice(0,3)+" "+p.ano;$("m-mes-sub").textContent=l.length+" lançamentos · "+(M.contagem(l,x=>x.estado)[0]?.[0]||"")}}
 if(Object.keys(dados).length)$("dash-badge").textContent="Dados do Excel";
}
function render(){renderEmployees();renderContratos();renderAdendas();renderSalarios();renderFerias();renderDashboard()}

$("employee-search").addEventListener("input",renderEmployees);
$("employee-filter").addEventListener("change",renderEmployees);
$("addenda-search").addEventListener("input",renderAdendas);
$("addenda-filter").addEventListener("change",renderAdendas);
$("pay-search").addEventListener("input",renderSalarios);
document.querySelectorAll("[data-reload]").forEach(b=>b.addEventListener("click",()=>carregar(b.dataset.reload,true)));
const month=$("pay-month"),year=$("pay-year");
M.NOME_MES.forEach((m,i)=>{const op=document.createElement("option");op.value=String(i+1);op.textContent=m;month.appendChild(op)});
const now=new Date();month.value=String(now.getMonth()+1);
for(let y=now.getFullYear()-2;y<=now.getFullYear()+1;y++){const op=document.createElement("option");op.value=String(y);op.textContent=String(y);year.appendChild(op)}year.value=String(now.getFullYear());
[month,year].forEach(s=>s.addEventListener("change",renderSalarios));
$("load-month").addEventListener("click",()=>{if(!ligado){showToast("Inicie sessão com a conta Microsoft da SIR para consultar os lançamentos.");return}carregar("lancamentos",true)});

// ---------- sessão ----------
function estadoLigacao(tipo,titulo,texto){
 ["source-dot","status-pill","auth-notice"].forEach(id=>{$(id).classList.toggle("live",tipo==="live");$(id).classList.toggle("error",tipo==="error")});
 $("notice-title").textContent=titulo;$("notice-text").textContent=texto;
}
function mostrarUtilizador(u){
 $("avatar").textContent=(u.nome||u.email||"RH").split(/\s+/).map(p=>p[0]).join("").slice(0,2).toUpperCase();
 $("avatar").title=(u.nome||u.email)+" · "+u.perfil;
 $("source-sub").textContent=u.perfil+(u.propor?"":" · só consulta");
}
function botoes(){const signed=Boolean(currentAccount());$("signin").hidden=signed||!authConfigured();$("signout").hidden=!signed;document.querySelectorAll("[data-reload]").forEach(b=>b.hidden=!signed)}
$("signin").addEventListener("click",()=>signIn().catch(e=>estadoLigacao("error","Sessão não iniciada",e.message)));
$("signout").addEventListener("click",()=>signOut().catch(e=>estadoLigacao("error","Erro ao terminar sessão",e.message)));

render();
(async()=>{
 if(!authConfigured()){botoes();return}
 try{
  await initAuth();botoes();
  const conta=currentAccount();
  if(!conta){estadoLigacao("","Sessão Microsoft não iniciada","Os números exibidos são fictícios. Entre com a conta Microsoft da SIR para ver os dados do Excel (só leitura).");return}
  ligado=true;
  $("status-pill").textContent="● Ligado ao Excel (só leitura)";$("source-title").textContent="Ligado ao SharePoint";$("source-sub").textContent=conta.username;
  estadoLigacao("live","Sessão iniciada: "+conta.username,"Os dados vêm do ficheiro Gestão de Colaboradores.xlsx através do Power Automate. Nesta fase o portal só consulta; não grava alterações.");
  (precisa[paginaActual]||[]).forEach(m=>carregar(m));
 }catch(e){estadoLigacao("error","Erro de autenticação",e.message)}
})();
