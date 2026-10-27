import {authConfigured, initAuth, currentAccount, signIn, signOut, getToken} from "./rh-auth.js";
import {iniciarRelatorios} from "./relatorios-ui.js";
import {iniciarPainel} from "./painel.js";
import {loadModulo, gravarColaborador, gravarNovoMes, gravarLancamento, finalizarMes} from "./rh-api.js";
import {urlGravacao, guardarUrlGravacao, urlGuardada, sessaoAcesso, entrarComCodigo, sairAcesso, chaveAcesso, validarEmail, validarCodigo, CODIGO_MINIMO} from "./rh-acesso.js";
import * as M from "./rh-modulos.js";
import {criarFicha} from "./ficha.js";
import {iniciarLancamentos} from "./lancamentos-ui.js";
import {unirAdendas} from "./rh-folha.js";
"use strict";

// Sem sessão Microsoft: dados fictícios. Com sessão: dados do Excel via RH_API_Consultar (só leitura).
const mockEmployees = [
 {codigo:"0001",nome:"Colaborador Exemplo A",empresa:"Empresa A",departamento:"Administração",profissao:"Técnico administrativo",regime:"Mensal",estado:"Activo"},
 {codigo:"0002",nome:"Colaborador Exemplo B",empresa:"Empresa B",departamento:"Operações",profissao:"Motorista",regime:"Diário",estado:"Activo"},
 {codigo:"0001",nome:"Colaborador Exemplo C",empresa:"Empresa B",departamento:"Manutenção",profissao:"Mecânico",regime:"Hora",estado:"Inactivo"},
 {codigo:"0004",nome:"Colaborador Exemplo D",empresa:"Empresa C",departamento:"Financeiro",profissao:"Contabilista",regime:"Mensal",estado:"Activo"}
];
const MAX_LINHAS = 300;
const $ = id => document.getElementById(id);
const labels={dashboard:"Visão geral",painel:"Dashboard RH",colaboradores:"Colaboradores",contratos:"Contratos",adendas:"Adendas",salarios:"Lançamento salarial",ferias:"Férias",relatorios:"Relatórios",configuracoes:"Configurações"};
// Página -> módulos do fluxo de que precisa.
const precisa={dashboard:["funcionarios","contratos","lancamentos","ferias"],colaboradores:["funcionarios"],contratos:["contratos","funcionarios"],adendas:["contratos"],salarios:["lancamentos"],ferias:["ferias"]};

const dados={};          // módulo -> {tabelas, carregado}
const pedidos={};        // módulo -> Promise em curso
let ligado=false, paginaActual="dashboard", utilizadorActual=null;

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
 if(page==="painel")painel.abrir();
 if(page==="relatorios")relatorios.abrir();
}
document.querySelectorAll("[data-page]").forEach(el=>el.addEventListener("click",()=>go(el.dataset.page)));
document.querySelectorAll("[data-go]").forEach(el=>el.addEventListener("click",()=>go(el.dataset.go)));
$("menu-toggle").addEventListener("click",()=>$("sidebar").classList.toggle("open"));
function showToast(message){const t=$("toast");t.textContent=message;t.classList.add("show");clearTimeout(showToast.timer);showToast.timer=setTimeout(()=>t.classList.remove("show"),4500)}
// ---------- lançamentos gravados em segundo plano: o Excel Online demora ~15–20 s por gravação ----------
// As gravações vão em fila (uma de cada vez no mesmo ficheiro) e o mês só é relido quando a fila do mês acaba:
// assim nenhuma alteração se perde nem é tapada por uma leitura feita antes de ela chegar ao Excel.
const pendentes=new Map(); // "ano-mes|Empresa|Codigo" -> {inicio, emFila, fim?, erro?}
let filaGravacao=Promise.resolve();
const segs=ms=>Math.max(0,Math.round(ms/1000));
function rotuloPendente(p){if(p.erro)return"Erro ao gravar";if(p.fim)return"Gravado · "+segs(p.fim-p.inicio)+" s";return(p.emFila?"Na fila… ":"A gravar… ")+segs(Date.now()-p.inicio)+" s"}
function gravacaoActiva(k){const p=pendentes.get(k);return Boolean(p&&!p.fim&&!p.erro)}
function actualizarPendentes(){
 document.querySelectorAll("[data-pendente]").forEach(b=>{const p=pendentes.get(b.dataset.pendente);if(p)b.textContent=rotuloPendente(p)});
 if(![...pendentes.keys()].some(gravacaoActiva)){clearInterval(actualizarPendentes.timer);actualizarPendentes.timer=null}
}
async function relerMesSeLivre(ano,mes){
 const pre=ano+"-"+mes+"|";
 if([...pendentes.keys()].some(k=>k.startsWith(pre)&&gravacaoActiva(k)))return; // ainda há gravações deste mês na fila
 try{await pedidos.lancamentos}catch{} // uma leitura já em curso foi pedida antes da gravação: espera e lê de novo
 meses.delete(ano+"-"+mes);mesesCompletos=false;
 try{await carregarLancamentos({ano,mes})}catch{}
 if([...pendentes.keys()].some(k=>k.startsWith(pre)&&gravacaoActiva(k)))return;
 for(const[k,p]of pendentes)if(k.startsWith(pre)&&p.fim&&!p.erro)pendentes.delete(k);
 renderSalarios();
}
function gravarEmFundo({ano,mes,chave,pedido}){
 const k=ano+"-"+mes+"|"+chave,quem=chave.replace("|"," | ");
 if(gravacaoActiva(k)){showToast("Este lançamento ainda está a ser gravado.");return}
 const naFila=[...pendentes.keys()].filter(gravacaoActiva).length;
 const p={inicio:Date.now(),emFila:naFila>0};pendentes.set(k,p);renderSalarios();
 showToast(naFila?"Na fila: "+quem+" (grava a seguir às "+naFila+" anterior(es)).":"A gravar "+quem+"… O Excel demora cerca de 15–20 s; pode continuar a trabalhar.");
 if(!actualizarPendentes.timer)actualizarPendentes.timer=setInterval(actualizarPendentes,1000);
 filaGravacao=filaGravacao.then(async()=>{
  p.emFila=false;actualizarPendentes();
  try{
   const res=await gravarLancamento(pedido,{credenciais:sessaoAcesso(),url:urlGravacao()});
   p.fim=Date.now();
   showToast("Alteração gravada: "+quem+" em "+segs(p.fim-p.inicio)+" s. Fórmulas recalculadas no Excel.");
   if(actualizarLinha(ano,mes,chave,res.linha)){pendentes.delete(k);renderSalarios();actualizarPendentes()}
   else{actualizarPendentes();p.reler=true;relerMesSeLivre(ano,mes)}
  }catch(e){
   p.erro=e.message||String(e);actualizarPendentes();renderSalarios();
   showToast("Não gravado ("+quem+", "+segs(Date.now()-p.inicio)+" s): "+p.erro);
   if([...pendentes.values()].some(x=>x.reler))relerMesSeLivre(ano,mes);
  }
 });
}
// Substitui na cache só a linha gravada, com os valores recalculados que o Office Script devolve.
// As outras linhas (e as que ainda estão a gravar) ficam como estão. Devolve false se não houver linha.
function actualizarLinha(ano,mes,chave,linhaExcel){
 if(!linhaExcel)return false;
 const igual=l=>M.numero(M.campo(l,"Ano"))===ano&&M.mes(M.campo(l,"Mês","Mes"))===mes&&M.chaveRH(M.campo(l,"Empresa"),M.campo(l,"Codigo"))===chave;
 const listas=[dados.lancamentos?.tabelas?.["tbl_lançamento"],meses.get(ano+"-"+mes)].filter(Array.isArray);
 let feito=false;
 for(const lista of listas)for(const l of lista)if(igual(l)){
  for(const[h,v]of Object.entries(linhaExcel))l[h]=v==null?"":(typeof v==="number"?String(v):v);
  feito=true;
 }
 return feito;
}
document.querySelectorAll("[data-action]").forEach(el=>el.addEventListener("click",()=>showToast("Gravação ainda não activada neste módulo. Já disponível: criar e editar colaboradores.")));

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
 if(modulo==="lancamentos")return carregarLancamentos(forcar?{ano:Number(year.value),mes:Number(month.value)}:null);
 if(!forcar&&dados[modulo])return Promise.resolve(dados[modulo]);
 if(pedidos[modulo])return pedidos[modulo];
 setSource(modulo,modulo==="lancamentos"?"A carregar do Excel (1800+ lançamentos; pode demorar cerca de 1 minuto)…":"A carregar do Excel…");
 pedidos[modulo]=loadModulo(modulo,transporte()).then(r=>{
  dados[modulo]={tabelas:r.tabelas,carregado:new Date()};
  if(r.utilizador)mostrarUtilizador(r.utilizador);
  setSource(modulo,"Fonte: Gestão de Colaboradores.xlsx via Power Automate · "+dados[modulo].carregado.toLocaleTimeString("pt-MZ",{hour:"2-digit",minute:"2-digit"}));
  render();
  return dados[modulo];
 }).catch(e=>{
  setSource(modulo,"Erro: "+e.message);
  estadoLigacao("error","Não foi possível carregar",e.message);
  if(e.code==="RH_FORBIDDEN"&&sessaoAcesso()){sairAcesso();desligar();abrirAcesso(e.message)}
  throw e;
 }).finally(()=>{delete pedidos[modulo]});
 pedidos[modulo].catch(()=>{});
 return pedidos[modulo];
}
// Lançamentos: o fluxo devolve só o mês pedido. Sem período indicado, procura o mês actual
// e recua até 6 meses até encontrar lançamentos.
function carregarLancamentos(periodo){
 if(!periodo&&dados.lancamentos)return Promise.resolve(dados.lancamentos);
 if(pedidos.lancamentos)return pedidos.lancamentos;
 const procurar=!periodo;
 let p=periodo||(()=>{const d=new Date();return{ano:d.getFullYear(),mes:d.getMonth()+1}})();
 const tarefa=(async()=>{
  for(let tentativa=0;tentativa<(procurar?7:1);tentativa++){
   setSource("lancamentos","A carregar "+M.NOME_MES[p.mes-1]+" "+p.ano+" do Excel…");
   const r=await loadModulo("lancamentos",{...transporte(),params:{ano:p.ano,mes:p.mes}});
   if(r.utilizador)mostrarUtilizador(r.utilizador);
   const todas=r.tabelas["tbl_lançamento"]||[];
   const n=M.lancamentos(todas,p.ano,p.mes).length;
   // Se o fluxo devolveu outros meses (sem filtro no servidor), usa o mais recente sem novo pedido.
   if(!n&&procurar){const ultimo=M.periodos(todas)[0];if(ultimo){p=ultimo;tentativa=6}}
   if(n||!procurar||tentativa===6){
    dados.lancamentos={tabelas:r.tabelas,periodo:p,carregado:new Date()};
    guardarMeses(todas);meses.set(p.ano+"-"+p.mes,doPeriodo(todas,p.ano,p.mes));
    year.value=String(p.ano);month.value=String(p.mes);
    setSource("lancamentos","Fonte: Gestão de Colaboradores.xlsx via Power Automate · "+dados.lancamentos.carregado.toLocaleTimeString("pt-MZ",{hour:"2-digit",minute:"2-digit"}));
    render();return dados.lancamentos;
   }
   p=p.mes===1?{ano:p.ano-1,mes:12}:{ano:p.ano,mes:p.mes-1};
  }
 })().catch(e=>{
  setSource("lancamentos","Erro: "+e.message);
  estadoLigacao("error","Não foi possível carregar",e.message);
  if(e.code==="RH_FORBIDDEN"&&sessaoAcesso()){sairAcesso();desligar();abrirAcesso(e.message)}
  throw e;
 }).finally(()=>{delete pedidos.lancamentos});
 pedidos.lancamentos=tarefa;tarefa.catch(()=>{});
 return tarefa;
}
function linhas(modulo,tabela){return dados[modulo]?.tabelas[tabela]||[]}
// Cache de meses de lançamento (linhas cruas do Excel), para comparativo e Novo mês.
// Se o fluxo devolver vários meses de uma vez (leitura sem filtro), todos ficam em cache.
const meses=new Map();let mesesCompletos=false;
const doPeriodo=(rows,ano,mes)=>rows.filter(l=>M.numero(M.campo(l,"Ano"))===ano&&M.mes(M.campo(l,"Mês","Mes"))===mes);
function guardarMeses(rows){
 const por=new Map();for(const l of rows){const a=M.numero(M.campo(l,"Ano")),m=M.mes(M.campo(l,"Mês","Mes"));if(!a||!m)continue;const k=a+"-"+m;if(!por.has(k))por.set(k,[]);por.get(k).push(l)}
 if(por.size>1){mesesCompletos=true;meses.clear()}
 por.forEach((v,k)=>meses.set(k,v));
}
async function carregarMes(ano,mes,forcar=false){
 const k=ano+"-"+mes;
 if(!forcar){if(meses.has(k))return meses.get(k);if(mesesCompletos&&meses.size)return []}
 const r=await loadModulo("lancamentos",{...transporte(),params:{ano,mes}});
 const todas=r.tabelas["tbl_lançamento"]||[];guardarMeses(todas);
 const rows=doPeriodo(todas,ano,mes);meses.set(k,rows);return rows;
}

// ---------- páginas ----------
function cascata(lista,selEmp,selDep){
 const o=M.opcoesFiltro(lista,selEmp.value);
 fillSelect(selEmp,o.empresas,"Todas");fillSelect(selDep,o.departamentos,"Todos");
}
function renderEmployees(){
 const lista=dados.funcionarios?M.colaboradores(linhas("funcionarios","tbl_funcionarios")):mockEmployees;
 const emp=$("employee-company"),dep=$("employee-department"),filtro=$("employee-filter");
 cascata(lista,emp,dep);
 const result=M.filtrarColaboradores(lista,$("employee-search").value,filtro.value,{empresa:emp.value,departamento:dep.value});
 const tbody=$("employees-body");tbody.replaceChildren();
 if(!result.length)emptyRow(tbody,8,"Nenhum colaborador corresponde aos filtros.");
 result.slice(0,MAX_LINHAS).forEach(e=>{const tr=document.createElement("tr");cell(tr,e.codigo);cell(tr,e.nome,{strong:true});cell(tr,e.empresa);cell(tr,e.departamento);cell(tr,e.profissao);cell(tr,e.regime);stateCell(tr,e.estado);
  const td=document.createElement("td"),btn=document.createElement("button");btn.type="button";btn.className="inline-btn";btn.textContent="Abrir ficha";btn.setAttribute("aria-label","Abrir ficha de "+e.nome+" ("+e.codigo+" · "+e.empresa+")");
  const abrir=()=>{if(!e.linha){showToast("Ficha disponível depois de entrar com o código de acesso.");return}ficha.ver(e.linha)};
  btn.addEventListener("click",ev=>{ev.stopPropagation();abrir()});tr.addEventListener("click",abrir);td.append(btn);tr.append(td);tbody.append(tr)});
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
function renderSalarios(){
 if(!dados.lancamentos)return;
 const todas=linhas("lancamentos","tbl_lançamento");
 showLoaded("lancamentos",true);
 const per=dados.lancamentos.periodo||{ano:Number(year.value),mes:Number(month.value)};
 const lista=M.lancamentos(todas,per.ano,per.mes);
 const emp=$("pay-company"),dep=$("pay-department"),col=$("pay-employee");cascata(lista,emp,dep);
 {const atual=col.value;const op=lista.filter(l=>(!emp.value||l.empresa===emp.value)&&(!dep.value||l.departamento===dep.value)).sort((a,b)=>a.codigo.localeCompare(b.codigo));
  col.replaceChildren(new Option("Todos",""),...op.map(l=>new Option(l.codigo+" · "+l.nome,l.codigo)));if(op.some(l=>l.codigo===atual))col.value=atual}
 const result=M.filtrarColaboradores(lista,$("pay-search").value,"",{empresa:emp.value,departamento:dep.value}).filter(l=>!col.value||l.codigo===col.value);
 $("pay-badge").textContent=M.NOME_MES[per.mes-1]+" "+per.ano+" · "+lista.length+" lançamento(s)";
 $("p-count").textContent=String(result.length);
 $("p-states").textContent=M.contagem(result,l=>l.estado).map(([k,n])=>k+": "+n).join(" · ")||"Sem lançamentos";
 $("p-acr").textContent=fmt(M.totais(result,"acrescimos"));$("p-desc").textContent=fmt(M.totais(result,"descontos"));$("p-liq").textContent=fmt(M.totais(result,"liquido"));
 const body=$("pay-body");body.replaceChildren();
 if(!result.length)emptyRow(body,13,lista.length?"Nenhum lançamento corresponde aos filtros.":"Sem lançamentos para este mês.");
 const v=x=>M.formatarValor(x);
 result.slice(0,MAX_LINHAS).forEach(l=>{const tr=document.createElement("tr");cell(tr,l.codigo);cell(tr,l.nome,{strong:true});cell(tr,l.empresa);cell(tr,l.departamento);cell(tr,l.regime);cell(tr,v(l.dias).replace(/,00$/,""),{num:true});cell(tr,v(l.base),{num:true});cell(tr,v(l.acrescimos),{num:true});cell(tr,v(l.descontos),{num:true});cell(tr,v(l.inssColab),{num:true});cell(tr,v(l.irps),{num:true});cell(tr,v(l.liquido),{num:true,strong:true});stateCell(tr,l.estado);tr.lastChild.firstChild.classList.toggle("off",!/finaliz/i.test(l.estado));const pk=per.ano+"-"+per.mes+"|"+M.chaveRH(M.campo(l.linha,"Empresa"),M.campo(l.linha,"Codigo")),pp=pendentes.get(pk);if(pp){const cls=pp.erro?"erro":pp.fim?"feito":"gravando";const b=tr.lastChild.firstChild;b.dataset.pendente=pk;b.textContent=rotuloPendente(pp);b.classList.add(cls);const n=document.createElement("small");n.className="pend state "+cls;n.dataset.pendente=pk;n.textContent=rotuloPendente(pp);tr.children[1].append(n);tr.classList.add("linha-"+cls);if(pp.erro){b.title=pp.erro;n.title=pp.erro}}tr.classList.add("clickable");tr.tabIndex=0;tr.title="Abrir lançamento";const abre=()=>gravacaoActiva(pk)?showToast("Este lançamento ainda está a ser gravado ("+segs(Date.now()-pp.inicio)+" s)."):lanc.editar(l.linha);tr.addEventListener("click",abre);tr.addEventListener("keydown",e=>{if(e.key==="Enter")abre()});body.append(tr)});
 const foot=$("pay-foot");foot.replaceChildren();
 if(result.length){const tr=document.createElement("tr");cell(tr,"Totais",{strong:true});["","","","",""].forEach(()=>cell(tr,""));["base","acrescimos","descontos","inssColab","irps","liquido"].forEach(k=>cell(tr,fmt(M.totais(result,k)),{num:true,strong:true}));cell(tr,"");foot.append(tr)}
 $("pay-count").textContent=result.length+" de "+lista.length+" lançamento(s)"+(result.length>MAX_LINHAS?" · a mostrar "+MAX_LINHAS:"")+" · INSS empresa: "+fmt(M.totais(result,"inssEmpresa"))+" MT";
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
 if(dados.lancamentos){const p=dados.lancamentos.periodo||M.periodos(linhas("lancamentos","tbl_lançamento"))[0];if(p){const l=M.lancamentos(linhas("lancamentos","tbl_lançamento"),p.ano,p.mes);$("m-mes").textContent=M.NOME_MES[p.mes-1].slice(0,3)+" "+p.ano;if(!l.length){$("m-mes-sub").textContent="Sem lançamentos nos últimos meses";}else $("m-mes-sub").textContent=l.length+" lançamentos · "+(M.contagem(l,x=>x.estado)[0]?.[0]||"")}}
 if(Object.keys(dados).length)$("dash-badge").textContent="Dados do Excel";
}
function render(){ficha.actualizar();renderEmployees();renderContratos();renderAdendas();renderSalarios();renderFerias();renderDashboard()}

$("employee-search").addEventListener("input",renderEmployees);
["employee-filter","employee-department"].forEach(id=>$(id).addEventListener("change",renderEmployees));
$("employee-company").addEventListener("change",()=>{$("employee-department").value="";renderEmployees()});
$("employee-clear").addEventListener("click",()=>{["employee-search","employee-company","employee-department","employee-filter"].forEach(id=>$(id).value="");renderEmployees()});
$("pay-company").addEventListener("change",()=>{$("pay-department").value="";$("pay-employee").value="";renderSalarios()});
$("pay-department").addEventListener("change",()=>{$("pay-employee").value="";renderSalarios()});
$("pay-employee").addEventListener("change",renderSalarios);
$("addenda-search").addEventListener("input",renderAdendas);
$("addenda-filter").addEventListener("change",renderAdendas);
$("pay-search").addEventListener("input",renderSalarios);
document.querySelectorAll("[data-reload]").forEach(b=>b.addEventListener("click",()=>carregar(b.dataset.reload,true)));
const month=$("pay-month"),year=$("pay-year");
M.NOME_MES.forEach((m,i)=>{const op=document.createElement("option");op.value=String(i+1);op.textContent=m;month.appendChild(op)});
const now=new Date();month.value=String(now.getMonth()+1);
for(let y=now.getFullYear()-2;y<=now.getFullYear()+1;y++){const op=document.createElement("option");op.value=String(y);op.textContent=String(y);year.appendChild(op)}year.value=String(now.getFullYear());
[month,year].forEach(s=>s.addEventListener("change",()=>{if(ligado)carregarLancamentos({ano:Number(year.value),mes:Number(month.value)}).catch(()=>{});else renderSalarios()}));
$("load-month").addEventListener("click",()=>{if(!ligado){abrirAcesso();return}carregar("lancamentos",true)});

// ---------- sessão ----------
// Dois caminhos: código de acesso pessoal (fluxo RH_API_Acesso, independente do Entra)
// ou conta Microsoft (quando o clientId estiver configurado).
function transporte(){const s=sessaoAcesso();return s?{credenciais:s,url:urlGuardada()}:{getToken}}
function estadoLigacao(tipo,titulo,texto){
 ["source-dot","status-pill","auth-notice"].forEach(id=>{$(id).classList.toggle("live",tipo==="live");$(id).classList.toggle("error",tipo==="error")});
 $("notice-title").textContent=titulo;$("notice-text").textContent=texto;
}
function mostrarUtilizador(u){
 utilizadorActual=u;
 $("avatar").textContent=(u.nome||u.email||"RH").split(/\s+/).map(p=>p[0]).join("").slice(0,2).toUpperCase();
 $("avatar").title=(u.nome||u.email)+" · "+u.perfil;
 $("source-sub").textContent=u.perfil+(u.propor?"":" · só consulta");
}
function botoes(){const conta=Boolean(currentAccount()),codigo=Boolean(sessaoAcesso());
 $("code-signin").hidden=conta||codigo;$("signin").hidden=conta||codigo||!authConfigured();$("signout").hidden=!(conta||codigo);
 document.querySelectorAll("[data-reload]").forEach(b=>b.hidden=!(conta||codigo))}
function ligar(quem,via){
 ligado=true;botoes();
 $("status-pill").textContent="● Ligado ao Excel";$("source-title").textContent="Ligado ao SharePoint";$("source-sub").textContent=quem;
 estadoLigacao("live","Sessão iniciada: "+quem,"Os dados vêm do ficheiro Gestão de Colaboradores.xlsx através do Power Automate ("+via+"). Os administradores RH podem criar e editar colaboradores; os restantes módulos são só de consulta.");
 (precisa[paginaActual]||[]).forEach(m=>carregar(m));
 if(paginaActual==="painel")painel.abrir();
 if(paginaActual==="relatorios")relatorios.abrir();
}
function desligar(){
 painel.limpar();relatorios.limpar();
 ligado=false;Object.keys(dados).forEach(k=>delete dados[k]);botoes();
 $("status-pill").textContent="● Ambiente de demonstração";$("source-title").textContent="Modo demonstração";$("source-sub").textContent="Sem ligação ao SharePoint";$("avatar").textContent="RH";$("avatar").title="";
 ["contratos","adendas","lancamentos","ferias"].forEach(p=>showLoaded(p,false));
 ["employee-company","employee-department","pay-company","pay-department","pay-employee"].forEach(id=>$(id).value="");meses.clear();mesesCompletos=false;
 if(ficha.aberta())document.getElementById("ficha").close();
 render();
}
// Diálogo de acesso por código
const dlg=$("access-dialog");
function mostrarCampoUrl(mostrar){$("access-url-field").hidden=!mostrar;$("access-change-url").hidden=mostrar}
function abrirAcesso(erro=""){
 mostrarCampoUrl(!urlGuardada());$("access-error").textContent=erro;$("access-code").value="";$("access-key").textContent="";
 if(typeof dlg.showModal==="function")dlg.showModal();else dlg.setAttribute("open","");
 ($("access-email").value?$("access-code"):$("access-email")).focus();
}
function fecharAcesso(){if(typeof dlg.close==="function")dlg.close();else dlg.removeAttribute("open")}
$("code-signin").addEventListener("click",()=>abrirAcesso());
$("access-cancel").addEventListener("click",fecharAcesso);
$("access-change-url").addEventListener("click",()=>{mostrarCampoUrl(true);$("access-url").focus()});
$("access-form").addEventListener("submit",async ev=>{
 ev.preventDefault();const btn=$("access-submit");btn.disabled=true;$("access-error").textContent="";
 try{
  const s=await entrarComCodigo({email:$("access-email").value,codigo:$("access-code").value,url:$("access-url").hidden?"":$("access-url").value});
  $("access-code").value="";$("access-url").value="";
  fecharAcesso();ligar(s.email,"código de acesso");
 }catch(e){$("access-error").textContent=e.message}
 finally{btn.disabled=false}
});
$("access-make-key").addEventListener("click",async()=>{
 const email=$("access-email").value,codigo=$("access-code").value;
 if(!validarEmail(email)||!validarCodigo(codigo)){$("access-error").textContent="Para gerar a chave, indique o e-mail e um código com pelo menos "+CODIGO_MINIMO+" caracteres.";return}
 $("access-error").textContent="";$("access-key").textContent=await chaveAcesso(email,codigo);
});
$("signin").addEventListener("click",()=>signIn().catch(e=>estadoLigacao("error","Sessão não iniciada",e.message)));
$("signout").addEventListener("click",()=>{
 if(sessaoAcesso()){sairAcesso();desligar();estadoLigacao("","Sessão terminada","Os números exibidos são fictícios. Entre com o seu código de acesso para ver os dados do Excel.");return}
 signOut().catch(e=>estadoLigacao("error","Erro ao terminar sessão",e.message));
});

// ---------- ficha do colaborador (gravação via RH_API_Actualizar) ----------
function podeGravar(){return Boolean(sessaoAcesso()&&utilizadorActual?.propor)}
const ficha=criarFicha({
 funcionarios:()=>linhas("funcionarios","tbl_funcionarios"),
 adendas:()=>dados.contratos?linhas("contratos","tbl_adendas"):null,
 parametros:()=>dados.contratos?linhas("contratos","tbl_parametros_pagamento"):null,
 carregarContratos:()=>carregar("contratos"),
 podeGravar,
 utilizador:()=>sessaoAcesso()?.email||"",
 urlGravacao,guardarUrlGravacao,
 gravar:(modo,pedido)=>gravarColaborador(modo,pedido,{credenciais:sessaoAcesso(),url:urlGravacao()}),
 recarregar:()=>carregar("funcionarios",true),
 aviso:showToast
});
$("new-employee").addEventListener("click",()=>{
 if(!ligado){abrirAcesso("Entre com o seu código de acesso para criar colaboradores.");return}
 if(!podeGravar()){showToast("O seu perfil ("+(utilizadorActual?.perfil||"desconhecido")+") só permite consultar. Gravar exige Superadministrador ou Administrador RH.");return}
 carregar("funcionarios").then(()=>ficha.novo()).catch(()=>{});
});

// ---------- lançamento salarial: folha, comparativo e Novo mês ----------
const lanc=iniciarLancamentos({
 periodo:()=>dados.lancamentos?.periodo||{ano:Number(year.value),mes:Number(month.value)},
 periodoMaisRecente:()=>dados.lancamentos?.periodo||null,
 lancamentosMes:()=>{const p=dados.lancamentos?.periodo;return p?doPeriodo(linhas("lancamentos","tbl_lançamento"),p.ano,p.mes):[]},
 filtros:()=>({empresa:$("pay-company").value,departamento:$("pay-department").value,codigo:$("pay-employee").value}),
 carregarMes,esquecerMes:(a,m)=>{meses.delete(a+"-"+m);mesesCompletos=false},
 funcionarios:()=>carregar("funcionarios").then(()=>linhas("funcionarios","tbl_funcionarios")),
 contratos:()=>carregar("contratos").then(()=>({adendas:unirAdendas(linhas("contratos","tbl_adendas"),linhas("contratos","tbl_adendas_colaborador"),linhas("contratos","tbl_adendas_detalhes")),parametros:linhas("contratos","tbl_parametros_pagamento")})),
 podeGravar,utilizador:()=>sessaoAcesso()?.email||"",urlGravacao,guardarUrlGravacao,
 gravarMes:pedido=>gravarNovoMes(pedido,{credenciais:sessaoAcesso(),url:urlGravacao()}),
 gravarLancamento:pedido=>gravarLancamento(pedido,{credenciais:sessaoAcesso(),url:urlGravacao()}),
 gravarEmFundo,
 finalizarMes:pedido=>finalizarMes(pedido,{credenciais:sessaoAcesso(),url:urlGravacao()}),
 recarregarMes:(ano,mes)=>{meses.delete(ano+"-"+mes);mesesCompletos=false;carregarLancamentos({ano,mes}).catch(()=>{})},
 mostrarMes:(ano,mes)=>{year.value=String(ano);month.value=String(mes);carregarLancamentos({ano,mes}).catch(()=>{})},
 aviso:showToast
});
// ---------- Dashboard RH (Streamlit no navegador) ----------
const painel=iniciarPainel({
 ligado:()=>ligado,
 funcionarios:forcar=>carregar("funcionarios",forcar).then(()=>linhas("funcionarios","tbl_funcionarios")),
 carregarMes:(ano,mes,forcar)=>carregarMes(ano,mes,forcar),
 periodoMaisRecente:()=>carregar("lancamentos").then(d=>d?.periodo||null).catch(()=>null)
});
// ---------- Relatórios formais (PDF e Word) ----------
const relatorios=iniciarRelatorios({
 ligado:()=>ligado,
 funcionarios:()=>carregar("funcionarios").then(()=>linhas("funcionarios","tbl_funcionarios")),
 carregarMes:(ano,mes)=>carregarMes(ano,mes),
 periodoMaisRecente:()=>carregar("lancamentos").then(d=>d?.periodo||null).catch(()=>null),
 utilizador:()=>utilizadorActual?.nome||sessaoAcesso()?.email||"",
 aviso:showToast
});
$("btn-preview").addEventListener("click",()=>lanc.preVisualizar());
$("btn-pdf").addEventListener("click",e=>{const b=e.currentTarget;b.disabled=true;lanc.exportarFolha("pdf").catch(x=>showToast(x.message)).finally(()=>b.disabled=false)});
$("btn-xlsx").addEventListener("click",e=>{const b=e.currentTarget;b.disabled=true;lanc.exportarFolha("xlsx").catch(x=>showToast(x.message)).finally(()=>b.disabled=false)});
$("btn-comparativo").addEventListener("click",()=>lanc.comparativo());
$("btn-finalizar").addEventListener("click",()=>{if(!ligado){abrirAcesso("Entre com o seu código de acesso para finalizar o mês.");return}lanc.finalizar()});
$("novo-mes").addEventListener("click",()=>{if(!ligado){abrirAcesso("Entre com o seu código de acesso para criar um novo mês.");return}lanc.novoMes()});

render();
(async()=>{
 const s=sessaoAcesso();
 if(s&&urlGuardada()){ligar(s.email,"código de acesso");return}
 if(location.hash==="#entrar"){history.replaceState(null,"",location.pathname);abrirAcesso()}
 botoes();
 if(!authConfigured())return;
 try{
  await initAuth();botoes();
  const conta=currentAccount();
  if(conta)ligar(conta.username,"conta Microsoft");
 }catch(e){estadoLigacao("error","Erro de autenticação",e.message)}
})();
