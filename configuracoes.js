import {evaluateFormula} from "./rh-engine.js";
import {readTables, resolveTables, simulateRecord} from "./rh-data.js";
import {authConfigured, initAuth, currentAccount, signIn, signOut, getToken} from "./rh-auth.js";
import {loadConfiguracoes, loadModulo, escaloesLer, escaloesGravar, escaloesAplicarMes} from "./rh-api.js";
import * as ESC from "./escaloes.js";
import {NOME_MES} from "./rh-modulos.js";
import {contarTiposContrato} from "./tipos-contrato.js";
import {sessaoAcesso, urlGuardada, sairAcesso, urlIndemnizacoes, guardarUrlIndemnizacoes} from "./rh-acesso.js";
"use strict";
const $=sel=>document.querySelector(sel),$$=sel=>[...document.querySelectorAll(sel)];
const nav=$$(".config-nav");
function section(name){nav.forEach(b=>b.classList.toggle("active",b.dataset.section===name));$$(".section-view").forEach(v=>v.classList.toggle("active",v.dataset.view===name));}
nav.forEach(b=>b.addEventListener("click",()=>section(b.dataset.section)));
// Ligação directa a uma secção (ex.: configuracoes.html#contratos, a partir de Colaboradores).
if(location.hash&&nav.some(b=>b.dataset.section===location.hash.slice(1)))section(location.hash.slice(1));
$("#menu-toggle").addEventListener("click",()=>$("#sidebar").classList.toggle("open"));
function toast(m){const e=$("#toast");e.textContent=m;e.classList.add("show");clearTimeout(toast.timer);toast.timer=setTimeout(()=>e.classList.remove("show"),4500)}
const level=$("#scope-level");
function scope(){let n=["geral","empresa","categoria","colaborador"].indexOf(level.value);["#company","#category","#employee"].forEach((s,i)=>{$(s).disabled=n<i+1; if(n<i+1)$(s).value=""});}
level.addEventListener("change",scope);scope();
let tables=null, sourceError="", sourceLabel="", utilizador=null;
let dirty=false;function markDirty(){dirty=true;$("#dirty").textContent="Alterações por validar";$("#dirty").style.background="#fff1dc";$("#dirty").style.color="#926320";$("#sim-result").dataset.validated="false"}
$$(".config-work input:not([data-livre]),.config-work select:not([data-livre])").forEach(el=>el.addEventListener("input",markDirty));
const scales=$("#scales-body");
function addScale(){let tr=document.createElement("tr");for(let i=0;i<8;i++){let td=document.createElement("td"),input=document.createElement("input");input.type="number";input.step="0.01";input.min="0";input.setAttribute("aria-label",["Limite inferior","Limite superior","Dependentes 0","Dependentes 1","Dependentes 2","Dependentes 3","Dependentes 4 ou mais","Coeficiente adicional"][i]);input.addEventListener("input",markDirty);td.append(input);tr.append(td)}scales.append(tr);markDirty()}
$("#add-scale").addEventListener("click",addScale);
addScale();dirty=false;$("#dirty").textContent="Sem alterações";
const library=$("#formula-library"),custom=$("#formula-custom");
library.addEventListener("change",()=>{custom.disabled=library.value!=="custom";if(custom.disabled)custom.value="";});
function get(name){return Number($('[data-param="'+name+'"]').value)}
function evaluate(expression,vars){return evaluateFormula(expression,vars);}
function simulate(){const target=$("#sim-result");try{
 if(sourceError)throw Error(sourceError);
 if(tables){
  const raw=$("#sim-variables").value.trim();const extra=raw?JSON.parse(raw):{};
  if(!extra||typeof extra!=="object"||Array.isArray(extra))throw Error("As variáveis devem ser um objecto JSON.");
  if(!$("#sim-base").value.trim()||!$("#sim-liquido").value.trim())throw Error("Preencha os valores da simulação.");
  const base=Number($("#sim-base").value),liquido=Number($("#sim-liquido").value);
  if(!Number.isFinite(base)||!Number.isFinite(liquido)||base<0||liquido<0)throw Error("Valores de simulação inválidos.");
  const output=simulateRecord(tables,library.value,context(),$("#reference-date").value,{...extra,BASE:base,LIQUIDO:liquido});
  target.className="simulation success";
  target.textContent="Resultado da simulação: "+String(output.result)+" · Fórmula: "+output.formula.Codigo_Formula+" · Versão: "+output.formula.Versao+". Não confirma equivalência com o VBA nem constitui aprovação.";
  target.dataset.validated="true";return;
 }

 const base=Number($("#sim-base").value),liquido=Number($("#sim-liquido").value),diasMes=get("diasMes"),horasMes=get("horasMes"),horasDia=get("horasDia");
 if([base,liquido,diasMes,horasMes,horasDia].some(v=>!Number.isFinite(v))||base<0||liquido<0||diasMes<=0||horasMes<=0||horasDia<=0)throw Error("Informe valores numéricos válidos e divisores positivos.");
 const diario=base/diasMes,vars={base,liquido,diasMes,horasMes,horasDia,diario};
 const expression=library.value==="custom"?custom.value.trim():library.value;
 if(!expression)throw Error("Introduza a fórmula personalizada.");
 const result=evaluate(expression,vars);
 if(typeof result!=="number"||!Number.isFinite(result))throw Error("A fórmula salarial deve devolver um número finito.");
 const fmt=x=>new Intl.NumberFormat("pt-MZ",{minimumFractionDigits:2,maximumFractionDigits:2}).format(x);
 target.className="simulation success";
 target.textContent="Validação sintáctica concluída. Resultado demonstrativo: "+fmt(result)+" MT · Fórmula: "+expression+". Os cálculos fiscais oficiais e a equivalência com o VBA ainda não foram validados.";
 target.dataset.validated="true";toast("Simulação demonstrativa concluída.");
 }catch(e){target.className="simulation error";target.textContent="Não foi possível validar: "+e.message;target.dataset.validated="false";}}
$("#simulate").addEventListener("click",simulate);
function reset(){location.reload()}$("#reset").addEventListener("click",reset);
$("#propose").addEventListener("click",()=>{
 if(utilizador&&!utilizador.propor){toast("O seu perfil ("+utilizador.perfil+") permite consultar, mas não propor alterações.");return}
 const since=$("#vigencia").value,reason=$("#justificacao").value.trim();
 if(!since){section("geral");toast("Indique a data de vigência.");return}
 if(!reason){section("geral");toast("Introduza uma justificação.");return}
 if(level.value==="empresa"&&!$("#company").value.trim()||["categoria","colaborador"].includes(level.value)&&(!$("#company").value.trim()||!$("#category").value.trim())||level.value==="colaborador"&&!$("#employee").value.trim()){toast("Complete a identificação do âmbito.");return}
 if($("#sim-result").dataset.validated!=="true"){section("motor");toast("É obrigatório executar a simulação antes de propor alterações.");return}
 toast("Proposta validada apenas na interface. Não foi gravada nem enviada para aprovação: integração e autenticação pendentes.");
});

function context(){return {empresa:$("#company").value.trim(),categoria:$("#category").value.trim(),codigo:$("#employee").value.trim()};}
function rowCells(body, values){const tr=document.createElement("tr");values.forEach(value=>{const td=document.createElement("td");td.textContent=value??"";tr.append(td);});body.append(tr);}
function refreshSource(){
 if(!tables)return;
 markDirty();const body=$("#resolved-configs");body.replaceChildren();library.replaceChildren();
 try{
  const resolved=resolveTables(tables,context(),$("#reference-date").value);
  resolved.parameters.forEach(r=>rowCells(body,[r.Codigo_Parametro,r.Valor,r.Nivel_Aplicacao,r.Versao]));
  resolved.formulas.forEach(r=>{const option=document.createElement("option");option.value=r.Codigo_Formula;option.textContent=r.Codigo_Formula+" — "+r.Expressao;library.append(option);});
  custom.disabled=true;custom.value="";
  sourceError="";$("#simulate").disabled=!resolved.formulas.length;
  $("#source-status").textContent="Tabelas carregadas nesta sessão. "+resolved.parameters.length+" parâmetros resolvidos; "+resolved.formulas.length+" fórmulas disponíveis; "+resolved.unavailable.length+" não validadas. "+(sourceLabel||"Excel/Power Automate não ligado.");
 }catch(error){sourceError=error.message;$("#source-status").textContent="Erro nas tabelas: "+error.message;$("#simulate").disabled=true;}
}
$("#reference-date").value=new Date().toLocaleDateString('en-CA',{timeZone:'Africa/Maputo'});
["#scope-level","#company","#category","#employee","#reference-date"].forEach(selector=>$(selector).addEventListener("input",refreshSource));
$("#reference-date").addEventListener("change",()=>{$("#sim-result").dataset.validated="false";refreshSource();});
$("#sim-variables").addEventListener("input",markDirty);
function applyTables(parsed,label){
 try{
  tables=parsed;sourceLabel=label;
  $$("[data-param], #scales-body input, #add-scale").forEach(el=>el.disabled=true);
  refreshSource();
  const body=$("#audit-body");body.replaceChildren();
  tables.tbl_auditoria_RH.forEach(r=>rowCells(body,[r.Data_Operacao,r.Codigo_Parametro||r.ID_Registo,r.Valor_Anterior,r.Valor_Novo,r.Utilizador_Operacao,r.Estado_Novo||r.Tipo_Operacao]));
  if(!tables.tbl_auditoria_RH.length)rowCells(body,["Sem eventos de auditoria importados."]);
  $("#sim-result").textContent="Seleccione a fórmula e indique as variáveis necessárias. Não é executado processamento salarial completo.";
 }catch(error){sourceError=error.message;$("#source-status").textContent="Não foi possível carregar: "+error.message;$("#simulate").disabled=true;$("#sim-result").dataset.validated="false";}
}
$("#tables-file").addEventListener("change",async event=>{
 const file=event.target.files[0];if(!file)return;
 try{
  if(file.size>5*1024*1024)throw Error("O ficheiro excede o limite de 5 MB.");
  applyTables(readTables(JSON.parse(await file.text())),"Fonte: ficheiro JSON importado nesta sessão.");
 }catch(error){sourceError=error.message;$("#source-status").textContent="Não foi possível carregar: "+error.message;$("#simulate").disabled=true;}
});

level.addEventListener("change",refreshSource);

// Ligação autenticada: Microsoft Entra -> Power Automate (RH_GET_Configuracoes) -> Gestão de Colaboradores.xlsx
function authView(state,message){
 const bar=$("#auth-bar");bar.classList.toggle("ok",state==="ok");bar.classList.toggle("error",state==="error");
 $("#auth-status").textContent=message;
 const codigo=Boolean(sessaoAcesso()&&urlGuardada()),signed=Boolean(currentAccount())||codigo;
 $("#signin").hidden=signed||!authConfigured();$("#code-signin").hidden=signed;$("#signout").hidden=!signed;$("#load-remote").hidden=!signed;
}
// Acesso por código pessoal (sessão iniciada na página principal, no mesmo separador).
function deps(){const s=sessaoAcesso(),url=urlGuardada();return s&&url?{credenciais:s,url}:{getToken}}
async function loadRemote(){
 const button=$("#load-remote");button.disabled=true;authView("","A carregar as configurações do Excel através do Power Automate…");
 try{
  const result=await loadConfiguracoes(deps());
  utilizador=result.utilizador;
  applyTables(result.tables,"Fonte: Excel no SharePoint via Power Automate ("+new Date().toLocaleTimeString("pt-MZ")+").");
  $("#integration-badge").textContent="Motor ligado · Power Automate";
  $("#propose").disabled=!utilizador.propor;
  authView("ok",(utilizador.nome||utilizador.email)+" · Perfil: "+utilizador.perfil+(utilizador.propor?"":" (apenas consulta)")+". Configurações carregadas.");
  carregarTiposContrato().then(carregarEscaloes);
 }catch(error){
  utilizador=null;authView("error",error.message);
 }finally{button.disabled=false;}
}
$("#signin").addEventListener("click",()=>signIn().catch(e=>authView("error",e.message)));
$("#signout").addEventListener("click",()=>{if(sessaoAcesso()){sairAcesso();location.reload();return}signOut().catch(e=>authView("error",e.message))});
$("#load-remote").addEventListener("click",loadRemote);
(async()=>{
 const s=sessaoAcesso();
 if(s&&urlGuardada()){authView("","Sessão iniciada: "+s.email+".");await loadRemote();return}
 // Sem sessão, as Configurações não se mostram: vai para o ecrã de entrada da página principal.
 if(!authConfigured()){location.replace("index.html");return}
 try{
  await initAuth();
  const account=currentAccount();
  if(account){authView("","Sessão iniciada: "+account.username+".");await loadRemote();}
  else location.replace("index.html");
 }catch(error){authView("error",error.message);}
})();

// ---------- Tipos de contrato ----------
// Contagem a partir do cadastro (tbl_funcionarios.TipoContrato). «+ Novo contrato» acrescenta um tipo proposto,
// mantido só nesta sessão até existir gravação de configurações no Excel.
let funcionarios=null;const propostos=[];
async function carregarTiposContrato(){
 const body=$("#tc-body");body.replaceChildren();rowCells(body,["A ler o cadastro de colaboradores…"]);body.firstChild.firstChild.colSpan=7;
 try{funcionarios=(await loadModulo("funcionarios",deps())).tabelas.tbl_funcionarios;}
 catch(e){funcionarios=null;body.replaceChildren();rowCells(body,["Não foi possível ler o cadastro: "+e.message]);body.firstChild.firstChild.colSpan=7;return}
 renderTiposContrato();
}
function renderTiposContrato(){
 const body=$("#tc-body");body.replaceChildren();
 const r=contarTiposContrato(funcionarios||[],propostos);
 $("#tc-n-tipos").textContent=funcionarios?String(r.tipos.length):(propostos.length?String(propostos.length):"—");
 $("#tc-n-activos").textContent=funcionarios?String(r.activos):"—";
 $("#tc-n-sem").textContent=funcionarios?String(r.semTipo):"—";
 if(!r.tipos.length){rowCells(body,["Sem tipos de contrato no cadastro."]);body.firstChild.firstChild.colSpan=7;return}
 r.tipos.forEach(t=>{const tr=document.createElement("tr");
  [t.nome,t.duracao||"—",t.regime||"—",t.activos,t.inactivos,t.total].forEach((v,i)=>{const td=document.createElement("td");if(i>=3)td.className="num";td.textContent=String(v);if(i===0)td.style.fontWeight="700";tr.append(td)});
  const td=document.createElement("td"),tag=document.createElement("span");tag.className="tc-tag"+(t.proposto?" proposto":"");tag.textContent=t.proposto?"Proposto · por gravar":"Cadastro";td.append(tag);tr.append(td);body.append(tr)});
}
const tcForm=$("#tc-form");
$("#novo-contrato").addEventListener("click",()=>{
 if(utilizador&&!utilizador.propor){toast("O seu perfil ("+utilizador.perfil+") permite consultar, mas não propor tipos de contrato.");return}
 tcForm.hidden=false;$("#tc-erro").textContent="";$("#tc-nome").focus();
});
$("#tc-cancelar").addEventListener("click",()=>{tcForm.reset();tcForm.hidden=true});
tcForm.addEventListener("submit",ev=>{
 ev.preventDefault();
 const nome=$("#tc-nome").value.trim().replace(/\s+/g," ");
 const existe=n=>contarTiposContrato(funcionarios||[],propostos).tipos.some(t=>t.nome.localeCompare(n,"pt",{sensitivity:"base"})===0);
 if(!nome){$("#tc-erro").textContent="Indique a designação do tipo de contrato.";return}
 if(existe(nome)){$("#tc-erro").textContent="Já existe um tipo de contrato «"+nome+"».";return}
 const dias=$("#tc-experimental").value;
 if(dias!==""&&!(Number.isInteger(Number(dias))&&Number(dias)>=0&&Number(dias)<=180)){$("#tc-erro").textContent="O período experimental deve ser um número inteiro de 0 a 180 dias.";return}
 propostos.push({nome,duracao:$("#tc-duracao").value,regime:$("#tc-regime").value,experimental:dias});
 tcForm.reset();tcForm.hidden=true;renderTiposContrato();
 toast("Tipo «"+nome+"» adicionado à lista como proposto. Ainda não foi gravado no Excel.");
});

// ---------- Escalões salariais (tbl_escaloes_lancamento, via RH_API_Indemnizacoes) ----------
let escDefs=new Map(),escFaixas=[],escCarregado=false;
const escTransp=()=>({credenciais:sessaoAcesso(),url:urlIndemnizacoes()});
const escFmt=v=>v===null||v===undefined||v===""?"":new Intl.NumberFormat("pt-MZ",{minimumFractionDigits:2,maximumFractionDigits:2}).format(v);
const escEstado=m=>{$("#esc-estado").textContent=m};
(function iniciarEscaloes(){
 const crit=$("#esc-criterio");ESC.CRITERIOS.forEach(c=>{const o=document.createElement("option");o.value=c.coluna;o.textContent=c.rotulo;crit.append(o)});
 const mes=$("#esc-mes");NOME_MES.forEach((n,i)=>{const o=document.createElement("option");o.value=n;o.textContent=n;mes.append(o)});
 const hoje=new Date();$("#esc-ano").value=String(hoje.getFullYear());mes.value=NOME_MES[hoje.getMonth()];
 $("#esc-empresa").addEventListener("change",escMostrarEmpresa);
 crit.addEventListener("change",escCalculos);
 $("#esc-add").addEventListener("click",()=>{escLerTabela();const ult=escFaixas.length?Number(escFaixas[escFaixas.length-1].inicio)||0:0;escFaixas.push({escalao:"",inicio:escFaixas.length?ult+1000:0});escRender(true)});
 $("#esc-repor").addEventListener("click",escMostrarEmpresa);
 $("#esc-gravar").addEventListener("click",escGravar);
 $("#esc-aplicar").addEventListener("click",escAplicar);
 $("#esc-url-guardar").addEventListener("click",()=>{try{guardarUrlIndemnizacoes($("#esc-url").value);$("#esc-url").value="";carregarEscaloes()}catch(e){$("#esc-erro").textContent=e.message}});
 {const o=document.createElement("option");o.value=ESC.TODAS;o.textContent="Todas as empresas";$("#esc-empresa").append(o)}
 escRender();
})();
function escEmpresas(){
 const nomes=new Map();(funcionarios||[]).forEach(f=>{const e=String(f.Empresa??"").trim();if(e)nomes.set(e.toUpperCase(),e)});
 escDefs.forEach(d=>{if(d.empresa.toUpperCase()!==ESC.TODAS.toUpperCase())nomes.set(d.empresa.toUpperCase(),d.empresa)});
 return [...nomes.values()].sort((a,b)=>a.localeCompare(b,"pt"));
}
async function carregarEscaloes(){
 const sel=$("#esc-empresa"),actual=sel.value;
 $("#esc-ligacao").hidden=Boolean(urlIndemnizacoes());
 if(!sessaoAcesso()){escEstado("Inicie sessão com o código de acesso (página principal) para ler e gravar os escalões.");}
 else if(!urlIndemnizacoes()){escEstado("Falta o endereço do fluxo RH_API_Indemnizacoes.");}
 else{
  escEstado("A ler os escalões do Excel…");
  try{escDefs=ESC.agrupar(await escaloesLer(escTransp()));escCarregado=true;escEstado(escDefs.size?"Escalões lidos da tbl_escaloes_lancamento ("+new Date().toLocaleTimeString("pt-MZ")+").":"Ainda não há escalões gravados. Defina-os abaixo; a tabela é criada na primeira gravação.");}
  catch(e){escEstado("Não foi possível ler os escalões: "+e.message);}
 }
 sel.replaceChildren();
 const op=(v,t)=>{const o=document.createElement("option");o.value=v;o.textContent=t+(escDefs.has(v.toUpperCase())?" · com escalões":"");sel.append(o)};
 op(ESC.TODAS,"Todas as empresas");escEmpresas().forEach(e=>op(e,e));
 if([...sel.options].some(o=>o.value===actual))sel.value=actual;
 escMostrarEmpresa();
}
function escMostrarEmpresa(){
 const d=escDefs.get($("#esc-empresa").value.toUpperCase());
 $("#esc-criterio").value=d?.criterio||ESC.CRITERIO_PADRAO;
 escFaixas=d?d.faixas.map(f=>({...f})):[];
 $("#esc-erro").textContent="";escRender();
}
function escLerTabela(){
 escFaixas=[...$("#esc-body").querySelectorAll("tr[data-i]")].map(tr=>({escalao:tr.querySelector("[data-k=n]").value,inicio:tr.querySelector("[data-k=v]").value}));
}
// Simulação: a própria empresa; em «Todas», as empresas que não têm escalões próprios.
function escActivos(){
 const emp=$("#esc-empresa").value.toUpperCase(),todas=emp===ESC.TODAS.toUpperCase();
 return (funcionarios||[]).filter(f=>{const e=String(f.Empresa??"").trim().toUpperCase();return /^activo$/i.test(String(f.Estado??"").trim())&&(todas?!escDefs.has(e):e===emp)});
}
// Desenha as linhas (só ao acrescentar, remover ou mudar de empresa); os cálculos actualizam-se no lugar,
// para não perder o campo em que o utilizador está a escrever.
function escRender(focar=false){
 const body=$("#esc-body");body.replaceChildren();
 if(!escFaixas.length){rowCells(body,["Sem escalões para esta empresa. Carregue em «+ Acrescentar escalão»."]);body.firstChild.firstChild.colSpan=5;$("#esc-sim").textContent="";return}
 escFaixas.forEach((f,i)=>{
  const tr=document.createElement("tr");tr.dataset.i=String(i);
  const nome=document.createElement("input");nome.dataset.k="n";nome.value=f.escalao;nome.placeholder="Ex.: Escalão 1";nome.setAttribute("aria-label","Nome do escalão "+(i+1));
  const val=document.createElement("input");val.dataset.k="v";val.type="number";val.step="0.01";val.min="0";val.value=f.inicio===null?"":String(f.inicio);val.setAttribute("aria-label","Valor inicial do escalão "+(i+1));
  [nome,val].forEach(x=>x.addEventListener("input",()=>{escLerTabela();escCalculos()}));
  [nome,val,document.createElement("span"),document.createElement("span")].forEach((c,j)=>{const td=document.createElement("td");if(j>=1)td.className="num";if(j===2)td.classList.add("esc-ate");td.append(c);tr.append(td)});
  const td=document.createElement("td"),rm=document.createElement("button");rm.type="button";rm.className="esc-remover";rm.textContent="×";rm.title="Remover escalão";rm.setAttribute("aria-label","Remover escalão "+(i+1));
  rm.addEventListener("click",()=>{escLerTabela();escFaixas.splice(i,1);escRender()});td.append(rm);tr.append(td);body.append(tr);
 });
 escCalculos();
 if(focar){const ult=body.querySelector("tr:last-child [data-k=n]");ult&&ult.focus()}
}
function escCalculos(){
 const validas=escFaixas.map(f=>({escalao:String(f.escalao).trim(),inicio:ESC.numero(f.inicio)})).filter(f=>f.escalao&&f.inicio!==null);
 const activos=escActivos(),dist=ESC.distribuicao(validas,$("#esc-criterio").value,activos),conta=new Map(dist.porEscalao.map(x=>[x.escalao,x.total]));
 const ate=new Map(ESC.intervalos(validas).map(x=>[x.escalao,x.ate]));
 $("#esc-body").querySelectorAll("tr[data-i]").forEach(tr=>{
  const n=tr.querySelector("[data-k=n]").value.trim(),cel=tr.querySelectorAll("td");
  const a=ate.has(n)?ate.get(n):undefined;
  cel[2].firstChild.textContent=a===undefined?"":a===null?"sem limite":"< "+escFmt(a);
  cel[3].firstChild.textContent=funcionarios&&a!==undefined?String(conta.get(n)??0):"—";
 });
 $("#esc-sim").textContent=funcionarios?(activos.length+" activos na simulação"+(dist.abaixo?" · "+dist.abaixo+" abaixo do 1.º escalão":"")+(dist.semValor?" · "+dist.semValor+" sem valor no cadastro":"")):"";
}
function escPodeGravar(){
 if(!sessaoAcesso()||!urlIndemnizacoes()){$("#esc-erro").textContent="Inicie sessão e indique o endereço do fluxo RH_API_Indemnizacoes.";return false}
 if(utilizador&&!utilizador.propor){$("#esc-erro").textContent="O seu perfil ("+utilizador.perfil+") só permite consultar.";return false}
 return true;
}
async function escGravar(){
 escLerTabela();$("#esc-erro").textContent="";
 const pedido={empresa:$("#esc-empresa").value,criterio:$("#esc-criterio").value,faixas:escFaixas};
 const erros=ESC.validar(pedido);if(erros.length){$("#esc-erro").textContent=erros.join(" ");return}
 if(!escPodeGravar())return;
 const b=$("#esc-gravar");b.disabled=true;b.textContent="A gravar…";
 try{const r=await escaloesGravar(ESC.pedidoGravacao(pedido),escTransp());escDefs=ESC.agrupar(r.linhas);toast(r.gravados+" escalões gravados para «"+(pedido.empresa===ESC.TODAS?"Todas as empresas":pedido.empresa)+"». Os próximos lançamentos já os usam.");await carregarEscaloes()}
 catch(e){$("#esc-erro").textContent="Não foi gravado: "+e.message}
 finally{b.disabled=false;b.textContent="Gravar escalões no Excel"}
}
async function escAplicar(){
 $("#esc-erro").textContent="";
 if(!escDefs.size){$("#esc-erro").textContent="Grave primeiro os escalões.";return}
 if(!escPodeGravar())return;
 const ano=Number($("#esc-ano").value),mesNome=$("#esc-mes").value;
 if(!Number.isInteger(ano)||ano<2000){$("#esc-erro").textContent="Ano inválido.";return}
 const b=$("#esc-aplicar");b.disabled=true;b.textContent="A aplicar…";
 try{const r=await escaloesAplicarMes({ano,mesNome},escTransp());toast(mesNome+"/"+ano+": "+r.alteradas+" linhas actualizadas"+(r.semEscalao?" · "+r.semEscalao+" sem escalão (sem valor ou abaixo do 1.º)":"")+".")}
 catch(e){$("#esc-erro").textContent="Não foi aplicado: "+e.message}
 finally{b.disabled=false;b.textContent="Aplicar escalões ao mês"}
}
