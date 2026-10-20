import {evaluateFormula} from "./rh-engine.js";
import {readTables, resolveTables, simulateRecord} from "./rh-data.js";
"use strict";
const $=sel=>document.querySelector(sel),$$=sel=>[...document.querySelectorAll(sel)];
const nav=$$(".config-nav");
function section(name){nav.forEach(b=>b.classList.toggle("active",b.dataset.section===name));$$(".section-view").forEach(v=>v.classList.toggle("active",v.dataset.view===name));}
nav.forEach(b=>b.addEventListener("click",()=>section(b.dataset.section)));
$("#menu-toggle").addEventListener("click",()=>$("#sidebar").classList.toggle("open"));
function toast(m){const e=$("#toast");e.textContent=m;e.classList.add("show");clearTimeout(toast.timer);toast.timer=setTimeout(()=>e.classList.remove("show"),4500)}
const level=$("#scope-level");
function scope(){let n=["geral","empresa","categoria","colaborador"].indexOf(level.value);["#company","#category","#employee"].forEach((s,i)=>{$(s).disabled=n<i+1; if(n<i+1)$(s).value=""});}
level.addEventListener("change",scope);scope();
let tables=null, sourceError="";
let dirty=false;function markDirty(){dirty=true;$("#dirty").textContent="Alterações por validar";$("#dirty").style.background="#fff1dc";$("#dirty").style.color="#926320";$("#sim-result").dataset.validated="false"}
$$(".config-work input,.config-work select").forEach(el=>el.addEventListener("input",markDirty));
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
  $("#source-status").textContent="Tabelas carregadas nesta sessão. "+resolved.parameters.length+" parâmetros resolvidos; "+resolved.formulas.length+" fórmulas disponíveis; "+resolved.unavailable.length+" não validadas. Excel/Power Automate não ligado.";
 }catch(error){sourceError=error.message;$("#source-status").textContent="Erro nas tabelas: "+error.message;$("#simulate").disabled=true;}
}
$("#reference-date").value=new Date().toLocaleDateString('en-CA',{timeZone:'Africa/Maputo'});
["#scope-level","#company","#category","#employee","#reference-date"].forEach(selector=>$(selector).addEventListener("input",refreshSource));
$("#reference-date").addEventListener("change",()=>{$("#sim-result").dataset.validated="false";refreshSource();});
$("#sim-variables").addEventListener("input",markDirty);
$("#tables-file").addEventListener("change",async event=>{
 const file=event.target.files[0];if(!file)return;
 try{
  if(file.size>5*1024*1024)throw Error("O ficheiro excede o limite de 5 MB.");
  const parsed=readTables(JSON.parse(await file.text()));
  tables=parsed;
  $$("[data-param], #scales-body input, #add-scale").forEach(el=>el.disabled=true);
  refreshSource();
  const body=$("#audit-body");body.replaceChildren();
  tables.tbl_auditoria_RH.forEach(r=>rowCells(body,[r.Data_Operacao,r.Codigo_Parametro||r.ID_Registo,r.Valor_Anterior,r.Valor_Novo,r.Utilizador_Operacao,r.Estado_Novo||r.Tipo_Operacao]));
  if(!tables.tbl_auditoria_RH.length)rowCells(body,["Sem eventos de auditoria importados."]);
  $("#sim-result").textContent="Seleccione a fórmula e indique as variáveis necessárias. Não é executado processamento salarial completo.";
 }catch(error){sourceError=error.message;$("#source-status").textContent="Não foi possível carregar: "+error.message;$("#simulate").disabled=true;$("#sim-result").dataset.validated="false";}
});

level.addEventListener("change",refreshSource);
