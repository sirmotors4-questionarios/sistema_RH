import {evaluateFormula} from "./rh-engine.js";
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
 const base=Number($("#sim-base").value),liquido=Number($("#sim-liquido").value),diasMes=get("diasMes"),horasMes=get("horasMes"),horasDia=get("horasDia");
 if([base,liquido,diasMes,horasMes,horasDia].some(v=>!Number.isFinite(v))||base<0||liquido<0||diasMes<=0||horasMes<=0||horasDia<=0)throw Error("Informe valores numéricos válidos e divisores positivos.");
 const diario=base/diasMes,vars={base,liquido,diasMes,horasMes,horasDia,diario};
 const expression=library.value==="custom"?custom.value.trim():library.value;
 if(!expression)throw Error("Introduza a fórmula personalizada.");
 const result=evaluate(expression,vars);
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
