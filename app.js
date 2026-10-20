"use strict";
// Atenção: apenas dados fictícios. Não integrar segredos/URLs Power Automate no browser.
const mockEmployees = [
 {codigo:"RH-001",nome:"Colaborador Exemplo A",empresa:"Empresa A",departamento:"Administração",regime:"Mensal",estado:"Activo"},
 {codigo:"RH-002",nome:"Colaborador Exemplo B",empresa:"Empresa B",departamento:"Operações",regime:"Diário",estado:"Activo"},
 {codigo:"RH-003",nome:"Colaborador Exemplo C",empresa:"Empresa A",departamento:"Manutenção",regime:"Hora",estado:"Inactivo"},
 {codigo:"RH-004",nome:"Colaborador Exemplo D",empresa:"Empresa C",departamento:"Financeiro",regime:"Mensal",estado:"Activo"}
];
const labels={dashboard:"Visão geral",colaboradores:"Colaboradores",contratos:"Contratos",adendas:"Adendas",salarios:"Lançamento salarial",ferias:"Férias",relatorios:"Relatórios",configuracoes:"Configurações"};
function go(page){
 if(page==="configuracoes"){window.location.href="configuracoes.html";return;}

 if(!Object.hasOwn(labels,page))return;
 document.querySelectorAll(".page").forEach(el=>el.classList.toggle("active",el.id===page));
 document.querySelectorAll("[data-page]").forEach(el=>el.classList.toggle("active",el.dataset.page===page));
 document.getElementById("crumb").textContent=labels[page];
 document.getElementById("sidebar").classList.remove("open");
 window.scrollTo({top:0,behavior:"instant"});
}
document.querySelectorAll("[data-page]").forEach(el=>el.addEventListener("click",()=>go(el.dataset.page)));
document.querySelectorAll("[data-go]").forEach(el=>el.addEventListener("click",()=>go(el.dataset.go)));
document.getElementById("menu-toggle").addEventListener("click",()=>document.getElementById("sidebar").classList.toggle("open"));
function showToast(message){const t=document.getElementById("toast");t.textContent=message;t.classList.add("show");clearTimeout(showToast.timer);showToast.timer=setTimeout(()=>t.classList.remove("show"),4500)}
document.querySelectorAll("[data-action]").forEach(el=>el.addEventListener("click",()=>showToast("Funcionalidade ainda não activada. A ligação segura ao Excel/SharePoint está pendente.")));
function appendCell(row,value,strong=false){const td=document.createElement("td");const node=strong?document.createElement("strong"):document.createElement("span");node.textContent=value;td.appendChild(node);row.appendChild(td)}
function renderEmployees(){
 const term=document.getElementById("employee-search").value.trim().toLocaleLowerCase("pt");
 const state=document.getElementById("employee-filter").value;
 const result=mockEmployees.filter(e=>(!state||e.estado===state)&&[e.codigo,e.nome,e.empresa,e.departamento].some(x=>x.toLocaleLowerCase("pt").includes(term)));
 const tbody=document.getElementById("employees-body");tbody.replaceChildren();
 if(!result.length){const tr=document.createElement("tr"),td=document.createElement("td");td.colSpan=7;td.textContent="Nenhum colaborador corresponde aos filtros.";tr.appendChild(td);tbody.appendChild(tr)}
 result.forEach(e=>{const tr=document.createElement("tr");appendCell(tr,e.codigo);appendCell(tr,e.nome,true);appendCell(tr,e.empresa);appendCell(tr,e.departamento);appendCell(tr,e.regime);
 const stateTd=document.createElement("td"),badge=document.createElement("span");badge.className="state"+(e.estado==="Activo"?"":" off");badge.textContent=e.estado;stateTd.appendChild(badge);tr.appendChild(stateTd);
 const td=document.createElement("td"),btn=document.createElement("button");btn.className="inline-btn";btn.textContent="Consultar";btn.addEventListener("click",()=>showToast("A consulta de detalhes será activada após a ligação autenticada."));td.appendChild(btn);tr.appendChild(td);tbody.appendChild(tr)});
 document.getElementById("employee-count").textContent=result.length+" registo(s) demonstrativo(s)";
}
document.getElementById("employee-search").addEventListener("input",renderEmployees);
document.getElementById("employee-filter").addEventListener("change",renderEmployees);
const month=document.getElementById("pay-month"),year=document.getElementById("pay-year");
["Janeiro","Fevereiro","Março","Abril","Maio","Junho","Julho","Agosto","Setembro","Outubro","Novembro","Dezembro"].forEach((m,i)=>{const op=document.createElement("option");op.value=String(i+1);op.textContent=m;month.appendChild(op)});
const now=new Date();month.value=String(now.getMonth()+1);
for(let y=now.getFullYear()-2;y<=now.getFullYear()+1;y++){const op=document.createElement("option");op.value=String(y);op.textContent=String(y);year.appendChild(op)}year.value=String(now.getFullYear());
renderEmployees();
