/* Motor RH — sem eval e sem acesso a dados pessoais.
 * Este módulo interpreta apenas fórmulas matemáticas whitelist e resolve parâmetros aprovados.
 * Não certifica conformidade fiscal nem substitui conferência com o VBA. */
export class RHError extends Error { constructor(message,code="RH_VALIDATION"){ super(message);this.name="RHError";this.code=code;} }
const norm=x=>String(x??"").trim();
const upper=x=>norm(x).toLocaleUpperCase("pt");
const num=x=>{if(typeof x==="number")return x;if(typeof x!=="string"||!norm(x))return NaN;return Number(x.replace(/\s/g,"").replace(",","."));};
const date=x=>{if(x==null||x==="")return null;if(typeof x==="number"||(/^\d+(\.\d+)?$/.test(norm(x))&&Number(x)>20000)){const d=new Date(Date.UTC(1899,11,30)+Math.floor(Number(x))*86400000);return Number.isNaN(d.getTime())?null:d;}const d=new Date(/^\d{4}-\d\d-\d\d$/.test(norm(x))?norm(x)+"T00:00:00Z":x);return Number.isNaN(d.getTime())?null:d;};
function effective(record, at){const start=date(record.Vigencia_Inicio),end=date(record.Vigencia_Fim);if((norm(record.Vigencia_Inicio)&&!start)||(norm(record.Vigencia_Fim)&&!end)||(start&&end&&start>end))throw new RHError("Vigência inválida");return (!start||start<=at)&&(!end||end>=at);}
const active=r=>["APROVADO","EM VIGOR","EM_VIGOR","ACTIVO","ATIVO"].includes(upper(r.Estado));
function rank(r,ctx){const lev=upper(r.Nivel_Aplicacao);if(["GERAL","GLOBAL"].includes(lev))return 0;if(lev==="EMPRESA"&&upper(r.Empresa)===upper(ctx.empresa)&&!!ctx.empresa)return 1;if(lev==="CATEGORIA"&&upper(r.Categoria)===upper(ctx.categoria)&&(!norm(r.Empresa)||upper(r.Empresa)===upper(ctx.empresa))&&!!ctx.categoria)return 2;if(lev==="COLABORADOR"&&upper(r.Codigo_Colaborador)===upper(ctx.codigo)&&(!norm(r.Empresa)||upper(r.Empresa)===upper(ctx.empresa))&&!!ctx.codigo)return 3;return -1;}
function select(rows,key,ctx={},at=new Date(),keyName="Codigo_Parametro"){const when=date(at);if(!when)throw new RHError("Data de referência inválida");const candidates=rows.filter(r=>upper(r[keyName])===upper(key)&&active(r)&&effective(r,when)).map(r=>({row:r,level:rank(r,ctx)})).filter(x=>x.level>=0).sort((a,b)=>b.level-a.level||(date(b.row.Vigencia_Inicio)?.getTime()??0)-(date(a.row.Vigencia_Inicio)?.getTime()??0)||num(b.row.Versao)-num(a.row.Versao));if(!candidates.length)return null;const chosen=candidates[0];if(candidates[1]&&candidates[1].level===chosen.level&&date(candidates[1].row.Vigencia_Inicio)?.getTime()===date(chosen.row.Vigencia_Inicio)?.getTime()&&num(candidates[1].row.Versao)===num(chosen.row.Versao))throw new RHError("Configurações ambíguas para "+key,"RH_AMBIGUOUS");return chosen.row;}
export function resolveConfig(rows,key,ctx={},at=new Date()){return select(rows,key,ctx,at);}
export function resolveFormula(rows,key,ctx={},at=new Date()){const chosen=select(rows,key,ctx,at,"Codigo_Formula");if(chosen&&upper(chosen.Validada)!=="SIM"&&upper(chosen.Validada)!=="TRUE")throw new RHError("Fórmula ainda não validada: "+key,"RH_UNVALIDATED");return chosen;}
export function configValues(rows,ctx={},at=new Date()){const values={};for(const key of new Set(rows.map(r=>r.Codigo_Parametro).filter(Boolean))){const r=resolveConfig(rows,key,ctx,at);if(r)values[key]=["INTEIRO","DECIMAL","PERCENTAGEM","NÚMERO","NUMERO"].includes(upper(r.Tipo_Valor))?num(r.Valor):r.Valor;}return values;}
export function validateScales(rows,tableId,ctx={},at=new Date()){const items=rows.filter(r=>r.ID_Tabela_Escaloes===tableId&&active(r)&&effective(r,date(at))&&rank(r,ctx)>=0).sort((a,b)=>num(a.Limite_Inferior)-num(b.Limite_Inferior));if(!items.length)throw new RHError("Nenhum escalão disponível para "+tableId);let previous=null;for(const item of items){const low=num(item.Limite_Inferior),high=norm(item.Limite_Superior)?num(item.Limite_Superior):Infinity;if(!Number.isFinite(low)||Number.isNaN(high)||high<low)throw new RHError("Limites inválidos no escalão "+item.ID_Escalao);if(previous!==null&&low<=previous)throw new RHError("Escalões sobrepostos ou duplicados","RH_SCALE_OVERLAP");if(previous===Infinity)throw new RHError("Escalão aberto antes do último","RH_SCALE_OVERLAP");previous=high;}return items;}
export function findScale(rows,tableId,amount,ctx={},at=new Date()){const n=num(amount);if(!Number.isFinite(n))throw new RHError("Base do escalão inválida");return validateScales(rows,tableId,ctx,at).find(r=>n>=num(r.Limite_Inferior)&&(norm(r.Limite_Superior)===""||n<=num(r.Limite_Superior)))??null;}
function parseSource(src){if(typeof src!=="string"||src.length>10000)throw new RHError("Expressão inválida ou demasiado longa");src=src.trim();const tokens=[];const re=/\s*(>=|<=|==|!=|[()+*\/,<>-]|"(?:[^"\\]|\\.)*"|'(?:[^'\\]|\\.)*'|\d+(?:\.\d+)?|[A-Za-z_][A-Za-z_0-9]*)/gy;let pos=0,m;while(pos<src.length){re.lastIndex=pos;m=re.exec(src);if(!m)throw new RHError("Símbolo não permitido na posição "+pos);tokens.push(m[1]);pos=re.lastIndex;if(tokens.length>1500)throw new RHError("Expressão demasiado complexa");}return tokens;}
export function evaluateFormula(expression, variables={}) {
 const tokens=parseSource(expression);let i=0,depth=0;
 const peek=()=>tokens[i],eat=x=>peek()===x?(i++,true):false;
 function primary(){
  if(++depth>100)throw new RHError("Fórmula demasiado profunda");
  try {
   const token=tokens[i++];
   if(token===undefined)throw new RHError("Expressão incompleta");
   if(token==="+"||token==="-")return {type:"unary",op:token,value:primary()};
   if(token==="("){const value=compare();if(!eat(")"))throw new RHError("Falta fechar parênteses");return value;}
   if(/^\d/.test(token))return {type:"literal",value:Number(token)};
   if(token[0]==='"'||token[0]==="'")return {type:"literal",value:token.slice(1,-1)};
   if(/^[A-Za-z_]/.test(token)){
    if(eat("(")){
     const args=[];if(!eat(")")){do{args.push(compare());}while(eat(","));if(!eat(")"))throw new RHError("Argumentos inválidos");}
     const name=upper(token),arity={IF:[3,3],ROUND:[2,2],ROUNDUP:[2,2],ROUNDDOWN:[2,2],MIN:[1,100],MAX:[1,100],ABS:[1,1],AND:[1,100],OR:[1,100],NOT:[1,1],NETWORKDAYS:[2,3]}[name];
     if(!arity||args.length<arity[0]||args.length>arity[1])throw new RHError("Função desconhecida ou argumentos inválidos: "+token);
     return {type:"call",name,args};
    }
    return {type:"variable",name:upper(token)};
   }
   throw new RHError("Símbolo inválido");
  }finally{depth--;}
 }
 function binary(next,ops){let left=next();while(ops.includes(peek())){const op=tokens[i++];left={type:"binary",op,left,right:next()};}return left;}
 const mul=()=>binary(primary,["*","/"]),add=()=>binary(mul,["+","-"]),compare=()=>binary(add,[">","<",">=","<=","==","!="]);
 const ast=compare();if(i!==tokens.length)throw new RHError("Elementos inesperados após a fórmula");
 const vars=new Map(Object.entries(variables).map(([key,value])=>[upper(key),value]));
 function numeric(value){if(value===null||typeof value==="boolean"||value===""||!['number','string'].includes(typeof value)||!Number.isFinite(Number(value)))throw new RHError("Operação com valor não numérico");return Number(value);}
 function round(value,digits,mode){const n=numeric(value),d=numeric(digits);if(!Number.isInteger(d)||d<0||d>8)throw new RHError("Arredondamento inválido");const k=10**d,scaled=Math.abs(n)*k,eps=Number.EPSILON*Math.max(1,scaled)*2;return Math.sign(n)*(mode==="ROUND"?Math.floor(scaled+.5+eps):mode==="ROUNDUP"?Math.ceil(scaled-eps):Math.floor(scaled+eps))/k;}
 function networkdays(a,b,holidays=[]){
  const s=date(a),e=date(b);if(!s||!e||!Array.isArray(holidays))throw new RHError("Datas ou feriados inválidos em NETWORKDAYS");
  const days=holidays.map(h=>date(h));if(days.some(d=>!d))throw new RHError("Feriado inválido");
  const set=new Set(days.map(d=>d.toISOString().slice(0,10))),start=Date.UTC(s.getUTCFullYear(),s.getUTCMonth(),s.getUTCDate()),end=Date.UTC(e.getUTCFullYear(),e.getUTCMonth(),e.getUTCDate()),sign=start<=end?1:-1;
  const span=Math.abs((end-start)/86400000);if(span>3660)throw new RHError("Período de NETWORKDAYS excessivo");let count=0;
  for(let k=0;k<=span;k++){const d=new Date(Math.min(start,end)+k*86400000);if(d.getUTCDay()!==0&&d.getUTCDay()!==6&&!set.has(d.toISOString().slice(0,10)))count++;}return count*sign;
 }
 function run(node){
  if(node.type==="literal")return node.value;
  if(node.type==="variable"){
   if(node.name==="TRUE")return true;if(node.name==="FALSE")return false;
   if(!vars.has(node.name))throw new RHError("Variável desconhecida: "+node.name,"RH_MISSING_VARIABLE");return vars.get(node.name);
  }
  if(node.type==="unary")return numeric(run(node.value))*(node.op==="-"?-1:1);
  if(node.type==="binary"){
   const a=run(node.left),b=run(node.right),op=node.op;
   if(["+","-","*","/"].includes(op)){const x=numeric(a),y=numeric(b);if(op==="/"&&y===0)throw new RHError("Divisão por zero");return op==="+"?x+y:op==="-"?x-y:op==="*"?x*y:x/y;}
   const strings=typeof a==="string"&&typeof b==="string",bools=typeof a==="boolean"&&typeof b==="boolean";
   const x=strings?upper(a):bools?a:numeric(a),y=strings?upper(b):bools?b:numeric(b);
   return op===">"?x>y:op==="<"?x<y:op===">="?x>=y:op==="<="?x<=y:op==="=="?x===y:x!==y;
  }
  // IF só avalia o ramo seleccionado, como as condições das fórmulas salariais exigem.
  if(node.name==="IF")return run(node.args[0])?run(node.args[1]):run(node.args[2]);
  const args=node.args.map(run);
  switch(node.name){
   case"ROUND":case"ROUNDUP":case"ROUNDDOWN":return round(...args,node.name);
   case"MIN":return Math.min(...args.map(numeric));case"MAX":return Math.max(...args.map(numeric));case"ABS":return Math.abs(numeric(args[0]));
   case"AND":return args.every(Boolean);case"OR":return args.some(Boolean);case"NOT":return !args[0];case"NETWORKDAYS":return networkdays(...args);
  }
 }
 const result=run(ast);if(typeof result==="number"&&!Number.isFinite(result))throw new RHError("Resultado não finito");return result;
}
