// Página «Acessos»: utilizadores da tbl_utilizadores_RH (só Superadministrador; o Office Script confirma o perfil).
// O código de cada pessoa nunca é lido nem mostrado: ao criar ou repor, o site gera um código temporário,
// calcula a chave SHA-256(email:código) e só a chave vai para o Excel (Trocar_Codigo = Sim).
const $ = id => document.getElementById(id);
const texto = v => String(v ?? "").trim();
export const PERFIS = ["Gestor", "Técnico RH", "Administrador RH", "Superadministrador"];

// Linhas do script -> objectos da página.
export function utilizadores(linhas = []) {
  return linhas.map(l => ({
    id: texto(l.ID_Utilizador), nome: texto(l.Nome), email: texto(l.Email).toLowerCase(), perfil: texto(l.Perfil),
    empresa: texto(l.Empresa), estado: texto(l.Estado) || "Activo", temCodigo: l.Tem_Codigo === true || l.Tem_Codigo === "true",
    trocar: /^sim$/i.test(texto(l.Trocar_Codigo)), alterado: l["Data Ultima Alteracao"], alteradoPor: texto(l["Utilizador Alteracao"])
  })).filter(u => u.email);
}
export function filtrar(lista, {procurar = "", estado = ""} = {}) {
  const q = procurar.trim().toLowerCase();
  return lista.filter(u => (!estado || u.estado.toLowerCase() === estado.toLowerCase()) &&
    (!q || u.nome.toLowerCase().includes(q) || u.email.includes(q)))
    .sort((a, b) => (a.estado === b.estado ? 0 : a.estado === "Activo" ? -1 : 1) || a.nome.localeCompare(b.nome, "pt"));
}
export function validar(u, {novo, lista = [], codigo = "", minimo = 8}) {
  const erros = [];
  if (!texto(u.nome)) erros.push("Indique o nome.");
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(texto(u.email))) erros.push("E-mail inválido.");
  else if (novo && lista.some(x => x.email === texto(u.email).toLowerCase())) erros.push("Já existe um utilizador com este e-mail.");
  if (!PERFIS.includes(u.perfil)) erros.push("Perfil inválido.");
  if ((novo || codigo) && codigo.length < minimo) erros.push("O código temporário tem de ter pelo menos " + minimo + " caracteres.");
  return erros;
}
function dataSerie(v) {
  const n = Number(v);
  if (!Number.isFinite(n) || n <= 0) return "";
  const d = new Date(Math.round((n - 25569) * 86400000));
  return d.toLocaleDateString("pt-MZ", {timeZone: "UTC"});
}

export function iniciarAcessos(ctx) {
  let lista = [], editar = null, aCarregar = false;
  const dlg = $("ac-dlg"), feito = $("ac-feito");
  const abrirDlg = d => (typeof d.showModal === "function" ? d.showModal() : d.setAttribute("open", ""));
  const fecharDlg = d => (typeof d.close === "function" ? d.close() : d.removeAttribute("open"));

  function render() {
    const body = $("ac-body");
    body.replaceChildren();
    const vis = filtrar(lista, {procurar: $("ac-procurar").value, estado: $("ac-estado").value});
    if (!vis.length) {
      const tr = document.createElement("tr"), td = document.createElement("td");
      td.colSpan = 7; td.textContent = aCarregar ? "A ler os utilizadores…" : lista.length ? "Nenhum utilizador corresponde ao filtro." : "Sem utilizadores.";
      tr.append(td); body.append(tr);
    }
    for (const u of vis) {
      const tr = document.createElement("tr");
      tr.tabIndex = 0;
      const cel = [u.nome, u.email, u.perfil, u.empresa || "—", null, null, dataSerie(u.alterado) + (u.alteradoPor ? " · " + u.alteradoPor : "")];
      cel.forEach((c, i) => {
        const td = document.createElement("td");
        if (i === 4) { const s = document.createElement("span"); s.className = "state" + (u.estado === "Activo" ? "" : " off"); s.textContent = u.estado; td.append(s); }
        else if (i === 5) { const s = document.createElement("span"); s.className = "tag" + (u.temCodigo && !u.trocar ? " green" : ""); s.textContent = !u.temCodigo ? "Sem código" : u.trocar ? "Temporário" : "Definido"; td.append(s); }
        else { td.textContent = c; if (i === 0) td.style.fontWeight = "700"; }
        tr.append(td);
      });
      const abrir = () => abrirFicha(u);
      tr.addEventListener("click", abrir);
      tr.addEventListener("keydown", e => { if (e.key === "Enter") abrir(); });
      body.append(tr);
    }
    const act = lista.filter(u => u.estado === "Activo").length;
    $("ac-total").textContent = lista.length ? lista.length + " utilizadores · " + act + " activos" : "—";
  }

  async function carregar() {
    $("ac-ligacao").hidden = ctx.temUrl();
    if (!ctx.temUrl()) { lista = []; $("ac-fonte").textContent = "Falta o endereço de gestão."; render(); return; }
    aCarregar = true; render();
    try { lista = utilizadores(await ctx.listar()); $("ac-fonte").textContent = "Actualizado às " + new Date().toLocaleTimeString("pt-MZ"); }
    catch (e) { lista = []; $("ac-fonte").textContent = ""; ctx.aviso("Não foi possível ler os acessos: " + e.message); }
    finally { aCarregar = false; render(); }
  }

  function codigoCampos(mostrar) {
    $("ac-codigo-campos").hidden = !mostrar;
    if (mostrar && !$("ac-codigo").value) $("ac-codigo").value = ctx.gerarCodigo();
  }
  function abrirFicha(u) {
    editar = u || null;
    $("ac-dlg-title").textContent = u ? "Alterar acesso" : "Novo utilizador";
    $("ac-nome").value = u?.nome || ""; $("ac-email").value = u?.email || ""; $("ac-email").readOnly = Boolean(u);
    $("ac-perfil").value = u?.perfil && PERFIS.includes(u.perfil) ? u.perfil : "Gestor";
    $("ac-empresa").value = u?.empresa || ""; $("ac-estado-dlg").value = u?.estado === "Inactivo" ? "Inactivo" : "Activo";
    $("ac-estado-campo").hidden = !u;
    const proprio = u && u.email === ctx.eu();
    $("ac-perfil").disabled = Boolean(proprio); $("ac-estado-dlg").disabled = Boolean(proprio);
    $("ac-novo-codigo").checked = !u; $("ac-novo-codigo").parentElement.hidden = !u;
    $("ac-codigo").value = ""; codigoCampos(!u);
    $("ac-erro").textContent = proprio ? "É a sua conta: perfil e estado não podem ser alterados aqui." : "";
    const dl = $("ac-empresas"); dl.replaceChildren(); ctx.empresas().forEach(e => { const o = document.createElement("option"); o.value = e; dl.append(o); });
    abrirDlg(dlg); $("ac-nome").focus();
  }

  async function gravar(ev) {
    ev.preventDefault();
    const novo = !editar, comCodigo = novo || $("ac-novo-codigo").checked;
    const u = {nome: $("ac-nome").value.trim(), email: $("ac-email").value.trim().toLowerCase(), perfil: $("ac-perfil").value,
      empresa: $("ac-empresa").value.trim(), estado: novo ? "Activo" : $("ac-estado-dlg").value};
    const codigo = comCodigo ? $("ac-codigo").value.trim() : "";
    const erros = validar(u, {novo, lista, codigo: comCodigo ? codigo : "", minimo: ctx.minimo});
    if (erros.length) { $("ac-erro").textContent = erros.join(" "); return; }
    const b = $("ac-gravar"); b.disabled = true; b.textContent = "A gravar…"; $("ac-erro").textContent = "";
    try {
      const chave = comCodigo ? await ctx.chave(u.email, codigo) : "";
      let r = null;
      const mudou = !editar || ["nome", "perfil", "empresa", "estado"].some(k => texto(editar[k]) !== texto(u[k]));
      if (mudou) r = await ctx.gravar(novo ? {...u, chave} : u);
      if (!novo && comCodigo) r = await ctx.codigo({email: u.email, chave});
      if (r) lista = utilizadores(r);
      fecharDlg(dlg); render();
      if (comCodigo && u.email !== ctx.eu()) {
        $("ac-feito-texto").textContent = (novo ? "Criado o acesso de " : "Novo código temporário para ") + u.nome + " (" + u.perfil + "). Envie-lhe o link e o código.";
        try { $("ac-link").value = ctx.link(u); } catch (e) { $("ac-link").value = e.message; }
        $("ac-feito-codigo").value = codigo;
        abrirDlg(feito);
      } else ctx.aviso(mudou || comCodigo ? "Acesso gravado." : "Sem alterações.");
    } catch (e) { $("ac-erro").textContent = "Não foi gravado: " + e.message; }
    finally { b.disabled = false; b.textContent = "Gravar"; }
  }

  $("ac-novo").addEventListener("click", () => { if (!ctx.temUrl()) { ctx.aviso("Indique primeiro o endereço de gestão."); return; } abrirFicha(null); });
  $("ac-form").addEventListener("submit", gravar);
  $("ac-cancelar").addEventListener("click", () => fecharDlg(dlg));
  $("ac-novo-codigo").addEventListener("change", e => codigoCampos(e.target.checked));
  $("ac-gerar").addEventListener("click", () => { $("ac-codigo").value = ctx.gerarCodigo(); });
  $("ac-feito-fechar").addEventListener("click", () => { $("ac-feito-codigo").value = ""; $("ac-link").value = ""; fecharDlg(feito); });
  feito.addEventListener("close", () => { $("ac-feito-codigo").value = ""; $("ac-link").value = ""; });
  document.querySelectorAll("[data-copiar]").forEach(btn => btn.addEventListener("click", async () => {
    const v = $(btn.dataset.copiar).value;
    try { await navigator.clipboard.writeText(v); btn.textContent = "Copiado"; setTimeout(() => { btn.textContent = "Copiar"; }, 1500); }
    catch { $(btn.dataset.copiar).select(); }
  }));
  $("ac-procurar").addEventListener("input", render);
  $("ac-estado").addEventListener("change", render);
  $("ac-url-guardar").addEventListener("click", () => { try { ctx.guardarUrl($("ac-url").value); $("ac-url").value = ""; carregar(); } catch (e) { ctx.aviso(e.message); } });

  return {abrir: carregar, limpar: () => { lista = []; render(); }};
}
