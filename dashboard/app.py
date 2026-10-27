# Dashboard RH — SIR Grupo
# Corre dentro do Sistema RH (stlite: Streamlit no navegador). O site escreve em dados/ o cadastro
# (tbl_funcionarios) e os lançamentos dos últimos meses (tbl_lançamento), lidos com a sessão do
# utilizador. Não há credenciais nem ligações a bases de dados neste ficheiro.
# Estrutura inspirada no Dashboard.py original: menu lateral de painéis + filtros + métricas + gráficos.

import json
import re
import unicodedata
from datetime import date, datetime, timedelta
from pathlib import Path

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

st.set_page_config(page_title="Dashboard RH", page_icon=":bar_chart:", layout="wide", initial_sidebar_state="expanded")

DADOS = Path("dados")
VERMELHO, TINTA, CINZA = "#be2537", "#182330", "#738092"
# Tons de vermelho da marca (do mais escuro ao mais claro)
PALETA = ["#9e1d2c", "#d2505e", "#5c0d18", "#efb0b6", "#be2537", "#7a1220", "#e27f89", "#f7d9dc"]
MESES = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]
MESES_CURTOS = ["Jan", "Fev", "Mar", "Abr", "Mai", "Jun", "Jul", "Ago", "Set", "Out", "Nov", "Dez"]

st.markdown(
    f"""<style>
    .block-container {{padding-top: 1.2rem; padding-bottom: 2rem;}}
    h1, h2, h3 {{color: {TINTA}; letter-spacing: -.3px;}}
    /* Cartões: rótulo, valor e nota sempre à mesma altura em todos os cartões */
    .kpi {{background:#fff;border:1px solid #e8edf2;border-radius:12px;padding:14px 16px;height:140px;box-sizing:border-box;
           display:flex;flex-direction:column;justify-content:flex-start;text-align:left;
           transition:transform .18s ease, box-shadow .18s ease, border-color .18s ease;cursor:default;}}
    .kpi:hover {{transform:translateY(-3px);border-color:#efb0b6;box-shadow:0 10px 24px rgba(158,29,44,.14);}}
    .kpi span {{font-size:12px;color:#69788a;font-weight:600;line-height:1.35;min-height:32px;
               display:-webkit-box;-webkit-line-clamp:2;-webkit-box-orient:vertical;overflow:hidden;}}
    .kpi b {{font-size:21px;color:{TINTA};margin:6px 0 4px;letter-spacing:-.5px;white-space:nowrap;line-height:1.2;transition:color .18s ease;}}
    .kpi:hover b {{color:{VERMELHO};}}
    .kpi small {{font-size:11px;color:#91a0ad;line-height:1.35;display:-webkit-box;-webkit-line-clamp:2;-webkit-box-orient:vertical;overflow:hidden;}}
    .nota {{font-size:12px;color:{CINZA};}}
    </style>""",
    unsafe_allow_html=True,
)


# ---------------------------------------------------------------- utilitários de dados
def _sem_acentos(s):
    return "".join(c for c in unicodedata.normalize("NFD", str(s)) if unicodedata.category(c) != "Mn").strip().lower()


def numero(v):
    """Aceita 17000, "17000", "17.000,50" (pt), "17,000.50" (en), "10 845,00"."""
    if v is None:
        return 0.0
    if isinstance(v, (int, float)):
        return float(v) if pd.notna(v) else 0.0
    s = re.sub(r"\s|MT|MZN", "", str(v), flags=re.I)
    if not s:
        return 0.0
    if re.search(r",\d{1,2}$", s):
        s = s.replace(".", "").replace(",", ".")
    else:
        s = s.replace(",", "")
    try:
        return float(s)
    except ValueError:
        return 0.0


def data(v):
    """Datas do Excel: número de série, dd/mm/aaaa, aaaa-mm-dd ou ISO."""
    if v is None or v == "":
        return pd.NaT
    if isinstance(v, (int, float)) or re.fullmatch(r"\d{4,6}(\.\d+)?", str(v).strip()):
        n = float(v)
        if 20000 < n < 80000:
            return pd.Timestamp(date(1899, 12, 30) + timedelta(days=int(n)))
        return pd.NaT
    s = str(v).strip()
    for fmt in ("%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d", "%d/%m/%y"):
        try:
            return pd.Timestamp(datetime.strptime(s[:10], fmt))
        except ValueError:
            pass
    return pd.to_datetime(s, errors="coerce", dayfirst=True)


def mes_num(v):
    if isinstance(v, (int, float)) and 1 <= int(v) <= 12:
        return int(v)
    k = _sem_acentos(v)[:3]
    for i, m in enumerate(MESES):
        if _sem_acentos(m)[:3] == k:
            return i + 1
    try:
        n = int(float(str(v)))
        return n if 1 <= n <= 12 else None
    except ValueError:
        return None


# Regras iguais às do dashboard VBA (modRH_Dashboard.bas)
IDADE_REFORMA = {"Masculino": 60, "Feminino": 55}


def sexo(v):
    k = _sem_acentos(v).upper()
    if k in ("MASCULINO", "M"):
        return "Masculino"
    if k in ("FEMININO", "F"):
        return "Feminino"
    return "Não indicado"


def reformado(v):
    return _sem_acentos(v).upper() in ("SIM", "REFORMADO", "REFORMADA")


def activo(v):
    return _sem_acentos(v).upper() in ("ACTIVO", "ATIVO")


def idade_completa(nasc, hoje):
    """Anos completos (como RH_Dash_IdadeActual)."""
    if pd.isna(nasc):
        return None
    anos = hoje.year - nasc.year - ((hoje.month, hoje.day) < (nasc.month, nasc.day))
    return anos


def antiguidade_anos(adm, hoje):
    """Meses completos / 12 (como RH_Dash_AntiguidadeActual); datas futuras ficam sem valor."""
    if pd.isna(adm) or adm > hoje:
        return None
    meses = (hoje.year - adm.year) * 12 + hoje.month - adm.month - (hoje.day < adm.day)
    return meses / 12


def chave(empresa, codigo):
    c = str(codigo or "").strip()
    c = c.zfill(4) if c.isdigit() else c.upper()
    return f"{str(empresa or '').strip().upper()}|{c}"


def mt(v, casas=0):
    s = f"{v:,.{casas}f}".replace(",", " ").replace(".", ",")
    return f"{s} MT"


def mt_curto(v):
    """Valores grandes em milhões para caberem nos cartões: 49,96 M MT."""
    if abs(v) >= 1e6:
        return f"{v / 1e6:,.2f} M MT".replace(",", " ").replace(".", ",")
    return mt(v)


def inteiro(v):
    return f"{int(round(v)):,}".replace(",", " ")


def pct(v):
    return f"{v:.1f}%".replace(".", ",")


# ---------------------------------------------------------------- carregamento
@st.cache_data(show_spinner=False)
def carregar(versao):
    def ler(nome):
        f = DADOS / nome
        return json.loads(f.read_text(encoding="utf-8")) if f.exists() else []

    meta = ler("meta.json") or {}
    func = pd.DataFrame(ler("funcionarios.json"))
    lanc = pd.DataFrame(ler("lancamentos.json"))

    # Cadastro
    if not func.empty:
        func["chave"] = [chave(e, c) for e, c in zip(func.get("Empresa", ""), func.get("Codigo", ""))]
        func["Sexo"] = func.get("Sexo", pd.Series([""] * len(func))).map(sexo)
        func["nascimento"] = func.get("Data Nascimento", pd.Series([None] * len(func))).map(data)
        func["admissao"] = func.get("Data Admissao", pd.Series([None] * len(func))).map(data)
        func["saida"] = func.get("Data Saida", pd.Series([None] * len(func))).map(data)
        func["activo"] = func.get("Estado", pd.Series([""] * len(func))).map(activo)
        func["Situação"] = func["activo"].map({True: "Activo", False: "Inactivo"})
        func["reformado"] = func.get("Estado de Reforma", pd.Series([""] * len(func))).map(reformado)
        for c in ("Empresa", "Departamento", "Profissao", "TipoContrato", "RegimePagamento", "Nacionalidade", "Estado Civil"):
            if c not in func:
                func[c] = ""
            func[c] = func[c].fillna("").astype(str).str.strip().replace("", "Não indicado")
        hoje = pd.Timestamp(date.today())
        func["idade"] = pd.to_numeric(func["nascimento"].map(lambda d: idade_completa(d, hoje)), errors="coerce")
        func["antiguidade"] = pd.to_numeric(func["admissao"].map(lambda d: antiguidade_anos(d, hoje)), errors="coerce")
        func["em_idade_reforma"] = [pd.notna(i) and sx in IDADE_REFORMA and i >= IDADE_REFORMA[sx] for i, sx in zip(func["idade"], func["Sexo"])]
        func["Nome"] = func.get("Nome", pd.Series([""] * len(func))).fillna("").astype(str).str.strip()

    # Lançamentos
    if not lanc.empty:
        lanc["chave"] = [chave(e, c) for e, c in zip(lanc.get("Empresa", ""), lanc.get("Codigo", ""))]
        lanc["ano"] = lanc.get("Ano", pd.Series([None] * len(lanc))).map(lambda v: int(numero(v)) if numero(v) else None)
        lanc["mes"] = lanc.get("Mês", pd.Series([None] * len(lanc))).map(mes_num)
        lanc = lanc.dropna(subset=["ano", "mes"])
        lanc["ano"] = lanc["ano"].astype(int)
        lanc["mes"] = lanc["mes"].astype(int)
        lanc["periodo"] = pd.to_datetime(dict(year=lanc["ano"], month=lanc["mes"], day=1))
        lanc["Período"] = [f"{MESES_CURTOS[m - 1]} {a}" for a, m in zip(lanc["ano"], lanc["mes"])]
        for c in VALORES:
            lanc[c] = lanc[c].map(numero) if c in lanc else 0.0
        for c in ("Empresa", "Departamento", "Profissao", "Estado do Lançamento", "RegimePagamento"):
            if c not in lanc:
                lanc[c] = ""
            lanc[c] = lanc[c].fillna("").astype(str).str.strip().replace("", "Não indicado")
        lanc["Custo empresa"] = lanc["Vencimento"] + lanc["INSS Empresa"]
        # Junta o perfil do cadastro (sexo, contrato, reforma, idade) a cada lançamento
        if not func.empty:
            perfil = func.drop_duplicates("chave").set_index("chave")[["Sexo", "TipoContrato", "reformado", "idade", "Situação"]]
            lanc = lanc.join(perfil, on="chave")
        lanc["Nome"] = lanc.get("Nome", pd.Series([""] * len(lanc))).fillna("").astype(str).str.strip()
        for c, v in (("Sexo", "Não indicado"), ("TipoContrato", "Não indicado"), ("reformado", False), ("Situação", "Não indicado")):
            if c not in lanc:
                lanc[c] = v
            lanc[c] = lanc[c].fillna(v)
    return meta, func, lanc


VALORES = ["Vencimento Base", "Vencimento", "Salario Liquido", "INSS Colaborador", "INSS Empresa", "Total INSS", "IRPS",
           "Total de Acréscimos", "Total de Descontos", "Subsídios", "Bónus", "Outros Acréscimos", "Horas Extras",
           "Faltas Justificadas", "Faltas Injustificadas", "Adiantamentos", "Empréstimos", "Outros Descontos", "Dias Trabalhados"]

versao_f = DADOS / "versao.txt"
meta, func, lanc = carregar(versao_f.read_text() if versao_f.exists() else "0")

if func.empty and lanc.empty:
    st.title("Dashboard RH")
    st.info("Sem dados. Abra o dashboard a partir do Sistema RH depois de iniciar sessão.")
    st.stop()


# ---------------------------------------------------------------- filtros (barra lateral)
PAINEIS = ["Custo do Capital Humano", "Perfil dos Colaboradores", "Movimentação e Presença", "Remuneração e Equidade"]
escolha = st.sidebar.selectbox("Escolha um Dashboard", PAINEIS)

periodos = sorted(lanc[["ano", "mes"]].drop_duplicates().itertuples(index=False, name=None)) if not lanc.empty else []
anos = sorted({a for a, _ in periodos}, reverse=True)
ano_sel = st.sidebar.selectbox("Seleccione o Ano:", ["Todos"] + [str(a) for a in anos])
meses_do_ano = [m for a, m in periodos if ano_sel == "Todos" or a == int(ano_sel)]
mes_sel = st.sidebar.selectbox("Seleccione o Mês:", ["Todos"] + [MESES[m - 1] for m in sorted(set(meses_do_ano))])


def opcoes(col):
    vals = set()
    if col in func:
        vals |= set(func[col].dropna().astype(str))
    if col in lanc:
        vals |= set(lanc[col].dropna().astype(str))
    return sorted(v for v in vals if v)


empresas = st.sidebar.multiselect("Empresa:", opcoes("Empresa"), placeholder="Todas")
deps_disp = sorted(set(
    (func[func["Empresa"].isin(empresas)] if empresas and not func.empty else func).get("Departamento", pd.Series(dtype=str)).astype(str))
    | set((lanc[lanc["Empresa"].isin(empresas)] if empresas and not lanc.empty else lanc).get("Departamento", pd.Series(dtype=str)).astype(str)))
departamentos = st.sidebar.multiselect("Departamento:", [d for d in deps_disp if d], placeholder="Todos")
nomes_disp = func
if not func.empty and empresas:
    nomes_disp = nomes_disp[nomes_disp["Empresa"].isin(empresas)]
if not func.empty and departamentos:
    nomes_disp = nomes_disp[nomes_disp["Departamento"].isin(departamentos)]
nome_sel = st.sidebar.selectbox("Colaborador:", ["Todos"] + sorted(n for n in set(nomes_disp.get("Nome", pd.Series(dtype=str))) if n))
estado_sel = st.sidebar.selectbox("Estado:", ["Todos", "Activo", "Inactivo"])
sexos = st.sidebar.multiselect("Sexo:", ["Masculino", "Feminino", "Não indicado"], placeholder="Todos")
regimes = st.sidebar.multiselect("Regime de pagamento:", opcoes("RegimePagamento"), placeholder="Todos")
contratos = st.sidebar.multiselect("Tipo de contrato:", opcoes("TipoContrato"), placeholder="Todos")
st.sidebar.markdown("---")
st.sidebar.markdown(
    f"<p class='nota'>Fonte: Gestão de Colaboradores.xlsx via Sistema RH<br>"
    f"{meta.get('meses', '')}<br>Actualizado: {meta.get('gerado', '—')}</p>", unsafe_allow_html=True)


def filtrar(df, periodo=True):
    if df.empty:
        return df
    m = pd.Series(True, index=df.index)
    if empresas and "Empresa" in df:
        m &= df["Empresa"].isin(empresas)
    if departamentos and "Departamento" in df:
        m &= df["Departamento"].isin(departamentos)
    if sexos and "Sexo" in df:
        m &= df["Sexo"].isin(sexos)
    if contratos and "TipoContrato" in df:
        m &= df["TipoContrato"].isin(contratos)
    if nome_sel != "Todos" and "Nome" in df:
        m &= df["Nome"].str.casefold() == nome_sel.casefold()
    if estado_sel != "Todos" and "Situação" in df:
        m &= df["Situação"] == estado_sel
    if regimes and "RegimePagamento" in df:
        m &= df["RegimePagamento"].isin(regimes)
    if periodo and "ano" in df:
        if ano_sel != "Todos":
            m &= df["ano"] == int(ano_sel)
        if mes_sel != "Todos":
            m &= df["mes"] == MESES.index(mes_sel) + 1
    return df[m]


L_todos = filtrar(lanc, periodo=False)  # tendências: todos os meses carregados, com os outros filtros
L = filtrar(lanc)                       # período seleccionado
F = filtrar(func)
F_act = F[F["activo"]] if not F.empty else F
def mes_referencia():
    """Mês para massa salarial, salário médio e variação: o seleccionado ou, sem mês, o último com dados."""
    if L_todos.empty:
        return None
    pers = sorted(L_todos[["ano", "mes"]].drop_duplicates().itertuples(index=False, name=None))
    if mes_sel != "Todos":
        a = int(ano_sel) if ano_sel != "Todos" else max(a for a, m in pers if m == MESES.index(mes_sel) + 1)
        return a, MESES.index(mes_sel) + 1
    cand = [p for p in pers if ano_sel == "Todos" or p[0] == int(ano_sel)]
    return cand[-1] if cand else None


def massa(ano, mes):
    d = L_todos[(L_todos["ano"] == ano) & (L_todos["mes"] == mes)]
    return d["Vencimento"].sum(), len(d)


def entradas_saidas(df):
    """Como RH_Dash_EntradasSaidas: admissões e saídas no mês seleccionado; sem mês, no ano seleccionado
    ou nos meses carregados."""
    if df.empty:
        return 0, 0, "—"
    if mes_sel != "Todos" and REF:
        a, m_ = REF
        f_ = lambda s_: (s_.dt.year == a) & (s_.dt.month == m_)
        rot = f"em {MESES[m_ - 1]} {a}"
    elif ano_sel != "Todos":
        f_ = lambda s_: s_.dt.year == int(ano_sel)
        rot = f"em {ano_sel}"
    else:
        ini = L_todos["periodo"].min() if not L_todos.empty else pd.Timestamp(date.today().replace(day=1)) - pd.DateOffset(months=11)
        f_ = lambda s_: s_ >= ini
        rot = "desde " + ini.strftime("%m/%Y")
    return int(f_(df["admissao"]).sum()), int(f_(df["saida"]).sum()), rot


REF = mes_referencia()
n_meses = max(L[["ano", "mes"]].drop_duplicates().shape[0], 1) if not L.empty else 1
rotulo_periodo = ("todos os meses carregados" if ano_sel == "Todos" and mes_sel == "Todos"
                  else f"{mes_sel if mes_sel != 'Todos' else 'todos os meses'} {'' if ano_sel == 'Todos' else ano_sel}").strip()


# ---------------------------------------------------------------- componentes visuais
def kpis(itens):
    cols = st.columns(len(itens))
    for col, (rotulo, valor, sub) in zip(cols, itens):
        col.markdown(f"<div class='kpi'><span>{rotulo}</span><b>{valor}</b><small>{sub}</small></div>", unsafe_allow_html=True)


def estilo(fig, altura=360, legenda=True):
    fig.update_layout(height=altura, margin=dict(l=10, r=10, t=40, b=10), paper_bgcolor="white", plot_bgcolor="white",
                      font=dict(family="Inter, Segoe UI, Arial", color=TINTA, size=12), title_font=dict(size=15),
                      showlegend=legenda, legend=dict(orientation="h", yanchor="bottom", y=1.0, xanchor="right", x=1, title_text=""),
                      separators=", ")
    if fig.layout.title.text:
        fig.update_layout(margin=dict(t=70))
    fig.update_xaxes(showgrid=False)
    fig.update_yaxes(gridcolor="#eef1f5")
    # Barras horizontais: o nome da categoria já está em cada barra, o título do eixo só ocupa espaço.
    if any(getattr(t, "orientation", None) == "h" for t in fig.data):
        fig.update_yaxes(title_text="")
    if fig.layout.xaxis.title.text in ("Período", "Mês"):
        fig.update_xaxes(title_text="")
    # Gráficos circulares: legenda por baixo, para não tapar o título.
    if any(t.type == "pie" for t in fig.data):
        fig.update_layout(legend=dict(orientation="h", yanchor="top", y=-0.02, xanchor="center", x=0.5), margin=dict(t=50, b=10))
    return fig


def grafico(fig, altura=360, legenda=True):
    st.plotly_chart(estilo(fig, altura, legenda), width="stretch", config={"displaylogo": False})


def sem_dados(msg="Sem dados para os filtros seleccionados."):
    st.info(msg)


def por_mes(df, cols):
    g = df.groupby(["periodo", "Período"], as_index=False)[cols].sum().sort_values("periodo")
    return g


# ================================================================ 1. Custo do Capital Humano
if escolha == "Custo do Capital Humano":
    st.title("Custo do Capital Humano")
    st.caption(f"Período: {rotulo_periodo}. Custo empresa = Vencimento + INSS a cargo da empresa.")
    if L.empty:
        sem_dados()
        st.stop()
    pessoas = L["chave"].nunique()
    custo = L["Custo empresa"].sum()
    kpis([
        ("Custo total empresa", mt_curto(custo), f"{n_meses} mês(es)"),
        ("Massa salarial (vencimento)", mt_curto(L["Vencimento"].sum()), "bruto"),
        ("Salário líquido pago", mt_curto(L["Salario Liquido"].sum()), "a transferir"),
        ("INSS total", mt_curto(L["INSS Colaborador"].sum() + L["INSS Empresa"].sum()),
         f"colab. {mt(L['INSS Colaborador'].sum())} · emp. {mt(L['INSS Empresa'].sum())}"),
        ("IRPS retido", mt_curto(L["IRPS"].sum()), pct(L["IRPS"].sum() / L["Vencimento"].sum() * 100) + " do vencimento" if L["Vencimento"].sum() else ""),
        ("Custo médio / colaborador", mt(custo / pessoas / n_meses if pessoas else 0), f"por mês · {inteiro(pessoas)} colaboradores"),
    ])
    # Indicadores do dashboard VBA: massa salarial, salário médio e variação face ao mês anterior
    if REF:
        a, m_ = REF
        ant = (a - 1, 12) if m_ == 1 else (a, m_ - 1)
        massa_ref, n_ref = massa(a, m_)
        massa_ant, _ = massa(*ant)
        var = massa_ref - massa_ant
        nome_ref = f"{MESES[m_ - 1]} {a}"
        st.markdown("<div style='height:12px'></div>", unsafe_allow_html=True)
        kpis([
            ("Massa salarial do mês", mt_curto(massa_ref), nome_ref + ("" if mes_sel != "Todos" else " (último mês)")),
            ("Salário médio", mt(massa_ref / n_ref if n_ref else 0), f"{inteiro(n_ref)} lançamentos em {nome_ref}"),
            ("Massa do mês anterior", mt_curto(massa_ant), f"{MESES[ant[1] - 1]} {ant[0]}" if massa_ant else f"{MESES[ant[1] - 1]} {ant[0]} · sem dados"),
            ("Variação MT", ("+" if var > 0 else "") + mt_curto(var), "face ao mês anterior"),
            ("Variação %", pct(var / massa_ant * 100) if massa_ant else "—", "face ao mês anterior"),
            ("Total de acréscimos", mt_curto(L["Total de Acréscimos"].sum()), f"descontos: {mt_curto(L['Total de Descontos'].sum())}"),
        ])
    st.markdown("####")

    m = por_mes(L_todos, ["Salario Liquido", "IRPS", "INSS Colaborador", "INSS Empresa", "Total de Descontos", "Custo empresa"])
    if not m.empty:
        m["Outros descontos"] = (m["Total de Descontos"] - m["IRPS"] - m["INSS Colaborador"]).clip(lower=0)
        long = m.melt(id_vars=["periodo", "Período"], value_vars=["Salario Liquido", "IRPS", "INSS Colaborador", "Outros descontos", "INSS Empresa"],
                      var_name="Componente", value_name="MT")
        long["Componente"] = long["Componente"].replace({"Salario Liquido": "Salário líquido", "INSS Colaborador": "INSS colaborador", "INSS Empresa": "INSS empresa"})
        fig = px.bar(long, x="Período", y="MT", color="Componente", title="Evolução mensal do custo (todos os meses carregados)",
                     color_discrete_sequence=["#7a1220", VERMELHO, "#efb0b6", "#f7d9dc", "#e27f89"])
        fig.add_trace(go.Scatter(x=m["Período"], y=m["Custo empresa"], mode="lines+markers", name="Custo empresa", line=dict(color="#3d0810", width=2)))
        grafico(fig, 400)

    c1, c2 = st.columns(2)
    with c1:
        e = L.groupby("Empresa", as_index=False)[["Custo empresa", "Salario Liquido", "IRPS", "INSS Empresa"]].sum().sort_values("Custo empresa")
        grafico(px.bar(e, x="Custo empresa", y="Empresa", orientation="h", title="Custo por empresa", text_auto=".3s",
                       color_discrete_sequence=[VERMELHO]), max(300, 40 * len(e) + 100), False)
    with c2:
        comp = pd.DataFrame({"Componente": ["Salário líquido", "IRPS", "INSS colaborador", "INSS empresa", "Outros descontos"],
                             "MT": [L["Salario Liquido"].sum(), L["IRPS"].sum(), L["INSS Colaborador"].sum(), L["INSS Empresa"].sum(),
                                    max(L["Total de Descontos"].sum() - L["IRPS"].sum() - L["INSS Colaborador"].sum(), 0)]})
        grafico(px.pie(comp, names="Componente", values="MT", hole=.55, title="Para onde vai o custo",
                       color_discrete_sequence=["#7a1220", VERMELHO, "#efb0b6", "#e27f89", "#f7d9dc"]), 360)

    d = L.groupby("Departamento", as_index=False).agg(custo=("Custo empresa", "sum"), pessoas=("chave", "nunique")).sort_values("custo", ascending=False).head(15)
    grafico(px.bar(d.sort_values("custo"), x="custo", y="Departamento", orientation="h", title="Custo por departamento (top 15)",
                   text_auto=".3s", color_discrete_sequence=["#7a1220"], labels={"custo": "MT"}), max(320, 30 * len(d) + 100), False)

    c3, c4 = st.columns(2)
    with c3:
        acr = pd.DataFrame({"Rubrica": ["Subsídios", "Bónus", "Horas extra (horas)", "Outros acréscimos"],
                            "Valor": [L["Subsídios"].sum(), L["Bónus"].sum(), L["Horas Extras"].sum(), L["Outros Acréscimos"].sum()]})
        acr = acr[acr["Rubrica"] != "Horas extra (horas)"]
        grafico(px.bar(acr, x="Rubrica", y="Valor", title="Acréscimos pagos", text_auto=".3s", color_discrete_sequence=["#e27f89"],
                       labels={"Valor": "MT"}), 320, False)
    with c4:
        dsc = pd.DataFrame({"Rubrica": ["Faltas e outros", "Adiantamentos", "Empréstimos"],
                            "Valor": [L["Outros Descontos"].sum(), L["Adiantamentos"].sum(), L["Empréstimos"].sum()]})
        grafico(px.bar(dsc, x="Rubrica", y="Valor", title="Descontos não fiscais", text_auto=".3s", color_discrete_sequence=["#d2505e"],
                       labels={"Valor": "MT"}), 320, False)

    st.subheader("Resumo por empresa e departamento")
    t = L.groupby(["Empresa", "Departamento"], as_index=False).agg(
        Colaboradores=("chave", "nunique"), Vencimento=("Vencimento", "sum"), INSS_colab=("INSS Colaborador", "sum"),
        INSS_emp=("INSS Empresa", "sum"), IRPS=("IRPS", "sum"), Liquido=("Salario Liquido", "sum"), Custo=("Custo empresa", "sum"),
    ).sort_values("Custo", ascending=False)
    t = t.rename(columns={"INSS_colab": "INSS colab.", "INSS_emp": "INSS emp.", "Liquido": "Líquido", "Custo": "Custo empresa"})
    st.dataframe(t, width="stretch", hide_index=True,
                 column_config={c: st.column_config.NumberColumn(format="%.2f") for c in ["Vencimento", "INSS colab.", "INSS emp.", "IRPS", "Líquido", "Custo empresa"]})

# ================================================================ 2. Perfil dos Colaboradores
elif escolha == "Perfil dos Colaboradores":
    st.title("Perfil dos Colaboradores")
    st.caption("Cadastro (tbl_funcionarios) com os filtros da barra lateral, como no dashboard do VBA. "
               "Use o filtro «Estado» para ver só activos ou inactivos.")
    if F.empty:
        sem_dados()
        st.stop()
    n = len(F)
    act = int(F["activo"].sum())
    homens = int((F["Sexo"] == "Masculino").sum())
    mulheres = int((F["Sexo"] == "Feminino").sum())
    idade_media = F["idade"].dropna().mean() if F["idade"].notna().any() else None
    antig_media = F["antiguidade"].dropna().mean() if F["antiguidade"].notna().any() else None
    ent, sai, rot_es = entradas_saidas(F)
    um = lambda v: f"{v:.1f}".replace(".", ",")
    kpis([
        ("Total", inteiro(n), "colaboradores no cadastro"),
        ("Activos", inteiro(act), pct(act / n * 100) + " do total"),
        ("Inactivos", inteiro(n - act), pct((n - act) / n * 100) + " do total"),
        ("Homens", inteiro(homens), pct(homens / n * 100)),
        ("Mulheres", inteiro(mulheres), pct(mulheres / n * 100)),
        ("Idade média", um(idade_media) + " anos" if idade_media is not None else "—", f"{inteiro(F['idade'].notna().sum())} com data de nascimento"),
    ])
    st.markdown("<div style='height:12px'></div>", unsafe_allow_html=True)
    kpis([
        ("Em idade de reforma", inteiro(F["em_idade_reforma"].sum()), "homens ≥ 60 · mulheres ≥ 55 anos"),
        ("Reformados", inteiro(F["reformado"].sum()), "«Estado de Reforma» = Sim / Reformado(a)"),
        ("Antiguidade média", um(antig_media) + " anos" if antig_media is not None else "—", "desde a admissão"),
        ("Entradas", inteiro(ent), rot_es + " · saldo " + ("+" if ent - sai > 0 else "") + inteiro(ent - sai)),
        ("Saídas", inteiro(sai), rot_es),
        ("Admitidos em " + str(date.today().year), inteiro(((F["admissao"].dt.year == date.today().year) & F["activo"]).sum()), "activos"),
    ])
    st.markdown("####")

    CORES_SEXO = {"Masculino": "#7a1220", "Feminino": VERMELHO, "Não indicado": "#e6e1e3"}
    c1, c2, c3 = st.columns(3)
    with c1:
        e = F.groupby("Empresa").size().sort_values().reset_index(name="Colaboradores")
        grafico(px.bar(e, x="Colaboradores", y="Empresa", orientation="h", title="Por empresa", text_auto=True,
                       color_discrete_sequence=["#9e1d2c"]), max(300, 34 * len(e) + 110), False)
    with c2:
        s_ = F["Sexo"].value_counts().reset_index()
        s_.columns = ["Sexo", "Colaboradores"]
        grafico(px.pie(s_, names="Sexo", values="Colaboradores", hole=.55, title="Sexo", color="Sexo", color_discrete_map=CORES_SEXO), 330)
    with c3:
        rg = F["RegimePagamento"].value_counts().reset_index()
        rg.columns = ["Regime", "Colaboradores"]
        grafico(px.pie(rg, names="Regime", values="Colaboradores", hole=.55, title="Regime de pagamento",
                       color_discrete_sequence=PALETA), 330)

    c4, c5 = st.columns(2)
    with c4:
        # Faixas etárias do VBA (RH_Dash_FaixasEtarias)
        faixas = pd.cut(F["idade"], bins=[-1, 24.999, 34.999, 44.999, 54.999, 59.999, 200],
                        labels=["< 25", "25 - 34", "35 - 44", "45 - 54", "55 - 59", "60+"])
        p_ = F.assign(Faixa=faixas).dropna(subset=["Faixa"]).groupby(["Faixa", "Sexo"], observed=False).size().reset_index(name="n")
        if p_["n"].sum() == 0:
            sem_dados("Sem datas de nascimento no cadastro.")
        else:
            grafico(px.bar(p_, x="Faixa", y="n", color="Sexo", title="Faixa etária", color_discrete_map=CORES_SEXO,
                           labels={"n": "Colaboradores", "Faixa": ""}, text_auto=True), 360)
    with c5:
        # Faixas de antiguidade do VBA (RH_Dash_AntiguidadeFaixas)
        ant = pd.cut(F["antiguidade"], bins=[-1, 0.999, 3.999, 5.999, 10.999, 200],
                     labels=["< 1 ano", "1 - 3 anos", "4 - 5 anos", "6 - 10 anos", "> 10 anos"])
        a_ = ant.value_counts().reindex(ant.cat.categories, fill_value=0).reset_index()
        a_.columns = ["Antiguidade", "Colaboradores"]
        grafico(px.bar(a_, x="Antiguidade", y="Colaboradores", title="Antiguidade", text_auto=True,
                       color_discrete_sequence=["#9e1d2c"], labels={"Antiguidade": ""}), 360, False)

    c6, c7 = st.columns(2)
    with c6:
        d = F.groupby("Departamento").size().sort_values(ascending=False).head(20).sort_values().reset_index(name="Colaboradores")
        grafico(px.bar(d, x="Colaboradores", y="Departamento", orientation="h", title="Por departamento (top 20)",
                       text_auto=True, color_discrete_sequence=["#7a1220"]), max(320, 26 * len(d) + 100), False)
    with c7:
        r = F.assign(Reforma=np.where(F["reformado"], "Reformado", np.where(F["em_idade_reforma"], "Em idade de reforma", "Abaixo da idade"))) \
            ["Reforma"].value_counts().reset_index()
        r.columns = ["Situação", "Colaboradores"]
        grafico(px.pie(r, names="Situação", values="Colaboradores", hole=.55, title="Reforma",
                       color="Situação", color_discrete_map={"Reformado": "#5c0d18", "Em idade de reforma": VERMELHO, "Abaixo da idade": "#f7d9dc"}), 360)

    c8, c9 = st.columns(2)
    with c8:
        faixas5 = pd.cut(F["idade"], bins=[0, 25, 30, 35, 40, 45, 50, 55, 60, 65, 120],
                         labels=["<25", "25–29", "30–34", "35–39", "40–44", "45–49", "50–54", "55–59", "60–64", "65+"], right=False)
        pp_ = F.assign(Faixa=faixas5).dropna(subset=["Faixa"]).groupby(["Faixa", "Sexo"], observed=True).size().reset_index(name="n")
        if pp_.empty:
            sem_dados("Sem datas de nascimento no cadastro.")
        else:
            pp_["Colaboradores"] = pp_.apply(lambda x: -x["n"] if x["Sexo"] == "Masculino" else x["n"], axis=1)
            fig = px.bar(pp_, y="Faixa", x="Colaboradores", color="Sexo", orientation="h", title="Pirâmide etária",
                         color_discrete_map=CORES_SEXO, hover_data={"n": True, "Colaboradores": False})
            fig.update_layout(barmode="relative")
            fig.update_xaxes(tickvals=[], title="← Masculino · Feminino →")
            grafico(fig, 380)
    with c9:
        es = F.groupby(["Empresa", "Sexo"]).size().reset_index(name="Colaboradores")
        grafico(px.bar(es, x="Empresa", y="Colaboradores", color="Sexo", title="Colaboradores por empresa e sexo",
                       color_discrete_map=CORES_SEXO), 380)

    pr = F["Profissao"].value_counts().head(15).sort_values().reset_index()
    pr.columns = ["Profissão", "Colaboradores"]
    grafico(px.bar(pr, x="Colaboradores", y="Profissão", orientation="h", title="Profissões mais frequentes (top 15)",
                   text_auto=True, color_discrete_sequence=[VERMELHO]), 420, False)

    extra = [c for c in ("TipoContrato", "Nacionalidade", "Estado Civil") if c in F and F[c].nunique() > 1]
    if extra:
        cols = st.columns(len(extra))
        for col, c in zip(cols, extra):
            v = F[c].value_counts().head(8).reset_index()
            v.columns = [c, "Colaboradores"]
            with col:
                grafico(px.pie(v, names=c, values="Colaboradores", hole=.5, title={"TipoContrato": "Tipo de contrato"}.get(c, c),
                               color_discrete_sequence=PALETA), 300)

    # Lista do VBA: «Colaboradores abrangidos pelos filtros»
    st.subheader(f"Colaboradores abrangidos pelos filtros ({inteiro(n)})")
    lista = pd.DataFrame({
        "Código": F.get("Codigo", ""), "Nome": F["Nome"], "Empresa": F["Empresa"], "Departamento": F["Departamento"],
        "Profissão": F["Profissao"], "Sexo": F["Sexo"], "Idade": F["idade"], "Antiguidade (anos)": F["antiguidade"].round(1),
        "Regime": F["RegimePagamento"], "Estado": F.get("Estado", ""), "Estado de Reforma": F.get("Estado de Reforma", ""),
    }).sort_values(["Empresa", "Código"])
    st.dataframe(lista, width="stretch", hide_index=True, height=420,
                 column_config={"Idade": st.column_config.NumberColumn(format="%d"), "Antiguidade (anos)": st.column_config.NumberColumn(format="%.1f")})

# ================================================================ 3. Movimentação e Presença
elif escolha == "Movimentação e Presença":
    st.title("Movimentação e Presença")
    st.caption(f"Admissões e saídas do cadastro; faltas, horas extra e processamento dos lançamentos ({rotulo_periodo}).")
    meses_carregados = sorted(L_todos["periodo"].unique()) if not L_todos.empty else []
    inicio = pd.Timestamp(meses_carregados[0]) if meses_carregados else pd.Timestamp(date.today().replace(day=1)) - pd.DateOffset(months=11)
    adm = F[F["admissao"] >= inicio] if not F.empty else F
    sai = F[F["saida"] >= inicio] if not F.empty else F
    ent_p, sai_p, rot_p = entradas_saidas(F)
    media_pessoas = L_todos.groupby("periodo")["chave"].nunique().mean() if not L_todos.empty else 0
    rot = len(sai) / media_pessoas * 100 if media_pessoas else 0
    kpis([
        ("Admissões", inteiro(ent_p), rot_p + f" · {inteiro(len(adm))} desde " + inicio.strftime("%m/%Y")),
        ("Saídas", inteiro(sai_p), rot_p + f" · {inteiro(len(sai))} desde " + inicio.strftime("%m/%Y")),
        ("Rotatividade", pct(rot), "saídas / efectivo médio"),
        ("Faltas injustificadas", inteiro(L["Faltas Injustificadas"].sum()) if not L.empty else "0", f"justificadas: {inteiro(L['Faltas Justificadas'].sum()) if not L.empty else 0}"),
        ("Horas extra", inteiro(L["Horas Extras"].sum()) if not L.empty else "0", rotulo_periodo),
        ("Lançamentos finalizados", pct((L["Estado do Lançamento"].str.lower().str.startswith("finaliz")).mean() * 100) if not L.empty else "—", f"{inteiro(len(L))} lançamentos"),
    ])
    st.markdown("####")

    mov = []
    for _, r in adm.iterrows():
        mov.append({"periodo": r["admissao"].to_period("M").to_timestamp(), "Movimento": "Admissões", "n": 1})
    for _, r in sai.iterrows():
        mov.append({"periodo": r["saida"].to_period("M").to_timestamp(), "Movimento": "Saídas", "n": 1})
    c1, c2 = st.columns(2)
    with c1:
        if mov:
            mv = pd.DataFrame(mov).groupby(["periodo", "Movimento"], as_index=False)["n"].sum().sort_values("periodo")
            mv["Mês"] = mv["periodo"].map(lambda p: f"{MESES_CURTOS[p.month - 1]} {p.year}")
            mv.loc[mv["Movimento"] == "Saídas", "n"] *= -1
            fig = px.bar(mv, x="Mês", y="n", color="Movimento", title="Admissões e saídas por mês",
                         color_discrete_map={"Admissões": "#e27f89", "Saídas": VERMELHO}, labels={"n": "Colaboradores"})
            fig.update_layout(barmode="relative")
            grafico(fig, 360)
        else:
            sem_dados("Sem admissões nem saídas registadas no período carregado.")
    with c2:
        if not L_todos.empty:
            ef = L_todos.groupby(["periodo", "Período"], as_index=False)["chave"].nunique().sort_values("periodo")
            grafico(px.line(ef, x="Período", y="chave", markers=True, title="Efectivo processado por mês",
                            labels={"chave": "Colaboradores"}, color_discrete_sequence=["#9e1d2c"]), 360, False)

    c3, c4 = st.columns(2)
    with c3:
        if not L_todos.empty:
            f = por_mes(L_todos, ["Faltas Justificadas", "Faltas Injustificadas"]).melt(
                id_vars=["periodo", "Período"], var_name="Tipo", value_name="Dias")
            grafico(px.bar(f, x="Período", y="Dias", color="Tipo", barmode="group", title="Faltas por mês",
                           color_discrete_map={"Faltas Justificadas": "#efb0b6", "Faltas Injustificadas": VERMELHO}), 340)
    with c4:
        if not L_todos.empty:
            h = por_mes(L_todos, ["Horas Extras"])
            grafico(px.bar(h, x="Período", y="Horas Extras", title="Horas extra por mês", color_discrete_sequence=["#7a1220"]), 340, False)

    if not L.empty:
        c5, c6 = st.columns(2)
        with c5:
            fd = L.groupby("Departamento", as_index=False)[["Faltas Injustificadas", "Faltas Justificadas"]].sum()
            fd = fd[(fd["Faltas Injustificadas"] + fd["Faltas Justificadas"]) > 0].sort_values("Faltas Injustificadas").tail(15)
            if fd.empty:
                sem_dados("Sem faltas registadas no período.")
            else:
                grafico(px.bar(fd.melt(id_vars="Departamento", var_name="Tipo", value_name="Dias"), x="Dias", y="Departamento",
                               color="Tipo", orientation="h", title="Faltas por departamento (top 15)",
                               color_discrete_map={"Faltas Justificadas": "#efb0b6", "Faltas Injustificadas": VERMELHO}), 420)
        with c6:
            est = L_todos.groupby(["periodo", "Período", "Estado do Lançamento"], as_index=False).size().sort_values("periodo")
            grafico(px.bar(est, x="Período", y="size", color="Estado do Lançamento", title="Estado do processamento por mês",
                           labels={"size": "Lançamentos"}, color_discrete_sequence=["#f7d9dc", "#efb0b6", "#e27f89", VERMELHO, "#7a1220"]), 420)

    if not sai.empty and "Motivo de saída" in sai:
        mt_ = sai["Motivo de saída"].fillna("").replace("", "Não indicado").value_counts().reset_index()
        mt_.columns = ["Motivo", "Saídas"]
        grafico(px.bar(mt_, x="Saídas", y="Motivo", orientation="h", title="Motivos de saída", color_discrete_sequence=[VERMELHO]), 300, False)

# ================================================================ 4. Remuneração e Equidade
else:
    st.title("Remuneração e Equidade")
    st.caption(f"Vencimento por colaborador e por mês ({rotulo_periodo}). Médias calculadas por pessoa.")
    if L.empty:
        sem_dados()
        st.stop()
    pp = L.groupby(["chave", "Sexo", "Empresa", "Departamento", "Profissao", "TipoContrato"], as_index=False).agg(
        Vencimento=("Vencimento", "mean"), Liquido=("Salario Liquido", "mean"), IRPS=("IRPS", "mean"), Custo=("Custo empresa", "mean"))
    med_h = pp.loc[pp["Sexo"] == "Masculino", "Vencimento"].mean()
    med_m = pp.loc[pp["Sexo"] == "Feminino", "Vencimento"].mean()
    gap = (med_h - med_m) / med_h * 100 if med_h and med_m and pd.notna(med_h) and pd.notna(med_m) else None
    kpis([
        ("Vencimento médio", mt(pp["Vencimento"].mean()), "por colaborador / mês"),
        ("Vencimento mediano", mt(pp["Vencimento"].median()), "metade ganha menos"),
        ("Líquido médio", mt(pp["Liquido"].mean()), "por colaborador / mês"),
        ("Diferença H/M", pct(gap) if gap is not None else "—", "vencimento médio (homens − mulheres) / homens"),
        ("Taxa efectiva IRPS", pct(L["IRPS"].sum() / L["Vencimento"].sum() * 100) if L["Vencimento"].sum() else "—", "IRPS / vencimento"),
        ("Rácio máx/mín", f"{pp['Vencimento'].max() / pp.loc[pp['Vencimento'] > 0, 'Vencimento'].min():.1f}×".replace(".", ",") if (pp["Vencimento"] > 0).any() else "—", "vencimento"),
    ])
    st.markdown("####")
    c1, c2 = st.columns(2)
    with c1:
        s = pp.groupby("Sexo", as_index=False)[["Vencimento", "Liquido"]].mean()
        grafico(px.bar(s.melt(id_vars="Sexo", var_name="Medida", value_name="MT"), x="Sexo", y="MT", color="Medida", barmode="group",
                       title="Médias por sexo", text_auto=".3s", color_discrete_sequence=["#7a1220", "#e27f89"]), 360)
    with c2:
        faixas = pd.cut(pp["Vencimento"], bins=[-1, 10000, 15000, 20000, 30000, 50000, 100000, 1e12],
                        labels=["<10 mil", "10–15 mil", "15–20 mil", "20–30 mil", "30–50 mil", "50–100 mil", "100 mil+"])
        fx = pp.assign(Faixa=faixas).groupby(["Faixa", "Sexo"], observed=True).size().reset_index(name="Colaboradores")
        grafico(px.bar(fx, x="Faixa", y="Colaboradores", color="Sexo", title="Distribuição do vencimento (MT/mês)",
                       color_discrete_map={"Masculino": "#7a1220", "Feminino": VERMELHO, "Não indicado": "#e6e1e3"}), 360)

    c3, c4 = st.columns(2)
    with c3:
        pr = pp.groupby("Profissao", as_index=False).agg(Vencimento=("Vencimento", "mean"), n=("chave", "count"))
        pr = pr[pr["n"] >= 2].sort_values("Vencimento").tail(15)
        grafico(px.bar(pr, x="Vencimento", y="Profissao", orientation="h", title="Vencimento médio por profissão (top 15, ≥2 pessoas)",
                       text_auto=".3s", color_discrete_sequence=[VERMELHO], hover_data={"n": True}, labels={"Profissao": "Profissão"}), 440, False)
    with c4:
        dp = pp.groupby("Departamento", as_index=False)["Vencimento"].mean().sort_values("Vencimento").tail(15)
        grafico(px.bar(dp, x="Vencimento", y="Departamento", orientation="h", title="Vencimento médio por departamento (top 15)",
                       text_auto=".3s", color_discrete_sequence=["#7a1220"]), 440, False)

    c5, c6 = st.columns(2)
    with c5:
        grafico(px.box(pp, x="Empresa", y="Vencimento", color="Empresa", points="outliers", title="Dispersão do vencimento por empresa",
                       color_discrete_sequence=PALETA), 380, False)
    with c6:
        tc = pp.groupby("TipoContrato", as_index=False).agg(Vencimento=("Vencimento", "mean"), Colaboradores=("chave", "count"))
        grafico(px.bar(tc, x="TipoContrato", y="Vencimento", title="Vencimento médio por tipo de contrato", text_auto=".3s",
                       color_discrete_sequence=["#9e1d2c"], hover_data={"Colaboradores": True}, labels={"TipoContrato": "Tipo de contrato"}), 380, False)
