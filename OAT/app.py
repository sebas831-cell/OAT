import streamlit as st
import pandas as pd
import plotly.express as px
from pathlib import Path

# ============================================================
# CONFIGURACIÓN
# ============================================================

st.set_page_config(
    page_title="Dashboard Ejecutivo FEN 2026",
    page_icon="🌱",
    layout="wide",
)

ARCHIVO = Path(__file__).parent / "datos_fen_2026.xlsx"
HOJA = "Hoja1"

KIT_ALIMENTO = "Alimento suplementario"

NOMBRES_CORTOS = [
    ("ALIMENTO SUPLEMENTARIO", KIT_ALIMENTO),
    ("INSUMOS AGRICOLAS", "Insumos agrícolas"),
    ("VETERINARIO", "Veterinario"),
    ("RIEGO", "Riego"),
    ("SEMILLAS", "Semillas"),
    ("PRADERAS", "Praderas"),
    ("ABREVADERO", "Abrevadero"),
]

COLORES_KIT = {
    KIT_ALIMENTO: "#0284C7",
    "Insumos agrícolas": "#16A34A",
    "Veterinario": "#F97316",
    "Riego": "#06B6D4",
    "Semillas": "#84CC16",
    "Praderas": "#14B8A6",
    "Abrevadero": "#EAB308",
}

VERDES = [
    "#0284C7", "#16A34A", "#F97316", "#06B6D4", "#84CC16",
    "#14B8A6", "#EAB308", "#8B5CF6", "#EC4899", "#6366F1",
]
GRIS_OTROS = "#CBD5E1"
COLOR_UNICO = "#0284C7"
PLANTILLA = "plotly_white"
ALTO_DONA = 430
ALTO_FILA = 520


def caja(nombre):
    try:
        return st.container(border=True, key=f"card_{nombre}")
    except TypeError:
        return st.container(border=True)


# ============================================================
# ESTILO
# ============================================================

st.markdown(
    """
<style>
.stApp {background: linear-gradient(180deg, #CFE8D8 0%, #E3F2E8 35%, #DCEEE2 100%);}
header[data-testid="stHeader"] {background: transparent;}
section[data-testid="stSidebar"], div[data-testid="collapsedControl"],
div[data-testid="stSidebarCollapsedControl"] {display: none;}
.block-container {padding-top: 2.6rem; padding-bottom: 2rem; max-width: 1500px;}
.banner {background: linear-gradient(90deg, #14532D 0%, #16A34A 100%);
         border-radius: 14px; padding: 22px 28px; margin-bottom: 18px;}
.banner .t {color: white; font-size: 30px; font-weight: 800; line-height: 1.2;}
.banner .s {color: #DCFCE7; font-size: 14px; margin-top: 4px;}
.kpi {background: white; border: 1px solid #D1E7D8; border-radius: 12px;
      padding: 16px 18px; min-height: 108px; border-top: 5px solid var(--c, #16A34A);}
.kpi-titulo {color: #475569; font-size: 12px; font-weight: 700; letter-spacing: .05em;}
.kpi-valor {color: #14532D; font-size: 30px; font-weight: 800; margin-top: 4px;}
.kpi-sub {color: #64748B; font-size: 12px;}
.seccion {font-size: 19px; font-weight: 800; color: #14532D;
          border-left: 6px solid #16A34A; padding-left: 10px; margin: 26px 0 10px 0;}
.lider {color: #14532D; font-size: 19px; font-weight: 800; margin-top: 6px; line-height: 1.25; min-height: 48px;}
div[data-testid="stPlotlyChart"], div[data-testid="stPlotlyChart"] > div,
div[data-testid="stPlotlyChart"] .js-plotly-plot, div[data-testid="stPlotlyChart"] .plot-container {
    background: #FFFFFF !important;}
div[data-testid="stDataFrame"] {background: #FFFFFF !important;}
div[class*="st-key-card_"] {background: #FFFFFF !important; border-radius: 12px;
    border: 1px solid #B7DCC4 !important; padding: 6px 8px;}
div[data-testid="stVerticalBlockBorderWrapper"] {
    background: white; border-radius: 12px; border-color: #D1E7D8 !important;}
</style>
""",
    unsafe_allow_html=True,
)


# ============================================================
# CARGA Y TRANSFORMACIÓN DEL EXCEL
# ============================================================

def nombre_corto(nombre_kit):
    n = (
        str(nombre_kit).upper()
        .replace("Á", "A").replace("É", "E").replace("Í", "I")
        .replace("Ó", "O").replace("Ú", "U")
    )
    for clave, corto in NOMBRES_CORTOS:
        if clave in n:
            return corto
    return str(nombre_kit).title()


@st.cache_data
def cargar_datos():
    crudo = pd.read_excel(ARCHIVO, sheet_name=HOJA, header=None)

    fila_enc = None
    for i in range(len(crudo)):
        if str(crudo.iat[i, 1]).strip().upper() == "DEPARTAMENTO":
            fila_enc = i
            break
    if fila_enc is None:
        raise ValueError("No se encontró la fila con 'DEPARTAMENTO' en la hoja.")

    fila_kits = fila_enc - 1

    datos = crudo.iloc[fila_enc + 1:].copy()
    datos = datos[datos[1].notna()]
    datos = datos[~datos[0].astype(str).str.upper().str.contains("TOTAL", na=False)]

    partes = []
    for col in range(4, crudo.shape[1]):
        nombre = crudo.iat[fila_kits, col]
        if pd.isna(nombre):
            continue
        partes.append(
            pd.DataFrame(
                {
                    "DEPARTAMENTO": datos[1].astype(str).str.strip(),
                    "PROVINCIA": datos[2].astype(str).str.strip(),
                    "DISTRITO": datos[3].astype(str).str.strip(),
                    "KIT": nombre_corto(nombre),
                    "META": pd.to_numeric(datos[col], errors="coerce").fillna(0),
                    "BENEFICIARIOS": pd.to_numeric(datos[col + 1], errors="coerce").fillna(0),
                }
            )
        )

    largo = pd.concat(partes, ignore_index=True)

    largo["BENEF_EXCEL"] = largo["BENEFICIARIOS"]
    es_prad = largo["KIT"] == "Praderas"
    largo.loc[es_prad, "BENEFICIARIOS"] = largo.loc[es_prad, "META"] / 5 * 3
    largo["ATENDIDO"] = (largo["META"] > 0) | (largo["BENEFICIARIOS"] > 0)
    largo["UBIGEO"] = largo["DEPARTAMENTO"] + " / " + largo["PROVINCIA"] + " / " + largo["DISTRITO"]
    largo["DISTRITO_LBL"] = largo["DISTRITO"] + " (" + largo["PROVINCIA"] + ")"
    largo["PROVINCIA_LBL"] = largo["PROVINCIA"] + " (" + largo["DEPARTAMENTO"].str.title() + ")"
    return largo


try:
    base = cargar_datos()
except Exception as e:
    st.error(f"No se pudo leer el archivo '{ARCHIVO.name}': {e}")
    st.stop()


# ============================================================
# ENCABEZADO Y FILTROS (ARRIBA DE LA PÁGINA)
# ============================================================

st.markdown(
    """
<div class="banner">
    <div class="t">🌱 Dashboard Ejecutivo FEN 2026</div>
    <div class="s">Programación de beneficiarios por intervención ·
    Fuente actualizada al 16.09.26</div>
</div>
""",
    unsafe_allow_html=True,
)


def limpiar_filtros():
    for clave in ("f_dep", "f_prov", "f_dist", "f_kit"):
        st.session_state[clave] = []


def opciones_validas(clave, opciones):
    st.session_state[clave] = [
        x for x in st.session_state.get(clave, []) if x in opciones
    ]
    return opciones


with caja("filtros"):
    st.markdown("**🎛️ Filtros** · elige uno o varios; los gráficos se actualizan solos")

    metrica = "Beneficiarios"
    col_valor = "BENEFICIARIOS"

    data = base.copy()

    f1, f2, f3, f4, f5 = st.columns([3, 3, 3, 3, 1])
    with f1:
        ops = opciones_validas("f_dep", sorted(data["DEPARTAMENTO"].unique()))
        sel_dep = st.multiselect("Departamento", ops, key="f_dep", placeholder="Todos")
    if sel_dep:
        data = data[data["DEPARTAMENTO"].isin(sel_dep)]
    with f2:
        ops = opciones_validas("f_prov", sorted(data["PROVINCIA"].unique()))
        sel_prov = st.multiselect("Provincia", ops, key="f_prov", placeholder="Todas")
    if sel_prov:
        data = data[data["PROVINCIA"].isin(sel_prov)]
    with f3:
        ops = opciones_validas("f_dist", sorted(data["DISTRITO"].unique()))
        sel_dist = st.multiselect("Distrito", ops, key="f_dist", placeholder="Todos")
    if sel_dist:
        data = data[data["DISTRITO"].isin(sel_dist)]
    with f4:
        ops = opciones_validas("f_kit", sorted(data["KIT"].unique()))
        sel_kit = st.multiselect("Kit / intervención", ops, key="f_kit", placeholder="Todos")
    if sel_kit:
        data = data[data["KIT"].isin(sel_kit)]
    with f5:
        st.write("")
        st.write("")
        st.button("🧹 Limpiar", on_click=limpiar_filtros, help="Limpiar filtros")

d = data[data["ATENDIDO"]].copy()
d["VALOR"] = d[col_valor]

if d.empty:
    st.warning("No hay datos con los filtros seleccionados.")
    st.stop()


# ============================================================
# FUNCIONES AUXILIARES
# ============================================================

def seccion(texto):
    st.markdown(f'<div class="seccion">{texto}</div>', unsafe_allow_html=True)


def kpi(titulo, valor, sub, color):
    return f"""
    <div class="kpi" style="--c:{color}">
        <div class="kpi-titulo">{titulo}</div>
        <div class="kpi-valor">{valor}</div>
        <div class="kpi-sub">{sub}</div>
    </div>"""


def agrupar(df, col, top=None):
    g = df.groupby(col)["VALOR"].sum().sort_values(ascending=False)
    if top is not None and len(g) > top:
        resto = g.iloc[top:].sum()
        g = g.iloc[:top]
        g.loc["Otros"] = resto
    return g


def dona(df, col, titulo, top=None, colores=None):
    g = agrupar(df, col, top)
    if colores is None:
        colores = {}
        i = 0
        for nombre in g.index:
            if nombre == "Otros":
                colores[nombre] = GRIS_OTROS
            else:
                colores[nombre] = VERDES[i % len(VERDES)]
                i += 1
    fig = px.pie(
        names=g.index, values=g.values, hole=0.58, title=titulo,
        color=g.index, color_discrete_map=colores, template=PLANTILLA,
    )
    fig.update_traces(
        textposition="inside",
        textinfo="percent+value",
        hovertemplate="<b>%{label}</b><br>%{value:,.0f}<br>%{percent}<extra></extra>",
        sort=False,
        marker=dict(line=dict(color="white", width=2)),
    )
    fig.add_annotation(
        text=f"<b>{g.sum():,.0f}</b><br><span style='font-size:12px;color:#64748B'>total</span>",
        showarrow=False, font=dict(size=22, color="#14532D"),
    )
    fig.update_layout(
        height=ALTO_DONA,
        margin=dict(t=50, b=10, l=10, r=10),
        legend=dict(orientation="h", yanchor="top", y=-0.02, xanchor="center", x=0.5,
                    font=dict(size=11)),
        title=dict(font=dict(size=16, color="#14532D")),
    )
    return fig


def barras_h(df, col, titulo, top=15, apilado_kit=False, color=COLOR_UNICO, alto=ALTO_FILA):
    orden = agrupar(df, col, top=None)
    top_items = orden.head(top).index.tolist()
    sub = df[df[col].isin(top_items)]
    if apilado_kit:
        g = sub.groupby([col, "KIT"], as_index=False)["VALOR"].sum()
        fig = px.bar(
            g, x="VALOR", y=col, color="KIT", orientation="h",
            color_discrete_map=COLORES_KIT, title=titulo, template=PLANTILLA,
        )
        fig.update_layout(
            legend=dict(orientation="h", yanchor="top", y=-0.06, xanchor="center", x=0.5),
            legend_title_text="",
        )
    else:
        g = sub.groupby(col, as_index=False)["VALOR"].sum()
        fig = px.bar(
            g, x="VALOR", y=col, orientation="h", title=titulo, template=PLANTILLA,
            color_discrete_sequence=[color], text="VALOR",
        )
        fig.update_traces(texttemplate="%{text:,.0f}", textposition="outside", cliponaxis=False)
    fig.update_yaxes(categoryorder="array", categoryarray=top_items[::-1], title=None)
    fig.update_xaxes(title=None if apilado_kit else metrica)
    fig.update_layout(
        height=alto,
        margin=dict(t=50, b=90 if apilado_kit else 10, l=10, r=50),
        title=dict(font=dict(size=16, color="#14532D")),
    )
    return fig


def tarjeta(fig, clave):
    fig.update_layout(paper_bgcolor="#FFFFFF", plot_bgcolor="#FFFFFF")
    with caja(clave):
        st.plotly_chart(fig, key=clave, theme=None)


# ============================================================
# KPIS
# ============================================================

n_dep = d["DEPARTAMENTO"].nunique()
n_prov = d.groupby(["DEPARTAMENTO", "PROVINCIA"]).ngroups
n_dist = d["UBIGEO"].nunique()
n_kits = d["KIT"].nunique()
total_benef = d["BENEFICIARIOS"].sum()
total_meta = d["META"].sum()

k = st.columns(5)
k[0].markdown(kpi("DEPARTAMENTOS", f"{n_dep}", "con intervención", "#14532D"), unsafe_allow_html=True)
k[1].markdown(kpi("PROVINCIAS", f"{n_prov}", "con intervención", "#16A34A"), unsafe_allow_html=True)
k[2].markdown(kpi("DISTRITOS", f"{n_dist:,}", "atendidos", "#4ADE80"), unsafe_allow_html=True)
k[3].markdown(kpi("TIPOS DE KIT", f"{n_kits}", "intervenciones distintas", "#A3E635"), unsafe_allow_html=True)
k[4].markdown(
    kpi("BENEFICIARIOS", f"{total_benef:,.0f}", f"Meta física total: {total_meta:,.0f}", "#0D9488"),
    unsafe_allow_html=True,
)
st.caption(
    "ℹ️ Los beneficiarios se suman por intervención: una misma familia que recibe "
    "dos kits se cuenta en cada uno."
)


# ============================================================
# LO MÁS DESTACADO
# ============================================================

def lider(df, col):
    g = df.groupby(col)["VALOR"].sum().sort_values(ascending=False)
    return g.index[0], g.iloc[0], g.iloc[0] / g.sum() * 100


def tarjeta_lider(titulo, nombre, valor, pct, color):
    return (
        f'<div class="kpi" style="--c:{color}">'
        f'<div class="kpi-titulo">{titulo}</div>'
        f'<div class="lider">{nombre}</div>'
        f'<div class="kpi-sub"><b>{valor:,.0f}</b> {metrica.lower()} · {pct:.0f}% del total</div>'
        f'</div>'
    )


seccion("⭐ Lo más destacado")
l1, l2, l3, l4 = st.columns(4)
n, v, p = lider(d, "DEPARTAMENTO")
l1.markdown(tarjeta_lider("DEPARTAMENTO CON MÁS BENEFICIARIOS", n.title(), v, p, "#14532D"), unsafe_allow_html=True)
n, v, p = lider(d, "PROVINCIA_LBL")
l2.markdown(tarjeta_lider("PROVINCIA CON MÁS BENEFICIARIOS", n, v, p, "#16A34A"), unsafe_allow_html=True)
n, v, p = lider(d, "DISTRITO_LBL")
l3.markdown(tarjeta_lider("DISTRITO CON MÁS BENEFICIARIOS", n, v, p, "#4ADE80"), unsafe_allow_html=True)
n, v, p = lider(d, "KIT")
l4.markdown(tarjeta_lider("KIT CON MÁS BENEFICIARIOS", n, v, p, "#0D9488"), unsafe_allow_html=True)


# ============================================================
# 1. RESUMEN GENERAL
# ============================================================

seccion("Resumen general")
c1, c2 = st.columns(2)
with c1:
    tarjeta(dona(d, "KIT", f"{metrica} por kit", colores=COLORES_KIT, top=None), "dona_kit")
with c2:
    tarjeta(dona(d, "DEPARTAMENTO", f"{metrica} por departamento", top=None), "dona_dep")

_pr = d[d["KIT"] == "Praderas"]
if not _pr.empty:
    _ha = _pr["META"].sum()
    st.caption(
        f"🌱 Regla de Praderas aplicada en estos gráficos: {_ha:,.0f} ha ÷ 5 × 3 = "
        f"**{_ha / 5 * 3:,.0f} beneficiarios**."
    )


# ============================================================
# 2. DEPARTAMENTOS
# ============================================================

seccion("Departamentos")
c1, c2 = st.columns(2)
with c1:
    tarjeta(
        dona(d, "DEPARTAMENTO", f"{metrica} por departamento", top=None),
        "dep_dona",
    )
with c2:
    dist_dep = (
        d.groupby("DEPARTAMENTO")["UBIGEO"].nunique()
        .sort_values(ascending=False).reset_index(name="DISTRITOS")
    )
    fig = px.bar(
        dist_dep, x="DISTRITOS", y="DEPARTAMENTO", orientation="h", text="DISTRITOS",
        title="Distritos atendidos por departamento", template=PLANTILLA,
        color_discrete_sequence=[COLOR_UNICO],
    )
    fig.update_yaxes(categoryorder="array",
                     categoryarray=dist_dep["DEPARTAMENTO"].tolist()[::-1], title=None)
    fig.update_traces(textposition="outside", cliponaxis=False)
    fig.update_layout(height=ALTO_DONA, margin=dict(t=50, b=10, l=10, r=50),
                      title=dict(font=dict(size=16, color="#14532D")))
    tarjeta(fig, "dep_distritos")


# ============================================================
# 3. PROVINCIAS
# ============================================================

seccion("Provincias")
c1, c2 = st.columns([3, 2])
with c1:
    tarjeta(
        barras_h(d, "PROVINCIA_LBL", f"Top 15 provincias por {metrica.lower()}", top=15,
                 color=COLOR_UNICO, alto=ALTO_DONA),
        "prov_top",
    )
with c2:
    tarjeta(dona(d, "PROVINCIA", "Participación de provincias", top=10), "prov_dona")


# ============================================================
# 4. DISTRITOS
# ============================================================

seccion("Distritos")
c1, c2 = st.columns(2)
with c1:
    tarjeta(
        barras_h(d, "DISTRITO_LBL", f"Top 15 distritos por {metrica.lower()}",
                 top=15, color=COLOR_UNICO, alto=ALTO_DONA),
        "dist_top",
    )
with c2:
    tarjeta(
        barras_h(d, "DISTRITO_LBL", f"Top 15 distritos por tipo de kit",
                 top=15, apilado_kit=True, alto=ALTO_DONA),
        "dist_top_kit",
    )


# ============================================================
# 5. KITS (TABLA EN EXPANDER)
# ============================================================

seccion("Kits / intervenciones")
c1, c2 = st.columns(2)
with c1:
    rk = d.groupby("KIT", as_index=False)["VALOR"].sum().sort_values("VALOR", ascending=False)
    fig = px.bar(
        rk, x="KIT", y="VALOR", color="KIT", text="VALOR",
        color_discrete_map=COLORES_KIT, template=PLANTILLA, title=f"{metrica} por kit",
    )
    fig.update_traces(texttemplate="%{text:,.0f}", textposition="outside", cliponaxis=False)
    fig.update_layout(showlegend=False, height=ALTO_DONA, xaxis_title=None,
                      yaxis_title=metrica, margin=dict(t=50, b=10),
                      title=dict(font=dict(size=16, color="#14532D")))
    tarjeta(fig, "kit_valor")
with c2:
    cob = (
        d.groupby("KIT")["UBIGEO"].nunique()
        .sort_values(ascending=False).reset_index(name="DISTRITOS")
    )
    fig = px.bar(
        cob, x="KIT", y="DISTRITOS", color="KIT", text="DISTRITOS",
        color_discrete_map=COLORES_KIT, template=PLANTILLA,
        title="Cobertura: distritos atendidos por kit",
    )
    fig.update_traces(textposition="outside", cliponaxis=False)
    fig.update_layout(showlegend=False, height=ALTO_DONA, xaxis_title=None,
                      margin=dict(t=50, b=10),
                      title=dict(font=dict(size=16, color="#14532D")))
    tarjeta(fig, "kit_cobertura")

with st.expander("📂 Desplegar resumen por kit", expanded=False):
    with caja("c4"):
        tabla_kit = (
            d.groupby("KIT")
            .agg(Distritos=("UBIGEO", "nunique"), Meta_fisica=("META", "sum"),
                 Beneficiarios=("BENEFICIARIOS", "sum"))
            .reset_index().sort_values("Beneficiarios", ascending=False)
        )
        tabla_kit["% beneficiarios"] = (
            tabla_kit["Beneficiarios"] / tabla_kit["Beneficiarios"].sum() * 100
        ).round(1)
        tabla_kit.columns = ["Kit / intervención", "Distritos", "Meta física",
                             "Beneficiarios", "% beneficiarios"]
        st.markdown("**Resumen por kit**")
        st.dataframe(tabla_kit, hide_index=True)


# ============================================================
# 6. PRADERAS
# ============================================================

seccion("Praderas (regla: 5 ha = 3 beneficiarios)")
pr = d[d["KIT"] == "Praderas"]
if pr.empty:
    st.info("No hay registros de Praderas con los filtros seleccionados.")
else:
    ha = pr["META"].sum()
    benef_regla = ha / 5 * 3
    benef_excel = pr["BENEF_EXCEL"].sum()

    m = st.columns(4)
    m[0].markdown(kpi("HECTÁREAS", f"{ha:,.0f}", "meta física", "#0D9488"), unsafe_allow_html=True)
    m[1].markdown(kpi("GRUPOS DE 5 HA", f"{ha / 5:,.0f}", "ha ÷ 5", "#16A34A"), unsafe_allow_html=True)
    m[2].markdown(kpi("BENEFICIARIOS", f"{benef_regla:,.0f}", "grupos × 3", "#14532D"), unsafe_allow_html=True)
    m[3].markdown(kpi("DISTRITOS", f"{pr['UBIGEO'].nunique()}", "con praderas", "#4ADE80"), unsafe_allow_html=True)

    st.write("")
    if abs(benef_regla - benef_excel) < 0.5:
        st.success(f"✅ La regla coincide con los beneficiarios del Excel ({benef_excel:,.0f}).")
    else:
        st.warning(
            f"⚠️ La regla da {benef_regla:,.0f} pero el Excel registra "
            f"{benef_excel:,.0f} beneficiarios. Conviene revisar."
        )

    pr2 = pr.copy()
    pr2["VALOR"] = pr2["META"]
    c1, c2 = st.columns(2)
    with c1:
        tarjeta(dona(pr2, "DEPARTAMENTO", "Participación en hectáreas", top=8), "prad_dona")
    with c2:
        tarjeta(
            barras_h(pr2, "DEPARTAMENTO", "Hectáreas por departamento (barras)", top=16,
                     color=COLOR_UNICO, alto=ALTO_DONA),
            "prad_barras",
        )


# ============================================================
# 7. DETALLE (TABLA EN EXPANDER)
# ============================================================

seccion("Detalle por distrito")
with st.expander("📂 Desplegar tabla completa de detalle por distrito", expanded=False):
    ancho = data.pivot_table(
        index=["DEPARTAMENTO", "PROVINCIA", "DISTRITO"], columns="KIT",
        values="BENEFICIARIOS", aggfunc="sum", fill_value=0,
    )
    ancho["TOTAL"] = ancho.sum(axis=1)
    ancho = ancho[ancho["TOTAL"] > 0].sort_values("TOTAL", ascending=False).reset_index()
    with caja("c5"):
        st.dataframe(ancho, hide_index=True, height=380)
        st.download_button(
            "⬇️ Descargar esta tabla (CSV)",
            ancho.to_csv(index=False).encode("utf-8-sig"),
            file_name="fen_2026_beneficiarios_por_distrito.csv",
            mime="text/csv",
        )

st.markdown("---")
st.caption(
    "Dashboard Ejecutivo FEN 2026 · Fuente: datos_fen_2026.xlsx (hoja Hoja1). "
    "Tiempo, costo y riesgos no están en la fuente de datos."
)