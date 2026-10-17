import unicodedata
from pathlib import Path

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

# ============================================================
# CONFIGURACIÓN
# ============================================================

st.set_page_config(page_title="Dashboard Ejecutivo FEN 2026 - UAT", page_icon="🌱", layout="wide")

CARPETA = Path(__file__).parent
ARCHIVO = CARPETA / "datos_fen_2026.xlsx"
HOJA = "Hoja1"
ARCHIVO_COORD = CARPETA / "coordenadas_distritos.xlsx"       # lat/lon por distrito
ARCHIVOS_AVANCE = [CARPETA / "avance_fen_2026.xlsx", CARPETA / "avance_fen_2026.csv"]

NIVELES = ["DEPARTAMENTO", "PROVINCIA", "DISTRITO"]
ETIQ_NIVEL = {"DEPARTAMENTO": "departamento", "PROVINCIA": "provincia", "DISTRITO": "distrito"}

NOMBRES_CORTOS = [
    ("ALIMENTO SUPLEMENTARIO", "Alimento suplementario"),
    ("INSUMOS AGRICOLAS", "Insumos agrícolas"),
    ("VETERINARIO", "Veterinario"),
    ("RIEGO", "Riego"),
    ("SEMILLAS", "Semillas"),
    ("PRADERAS", "Praderas"),
    ("ABREVADERO", "Abrevadero"),
]
COLORES_KIT = {
    "Alimento suplementario": "#0284C7", "Insumos agrícolas": "#16A34A",
    "Veterinario": "#F97316", "Riego": "#06B6D4", "Semillas": "#84CC16",
    "Praderas": "#14B8A6", "Abrevadero": "#EAB308",
}
COLOR_PROG = "#0284C7"       # programado (sin avance)
COLOR_PROG_SUAVE = "#BAE6FD"  # programado (con avance encima)
COLOR_AVANCE = "#16A34A"
PLANTILLA = "plotly_white"
ALTO = 520


def caja(nombre):
    try:
        return st.container(border=True, key=f"card_{nombre}")
    except TypeError:
        return st.container(border=True)


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
      padding: 16px 18px; min-height: 124px; border-top: 5px solid var(--c, #16A34A);}
.kpi-titulo {color: #475569; font-size: 12px; font-weight: 700; letter-spacing: .05em;}
.kpi-valor {color: #14532D; font-size: 30px; font-weight: 800; margin-top: 4px; line-height: 1.1;}
.kpi-sub {color: #64748B; font-size: 12px; margin-top: 4px;}
.lider {color: #14532D; font-size: 17px; font-weight: 800; margin-top: 6px; line-height: 1.2; min-height: 40px;}
.kpi.alta {height: 215px; position: relative; box-sizing: border-box;}
.kpi.alta .kpi-titulo {min-height: 30px;}
.kpi.alta .pbar {position: absolute; left: 18px; right: 18px; bottom: 44px; margin: 0;}
.kpi.alta .kpi-sub {position: absolute; left: 18px; right: 18px; bottom: 14px; margin: 0;}
.pbar {background: #E2E8F0; border-radius: 6px; height: 8px; margin-top: 8px;}
.pbar > div {background: var(--c, #16A34A); height: 8px; border-radius: 6px;}
.seccion {font-size: 19px; font-weight: 800; color: #14532D;
          border-left: 6px solid #16A34A; padding-left: 10px; margin: 26px 0 10px 0;}
.ruta {font-size: 17px; font-weight: 700; color: #14532D; padding-top: 6px;}
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
# CARGA Y TRANSFORMACIÓN
# ============================================================

def norm(s):
    t = unicodedata.normalize("NFD", str(s).upper().strip())
    return " ".join("".join(c for c in t if unicodedata.category(c) != "Mn").split())


def nombre_corto(nombre_kit):
    n = norm(nombre_kit)
    for clave, corto in NOMBRES_CORTOS:
        if clave in n:
            return corto
    return str(nombre_kit).strip().title()


def leer_avance():
    """Lee el reporte de avance (opcional). Devuelve (df | None, mensaje)."""
    for ruta in ARCHIVOS_AVANCE:
        if not ruta.exists():
            continue
        try:
            if ruta.suffix == ".xlsx":
                a = pd.read_excel(ruta, dtype=str)
            else:
                a = pd.read_csv(ruta, dtype=str, sep=None, engine="python", encoding="utf-8-sig")
        except Exception as e:
            return None, f"No se pudo leer {ruta.name}: {e}"
        a.columns = [str(c).strip().upper() for c in a.columns]
        faltan = [c for c in ("DEPARTAMENTO", "PROVINCIA", "DISTRITO", "KIT") if c not in a.columns]
        if faltan or not ({"AVANCE_META", "AVANCE_BENEFICIARIOS"} & set(a.columns)):
            return None, (f"{ruta.name} debe tener las columnas DEPARTAMENTO, PROVINCIA, DISTRITO, KIT y "
                          "AVANCE_META / AVANCE_BENEFICIARIOS (opcional FECHA_REPORTE).")
        return a, ""
    return None, ""


@st.cache_data
def cargar_datos(firma):
    crudo = pd.read_excel(ARCHIVO, sheet_name=HOJA, header=None)

    fila_enc = next((i for i in range(len(crudo))
                     if str(crudo.iat[i, 1]).strip().upper() == "DEPARTAMENTO"), None)
    if fila_enc is None:
        raise ValueError("No se encontró la fila con 'DEPARTAMENTO' en la hoja.")
    fila_kits = fila_enc - 1

    fecha_prog = "16.09.26"
    try:
        txt = str(crudo.iat[0, 19])
        if "ctualizado" in txt:
            fecha_prog = txt.lower().replace("actualizado", "").strip()
    except Exception:
        pass

    datos = crudo.iloc[fila_enc + 1:].copy()
    datos = datos[datos[1].notna()]
    datos = datos[~datos[0].astype(str).str.upper().str.contains("TOTAL", na=False)]

    partes = []
    for col in range(4, crudo.shape[1]):
        nombre = crudo.iat[fila_kits, col]
        if pd.isna(nombre):
            continue
        partes.append(pd.DataFrame({
            "DEPARTAMENTO": datos[1].astype(str).str.strip(),
            "PROVINCIA": datos[2].astype(str).str.strip(),
            "DISTRITO": datos[3].astype(str).str.strip(),
            "KIT": nombre_corto(nombre),
            "META": pd.to_numeric(datos[col], errors="coerce").fillna(0),
            "BENEF_EXCEL": pd.to_numeric(datos[col + 1], errors="coerce").fillna(0),
        }))
    b = pd.concat(partes, ignore_index=True)

    # Regla de Praderas: 5 ha = 3 beneficiarios
    b["PROG"] = b["BENEF_EXCEL"]
    es_prad = b["KIT"] == "Praderas"
    b.loc[es_prad, "PROG"] = b.loc[es_prad, "META"] / 5 * 3
    b["ATENDIDO"] = (b["META"] > 0) | (b["PROG"] > 0)
    b["UBIGEO"] = b["DEPARTAMENTO"] + " / " + b["PROVINCIA"] + " / " + b["DISTRITO"]
    b["K_DEP"], b["K_PROV"], b["K_DIST"] = (b[c].map(norm) for c in NIVELES)

    # Coordenadas
    hay_coord = ARCHIVO_COORD.exists()
    b["LAT"], b["LON"], b["APROX"] = np.nan, np.nan, False
    if hay_coord:
        c = pd.read_excel(ARCHIVO_COORD)
        c["K_DEP"], c["K_PROV"], c["K_DIST"] = (c[x].map(norm) for x in NIVELES)
        c = c.drop_duplicates(["K_DEP", "K_PROV", "K_DIST"])
        b = b.drop(columns=["LAT", "LON", "APROX"]).merge(
            c[["K_DEP", "K_PROV", "K_DIST", "LAT", "LON", "APROXIMADA"]]
            .rename(columns={"APROXIMADA": "APROX"}),
            on=["K_DEP", "K_PROV", "K_DIST"], how="left")
        b["APROX"] = b["APROX"].fillna(False).astype(bool)

    # Avance (opcional)
    b["AV_META"], b["AV"] = 0.0, 0.0
    hay_av, fecha_av, msg_av, av_sin = False, "", "", pd.DataFrame()
    a, msg_av = leer_avance()
    if a is not None:
        hay_av = True
        for col in ("AVANCE_META", "AVANCE_BENEFICIARIOS"):
            a[col] = pd.to_numeric(a[col], errors="coerce").fillna(0) if col in a.columns else 0.0
        a["K_DEP"], a["K_PROV"], a["K_DIST"] = (a[c].map(norm) for c in NIVELES)
        a["KIT"] = a["KIT"].map(nombre_corto)
        if "FECHA_REPORTE" in a.columns:
            f = pd.to_datetime(a["FECHA_REPORTE"], dayfirst=True, errors="coerce").max()
            fecha_av = "" if pd.isna(f) else f.strftime("%d/%m/%Y")
        g = a.groupby(["K_DEP", "K_PROV", "K_DIST", "KIT"], as_index=False)[
            ["AVANCE_META", "AVANCE_BENEFICIARIOS"]].sum()
        claves = b[["K_DEP", "K_PROV", "K_DIST", "KIT"]].drop_duplicates()
        chk = g.merge(claves, how="left", indicator=True)
        av_sin = chk[chk["_merge"] == "left_only"].drop(columns="_merge")
        b = b.drop(columns=["AV_META", "AV"]).merge(g, on=["K_DEP", "K_PROV", "K_DIST", "KIT"], how="left")
        b["AV_META"] = b["AVANCE_META"].fillna(0)
        b["AV"] = b["AVANCE_BENEFICIARIOS"].fillna(0)
        es_prad = b["KIT"] == "Praderas"
        b.loc[es_prad, "AV"] = b.loc[es_prad, "AV_META"] / 5 * 3
        b = b.drop(columns=["AVANCE_META", "AVANCE_BENEFICIARIOS"])

    b["PROVINCIA_LBL"] = b["PROVINCIA"] + " (" + b["DEPARTAMENTO"].str.title() + ")"
    b["DISTRITO_LBL"] = b["DISTRITO"] + " (" + b["PROVINCIA"] + ")"
    return {"base": b, "fecha_prog": fecha_prog, "fecha_av": fecha_av, "hay_av": hay_av,
            "msg_av": msg_av, "av_sin": av_sin, "hay_coord": hay_coord}


def _mt(p):
    return p.stat().st_mtime if p.exists() else 0


try:
    R = cargar_datos((_mt(ARCHIVO), _mt(ARCHIVO_COORD), *[_mt(p) for p in ARCHIVOS_AVANCE]))
except Exception as e:
    st.error(f"No se pudo leer el archivo '{ARCHIVO.name}': {e}")
    st.stop()

base, HAY_AV, HAY_COORD = R["base"], R["hay_av"], R["hay_coord"]
d = base[base["ATENDIDO"]].copy()
FECHA_AV = R["fecha_av"] or "fecha no indicada"


# ============================================================
# FUNCIONES AUXILIARES
# ============================================================

def seccion(t):
    st.markdown(f'<div class="seccion">{t}</div>', unsafe_allow_html=True)


def pct(av, prog):
    return (av / prog * 100) if prog else 0.0


def kpi(titulo, valor, sub, color, barra=None):
    b = f'<div class="pbar"><div style="width:{min(barra, 100):.0f}%"></div></div>' if barra is not None else ""
    return (f'<div class="kpi" style="--c:{color}"><div class="kpi-titulo">{titulo}</div>'
            f'<div class="kpi-valor">{valor}</div>{b}<div class="kpi-sub">{sub}</div></div>')


def tarjeta_lider(titulo, nombre, valor_txt, sub, color, barra=None):
    b = f'<div class="pbar"><div style="width:{min(barra, 100):.0f}%"></div></div>' if barra is not None else ""
    return (f'<div class="kpi alta" style="--c:{color}"><div class="kpi-titulo">{titulo}</div>'
            f'<div class="lider">{nombre}</div><div class="kpi-valor">{valor_txt}</div>{b}'
            f'<div class="kpi-sub">{sub}</div></div>')


def resumen(df, col):
    g = (df.groupby(col).agg(PROG=("PROG", "sum"), AV=("AV", "sum"),
                             META=("META", "sum"), DIST=("UBIGEO", "nunique"))
         .reset_index().rename(columns={col: "NOMBRE"}))
    if col == "KIT":
        g["LAT"], g["LON"] = df["LAT"].mean(), df["LON"].mean()
    else:
        co = (df.drop_duplicates("UBIGEO").groupby(col)[["LAT", "LON"]].mean()
              .reset_index().rename(columns={col: "NOMBRE"}))
        g = g.merge(co, on="NOMBRE", how="left")
    g["PCT"] = np.where(g["PROG"] > 0, g["AV"] / g["PROG"] * 100, 0)
    return g.sort_values("PROG", ascending=False).reset_index(drop=True)


def mapa(res, key, seleccionable=True):
    r = res.dropna(subset=["LAT", "LON"]).copy()
    if r.empty:
        st.info("No hay coordenadas disponibles para esta selección.")
        return None
    r["TAM"] = r["PROG"].clip(lower=1)
    span = max(np.ptp(r["LAT"]), np.ptp(r["LON"]) * 0.7, 0.12) * 1.35
    zoom = float(np.clip(np.log2(675 / span), 4, 10.5))
    kw = dict(lat="LAT", lon="LON", size="TAM", size_max=38 if len(r) > 1 else 24,
              custom_data=["NOMBRE"], hover_name="NOMBRE", zoom=zoom,
              center=dict(lat=float(r["LAT"].mean()), lon=float(r["LON"].mean())),
              hover_data={"PROG": ":,.0f", "AV": ":,.0f", "PCT": ":.0f",
                          "TAM": False, "LAT": False, "LON": False},
              labels={"PROG": "Beneficiarios", "AV": "Avance", "PCT": "% avance"})
    if HAY_AV:
        kw.update(color="PCT", range_color=[0, 100],
                  color_continuous_scale=[[0, "#FDE68A"], [0.5, "#4ADE80"], [1, "#14532D"]])
    else:
        kw.update(color_discrete_sequence=[COLOR_PROG])
    if hasattr(px, "scatter_map"):
        fig = px.scatter_map(r, map_style="carto-positron", **kw)
    else:
        fig = px.scatter_mapbox(r, mapbox_style="carto-positron", **kw)
    fig.update_traces(marker=dict(opacity=0.78))
    if len(r) <= 16:
        try:
            fig.update_traces(text=r["NOMBRE"].tolist(), mode="markers+text", textposition="top right")
        except Exception:
            pass
    fig.update_layout(height=ALTO, margin=dict(t=10, b=0, l=0, r=0),
                      coloraxis_showscale=HAY_AV, paper_bgcolor="#FFFFFF",
                      coloraxis_colorbar=dict(title="% avance", thickness=12, len=0.5))
    return chart(fig, key, seleccionable)


def ranking(res, titulo, key, seleccionable=True, colores=None):
    r = res.sort_values("PROG", ascending=True)
    tot = r["PROG"].sum()
    if HAY_AV:
        txt = [f"{a:,.0f} / {p:,.0f} ({q:.0f}%)" for a, p, q in zip(r["AV"], r["PROG"], r["PCT"])]
    else:
        txt = [f"{p:,.0f} · {p / tot * 100:.0f}%" if tot else f"{p:,.0f}" for p in r["PROG"]]
    fig = go.Figure()
    marcador = (COLOR_PROG_SUAVE if HAY_AV
                else ([colores.get(n, COLOR_PROG) for n in r["NOMBRE"]] if colores else COLOR_PROG))
    fig.add_bar(y=r["NOMBRE"], x=r["PROG"], orientation="h", name="Total",
                marker_color=marcador, customdata=r["NOMBRE"], text=txt, textposition="outside",
                cliponaxis=False, hovertemplate="<b>%{y}</b><br>Beneficiarios: %{x:,.0f}<extra></extra>")
    if HAY_AV:
        fig.add_bar(y=r["NOMBRE"], x=r["AV"], orientation="h", name="Avance",
                    marker_color=COLOR_AVANCE, customdata=r["NOMBRE"],
                    hovertemplate="<b>%{y}</b><br>Avance: %{x:,.0f}<extra></extra>")
    fig.update_layout(
        template=PLANTILLA, barmode="overlay", height=ALTO, title=dict(text=titulo, font=dict(size=16, color="#14532D")),
        margin=dict(t=50, b=10, l=10, r=110), xaxis_title="Beneficiarios", yaxis_title=None,
        showlegend=HAY_AV, legend=dict(orientation="h", y=1.07, x=1, xanchor="right"),
        paper_bgcolor="#FFFFFF", plot_bgcolor="#FFFFFF")
    fig.update_yaxes(categoryorder="array", categoryarray=r["NOMBRE"].tolist())
    return chart(fig, key, seleccionable)


def chart(fig, key, seleccionable):
    """Dibuja el gráfico; si es seleccionable devuelve el nombre del elemento clicado."""
    if not seleccionable:
        st.plotly_chart(fig, key=key, theme=None)
        return None
    try:
        ev = st.plotly_chart(fig, key=key, theme=None, on_select="rerun", selection_mode="points")
    except TypeError:       # Streamlit antiguo: sin clic, se navega con el selector "Ir a"
        st.plotly_chart(fig, key=key, theme=None)
        return None
    try:
        pts = ev.selection.points
    except Exception:
        pts = []
    if not pts:
        return None
    cd = pts[0].get("customdata")
    if isinstance(cd, (list, tuple)):
        cd = cd[0] if cd else None
    return cd or pts[0].get("y")


# ============================================================
# ENCABEZADO
# ============================================================

sub_av = f"Avance reportado al {FECHA_AV}" if HAY_AV else "Resultados finales"
st.markdown(
    f"""<div class="banner"><div class="t">🌱 Dashboard Ejecutivo FEN 2026</div>
    <div class="s">Metas físicas y beneficiarios por intervención · {sub_av} · Datos actualizados al
    {R['fecha_prog']}</div></div>""",
    unsafe_allow_html=True,
)

# ============================================================
# 1. VISIÓN GENERAL (SIEMPRE TODO EL PROYECTO)
# ============================================================

PROG_T, AV_T = d["PROG"].sum(), d["AV"].sum()
seccion("Alcance del proyecto")
k = st.columns(5)
k[0].markdown(kpi("DEPARTAMENTOS", d["DEPARTAMENTO"].nunique(), "en el alcance del proyecto", "#14532D"), unsafe_allow_html=True)
k[1].markdown(kpi("PROVINCIAS", d.groupby(["DEPARTAMENTO", "PROVINCIA"]).ngroups, "en el alcance del proyecto", "#16A34A"), unsafe_allow_html=True)
dist_av = d.loc[d["AV"] > 0, "UBIGEO"].nunique()
k[2].markdown(kpi("DISTRITOS", f"{d['UBIGEO'].nunique():,}",
                  f"en el alcance · {dist_av:,} con avance" if HAY_AV else "en el alcance del proyecto",
                  "#4ADE80"), unsafe_allow_html=True)
k[3].markdown(kpi("INTERVENCIONES", d["KIT"].nunique(), "tipos de kit del proyecto", "#A3E635"), unsafe_allow_html=True)
k[4].markdown(kpi("BENEFICIARIOS", f"{PROG_T:,.0f}", "total del proyecto", "#0D9488"), unsafe_allow_html=True)
st.caption("ℹ️ Beneficiarios = unidades de atención de cada intervención (no necesariamente personas) y se "
           "suman entre intervenciones. Definición por validar con todos los involucrados.")

# ---- Lo más destacado
seccion("⭐ Lo más destacado")
h = st.columns(5)
if HAY_AV:
    p = pct(AV_T, PROG_T)
    h[0].markdown(tarjeta_lider("AVANCE GENERAL", "Todo el proyecto", f"{p:.1f}%",
                                f"{AV_T:,.0f} de {PROG_T:,.0f} · al {FECHA_AV}", "#14532D", barra=p),
                  unsafe_allow_html=True)
else:
    h[0].markdown(tarjeta_lider("TOTAL DEL PROYECTO", "Todo el proyecto", f"{PROG_T:,.0f}",
                                f"beneficiarios en {d['UBIGEO'].nunique():,} distritos", "#14532D"),
                  unsafe_allow_html=True)

COL_LID = [("DEPARTAMENTO", "DEPARTAMENTO", "#16A34A"), ("PROVINCIA_LBL", "PROVINCIA", "#4ADE80"),
           ("DISTRITO_LBL", "DISTRITO", "#0D9488"), ("KIT", "INTERVENCIÓN", "#A3E635")]
for i, (col, etiqueta, color) in enumerate(COL_LID, start=1):
    g = d.groupby(col).agg(PROG=("PROG", "sum"), AV=("AV", "sum"))
    g = g[g["PROG"] > 0]
    if HAY_AV:
        g["PCT"] = g["AV"] / g["PROG"] * 100
        top = g.sort_values(["PCT", "PROG"], ascending=False).iloc[0]
        nombre = g.sort_values(["PCT", "PROG"], ascending=False).index[0]
        h[i].markdown(tarjeta_lider(f"{etiqueta} · MAYOR AVANCE", nombre.title() if col == "DEPARTAMENTO" else nombre,
                                    f"{top['PCT']:.0f}%",
                                    f"{top['AV']:,.0f} de {top['PROG']:,.0f} · al {FECHA_AV}", color, barra=top["PCT"]),
                      unsafe_allow_html=True)
    else:
        g = g.sort_values("PROG", ascending=False)
        nombre, top = g.index[0], g.iloc[0]
        share = top["PROG"] / g["PROG"].sum() * 100
        h[i].markdown(tarjeta_lider(f"{etiqueta} · MÁS BENEFICIARIOS", nombre.title() if col == "DEPARTAMENTO" else nombre,
                                    f"{share:.0f}%", f"{top['PROG']:,.0f} de {PROG_T:,.0f} beneficiarios del total",
                                    color, barra=share), unsafe_allow_html=True)


# ============================================================
# 2. EXPLORACIÓN: CLIC PARA PROFUNDIZAR
# ============================================================

seccion("Explora el detalle · haz clic en un círculo o barra para profundizar")

ruta = st.session_state.setdefault("ruta", [])


def subir():
    st.session_state["ruta"] = st.session_state["ruta"][:-1]


def inicio():
    st.session_state["ruta"] = []


def entrar(nombre):
    if len(st.session_state["ruta"]) < 3:
        st.session_state["ruta"] = st.session_state["ruta"] + [nombre]
        st.rerun()


kits_disp = ["Todas las intervenciones"] + sorted(d["KIT"].unique())
c_ruta, c_kit, c_ir, c_b1, c_b2 = st.columns([4, 2.3, 2.3, 1.1, 1.1])
c_ruta.markdown('<div class="ruta">🇵🇪 ' + " › ".join(["Perú"] + [r.title() if i == 0 else r for i, r in enumerate(ruta)])
                + "</div>", unsafe_allow_html=True)
kit_sel = c_kit.selectbox("Intervención", kits_disp, key="f_kit", label_visibility="collapsed")
c_b1.button("⬆️ Subir", on_click=subir, disabled=not ruta, use_container_width=True)
c_b2.button("🏠 Perú", on_click=inicio, disabled=not ruta, use_container_width=True)

sc = d.copy()
for i, v in enumerate(ruta):
    sc = sc[sc[NIVELES[i]] == v]
if kit_sel != kits_disp[0]:
    sc = sc[sc["KIT"] == kit_sel]
if sc.empty:
    st.warning("Esa combinación no tiene datos (la intervención elegida no llega a este lugar).")
    st.stop()

nivel = len(ruta)
if nivel < 3:
    hijos = resumen(sc, NIVELES[nivel])
    sel_ir = c_ir.selectbox("Ir a…", hijos["NOMBRE"].tolist(), index=None,
                            placeholder=f"Ir a {ETIQ_NIVEL[NIVELES[nivel]]}…",
                            key="ir_" + "|".join(ruta), label_visibility="collapsed")
    if sel_ir:
        entrar(sel_ir)
else:
    c_ir.empty()
    hijos = resumen(sc, "KIT")

# Resumen de la selección
s = st.columns(4)
p_sc = pct(sc["AV"].sum(), sc["PROG"].sum())
p_tot = sc["PROG"].sum() / PROG_T * 100
s[0].markdown(kpi("BENEFICIARIOS", f"{sc['PROG'].sum():,.0f}", "en esta selección", "#0D9488"), unsafe_allow_html=True)
if HAY_AV:
    s[1].markdown(kpi("AVANCE", f"{sc['AV'].sum():,.0f}", f"{p_sc:.1f}% de lo proyectado · al {FECHA_AV}",
                      "#16A34A", barra=p_sc), unsafe_allow_html=True)
else:
    s[1].markdown(kpi("PESO EN EL PROYECTO", f"{p_tot:.1f}%", "del total de beneficiarios", "#16A34A", barra=p_tot),
                  unsafe_allow_html=True)
s[2].markdown(kpi("DISTRITOS", f"{sc['UBIGEO'].nunique():,}", "en esta selección", "#4ADE80"), unsafe_allow_html=True)
s[3].markdown(kpi("INTERVENCIONES", sc["KIT"].nunique(), "en esta selección", "#A3E635"), unsafe_allow_html=True)
st.write("")

clave = "|".join(ruta + [kit_sel])
m1, m2 = st.columns([3, 2])
nombre_nivel = ["departamentos", "provincias", "distritos", "intervenciones"][nivel]
with m1:
    with caja("mapa"):
        if HAY_COORD:
            res_mapa = hijos if nivel < 3 else resumen(sc, "DISTRITO")
            clic_mapa = mapa(res_mapa, "mapa_" + clave, seleccionable=nivel < 3)
        else:
            clic_mapa = None
            st.info("Falta el archivo coordenadas_distritos.xlsx junto a este programa para mostrar el mapa.")
with m2:
    with caja("ranking"):
        clic_barra = ranking(hijos, f"Beneficiarios por {nombre_nivel[:-1] if nivel < 3 else 'intervención'}"
                                    f"{' (clic para entrar)' if nivel < 3 else ''}",
                             "rank_" + clave, seleccionable=nivel < 3,
                             colores=COLORES_KIT if nivel == 3 else None)
if nivel < 3 and (clic_mapa or clic_barra):
    entrar(clic_mapa or clic_barra)

if nivel < 3:
    with caja("kits_sel"):
        por_kit = resumen(sc, "KIT")
        ranking(por_kit, "Intervenciones en esta selección", "kits_" + clave,
                seleccionable=False, colores=COLORES_KIT)
else:
    with caja("tabla_kit"):
        t = sc.groupby("KIT").agg(Meta=("META", "sum"), Proyectado=("PROG", "sum"), Avance=("AV", "sum")).reset_index()
        t["% avance"] = np.where(t["Proyectado"] > 0, (t["Avance"] / t["Proyectado"] * 100).round(1), 0)
        if not HAY_AV:
            t = t.drop(columns=["Avance", "% avance"])
        t = t.rename(columns={"KIT": "Intervención", "Meta": "Meta física"})
        st.markdown("**Detalle por intervención del distrito**")
        st.dataframe(t.sort_values("Proyectado", ascending=False), hide_index=True)

# ============================================================
# 3. TABLA Y VALIDACIONES
# ============================================================

with st.expander("📂 Tabla de detalle por distrito (selección actual)", expanded=False):
    ancho = sc.pivot_table(index=NIVELES, columns="KIT", values="PROG", aggfunc="sum", fill_value=0)
    ancho["TOTAL PROYECTADO"] = ancho.sum(axis=1)
    if HAY_AV:
        ancho["AVANCE"] = sc.groupby(NIVELES)["AV"].sum()
        ancho["% AVANCE"] = (ancho["AVANCE"] / ancho["TOTAL PROYECTADO"] * 100).round(1)
    ancho = ancho.sort_values("TOTAL PROYECTADO", ascending=False).reset_index()
    st.dataframe(ancho, hide_index=True, height=380)
    st.download_button("⬇️ Descargar (CSV)", ancho.to_csv(index=False).encode("utf-8-sig"),
                       file_name="fen_2026_detalle_distritos.csv", mime="text/csv")

with st.expander("🔎 Notas metodológicas y validaciones de datos", expanded=False):
    st.markdown(
        "- **Beneficiarios:** unidades de atención de cada intervención (no necesariamente personas); se suman entre intervenciones (definición por validar).\n"
        "- **Praderas:** se aplica la regla 5 ha = 3 beneficiarios sobre la meta física.\n"
        "- Las cifras del Excel se toman como resultados finales. El avance es opcional (archivo aparte)."
    )
    pr = d[d["KIT"] == "Praderas"]
    if not pr.empty:
        reg, exc = pr["META"].sum() / 5 * 3, pr["BENEF_EXCEL"].sum()
        (st.success if abs(reg - exc) < 0.5 else st.warning)(
            f"Praderas: {pr['META'].sum():,.0f} ha ÷ 5 × 3 = {reg:,.0f} beneficiarios · Excel registra {exc:,.0f}.")
    if HAY_COORD:
        ap = d.loc[d["APROX"], ["DEPARTAMENTO", "PROVINCIA", "DISTRITO"]].drop_duplicates()
        st.markdown(f"**Mapa:** {len(ap)} distritos se ubican de forma aproximada (centro de la provincia o "
                    "nombre similar en la base de referencia).")
        if len(ap):
            st.dataframe(ap, hide_index=True)
    if HAY_AV and len(R["av_sin"]):
        st.warning(f"{len(R['av_sin'])} filas del archivo de avance no coinciden con la programación "
                   "(revisar nombres de distrito/intervención):")
        st.dataframe(R["av_sin"], hide_index=True)

st.markdown("---")
st.caption("Dashboard Ejecutivo FEN 2026 · Fuente: datos_fen_2026.xlsx (hoja Hoja1). "
           "Tiempo, costo y riesgos no están en la fuente de datos.")
