
from pathlib import Path
import unicodedata
import pandas as pd
import streamlit as st

# ============================================================
# CONFIGURACIÓN
# ============================================================
st.set_page_config(
    page_title="Estrategia Salud en Control",
    page_icon="🏥",
    layout="wide"
)

# Esta carpeta es la que indicaste. Si guardas app.py en otra
# ubicación, cambia esta ruta por la carpeta donde están los Excel.
# En Streamlit Community Cloud, los archivos Excel estarán junto a app.py
CARPETA_DATOS = Path(__file__).resolve().parent

def buscar_archivo(patron):
    """Busca un Excel por patrón dentro de la carpeta de trabajo."""
    encontrados = list(CARPETA_DATOS.glob(patron))
    if not encontrados:
        raise FileNotFoundError(
            f"No encontré un archivo con el patrón '{patron}' en:\n{CARPETA_DATOS}"
        )
    return encontrados[0]

@st.cache_data
def cargar_datos():
    archivo_cruda = buscar_archivo("Cédula_de_la_Estrategia*.xlsx")
    archivo_meta = buscar_archivo("Meta*.xlsx")
    archivo_equivalencias = buscar_archivo("Identificador*clues*.xlsx")
    archivo_tablas = buscar_archivo("Tablas independientes*.xlsx")

    base = pd.read_excel(archivo_cruda)
    metas = pd.read_excel(archivo_meta, sheet_name=0)
    equivalencias_edo = pd.read_excel(archivo_equivalencias, sheet_name="EDO")
    tablas_excel = pd.ExcelFile(archivo_tablas)

    candidatos = pd.read_excel(tablas_excel, sheet_name="Candidatos")
    fidelizados = pd.read_excel(tablas_excel, sheet_name="Fidelizados")

    return base, metas, equivalencias_edo, candidatos, fidelizados

# ============================================================
# FUNCIONES
# ============================================================
def a_numero(serie):
    return pd.to_numeric(serie, errors="coerce").fillna(0)

def total_por_prefijo(df, prefijo):
    # Las columnas sum_*_tot ya contienen los totales por grupo de edad.
    columnas = [
        c for c in df.columns
        if c.startswith(prefijo) and c.endswith("_tot")
    ]
    if not columnas:
        raise ValueError(
            f"No encontré columnas totalizadoras para el prefijo '{prefijo}'."
        )
    return int(df[columnas].apply(a_numero).sum().sum())

def normalizar_texto(valor):
    if pd.isna(valor):
        return ""
    return str(valor).strip().upper()

# ============================================================
# CARGA Y PREPARACIÓN
# ============================================================
try:
    base, metas, equivalencias_edo, candidatos, fidelizados = cargar_datos()
except Exception as error:
    st.error("No fue posible cargar los archivos de Excel.")
    st.exception(error)
    st.stop()

# No se modifica ni se deduplica ninguna tabla independiente.
# La separación por grupo/ciclo ya viene en el archivo de Tablas independientes.

base["edo_num"] = pd.to_numeric(base["edo"], errors="coerce")
base["CLUES_num"] = pd.to_numeric(base["CLUES"], errors="coerce")
base["mes_num"] = pd.to_numeric(base["mes_actual"], errors="coerce")

equivalencias_edo["ID_num"] = pd.to_numeric(
    equivalencias_edo["ID"], errors="coerce"
)
equivalencias_edo["Entidad_limpia"] = equivalencias_edo["Entidad"].map(
    normalizar_texto
)

base = base.merge(
    equivalencias_edo[["ID_num", "Entidad_limpia"]],
    left_on="edo_num",
    right_on="ID_num",
    how="left"
)
base.rename(columns={"Entidad_limpia": "Entidad"}, inplace=True)

metas["Entidad"] = metas["Entidad federativa"].map(normalizar_texto)
metas["Meta anual"] = a_numero(metas["Meta anual"])
metas["Meta por ciclo"] = a_numero(metas["Meta por ciclo"])

# ============================================================
# INDICADORES SUPERIORES
# ============================================================
meta_anual_nacional = metas["Meta anual"].sum()
meta_ciclo_nacional = metas["Meta por ciclo"].sum()

# Conteo de CLUES únicas con al menos un registro en la base cruda.
# No se elimina ningún registro de la base: solo se cuentan claves únicas
# para construir este indicador.
clues_con_registro = base["CLUES_num"].nunique()
porcentaje_meta = (
    clues_con_registro / meta_anual_nacional * 100
    if meta_anual_nacional else 0
)

total_candidatos = total_por_prefijo(candidatos, "sum_candidat_")
total_fidelizados = total_por_prefijo(fidelizados, "sum_candidat_fideliz_")

# ============================================================
# CUADRO 1. META ANUAL Y META POR CICLO POR ENTIDAD
# ============================================================
cuadro1 = metas[["Entidad", "Meta anual", "Meta por ciclo"]].copy()
cuadro1 = cuadro1.sort_values("Entidad")

fila_total_1 = pd.DataFrame([{
    "Entidad": "TOTAL",
    "Meta anual": cuadro1["Meta anual"].sum(),
    "Meta por ciclo": cuadro1["Meta por ciclo"].sum()
}])
cuadro1 = pd.concat([cuadro1, fila_total_1], ignore_index=True)

# ============================================================
# CUADRO 2. REGISTROS POR MES DE CAPTURA
# Se toman las CLUES únicas por entidad para mes_actual 1, 4, 7 y 10.
# Es una primera aproximación de la tabla de la captura; no se deduce
# el ciclo a partir de submission_time.
# ============================================================
meses_reporte = {
    1: "Enero",
    4: "Abril",
    7: "Julio",
    10: "Octubre"
}

base_meses = base[
    base["mes_num"].isin(meses_reporte.keys())
    & base["Entidad"].ne("")
    & base["CLUES_num"].notna()
].copy()

cuadro2 = (
    base_meses
    .groupby(["Entidad", "mes_num"])["CLUES_num"]
    .nunique()
    .unstack(fill_value=0)
    .reset_index()
)

cuadro2.rename(columns=meses_reporte, inplace=True)

for nombre_mes in meses_reporte.values():
    if nombre_mes not in cuadro2.columns:
        cuadro2[nombre_mes] = 0

orden_columnas = ["Entidad", "Enero", "Abril", "Julio", "Octubre"]
cuadro2 = cuadro2[orden_columnas]
cuadro2["Total"] = cuadro2[["Enero", "Abril", "Julio", "Octubre"]].sum(axis=1)
cuadro2 = cuadro2.sort_values("Entidad")

fila_total_2 = {"Entidad": "TOTAL"}
for columna in ["Enero", "Abril", "Julio", "Octubre", "Total"]:
    fila_total_2[columna] = int(cuadro2[columna].sum())

cuadro2 = pd.concat([cuadro2, pd.DataFrame([fila_total_2])], ignore_index=True)

# ============================================================
# DISEÑO
# ============================================================
st.markdown("""
<style>
.stApp { background: #ffffff; color: #202124; }
.block-container { padding-top: 2.5rem; padding-bottom: 1rem; max-width: 1550px; }
.encabezado {
    background: linear-gradient(90deg, #75132f, #a51d49);
    color: white; padding: 20px 22px; border-radius: 4px;
    margin-bottom: 10px; min-height: 88px; box-sizing: border-box; overflow: visible;
}
.titulo { font-size: 25px; font-weight: 800; line-height: 1.35; margin: 0 0 4px 0; }
.subtitulo { font-size: 13px; font-weight: 600; }
.tarjeta {
    border: 1px solid #dedede; border-radius: 15px;
    box-shadow: 0 2px 3px #0002; padding: 12px 8px;
    text-align: center; min-height: 95px; background: #fff;
}
.tarjeta-titulo { font-size: 13px; font-weight: 800; }
.tarjeta-valor { font-size: 25px; font-weight: 600; margin-top: 5px; }
.tarjeta-detalle { font-size: 11px; color: #666; }
.seccion {
    text-align: center; font-size: 17px; font-weight: 800;
    border-top: 1px solid #aaa; border-bottom: 1px solid #aaa;
    padding: 5px; margin: 15px 0 8px;
}
.nota { font-size: 11px; color: #555; line-height: 1.45; }
.calendario {
    border: 1px solid #d2d2d2; border-radius: 13px;
    box-shadow: 0 2px 3px #0002; padding: 12px 16px;
}
.calendario p { margin: 5px 0; font-size: 12px; }
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="encabezado">
  <div class="titulo">ESTRATEGIA SALUD EN CONTROL</div>
  <div class="subtitulo">Coordinación de Unidades de Primer Nivel</div>
</div>
""", unsafe_allow_html=True)

st.caption("Vista nacional · Versión de prueba sin filtros")

col1, col2, col3 = st.columns(3)

with col1:
    st.markdown(f"""
    <div class="tarjeta">
      <div class="tarjeta-titulo">META NACIONAL</div>
      <div class="tarjeta-valor">{clues_con_registro:,.0f}
        <span style="font-size:15px;color:#666">{porcentaje_meta:.1f}%</span>
      </div>
      <div class="tarjeta-detalle">
        de {meta_anual_nacional:,.0f} unidades · Meta por ciclo {meta_ciclo_nacional:,.0f}
      </div>
    </div>
    """, unsafe_allow_html=True)

with col2:
    st.markdown(f"""
    <div class="tarjeta">
      <div class="tarjeta-titulo">CANDIDATOS</div>
      <div class="tarjeta-valor">{total_candidatos:,.0f}</div>
    </div>
    """, unsafe_allow_html=True)

with col3:
    st.markdown(f"""
    <div class="tarjeta">
      <div class="tarjeta-titulo">FIDELIZADOS</div>
      <div class="tarjeta-valor">{total_fidelizados:,.0f}</div>
    </div>
    """, unsafe_allow_html=True)

st.markdown(
    '<div class="seccion">AVANCE ESTATAL DE IMPLEMENTACIÓN DE LA ESTRATEGIA</div>',
    unsafe_allow_html=True
)
st.markdown("""
<div style="text-align:center;font-size:11px;margin-bottom:9px">
  <b style="color:#08745b">▲ Aumento</b> (desfase y/o mayor participación)
  &nbsp;|&nbsp; <b style="color:#b07b00">▬ Consistente</b> (registro adecuado)
  &nbsp;|&nbsp; <b style="color:#a51d49">▼ Disminución</b> (alerta, registro no oportuno)
</div>
""", unsafe_allow_html=True)

izquierda, derecha = st.columns([1, 1.45], gap="large")

with izquierda:
    st.markdown(
        "**Cuadro 1. Meta anual y por ciclo de Unidades de primer nivel fijas que deben participar en la estrategia**"
    )
    st.dataframe(
        cuadro1.style.format({
            "Meta anual": "{:,.0f}",
            "Meta por ciclo": "{:,.0f}"
        }),
        use_container_width=True,
        hide_index=True,
        height=490
    )

with derecha:
    st.markdown(
        "**Cuadro 2. Unidades de primer nivel fijas con registro de la estrategia según el mes de captura**"
    )
    st.dataframe(
        cuadro2.style.format({
            "Enero": "{:,.0f}",
            "Abril": "{:,.0f}",
            "Julio": "{:,.0f}",
            "Octubre": "{:,.0f}",
            "Total": "{:,.0f}"
        }),
        use_container_width=True,
        hide_index=True,
        height=490
    )

st.markdown("""
<div class="nota">
<b>Fuente:</b> Portal de Gestión de Salud Poblacional (PGSP).<br>
<b>Meta anual:</b> corresponde a la participación del 50% de las unidades de primer nivel de la entidad.<br>
<b>Meta por ciclo:</b> corresponde a la participación del 25% de la meta anual.<br>
<b>Nota de prueba:</b> el cuadro de registro cuenta CLUES únicas para los meses 1, 4, 7 y 10 de la variable
<code>mes_actual</code>. Se debe validar esta regla contra el cálculo definitivo del visualizador original.
</div>
""", unsafe_allow_html=True)

st.markdown(
    '<div class="seccion">CALENDARIO DE REGISTRO</div>',
    unsafe_allow_html=True
)
st.markdown("""
<div class="calendario">
  <b>MESES DE CAPTURA</b>
  <p>• <b>Enero:</b> Ingreso de nuevos candidatos (1.er ciclo) y fidelizados del 2.º ciclo del año previo.</p>
  <p>• <b>Abril:</b> Seguimiento de candidatos (1.er ciclo).</p>
  <p>• <b>Julio:</b> Cierre de fidelizados (1.er ciclo) e ingreso de nuevos candidatos (2.º ciclo).</p>
  <p>• <b>Octubre:</b> Seguimiento de candidatos (2.º ciclo).</p>
</div>
""", unsafe_allow_html=True)

with st.expander("Comprobación de cifras y archivos"):
    st.write("Archivos leídos desde:", str(CARPETA_DATOS))
    st.write("Registros de la base cruda:", f"{len(base):,}")
    st.write("CLUES únicas en la base cruda:", f"{clues_con_registro:,}")
    st.write("Meta anual nacional:", f"{meta_anual_nacional:,.2f}")
    st.write("Meta por ciclo nacional:", f"{meta_ciclo_nacional:,.2f}")
    st.write("Total de candidatos:", f"{total_candidatos:,}")
    st.write("Total de fidelizados:", f"{total_fidelizados:,}")
    st.write("La aplicación no aplica filtros ni deduplica las tablas independientes.")
