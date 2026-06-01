from pathlib import Path
import re

import altair as alt
import numpy as np
import pandas as pd
import streamlit as st


CARPETA_APP = Path(__file__).resolve().parent
RUTA_LOGO = CARPETA_APP / "Logo_PROA_HRS.png"

ARCHIVOS_SERVICIOS = {
    "URGENCIAS ADULTOS": "HRS_urgencias_adultos.xlsx",
    "URGENCIAS PEDIATRIA": "HRS_urgencias_pediatria.xlsx",
    "HOSPITALIZACION ADULTOS": "HRS_hospitalizacion_adultos.xlsx",
    "HOSPITALIZACION PEDIATRIA": "HRS_hospitalizacion_pediatria.xlsx",
    "SALA DE PARTOS": "HRS_sala_partos.xlsx",
    "UCI ADULTOS": "HRS_UCI_adultos.xlsx",
    "UCI NEONATAL": "HRS_UCI_neonatal.xlsx",
}

ANTIBIOTICOS = {
    "AMK": "Amikacina",
    "PEN": "Penicilina",
    "AMP": "Ampicilina",
    "SAM": "Ampi/Sulbactam",
    "CZO": "Cefazolina",
    "CXM": "Cefuroxima",
    "CRO": "Ceftriaxona",
    "FEP": "Cefepime",
    "TZP": "Pip/Tazo",
    "ETP": "Ertapenem",
    "MEM": "Meropenem",
    "CIP": "Ciprofloxacino",
    "SXT": "TMP/SMX",
    "OXA": "Oxacilina",
    "VAN": "Vancomicina",
    "LNZ": "Linezolid",
    "CLI": "Clindamicina",
    "TET": "Tetraciclina",
    "LVX": "Levofloxacina",
    "CAZ": "Ceftazidima",
    "FOS": "Fosfomicina",
    "FLU": "Fluconazol",
    "VOR": "Voriconazol",
    "CAS": "Caspofungina",
    "AMB": "Anfotericina B",
}

TIPOS_MUESTRA = {
    "Todas": [],
    "Urocultivos": ["orina"],
    "Hemocultivos": ["sang"],
    "Muestras respiratorias": [
        "aspirado traqueal", "bronquial", "empiema", "empiema pleural",
        "esputo", "esputo inducido", "faringe", "mini-bal", "mini bal",
        "nasofaringe", "pulmones", "respiratorio", "respiratorio alto",
        "respiratorio bajo", "traqueal",
    ],
    "Muestras abdominales": [
        "abdomen", "abdominal", "absceso abdominal", "duodeno", "fistula",
        "fístula", "higado", "hígado", "liquido abdominal",
        "líquido abdominal", "liquido de dialisis", "líquido de diálisis",
        "liquido gastrico", "líquido gástrico", "pancreas", "páncreas",
        "ovario", "pelvis", "receso rectouterino", "sitio de gastrostomia",
        "sitio de gastrostomía", "vesicula biliar", "vesícula biliar",
    ],
}


def normalizar_texto(texto):
    texto = str(texto).lower().strip()
    reemplazos = {
        "á": "a", "é": "e", "í": "i", "ó": "o", "ú": "u", "ñ": "n",
        "Ã¡": "a", "Ã©": "e", "Ã­": "i", "Ã³": "o", "Ãº": "u", "Ã±": "n",
        "ÃƒÂ¡": "a", "ÃƒÂ©": "e", "ÃƒÂ­": "i", "ÃƒÂ³": "o", "ÃƒÂº": "u", "ÃƒÂ±": "n",
    }
    for origen, destino in reemplazos.items():
        texto = texto.replace(origen, destino)
    return re.sub(r"[^a-z0-9]", "", texto)


def normalizar_microorganismo(valor):
    germen = str(valor).strip()
    germen_normalizado = normalizar_texto(germen)

    if "escherichiacoli" in germen_normalizado:
        return "Escherichia coli"
    if "klebsiellapneumoniae" in germen_normalizado:
        return "Klebsiella pneumoniae"
    if "klebsiellaoxytoca" in germen_normalizado:
        return "Klebsiella oxytoca"
    if "proteusmirabilis" in germen_normalizado:
        return "Proteus mirabilis"

    return germen


def buscar_columna(df, opciones):
    for columna in opciones:
        if columna in df.columns:
            return columna
    raise ValueError(f"No se encontro ninguna de estas columnas: {opciones}")


def resolver_archivo_excel(nombre_archivo):
    ruta = CARPETA_APP / nombre_archivo
    if ruta.exists():
        return ruta

    ruta_sin_extension = CARPETA_APP / Path(nombre_archivo).stem
    for extension in [".xlsx", ".xls", ".xlsm"]:
        ruta_alternativa = ruta_sin_extension.with_suffix(extension)
        if ruta_alternativa.exists():
            return ruta_alternativa

    raise FileNotFoundError(f"No se encontro {nombre_archivo} en {CARPETA_APP}")


def filtrar_tipo_muestra(df, tipo):
    palabras = TIPOS_MUESTRA[tipo]
    if not palabras:
        return df

    claves = [normalizar_texto(palabra) for palabra in palabras]
    muestras = df["Tipo de muestra"].apply(normalizar_texto)
    return df[muestras.apply(lambda valor: any(clave in valor for clave in claves))]


def top_microorganismos(df, n=10):
    if df.empty:
        return pd.DataFrame(columns=["Microorganismo", "#", "(%)"])

    conteo = df["Microorganismo"].value_counts().head(n)
    total = len(df)
    return pd.DataFrame({
        "Microorganismo": conteo.index,
        "#": conteo.values,
        "(%)": (conteo.values / total * 100).round(1),
    })


def es_urocultivo(df):
    if "Grupo_muestra" in df.columns:
        grupos = df["Grupo_muestra"].astype(str)
        if grupos.str.contains("urocultivo|urocultivos", case=False, na=False).any():
            return True

    if "Tipo de muestra" in df.columns:
        muestras = df["Tipo de muestra"].apply(normalizar_texto)
        return muestras.str.contains("orina|urocultivo|urocultivos", case=False, na=False).any()

    return False


def obtener_panel_clinico(germen, df_f):
    germen = normalizar_texto(germen)
    incluir_fosfomicina = True

    if "pseudomonas" in germen:
        panel = {
            "AMK": "Amikacina",
            "FEP": "Cefepime",
            "TZP": "Piperacilina tazobactam",
            "MEM": "Meropenem",
            "CIP": "Ciprofloxacino",
        }
    elif "acinetobacter" in germen:
        panel = {
            "FEP": "Cefepime",
            "TZP": "Piperacilina tazobactam",
            "MEM": "Meropenem",
            "CIP": "Ciprofloxacino",
            "SAM": "Ampicilina sulbactam",
        }
    elif (
        "klebsiellaaerogenes" in germen
        or "citrobacter" in germen
        or "enterobacter" in germen
        or "serratia" in germen
        or "morganella" in germen
        or "providencia" in germen
    ):
        panel = {
            "AMK": "Amikacina",
            "FEP": "Cefepime",
            "ETP": "Ertapenem",
            "MEM": "Meropenem",
            "CIP": "Ciprofloxacino",
            "SXT": "Trimetoprim sulfametoxazol",
        }
        incluir_fosfomicina = False
    elif "enterococcus" in germen:
        panel = {
            "PEN": "Penicilina",
            "AMP": "Ampicilina",
            "VAN": "Vancomicina",
            "LNZ": "Linezolid",
            "TET": "Tetraciclina",
        }
        incluir_fosfomicina = False
    elif "staphyl" in germen:
        panel = {
            "OXA": "Oxacilina",
            "SXT": "Trimetoprim sulfametoxazol",
            "CLI": "Clindamicina",
            "TET": "Tetraciclina",
        }
    elif "stenotrophomonas" in germen:
        panel = {
            "SXT": "Trimetoprim sulfametoxazol",
            "LVX": "Levofloxacina",
        }
    elif "burkholderia" in germen:
        panel = {
            "SXT": "Trimetoprim sulfametoxazol",
            "LVX": "Levofloxacina",
            "MEM": "Meropenem",
            "CAZ": "Ceftazidima",
        }
    elif "candida" in germen:
        panel = {
            "FLU": "Fluconazol",
            "CAS": "Caspofungina",
            "VOR": "Voriconazol",
            "AMB": "Anfotericina B",
        }
    else:
        panel = {
            "AMK": "Amikacina",
            "SAM": "Ampicilina sulbactam",
            "CZO": "Cefazolina",
            "CXM": "Cefuroxima",
            "CRO": "Ceftriaxona",
            "FEP": "Cefepime",
            "TZP": "Piperacilina tazobactam",
            "ETP": "Ertapenem",
            "MEM": "Meropenem",
            "CIP": "Ciprofloxacino",
            "SXT": "Trimetoprim sulfametoxazol",
        }

    if incluir_fosfomicina and es_urocultivo(df_f):
        panel["FOS"] = "Fosfomicina"

    return panel


def perfil_resistencia(df, top):
    filas = []
    columnas_panel = []

    for _, item in top.iterrows():
        germen = item["Microorganismo"]
        df_germen = df[df["Microorganismo"] == germen]
        panel = obtener_panel_clinico(germen, df_germen)
        fila = {
            "Microorganismo": germen,
            "#": int(item["#"]),
            "(%)": float(item["(%)"]),
        }

        for codigo, nombre in panel.items():
            if nombre not in columnas_panel:
                columnas_panel.append(nombre)

            if codigo in df_germen.columns:
                total = df_germen[codigo].notna().sum()
                resistentes = df_germen[codigo].isin(["R"]).sum()
                fila[nombre] = round(resistentes / total * 100, 1) if total > 0 else np.nan
            else:
                fila[nombre] = np.nan

        filas.append(fila)

    columnas = ["Microorganismo", "#", "(%)"] + columnas_panel
    return pd.DataFrame(filas).reindex(columns=columnas)


def indicadores_proa(df):
    indicadores = [
        ("E. coli BLEE", df[df["Microorganismo"].str.contains("coli", case=False, na=False)], "CRO"),
        ("Klebsiella BLEE", df[df["Microorganismo"].str.contains("kleb", case=False, na=False)], "CRO"),
        ("MRSA", df[df["Microorganismo"].str.contains("aureus", case=False, na=False)], "OXA"),
        ("Resistencia a carbapenemicos", df, "MEM"),
    ]

    resultados = []
    for nombre, subdf, antibiotico in indicadores:
        if antibiotico in subdf.columns:
            total = subdf[antibiotico].notna().sum()
            resistentes = subdf[antibiotico].isin(["R"]).sum()
            valor = round(resistentes / total * 100, 1) if total > 0 else np.nan
        else:
            valor = np.nan
        resultados.append({"Indicador": nombre, "%": valor})

    return pd.DataFrame(resultados)


def clasificar_proa(valor):
    if pd.isna(valor):
        return ""
    if valor <= 20:
        return "Adecuado"
    if valor <= 30:
        return "Precaucion"
    return "Evitar uso empirico"


def colorear_resistencia(valor):
    if pd.isna(valor):
        return ""

    texto = "color: #0B1F4D; font-weight: 700"
    if valor <= 20:
        return f"background-color: #C6EFCE; {texto}"
    if valor <= 30:
        return f"background-color: #FFEB9C; {texto}"
    return f"background-color: #FFC7CE; {texto}"


@st.cache_data(show_spinner="Cargando archivos del hospital...")
def cargar_datos():
    bases = []
    archivos_cargados = []

    for servicio, archivo in ARCHIVOS_SERVICIOS.items():
        ruta = resolver_archivo_excel(archivo)
        df_servicio = pd.read_excel(ruta)

        col_microorganismo = buscar_columna(
            df_servicio,
            [
                "Codigo de microorganismo local",
                "Código de microorganismo local",
                "CÃ³digo de microorganismo local",
                "CÃƒÂ³digo de microorganismo local",
            ],
        )
        col_muestra = buscar_columna(
            df_servicio,
            [
                "Codigo de muestra local",
                "Código de muestra local",
                "CÃ³digo de muestra local",
                "CÃƒÂ³digo de muestra local",
            ],
        )

        df_servicio["Microorganismo"] = df_servicio[col_microorganismo].apply(normalizar_microorganismo)
        df_servicio["Tipo de muestra"] = df_servicio[col_muestra].astype(str).str.strip().str.lower()
        df_servicio["Servicio"] = servicio
        df_servicio["Archivo fuente"] = ruta.name

        bases.append(df_servicio)
        archivos_cargados.append({"Servicio": servicio, "Archivo": ruta.name, "Registros": len(df_servicio)})

    return pd.concat(bases, ignore_index=True), pd.DataFrame(archivos_cargados)


st.set_page_config(page_title="Perfil microbiologico HRS", layout="wide")

if RUTA_LOGO.exists():
    st.image(str(RUTA_LOGO), width=1100)
else:
    st.warning(f"No se encontro el logo: {RUTA_LOGO.name}")

st.title("Perfil microbiologico hospitalario")

try:
    df, archivos = cargar_datos()
except Exception as exc:
    st.error(str(exc))
    st.stop()

servicios_dashboard = ["Perfil General"] + list(ARCHIVOS_SERVICIOS.keys())

with st.sidebar:
    if RUTA_LOGO.exists():
        st.image(str(RUTA_LOGO), width="stretch")
    st.header("Filtros")
    servicio = st.selectbox("Servicio", servicios_dashboard)
    tipo_muestra = st.selectbox("Tipo de muestra", list(TIPOS_MUESTRA.keys()))

if servicio == "Perfil General":
    df_filtrado = df.copy()
else:
    df_filtrado = df[df["Servicio"] == servicio].copy()

df_filtrado = filtrar_tipo_muestra(df_filtrado, tipo_muestra)

if servicio == "Perfil General" or tipo_muestra == "Todas":
    top_n = 20
else:
    top_n = 5

col1, col2, col3 = st.columns(3)
col1.metric("Registros", f"{len(df_filtrado):,}")
col2.metric("Servicios incluidos", df_filtrado["Servicio"].nunique() if not df_filtrado.empty else 0)
col3.metric("Microorganismos", df_filtrado["Microorganismo"].nunique() if not df_filtrado.empty else 0)

st.divider()

tab_perfil, tab_proa, tab_archivos = st.tabs(["Perfil", "Resumen PROA", "Archivos fuente"])

with tab_perfil:
    st.subheader(f"{servicio} - {tipo_muestra}")

    if df_filtrado.empty:
        st.warning("Sin datos para los filtros seleccionados.")
    else:
        top = top_microorganismos(df_filtrado, top_n)
        perfil = perfil_resistencia(df_filtrado, top)

        grafico = top.sort_values("#", ascending=False).copy()
        barras = (
            alt.Chart(grafico)
            .mark_bar()
            .encode(
                x=alt.X(
                    "Microorganismo:N",
                    sort=alt.EncodingSortField(field="#", order="descending"),
                    title=None,
                ),
                y=alt.Y("#:Q", title="#"),
                tooltip=["Microorganismo", "#", alt.Tooltip("(%)", format=".1f")],
            )
        )
        st.altair_chart(barras, width="stretch")

        perfil["(%)"] = perfil["(%)"].apply(lambda valor: f"{valor:.1f}%" if pd.notna(valor) else "")
        columnas_resistencia = [col for col in perfil.columns if col not in ["Microorganismo", "#", "(%)"]]
        perfil_tabla = perfil.set_index("Microorganismo")
        tabla = (
            perfil_tabla.style
            .format({col: "{:.0f}%" for col in columnas_resistencia}, na_rep="")
            .map(colorear_resistencia, subset=columnas_resistencia)
        )
        st.dataframe(tabla, width="stretch")

with tab_proa:
    resumen = indicadores_proa(df_filtrado)
    resumen["Interpretacion"] = resumen["%"].apply(clasificar_proa)

    st.dataframe(
        resumen.style.format({"%": "{:.1f}%"}, na_rep=""),
        width="stretch",
        hide_index=True,
    )

with tab_archivos:
    st.dataframe(archivos, width="stretch", hide_index=True)
    st.caption("Los archivos deben estar en la misma carpeta del dashboard.py al publicar en Streamlit.")
