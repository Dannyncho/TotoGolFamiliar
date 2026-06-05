import os
from pathlib import Path
from io import BytesIO
from datetime import datetime

import pandas as pd
import streamlit as st


BASE_DIR = Path(__file__).parent
DATA_DIR = BASE_DIR / "data"
JUGADORES_DIR = DATA_DIR / "Jugadores"
OUTPUT_DIR = BASE_DIR / "Output"

PARTIDOS_FILE = DATA_DIR / "partidos.xlsx"
RESULTADOS_FILE = DATA_DIR / "resultados.xlsx"

COLUMNAS_PARTIDOS = ["id_partido", "grupo", "local", "visitante"]
COLUMNAS_RESULTADOS = ["id_partido", "goles_local_real", "goles_visitante_real"]
COLUMNAS_JUGADOR = ["id_partido", "goles_local", "goles_visitante"]

PUNTOS_MARCADOR_EXACTO = 5
PUNTOS_RESULTADO_CORRECTO = 3
PUNTOS_RESULTADO_INCORRECTO = 0


st.set_page_config(
    page_title="Toto Gol Familiar",
    page_icon="⚽",
    layout="wide",
)


def crear_carpetas_necesarias():
    DATA_DIR.mkdir(exist_ok=True)
    JUGADORES_DIR.mkdir(parents=True, exist_ok=True)
    OUTPUT_DIR.mkdir(exist_ok=True)


def aplicar_estilos():
    st.markdown(
        """
        <style>
        .main {
            background: linear-gradient(180deg, #f7fafc 0%, #eef2f7 100%);
        }
        .titulo-principal {
            text-align: center;
            padding: 18px;
            border-radius: 18px;
            background: linear-gradient(135deg, #0f5132, #198754);
            color: white;
            margin-bottom: 20px;
            box-shadow: 0 6px 20px rgba(0,0,0,0.15);
        }
        .subtitulo {
            text-align: center;
            color: #e9f7ef;
            font-size: 18px;
        }
        .card-periodista {
            background: white;
            padding: 22px;
            border-radius: 18px;
            border-left: 8px solid #198754;
            box-shadow: 0 6px 20px rgba(0,0,0,0.10);
            font-size: 19px;
            line-height: 1.6;
        }
        .whatsapp-box {
            background: #e7f8ef;
            padding: 18px;
            border-radius: 16px;
            border: 1px solid #b7ebcd;
            font-family: monospace;
            white-space: pre-wrap;
        }
        .stMetric {
            background: white;
            padding: 14px;
            border-radius: 16px;
            box-shadow: 0 4px 14px rgba(0,0,0,0.08);
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def mostrar_encabezado():
    st.markdown(
        """
        <div class="titulo-principal">
            <h1>⚽ Toto Gol Familiar</h1>
            <div class="subtitulo">Dashboard deportivo de predicciones familiares</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def leer_excel(ruta: Path, columnas_requeridas: list[str], nombre_archivo: str) -> pd.DataFrame:
    if not ruta.exists():
        st.error(f"❌ No se encontró el archivo requerido: {nombre_archivo}")
        st.stop()

    try:
        df = pd.read_excel(ruta)
    except Exception as error:
        st.error(f"❌ No se pudo leer {nombre_archivo}. Verifica que sea un archivo Excel válido.")
        st.caption(str(error))
        st.stop()

    validar_columnas(df, columnas_requeridas, nombre_archivo)
    return df


def validar_columnas(df: pd.DataFrame, columnas_requeridas: list[str], nombre_archivo: str):
    columnas_faltantes = [col for col in columnas_requeridas if col not in df.columns]

    if columnas_faltantes:
        st.error(
            f"❌ El archivo {nombre_archivo} no tiene el formato correcto. "
            f"Faltan estas columnas: {', '.join(columnas_faltantes)}"
        )
        st.stop()


def obtener_archivos_jugadores() -> list[Path]:
    if not JUGADORES_DIR.exists():
        st.error("❌ No existe la carpeta data/Jugadores.")
        st.stop()

    archivos = sorted(JUGADORES_DIR.glob("*.xlsx"))

    if not archivos:
        st.warning("⚠️ No hay participantes cargados en data/Jugadores.")
        st.stop()

    return archivos


def limpiar_id_partido(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df["id_partido"] = df["id_partido"].astype(str).str.strip()
    return df


def convertir_goles_a_numero(df: pd.DataFrame, columnas: list[str]) -> pd.DataFrame:
    df = df.copy()

    for columna in columnas:
        df[columna] = pd.to_numeric(df[columna], errors="coerce")

    return df


def resultado_partido(goles_local, goles_visitante) -> str:
    if goles_local > goles_visitante:
        return "local"
    if goles_local < goles_visitante:
        return "visitante"
    return "empate"


def calcular_puntos_fila(fila: pd.Series) -> tuple[int, bool, bool, bool]:
    goles_local = fila["goles_local"]
    goles_visitante = fila["goles_visitante"]
    goles_local_real = fila["goles_local_real"]
    goles_visitante_real = fila["goles_visitante_real"]

    marcador_exacto = (
        goles_local == goles_local_real
        and goles_visitante == goles_visitante_real
    )

    if marcador_exacto:
        return PUNTOS_MARCADOR_EXACTO, True, False, False

    resultado_pronostico = resultado_partido(goles_local, goles_visitante)
    resultado_real = resultado_partido(goles_local_real, goles_visitante_real)

    resultado_correcto = resultado_pronostico == resultado_real

    if resultado_correcto:
        return PUNTOS_RESULTADO_CORRECTO, False, True, False

    return PUNTOS_RESULTADO_INCORRECTO, False, False, True


def procesar_jugador(
    archivo_jugador: Path,
    partidos: pd.DataFrame,
    resultados: pd.DataFrame,
) -> tuple[dict, pd.DataFrame]:
    nombre_jugador = archivo_jugador.stem.replace(".xlsx", "")

    try:
        pronosticos = pd.read_excel(archivo_jugador)
    except Exception as error:
        st.error(f"❌ No se pudo leer el archivo del jugador {archivo_jugador.name}.")
        st.caption(str(error))
        st.stop()

    validar_columnas(pronosticos, COLUMNAS_JUGADOR, archivo_jugador.name)

    pronosticos = limpiar_id_partido(pronosticos)
    pronosticos = convertir_goles_a_numero(pronosticos, ["goles_local", "goles_visitante"])
    pronosticos = pronosticos.dropna(subset=["goles_local", "goles_visitante"])

    detalle = (
        pronosticos
        .merge(resultados, on="id_partido", how="inner")
        .merge(partidos, on="id_partido", how="left")
    )

    if detalle.empty:
        resumen = {
            "Jugador": nombre_jugador,
            "Puntos": 0,
            "Marcadores exactos": 0,
            "Resultados correctos": 0,
            "Predicciones incorrectas": 0,
        }

        detalle["Jugador"] = nombre_jugador
        detalle["Puntos"] = []
        return resumen, detalle

    detalle[["Puntos", "Marcador exacto", "Resultado correcto", "Predicción incorrecta"]] = detalle.apply(
        lambda fila: pd.Series(calcular_puntos_fila(fila)),
        axis=1,
    )

    detalle["Jugador"] = nombre_jugador

    resumen = {
        "Jugador": nombre_jugador,
        "Puntos": int(detalle["Puntos"].sum()),
        "Marcadores exactos": int(detalle["Marcador exacto"].sum()),
        "Resultados correctos": int(detalle["Resultado correcto"].sum()),
        "Predicciones incorrectas": int(detalle["Predicción incorrecta"].sum()),
    }

    return resumen, detalle


def calcular_ranking(partidos: pd.DataFrame, resultados: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    archivos_jugadores = obtener_archivos_jugadores()

    resumenes = []
    detalles = []

    for archivo in archivos_jugadores:
        resumen, detalle = procesar_jugador(archivo, partidos, resultados)
        resumenes.append(resumen)
        detalles.append(detalle)

    ranking = pd.DataFrame(resumenes)

    ranking = ranking.sort_values(
        by=["Puntos", "Marcadores exactos", "Jugador"],
        ascending=[False, False, True],
    ).reset_index(drop=True)

    ranking.insert(0, "Posición", range(1, len(ranking) + 1))

    detalle_general = pd.concat(detalles, ignore_index=True) if detalles else pd.DataFrame()

    return ranking, detalle_general


def agregar_medallas(ranking: pd.DataFrame) -> pd.DataFrame:
    ranking_visual = ranking.copy()

    ranking_visual["Medalla"] = ""

    ranking_visual.loc[
        ranking_visual["Posición"] == 1,
        "Medalla"
    ] = "🥇"

    ranking_visual.loc[
        ranking_visual["Posición"] == 2,
        "Medalla"
    ] = "🥈"

    ranking_visual.loc[
        ranking_visual["Posición"] == 3,
        "Medalla"
    ] = "🥉"

    columnas = ["Medalla"] + [
        col for col in ranking_visual.columns
        if col != "Medalla"
    ]

    return ranking_visual[columnas]


def generar_excel_ranking(ranking: pd.DataFrame) -> BytesIO:
    buffer = BytesIO()

    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        ranking.to_excel(writer, sheet_name="Ranking", index=False)

    buffer.seek(0)
    return buffer


def guardar_ranking_output(ranking: pd.DataFrame):
    archivo_salida = OUTPUT_DIR / "ranking_toto_gol.xlsx"

    try:
        ranking.to_excel(archivo_salida, index=False)
    except Exception:
        pass


def unir_nombres(nombres: list[str]) -> str:
    if len(nombres) == 0:
        return ""
    if len(nombres) == 1:
        return nombres[0]
    if len(nombres) == 2:
        return f"{nombres[0]} y {nombres[1]}"
    return ", ".join(nombres[:-1]) + f" y {nombres[-1]}"


def generar_comentario_periodista(ranking: pd.DataFrame) -> str:
    if ranking.empty:
        return "La jornada aún no tiene líder. El Toto Gol Familiar espera sus primeros resultados oficiales."

    ranking_ordenado = ranking.sort_values(
        by=["Puntos", "Marcadores exactos", "Jugador"],
        ascending=[False, False, True],
    ).reset_index(drop=True)

    puntos_lider = int(ranking_ordenado.iloc[0]["Puntos"])
    puntos_ultimo = int(ranking_ordenado.iloc[-1]["Puntos"])

    lideres = ranking_ordenado[ranking_ordenado["Puntos"] == puntos_lider]
    ultimos = ranking_ordenado[ranking_ordenado["Puntos"] == puntos_ultimo]

    lideres_texto = unir_nombres(lideres["Jugador"].tolist())
    ultimos_texto = unir_nombres(ultimos["Jugador"].tolist())

    if puntos_lider == 0:
        return (
            "¡Arranca la emoción del Toto Gol Familiar! "
            "La tabla todavía está completamente abierta y cualquier marcador exacto puede cambiar la historia. "
            f"Por ahora, {lideres_texto} comparten el inicio de esta competencia familiar."
        )

    if len(lideres) == 1:
        intro = (
            f"🏆 ¡{lideres_texto} toma el mando del Toto Gol Familiar con {puntos_lider} puntos! "
        )
    else:
        intro = (
            f"🔥 ¡Empate total en la cima del Toto Gol Familiar! "
            f"{lideres_texto} comparten el liderato con {puntos_lider} puntos. "
        )

    perseguidores = ranking_ordenado[
        (ranking_ordenado["Puntos"] < puntos_lider)
        & (ranking_ordenado["Puntos"] >= puntos_lider - 3)
    ]

    if not perseguidores.empty:
        nombres_perseguidores = unir_nombres(perseguidores["Jugador"].tolist())
        diferencia = puntos_lider - int(perseguidores.iloc[0]["Puntos"])

        if diferencia == 1:
            parte_persecucion = (
                f"{nombres_perseguidores} se mantienen muy cerca, a tan solo un punto de distancia. "
            )
        else:
            parte_persecucion = (
                f"{nombres_perseguidores} siguen en la pelea, a menos de tres puntos de la cima. "
            )
    else:
        parte_persecucion = (
            "Por ahora, la parte alta de la tabla empieza a marcar diferencias importantes. "
        )

    if len(ultimos) == len(ranking_ordenado):
        parte_baja = (
            "La clasificación todavía no marca diferencias en la parte baja de la tabla. "
        )
    elif len(ultimos) == 1:
        parte_baja = (
            f"En la parte baja, {ultimos_texto} ocupa momentáneamente la última posición, "
            "aunque todavía queda mucho torneo por disputar. "
        )
    else:
        parte_baja = (
            f"En la parte baja, {ultimos_texto} ocupan momentáneamente la última posición, "
            "aunque todavía queda mucho torneo por disputar. "
        )

    diferencia_total = puntos_lider - puntos_ultimo

    cierre = (
        f"La diferencia entre la cima y el fondo de la tabla es de {diferencia_total} punto(s). "
        "Con tantos partidos por jugar, un marcador exacto puede cambiar completamente la historia."
    )

    return intro + parte_persecucion + parte_baja + cierre


def generar_resumen_whatsapp(ranking: pd.DataFrame) -> str:
    if ranking.empty:
        return "⚽ Toto Gol Familiar\n\nAún no hay datos disponibles."

    fecha = datetime.now().strftime("%d/%m/%Y %H:%M")

    ranking_ordenado = ranking.sort_values(
        by=["Puntos", "Marcadores exactos", "Jugador"],
        ascending=[False, False, True],
    ).reset_index(drop=True)

    puntos_lider = int(ranking_ordenado.iloc[0]["Puntos"])
    puntos_ultimo = int(ranking_ordenado.iloc[-1]["Puntos"])

    lideres = ranking_ordenado[ranking_ordenado["Puntos"] == puntos_lider]
    ultimos = ranking_ordenado[ranking_ordenado["Puntos"] == puntos_ultimo]

    lideres_texto = unir_nombres(lideres["Jugador"].tolist())
    ultimos_texto = unir_nombres(ultimos["Jugador"].tolist())

    perseguidores = ranking_ordenado[
        (ranking_ordenado["Puntos"] < puntos_lider)
        & (ranking_ordenado["Puntos"] >= puntos_lider - 1)
    ]

    lineas = [
        "⚽ *Toto Gol Familiar*",
        f"📅 Actualizado: {fecha}",
        "",
        "🏆 *Líderes actuales*",
        f"🥇 {lideres_texto} ({puntos_lider} pts)",
    ]

    if not perseguidores.empty:
        perseguidores_texto = unir_nombres(perseguidores["Jugador"].tolist())

        lineas.extend(
            [
                "",
                "🔥 *A la caza del liderato*",
                f"{perseguidores_texto} están a solo 1 punto de la cima.",
            ]
        )

    if len(ultimos) != len(ranking_ordenado):
        lineas.extend(
            [
                "",
                "⚠️ *Última posición*",
                f"{ultimos_texto} ocupan momentáneamente el último lugar.",
            ]
        )

    lineas.extend(
        [
            "",
            f"📊 Diferencia entre primero y último: {puntos_lider - puntos_ultimo} puntos",
            "",
            "⚽ ¡La competencia sigue completamente abierta!",
        ]
    )

    return "\n".join(lineas)


def mostrar_metricas(ranking: pd.DataFrame, detalle_general: pd.DataFrame, resultados: pd.DataFrame):
    total_jugadores = len(ranking)
    total_partidos_procesados = resultados["id_partido"].nunique()
    total_exactos = int(ranking["Marcadores exactos"].sum()) if not ranking.empty else 0
    lider = ranking.iloc[0]["Jugador"] if not ranking.empty else "Sin líder"

    col1, col2, col3, col4 = st.columns(4)

    col1.metric("🏆 Líder actual", lider)
    col2.metric("👥 Total de jugadores", total_jugadores)
    col3.metric("📅 Partidos procesados", total_partidos_procesados)
    col4.metric("🎯 Marcadores exactos", total_exactos)


def mostrar_dashboard(ranking: pd.DataFrame, detalle_general: pd.DataFrame, resultados: pd.DataFrame):
    mostrar_metricas(ranking, detalle_general, resultados)

    plantilla_path = OUTPUT_DIR / "Plantilla_Jugador.xlsx"

    if plantilla_path.exists():
        with open(plantilla_path, "rb") as archivo:
            st.download_button(
                label="📄 Descargar plantilla de jugador",
                data=archivo,
                file_name="Plantilla_Jugador.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
    else:
        st.warning(
            "⚠️ No se encontró Output/Plantilla_Jugador.xlsx. Ejecuta primero: py crear_plantilla.py"
        )

    st.markdown("## 🏆 Podio actual")

    if len(ranking) >= 3:
        col1, col2, col3 = st.columns(3)

        col1.metric(
            "🥇 Primer Lugar",
            ranking.iloc[0]["Jugador"],
            f"{int(ranking.iloc[0]['Puntos'])} pts"
        )

        col2.metric(
            "🥈 Segundo Lugar",
            ranking.iloc[1]["Jugador"],
            f"{int(ranking.iloc[1]['Puntos'])} pts"
        )

        col3.metric(
            "🥉 Tercer Lugar",
            ranking.iloc[2]["Jugador"],
            f"{int(ranking.iloc[2]['Puntos'])} pts"
        )

        st.divider()
    else:
        st.info("Aún no hay suficientes jugadores para mostrar el podio.")

    st.markdown("## 🏆 Tabla de posiciones general")

    ranking_visual = agregar_medallas(ranking)

    st.dataframe(
        ranking_visual,
        use_container_width=True,
        hide_index=True,
    )

    guardar_ranking_output(ranking)

    st.markdown("## 📊 Gráfico de posiciones")

    grafico = ranking.sort_values(
        by="Puntos",
        ascending=False
    ).set_index("Jugador")["Puntos"]

    st.bar_chart(grafico)

    excel = generar_excel_ranking(ranking)

    st.download_button(
        label="📥 Exportar ranking a Excel",
        data=excel,
        file_name="ranking_toto_gol.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    st.markdown("## 🎙️ Comentario deportivo")
    comentario = generar_comentario_periodista(ranking)
    st.markdown(
        f"<div class='card-periodista'>{comentario}</div>",
        unsafe_allow_html=True,
    )

    st.markdown("## 📲 Resumen para WhatsApp")
    resumen = generar_resumen_whatsapp(ranking)
    st.markdown(
        f"<div class='whatsapp-box'>{resumen}</div>",
        unsafe_allow_html=True,
    )


def mostrar_detalle_jugador(ranking: pd.DataFrame, detalle_general: pd.DataFrame):
    st.markdown("## 👤 Detalle de Jugador")

    if detalle_general.empty:
        st.warning("⚠️ Todavía no hay partidos procesados para mostrar detalles.")
        return

    jugadores = sorted(detalle_general["Jugador"].dropna().unique())

    jugador_seleccionado = st.selectbox(
        "Selecciona un jugador",
        jugadores,
    )

    detalle = detalle_general[detalle_general["Jugador"] == jugador_seleccionado].copy()

    if detalle.empty:
        st.warning("⚠️ Este jugador aún no tiene partidos con resultado oficial.")
        return

    detalle = detalle.sort_values(by="id_partido")

    columnas_mostrar = [
        "id_partido",
        "grupo",
        "local",
        "visitante",
        "goles_local",
        "goles_visitante",
        "goles_local_real",
        "goles_visitante_real",
        "Puntos",
    ]

    columnas_disponibles = [col for col in columnas_mostrar if col in detalle.columns]

    total = int(detalle["Puntos"].sum())
    exactos = int(detalle["Marcador exacto"].sum())
    correctos = int(detalle["Resultado correcto"].sum())
    incorrectos = int(detalle["Predicción incorrecta"].sum())

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Total acumulado", total)
    col2.metric("Marcadores exactos", exactos)
    col3.metric("Resultados correctos", correctos)
    col4.metric("Incorrectas", incorrectos)

    st.dataframe(
        detalle[columnas_disponibles],
        use_container_width=True,
        hide_index=True,
    )


def preparar_datos() -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    partidos = leer_excel(PARTIDOS_FILE, COLUMNAS_PARTIDOS, "data/partidos.xlsx")
    resultados = leer_excel(RESULTADOS_FILE, COLUMNAS_RESULTADOS, "data/resultados.xlsx")

    partidos = limpiar_id_partido(partidos)
    resultados = limpiar_id_partido(resultados)

    resultados = convertir_goles_a_numero(
        resultados,
        ["goles_local_real", "goles_visitante_real"],
    )

    resultados = resultados.dropna(
        subset=["goles_local_real", "goles_visitante_real"]
    )

    ranking, detalle_general = calcular_ranking(partidos, resultados)

    return partidos, resultados, ranking, detalle_general


def mostrar_sidebar():
    st.sidebar.title("⚽ Toto Gol Familiar")
    st.sidebar.info(
        "Los datos se leen desde archivos Excel locales dentro de la carpeta data."
    )

    st.sidebar.markdown("### Estructura esperada")
    st.sidebar.code(
        """data/
├── Jugadores/
│   ├── Danny.xlsx
│   ├── Maria.xlsx
│   └── Carlos.xlsx
├── partidos.xlsx
└── resultados.xlsx

Output/""",
        language="text",
    )

    st.sidebar.markdown("### Puntuación")
    st.sidebar.write("🎯 Marcador exacto: 5 puntos")
    st.sidebar.write("✅ Resultado correcto: 3 puntos")
    st.sidebar.write("❌ Resultado incorrecto: 0 puntos")
    st.sidebar.markdown("### DESCARGA TEST 123")

    plantilla_path = OUTPUT_DIR / "Plantilla_Jugador.xlsx"

    if plantilla_path.exists():
        with open(plantilla_path, "rb") as archivo:
            st.sidebar.download_button(
                label="📄 Descargar plantilla",
                data=archivo,
                file_name="Plantilla_Jugador.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
    else:
        st.sidebar.warning("No se encontró la plantilla.")


def main():
    crear_carpetas_necesarias()
    aplicar_estilos()
    mostrar_encabezado()
    mostrar_sidebar()

    partidos, resultados, ranking, detalle_general = preparar_datos()

    pagina = st.sidebar.radio(
        "Navegación",
        ["Dashboard Público", "Detalle de Jugador"],
    )

    if pagina == "Dashboard Público":
        mostrar_dashboard(ranking, detalle_general, resultados)
    else:
        mostrar_detalle_jugador(ranking, detalle_general)


if __name__ == "__main__":
    main()
