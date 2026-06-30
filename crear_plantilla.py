import pandas as pd
from pathlib import Path

archivo_partidos = Path("data") / "partidos.xlsx"
archivo_salida = Path("Output") / "Plantilla_Jugador.xlsx"

df = pd.read_excel(archivo_partidos)

print("Partidos leídos:", len(df))
print("Fechas encontradas:", df["fecha"].notna().sum())

plantilla = df[["id_partido", "fecha", "grupo", "local", "visitante"]].copy()
plantilla["fecha"] = pd.to_datetime(plantilla["fecha"], dayfirst=True, errors="coerce").dt.strftime("%d/%m/%Y")
plantilla["goles_local"] = ""
plantilla["goles_visitante"] = ""

archivo_salida.parent.mkdir(exist_ok=True)
plantilla.to_excel(archivo_salida, index=False)

print("Plantilla creada correctamente en:", archivo_salida)
print("Fechas en plantilla:", plantilla["fecha"].notna().sum())