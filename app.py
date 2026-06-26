import io
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Liquidador de Marcaciones Part Time",
    layout="wide"
)

st.title("Liquidador de Marcaciones Part Time")
st.caption("Liquida horas entre fecha/hora entrada y fecha/hora salida en bloques completos de 30 minutos.")

uploaded = st.file_uploader(
    "Carga el reporte de marcaciones",
    type=["xlsx", "xls", "csv"]
)

def read_file(file):
    name = file.name.lower()
    if name.endswith(".csv"):
        return pd.read_csv(file)
    return pd.read_excel(file)

def combine_fecha_hora(fecha_col, hora_col):
    fecha = pd.to_datetime(fecha_col, errors="coerce", dayfirst=True)

    hora_txt = hora_col.astype(str).str.strip()

    hora = pd.to_datetime(
        hora_txt,
        errors="coerce"
    ).dt.strftime("%H:%M:%S")

    combinado = fecha.dt.strftime("%Y-%m-%d") + " " + hora

    return pd.to_datetime(combinado, errors="coerce")

def liquidar_bloques_30(minutos):
    if pd.isna(minutos) or minutos < 0:
        return 0.0

    bloques = int(minutos // 30)
    return bloques * 0.5

def formato_horas(valor):
    try:
        return str(valor).replace(".", ",")
    except Exception:
        return valor

if uploaded:
    try:
        df = read_file(uploaded)
    except Exception as e:
        st.error(f"No fue posible leer el archivo: {e}")
        st.stop()

    required_cols = [
        "FechaEntrada",
        "HoraEntrada",
        "FechaSalida",
        "HoraSalida"
    ]

    missing = [c for c in required_cols if c not in df.columns]

    if missing:
        st.error(f"Faltan columnas obligatorias: {missing}")
        st.stop()

    result = df.copy()

    result["FechaHoraEntrada"] = combine_fecha_hora(
        result["FechaEntrada"],
        result["HoraEntrada"]
    )

    result["FechaHoraSalida"] = combine_fecha_hora(
        result["FechaSalida"],
        result["HoraSalida"]
    )

    result["Minutos_Marcacion"] = (
        result["FechaHoraSalida"] - result["FechaHoraEntrada"]
    ).dt.total_seconds() / 60

    result["Horas_Liquidadas"] = result["Minutos_Marcacion"].apply(liquidar_bloques_30)

    result["Cruza_Medianoche"] = "NO"
    result.loc[
        result["FechaHoraSalida"].dt.date > result["FechaHoraEntrada"].dt.date,
        "Cruza_Medianoche"
    ] = "SI"

    result["Observacion"] = ""

    result.loc[
        result["FechaHoraEntrada"].isna(),
        "Observacion"
    ] = "Fecha/hora entrada inválida"

    result.loc[
        result["FechaHoraSalida"].isna(),
        "Observacion"
    ] = "Fecha/hora salida inválida"

    result.loc[
        result["Minutos_Marcacion"] < 0,
        "Observacion"
    ] = "Marcación salida menor que entrada"

    result.loc[
        result["Minutos_Marcacion"] < 0,
        "Horas_Liquidadas"
    ] = 0.0

    result["Horas_Liquidadas_Texto"] = result["Horas_Liquidadas"].apply(formato_horas)

    st.success("Liquidación generada correctamente.")

    c1, c2, c3, c4 = st.columns(4)

    c1.metric("Registros", len(result))
    c2.metric("Con error", int((result["Observacion"] != "").sum()))
    c3.metric("Cruzan medianoche", int((result["Cruza_Medianoche"] == "SI").sum()))
    c4.metric("Total horas", round(result["Horas_Liquidadas"].sum(), 2))

    st.subheader("Resultado")
    st.dataframe(result, use_container_width=True)

    output = io.BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        result.to_excel(writer, index=False, sheet_name="Marcaciones_Liquidadas")

        resumen = pd.DataFrame({
            "Indicador": [
                "Registros",
                "Registros con error",
                "Registros que cruzan medianoche",
                "Total horas liquidadas"
            ],
            "Valor": [
                len(result),
                int((result["Observacion"] != "").sum()),
                int((result["Cruza_Medianoche"] == "SI").sum()),
                round(result["Horas_Liquidadas"].sum(), 2)
            ]
        })

        resumen.to_excel(writer, index=False, sheet_name="Resumen")

    st.download_button(
        label="Descargar Excel liquidado",
        data=output.getvalue(),
        file_name="Marcaciones_Liquidadas.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )