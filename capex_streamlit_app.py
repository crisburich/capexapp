import io
from datetime import date

import numpy as np
import pandas as pd
import streamlit as st


# =========================================================
# CONFIG
# =========================================================
IPC_FILE_PATH = "IPC.xlsx"
IPC_SHEET_NAME = "pautas_py"

APP_TITLE = "Ajuste CAPEX ROI"
REQUIRED_COLUMNS = [
    "CAPEX ROI",
    "Valuación CAPEX ROI",
    "INICIO OBRA",
    "FIN OBRA",
    "Moneda",
]


# =========================================================
# CARGA DE SERIES
# =========================================================
@st.cache_data

def load_macro_series(file_path: str = IPC_FILE_PATH, sheet_name: str = IPC_SHEET_NAME):
    """
    Espera un Excel con una hoja similar a la de tu notebook, donde:
    - la fila 0 corresponde a IPC
    - la fila 1 corresponde a Dólar BNA
    - las columnas desde la segunda en adelante representan meses/fechas

    Si tu formato real difiere, esta función es la que habría que ajustar.
    """
    pautas = pd.read_excel(file_path, sheet_name=sheet_name, header=0)

    ipc_row = pautas.iloc[0]
    dolar_row = pautas.iloc[1]

    ipc_df = _series_row_to_df(ipc_row, value_name="ipc")
    dolar_bna_df = _series_row_to_df(dolar_row, value_name="dolar_bna")

    return ipc_df, dolar_bna_df


def _series_row_to_df(row: pd.Series, value_name: str) -> pd.DataFrame:
    records = []

    for col, value in row.items():
        if pd.isna(value):
            continue

        fecha = pd.to_datetime(col, errors="coerce", dayfirst=True)
        if pd.isna(fecha):
            continue

        records.append({"fecha": fecha, value_name: pd.to_numeric(value, errors="coerce")})

    df = pd.DataFrame(records).dropna().sort_values("fecha").reset_index(drop=True)
    return df


# =========================================================
# LÓGICA DE NEGOCIO
# =========================================================

def calcular_midpoint(inicio_obra, fin_obra):
    inicio = pd.to_datetime(inicio_obra, errors="coerce")
    fin = pd.to_datetime(fin_obra, errors="coerce")

    if pd.isna(inicio) or pd.isna(fin):
        return pd.NaT

    return inicio + (fin - inicio) / 2



def get_monthly_ipc_detail(start_date, end_date, ipc_df):
    """
    Replica la lógica de prorrateo mensual que usás en tu notebook.
    Devuelve detalle mensual + factor total.
    """
    start_date = pd.to_datetime(start_date, errors="coerce")
    end_date = pd.to_datetime(end_date, errors="coerce")

    if pd.isna(start_date) or pd.isna(end_date):
        return pd.DataFrame(), np.nan

    first_month = start_date.replace(day=1)
    end_of_month = end_date.replace(day=1) + pd.offsets.MonthEnd(1)
    months = pd.date_range(start=first_month, end=end_of_month, freq="M")

    detail_rows = []
    adjustment_factor = 1.0

    for month_end in months:
        month_start = month_end.replace(day=1)
        total_days = month_end.day

        if month_start.year == start_date.year and month_start.month == start_date.month:
            days_applied = total_days - start_date.day + 1
        elif month_start.year == end_date.year and month_start.month == end_date.month:
            days_applied = end_date.day
        else:
            days_applied = total_days

        mask = ipc_df["fecha"].dt.to_period("M") == month_end.to_period("M")
        ipc_value = ipc_df.loc[mask]

        if ipc_value.empty:
            detail_rows.append(
                {
                    "mes": month_end.to_period("M").strftime("%Y-%m"),
                    "ipc": np.nan,
                    "dias_aplicados": days_applied,
                    "dias_mes": total_days,
                    "ipc_efectivo": np.nan,
                    "factor_mes": np.nan,
                    "estado": "Sin dato IPC",
                }
            )
            continue

        ipc_rate = float(ipc_value["ipc"].iloc[0])
        effective_ipc = ipc_rate * (days_applied / total_days)
        factor_mes = 1 + effective_ipc
        adjustment_factor *= factor_mes

        detail_rows.append(
            {
                "mes": month_end.to_period("M").strftime("%Y-%m"),
                "ipc": ipc_rate,
                "dias_aplicados": days_applied,
                "dias_mes": total_days,
                "ipc_efectivo": effective_ipc,
                "factor_mes": factor_mes,
                "estado": "OK",
            }
        )

    detail_df = pd.DataFrame(detail_rows)
    return detail_df, adjustment_factor



def ajustar_capex_row(row: pd.Series, ipc_df: pd.DataFrame, dolar_bna_df: pd.DataFrame) -> pd.Series:
    start_date = pd.to_datetime(row.get("Valuación CAPEX ROI"), errors="coerce")
    inicio_obra = pd.to_datetime(row.get("INICIO OBRA"), errors="coerce")
    fin_obra = pd.to_datetime(row.get("FIN OBRA"), errors="coerce")
    moneda = row.get("Moneda")
    original_amount = pd.to_numeric(row.get("CAPEX ROI"), errors="coerce")

    midpoint = calcular_midpoint(inicio_obra, fin_obra)

    result = {
        "MIDPOINT": midpoint,
        "FACTOR_AJUSTE": np.nan,
        "CAPEX_AJUSTADO": np.nan,
        "SERIE_USADA": None,
        "OBSERVACIONES": None,
    }

    if pd.isna(original_amount):
        result["OBSERVACIONES"] = "CAPEX ROI inválido"
        return pd.Series(result)

    if pd.isna(start_date):
        result["OBSERVACIONES"] = "Valuación CAPEX ROI inválida"
        return pd.Series(result)

    if pd.isna(midpoint):
        result["OBSERVACIONES"] = "Fechas de obra inválidas"
        return pd.Series(result)

    if midpoint < start_date:
        result["OBSERVACIONES"] = "MIDPOINT anterior a valuación"
        return pd.Series(result)

    if moneda == "$":
        detail_df, factor = get_monthly_ipc_detail(start_date, midpoint, ipc_df)

        if detail_df.empty:
            result["OBSERVACIONES"] = "No se pudo calcular IPC"
            return pd.Series(result)

        if detail_df["ipc"].isna().any():
            result["OBSERVACIONES"] = "Faltan meses de IPC en la serie"
            result["SERIE_USADA"] = "IPC"
            return pd.Series(result)

        result["FACTOR_AJUSTE"] = factor
        result["CAPEX_AJUSTADO"] = original_amount * factor
        result["SERIE_USADA"] = "IPC"
        result["OBSERVACIONES"] = "OK"
        return pd.Series(result)

    if moneda == "USD":
        midpoint_period = midpoint.to_period("M")
        dolar_value = dolar_bna_df.loc[
            dolar_bna_df["fecha"].dt.to_period("M") == midpoint_period, "dolar_bna"
        ]

        if dolar_value.empty:
            result["OBSERVACIONES"] = "No hay dólar BNA para el mes del midpoint"
            result["SERIE_USADA"] = "Dólar BNA"
            return pd.Series(result)

        fx = float(dolar_value.iloc[0])
        result["FACTOR_AJUSTE"] = fx
        result["CAPEX_AJUSTADO"] = original_amount * fx
        result["SERIE_USADA"] = "Dólar BNA"
        result["OBSERVACIONES"] = "OK"
        return pd.Series(result)

    result["OBSERVACIONES"] = "Moneda no reconocida"
    return pd.Series(result)



def procesar_archivo(df_input: pd.DataFrame, ipc_df: pd.DataFrame, dolar_bna_df: pd.DataFrame):
    missing = [c for c in REQUIRED_COLUMNS if c not in df_input.columns]
    if missing:
        raise ValueError(f"Faltan columnas requeridas: {missing}")

    df = df_input.copy()
    df["Valuación CAPEX ROI"] = pd.to_datetime(df["Valuación CAPEX ROI"], errors="coerce")
    df["INICIO OBRA"] = pd.to_datetime(df["INICIO OBRA"], errors="coerce")
    df["FIN OBRA"] = pd.to_datetime(df["FIN OBRA"], errors="coerce")
    df["CAPEX ROI"] = pd.to_numeric(df["CAPEX ROI"], errors="coerce")
    df["Moneda"] = df["Moneda"].astype(str).str.strip()

    results = df.apply(ajustar_capex_row, axis=1, ipc_df=ipc_df, dolar_bna_df=dolar_bna_df)
    output = pd.concat([df, results], axis=1)
    return output



def build_template() -> pd.DataFrame:
    return pd.DataFrame(
        {
            "CAPEX ROI": [15000000, 220000],
            "Valuación CAPEX ROI": ["2025-01-15", "2025-02-01"],
            "INICIO OBRA": ["2025-03-01", "2025-04-01"],
            "FIN OBRA": ["2025-10-15", "2025-12-20"],
            "Moneda": ["$", "USD"],
        }
    )



def dataframe_to_excel_bytes(df: pd.DataFrame, detail_df: pd.DataFrame | None = None) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="resultado")
        if detail_df is not None and not detail_df.empty:
            detail_df.to_excel(writer, index=False, sheet_name="detalle")
    output.seek(0)
    return output.getvalue()


# =========================================================
# UI
# =========================================================

def main():
    st.set_page_config(page_title=APP_TITLE, layout="wide")
    st.title(APP_TITLE)
    st.caption(
        "Calculá el CAPEX ajustado por IPC o dólar BNA mediante carga manual o carga masiva."
    )

    try:
        ipc_df, dolar_bna_df = load_macro_series()
        max_ipc = ipc_df["fecha"].max() if not ipc_df.empty else None
        max_dolar = dolar_bna_df["fecha"].max() if not dolar_bna_df.empty else None
    except Exception as e:
        st.error(f"No se pudieron cargar las series internas: {e}")
        st.stop()

    with st.expander("Información de series utilizadas", expanded=False):
        col1, col2 = st.columns(2)
        with col1:
            st.write("**Serie IPC**")
            st.write(f"Última fecha disponible: {max_ipc.date() if pd.notna(max_ipc) else 'N/A'}")
            st.dataframe(ipc_df.tail(12), use_container_width=True)
        with col2:
            st.write("**Serie Dólar BNA**")
            st.write(f"Última fecha disponible: {max_dolar.date() if pd.notna(max_dolar) else 'N/A'}")
            st.dataframe(dolar_bna_df.tail(12), use_container_width=True)

    tab_manual, tab_masiva = st.tabs(["Carga manual", "Carga masiva"])

    # -----------------------------------------------------
    # TAB MANUAL
    # -----------------------------------------------------
    with tab_manual:
        col1, col2 = st.columns(2)

        with col1:
            capex_roi = st.number_input("CAPEX ROI", min_value=0.0, step=1000.0, format="%.2f")
            valuacion = st.date_input("Valuación CAPEX ROI", value=date.today())
            moneda = st.selectbox("Moneda", ["$", "USD"])

        with col2:
            inicio_obra = st.date_input("INICIO OBRA", value=date.today(), key="inicio_manual")
            fin_obra = st.date_input("FIN OBRA", value=date.today(), key="fin_manual")

        if st.button("Calcular", type="primary"):
            input_row = pd.DataFrame(
                [
                    {
                        "CAPEX ROI": capex_roi,
                        "Valuación CAPEX ROI": valuacion,
                        "INICIO OBRA": inicio_obra,
                        "FIN OBRA": fin_obra,
                        "Moneda": moneda,
                    }
                ]
            )

            result_df = procesar_archivo(input_row, ipc_df, dolar_bna_df)
            result_row = result_df.iloc[0]

            st.subheader("Resultado")
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("CAPEX ajustado", f"{result_row['CAPEX_AJUSTADO']:,.2f}" if pd.notna(result_row['CAPEX_AJUSTADO']) else "N/A")
            m2.metric("Factor ajuste", f"{result_row['FACTOR_AJUSTE']:.6f}" if pd.notna(result_row['FACTOR_AJUSTE']) else "N/A")
            m3.metric("Midpoint", str(result_row['MIDPOINT'].date()) if pd.notna(result_row['MIDPOINT']) else "N/A")
            m4.metric("Serie usada", result_row["SERIE_USADA"] if pd.notna(result_row["SERIE_USADA"]) else "N/A")

            if result_row["OBSERVACIONES"] != "OK":
                st.warning(f"Observación: {result_row['OBSERVACIONES']}")
            else:
                st.success("Cálculo realizado correctamente.")

            st.dataframe(result_df, use_container_width=True)

            if moneda == "$" and pd.notna(result_row["MIDPOINT"]):
                detail_df, _ = get_monthly_ipc_detail(valuacion, result_row["MIDPOINT"], ipc_df)
                with st.expander("Ver detalle mensual IPC", expanded=False):
                    st.dataframe(detail_df, use_container_width=True)
            else:
                detail_df = None

            st.download_button(
                "Descargar resultado en Excel",
                data=dataframe_to_excel_bytes(result_df, detail_df),
                file_name="capex_ajustado_manual.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

    # -----------------------------------------------------
    # TAB MASIVA
    # -----------------------------------------------------
    with tab_masiva:
        st.write("Subí un Excel con estas columnas:")
        st.code(", ".join(REQUIRED_COLUMNS))

        template_df = build_template()
        st.download_button(
            "Descargar template",
            data=dataframe_to_excel_bytes(template_df),
            file_name="template_capex_ajuste.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        uploaded_file = st.file_uploader("Subir archivo Excel", type=["xlsx"])

        if uploaded_file is not None:
            try:
                df_input = pd.read_excel(uploaded_file)
                st.write("Vista previa del archivo cargado")
                st.dataframe(df_input.head(20), use_container_width=True)

                if st.button("Procesar archivo", type="primary"):
                    output_df = procesar_archivo(df_input, ipc_df, dolar_bna_df)

                    st.subheader("Resultado")
                    total = len(output_df)
                    ok = int((output_df["OBSERVACIONES"] == "OK").sum())
                    errores = total - ok
                    c1, c2, c3 = st.columns(3)
                    c1.metric("Filas procesadas", total)
                    c2.metric("OK", ok)
                    c3.metric("Con observaciones", errores)

                    st.dataframe(output_df.head(50), use_container_width=True)

                    st.download_button(
                        "Descargar output",
                        data=dataframe_to_excel_bytes(output_df),
                        file_name="capex_ajustado_output.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )
            except Exception as e:
                st.error(f"No se pudo procesar el archivo: {e}")


if __name__ == "__main__":
    main()
