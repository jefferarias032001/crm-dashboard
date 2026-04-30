from __future__ import annotations

from datetime import datetime, timedelta
from pathlib import Path
from io import BytesIO
import json

import pandas as pd
from flask import Flask, render_template, request, send_file

app = Flask(__name__)

DEFAULT_EXCEL_PATH = "BD CRM 14042026.xlsx"
SHEET_NAME = "BD"

ESTADOS_COTIZACION_VALIDOS = {
    "Aprobada",
    "Enviada a Operaciones",
    "Enviada al cliente",
    "Enviar a operaciones",
    "Perdida",
    "Sin cobertura",
    "No aplica",
}


def find_column(df: pd.DataFrame, possible_names: list[str]) -> str | None:
    normalized = {str(col).strip().lower(): col for col in df.columns}

    for name in possible_names:
        key = name.strip().lower()
        if key in normalized:
            return normalized[key]

    return None


def clean_text(series: pd.Series) -> pd.Series:
    return (
        series.fillna("")
        .astype(str)
        .str.strip()
        .replace({"nan": "", "None": ""})
    )


def normalize_estado_cotizacion(series: pd.Series) -> pd.Series:
    s = clean_text(series).str.lower()

    replacements = {
        "aprobada": "Aprobada",
        "enviada a operaciones": "Enviada a Operaciones",
        "enviada a operacion": "Enviada a Operaciones",
        "enviada a operaciones tractocar": "Enviada a Operaciones",
        "enviada al cliente": "Enviada al cliente",
        "enviar a operaciones": "Enviar a operaciones",
        "perdida": "Perdida",
        "perdída": "Perdida",
        "sin cobertura": "Sin cobertura",
        "no aplica": "No aplica",
    }

    normalized = s.replace(replacements)
    normalized = normalized.where(normalized.isin(ESTADOS_COTIZACION_VALIDOS), "")

    return normalized


def to_datetime_safe(series: pd.Series) -> pd.Series:
    return pd.to_datetime(series, errors="coerce", dayfirst=True)


def to_numeric_safe(series: pd.Series) -> pd.Series:
    if series.dtype == object:
        cleaned = (
            series.astype(str)
            .str.replace(r"[^0-9,.-]", "", regex=True)
            .str.replace(".", "", regex=False)
            .str.replace(",", ".", regex=False)
        )
        return pd.to_numeric(cleaned, errors="coerce")

    return pd.to_numeric(series, errors="coerce")


def normalize_time_text(value: str) -> str:
    if pd.isna(value):
        return ""

    text = str(value).strip().lower()
    text = text.replace("a. m.", "AM").replace("p. m.", "PM")
    text = text.replace("a.m.", "AM").replace("p.m.", "PM")
    text = text.replace("am", "AM").replace("pm", "PM")

    return text.strip()


def combine_date_and_time(date_series: pd.Series, time_series: pd.Series) -> pd.Series:
    date_part = pd.to_datetime(date_series, errors="coerce", dayfirst=True)
    time_text = time_series.apply(normalize_time_text)

    combined = date_part.dt.strftime("%Y-%m-%d") + " " + time_text.astype(str)
    result = pd.to_datetime(combined, errors="coerce")

    fallback = pd.to_datetime(date_series, errors="coerce", dayfirst=True)

    return result.fillna(fallback)


def month_label(period) -> str:
    if pd.isna(period):
        return ""

    dt = period.to_timestamp()

    months = {
        1: "Enero",
        2: "Febrero",
        3: "Marzo",
        4: "Abril",
        5: "Mayo",
        6: "Junio",
        7: "Julio",
        8: "Agosto",
        9: "Septiembre",
        10: "Octubre",
        11: "Noviembre",
        12: "Diciembre",
    }

    return f"{months.get(dt.month, dt.month)} {dt.year}"


def month_short_label(period) -> str:
    if pd.isna(period):
        return ""

    dt = period.to_timestamp()

    months = {
        1: "Ene",
        2: "Feb",
        3: "Mar",
        4: "Abr",
        5: "May",
        6: "Jun",
        7: "Jul",
        8: "Ago",
        9: "Sep",
        10: "Oct",
        11: "Nov",
        12: "Dic",
    }

    return months.get(dt.month, "")


def hours_to_text(value: float | int | None) -> str:
    if value is None or pd.isna(value):
        return ""

    value = float(value)

    if value < 24:
        return f"{value:.1f} h"

    days = value / 24
    return f"{days:.1f} días"


def days_to_text(value: float | int | None) -> str:
    if value is None or pd.isna(value):
        return ""

    return f"{float(value):.1f} días"


def get_filter_options(df: pd.DataFrame) -> dict:
    def options_for(col: str) -> list[str]:
        if col not in df.columns:
            return []

        vals = df[col].dropna().astype(str).str.strip()
        vals = vals[vals != ""]

        return sorted(vals.unique().tolist())

    options = {
        "canal": options_for("canal"),
        "tipo_cliente": options_for("tipo_cliente"),
        "tipo_solicitud": options_for("tipo_solicitud"),
        "tipo_operacion": options_for("tipo_operacion"),
        "estado": options_for("estado"),
        "estado_cotizacion": options_for("estado_cotizacion"),
        "responsable_cotizacion": options_for("responsable_cotizacion"),
        "mes": [],
    }

    if "periodo" in df.columns:
        periods = sorted(df["periodo"].dropna().unique())
        options["mes"] = [str(p) for p in periods]

    return options


def apply_filters(df: pd.DataFrame, filters: dict) -> pd.DataFrame:
    filtered = df.copy()

    column_map = {
        "canal": "canal",
        "tipo_cliente": "tipo_cliente",
        "tipo_solicitud": "tipo_solicitud",
        "tipo_operacion": "tipo_operacion",
        "estado": "estado",
        "estado_cotizacion": "estado_cotizacion",
        "responsable_cotizacion": "responsable_cotizacion",
    }

    for filter_key, col_name in column_map.items():
        value = filters.get(filter_key, "").strip()

        if value and col_name in filtered.columns:
            filtered = filtered[filtered[col_name] == value]

    mes = filters.get("mes", "").strip()

    if mes and "periodo" in filtered.columns:
        filtered = filtered[filtered["periodo"].astype(str) == mes]

    return filtered


def parse_filters_from_request() -> dict:
    return {
        "canal": request.args.get("canal", ""),
        "tipo_cliente": request.args.get("tipo_cliente", ""),
        "tipo_solicitud": request.args.get("tipo_solicitud", ""),
        "tipo_operacion": request.args.get("tipo_operacion", ""),
        "estado": request.args.get("estado", ""),
        "estado_cotizacion": request.args.get("estado_cotizacion", ""),
        "responsable_cotizacion": request.args.get("responsable_cotizacion", ""),
        "mes": request.args.get("mes", ""),
    }


def load_data() -> pd.DataFrame:
    excel_path = Path(DEFAULT_EXCEL_PATH)

    if not excel_path.exists():
        raise FileNotFoundError(
            f"No se encontró el archivo Excel en: {excel_path}. "
            "Pon el archivo dentro de la carpeta del proyecto o ajusta DEFAULT_EXCEL_PATH."
        )

    df = pd.read_excel(
        excel_path,
        sheet_name=SHEET_NAME,
        engine="openpyxl",
    )

    df.columns = [str(c).strip() for c in df.columns]

    col_canal = find_column(df, ["Canal", "CANAL", "Medio", "MEDIO"])
    col_fecha_ingreso_sol = find_column(df, ["Fecha Ingreso Solicitud", "FECHA INGRESO SOLICITUD"])
    col_hora_ingreso_sol = find_column(df, ["Hora ingreso Solicitud", "HORA INGRESO SOLICITUD"])
    col_fecha_respuesta_sol = find_column(df, ["Fecha de respuesta Solicitud", "FECHA DE RESPUESTA SOLICITUD"])
    col_hora_respuesta_sol = find_column(df, ["Hora de respuesta Solicitud", "HORA DE RESPUESTA SOLICITUD"])
    col_fecha_ingreso_cot = find_column(df, ["Fecha Ingreso Cotizacion", "Fecha Ingreso Cotización"])
    col_fecha_cierre_cot = find_column(df, ["Fecha Cierre Cotizacion", "Fecha Cierre Cotización"])
    col_cliente = find_column(df, ["Cliente", "CLIENTE"])
    col_tipo_cliente = find_column(df, ["Tipo de Cliente", "TIPO DE CLIENTE"])
    col_contacto = find_column(df, ["Contacto", "CONTACTO"])
    col_celular = find_column(df, ["Celular", "CELULAR"])
    col_tipo_solicitud = find_column(df, ["Tipo de Solicitud", "TIPO DE SOLICITUD"])
    col_tipo_operacion = find_column(df, ["Tipo de Operación", "TIPO DE OPERACIÓN"])
    col_estado = find_column(df, ["Estado", "ESTADO"])
    col_estado_cot = find_column(df, ["Estado Cotizacion", "Estado Cotización"])
    col_responsable_cot = find_column(df, ["Responsable Cotizacion", "Responsable Cotización"])
    col_accion_seguir = find_column(df, ["Accion a Seguir", "ACCIÓN A SEGUIR", "ACCION A SEGUIR"])
    col_valor = find_column(df, ["Valor Estimado Negocio", "VALOR ESTIMADO NEGOCIO", "VALOR ESTIMADO"])
    col_observacion = find_column(df, ["Observación", "OBSERVACIÓN", "OBSERVACION"])

    rename_map = {}

    if col_canal:
        rename_map[col_canal] = "canal"
    if col_fecha_ingreso_sol:
        rename_map[col_fecha_ingreso_sol] = "fecha_ingreso_solicitud"
    if col_hora_ingreso_sol:
        rename_map[col_hora_ingreso_sol] = "hora_ingreso_solicitud"
    if col_fecha_respuesta_sol:
        rename_map[col_fecha_respuesta_sol] = "fecha_respuesta_solicitud"
    if col_hora_respuesta_sol:
        rename_map[col_hora_respuesta_sol] = "hora_respuesta_solicitud"
    if col_fecha_ingreso_cot:
        rename_map[col_fecha_ingreso_cot] = "fecha_ingreso_cotizacion"
    if col_fecha_cierre_cot:
        rename_map[col_fecha_cierre_cot] = "fecha_cierre_cotizacion"
    if col_cliente:
        rename_map[col_cliente] = "cliente"
    if col_tipo_cliente:
        rename_map[col_tipo_cliente] = "tipo_cliente"
    if col_contacto:
        rename_map[col_contacto] = "contacto"
    if col_celular:
        rename_map[col_celular] = "celular"
    if col_tipo_solicitud:
        rename_map[col_tipo_solicitud] = "tipo_solicitud"
    if col_tipo_operacion:
        rename_map[col_tipo_operacion] = "tipo_operacion"
    if col_estado:
        rename_map[col_estado] = "estado"
    if col_estado_cot:
        rename_map[col_estado_cot] = "estado_cotizacion"
    if col_responsable_cot:
        rename_map[col_responsable_cot] = "responsable_cotizacion"
    if col_accion_seguir:
        rename_map[col_accion_seguir] = "accion_a_seguir"
    if col_valor:
        rename_map[col_valor] = "valor_estimado_negocio"
    if col_observacion:
        rename_map[col_observacion] = "observacion"

    df = df.rename(columns=rename_map)

    if "fecha_ingreso_solicitud" not in df.columns:
        raise ValueError("Falta la columna Fecha Ingreso Solicitud.")

    for text_col in [
        "canal",
        "cliente",
        "tipo_cliente",
        "contacto",
        "celular",
        "tipo_solicitud",
        "tipo_operacion",
        "estado",
        "responsable_cotizacion",
        "accion_a_seguir",
        "observacion",
    ]:
        if text_col in df.columns:
            df[text_col] = clean_text(df[text_col])

    if "estado_cotizacion" in df.columns:
        df["estado_cotizacion"] = normalize_estado_cotizacion(df["estado_cotizacion"])

    df["fecha_ingreso_solicitud_dt"] = combine_date_and_time(
        df["fecha_ingreso_solicitud"],
        df["hora_ingreso_solicitud"] if "hora_ingreso_solicitud" in df.columns else pd.Series("", index=df.index),
    )

    if "fecha_respuesta_solicitud" in df.columns:
        df["fecha_respuesta_solicitud_dt"] = combine_date_and_time(
            df["fecha_respuesta_solicitud"],
            df["hora_respuesta_solicitud"] if "hora_respuesta_solicitud" in df.columns else pd.Series("", index=df.index),
        )
    else:
        df["fecha_respuesta_solicitud_dt"] = pd.NaT

    if "fecha_ingreso_cotizacion" in df.columns:
        df["fecha_ingreso_cotizacion_dt"] = to_datetime_safe(df["fecha_ingreso_cotizacion"])
    else:
        df["fecha_ingreso_cotizacion_dt"] = pd.NaT

    if "fecha_cierre_cotizacion" in df.columns:
        df["fecha_cierre_cotizacion_dt"] = to_datetime_safe(df["fecha_cierre_cotizacion"])
    else:
        df["fecha_cierre_cotizacion_dt"] = pd.NaT

    if "valor_estimado_negocio" in df.columns:
        df["valor_estimado_negocio"] = to_numeric_safe(df["valor_estimado_negocio"]).fillna(0)
    else:
        df["valor_estimado_negocio"] = 0.0

    df = df[df["fecha_ingreso_solicitud_dt"].notna()].copy()

    df["periodo"] = df["fecha_ingreso_solicitud_dt"].dt.to_period("M")

    iso = df["fecha_ingreso_solicitud_dt"].dt.isocalendar()
    df["anio"] = iso.year.astype(int)
    df["semana_num"] = iso.week.astype(int)
    df["semana_label"] = "Sem " + df["semana_num"].astype(str)
    df["mes_semana_label"] = df["periodo"].apply(month_short_label)

    month_df = pd.DataFrame(
        {
            "periodo": df["periodo"],
            "semana_num": df["semana_num"],
        }
    ).drop_duplicates().sort_values(["periodo", "semana_num"])

    month_df["orden_semana_mes"] = month_df.groupby("periodo").cumcount() + 1

    df = df.merge(
        month_df,
        on=["periodo", "semana_num"],
        how="left",
    )

    df["horas_respuesta_solicitud"] = (
        (df["fecha_respuesta_solicitud_dt"] - df["fecha_ingreso_solicitud_dt"]).dt.total_seconds() / 3600
    )

    df.loc[df["horas_respuesta_solicitud"] < 0, "horas_respuesta_solicitud"] = pd.NA

    df["dias_cierre_cotizacion"] = (
        (df["fecha_cierre_cotizacion_dt"] - df["fecha_ingreso_cotizacion_dt"]).dt.total_seconds() / 86400
    )

    df.loc[df["dias_cierre_cotizacion"] < 0, "dias_cierre_cotizacion"] = pd.NA

    df["tiene_respuesta_solicitud"] = df["fecha_respuesta_solicitud_dt"].notna()
    df["tiene_cotizacion"] = df["fecha_ingreso_cotizacion_dt"].notna()
    df["cotizacion_cerrada"] = df["fecha_cierre_cotizacion_dt"].notna()

    if "estado" in df.columns and "estado_cotizacion" in df.columns:
        estado_lower = clean_text(df["estado"]).str.lower()
        df["estado_cotizacion_valido"] = (
            (estado_lower == "cotizacion")
            & (df["estado_cotizacion"].isin(ESTADOS_COTIZACION_VALIDOS))
        )
    else:
        df["estado_cotizacion_valido"] = False

    return df


def build_dashboard_data(df_filtered: pd.DataFrame, df_unfiltered_for_trend: pd.DataFrame) -> dict:
    if df_filtered.empty and df_unfiltered_for_trend.empty:
        return {}

    working_df = df_filtered.copy() if not df_filtered.empty else df_unfiltered_for_trend.copy()
    cot_df = working_df[working_df["estado_cotizacion_valido"]].copy()

    now_col = pd.Timestamp(datetime.utcnow() - timedelta(hours=5))

    current_period = working_df["periodo"].max()
    previous_period = current_period - 1 if pd.notna(current_period) else None

    current_df = (
        working_df[working_df["periodo"] == current_period].copy()
        if pd.notna(current_period)
        else pd.DataFrame()
    )

    previous_df = (
        working_df[working_df["periodo"] == previous_period].copy()
        if previous_period is not None
        else pd.DataFrame()
    )

    total_solicitudes = int(len(working_df))

    total_cotizaciones = (
        int(working_df["tiene_cotizacion"].sum())
        if "tiene_cotizacion" in working_df.columns
        else 0
    )

    tasa_cotizacion = (
        round((total_cotizaciones / total_solicitudes) * 100, 1)
        if total_solicitudes
        else 0
    )

    valor_total = (
        float(working_df["valor_estimado_negocio"].fillna(0).sum())
        if "valor_estimado_negocio" in working_df.columns
        else 0
    )

    tiempos_respuesta_validos = working_df[
        (working_df["horas_respuesta_solicitud"].notna())
        & (working_df["horas_respuesta_solicitud"] >= 0)
    ]

    tiempo_respuesta_promedio_horas = (
        float(tiempos_respuesta_validos["horas_respuesta_solicitud"].mean())
        if len(tiempos_respuesta_validos) > 0
        else 0
    )

    tiempos_cierre_validos = working_df[
        (working_df["fecha_ingreso_cotizacion_dt"].notna())
        & (working_df["fecha_cierre_cotizacion_dt"].notna())
        & (working_df["dias_cierre_cotizacion"].notna())
        & (working_df["dias_cierre_cotizacion"] >= 0)
    ]

    tiempo_cierre_promedio_dias = (
        float(tiempos_cierre_validos["dias_cierre_cotizacion"].mean())
        if len(tiempos_cierre_validos) > 0
        else 0
    )

    pendientes_operaciones_df = working_df[
        (working_df["estado_cotizacion"] == "Enviada a Operaciones")
        & (working_df["fecha_cierre_cotizacion_dt"].notna())
    ].copy()

    if not pendientes_operaciones_df.empty:
        pendientes_operaciones_df["dias_espera_operaciones"] = (
            (now_col - pendientes_operaciones_df["fecha_cierre_cotizacion_dt"]).dt.total_seconds() / 86400
        )

        pendientes_operaciones_df = pendientes_operaciones_df[
            pendientes_operaciones_df["dias_espera_operaciones"] >= 0
        ]
    else:
        pendientes_operaciones_df["dias_espera_operaciones"] = pd.Series(dtype=float)

    promedio_espera_operaciones_dias = (
        float(pendientes_operaciones_df["dias_espera_operaciones"].mean())
        if len(pendientes_operaciones_df) > 0
        else 0
    )

    pendientes_operaciones_count = int(len(pendientes_operaciones_df))

    pendientes_operaciones_criticas = (
        int((pendientes_operaciones_df["dias_espera_operaciones"] > 3).sum())
        if len(pendientes_operaciones_df) > 0
        else 0
    )

    current_solicitudes = int(len(current_df))

    current_cotizaciones = (
        int(current_df["tiene_cotizacion"].sum())
        if not current_df.empty
        else 0
    )

    current_tasa_cotizacion = (
        round((current_cotizaciones / current_solicitudes) * 100, 1)
        if current_solicitudes
        else 0
    )

    previous_solicitudes = int(len(previous_df))

    previous_cotizaciones = (
        int(previous_df["tiene_cotizacion"].sum())
        if not previous_df.empty
        else 0
    )

    previous_tasa_cotizacion = (
        round((previous_cotizaciones / previous_solicitudes) * 100, 1)
        if previous_solicitudes
        else 0
    )

    negocios_ganados = (
        int((cot_df["estado_cotizacion"] == "Aprobada").sum())
        if "estado_cotizacion" in cot_df.columns
        else 0
    )

    negocios_perdidos = (
        int((cot_df["estado_cotizacion"] == "Perdida").sum())
        if "estado_cotizacion" in cot_df.columns
        else 0
    )

    total_con_resultado = negocios_ganados + negocios_perdidos

    tasa_exito = (
        round((negocios_ganados / total_con_resultado) * 100, 1)
        if total_con_resultado > 0
        else 0
    )

    monthly = (
        df_unfiltered_for_trend.groupby("periodo")
        .agg(
            solicitudes=("fecha_ingreso_solicitud_dt", "count"),
            cotizaciones=("tiene_cotizacion", "sum"),
            valor=("valor_estimado_negocio", "sum"),
        )
        .reset_index()
        .sort_values("periodo")
    ) if not df_unfiltered_for_trend.empty else pd.DataFrame(
        columns=["periodo", "solicitudes", "cotizaciones", "valor"]
    )

    if not monthly.empty:
        monthly["mes"] = monthly["periodo"].apply(month_label)
        monthly_data = monthly[
            ["mes", "solicitudes", "cotizaciones", "valor"]
        ].to_dict(orient="records")
    else:
        monthly_data = []

    weekly = (
        df_unfiltered_for_trend.groupby(
            ["periodo", "semana_num", "semana_label", "mes_semana_label"]
        )
        .agg(
            solicitudes=("fecha_ingreso_solicitud_dt", "count"),
            cotizaciones=("tiene_cotizacion", "sum"),
            valor=("valor_estimado_negocio", "sum"),
        )
        .reset_index()
        .sort_values(["periodo", "semana_num"])
    ) if not df_unfiltered_for_trend.empty else pd.DataFrame(
        columns=[
            "periodo",
            "semana_num",
            "semana_label",
            "mes_semana_label",
            "solicitudes",
            "cotizaciones",
            "valor",
        ]
    )

    if not weekly.empty:
        weekly = weekly.tail(12).copy()
        weekly_data = weekly[
            ["semana_label", "mes_semana_label", "solicitudes", "cotizaciones", "valor"]
        ].rename(
            columns={
                "semana_label": "semana",
                "mes_semana_label": "mes",
            }
        ).to_dict(orient="records")
    else:
        weekly_data = []

    trend_summary = {}

    if len(monthly) >= 2:
        last_row = monthly.iloc[-1]
        prev_row = monthly.iloc[-2]

        delta_sol = int(last_row["solicitudes"] - prev_row["solicitudes"])
        delta_cot = int(last_row["cotizaciones"] - prev_row["cotizaciones"])

        pct_sol = (
            round((delta_sol / prev_row["solicitudes"]) * 100, 1)
            if prev_row["solicitudes"] > 0
            else 0
        )

        pct_cot = (
            round((delta_cot / prev_row["cotizaciones"]) * 100, 1)
            if prev_row["cotizaciones"] > 0
            else 0
        )

        trend_summary = {
            "current_month": last_row["mes"],
            "previous_month": prev_row["mes"],
            "delta_solicitudes": delta_sol,
            "delta_cotizaciones": delta_cot,
            "pct_solicitudes": pct_sol,
            "pct_cotizaciones": pct_cot,
        }

    performance_semanal = {
        "delta_solicitudes": 0,
        "delta_cotizaciones": 0,
        "pct_solicitudes": 0,
        "pct_cotizaciones": 0,
    }

    if pd.notna(current_period):
        current_cut_weeks = sorted(
            current_df["orden_semana_mes"].dropna().astype(int).unique().tolist()
        )

        if previous_period is not None and current_cut_weeks:
            previous_cut_df = previous_df[
                previous_df["orden_semana_mes"].isin(current_cut_weeks)
            ].copy()

            current_cut_df = current_df[
                current_df["orden_semana_mes"].isin(current_cut_weeks)
            ].copy()

            cut_solicitudes = int(len(current_cut_df))
            prev_cut_solicitudes = int(len(previous_cut_df))

            cut_cotizaciones = (
                int(current_cut_df["tiene_cotizacion"].sum())
                if not current_cut_df.empty
                else 0
            )

            prev_cut_cotizaciones = (
                int(previous_cut_df["tiene_cotizacion"].sum())
                if not previous_cut_df.empty
                else 0
            )

            delta_sol_cut = cut_solicitudes - prev_cut_solicitudes
            delta_cot_cut = cut_cotizaciones - prev_cut_cotizaciones

            pct_sol_cut = (
                round((delta_sol_cut / prev_cut_solicitudes) * 100, 1)
                if prev_cut_solicitudes > 0
                else 0
            )

            pct_cot_cut = (
                round((delta_cot_cut / prev_cut_cotizaciones) * 100, 1)
                if prev_cut_cotizaciones > 0
                else 0
            )

            performance_semanal = {
                "delta_solicitudes": int(delta_sol_cut),
                "delta_cotizaciones": int(delta_cot_cut),
                "pct_solicitudes": pct_sol_cut,
                "pct_cotizaciones": pct_cot_cut,
            }

    canal_data = []

    if "canal" in working_df.columns:
        canal = (
            working_df.groupby("canal")
            .size()
            .reset_index(name="cantidad")
            .sort_values("cantidad", ascending=False)
        )

        canal = canal[
            (canal["canal"] != "")
            & (canal["cantidad"] > 0)
        ]

        total_canal = canal["cantidad"].sum()

        for _, row in canal.iterrows():
            canal_data.append(
                {
                    "canal": row["canal"],
                    "cantidad": int(row["cantidad"]),
                    "porcentaje": round((row["cantidad"] / total_canal) * 100, 1)
                    if total_canal
                    else 0,
                }
            )

    tipo_cliente_data = []

    if "tipo_cliente" in working_df.columns:
        tipo_cliente = (
            working_df.groupby("tipo_cliente")
            .size()
            .reset_index(name="cantidad")
            .sort_values("cantidad", ascending=False)
        )

        tipo_cliente = tipo_cliente[
            (tipo_cliente["tipo_cliente"] != "")
            & (tipo_cliente["cantidad"] > 0)
        ]

        tipo_cliente_data = tipo_cliente.to_dict(orient="records")

    tipo_operacion_data = []

    if "tipo_operacion" in working_df.columns and "valor_estimado_negocio" in working_df.columns:
        tipo_operacion = (
            working_df.groupby("tipo_operacion")
            .agg(
                cantidad=("tipo_operacion", "count"),
                valor_total=("valor_estimado_negocio", "sum"),
            )
            .reset_index()
            .sort_values("valor_total", ascending=False)
        )

        tipo_operacion = tipo_operacion[
            (tipo_operacion["tipo_operacion"] != "")
            & (tipo_operacion["valor_total"] > 0)
        ]

        total_valor_operacion = float(tipo_operacion["valor_total"].sum())

        tipo_operacion["porcentaje_participacion"] = (
            (tipo_operacion["valor_total"] / total_valor_operacion) * 100
        ).round(1) if total_valor_operacion > 0 else 0

        tipo_operacion_data = tipo_operacion.to_dict(orient="records")

    tipo_operacion_count_data = []

    if "tipo_operacion" in df_unfiltered_for_trend.columns:
        tipo_op_count = (
            df_unfiltered_for_trend.groupby("tipo_operacion")
            .size()
            .reset_index(name="cantidad")
            .sort_values("cantidad", ascending=False)
        )

        tipo_op_count = tipo_op_count[
            (tipo_op_count["tipo_operacion"] != "")
            & (tipo_op_count["cantidad"] > 0)
        ]

        tipo_operacion_count_data = tipo_op_count.to_dict(orient="records")

    tipo_operacion_mensual = []

    if "tipo_operacion" in df_unfiltered_for_trend.columns:
        tipo_op_month = (
            df_unfiltered_for_trend.groupby(["periodo", "tipo_operacion"])
            .size()
            .reset_index(name="cantidad")
            .sort_values(["periodo", "cantidad"], ascending=[True, False])
        )

        tipo_op_month = tipo_op_month[
            (tipo_op_month["tipo_operacion"] != "")
            & (tipo_op_month["cantidad"] > 0)
        ].copy()

        tipo_op_month["mes"] = tipo_op_month["periodo"].apply(month_label)
        tipo_op_month["periodo_texto"] = tipo_op_month["periodo"].astype(str)

        tipo_operacion_mensual = tipo_op_month[
            [
                "periodo_texto",
                "mes",
                "tipo_operacion",
                "cantidad",
            ]
        ].to_dict(orient="records")

    valor_cliente_data = []

    if "cliente" in working_df.columns and "valor_estimado_negocio" in working_df.columns:
        cliente = (
            working_df.groupby("cliente")
            .agg(
                cantidad=("cliente", "count"),
                valor_total=("valor_estimado_negocio", "sum"),
            )
            .reset_index()
            .sort_values("valor_total", ascending=False)
        )

        cliente = cliente[
            (cliente["cliente"] != "")
            & (cliente["valor_total"] > 0)
        ]

        total_valor_clientes = float(cliente["valor_total"].sum())

        cliente["porcentaje_participacion"] = (
            (cliente["valor_total"] / total_valor_clientes) * 100
        ).round(1) if total_valor_clientes > 0 else 0

        valor_cliente_data = cliente.head(10).to_dict(orient="records")

    estado_data = []

    if "estado" in working_df.columns:
        estado = (
            working_df.groupby("estado")
            .size()
            .reset_index(name="cantidad")
            .sort_values("cantidad", ascending=False)
        )

        estado = estado[
            (estado["estado"] != "")
            & (estado["cantidad"] > 0)
        ]

        estado_data = estado.to_dict(orient="records")

    estado_cotizacion_data = []

    if "estado_cotizacion" in cot_df.columns:
        estado_cot = (
            cot_df.groupby("estado_cotizacion")
            .size()
            .reset_index(name="cantidad")
            .sort_values("cantidad", ascending=False)
        )

        estado_cot = estado_cot[
            (estado_cot["estado_cotizacion"] != "")
            & (estado_cot["cantidad"] > 0)
        ]

        estado_cotizacion_data = estado_cot.to_dict(orient="records")

    responsable_data = []

    if "responsable_cotizacion" in working_df.columns:
        base_responsable = working_df[
            working_df["responsable_cotizacion"].fillna("").str.strip() != ""
        ].copy()

        responsable = (
            base_responsable.groupby("responsable_cotizacion")
            .agg(
                cantidad=("responsable_cotizacion", "count"),
                valor_total=("valor_estimado_negocio", "sum"),
                tiempo_respuesta_solicitud=("horas_respuesta_solicitud", "mean"),
                tiempo_respuesta_cotizacion=("dias_cierre_cotizacion", "mean"),
            )
            .reset_index()
        )

        espera_operaciones = (
            pendientes_operaciones_df[
                pendientes_operaciones_df["responsable_cotizacion"].fillna("").str.strip()
                != ""
            ]
            .groupby("responsable_cotizacion")
            .agg(
                espera_operaciones=("dias_espera_operaciones", "mean")
            )
            .reset_index()
        )

        responsable = responsable.merge(
            espera_operaciones,
            on="responsable_cotizacion",
            how="left",
        )

        responsable = responsable[
            (responsable["responsable_cotizacion"] != "")
            & (
                (responsable["valor_total"] > 0)
                | (responsable["cantidad"] > 0)
            )
        ].copy()

        total_valor_responsable = float(responsable["valor_total"].sum())

        responsable["porcentaje_participacion"] = (
            (responsable["valor_total"] / total_valor_responsable) * 100
        ).round(1) if total_valor_responsable > 0 else 0

        responsable["tiempo_respuesta_solicitud"] = responsable["tiempo_respuesta_solicitud"].round(1)
        responsable["tiempo_respuesta_cotizacion"] = responsable["tiempo_respuesta_cotizacion"].round(1)
        responsable["espera_operaciones"] = responsable["espera_operaciones"].round(1)

        responsable = responsable.sort_values(
            ["valor_total", "cantidad"],
            ascending=[False, False],
        )

        responsable_data = responsable.to_dict(orient="records")

    aging_operaciones_data = []

    if len(pendientes_operaciones_df) > 0:
        pendientes_operaciones_df["rango_aging"] = pd.cut(
            pendientes_operaciones_df["dias_espera_operaciones"],
            bins=[-1, 1, 2, 3, 9999],
            labels=["0-1 días", "1-2 días", "2-3 días", "+3 días"],
        )

        aging = (
            pendientes_operaciones_df.groupby("rango_aging")
            .size()
            .reset_index(name="cantidad")
        )

        aging["rango_aging"] = aging["rango_aging"].astype(str)
        aging_operaciones_data = aging.to_dict(orient="records")

    detalle = working_df.copy()

    detalle["fecha_ingreso_solicitud_fmt"] = detalle["fecha_ingreso_solicitud_dt"].dt.strftime(
        "%Y-%m-%d %H:%M"
    )

    detalle_cols = [
        c
        for c in [
            "fecha_ingreso_solicitud_fmt",
            "canal",
            "cliente",
            "tipo_cliente",
            "contacto",
            "tipo_solicitud",
            "tipo_operacion",
            "estado",
            "estado_cotizacion",
            "responsable_cotizacion",
            "accion_a_seguir",
            "valor_estimado_negocio",
            "observacion",
        ]
        if c in detalle.columns
    ]

    detalle_data = (
        detalle[detalle_cols]
        .sort_values("fecha_ingreso_solicitud_fmt", ascending=False)
        .head(50)
        .to_dict(orient="records")
    )

    insights = []

    if trend_summary:
        insights.append(
            f"En {trend_summary['current_month']}, las solicitudes variaron "
            f"{trend_summary['pct_solicitudes']}% y las cotizaciones "
            f"{trend_summary['pct_cotizaciones']}% frente a "
            f"{trend_summary['previous_month']}."
        )

    insights.append(
        f"Al corte de semanas equivalentes del mes, las solicitudes variaron "
        f"{performance_semanal['pct_solicitudes']}% y las cotizaciones "
        f"{performance_semanal['pct_cotizaciones']}% frente al mes anterior."
    )

    if tiempo_cierre_promedio_dias > 0:
        insights.append(
            f"El tiempo promedio de cierre de cotización es de "
            f"{tiempo_cierre_promedio_dias:.1f} días."
        )

    if pendientes_operaciones_count > 0:
        insights.append(
            f"Hay {pendientes_operaciones_count} cotizaciones actualmente en estado "
            f"'Enviada a Operaciones', con una espera promedio de "
            f"{promedio_espera_operaciones_dias:.1f} días desde su envío a operaciones."
        )

    if pendientes_operaciones_criticas > 0:
        insights.append(
            f"Se identifican {pendientes_operaciones_criticas} cotizaciones críticas "
            f"con más de 3 días en operaciones sin cambio de estado."
        )

    if canal_data:
        top_canal = canal_data[0]
        insights.append(
            f"El canal con mayor ingreso de solicitudes es {top_canal['canal']}, "
            f"con {top_canal['cantidad']} registros "
            f"({top_canal['porcentaje']}% del total filtrado)."
        )

    if tipo_operacion_data:
        top_operacion = tipo_operacion_data[0]
        insights.append(
            f"El tipo de operación con mayor valor estimado es "
            f"{top_operacion['tipo_operacion']}, con "
            f"{top_operacion['porcentaje_participacion']}% del valor del negocio filtrado."
        )

    chart_labels = [row["mes"] for row in monthly_data]
    chart_solicitudes = [int(row["solicitudes"]) for row in monthly_data]
    chart_cotizaciones = [int(row["cotizaciones"]) for row in monthly_data]

    weekly_labels = [row["semana"] for row in weekly_data]
    weekly_months = [row["mes"] for row in weekly_data]
    weekly_solicitudes = [int(row["solicitudes"]) for row in weekly_data]
    weekly_cotizaciones = [int(row["cotizaciones"]) for row in weekly_data]

    chart_tipo_operacion_labels = [
        row["tipo_operacion"] for row in tipo_operacion_data
    ]

    chart_tipo_operacion_valores = [
        float(row["valor_total"]) for row in tipo_operacion_data
    ]

    chart_cliente_labels = [row["cliente"] for row in valor_cliente_data]
    chart_cliente_valores = [float(row["valor_total"]) for row in valor_cliente_data]

    chart_canal_labels = [row["canal"] for row in canal_data]
    chart_canal_valores = [int(row["cantidad"]) for row in canal_data]

    chart_tipo_cliente_labels = [row["tipo_cliente"] for row in tipo_cliente_data]
    chart_tipo_cliente_valores = [int(row["cantidad"]) for row in tipo_cliente_data]

    chart_estado_cot_labels = [
        row["estado_cotizacion"] for row in estado_cotizacion_data
    ]

    chart_estado_cot_valores = [
        int(row["cantidad"]) for row in estado_cotizacion_data
    ]

    chart_aging_labels = [
        row["rango_aging"] for row in aging_operaciones_data
    ]

    chart_aging_valores = [
        int(row["cantidad"]) for row in aging_operaciones_data
    ]

    chart_tipo_operacion_count_labels = [
        row["tipo_operacion"] for row in tipo_operacion_count_data
    ]

    chart_tipo_operacion_count_valores = [
        int(row["cantidad"]) for row in tipo_operacion_count_data
    ]

    chart_tipo_operacion_mensual_json = json.dumps(
        tipo_operacion_mensual,
        ensure_ascii=False,
    )

    return {
        "generated_at": now_col.strftime("%Y-%m-%d %H:%M"),
        "current_period_label": month_label(current_period)
        if pd.notna(current_period)
        else "",
        "previous_period_label": month_label(previous_period)
        if previous_period is not None
        else "",
        "total_solicitudes": total_solicitudes,
        "total_cotizaciones": total_cotizaciones,
        "tasa_cotizacion": tasa_cotizacion,
        "tiempo_respuesta_promedio_horas": tiempo_respuesta_promedio_horas,
        "tiempo_respuesta_promedio_txt": hours_to_text(
            tiempo_respuesta_promedio_horas
        ),
        "tiempo_cierre_promedio_dias": tiempo_cierre_promedio_dias,
        "tiempo_cierre_promedio_txt": days_to_text(
            tiempo_cierre_promedio_dias
        ),
        "valor_total": valor_total,
        "valor_total_mm": valor_total / 1_000_000,
        "valor_total_b": valor_total / 1_000_000_000,
        "promedio_espera_operaciones_dias": promedio_espera_operaciones_dias,
        "promedio_espera_operaciones_txt": days_to_text(
            promedio_espera_operaciones_dias
        ),
        "pendientes_operaciones_count": pendientes_operaciones_count,
        "pendientes_operaciones_criticas": pendientes_operaciones_criticas,
        "negocios_ganados": negocios_ganados,
        "negocios_perdidos": negocios_perdidos,
        "tasa_exito": tasa_exito,
        "current_solicitudes": current_solicitudes,
        "current_cotizaciones": current_cotizaciones,
        "current_tasa_cotizacion": current_tasa_cotizacion,
        "previous_solicitudes": previous_solicitudes,
        "previous_cotizaciones": previous_cotizaciones,
        "previous_tasa_cotizacion": previous_tasa_cotizacion,
        "monthly_data": monthly_data,
        "weekly_data": weekly_data,
        "trend_summary": trend_summary,
        "performance_semanal": performance_semanal,
        "canal_data": canal_data,
        "tipo_cliente_data": tipo_cliente_data,
        "tipo_operacion_data": tipo_operacion_data,
        "tipo_operacion_count_data": tipo_operacion_count_data,
        "tipo_operacion_mensual": tipo_operacion_mensual,
        "estado_data": estado_data,
        "estado_cotizacion_data": estado_cotizacion_data,
        "valor_cliente_data": valor_cliente_data,
        "responsable_data": responsable_data,
        "aging_operaciones_data": aging_operaciones_data,
        "detalle_data": detalle_data,
        "insights": insights,
        "chart_labels_json": json.dumps(chart_labels, ensure_ascii=False),
        "chart_solicitudes_json": json.dumps(chart_solicitudes),
        "chart_cotizaciones_json": json.dumps(chart_cotizaciones),
        "chart_weekly_labels_json": json.dumps(weekly_labels, ensure_ascii=False),
        "chart_weekly_months_json": json.dumps(weekly_months, ensure_ascii=False),
        "chart_weekly_solicitudes_json": json.dumps(weekly_solicitudes),
        "chart_weekly_cotizaciones_json": json.dumps(weekly_cotizaciones),
        "chart_tipo_operacion_labels_json": json.dumps(
            chart_tipo_operacion_labels,
            ensure_ascii=False,
        ),
        "chart_tipo_operacion_valores_json": json.dumps(
            chart_tipo_operacion_valores
        ),
        "chart_cliente_labels_json": json.dumps(
            chart_cliente_labels,
            ensure_ascii=False,
        ),
        "chart_cliente_valores_json": json.dumps(chart_cliente_valores),
        "chart_canal_labels_json": json.dumps(
            chart_canal_labels,
            ensure_ascii=False,
        ),
        "chart_canal_valores_json": json.dumps(chart_canal_valores),
        "chart_tipo_cliente_labels_json": json.dumps(
            chart_tipo_cliente_labels,
            ensure_ascii=False,
        ),
        "chart_tipo_cliente_valores_json": json.dumps(
            chart_tipo_cliente_valores
        ),
        "chart_estado_cot_labels_json": json.dumps(
            chart_estado_cot_labels,
            ensure_ascii=False,
        ),
        "chart_estado_cot_valores_json": json.dumps(
            chart_estado_cot_valores
        ),
        "chart_aging_labels_json": json.dumps(
            chart_aging_labels,
            ensure_ascii=False,
        ),
        "chart_aging_valores_json": json.dumps(chart_aging_valores),
        "chart_tipo_operacion_count_labels_json": json.dumps(
            chart_tipo_operacion_count_labels,
            ensure_ascii=False,
        ),
        "chart_tipo_operacion_count_valores_json": json.dumps(
            chart_tipo_operacion_count_valores
        ),
        "chart_tipo_operacion_mensual_json": chart_tipo_operacion_mensual_json,
    }


@app.template_filter("currency")
def currency_filter(value):
    try:
        value = float(value)
        if value == 0:
            return ""
        return f"${value:,.0f}".replace(",", ".")
    except Exception:
        return ""


@app.route("/")
def index():
    try:
        df = load_data()
        filters = parse_filters_from_request()
        filter_options = get_filter_options(df)

        theme = request.args.get("theme", "A").upper()
        if theme not in ["A", "B"]:
            theme = "A"

        filtered_df = apply_filters(df, filters)

        filters_without_month = filters.copy()
        filters_without_month["mes"] = ""

        trend_df = apply_filters(df, filters_without_month)

        data = build_dashboard_data(filtered_df, trend_df)

        return render_template(
            "index.html",
            data=data,
            error=None,
            filters=filters,
            filter_options=filter_options,
            theme=theme,
        )

    except Exception as e:
        return render_template(
            "index.html",
            data=None,
            error=str(e),
            filters={},
            filter_options={},
            theme="A",
        )


@app.route("/exportar_excel")
def exportar_excel():
    df = load_data()
    filters = parse_filters_from_request()
    filtered_df = apply_filters(df, filters).copy()

    export_cols = [
        "fecha_ingreso_solicitud_dt",
        "fecha_respuesta_solicitud_dt",
        "fecha_ingreso_cotizacion_dt",
        "fecha_cierre_cotizacion_dt",
        "canal",
        "cliente",
        "tipo_cliente",
        "contacto",
        "celular",
        "tipo_solicitud",
        "tipo_operacion",
        "estado",
        "estado_cotizacion",
        "responsable_cotizacion",
        "accion_a_seguir",
        "valor_estimado_negocio",
        "observacion",
    ]

    export_cols = [c for c in export_cols if c in filtered_df.columns]
    export_df = filtered_df[export_cols].copy()

    rename_export = {
        "fecha_ingreso_solicitud_dt": "Fecha ingreso solicitud",
        "fecha_respuesta_solicitud_dt": "Fecha respuesta solicitud",
        "fecha_ingreso_cotizacion_dt": "Fecha ingreso cotización",
        "fecha_cierre_cotizacion_dt": "Fecha cierre cotización",
        "canal": "Canal",
        "cliente": "Cliente",
        "tipo_cliente": "Tipo de cliente",
        "contacto": "Contacto",
        "celular": "Celular",
        "tipo_solicitud": "Tipo de solicitud",
        "tipo_operacion": "Tipo de operación",
        "estado": "Estado",
        "estado_cotizacion": "Estado cotización",
        "responsable_cotizacion": "Responsable cotización",
        "accion_a_seguir": "Acción a seguir",
        "valor_estimado_negocio": "Valor estimado negocio",
        "observacion": "Observación",
    }

    export_df = export_df.rename(columns=rename_export)

    output = BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        export_df.to_excel(writer, index=False, sheet_name="CRM Filtrado")

    output.seek(0)

    return send_file(
        output,
        as_attachment=True,
        download_name="crm_filtrado.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


if __name__ == "__main__":
    app.run(debug=True)