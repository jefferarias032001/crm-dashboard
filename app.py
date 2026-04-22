from __future__ import annotations

from datetime import datetime
from pathlib import Path
import json

import pandas as pd
from flask import Flask, render_template, request

app = Flask(__name__)

# En Render debe quedar así, con el archivo dentro del proyecto
# En local también te funciona si el Excel está en la carpeta del proyecto
DEFAULT_EXCEL_PATH = "BD CRM 14042026.xlsx"
SHEET_NAME = "BD"


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


def hours_to_text(value: float | int | None) -> str:
    if value is None or pd.isna(value):
        return ""
    value = float(value)
    if value < 24:
        return f"{value:.1f} h"
    days = value / 24
    return f"{days:.1f} días"


def get_filter_options(df: pd.DataFrame) -> dict:
    def options_for(col: str) -> list[str]:
        if col not in df.columns:
            return []
        vals = (
            df[col]
            .dropna()
            .astype(str)
            .str.strip()
        )
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
        "mes": []
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


def load_data() -> pd.DataFrame:
    excel_path = Path(DEFAULT_EXCEL_PATH)

    if not excel_path.exists():
        raise FileNotFoundError(
            f"No se encontró el archivo Excel en: {excel_path}. "
            "Pon el archivo dentro de la carpeta del proyecto o ajusta DEFAULT_EXCEL_PATH."
        )

    df = pd.read_excel(excel_path, sheet_name=SHEET_NAME)
    df.columns = [str(c).strip() for c in df.columns]

    col_canal = find_column(df, ["CANAL", "Canal", "Medio", "MEDIO", "MEDIO DE CONTACTO"])
    col_fecha_ingreso_sol = find_column(df, ["Fecha Ingreso Solicitud", "FECHA INGRESO SOLICITUD", "FECHA SOLICITUD"])
    col_hora_ingreso_sol = find_column(df, ["Hora ingreso Solicitud", "HORA INGRESO SOLICITUD"])
    col_fecha_respuesta_sol = find_column(df, ["Fecha de respuesta Solicitud", "FECHA DE RESPUESTA SOLICITUD"])
    col_hora_respuesta_sol = find_column(df, ["Hora de respuesta Solicitud", "HORA DE RESPUESTA SOLICITUD"])
    col_fecha_ingreso_cot = find_column(df, ["Fecha Ingreso Cotizacion", "Fecha Ingreso Cotización", "FECHA INGRESO COTIZACION"])
    col_fecha_cierre_cot = find_column(df, ["Fecha Cierre Cotizacion", "Fecha Cierre Cotización", "FECHA CIERRE COTIZACION"])
    col_cliente = find_column(df, ["Cliente", "CLIENTE"])
    col_tipo_cliente = find_column(df, ["Tipo de Cliente", "TIPO DE CLIENTE"])
    col_contacto = find_column(df, ["Contacto", "CONTACTO"])
    col_celular = find_column(df, ["Celular", "CELULAR"])
    col_tipo_solicitud = find_column(df, ["Tipo de Solicitud", "TIPO DE SOLICITUD"])
    col_tipo_operacion = find_column(df, ["Tipo de Operación", "TIPO DE OPERACIÓN"])
    col_estado = find_column(df, ["Estado", "ESTADO"])
    col_estado_cot = find_column(df, ["Estado Cotizacion", "Estado Cotización", "ESTADO COTIZACION"])
    col_responsable_cot = find_column(df, ["Responsable Cotizacion", "Responsable Cotización", "RESPONSABLE COTIZACION"])
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
        "estado_cotizacion",
        "responsable_cotizacion",
        "accion_a_seguir",
        "observacion",
    ]:
        if text_col in df.columns:
            df[text_col] = clean_text(df[text_col])

    # Datetimes clave
    df["fecha_ingreso_solicitud_dt"] = combine_date_and_time(
        df["fecha_ingreso_solicitud"],
        df["hora_ingreso_solicitud"] if "hora_ingreso_solicitud" in df.columns else pd.Series("", index=df.index)
    )

    if "fecha_respuesta_solicitud" in df.columns:
        df["fecha_respuesta_solicitud_dt"] = combine_date_and_time(
            df["fecha_respuesta_solicitud"],
            df["hora_respuesta_solicitud"] if "hora_respuesta_solicitud" in df.columns else pd.Series("", index=df.index)
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
    df["mes_label"] = df["periodo"].apply(month_label)

    # Indicadores de tiempo
    df["horas_respuesta_solicitud"] = (
        (df["fecha_respuesta_solicitud_dt"] - df["fecha_ingreso_solicitud_dt"]).dt.total_seconds() / 3600
    )

    df["dias_cierre_cotizacion"] = (
        (df["fecha_cierre_cotizacion_dt"] - df["fecha_ingreso_cotizacion_dt"]).dt.total_seconds() / 86400
    )

    # Flags de negocio
    df["tiene_respuesta_solicitud"] = df["fecha_respuesta_solicitud_dt"].notna()
    df["tiene_cotizacion"] = df["fecha_ingreso_cotizacion_dt"].notna()
    df["cotizacion_cerrada"] = df["fecha_cierre_cotizacion_dt"].notna()

    return df


def build_dashboard_data(df: pd.DataFrame) -> dict:
    if df.empty:
        return {}

    current_period = df["periodo"].max()
    previous_period = current_period - 1

    current_df = df[df["periodo"] == current_period].copy()
    previous_df = df[df["periodo"] == previous_period].copy()

    total_solicitudes = int(len(df))
    total_cotizaciones = int(df["tiene_cotizacion"].sum())
    tasa_cotizacion = round((total_cotizaciones / total_solicitudes) * 100, 1) if total_solicitudes else 0

    valor_total = float(df["valor_estimado_negocio"].fillna(0).sum())
    negocios_con_valor = df[df["valor_estimado_negocio"] > 0].copy()
    valor_promedio_negocio = float(negocios_con_valor["valor_estimado_negocio"].mean()) if len(negocios_con_valor) > 0 else 0

    tiempos_respuesta_validos = df[
        (df["horas_respuesta_solicitud"].notna()) & (df["horas_respuesta_solicitud"] >= 0)
    ]
    tiempo_respuesta_promedio_horas = (
        float(tiempos_respuesta_validos["horas_respuesta_solicitud"].mean())
        if len(tiempos_respuesta_validos) > 0 else 0
    )

    tiempos_cierre_validos = df[
        (df["dias_cierre_cotizacion"].notna()) & (df["dias_cierre_cotizacion"] >= 0)
    ]
    tiempo_cierre_promedio_dias = (
        float(tiempos_cierre_validos["dias_cierre_cotizacion"].mean())
        if len(tiempos_cierre_validos) > 0 else 0
    )

    current_solicitudes = int(len(current_df))
    current_cotizaciones = int(current_df["tiene_cotizacion"].sum())
    current_tasa_cotizacion = round((current_cotizaciones / current_solicitudes) * 100, 1) if current_solicitudes else 0

    previous_solicitudes = int(len(previous_df))
    previous_cotizaciones = int(previous_df["tiene_cotizacion"].sum())
    previous_tasa_cotizacion = round((previous_cotizaciones / previous_solicitudes) * 100, 1) if previous_solicitudes else 0

    # Tendencia mensual
    monthly = (
        df.groupby("periodo")
        .agg(
            solicitudes=("fecha_ingreso_solicitud_dt", "count"),
            cotizaciones=("tiene_cotizacion", "sum"),
            valor=("valor_estimado_negocio", "sum"),
        )
        .reset_index()
        .sort_values("periodo")
    )
    monthly["mes"] = monthly["periodo"].apply(month_label)
    monthly_data = monthly[["mes", "solicitudes", "cotizaciones", "valor"]].to_dict(orient="records")

    trend_summary = {}
    if len(monthly) >= 2:
        last_row = monthly.iloc[-1]
        prev_row = monthly.iloc[-2]

        delta_sol = int(last_row["solicitudes"] - prev_row["solicitudes"])
        delta_cot = int(last_row["cotizaciones"] - prev_row["cotizaciones"])

        pct_sol = round((delta_sol / prev_row["solicitudes"]) * 100, 1) if prev_row["solicitudes"] > 0 else 0
        pct_cot = round((delta_cot / prev_row["cotizaciones"]) * 100, 1) if prev_row["cotizaciones"] > 0 else 0

        trend_summary = {
            "current_month": last_row["mes"],
            "previous_month": prev_row["mes"],
            "delta_solicitudes": delta_sol,
            "delta_cotizaciones": delta_cot,
            "pct_solicitudes": pct_sol,
            "pct_cotizaciones": pct_cot,
        }

    # Canal
    canal_data = []
    if "canal" in df.columns:
        canal = (
            df.groupby("canal")
            .size()
            .reset_index(name="cantidad")
            .sort_values("cantidad", ascending=False)
        )
        canal = canal[(canal["canal"] != "") & (canal["cantidad"] > 0)]
        total_canal = canal["cantidad"].sum()
        for _, row in canal.iterrows():
            canal_data.append({
                "canal": row["canal"],
                "cantidad": int(row["cantidad"]),
                "porcentaje": round((row["cantidad"] / total_canal) * 100, 1) if total_canal else 0
            })

    # Tipo cliente
    tipo_cliente_data = []
    if "tipo_cliente" in df.columns:
        tipo_cliente = (
            df.groupby("tipo_cliente")
            .size()
            .reset_index(name="cantidad")
            .sort_values("cantidad", ascending=False)
        )
        tipo_cliente = tipo_cliente[(tipo_cliente["tipo_cliente"] != "") & (tipo_cliente["cantidad"] > 0)]
        tipo_cliente_data = tipo_cliente.to_dict(orient="records")

    # Tipo solicitud
    tipo_solicitud_data = []
    if "tipo_solicitud" in df.columns:
        tipo_solicitud = (
            df.groupby("tipo_solicitud")
            .size()
            .reset_index(name="cantidad")
            .sort_values("cantidad", ascending=False)
        )
        tipo_solicitud = tipo_solicitud[(tipo_solicitud["tipo_solicitud"] != "") & (tipo_solicitud["cantidad"] > 0)]
        tipo_solicitud_data = tipo_solicitud.head(10).to_dict(orient="records")

    # Tipo operación
    tipo_operacion_data = []
    if "tipo_operacion" in df.columns:
        tipo_operacion = (
            df.groupby("tipo_operacion")
            .agg(
                cantidad=("tipo_operacion", "count"),
                valor_total=("valor_estimado_negocio", "sum")
            )
            .reset_index()
            .sort_values("valor_total", ascending=False)
        )
        tipo_operacion = tipo_operacion[
            (tipo_operacion["tipo_operacion"] != "") &
            ((tipo_operacion["cantidad"] > 0) | (tipo_operacion["valor_total"] > 0))
        ]

        total_valor_operacion = tipo_operacion["valor_total"].sum()
        tipo_operacion["porcentaje_valor"] = (
            (tipo_operacion["valor_total"] / total_valor_operacion) * 100
        ).round(1) if total_valor_operacion > 0 else 0

        tipo_operacion_data = tipo_operacion.to_dict(orient="records")

    # Estado
    estado_data = []
    if "estado" in df.columns:
        estado = (
            df.groupby("estado")
            .size()
            .reset_index(name="cantidad")
            .sort_values("cantidad", ascending=False)
        )
        estado = estado[(estado["estado"] != "") & (estado["cantidad"] > 0)]
        estado_data = estado.to_dict(orient="records")

    # Estado cotización
    estado_cotizacion_data = []
    if "estado_cotizacion" in df.columns:
        estado_cot = (
            df.groupby("estado_cotizacion")
            .size()
            .reset_index(name="cantidad")
            .sort_values("cantidad", ascending=False)
        )
        estado_cot = estado_cot[(estado_cot["estado_cotizacion"] != "") & (estado_cot["cantidad"] > 0)]
        estado_cotizacion_data = estado_cot.to_dict(orient="records")

    # Valor por cliente
    valor_cliente_data = []
    if "cliente" in df.columns:
        cliente = (
            df.groupby("cliente")
            .agg(
                cantidad=("cliente", "count"),
                valor_total=("valor_estimado_negocio", "sum")
            )
            .reset_index()
            .sort_values("valor_total", ascending=False)
        )
        cliente = cliente[(cliente["cliente"] != "") & ((cliente["cantidad"] > 0) | (cliente["valor_total"] > 0))]
        cliente["valor_promedio"] = (
            cliente["valor_total"] / cliente["cantidad"]
        ).fillna(0)
        valor_cliente_data = cliente.head(10).to_dict(orient="records")

    # Responsable
    responsable_data = []
    if "responsable_cotizacion" in df.columns:
        responsable = (
            df.groupby("responsable_cotizacion")
            .agg(
                cantidad=("responsable_cotizacion", "count"),
                valor_total=("valor_estimado_negocio", "sum")
            )
            .reset_index()
            .sort_values(["cantidad", "valor_total"], ascending=[False, False])
        )
        responsable = responsable[(responsable["responsable_cotizacion"] != "") & (responsable["cantidad"] > 0)]
        responsable_data = responsable.to_dict(orient="records")

    # Tabla detalle
    detalle = df.copy()
    detalle["fecha_ingreso_solicitud_fmt"] = detalle["fecha_ingreso_solicitud_dt"].dt.strftime("%Y-%m-%d %H:%M")
    detalle["fecha_respuesta_solicitud_fmt"] = detalle["fecha_respuesta_solicitud_dt"].dt.strftime("%Y-%m-%d %H:%M")
    detalle["fecha_ingreso_cotizacion_fmt"] = detalle["fecha_ingreso_cotizacion_dt"].dt.strftime("%Y-%m-%d")
    detalle["fecha_cierre_cotizacion_fmt"] = detalle["fecha_cierre_cotizacion_dt"].dt.strftime("%Y-%m-%d")

    detalle_cols = [
        c for c in [
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
        ] if c in detalle.columns
    ]

    detalle_data = (
        detalle[detalle_cols]
        .sort_values("fecha_ingreso_solicitud_fmt", ascending=False)
        .head(50)
        .to_dict(orient="records")
    )

    # Insights ejecutivos
    insights = []

    if trend_summary:
        insights.append(
            f"En {trend_summary['current_month']}, las solicitudes variaron {trend_summary['pct_solicitudes']}% "
            f"y las cotizaciones {trend_summary['pct_cotizaciones']}% frente a {trend_summary['previous_month']}."
        )

    if canal_data:
        top_canal = canal_data[0]
        insights.append(
            f"El canal con mayor ingreso de solicitudes es {top_canal['canal']}, con {top_canal['cantidad']} casos "
            f"({top_canal['porcentaje']}% del total filtrado)."
        )

    if tipo_operacion_data:
        top_operacion = tipo_operacion_data[0]
        if float(top_operacion["valor_total"]) > 0:
            insights.append(
                f"El tipo de operación con mayor valor estimado es {top_operacion['tipo_operacion']}, "
                f"con {top_operacion['porcentaje_valor']}% del valor total del negocio."
            )

    if estado_cotizacion_data:
        top_estado_cot = estado_cotizacion_data[0]
        insights.append(
            f"El estado de cotización más frecuente es {top_estado_cot['estado_cotizacion']}, "
            f"con {top_estado_cot['cantidad']} registros."
        )

    chart_labels = [row["mes"] for row in monthly_data]
    chart_solicitudes = [row["solicitudes"] for row in monthly_data]
    chart_cotizaciones = [row["cotizaciones"] for row in monthly_data]

    chart_tipo_operacion_labels = [row["tipo_operacion"] for row in tipo_operacion_data]
    chart_tipo_operacion_valores = [float(row["valor_total"]) for row in tipo_operacion_data]

    chart_canal_labels = [row["canal"] for row in canal_data]
    chart_canal_valores = [row["cantidad"] for row in canal_data]

    chart_tipo_cliente_labels = [row["tipo_cliente"] for row in tipo_cliente_data]
    chart_tipo_cliente_valores = [row["cantidad"] for row in tipo_cliente_data]

    return {
        "generated_at": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "current_period_label": month_label(current_period),
        "previous_period_label": month_label(previous_period),

        "total_solicitudes": total_solicitudes,
        "total_cotizaciones": total_cotizaciones,
        "tasa_cotizacion": tasa_cotizacion,
        "tiempo_respuesta_promedio_horas": tiempo_respuesta_promedio_horas,
        "tiempo_respuesta_promedio_txt": hours_to_text(tiempo_respuesta_promedio_horas),
        "tiempo_cierre_promedio_dias": tiempo_cierre_promedio_dias,
        "tiempo_cierre_promedio_txt": hours_to_text(tiempo_cierre_promedio_dias * 24 if tiempo_cierre_promedio_dias else 0),
        "valor_total": valor_total,
        "valor_promedio_negocio": valor_promedio_negocio,

        "current_solicitudes": current_solicitudes,
        "current_cotizaciones": current_cotizaciones,
        "current_tasa_cotizacion": current_tasa_cotizacion,

        "previous_solicitudes": previous_solicitudes,
        "previous_cotizaciones": previous_cotizaciones,
        "previous_tasa_cotizacion": previous_tasa_cotizacion,

        "monthly_data": monthly_data,
        "trend_summary": trend_summary,
        "canal_data": canal_data,
        "tipo_cliente_data": tipo_cliente_data,
        "tipo_solicitud_data": tipo_solicitud_data,
        "tipo_operacion_data": tipo_operacion_data,
        "estado_data": estado_data,
        "estado_cotizacion_data": estado_cotizacion_data,
        "valor_cliente_data": valor_cliente_data,
        "responsable_data": responsable_data,
        "detalle_data": detalle_data,
        "insights": insights,

        "chart_labels_json": json.dumps(chart_labels, ensure_ascii=False),
        "chart_solicitudes_json": json.dumps(chart_solicitudes),
        "chart_cotizaciones_json": json.dumps(chart_cotizaciones),
        "chart_tipo_operacion_labels_json": json.dumps(chart_tipo_operacion_labels, ensure_ascii=False),
        "chart_tipo_operacion_valores_json": json.dumps(chart_tipo_operacion_valores),
        "chart_canal_labels_json": json.dumps(chart_canal_labels, ensure_ascii=False),
        "chart_canal_valores_json": json.dumps(chart_canal_valores),
        "chart_tipo_cliente_labels_json": json.dumps(chart_tipo_cliente_labels, ensure_ascii=False),
        "chart_tipo_cliente_valores_json": json.dumps(chart_tipo_cliente_valores),
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

        filters = {
            "canal": request.args.get("canal", ""),
            "tipo_cliente": request.args.get("tipo_cliente", ""),
            "tipo_solicitud": request.args.get("tipo_solicitud", ""),
            "tipo_operacion": request.args.get("tipo_operacion", ""),
            "estado": request.args.get("estado", ""),
            "estado_cotizacion": request.args.get("estado_cotizacion", ""),
            "responsable_cotizacion": request.args.get("responsable_cotizacion", ""),
            "mes": request.args.get("mes", ""),
        }

        filter_options = get_filter_options(df)
        filtered_df = apply_filters(df, filters)
        data = build_dashboard_data(filtered_df)

        return render_template(
            "index.html",
            data=data,
            error=None,
            filters=filters,
            filter_options=filter_options
        )
    except Exception as e:
        return render_template(
            "index.html",
            data=None,
            error=str(e),
            filters={},
            filter_options={}
        )


if __name__ == "__main__":
    app.run(debug=True)