from __future__ import annotations

from datetime import datetime
from pathlib import Path
import json

import pandas as pd
from flask import Flask, render_template

app = Flask(__name__)

# Si el Excel está dentro de la carpeta del proyecto
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


def load_data() -> pd.DataFrame:
    excel_path = Path(DEFAULT_EXCEL_PATH)

    if not excel_path.exists():
        raise FileNotFoundError(
            f"No se encontró el archivo Excel en: {excel_path}\n"
            "Pon el archivo dentro de la carpeta del proyecto o ajusta DEFAULT_EXCEL_PATH."
        )

    df = pd.read_excel(excel_path, sheet_name=SHEET_NAME)
    df.columns = [str(c).strip() for c in df.columns]

    col_medio = find_column(df, ["Medio", "MEDIO", "MEDIO DE CONTACTO", "CANAL"])
    col_fecha_sol = find_column(df, ["Fecha Ingreso Solicitud", "FECHA SOLICITUD", "FECHA DE SOLICITUD", "FECHA"])
    col_cliente = find_column(df, ["Cliente", "CLIENTE", "RAZON SOCIAL"])
    col_servicio = find_column(df, ["Tipo de Solicitud", "TIPO DE SOLICITUD", "SERVICIO", "TIPO DE SERVICIO"])
    col_estado = find_column(df, ["Accion a Seguir", "ACCION A SEGUIR", "ACCIÓN A SEGUIR", "Estado Cotizacion", "Estado Cotización", "Estado", "STATUS"])
    col_responsable = find_column(df, ["Responsable Cotizacion", "Responsable Cotización", "RESPONSABLE", "ASESOR", "COMERCIAL"])
    col_fecha_cot = find_column(df, ["Fecha Ingreso Cotizacion", "Fecha Ingreso Cotización", "FECHA COTIZACION", "FECHA DE COTIZACIÓN", "FECHA COTIZADA"])
    col_valor = find_column(df, ["Valor Estimado Negocio", "VALOR ESTIMADO", "VALOR", "MONTO", "NEGOCIO ESTIMADO"])
    col_observacion = find_column(df, ["Observación", "OBSERVACION", "OBSERVACIÓN", "COMENTARIOS"])
    col_tipo_operacion = find_column(df, ["Tipo de Operación", "TIPO DE OPERACIÓN"])
    col_accion = find_column(df, ["Accion a Seguir", "ACCION A SEGUIR", "ACCIÓN A SEGUIR"])
    col_estado_general = find_column(df, ["Estado", "ESTADO"])
    col_estado_cotizacion = find_column(df, ["Estado Cotizacion", "Estado Cotización", "ESTADO COTIZACION"])

    rename_map = {}
    if col_medio:
        rename_map[col_medio] = "medio"
    if col_fecha_sol:
        rename_map[col_fecha_sol] = "fecha_solicitud"
    if col_cliente:
        rename_map[col_cliente] = "cliente"
    if col_servicio:
        rename_map[col_servicio] = "servicio"
    if col_estado:
        rename_map[col_estado] = "estado"
    if col_responsable:
        rename_map[col_responsable] = "responsable"
    if col_fecha_cot:
        rename_map[col_fecha_cot] = "fecha_cotizacion"
    if col_valor:
        rename_map[col_valor] = "valor_estimado"
    if col_observacion:
        rename_map[col_observacion] = "observacion"
    if col_tipo_operacion:
        rename_map[col_tipo_operacion] = "tipo_operacion"
    if col_accion:
        rename_map[col_accion] = "accion_seguir"
    if col_estado_general:
        rename_map[col_estado_general] = "estado_general"
    if col_estado_cotizacion:
        rename_map[col_estado_cotizacion] = "estado_cotizacion"

    df = df.rename(columns=rename_map)

    if "fecha_solicitud" not in df.columns:
        raise ValueError("Falta la columna requerida: fecha_solicitud")

    for text_col in [
        "medio",
        "cliente",
        "servicio",
        "estado",
        "responsable",
        "observacion",
        "tipo_operacion",
        "accion_seguir",
        "estado_general",
        "estado_cotizacion",
    ]:
        if text_col in df.columns:
            df[text_col] = clean_text(df[text_col])

    df["fecha_solicitud"] = to_datetime_safe(df["fecha_solicitud"])

    if "fecha_cotizacion" in df.columns:
        df["fecha_cotizacion"] = to_datetime_safe(df["fecha_cotizacion"])
    else:
        df["fecha_cotizacion"] = pd.NaT

    if "valor_estimado" in df.columns:
        df["valor_estimado"] = to_numeric_safe(df["valor_estimado"]).fillna(0)
    else:
        df["valor_estimado"] = 0.0

    df = df[df["fecha_solicitud"].notna()].copy()

    df["periodo"] = df["fecha_solicitud"].dt.to_period("M")
    df["tiene_cotizacion"] = df["fecha_cotizacion"].notna()

    estado_lower = df["estado"].str.lower() if "estado" in df.columns else pd.Series("", index=df.index)
    accion_lower = df["accion_seguir"].str.lower() if "accion_seguir" in df.columns else pd.Series("", index=df.index)
    estado_cot_lower = df["estado_cotizacion"].str.lower() if "estado_cotizacion" in df.columns else pd.Series("", index=df.index)

    df["es_pendiente"] = (
        ~df["tiene_cotizacion"]
        | accion_lower.str.contains("pendiente", na=False)
        | accion_lower.str.contains("aprobacion", na=False)
        | estado_cot_lower.str.contains("enviada", na=False)
        | estado_lower.str.contains("pend", na=False)
        | estado_lower.str.contains("analisis", na=False)
        | estado_lower.str.contains("análisis", na=False)
    )

    df["dias_desde_ingreso"] = (
        pd.Timestamp.today().normalize() - df["fecha_solicitud"]
    ).dt.days

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
    conversion_total = round((total_cotizaciones / total_solicitudes) * 100, 1) if total_solicitudes else 0

    pendientes_df = df[df["es_pendiente"]].copy()
    pendientes_count = int(len(pendientes_df))

    valor_total = float(df["valor_estimado"].fillna(0).sum())

    negocios_con_valor = df[df["valor_estimado"].fillna(0) > 0].copy()
    cantidad_negocios_con_valor = int(len(negocios_con_valor))
    valor_promedio_por_negocio = (
        float(negocios_con_valor["valor_estimado"].mean())
        if cantidad_negocios_con_valor > 0
        else 0.0
    )

    current_solicitudes = int(len(current_df))
    current_cotizaciones = int(current_df["tiene_cotizacion"].sum())
    current_conversion = round((current_cotizaciones / current_solicitudes) * 100, 1) if current_solicitudes else 0

    previous_solicitudes = int(len(previous_df))
    previous_cotizaciones = int(previous_df["tiene_cotizacion"].sum())
    previous_conversion = round((previous_cotizaciones / previous_solicitudes) * 100, 1) if previous_solicitudes else 0

    monthly = (
        df.groupby("periodo")
        .agg(
            solicitudes=("fecha_solicitud", "count"),
            cotizaciones=("tiene_cotizacion", "sum"),
            valor=("valor_estimado", "sum"),
        )
        .reset_index()
        .sort_values("periodo")
    )
    monthly["mes"] = monthly["periodo"].apply(month_label)
    monthly_data = monthly[["mes", "solicitudes", "cotizaciones", "valor"]].to_dict(orient="records")

    medio_data = []
    if "medio" in df.columns:
        medio = (
            df.groupby("medio")
            .size()
            .reset_index(name="cantidad")
            .sort_values("cantidad", ascending=False)
        )
        medio = medio[(medio["medio"] != "") & (medio["cantidad"] > 0)]
        total_medio = medio["cantidad"].sum()
        for _, row in medio.iterrows():
            medio_data.append({
                "medio": row["medio"],
                "cantidad": int(row["cantidad"]),
                "porcentaje": round((row["cantidad"] / total_medio) * 100, 1) if total_medio else 0,
            })

    servicios_data = []
    if "servicio" in df.columns:
        servicios = (
            df.groupby("servicio")
            .size()
            .reset_index(name="cantidad")
            .sort_values("cantidad", ascending=False)
        )
        servicios = servicios[(servicios["servicio"] != "") & (servicios["cantidad"] > 0)].head(10)
        servicios_data = servicios.to_dict(orient="records")

    estados_data = []
    if "estado_cotizacion" in df.columns:
        estados = (
            df.groupby("estado_cotizacion")
            .size()
            .reset_index(name="cantidad")
            .sort_values("cantidad", ascending=False)
        )
        estados = estados[(estados["estado_cotizacion"] != "") & (estados["cantidad"] > 0)]
        estados_data = estados.to_dict(orient="records")
    elif "estado_general" in df.columns:
        estados = (
            df.groupby("estado_general")
            .size()
            .reset_index(name="cantidad")
            .sort_values("cantidad", ascending=False)
        )
        estados = estados[(estados["estado_general"] != "") & (estados["cantidad"] > 0)]
        estados_data = estados.to_dict(orient="records")

    responsables_data = []
    if "responsable" in pendientes_df.columns:
        responsables = (
            pendientes_df.groupby("responsable")
            .agg(
                pendientes=("responsable", "count"),
                antiguedad_promedio=("dias_desde_ingreso", "mean"),
                caso_mas_antiguo=("dias_desde_ingreso", "max"),
            )
            .reset_index()
            .sort_values(["pendientes", "antiguedad_promedio"], ascending=[False, False])
        )
        responsables = responsables[(responsables["responsable"] != "") & (responsables["pendientes"] > 0)]
        responsables["antiguedad_promedio"] = responsables["antiguedad_promedio"].round(1)
        responsables_data = responsables.to_dict(orient="records")

    clientes_valor_data = []
    if "cliente" in df.columns:
        clientes = (
            df.groupby("cliente")
            .agg(
                valor=("valor_estimado", "sum"),
                cantidad_negocios=("cliente", "count")
            )
            .reset_index()
            .sort_values("valor", ascending=False)
        )
        clientes = clientes[(clientes["cliente"] != "") & (clientes["valor"] > 0)].head(10)
        clientes["valor_promedio_negocio"] = (
            clientes["valor"] / clientes["cantidad_negocios"]
        ).fillna(0)
        clientes_valor_data = clientes.to_dict(orient="records")

    tipo_operacion_data = []
    if "tipo_operacion" in df.columns:
        tipo_operacion = (
            df.groupby("tipo_operacion")
            .agg(
                valor_total=("valor_estimado", "sum"),
                cantidad=("tipo_operacion", "count")
            )
            .reset_index()
            .sort_values("valor_total", ascending=False)
        )

        tipo_operacion = tipo_operacion[
            (tipo_operacion["tipo_operacion"] != "") &
            (tipo_operacion["valor_total"] > 0)
        ]

        total_valor_tipo = tipo_operacion["valor_total"].sum()

        if total_valor_tipo > 0:
            tipo_operacion["porcentaje"] = (
                (tipo_operacion["valor_total"] / total_valor_tipo) * 100
            ).round(1)
        else:
            tipo_operacion["porcentaje"] = 0

        tipo_operacion_data = tipo_operacion.to_dict(orient="records")

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

    temp = pendientes_df.copy()
    temp["fecha_solicitud_fmt"] = temp["fecha_solicitud"].dt.strftime("%Y-%m-%d")

    show_cols = [
        c for c in [
            "cliente",
            "servicio",
            "tipo_operacion",
            "responsable",
            "estado_general",
            "estado_cotizacion",
            "accion_seguir",
        ]
        if c in temp.columns
    ]

    columns_table = ["fecha_solicitud_fmt", "dias_desde_ingreso"] + show_cols
    if "valor_estimado" in temp.columns:
        columns_table.append("valor_estimado")

    pending_table = (
        temp[columns_table]
        .sort_values("dias_desde_ingreso", ascending=False)
        .head(30)
        .to_dict(orient="records")
    )

    insights = []

    if responsables_data:
        top_resp = responsables_data[0]
        insights.append(
            f"{top_resp['responsable']} concentra el mayor volumen de pendientes: "
            f"{top_resp['pendientes']} casos, con antigüedad promedio de {top_resp['antiguedad_promedio']} días."
        )

    older_cases = pendientes_df[pendientes_df["dias_desde_ingreso"] >= 28]
    if len(older_cases) > 0:
        insights.append(
            f"{len(older_cases)} cotizaciones pendientes acumulan 28 días o más desde su ingreso."
        )

    if trend_summary:
        insights.append(
            f"En {trend_summary['current_month']}, las solicitudes variaron {trend_summary['pct_solicitudes']}% "
            f"y las cotizaciones {trend_summary['pct_cotizaciones']}% frente a {trend_summary['previous_month']}."
        )

    # Datos para gráficos
    chart_labels = [row["mes"] for row in monthly_data]
    chart_solicitudes = [row["solicitudes"] for row in monthly_data]
    chart_cotizaciones = [row["cotizaciones"] for row in monthly_data]

    chart_tipo_operacion_labels = [row["tipo_operacion"] for row in tipo_operacion_data]
    chart_tipo_operacion_valores = [float(row["valor_total"]) for row in tipo_operacion_data]

    return {
        "generated_at": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "current_period_label": month_label(current_period),
        "previous_period_label": month_label(previous_period),

        "total_solicitudes": total_solicitudes,
        "total_cotizaciones": total_cotizaciones,
        "conversion_total": conversion_total,
        "pendientes_count": pendientes_count,
        "valor_total": valor_total,
        "cantidad_negocios_con_valor": cantidad_negocios_con_valor,
        "valor_promedio_por_negocio": valor_promedio_por_negocio,

        "current_solicitudes": current_solicitudes,
        "current_cotizaciones": current_cotizaciones,
        "current_conversion": current_conversion,

        "previous_solicitudes": previous_solicitudes,
        "previous_cotizaciones": previous_cotizaciones,
        "previous_conversion": previous_conversion,

        "monthly_data": monthly_data,
        "trend_summary": trend_summary,
        "medio_data": medio_data,
        "servicios_data": servicios_data,
        "estados_data": estados_data,
        "responsables_data": responsables_data,
        "clientes_valor_data": clientes_valor_data,
        "tipo_operacion_data": tipo_operacion_data,
        "pending_table": pending_table,
        "insights": insights,

        "chart_labels_json": json.dumps(chart_labels, ensure_ascii=False),
        "chart_solicitudes_json": json.dumps(chart_solicitudes),
        "chart_cotizaciones_json": json.dumps(chart_cotizaciones),
        "chart_tipo_operacion_labels_json": json.dumps(chart_tipo_operacion_labels, ensure_ascii=False),
        "chart_tipo_operacion_valores_json": json.dumps(chart_tipo_operacion_valores),
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
        data = build_dashboard_data(df)
        return render_template("index.html", data=data, error=None)
    except Exception as e:
        return render_template("index.html", data=None, error=str(e))


if __name__ == "__main__":
    app.run(debug=True)