from datetime import datetime, timedelta
from zoneinfo import ZoneInfo
import pandas as pd

TZ_AR = ZoneInfo("America/Argentina/Buenos_Aires")

# =========================
# CONSTANTES
# =========================

INCIDENCES_MAP = {
    "ABSENT": "Ausencia sin aviso",
    "LATE": "Tardanza",
    "UNDERWORKED": "Trabajo insuficiente",
}

WEEKDAY_ES_MAP = {
    0: "Lunes",
    1: "Martes",
    2: "Miércoles",
    3: "Jueves",
    4: "Viernes",
    5: "Sábado",
    6: "Domingo",
}

# =========================
# FECHAS / HORAS
# =========================

def iso_to_dt(value, tz=TZ_AR):
    if not value:
        return pd.NaT
    try:
        return datetime.fromisoformat(value.replace("Z", "+00:00")).astimezone(tz)
    except Exception:
        return pd.NaT


def floor_minute(dt):
    if pd.isna(dt):
        return dt
    return dt.replace(second=0, microsecond=0)


def weekday_es(date_str: str) -> str:
    try:
        d = datetime.strptime(date_str, "%Y-%m-%d")
        return WEEKDAY_ES_MAP[d.weekday()]
    except Exception:
        return ""


def fmt_range(start, end):
    if pd.isna(start) or pd.isna(end):
        return ""
    return f"{start.strftime('%H:%M')} - {end.strftime('%H:%M')}"


def calc_delta_hours(real, sched, tolerance_seconds=0):
    if pd.isna(real) or pd.isna(sched):
        return 0.0
    delta = (real - sched).total_seconds() - tolerance_seconds
    return round(delta / 3600, 2) if delta > 0 else 0.0

def duracion_horas_puras(start, end):
    if pd.isna(start) or pd.isna(end):
        return 0.0
    return (end - start).total_seconds() / 3600


# =========================
# INCIDENCIAS
# =========================

def build_observaciones(it):
    lines = []

    # ================= INCIDENCIAS =================
    incidencias_map = {
        "ABSENT": "Ausencia",
        "LATE": "Tardanza",
        "UNDERWORKED": "Trabajo insuficiente",
    }

    incid = it.get("incidences") or []
    for inc in incid:
        inc_norm = str(inc).upper().strip()
        if inc_norm in incidencias_map:
            lines.append(incidencias_map[inc_norm])
        else:
            lines.append(inc_norm)  # fallback por si aparece algo nuevo

    # ================= LICENCIAS =================
    tor = it.get("timeOffRequests") or []
    if tor:
        names = [t.get("name") for t in tor if isinstance(t, dict) and t.get("name")]
        if names:
            lines.append("Licencia: " + ", ".join(names))

    # ================= FERIADOS =================
    hol = it.get("holidays") or []
    if hol:
        names = [h.get("name") for h in hol if isinstance(h, dict) and h.get("name")]
        lines.append("Feriado: " + (", ".join(names) if names else ""))

    # ================= AUSENCIA SIN FICHADAS =================
    slots = it.get("timeSlots") or []
    entries = it.get("entries") or []
    has_time_off = bool(tor)
    has_absent_inc = any(str(x).upper() == "ABSENT" for x in incid)

    if slots and not entries and not has_time_off and not has_absent_inc:
        lines.append("Ausencia sin aviso")

    # 👉 salto de línea entre observaciones
    return "\n".join(lines)


def timeslot_a_horas_decimales(start: str, end: str) -> float:
    """
    Convierte un rango HH:MM - HH:MM a horas decimales.
    Ej: '12:00', '12:50' -> 0.83
    """
    h1, m1 = map(int, start.split(":"))
    h2, m2 = map(int, end.split(":"))

    minutos = (h2 * 60 + m2) - (h1 * 60 + m1)
    return round(minutos / 60, 4)





# =========================
# EXPORTACIÓN EXCEL
# =========================


def aplicar_nocturnidad_50_100(df: pd.DataFrame) -> pd.DataFrame:
    # 1) asegurar numérico
    df["ADICIONAL_NOCTURNIDAD"] = pd.to_numeric(df["ADICIONAL_NOCTURNIDAD"], errors="coerce").fillna(0.0)

    # 2) crear columnas destino (con los nombres EXACTOS)
    col50  = "Horas Extra 50% (por nocturnidad)"
    col100 = "Horas Extra 100% (por nocturnidad)"
    df[col50]  = 0.0
    df[col100] = 0.0

    # 3) definir si es 100% según data API: sábado/domingo/feriado
    weekday = df["_weekday_api"].fillna("").astype(str).str.upper().str.strip()
    is_weekend = weekday.isin(["SATURDAY", "SUNDAY"])
    is_holiday = df["_hasHoliday_api"].fillna(False).astype(bool)

    is_100 = is_weekend | is_holiday

    # 4) asignación: si hay adicional nocturnidad (>0), va a 50 o 100 según is_100
    has_add = df["ADICIONAL_NOCTURNIDAD"] > 0

    df.loc[has_add & is_100,  col100] = df.loc[has_add & is_100,  "ADICIONAL_NOCTURNIDAD"]
    df.loc[has_add & ~is_100, col50]  = df.loc[has_add & ~is_100, "ADICIONAL_NOCTURNIDAD"]

    return df


def export_detalle_diario_excel(
    df_export: pd.DataFrame,
    out: str,
    START_DATE: str,
    END_DATE: str,
    generated_at: str,
    EXPORTAR_DECIMAL: bool,
    COLS_HORAS_DETALLE: list,
):
    with pd.ExcelWriter(out, engine="xlsxwriter") as writer:
        workbook = writer.book

        fmt_title = workbook.add_format({"bold": True, "font_size": 14})
        fmt_sub   = workbook.add_format({"font_size": 11})
        fmt_wrap  = workbook.add_format({"text_wrap": True})
        fmt_hhmm  = workbook.add_format({"num_format": "[h]:mm"}) if not EXPORTAR_DECIMAL else None
        fmt_dec   = workbook.add_format({"num_format": "0.00"})   if EXPORTAR_DECIMAL else None

        # ----- Detalle diario -----
        startrow = 4

        # Ordenar
        df_export = df_export.sort_values(by=["ID", "Fecha"], ascending=[True, True]).reset_index(drop=True)

        # Escribir dataframe
        df_export.to_excel(writer, index=False, sheet_name="Detalle diario", startrow=startrow)
        ws1 = writer.sheets["Detalle diario"]

        # Encabezado
        ws1.write("A1", "DETALLE DIARIO DE ASISTENCIA", fmt_title)
        ws1.write("A2", f"Período: {START_DATE} al {END_DATE}", fmt_sub)
        ws1.write("A3", f"Generado: {generated_at}", fmt_sub)

        # Anchos + formatos
        for idx, col in enumerate(df_export.columns):
            if col == "Observaciones":
                ws1.set_column(idx, idx, 45, fmt_wrap)
            elif col in COLS_HORAS_DETALLE:
                ws1.set_column(idx, idx, 22, fmt_dec if EXPORTAR_DECIMAL else fmt_hhmm)
            else:
                ws1.set_column(idx, idx, 26)

    print("✅ Excel generado:", out)
