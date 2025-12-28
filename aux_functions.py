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


# =========================
# CATEGORÍAS / HORAS
# =========================

def split_categorized_hours(categorized_hours, categorias_validas):
    """
    Devuelve dict:
    HORAS_<CATEGORIA> = horas
    """
    out = {}
    for c in categorias_validas:
        out[f"HORAS_{c}"] = 0.0

    for ch in categorized_hours or []:
        name = (ch.get("category", {}).get("name") or "").upper().strip()
        if name in categorias_validas:
            out[f"HORAS_{name}"] += float(ch.get("hours") or 0)

    return {k: round(v, 2) for k, v in out.items()}


# =========================
# INCIDENCIAS
# =========================

def build_observaciones(day):
    incs = day.get("incidences") or []
    textos = [INCIDENCES_MAP.get(i, i) for i in incs]
    return "\n".join(textos)


# =========================
# JORNADA / CLASIFICACIÓN
# =========================

def clasificar_empleado_por_scheduled_max(df, col_sched="SCHEDULED_HOURS"):
    """
    FULL-TIME  : max scheduled >= 8
    PART-TIME  : max scheduled < 8
    """
    res = {}
    grouped = df.groupby("ID")[col_sched].max()

    for emp, max_h in grouped.items():
        try:
            max_h = float(max_h or 0)
        except Exception:
            max_h = 0

        if max_h >= 8:
            res[emp] = "FULL-TIME"
        else:
            res[emp] = "PART-TIME"

    return res


# =========================
# NOCTURNIDAD
# =========================

def nocturnidad_es_100(row):
    """
    Regla:
    - Domingo
    - Feriado
    - No laborable
    - Sábado (fallback conservador)
    """
    weekday = row.get("_weekday_api", "")
    if weekday in ("SUNDAY",):
        return True

    if row.get("_hasHoliday_api", False):
        return True

    if row.get("_isWorkday_api") is False:
        return True

    if weekday == "SATURDAY":
        return True

    return False


# =========================
# STRINGS / NOMBRES
# =========================

def split_apellido_nombre(value):
    if not value or "," not in value:
        return "", ""
    apellido, nombre = value.split(",", 1)
    return apellido.strip(), nombre.strip()


# =========================
# EXPORTACIÓN EXCEL
# =========================

def horas_para_excel(value, usar_decimal=True):
    """
    - 0 => celda vacía
    - decimal => float
    - hh:mm => fracción de día
    """
    try:
        v = float(value)
    except Exception:
        return ""

    if v == 0:
        return ""

    if usar_decimal:
        return round(v, 2)

    return v / 24.0
