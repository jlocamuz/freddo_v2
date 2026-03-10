import math
import requests
import pandas as pd
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo
import numpy as np

from aux_functions import (
    aplicar_nocturnidad_50_100,
    aplicar_nocturnidad_50_100_real,
    calc_delta_hours,
    duracion_horas_puras,
    export_detalle_diario_excel,
    weekday_es,
    iso_to_dt,
    floor_minute,
    fmt_range,
    build_observaciones,
)

# ================= CONFIG =================
BASE = "https://api-prod.humand.co/public/api/v1"
AUTH = "Basic Njc3NDUzMjpkNWR2Z1pzNXQ3ZEZ2XzE2Z2pfbV9XNklpVFNPU0NmMQ=="

MINUTOS_ADICIONAL_POR_HORA = 8
FACTOR_ADICIONAL = MINUTOS_ADICIONAL_POR_HORA / 60

TOLERANCIA_TARDANZA_SEG = 10 * 60
TOLERANCIA_RETIRO_SEG  = 10 * 60

START_DATE = "2026-01-19"
END_DATE   = "2026-02-19"

LIMIT_USERS = 50
LIMIT_DAYS  = 500
BATCH_SIZE  = 25
MAX_WORKERS = 8

TZ_AR = ZoneInfo("America/Argentina/Buenos_Aires")

NORMALIZAR_A_MINUTO = True
FILTRAR_SIN_ESQUEMA_JORNADA = True  # False = no filtra / True = filtra

CATEGORIAS = [
    "REGULAR",
    "NOCTURNA",
    "EXTRA",
    "EXTRA AL 50",
    "EXTRA NOCTURNA",
    "EXTRA AL 100",
    "EXTRA AL 100 NOCTURNA",
    "HORA FERIADO",
    "HORA FERIADO NOCTURNA",
]

# ================= SESSION =================
s = requests.Session()
s.headers.update({"Authorization": AUTH})

def get(url, params):
    r = s.get(url, params=params, timeout=60)
    r.raise_for_status()
    return r.json()

# ================= USERS =================
def fetch_users():
    first = get(f"{BASE}/users", {"page": 1, "limit": LIMIT_USERS})
    pages = math.ceil(first["count"] / LIMIT_USERS)

    users = first["users"]
    for p in range(2, pages + 1):
        users += get(f"{BASE}/users", {"page": p, "limit": LIMIT_USERS})["users"]

    user_map, legajo_map, esquema_map, employee_ids = {}, {}, {}, []

    for u in users:
        emp = u.get("employeeInternalId")
        if not emp:
            continue

        employee_ids.append(emp)
        user_map[emp] = f"{u.get('lastName','')}, {u.get('firstName','')}"

        legajo = ""
        for f in (u.get("fields") or []):
            if f.get("name") == "Legajo":
                legajo = f.get("value")
                break
        legajo_map[emp] = legajo

        esquema = ""
        for seg in (u.get("segmentations") or []):
            if (seg.get("group") or "").strip() == "Esquema de Jornada":
                esquema = seg.get("item") or ""
                break
        esquema_map[emp] = esquema

    return employee_ids, user_map, legajo_map, esquema_map

# ================= CATEGORÍAS =================
def split_categorized_hours_basic(categorized_hours, categorias_validas):
    valid_upper = {c.upper(): c for c in categorias_validas}
    out = {f"HORAS_{c}": 0.0 for c in categorias_validas}

    for ch in categorized_hours or []:
        name = (ch.get("category", {}) or {}).get("name") or ""
        name_u = str(name).upper().strip()
        if name_u in valid_upper:
            label = valid_upper[name_u]
            out[f"HORAS_{label}"] += float(ch.get("hours") or 0)

    return {k: round(v, 2) for k, v in out.items()}

# ================= DAY SUMMARIES =================
def fetch_batch(emp_ids, user_map, legajo_map, esquema_map):
    rows = []
    page = 1

    while True:
        data = get(
            f"{BASE}/time-tracking/day-summaries",
            {
                "employeeIds": ",".join(emp_ids),
                "startDate": START_DATE,
                "endDate": END_DATE,
                "limit": LIMIT_DAYS,
                "page": page,
            },
        )

        items = data.get("items", [])
        if not items:
            break

        for it in items:
            emp = it.get("employeeId")
            ref = (it.get("referenceDate") or it.get("date") or "")[:10]
            if not ref:
                continue

            entries = it.get("entries") or []
            slots   = it.get("timeSlots") or []
            incid   = it.get("incidences") or []
            tor     = it.get("timeOffRequests") or []
            hol     = it.get("holidays") or []
            cat     = it.get("categorizedHours") or []

            # ── Horas trabajadas desde la API (campo worked) ──
            worked_api = float((it.get("hours") or {}).get("worked") or 0)

            sched_start = sched_end = pd.NaT
            if slots and isinstance(slots, list):
                d = datetime.strptime(ref, "%Y-%m-%d")
                s0 = slots[0] if slots else {}
                if isinstance(s0, dict):
                    if s0.get("startTime"):
                        try:
                            h, m = map(int, s0["startTime"].split(":"))
                            sched_start = datetime(d.year, d.month, d.day, h, m, tzinfo=TZ_AR)
                        except Exception:
                            sched_start = pd.NaT
                    if s0.get("endTime"):
                        try:
                            h, m = map(int, s0["endTime"].split(":"))
                            sched_end = datetime(d.year, d.month, d.day, h, m, tzinfo=TZ_AR)
                            if not pd.isna(sched_start) and sched_end < sched_start:
                                sched_end += timedelta(days=1)
                        except Exception:
                            sched_end = pd.NaT

            real_start = real_end = pd.NaT
            if entries and isinstance(entries, list):
                for e in entries:
                    if isinstance(e, dict) and e.get("type") == "START" and pd.isna(real_start):
                        real_start = iso_to_dt(e.get("time") or e.get("date"), TZ_AR)
                        break
                for e in entries:
                    if isinstance(e, dict) and e.get("type") == "END":
                        real_end = iso_to_dt(e.get("time") or e.get("date"), TZ_AR)

            if NORMALIZAR_A_MINUTO:
                sched_start = floor_minute(sched_start)
                sched_end   = floor_minute(sched_end)
                real_start  = floor_minute(real_start)
                real_end    = floor_minute(real_end)

            cat_hours = split_categorized_hours_basic(cat, CATEGORIAS)

            row = {
                "ID": emp,
                "APELLIDO, NOMBRE": user_map.get(emp, ""),
                "LEGAJO": legajo_map.get(emp, ""),
                "ESQUEMA_JORNADA": esquema_map.get(emp, ""),

                "FECHA": ref,
                "DIA": weekday_es(ref),

                "_weekday_api": (it.get("weekday") or "").upper().strip(),
                "_isWorkday_api": bool(it.get("isWorkday", True)),
                "_hasHoliday_api": bool(it.get("holidays") or []),
                "_worked_api": worked_api,

                "_ss": sched_start,
                "_se": sched_end,
                "_rs": real_start,
                "_re": real_end,

                "HORARIO_OBLIGATORIO": fmt_range(sched_start, sched_end),
                "FICHADAS": fmt_range(real_start, real_end),
                "OBSERVACIONES": build_observaciones(it),
            }
            row.update(cat_hours)
            rows.append(row)

        if len(items) < LIMIT_DAYS:
            break
        page += 1

    return rows

def build_df(employee_ids, user_map, legajo_map, esquema_map):
    rows = []
    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as ex:
        futures = [
            ex.submit(fetch_batch, employee_ids[i:i + BATCH_SIZE], user_map, legajo_map, esquema_map)
            for i in range(0, len(employee_ids), BATCH_SIZE)
        ]
        for f in as_completed(futures):
            rows.extend(f.result())

    df = pd.DataFrame(rows)

    df["HORAS_REGULAR"]  = pd.to_numeric(df.get("HORAS_REGULAR", 0), errors="coerce").fillna(0.0)
    df["HORAS_NOCTURNA"] = pd.to_numeric(df.get("HORAS_NOCTURNA", 0), errors="coerce").fillna(0.0)

    # ── Horas Trabajadas desde API (campo worked) ──
    df["HORAS_TRABAJADAS"] = pd.to_numeric(df["_worked_api"], errors="coerce").fillna(0.0)

    # ── Horas Extra = Horas Trabajadas - 8 ──
    df["HORAS_EXTRA"] = (df["HORAS_TRABAJADAS"] - 8.0).clip(lower=0.0).round(4)

    df["ADICIONAL_NOCTURNIDAD"] = (pd.to_numeric(df["HORAS_NOCTURNA"], errors="coerce").fillna(0) * FACTOR_ADICIONAL).round(4)

    # ── CJN y Adicional Nocturnidad REAL (lógica Leandro) ──────────────────
    is_part_build = df["ESQUEMA_JORNADA"].astype(str).str.upper().str.contains(r"\bPART\s*TIME\b", regex=True, na=False)
    jornada_esperada = is_part_build.map({True: 4.0, False: 8.0})

    df["CJN"] = (jornada_esperada - df["HORAS_TRABAJADAS"]).clip(lower=0.0).round(4)
    df["ADICIONAL_NOCTURNIDAD_REAL"] = (df["ADICIONAL_NOCTURNIDAD"] - df["CJN"]).clip(lower=0.0).round(4)
    # ───────────────────────────────────────────────────────────────────────

    df = aplicar_nocturnidad_50_100(df)
    df = aplicar_nocturnidad_50_100_real(df)

    for c in [
        "HORAS_EXTRA AL 50",
        "HORAS_EXTRA AL 100",
        "HORAS_HORA FERIADO",
        "HORAS_HORA FERIADO NOCTURNA",
    ]:
        if c not in df.columns:
            df[c] = 0.0
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0.0)

    return df

def main():
    employee_ids, user_map, legajo_map, esquema_map = fetch_users()
    print(f"Usuarios ACTIVE: {len(employee_ids)}")

    df = build_df(employee_ids, user_map, legajo_map, esquema_map)

    df = df.sort_values(by=["ID", "FECHA"], ascending=[True, True]).reset_index(drop=True)

    if FILTRAR_SIN_ESQUEMA_JORNADA:
        df = df[
            df["ESQUEMA_JORNADA"].notna()
            & (df["ESQUEMA_JORNADA"].astype(str).str.strip() != "")
        ]

    df_export = df.copy()

    df_export["TARDANZA"] = df_export.apply(
        lambda r: round(max(0.0, calc_delta_hours(r["_rs"], r["_ss"], TOLERANCIA_TARDANZA_SEG)), 2),
        axis=1
    )
    df_export["RETIRO ANTICIPADO"] = df_export.apply(
        lambda r: round(max(0.0, calc_delta_hours(r["_se"], r["_re"], TOLERANCIA_RETIRO_SEG)), 2),
        axis=1
    )

    for c in ["_ss", "_se", "_rs", "_re"]:
        if c in df_export.columns:
            df_export[c] = pd.to_datetime(df_export[c], errors="coerce").dt.tz_localize(None)

    # RENAME FINAL PARA EXCEL
    rename_excel = {
        "ID": "ID",
        "APELLIDO, NOMBRE": "Apellido, Nombre",
        "LEGAJO": "Legajo",
        "ESQUEMA_JORNADA": "Esquema de Jornada",
        "FECHA": "Fecha",
        "DIA": "dia",

        "HORARIO_OBLIGATORIO": "Horario obligatorio",
        "FICHADAS": "Fichadas",
        "OBSERVACIONES": "Observaciones",

        "HORAS_TRABAJADAS": "Horas Trabajadas",
        "HORAS_EXTRA": "Horas extra",
        "HORAS_NOCTURNA": "Horas Nocturnas",
        "ADICIONAL_NOCTURNIDAD": "Adicional Nocturnidad",

        "CJN": "CJN",
        "ADICIONAL_NOCTURNIDAD_REAL": "Adicional Nocturnidad REAL",

        "HORAS_EXTRA AL 50": "Horas Extra 50%",
        "HORAS_EXTRA AL 100": "Horas Extra 100%",

        "Horas Extra 50% (por nocturnidad)": "Horas Extra 50% (por nocturnidad)",
        "Horas Extra 100% (por nocturnidad)": "Horas Extra 100% (por nocturnidad)",

        "Horas Extra 50% (por nocturnidad) Real": "Horas Extra 50% (por nocturnidad) REAL",
        "Horas Extra 100% (por nocturnidad) Real": "Horas Extra 100% (por nocturnidad) REAL",

        "HORAS_HORA FERIADO": "Horas Feriado",
        "HORAS_HORA FERIADO NOCTURNA": "Horas Feriado Nocturnas",

        "TARDANZA": "Tardanza",
        "RETIRO ANTICIPADO": "Retiro Anticipado",
    }

    df_export = df_export.rename(columns=rename_excel)

    # ==========================================================
    # AJUSTE PART TIME
    # ==========================================================
    for col in [
        "Horas Extra 50% (por nocturnidad)", "Horas Extra 100% (por nocturnidad)",
        "Horas Extra 50% (por nocturnidad) REAL", "Horas Extra 100% (por nocturnidad) REAL",
    ]:
        if col not in df_export.columns:
            df_export[col] = 0.0

    if "Complementaria por nocturnidad" not in df_export.columns:
        df_export["Complementaria por nocturnidad"] = 0.0
    if "Complementaria por nocturnidad REAL" not in df_export.columns:
        df_export["Complementaria por nocturnidad REAL"] = 0.0

    is_part_time = (
        df_export.get("Esquema de Jornada", "")
        .astype(str)
        .str.upper()
        .str.contains(r"\bPART\s*TIME\b", regex=True, na=False)
    )

    ad_noct      = pd.to_numeric(df_export.get("Adicional Nocturnidad", 0), errors="coerce").fillna(0.0)
    ad_noct_real = pd.to_numeric(df_export.get("Adicional Nocturnidad REAL", 0), errors="coerce").fillna(0.0)

    # Complementaria original (Part Time)
    df_export.loc[is_part_time & (ad_noct > 0), "Complementaria por nocturnidad"] = ad_noct
    df_export.loc[is_part_time, "Horas Extra 50% (por nocturnidad)"]  = 0.0
    df_export.loc[is_part_time, "Horas Extra 100% (por nocturnidad)"] = 0.0

    # Complementaria REAL (Part Time)
    df_export.loc[is_part_time & (ad_noct_real > 0), "Complementaria por nocturnidad REAL"] = ad_noct_real
    df_export.loc[is_part_time, "Horas Extra 50% (por nocturnidad) REAL"]  = 0.0
    df_export.loc[is_part_time, "Horas Extra 100% (por nocturnidad) REAL"] = 0.0

    # Totales 50/100 originales
    df_export["Horas Extra 50% (total)"] = (
        pd.to_numeric(df_export.get("Horas Extra 50%", 0), errors="coerce").fillna(0.0)
        + pd.to_numeric(df_export.get("Horas Extra 50% (por nocturnidad)", 0), errors="coerce").fillna(0.0)
    )
    df_export["Horas Extra 100% (total)"] = (
        pd.to_numeric(df_export.get("Horas Extra 100%", 0), errors="coerce").fillna(0.0)
        + pd.to_numeric(df_export.get("Horas Extra 100% (por nocturnidad)", 0), errors="coerce").fillna(0.0)
    )

    # Totales 50/100 REAL
    df_export["Horas Extra 50% (total) REAL"] = (
        pd.to_numeric(df_export.get("Horas Extra 50%", 0), errors="coerce").fillna(0.0)
        + pd.to_numeric(df_export.get("Horas Extra 50% (por nocturnidad) REAL", 0), errors="coerce").fillna(0.0)
    )
    df_export["Horas Extra 100% (total) REAL"] = (
        pd.to_numeric(df_export.get("Horas Extra 100%", 0), errors="coerce").fillna(0.0)
        + pd.to_numeric(df_export.get("Horas Extra 100% (por nocturnidad) REAL", 0), errors="coerce").fillna(0.0)
    )

    # Quitar columnas técnicas
    df_export = df_export.drop(
        columns=["_weekday_api", "_isWorkday_api", "_hasHoliday_api", "_ss", "_se", "_rs", "_re", "_worked_api"],
        errors="ignore"
    )

    # Orden final
    cols_final = [
        "ID",
        "Apellido, Nombre",
        "Legajo",
        "Esquema de Jornada",
        "Fecha",
        "dia",
        "Horario obligatorio",
        "Fichadas",
        "Observaciones",
        "Horas Trabajadas",
        "Horas extra",
        "Horas Nocturnas",
        # Columnas originales
        "Adicional Nocturnidad",
        "Horas Extra 50% (por nocturnidad)",
        "Horas Extra 50% (total)",
        "Horas Extra 100% (por nocturnidad)",
        "Horas Extra 100% (total)",
        "Complementaria por nocturnidad",
        # Columnas REAL
        "CJN",
        "Adicional Nocturnidad REAL",
        "Horas Extra 50% (por nocturnidad) REAL",
        "Horas Extra 50% (total) REAL",
        "Horas Extra 100% (por nocturnidad) REAL",
        "Horas Extra 100% (total) REAL",
        "Complementaria por nocturnidad REAL",
        # Resto
        "Horas Extra 50%",
        "Horas Extra 100%",
        "Horas Feriado",
        "Horas Feriado Nocturnas",
        "Tardanza",
        "Retiro Anticipado",
    ]

    for c in cols_final:
        if c not in df_export.columns:
            df_export[c] = 0.0

    cols_cero_vacio = [
        "Horas Trabajadas",
        "Horas extra",
        "Horas Nocturnas",
        "Adicional Nocturnidad",
        "Horas Extra 50% (por nocturnidad)",
        "Horas Extra 50% (total)",
        "Horas Extra 100% (por nocturnidad)",
        "Horas Extra 100% (total)",
        "Complementaria por nocturnidad",
        "CJN",
        "Adicional Nocturnidad REAL",
        "Horas Extra 50% (por nocturnidad) REAL",
        "Horas Extra 50% (total) REAL",
        "Horas Extra 100% (por nocturnidad) REAL",
        "Horas Extra 100% (total) REAL",
        "Complementaria por nocturnidad REAL",
        "Horas Extra 50%",
        "Horas Extra 100%",
        "Horas Feriado",
        "Horas Feriado Nocturnas",
        "Tardanza",
        "Retiro Anticipado",
    ]

    for c in cols_cero_vacio:
        if c in df_export.columns:
            df_export[c] = df_export[c].replace(0, "")

    df_export = df_export[cols_final]

    EXPORTAR_DECIMAL = True

    COLS_HORAS_DETALLE = [
        "Horas Trabajadas",
        "Horas extra",
        "Horas Nocturnas",
        "Adicional Nocturnidad",
        "Horas Extra 50% (por nocturnidad)",
        "Horas Extra 50% (total)",
        "Horas Extra 100% (por nocturnidad)",
        "Horas Extra 100% (total)",
        "Complementaria por nocturnidad",
        "CJN",
        "Adicional Nocturnidad REAL",
        "Horas Extra 50% (por nocturnidad) REAL",
        "Horas Extra 50% (total) REAL",
        "Horas Extra 100% (por nocturnidad) REAL",
        "Horas Extra 100% (total) REAL",
        "Complementaria por nocturnidad REAL",
        "Horas Extra 50%",
        "Horas Extra 100%",
        "Horas Feriado",
        "Horas Feriado Nocturnas",
        "Tardanza",
        "Retiro Anticipado",
    ]

    out = f"reporte_basico_{START_DATE}-{END_DATE}.xlsx"
    generated_at = datetime.now().strftime("%Y-%m-%d %H:%M")

    export_detalle_diario_excel(
        df_export=df_export,
        out=out,
        START_DATE=START_DATE,
        END_DATE=END_DATE,
        generated_at=generated_at,
        EXPORTAR_DECIMAL=EXPORTAR_DECIMAL,
        COLS_HORAS_DETALLE=COLS_HORAS_DETALLE,
    )

    print(f"Excel generado: {out}")

if __name__ == "__main__":
    main()