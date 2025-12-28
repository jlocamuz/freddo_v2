# ================= CONFIG =================
import math
import requests
import pandas as pd
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo

from aux_functions import (
    weekday_es,
    iso_to_dt,
    floor_minute,
    fmt_range,
    calc_delta_hours,
    split_categorized_hours,
    build_observaciones,
    clasificar_empleado_por_scheduled_max,
    nocturnidad_es_100,
    split_apellido_nombre,
    horas_para_excel,
)


BASE = "https://api-prod.humand.co/public/api/v1"
AUTH = "Basic Njc3NDUzMjpkNWR2Z1pzNXQ3ZEZ2XzE2Z2pfbV9XNklpVFNPU0NmMQ=="

START_DATE = "2025-10-21"
END_DATE   = "2025-11-20"

LIMIT_USERS = 50
LIMIT_DAYS  = 500
BATCH_SIZE  = 25
MAX_WORKERS = 8

TZ_AR = ZoneInfo("America/Argentina/Buenos_Aires")

# ✅ Flag: True = decimal / False = [h]:mm
EXPORTAR_DECIMAL = True

NORMALIZAR_A_MINUTO = True
TOLERANCIA_TARDANZA_SEG = 0
TOLERANCIA_RETIRO_SEG = 0

MONTHLY_CAP_FT = 190
MONTHLY_CAP_PT = 95

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

COLS_HORAS_DETALLE = [
    "Horas Trabajadas",
    "Horas Regulares",
    "Horas extra",
    "Horas Nocturnas",
    "Adicional Nocturnidad",
    "Horas Extra 50% (por nocturnidad)",
    "Horas Extra 50%",
    "Horas Extra 100% (por nocturnidad)",
    "Horas Extra 100%",
    "Horas Feriado",
    "Horas Feriado Nocturnas",
    "Tardanza",
    "Retiro Anticipado",
]

# PART-TIME no puede tener extras "normales"
COLS_EXTRAS_BLOQUEAR_EN_PT = [
    "Horas extra",
    "Horas Extra 50%",
    "Horas Extra 100%",
]

# ================= SESSION =================
s = requests.Session()
s.headers.update({"Authorization": AUTH})

def get(url, params):
    r = s.get(url, params=params, timeout=60)
    r.raise_for_status()
    return r.json()

def _num(x):
    return pd.to_numeric(x, errors="coerce").fillna(0.0)

# ================= USERS =================
def fetch_users():
    first = get(f"{BASE}/users", {"page": 1, "limit": LIMIT_USERS})
    pages = math.ceil(first["count"] / LIMIT_USERS)

    users = first["users"]
    for p in range(2, pages + 1):
        users += get(f"{BASE}/users", {"page": p, "limit": LIMIT_USERS})["users"]

    user_map, legajo_map, employee_ids = {}, {}, []
    for u in users:
        if u.get("status") != "ACTIVE":
            continue
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

    return employee_ids, user_map, legajo_map

# ================= DAY SUMMARIES =================
def fetch_batch(emp_ids, user_map, legajo_map):
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

            # ✅ No descartar días "válidos" si worked/scheduled > 0
            hours_obj = it.get("hours") or {}
            worked = float(hours_obj.get("worked") or 0)
            scheduled = float(hours_obj.get("scheduled") or 0)

            has_useful_data = any([entries, slots, incid, tor, hol, cat, worked > 0, scheduled > 0])
            if not has_useful_data:
                continue

            # Horario obligatorio (tomamos primer timeslot)
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

            # Fichadas (entries)
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

            # Categorías
            cat_hours = split_categorized_hours(cat, CATEGORIAS)

            row = {
                "ID": emp,
                "APELLIDO, NOMBRE": user_map.get(emp, ""),
                "LEGAJO": legajo_map.get(emp, ""),
                "SCHEDULED_HOURS": scheduled,

                "FECHA": ref,
                "DIA": weekday_es(ref),

                "_weekday_api": (it.get("weekday") or "").upper().strip(),
                "_isWorkday_api": bool(it.get("isWorkday", True)),
                "_hasHoliday_api": bool(it.get("holidays") or []),

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

def build_df(employee_ids, user_map, legajo_map):
    rows = []
    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as ex:
        futures = [
            ex.submit(fetch_batch, employee_ids[i:i + BATCH_SIZE], user_map, legajo_map)
            for i in range(0, len(employee_ids), BATCH_SIZE)
        ]
        for f in as_completed(futures):
            rows.extend(f.result())

    return pd.DataFrame(rows)

# ================= BUSINESS LOGIC =================
def apply_business_logic(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()

    # 1) detectar empleados sin horario (scheduled siempre 0 en todo el período) => quedan afuera
    df["SCHEDULED_HOURS"] = pd.to_numeric(df.get("SCHEDULED_HOURS"), errors="coerce")
    max_sched = df.groupby("ID")["SCHEDULED_HOURS"].max()
    sin_horario_ids = set(max_sched[(max_sched.fillna(0) <= 0)].index)

    # 2) clasificar SOLO empleados con horario
    df_con_horario = df[~df["ID"].isin(sin_horario_ids)].copy()
    df_con_horario["SCHEDULED_HOURS"] = df_con_horario["SCHEDULED_HOURS"].fillna(0.0)

    cat_map = clasificar_empleado_por_scheduled_max(df_con_horario, "SCHEDULED_HOURS")
    df["Jornada"] = df["ID"].map(cat_map)
    df.loc[df["ID"].isin(sin_horario_ids), "Jornada"] = ""  # explícito

    # 3) tardanza / retiro
    df["TARDANZA_HORAS"] = df.apply(
        lambda r: round(max(0.0, calc_delta_hours(r["_rs"], r["_ss"], TOLERANCIA_TARDANZA_SEG)), 2),
        axis=1
    )
    df["RETIRO_ANTICIPADO_HORAS"] = df.apply(
        lambda r: round(max(0.0, calc_delta_hours(r["_se"], r["_re"], TOLERANCIA_RETIRO_SEG)), 2),
        axis=1
    )

    # 4) horas trabajadas: REGULAR + EXTRA (API)
    df["HORAS_TRABAJADAS"] = (_num(df.get("HORAS_REGULAR", 0)) + _num(df.get("HORAS_EXTRA", 0))).round(2)

    # 5) adicional nocturnidad: 8 min por hora nocturna
    df["ADICIONAL_NOCTURNIDAD"] = (_num(df.get("HORAS_NOCTURNA", 0)) * (8/60)).round(2)

    # 6) replica nocturnidad a 50/100 según día
    mask_100 = df.apply(nocturnidad_es_100, axis=1)
    df["HORAS_EXTRA_50_POR_NOCT"]  = 0.0
    df["HORAS_EXTRA_100_POR_NOCT"] = 0.0
    df.loc[mask_100,  "HORAS_EXTRA_100_POR_NOCT"] = df.loc[mask_100,  "ADICIONAL_NOCTURNIDAD"]
    df.loc[~mask_100, "HORAS_EXTRA_50_POR_NOCT"]  = df.loc[~mask_100, "ADICIONAL_NOCTURNIDAD"]
    df["HORAS_EXTRA_50_POR_NOCT"]  = _num(df["HORAS_EXTRA_50_POR_NOCT"]).round(2)
    df["HORAS_EXTRA_100_POR_NOCT"] = _num(df["HORAS_EXTRA_100_POR_NOCT"]).round(2)

    # 7) compat: si API trae EXTRA AL 100 NOCTURNA, sumarla con replica 100 por noct
    if "HORAS_EXTRA AL 100 NOCTURNA" not in df.columns:
        df["HORAS_EXTRA AL 100 NOCTURNA"] = 0.0
    df["HORAS_EXTRA AL 100 NOCTURNA"] = (_num(df["HORAS_EXTRA AL 100 NOCTURNA"]) + _num(df["HORAS_EXTRA_100_POR_NOCT"])).round(2)

    # 8) renombres finales (detalle)
    df = df.rename(columns={
        "APELLIDO, NOMBRE": "Apellido, Nombre",
        "LEGAJO": "Legajo",
        "FECHA": "Fecha",
        "DIA": "dia",
        "HORARIO_OBLIGATORIO": "Horario obligatorio",
        "FICHADAS": "Fichadas",
        "OBSERVACIONES": "Observaciones",

        "HORAS_TRABAJADAS": "Horas Trabajadas",
        "HORAS_REGULAR": "Horas Regulares",
        "HORAS_EXTRA": "Horas extra",
        "HORAS_NOCTURNA": "Horas Nocturnas",
        "ADICIONAL_NOCTURNIDAD": "Adicional Nocturnidad",

        "HORAS_EXTRA AL 50": "Horas Extra 50%",
        "HORAS_EXTRA AL 100": "Horas Extra 100%",

        "HORAS_EXTRA_50_POR_NOCT": "Horas Extra 50% (por nocturnidad)",
        "HORAS_EXTRA AL 100 NOCTURNA": "Horas Extra 100% (por nocturnidad)",

        "HORAS_HORA FERIADO": "Horas Feriado",
        "HORAS_HORA FERIADO NOCTURNA": "Horas Feriado Nocturnas",

        "TARDANZA_HORAS": "Tardanza",
        "RETIRO_ANTICIPADO_HORAS": "Retiro Anticipado",
    })



    # 9) PART-TIME: bloquear extras normales (pero nocturnidad sí)
    for c in COLS_EXTRAS_BLOQUEAR_EN_PT:
        if c not in df.columns:
            df[c] = 0.0

    mask_pt_emp = df["Jornada"].fillna("").eq("PART-TIME")
    df.loc[mask_pt_emp, COLS_EXTRAS_BLOQUEAR_EN_PT] = 0.0

    return df

# ================= RESUMEN =================
def armar_resumen(df_det: pd.DataFrame, categoria: str) -> pd.DataFrame:
    tmp = df_det.copy()
    tmp["Jornada"] = tmp["Jornada"].fillna("")
    tmp = tmp[tmp["Jornada"].eq(categoria)].copy()

    cols_out = [
        "ID", "Nombre", "Apellido", "Legajo",
        "Total Horas", "Horas Nocturnas", "Horas Extra 50%", "Horas Extra 100%",
        "Adicional Nocturno", "Horas Extra 50% (por nocturnidad)", "Horas Extra 100% (por nocturnidad)",
    ]
    if tmp.empty:
        return pd.DataFrame(columns=cols_out)

    cols_sum = [
        "Horas Trabajadas",
        "Horas Nocturnas",
        "Horas Extra 50%",
        "Horas Extra 100%",
        "Adicional Nocturnidad",
        "Horas Extra 50% (por nocturnidad)",
        "Horas Extra 100% (por nocturnidad)",
    ]
    for c in cols_sum:
        tmp[c] = _num(tmp.get(c, 0))

    tmp[["Apellido", "Nombre"]] = tmp["Apellido, Nombre"].apply(lambda x: pd.Series(split_apellido_nombre(x)))

    res = (
        tmp.groupby("ID", as_index=False)
           .agg(
                Nombre=("Nombre", "first"),
                Apellido=("Apellido", "first"),
                Legajo=("Legajo", "first"),
                **{
                    "Total Horas": ("Horas Trabajadas", "sum"),
                    "Horas Nocturnas": ("Horas Nocturnas", "sum"),
                    "Horas Extra 50%": ("Horas Extra 50%", "sum"),
                    "Horas Extra 100%": ("Horas Extra 100%", "sum"),
                    "Adicional Nocturno": ("Adicional Nocturnidad", "sum"),
                    "Horas Extra 50% (por nocturnidad)": ("Horas Extra 50% (por nocturnidad)", "sum"),
                    "Horas Extra 100% (por nocturnidad)": ("Horas Extra 100% (por nocturnidad)", "sum"),
                }
            )
    )

    for c in cols_out[4:]:
        res[c] = _num(res[c]).round(2)

    return res[cols_out].sort_values(by=["Apellido", "Nombre"])

def agregar_columnas_nuevas(resumen_part: pd.DataFrame, resumen_full: pd.DataFrame):
    # PART-TIME
    if not resumen_part.empty:
        resumen_part["Total horas + Adicional Nocturno"] = (_num(resumen_part["Total Horas"]) + _num(resumen_part["Adicional Nocturno"])).round(2)
        resumen_part["Complementarias"] = (resumen_part["Total horas + Adicional Nocturno"] - MONTHLY_CAP_PT).clip(lower=0).round(2)
    else:
        resumen_part["Total horas + Adicional Nocturno"] = pd.Series(dtype=float)
        resumen_part["Complementarias"] = pd.Series(dtype=float)

    # FULL-TIME
    if not resumen_full.empty:
        resumen_full["Horas extra totales"] = (
            _num(resumen_full["Horas Extra 50%"]) +
            _num(resumen_full["Horas Extra 100%"]) +
            _num(resumen_full["Horas Extra 50% (por nocturnidad)"]) +
            _num(resumen_full["Horas Extra 100% (por nocturnidad)"])
        ).round(2)

        resumen_full["[Total horas + Adicional nocturnidad] - Horas extra"] = (
            (_num(resumen_full["Total Horas"]) + _num(resumen_full["Adicional Nocturno"])) - _num(resumen_full["Horas extra totales"])
        ).round(2)

        resumen_full["Complementarias"] = (
            resumen_full["[Total horas + Adicional nocturnidad] - Horas extra"] - MONTHLY_CAP_FT
        ).clip(lower=0).round(2)
    else:
        resumen_full["Horas extra totales"] = pd.Series(dtype=float)
        resumen_full["[Total horas + Adicional nocturnidad] - Horas extra"] = pd.Series(dtype=float)
        resumen_full["Complementarias"] = pd.Series(dtype=float)

    return resumen_part, resumen_full

# ================= EXPORT EXCEL =================
def export_excel(df_detalle, resumen_full, resumen_part):
    out = f"reporte_{START_DATE}_a_{END_DATE}.xlsx"
    generated_at = datetime.now(TZ_AR).strftime("%d/%m/%Y %H:%M")

    # convertir horas del detalle (0 => vacío)
    df_export = df_detalle.copy()
    for c in COLS_HORAS_DETALLE:
        if c in df_export.columns:
            df_export[c] = df_export[c].apply(lambda v: horas_para_excel(v, EXPORTAR_DECIMAL))

    # convertir horas del resumen (incluye columnas nuevas)
    def export_resumen_df(r: pd.DataFrame) -> pd.DataFrame:
        r2 = r.copy()
        cols_resumen_horas = [
            "Total Horas",
            "Horas Nocturnas",
            "Horas Extra 50%",
            "Horas Extra 100%",
            "Adicional Nocturno",
            "Horas Extra 50% (por nocturnidad)",
            "Horas Extra 100% (por nocturnidad)",

            # nuevas
            "Total horas + Adicional Nocturno",
            "Horas extra totales",
            "[Total horas + Adicional nocturnidad] - Horas extra",
            "Complementarias",
        ]
        for c in cols_resumen_horas:
            if c in r2.columns:
                r2[c] = r2[c].apply(lambda v: horas_para_excel(v, EXPORTAR_DECIMAL))
        return r2

    resumen_full_export = export_resumen_df(resumen_full)
    resumen_part_export = export_resumen_df(resumen_part)

    with pd.ExcelWriter(out, engine="xlsxwriter") as writer:
        workbook = writer.book
        fmt_title = workbook.add_format({"bold": True, "font_size": 14})
        fmt_sub   = workbook.add_format({"font_size": 11})
        fmt_wrap  = workbook.add_format({"text_wrap": True})
        fmt_hhmm  = workbook.add_format({"num_format": "[h]:mm"}) if not EXPORTAR_DECIMAL else None
        fmt_dec   = workbook.add_format({"num_format": "0.00"}) if EXPORTAR_DECIMAL else None

        # ----- Detalle diario -----
        startrow = 4
        # 👉 AGREGAR ESTO
        df_export = df_export.sort_values(
            by=["ID", "Fecha"],
            ascending=[True, True]
        ).reset_index(drop=True)

        df_export.to_excel(writer, index=False, sheet_name="Detalle diario", startrow=startrow)
        ws1 = writer.sheets["Detalle diario"]
        ws1.write("A1", "DETALLE DIARIO DE ASISTENCIA", fmt_title)
        ws1.write("A2", f"Período: {START_DATE} al {END_DATE}", fmt_sub)
        ws1.write("A3", f"Generado: {generated_at}", fmt_sub)

        for idx, col in enumerate(df_export.columns):
            if col == "Observaciones":
                ws1.set_column(idx, idx, 45, fmt_wrap)
            elif col in COLS_HORAS_DETALLE:
                ws1.set_column(idx, idx, 22, fmt_dec if EXPORTAR_DECIMAL else fmt_hhmm)
            else:
                ws1.set_column(idx, idx, 26)

        # ----- Resumen PART-TIME -----
        startrow = 4
        resumen_part_export.to_excel(writer, index=False, sheet_name="Resumen PART-TIME", startrow=startrow)
        ws_pt = writer.sheets["Resumen PART-TIME"]
        ws_pt.write("A1", "REPORTE DE ASISTENCIA - RESUMEN CONSOLIDADO PART-TIME", fmt_title)
        ws_pt.write("A2", f"Período: {START_DATE} al {END_DATE}", fmt_sub)
        ws_pt.write("A3", f"Generado: {generated_at}", fmt_sub)

        # columnas numéricas resumen (formato)
        resumen_num_cols = [
            "Total Horas",
            "Horas Nocturnas",
            "Horas Extra 50%",
            "Horas Extra 100%",
            "Adicional Nocturno",
            "Horas Extra 50% (por nocturnidad)",
            "Horas Extra 100% (por nocturnidad)",
            "Total horas + Adicional Nocturno",
            "Complementarias",
        ]
        for idx, col in enumerate(resumen_part_export.columns):
            if col in resumen_num_cols:
                ws_pt.set_column(idx, idx, 22, fmt_dec if EXPORTAR_DECIMAL else fmt_hhmm)
            else:
                ws_pt.set_column(idx, idx, 26)

        # ----- Resumen FULL-TIME -----
        startrow = 4
        resumen_full_export.to_excel(writer, index=False, sheet_name="Resumen FULL-TIME", startrow=startrow)
        ws_ft = writer.sheets["Resumen FULL-TIME"]
        ws_ft.write("A1", "REPORTE DE ASISTENCIA - RESUMEN CONSOLIDADO FULL-TIME", fmt_title)
        ws_ft.write("A2", f"Período: {START_DATE} al {END_DATE}", fmt_sub)
        ws_ft.write("A3", f"Generado: {generated_at}", fmt_sub)

        resumen_num_cols_ft = [
            "Total Horas",
            "Horas Nocturnas",
            "Horas Extra 50%",
            "Horas Extra 100%",
            "Adicional Nocturno",
            "Horas Extra 50% (por nocturnidad)",
            "Horas Extra 100% (por nocturnidad)",
            "Horas extra totales",
            "[Total horas + Adicional nocturnidad] - Horas extra",
            "Complementarias",
        ]
        for idx, col in enumerate(resumen_full_export.columns):
            if col in resumen_num_cols_ft:
                ws_ft.set_column(idx, idx, 22, fmt_dec if EXPORTAR_DECIMAL else fmt_hhmm)
            else:
                ws_ft.set_column(idx, idx, 26)

    print("✅ Excel generado:", out)

# ================= MAIN =================
def main():
    employee_ids, user_map, legajo_map = fetch_users()
    print(f"Usuarios ACTIVE: {len(employee_ids)}")

    df_raw = build_df(employee_ids, user_map, legajo_map)
    print(f"Filas day-summaries: {len(df_raw)}")

    df_logic = apply_business_logic(df_raw)

    # ✅ excluir del output a los "sin horario" (Jornada vacía)
    df_logic = df_logic[df_logic["Jornada"].isin(["FULL-TIME", "PART-TIME"])].copy()

    cols_detalle = [
        "ID",
        "Apellido, Nombre",
        "Legajo",
        "Jornada",
        "Fecha",
        "dia",
        "Horario obligatorio",
        "Fichadas",
        "Observaciones",
        "Horas Trabajadas",
        "Horas Regulares",
        "Horas extra",
        "Horas Nocturnas",
        "Adicional Nocturnidad",
        "Horas Extra 50% (por nocturnidad)",
        "Horas Extra 50%",
        "Horas Extra 100% (por nocturnidad)",
        "Horas Extra 100%",
        "Horas Feriado",
        "Horas Feriado Nocturnas",
        "Tardanza",
        "Retiro Anticipado",
    ]
    cols_detalle = [c for c in cols_detalle if c in df_logic.columns]
    df_detalle = df_logic[cols_detalle].sort_values(by=["ID", "Fecha"], ascending=[True, True])

    resumen_full = armar_resumen(df_detalle, "FULL-TIME")
    resumen_part = armar_resumen(df_detalle, "PART-TIME")

    resumen_part, resumen_full = agregar_columnas_nuevas(resumen_part, resumen_full)

    export_excel(df_detalle, resumen_full, resumen_part)

if __name__ == "__main__":
    main()
