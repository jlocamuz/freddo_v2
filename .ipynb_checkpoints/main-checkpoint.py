import os, math, requests, pandas as pd
from concurrent.futures import ThreadPoolExecutor, as_completed

BASE = "https://api-prod.humand.co/public/api/v1"
AUTH = "Basic Njc3NDUzMjpkNWR2Z1pzNXQ3ZEZ2XzE2Z2pfbV9XNklpVFNPU0NmMQ=="

START_DATE = "2025-11-17"
END_DATE   = "2025-12-16"

LIMIT_USERS = 50
LIMIT_DAYS  = 500
BATCH_SIZE  = 25
MAX_WORKERS = 8

# ================= SESSION =================
s = requests.Session()
s.headers.update({"Authorization": AUTH})

def get(url, params):
    r = s.get(url, params=params, timeout=60)
    r.raise_for_status()
    return r.json()

def flatten(d, p=""):
    o = {}
    for k, v in d.items():
        k2 = f"{p}.{k}" if p else k
        if isinstance(v, dict):
            o.update(flatten(v, k2))
        elif isinstance(v, list):
            o[k2] = str(v)
        else:
            o[k2] = v
    return o

# ================= USERS =================
first = get(f"{BASE}/users", {"page": 1, "limit": LIMIT_USERS})
pages = math.ceil(first["count"] / LIMIT_USERS)

users = first["users"]
for p in range(2, pages + 1):
    users += get(f"{BASE}/users", {"page": p, "limit": LIMIT_USERS})["users"]

employee_ids = [
    u["employeeInternalId"]
    for u in users
    if u["status"] == "ACTIVE"
]

print(f"Usuarios ACTIVE: {len(employee_ids)}")

# ================= DAY SUMMARIES (COMMA-SEPARATED) =================
def fetch_batch(emp_ids):
    rows = []
    page = 1
    emp_param = ",".join(emp_ids)   # 👈 CLAVE

    while True:
        data = get(
            f"{BASE}/time-tracking/day-summaries",
            {
                "employeeIds": emp_param,
                "startDate": START_DATE,
                "endDate": END_DATE,
                "limit": LIMIT_DAYS,
                "page": page,
            }
        )

        items = data.get("items", [])
        if not items:
            break

        for it in items:
            rows.append(flatten(it))

        if len(items) < LIMIT_DAYS:
            break

        page += 1

    return rows

# ================= EXEC =================
all_rows = []

with ThreadPoolExecutor(max_workers=MAX_WORKERS) as ex:
    futures = [
        ex.submit(fetch_batch, employee_ids[i:i+BATCH_SIZE])
        for i in range(0, len(employee_ids), BATCH_SIZE)
    ]

    for f in as_completed(futures):
        all_rows.extend(f.result())

print(f"TOTAL FILAS: {len(all_rows)}")

df = pd.DataFrame(all_rows)
df = df.sort_values(
    by=["employeeId", "referenceDate"],  # si la columna se llama "date"
    ascending=[True, True]
)
df.to_excel("day_summaries_2025-11-17_a_2025-12-16.xlsx", index=False)

print("✅ Excel generado")
