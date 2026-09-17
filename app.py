"""
Job Flow Tracker — Streamlit app with Oracle auto-sync
Run:   streamlit run app.py
Setup: copy .env.example → .env and fill in Oracle credentials
"""

import os
import sqlite3
import uuid
import io
import time
from datetime import date, datetime, timedelta
import pandas as pd
import streamlit as st
from dotenv import load_dotenv
from streamlit_autorefresh import st_autorefresh

load_dotenv()

try:
    import oracledb
    ORACLE_AVAILABLE = True
except ImportError:
    ORACLE_AVAILABLE = False

# ─── Page config ──────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Job Flow Tracker",
    page_icon="🖨️",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─── Oracle config (.env) ─────────────────────────────────────────────────────
ORA_HOST    = os.getenv("ORACLE_HOST", "10.0.3.117")
ORA_PORT    = int(os.getenv("ORACLE_PORT", "1521"))
ORA_SERVICE = os.getenv("ORACLE_SERVICE", "orcl.LSLGRP.CORP")
ORA_SID     = os.getenv("ORACLE_SID", "")  # Use if connecting via SID instead of Service Name
ORA_USER    = os.getenv("ORACLE_USER", "INTRAPRINT")
ORA_PASS    = os.getenv("ORACLE_PASSWORD", "INTRAPRINT")
ORA_MONTHS  = int(os.getenv("ORACLE_SYNC_MONTHS", "1"))

# ─── Field definitions ────────────────────────────────────────────────────────

# Oracle-sourced fields (auto-populated, shown read-only)
ORACLE_FIELDS = [
    {"key": "job_number",        "label": "Job Number",           "col": 0},
    {"key": "job_title",         "label": "Job Title",            "col": 1},
    {"key": "client_name",       "label": "Client",               "col": 2},
    {"key": "job_creation_date", "label": "Creation Date",        "col": 3},
    {"key": "job_created_by",    "label": "Created By (Initials)","col": 0},
    {"key": "date_of_delivery",  "label": "Delivery Date",        "col": 1},
    {"key": "quantity",          "label": "Quantity",             "col": 2},
    {"key": "product_code",      "label": "Product Code",         "col": 3},
    {"key": "element_numbers",   "label": "Element No(s)",        "col": 0},
    {"key": "colors_recto",      "label": "Colors Recto",         "col": 1},
    {"key": "colors_verso",      "label": "Colors Verso",         "col": 2},
    {"key": "total_sheets",      "label": "Total Sheets",         "col": 3},
    {"key": "num_makereadies",   "label": "No. of Makereadies",   "col": 0},
    {"key": "quick_paper",       "label": "Quick Paper",          "col": 1},
    {"key": "station_codes",     "label": "Production Stations",  "col": 2},
]

# Deduplicated key list for DB schema
ORACLE_KEYS = list(dict.fromkeys(f["key"] for f in ORACLE_FIELDS))

# Tracking stages (manually updated by team)
STAGES = [
    {"id": "planning",        "label": "Planning",         "icon": "🗓️", "fields": [
        {"key": "date_of_printing",  "label": "Date of Printing",  "type": "date"},
    ]},
    {"id": "prepress",        "label": "Pre-Press",        "icon": "🎨", "fields": [
        {"key": "pdf_conformity",      "label": "PDF Conformity Date", "type": "date"},
        {"key": "ctp_submission_date", "label": "CTP Submission Date", "type": "date"},
        {"key": "prepress_user",       "label": "Prepress User",       "type": "text"},
    ]},
    {"id": "workflow",        "label": "Workflow",         "icon": "⚙️", "fields": [
        {"key": "plates_manufacturing", "label": "Plates Manufacturing Date", "type": "date"},
        {"key": "plates_sent_to_press", "label": "Plates Sent to Press",      "type": "date"},
    ]},
    {"id": "store_logistics", "label": "Store & Logistics","icon": "📦", "fields": [
        {"key": "paper_store_issue",              "label": "Paper Store Issue Date",        "type": "date"},
        {"key": "paper_delivered_to_production", "label": "Paper Delivered to Production", "type": "date"},
    ]},
    {"id": "press",     "label": "Press",     "icon": "🖨️", "fields": [
        {"key": "printing_date", "label": "Printing Date", "type": "date"},
    ]},
    {"id": "cutting",   "label": "Cutting",   "icon": "✂️", "fields": [
        {"key": "cutting_date", "label": "Cutting Date", "type": "date"},
    ]},
    {"id": "finishing", "label": "Finishing", "icon": "🏁", "fields": [
        {"key": "finishing_date", "label": "Finishing Date", "type": "date"},
    ]},
    {"id": "delivery",  "label": "Delivery",  "icon": "🚚", "fields": [
        {"key": "delivered_to_store",  "label": "Delivered to Store",  "type": "date"},
        {"key": "delivered_to_client", "label": "Delivered to Client", "type": "date"},
    ]},
    {"id": "invoice",   "label": "Invoice",   "icon": "🧾", "fields": [
        {"key": "invoice_date", "label": "Invoice Date", "type": "date"},
    ]},
]

TRACKING_KEYS = [f["key"] for s in STAGES for f in s["fields"]]
DATE_KEYS     = {f["key"] for s in STAGES for f in s["fields"] if f["type"] == "date"}
ALL_DB_KEYS   = ORACLE_KEYS + TRACKING_KEYS

# Excel import column-name → field key
EXCEL_COL_MAP = {
    # Oracle source columns
    "dos_ndos": "job_number",            "job number": "job_number",
    "dos_libelle": "job_title",          "job title": "job_title",
    "dos_dcreation": "job_creation_date","creation date": "job_creation_date",
    "dos_qte": "quantity",               "quantity": "quantity",
    "dos_cproduit": "product_code",      "product code": "product_code",
    "dos_ini": "job_created_by",         "created by": "job_created_by",
                                         "created by (initials)": "job_created_by",
    "dos_date01": "date_of_delivery",    "delivery date": "date_of_delivery",
                                         "date of delivery": "date_of_delivery",
    "aux_nom": "client_name",            "client": "client_name",
    "elements": "element_numbers",       "element no(s)": "element_numbers",
                                         "element numbers": "element_numbers",
    "elm_nelement": "element_numbers",
    "colors_recto": "colors_recto",      "colors recto": "colors_recto",
    "elm_libcoulro": "colors_recto",
    "colors_verso": "colors_verso",      "colors verso": "colors_verso",
    "elm_libcoulvo": "colors_verso",
    "total_sheets": "total_sheets",      "total sheets": "total_sheets",
    "elm_ftotal": "total_sheets",
    "num_makereadies": "num_makereadies","no. of makereadies": "num_makereadies",
    "elm_nbcalages": "num_makereadies",
    "quick_paper": "quick_paper",        "quick paper": "quick_paper",
    "elm_papcrapide": "quick_paper",
    "stations": "station_codes",         "production stations": "station_codes",
    "station_codes": "station_codes",    "gam_cposte": "station_codes",
    # Tracking columns
    "date of printing": "date_of_printing",
    "pdf conformity date": "pdf_conformity",  "pdf conformity": "pdf_conformity",
    "ctp submission date": "ctp_submission_date",
    "prepress user": "prepress_user",
    "plates manufacturing date": "plates_manufacturing",
    "plates manufacturing": "plates_manufacturing",
    "plates sent to press": "plates_sent_to_press",
    "plates sent to press dept": "plates_sent_to_press",
    "paper store issue date": "paper_store_issue",
    "paper store issue": "paper_store_issue",
    "paper delivered to production": "paper_delivered_to_production",
    "printing date": "printing_date",
    "cutting date": "cutting_date",
    "finishing date": "finishing_date",
    "delivered to store": "delivered_to_store",
    "delivered to store date": "delivered_to_store",
    "delivered to client": "delivered_to_client",
    "delivered to client date": "delivered_to_client",
    "invoice date": "invoice_date",
}

# ─── Database ──────────────────────────────────────────────────────────────────
DB_PATH = "jobs.db"

@st.cache_resource
def get_db():
    conn = sqlite3.connect(DB_PATH, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    extra = ",\n            ".join(
        f"{k} TEXT DEFAULT ''" for k in ALL_DB_KEYS if k != "job_number"
    )
    conn.execute(f"""
        CREATE TABLE IF NOT EXISTS jobs (
            id               TEXT PRIMARY KEY,
            job_number        TEXT UNIQUE NOT NULL,
            {extra},
            source           TEXT DEFAULT 'manual',
            last_oracle_sync TEXT DEFAULT '',
            created_at       TEXT DEFAULT (datetime('now'))
        )
    """)
    conn.commit()
    return conn

def load_jobs() -> list[dict]:
    return [dict(r) for r in get_db().execute(
        "SELECT * FROM jobs ORDER BY created_at DESC"
    ).fetchall()]

def upsert_job(job: dict):
    db   = get_db()
    keys = ["id", "job_number"] + [k for k in ALL_DB_KEYS if k != "job_number" and k in job]
    ph   = ", ".join(f":{k}" for k in keys)
    ups  = ", ".join(f"{k}=:{k}" for k in keys if k not in ("id", "job_number"))
    db.execute(
        f"INSERT INTO jobs ({','.join(keys)}) VALUES ({ph}) "
        f"ON CONFLICT(job_number) DO UPDATE SET {ups}",
        {k: job.get(k, "") for k in keys}
    )
    db.commit()

def delete_job(job_id: str):
    db = get_db()
    db.execute("DELETE FROM jobs WHERE id=?", (job_id,))
    db.commit()

def clear_all_jobs():
    db = get_db()
    db.execute("DELETE FROM jobs")
    db.commit()

def update_oracle_fields(job: dict):
    """Update only Oracle-sourced fields for an existing job, preserving all tracking data."""
    db = get_db()
    update_keys = [k for k in ORACLE_KEYS if k != "job_number" and k in job]
    update_keys.append("last_oracle_sync")
    sets = ", ".join(f"{k}=:{k}" for k in update_keys)
    db.execute(
        f"UPDATE jobs SET {sets} WHERE job_number=:job_number",
        {k: job.get(k, "") for k in update_keys + ["job_number"]}
    )
    db.commit()

# ─── Oracle sync ──────────────────────────────────────────────────────────────

# Aggregated SQL — protected against concatenation string overflow
ORACLE_SQL = """
SELECT
    d.DOS_NDOS,
    d.DOS_LIBELLE,
    TO_CHAR(d.DOS_DCREATION, 'YYYY-MM-DD')  AS DOS_DCREATION,
    d.DOS_QTE,
    d.DOS_CPRODUIT,
    d.DOS_INI,
    TO_CHAR(d.DOS_DATE01,    'YYYY-MM-DD')  AS DOS_DATE01,
    a.AUX_NOM,
    LISTAGG(DISTINCT TO_CHAR(e.ELM_NELEMENT), ', ' ON OVERFLOW TRUNCATE '...')
        WITHIN GROUP (ORDER BY e.ELM_NELEMENT)  AS ELEMENTS,
    LISTAGG(DISTINCT TRIM(e.ELM_LIBCOULRO),  ', ' ON OVERFLOW TRUNCATE '...')
        WITHIN GROUP (ORDER BY e.ELM_NELEMENT)  AS COLORS_RECTO,
    LISTAGG(DISTINCT TRIM(e.ELM_LIBCOULVO),  ', ' ON OVERFLOW TRUNCATE '...')
        WITHIN GROUP (ORDER BY e.ELM_NELEMENT)  AS COLORS_VERSO,
    MAX(e.ELM_FTOTAL)                           AS TOTAL_SHEETS,
    MAX(e.ELM_NBCALAGES)                        AS NUM_MAKEREADIES,
    MAX(e.ELM_PAPCRAPIDE)                       AS QUICK_PAPER,
    LISTAGG(DISTINCT TRIM(g.GAM_CPOSTE),     ', ' ON OVERFLOW TRUNCATE '...')
        WITHIN GROUP (ORDER BY g.GAM_CPOSTE)    AS STATIONS
FROM       INTRAPRINT.F_DOSSIER  d
INNER JOIN INTRAPRINT.F_AUX      a ON d.DOS_CAUX    = a.AUX_CAUX
INNER JOIN INTRAPRINT.F_FABGAM   g ON d.DOS_NDOS    = g.GAM_NDOS
INNER JOIN INTRAPRINT.F_ELEMFAB  e ON g.GAM_NDOS    = e.ELM_NDOSSIER
                                  AND g.GAM_ELEMENT = e.ELM_NELEMENT
WHERE d.DOS_DCREATION >= :start_date
  AND d.DOS_DCREATION <  :end_date
  AND d.DOS_CPRODUIT  >= 'A99'
  AND d.DOS_CPRODUIT  <= 'S99'
GROUP BY
    d.DOS_NDOS, d.DOS_LIBELLE, d.DOS_DCREATION,
    d.DOS_QTE,  d.DOS_CPRODUIT, d.DOS_INI, d.DOS_DATE01, a.AUX_NOM
ORDER BY d.DOS_DCREATION DESC
"""

ORA_COL_MAP = {
    "dos_ndos":       "job_number",
    "dos_libelle":    "job_title",
    "dos_dcreation":  "job_creation_date",
    "dos_qte":        "quantity",
    "dos_cproduit":   "product_code",
    "dos_ini":        "job_created_by",
    "dos_date01":     "date_of_delivery",
    "aux_nom":        "client_name",
    "elements":       "element_numbers",
    "colors_recto":   "colors_recto",
    "colors_verso":   "colors_verso",
    "total_sheets":   "total_sheets",
    "num_makereadies":"num_makereadies",
    "quick_paper":    "quick_paper",
    "stations":       "station_codes",
}

@st.cache_resource
def init_oracle_pool():
    if not ORACLE_AVAILABLE:
        return None
    if not all([ORA_HOST, ORA_USER, ORA_PASS]):
        return None
    try:
        kwargs = {
            "user": ORA_USER,
            "password": ORA_PASS,
            "host": ORA_HOST,
            "port": ORA_PORT,
            "min": 1,
            "max": 5,
            "increment": 1
        }
        if ORA_SID:
            kwargs["sid"] = ORA_SID
        else:
            kwargs["service_name"] = ORA_SERVICE
            
        return oracledb.create_pool(**kwargs)
    except Exception as e:
        st.error(f"Oracle Connection Pool Initialization Failed: {e}")
        return None

def sync_from_oracle(months_back: int = ORA_MONTHS) -> tuple:
    """
    Pull jobs from Oracle: insert new ones, update Oracle fields on existing ones.
    Returns (new_count, updated_count, skipped_count, error_msg).
    """
    pool = init_oracle_pool()
    if not pool:
        return 0, 0, 0, "Oracle driver or connection environment incomplete."

    try:
        end   = datetime.now()
        start = end - timedelta(days=months_back * 31)

        with pool.acquire() as conn:
            with conn.cursor() as cur:
                cur.execute(ORACLE_SQL, {"start_date": start, "end_date": end})
                cols = [c[0].lower() for c in cur.description]
                rows = cur.fetchall()

        # Build lookup of existing jobs: job_number → id
        existing_jobs = {j.get("job_number", ""): j for j in load_jobs()}
        new_count = updated_count = skipped = 0
        now_str = datetime.now().isoformat()

        for row in rows:
            raw = dict(zip(cols, row))
            job = {"source": "oracle", "last_oracle_sync": now_str}
            for ora_col, db_key in ORA_COL_MAP.items():
                val = raw.get(ora_col)
                job[db_key] = "" if val is None else str(val).strip()

            num = job.get("job_number", "")
            if not num:
                skipped += 1
                continue

            try:
                if num in existing_jobs:
                    # Update Oracle-sourced fields only — tracking data is preserved
                    update_oracle_fields(job)
                    updated_count += 1
                else:
                    # Brand-new job from Oracle
                    job["id"] = str(uuid.uuid4())
                    upsert_job(job)
                    new_count += 1
            except Exception:
                skipped += 1

        return new_count, updated_count, skipped, ""
    except Exception as e:
        return 0, 0, 0, str(e)

# ─── Helpers ──────────────────────────────────────────────────────────────────
def stage_progress(job: dict) -> int:
    for i in range(len(STAGES)-1, -1, -1):
        if any(job.get(f["key"], "") for f in STAGES[i]["fields"]):
            return i
    return -1

def status_label(job: dict) -> str:
    i = stage_progress(job)
    if i == len(STAGES)-1: return "✅ Complete"
    if i >= 0:             return f"🔵 {STAGES[i]['label']}"
    return "⚪ New"

def progress_pct(job: dict) -> float:
    i = stage_progress(job)
    return 0.0 if i < 0 else (i+1)/len(STAGES)

def parse_date(val) -> "date | None":
    if not val or str(val).strip() in ("", "nan", "NaT", "None"): return None
    if isinstance(val, datetime): return val.date()
    if isinstance(val, date):     return val
    try: return datetime.strptime(str(val)[:10], "%Y-%m-%d").date()
    except: return None

def to_str(val) -> str:
    if val is None: return ""
    if isinstance(val, (date, datetime)): return val.strftime("%Y-%m-%d")
    return str(val)

def safe_date(v) -> str:
    if pd.isna(v) if hasattr(pd, 'isna') else False: return ""
    if str(v).strip() in ("", "nan", "NaT", "None"): return ""
    if isinstance(v, (date, datetime)):
        d = v if isinstance(v, datetime) else datetime.combine(v, datetime.min.time())
        return d.strftime("%Y-%m-%d")
    try: return pd.to_datetime(v).strftime("%Y-%m-%d")
    except: return str(v).strip()

# ─── CSS ──────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
[data-testid="stAppViewContainer"] > .main { background:#F0EDE8; }
[data-testid="stSidebar"]           { background:#1A2130 !important; }
[data-testid="stSidebar"] p,
[data-testid="stSidebar"] label,
[data-testid="stSidebar"] span      { color:#C8C4BC !important; }
[data-testid="stSidebar"] h1,
[data-testid="stSidebar"] h2,
[data-testid="stSidebar"] h3        { color:#E8E4DC !important; }
[data-testid="stSidebar"] [data-testid="stMetricValue"] { color:#FFF !important; }
[data-testid="stSidebar"] div[role="radiogroup"] label  { color:#E8E4DC !important; }
.stTabs [data-baseweb="tab"] { font-size:0.78rem; padding:6px 10px; }
button[kind="primary"]       { background-color:#1C5FA5 !important; }
.info-chip { background:#EEF2FA; border-radius:6px; padding:8px 12px; margin-bottom:6px; }
.info-lbl  { font-size:0.68rem; font-weight:700; text-transform:uppercase;
             letter-spacing:0.05em; color:#6B7280; margin-bottom:2px; }
.info-val  { font-size:0.88rem; color:#1A1A1A; font-weight:500; }
</style>
""", unsafe_allow_html=True)

# ─── Auto-refresh every 2 minutes ────────────────────────────────────────────
SYNC_INTERVAL_MS  = 2 * 60 * 1000   # 2 minutes in milliseconds
SYNC_INTERVAL_SEC = SYNC_INTERVAL_MS // 1000

refresh_count = st_autorefresh(interval=SYNC_INTERVAL_MS, key="oracle_auto_sync")

# ─── Sidebar ──────────────────────────────────────────────────────────────────
oracle_configured = ORACLE_AVAILABLE and all([ORA_HOST, ORA_USER, ORA_PASS])

def run_sync(source: str = "auto"):
    """Run Oracle sync and store result in session state."""
    n, u, s, err = sync_from_oracle()
    st.session_state["last_sync_time"]   = time.time()
    st.session_state["last_sync_new"]    = n
    st.session_state["last_sync_updated"]= u
    st.session_state["last_sync_err"]    = err
    st.session_state["last_sync_source"] = source
    return n, u, s, err

with st.sidebar:
    st.markdown("## 🖨️ Job Flow Tracker")

    if oracle_configured:
        target_name = ORA_SID if ORA_SID else ORA_SERVICE
        st.markdown(f"🟢 **Oracle** connected  \n`{target_name}@{ORA_HOST}`")
    else:
        st.markdown("🔴 **Oracle** — not configured")
        st.caption("Edit `.env` with your credentials")

    st.divider()
    page = st.radio("nav",
                    ["📋  All Jobs", "📥  Import Excel", "📤  Export"],
                    label_visibility="collapsed")
    st.divider()

    if oracle_configured:
        # Manual sync button
        if st.button("🔄  Sync Oracle Now", use_container_width=True, type="primary"):
            with st.spinner("Syncing from Oracle…"):
                n, u, s, err = run_sync("manual")
            if err:
                st.error(f"Sync error: {err}")
            else:
                st.success(f"✅ {n} new · {u} updated · {s} skipped")
                st.rerun()

        # Auto-sync status
        last_t   = st.session_state.get("last_sync_time", 0)
        last_n   = st.session_state.get("last_sync_new", 0)
        last_u   = st.session_state.get("last_sync_updated", 0)
        last_err = st.session_state.get("last_sync_err", "")
        last_src = st.session_state.get("last_sync_source", "")

        if last_t:
            elapsed   = int(time.time() - last_t)
            next_sync = max(0, SYNC_INTERVAL_SEC - elapsed)
            mins, secs = divmod(next_sync, 60)
            sync_time_str = datetime.fromtimestamp(last_t).strftime("%H:%M:%S")
            if last_err:
                st.caption(f"⚠️ Last sync failed at {sync_time_str}")
            else:
                src_icon = "👆" if last_src == "manual" else "⏱️"
                st.caption(
                    f"{src_icon} Last sync: {sync_time_str}  \n"
                    f"&nbsp;&nbsp;&nbsp;↳ {last_n} new · {last_u} updated  \n"
                    f"⏭️ Next sync in **{mins}m {secs:02d}s**"
                )
        else:
            st.caption(f"⏱️ Auto-syncs every {SYNC_INTERVAL_SEC//60} min")

        st.divider()

    # Clear database
    with st.expander("🗑️ Clear Database"):
        st.warning("Permanently deletes **all** jobs from the tracker.")
        if st.button("Delete All Jobs", key="clear_all_btn", type="primary",
                     use_container_width=True):
            clear_all_jobs()
            st.session_state.pop("last_sync_time", None)
            st.toast("All jobs cleared.", icon="🗑️")
            st.rerun()

    st.divider()
    jobs = load_jobs()
    ca, cb = st.columns(2)
    ca.metric("Total",    len(jobs))
    cb.metric("Complete", sum(1 for j in jobs if stage_progress(j)==len(STAGES)-1))
    st.metric("In Progress", sum(1 for j in jobs if 0<=stage_progress(j)<len(STAGES)-1))

# ─── Sync logic: initial load + every auto-refresh tick ──────────────────────
if oracle_configured:
    prev_count = st.session_state.get("prev_refresh_count", -1)

    # On first load (refresh_count == 0) or on every auto-refresh tick
    if refresh_count != prev_count:
        st.session_state["prev_refresh_count"] = refresh_count
        n, u, s, err = run_sync("auto" if refresh_count > 0 else "startup")
        if not err and n > 0:
            st.toast(f"🔄 {n} new job(s) added from Oracle", icon="✅")
        if not err and u > 0:
            st.toast(f"🔁 {u} job(s) updated from Oracle", icon="🔁")

jobs = load_jobs()

# ══════════════════════════════════════════════════════════════════════════════
# PAGE: All Jobs
# ══════════════════════════════════════════════════════════════════════════════
if "All Jobs" in page:
    st.title("All Jobs")

    c1, c2, c3 = st.columns([3, 1, 1])
    search   = c1.text_input("Search", placeholder="Job #, title, client…",
                              label_visibility="collapsed")
    f_stat   = c2.selectbox("Status", ["All", "New", "In Progress", "Complete"],
                              label_visibility="collapsed")
    f_src    = c3.selectbox("Source", ["All", "Oracle", "Manual"],
                              label_visibility="collapsed")

    flt = jobs
    if search:
        s = search.lower()
        flt = [j for j in flt if
               s in (j.get("job_number", "") or "").lower() or
               s in (j.get("job_title", "")  or "").lower() or
               s in (j.get("client_name", "") or "").lower()]
    if f_stat == "New":           flt = [j for j in flt if stage_progress(j)==-1]
    elif f_stat == "In Progress": flt = [j for j in flt if 0<=stage_progress(j)<len(STAGES)-1]
    elif f_stat == "Complete":    flt = [j for j in flt if stage_progress(j)==len(STAGES)-1]
    if f_src == "Oracle":         flt = [j for j in flt if j.get("source")=="oracle"]
    elif f_src == "Manual":       flt = [j for j in flt if j.get("source")!="oracle"]

    if not flt:
        st.info("No jobs found. Oracle syncs automatically, or use Import Excel.")
    else:
        st.dataframe(pd.DataFrame([{
            "Job #":     j.get("job_number", ""),
            "Title":     j.get("job_title", "—") or "—",
            "Client":    j.get("client_name", "—") or "—",
            "Qty":       j.get("quantity", "—") or "—",
            "Delivery":  j.get("date_of_delivery", "—") or "—",
            "Status":    status_label(j),
            "Progress":  f"{round(progress_pct(j)*100)}%",
            "Source":    "🏢 Oracle" if j.get("source")=="oracle" else "✍️ Manual",
        } for j in flt]), use_container_width=True, hide_index=True)

        st.divider()
        st.subheader("Update a Job")

        sel = st.selectbox("Pick a job to update",
                           ["— select —"] + [j.get("job_number", "") for j in flt],
                           label_visibility="collapsed")

        if sel and sel != "— select —":
            job = next((j for j in jobs if j.get("job_number")==sel), None)
            if job:
                pct_val = progress_pct(job)
                s_idx   = stage_progress(job)

                st.markdown(
                    f"**{job.get('job_number','')}** &nbsp;·&nbsp; "
                    f"{job.get('job_title','')} &nbsp;·&nbsp; {status_label(job)}"
                )
                st.progress(
                    pct_val,
                    text=f"{round(pct_val*100)}% complete  ·  "
                         f"{'Stage '+str(s_idx+1)+'/'+str(len(STAGES)) if s_idx>=0 else 'Not started'}"
                )

                # ── Oracle info card ──────────────────────────────────────
                with st.expander("📋 Job Details (from Oracle)", expanded=True):
                    INFO = [
                        ("client_name",        "Client"),
                        ("job_creation_date",  "Creation Date"),
                        ("job_created_by",     "Created By"),
                        ("date_of_delivery",   "Delivery Date"),
                        ("quantity",           "Quantity"),
                        ("product_code",       "Product Code"),
                        ("element_numbers",    "Element No(s)"),
                        ("colors_recto",       "Colors Recto"),
                        ("colors_verso",       "Colors Verso"),
                        ("total_sheets",       "Total Sheets"),
                        ("num_makereadies",    "Makereadies"),
                        ("quick_paper",        "Quick Paper"),
                        ("station_codes",      "Stations"),
                        ("job_title",          "Job Title"),
                    ]
                    cols4 = st.columns(4)
                    for i, (k, lbl) in enumerate(INFO):
                        v = job.get(k, "") or "—"
                        cols4[i % 4].markdown(
                            f'<div class="info-chip">'
                            f'<div class="info-lbl">{lbl}</div>'
                            f'<div class="info-val">{v}</div></div>',
                            unsafe_allow_html=True
                        )

                st.write("")

                # ── Tracking stage tabs ───────────────────────────────────
                tabs = st.tabs([f"{s['icon']} {s['label']}" for s in STAGES])
                for stage, tab in zip(STAGES, tabs):
                    done = all(job.get(f["key"], "") for f in stage["fields"])
                    with tab:
                        if done:
                            st.success("Stage complete ✓")
                        with st.form(f"frm_{job['id']}_{stage['id']}"):
                            vals  = {}
                            ncols = min(len(stage["fields"]), 2)
                            fcols = st.columns(ncols)
                            for i, field in enumerate(stage["fields"]):
                                lbl  = field["label"]
                                wkey = f"{job['id']}_{stage['id']}_{field['key']}"
                                with fcols[i % ncols]:
                                    if field["type"] == "date":
                                        vals[field["key"]] = st.date_input(
                                            lbl,
                                            value=parse_date(job.get(field["key"])),
                                            format="DD/MM/YYYY", key=wkey)
                                    else:
                                        vals[field["key"]] = st.text_input(
                                            lbl,
                                            value=job.get(field["key"], "") or "",
                                            key=wkey)

                            if st.form_submit_button("💾 Save stage", type="primary",
                                                     use_container_width=True):
                                updated = {**job, **{
                                    k: to_str(v) if isinstance(v, date) else (v or "")
                                    for k, v in vals.items()
                                }}
                                upsert_job(updated)
                                st.success("Saved!")
                                st.rerun()

                # ── Delete ────────────────────────────────────────────────
                with st.expander("⚠️ Danger zone"):
                    st.warning(f"Delete **{job.get('job_number','')}**? Cannot be undone.")
                    if st.button("Yes, delete this job", key=f"del_{job['id']}"):
                        delete_job(job["id"])
                        st.success("Deleted.")
                        st.rerun()

# ══════════════════════════════════════════════════════════════════════════════
# PAGE: Import Excel
# ══════════════════════════════════════════════════════════════════════════════
elif "Import" in page:
    st.title("Import from Excel")
    st.info(
        "Use this as a fallback when Oracle is unavailable, or to load historical data.  \n"
        "Accepted column names: Oracle field names (e.g. `DOS_NDOS`) **or** human labels "
        "(e.g. `Job Number`). The file may also be a previous export from this app."
    )

    uploaded = st.file_uploader("Choose an Excel file (.xlsx / .xls)", type=["xlsx", "xls"])
    if uploaded:
        try:
            file_bytes = io.BytesIO(uploaded.getvalue())
            raw = pd.read_excel(file_bytes, header=None)

            # Find header row
            header_row = None
            job_num_aliases = {k for k, v in EXCEL_COL_MAP.items() if v == "job_number"}
            for idx, row in raw.iterrows():
                if any(str(c).strip().lower() in job_num_aliases for c in row):
                    header_row = idx
                    break

            if header_row is None:
                st.error(
                    "Cannot find a Job Number column. "
                    "Expected a column named 'Job Number', 'DOS_NDOS', or 'Job #'."
                )
            else:
                file_bytes.seek(0)
                df = pd.read_excel(file_bytes, header=header_row)
                df.columns = df.columns.map(lambda c: str(c).strip().lower())

                # Rename using map
                df = df.rename(columns={k: v for k, v in EXCEL_COL_MAP.items() if k in df.columns})
                df = df[[c for c in df.columns if c in ALL_DB_KEYS]]

                if "job_number" not in df.columns:
                    st.error("No Job Number column found after mapping.")
                    st.stop()

                df["job_number"] = df["job_number"].astype(str).str.strip()
                df = df[df["job_number"].isin(
                    [v for v in df["job_number"] if v not in ("", "nan", "None")]
                )]

                # Convert date columns
                for col in DATE_KEYS:
                    if col in df.columns:
                        df[col] = df[col].apply(safe_date)

                df = df.fillna("").astype(str).replace({"nan": "", "NaT": "", "None": ""})

                # Duplicate detection
                existing = {j.get("job_number", "") for j in load_jobs()}
                df["⚠️ Exists"] = df["job_number"].isin(existing)
                n_new = (~df["⚠️ Exists"]).sum()
                n_dup =   df["⚠️ Exists"].sum()

                st.write(f"**{len(df)} rows found** — {n_new} new, {n_dup} already in tracker")

                preview_cols = [c for c in
                    ["job_number", "job_title", "client_name", "date_of_delivery", "⚠️ Exists"]
                    if c in df.columns or c == "⚠️ Exists"]
                st.dataframe(
                    df[preview_cols].rename(columns={
                        "job_number": "Job #", "job_title": "Title",
                        "client_name": "Client", "date_of_delivery": "Delivery"
                    }),
                    use_container_width=True, hide_index=True
                )

                dup_mode = "Skip"
                if n_dup > 0:
                    choice = st.radio(
                        f"**{n_dup} duplicate(s)** — what should happen?",
                        ["Skip duplicates (keep existing data)",
                         "Replace with imported data"],
                        horizontal=True,
                    )
                    dup_mode = "Skip" if "Skip" in choice else "Replace"

                if st.button("✅ Confirm Import", type="primary"):
                    imported = 0
                    for _, row in df.iterrows():
                        if row.get("⚠️ Exists") and dup_mode == "Skip":
                            continue
                        data = {k: v for k, v in row.items()
                                if k in ALL_DB_KEYS and v not in ("", "nan", "NaT", "None")}
                        data["id"]     = str(uuid.uuid4())
                        data["source"] = "manual"
                        try:
                            upsert_job(data)
                            imported += 1
                        except Exception as ex:
                            st.warning(f"Skipped {row.get('job_number','?')}: {ex}")
                    st.success(f"✅ Imported {imported} job(s).")
                    st.rerun()

        except Exception as e:
            st.error(f"Failed to read file: {e}")

# ══════════════════════════════════════════════════════════════════════════════
# PAGE: Export
# ══════════════════════════════════════════════════════════════════════════════
elif "Export" in page:
    st.title("Export Jobs")

    if not jobs:
        st.info("No jobs to export yet.")
    else:
        lbl_map = {f["key"]: f["label"] for f in ORACLE_FIELDS}
        lbl_map.update({f["key"]: f["label"] for s in STAGES for f in s["fields"]})

        rows = []
        for j in jobs:
            row = {}
            for k in ORACLE_KEYS:
                row[lbl_map.get(k, k)] = j.get(k, "") or ""
            for s in STAGES:
                for f in s["fields"]:
                    row[f["label"]] = j.get(f["key"], "") or ""
            row["Status"]     = status_label(j).replace("✅ ", "").replace("🔵 ", "").replace("⚪ ", "")
            row["Progress %"] = round(progress_pct(j)*100)
            row["Source"]     = j.get("source", "manual")
            rows.append(row)

        df_out = pd.DataFrame(rows)
        st.dataframe(df_out, use_container_width=True, hide_index=True)

        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as writer:
            df_out.to_excel(writer, index=False, sheet_name="Job Tracker")
        buf.seek(0)

        st.download_button(
            "📥 Download as Excel",
            data=buf.getvalue(),
            file_name=f"job_tracker_{date.today().isoformat()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True,
        )