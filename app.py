# -*- coding: utf-8 -*-
import math
import re
import hmac
import itertools
from pathlib import Path

import pandas as pd
import numpy as np
import streamlit as st

# -----------------------------
# Page config
# -----------------------------
st.set_page_config(
    page_title="מנוע תמהילי קרנות השתלמות",
    page_icon="📊",
    layout="wide",
)

# -----------------------------
# RTL + Theme (desktop light, mobile dark) + UI polish
# -----------------------------
st.markdown(
    """
<style>
/* RTL base */
html, body, [class*="css"]  { direction: rtl; text-align: right; }

/* Tabs RTL */
div[data-baseweb="tab-list"] { direction: rtl; }

/* Keep slider track stable: force LTR for slider widget itself */
div[data-testid="stSlider"] { direction: ltr; }
div[data-testid="stSlider"] label { direction: rtl; text-align: right; font-weight: 600; }

/* Wide layout padding */
.block-container { padding-top: 1.2rem; padding-bottom: 2rem; }

/* KPI cards */
.kpi-wrap { display: grid; grid-template-columns: repeat(3, 1fr); gap: 12px; }
.kpi {
  border-radius: 16px;
  padding: 14px 16px;
  border: 1px solid rgba(0,0,0,0.10);
  background: rgba(255,255,255,0.92);
}
.kpi h4 { margin: 0 0 6px 0; font-size: 0.95rem; opacity: 0.9; }
.kpi .big { font-size: 1.35rem; font-weight: 700; line-height: 1.2; }
.kpi .sub { margin-top: 6px; font-size: 0.9rem; opacity: 0.85; }

/* Table RTL */
div[data-testid="stDataFrame"], table, thead, tbody, tr, th, td { direction: rtl !important; }
thead tr th, tbody tr td { text-align: right !important; }

/* Mobile dark mode */
@media (max-width: 768px) {
  body { background: #0b0f14; color: #e6edf3; }
  .kpi { background: rgba(16, 23, 33, 0.92); border: 1px solid rgba(255,255,255,0.10); }
  .stMarkdown, .stText, .stCaption, .stAlert, label, p, span, div { color: #e6edf3 !important; }
  input, textarea { background: #121a24 !important; color: #e6edf3 !important; }
}
</style>
""",
    unsafe_allow_html=True,
)

# -----------------------------
# Embedded data (users cannot change)
# -----------------------------
APP_DIR = Path(__file__).resolve().parent
DATA_DIR = APP_DIR / "data"
HOLDINGS_XLSX = DATA_DIR / "קרנות השתלמות פברואר 2026.xlsx"
SERVICE_XLSX = DATA_DIR / "ציוני שירות.xlsx"

# -----------------------------
# Helpers
# -----------------------------
PARAM_ALIASES = {
    "stocks": [r"חשיפה\s*למניות", r"סך\s*חשיפה\s*למניות"],
    "foreign": [r"מושקעים\s*בחו\"ל", r"בחו\"ל", r"חשיפה\s*לחו\"ל"],
    "illiquid": [r"לא\s*סחיר", r"נכסים\s*לא\s*סחירים"],
    "fx": [r"מט\"ח", r"חשיפה\s*למט\"ח"],
    "sharpe": [r"שארפ", r"Sharpe"],
}

def parse_pct(x):
    if x is None or (isinstance(x, float) and np.isnan(x)):
        return np.nan
    if isinstance(x, (int, float)) and not isinstance(x, bool):
        return float(x)
    s = str(x).strip()
    s = s.replace(",", "")
    if s.endswith("%"):
        s = s[:-1].strip()
    s = s.replace("−", "-")
    try:
        return float(s)
    except:
        return np.nan

def normalize_param_name(param):
    s = str(param).strip()
    for key, pats in PARAM_ALIASES.items():
        for p in pats:
            if re.search(p, s):
                return key
    return None

def extract_manager_name(col_name: str) -> str:
    s = str(col_name).strip()
    if "קרן השתלמות" in s:
        left = s.split("קרן השתלמות")[0].strip()
        return left if left else s
    if "השתלמות" in s:
        left = s.split("השתלמות")[0].strip()
        return left if left else s
    parts = s.split()
    return " ".join(parts[:2]) if parts else s

def load_holdings_excel(path: Path) -> tuple[pd.DataFrame, list[str]]:
    xl = pd.ExcelFile(path)
    sheets = xl.sheet_names
    rows = []
    for sh in sheets:
        df = xl.parse(sh)
        if df.empty:
            continue
        param_col = None
        for c in df.columns:
            if str(c).strip() == "פרמטר":
                param_col = c
                break
        if param_col is None:
            param_col = df.columns[0]

        param_keys = df[param_col].apply(normalize_param_name)
        idx_map = {}
        for i, k in enumerate(param_keys.tolist()):
            if k and k not in idx_map:
                idx_map[k] = i

        if "foreign" not in idx_map or "illiquid" not in idx_map:
            continue

        for c in df.columns:
            if c == param_col:
                continue
            series = df[c]
            rec = {
                "sheet": sh,
                "fund_name": str(c).strip(),
                "manager": extract_manager_name(c),
                "stocks": parse_pct(series.iloc[idx_map["stocks"]]) if "stocks" in idx_map else np.nan,
                "foreign": parse_pct(series.iloc[idx_map["foreign"]]) if "foreign" in idx_map else np.nan,
                "illiquid": parse_pct(series.iloc[idx_map["illiquid"]]) if "illiquid" in idx_map else np.nan,
                "fx": parse_pct(series.iloc[idx_map["fx"]]) if "fx" in idx_map else np.nan,
                "sharpe": parse_pct(series.iloc[idx_map["sharpe"]]) if "sharpe" in idx_map else np.nan,
            }
            rec["israel"] = (100.0 - rec["foreign"]) if not np.isnan(rec["foreign"]) else np.nan  # Israel rule

            if all(np.isnan(rec[k]) for k in ["stocks","foreign","illiquid","fx","sharpe"]):
                continue
            rows.append(rec)

    out = pd.DataFrame(rows)
    for k in ["stocks","foreign","illiquid","fx","sharpe","israel"]:
        if k in out.columns:
            out[k] = pd.to_numeric(out[k], errors="coerce")
    return out, sheets

def load_service_scores_excel(path: Path) -> dict:
    df = pd.read_excel(path)
    if df.empty:
        return {}
    cols = [str(c).strip() for c in df.columns]
    manager_col = None
    score_col = None
    for c in cols:
        if "מנהל" in c or "גוף" in c or "חברה" in c:
            manager_col = c
            break
    for c in cols:
        if "שירות" in c or "ציון" in c or "score" in c.lower():
            score_col = c
            break
    if manager_col is None:
        manager_col = cols[0]
    if score_col is None:
        score_col = cols[1] if len(cols) > 1 else cols[0]

    df2 = df.rename(columns={manager_col: "manager", score_col: "service"})
    df2["manager"] = df2["manager"].astype(str).str.strip()
    df2["service"] = pd.to_numeric(df2["service"], errors="coerce")
    scores = {}
    for _, r in df2.iterrows():
        if pd.isna(r["service"]):
            continue
        scores[str(r["manager"]).strip()] = float(r["service"])
    return scores

def generate_weights(n: int, step: int):
    step = int(step)
    if n == 1:
        yield (100,)
        return
    if n == 2:
        for a in range(0, 101, step):
            yield (a, 100 - a)
        return
    for a in range(0, 101, step):
        for b in range(0, 101 - a, step):
            c = 100 - a - b
            if c < 0:
                continue
            if c % step != 0:
                continue
            yield (a, b, c)

def weighted_metrics(rows: list[dict], weights: list[float]) -> dict:
    keys = ["stocks","foreign","israel","fx","illiquid","sharpe"]
    w = np.array(weights, dtype=float)
    w = w / w.sum()
    out = {}
    for k in keys:
        vals = np.array([r.get(k, np.nan) for r in rows], dtype=float)
        mask = ~np.isnan(vals)
        if mask.sum() == 0:
            out[k] = np.nan
            continue
        ww = w[mask]
        ww = ww / ww.sum()
        out[k] = float(np.sum(vals[mask] * ww))
    return out

def weighted_service(managers: list[str], weights: list[float], service_scores: dict) -> float:
    w = np.array(weights, dtype=float)
    w = w / w.sum()
    vals = np.array([service_scores.get(m, np.nan) for m in managers], dtype=float)
    mask = ~np.isnan(vals)
    if mask.sum() == 0:
        return np.nan
    ww = w[mask]
    ww = ww / ww.sum()
    return float(np.sum(vals[mask] * ww))

def accuracy_score(metrics: dict, targets: dict) -> float:
    dist = 0.0
    for k, t in targets.items():
        if t is None or (isinstance(t, float) and np.isnan(t)):
            continue
        v = metrics.get(k, np.nan)
        if np.isnan(v):
            dist += 9999.0
        else:
            dist += abs(v - t)
    return dist

def primary_score(metrics: dict, objective: str, svc: float) -> float:
    # unified "lower is better"
    if objective == "דיוק":
        return 0.0
    if objective == "שארפ":
        v = metrics.get("sharpe", np.nan)
        return -(v if not np.isnan(v) else -9999.0)
    if objective == "שירות":
        return -(svc if not np.isnan(svc) else -9999.0)
    if objective == "מקסום מט״ח":
        v = metrics.get("fx", np.nan)
        return -(v if not np.isnan(v) else -9999.0)
    return 0.0

def find_best_solutions(
    df_long: pd.DataFrame,
    n_funds: int,
    step: int,
    manager_mode: str,
    objective: str,
    targets: dict,
    hard_min: dict,
    max_illiquid: float,
    service_scores: dict,
    exclude_managers: set[str] | None = None,
    top_k: int = 3000,
):
    exclude_managers = exclude_managers or set()
    items = df_long.copy()
    items = items[~items["manager"].isin(exclude_managers)].reset_index(drop=True)
    records = items.to_dict("records")

    cands = []
    eps = 1e-9

    for combo_idx in itertools.combinations(range(len(records)), n_funds):
        rows = [records[i] for i in combo_idx]
        managers = [r["manager"] for r in rows]

        if manager_mode == "פיזור בין מנהלים":
            if len(set(managers)) != n_funds:
                continue
        else:
            if len(set(managers)) != 1:
                continue

        for wts in generate_weights(n_funds, step):
            metrics = weighted_metrics(rows, list(wts))

            # Hard constraint: illiquid max
            ill = metrics.get("illiquid", np.nan)
            if pd.notna(max_illiquid) and pd.notna(ill) and ill > max_illiquid + eps:
                continue

            # Hard minimum constraints (e.g. FX >= target)
            ok = True
            for k, is_hard in hard_min.items():
                if not is_hard:
                    continue
                t = targets.get(k, np.nan)
                v = metrics.get(k, np.nan)
                if pd.notna(t) and pd.notna(v):
                    if v + eps < t:
                        ok = False
                        break
                else:
                    ok = False
                    break
            if not ok:
                continue

            svc = weighted_service(managers, list(wts), service_scores)
            acc = accuracy_score(metrics, targets)

            if objective == "דיוק":
                prim = acc
            else:
                prim = primary_score(metrics, objective, svc)

            cands.append({
                "rows": rows,
                "weights": list(wts),
                "metrics": metrics,
                "service": svc,
                "accuracy": acc,
                "primary": prim,
            })

    cands.sort(key=lambda x: x["primary"])
    return cands[:top_k]

def pick_distinct_manager_solution(cands, used_managers: set[str]):
    for c in cands:
        mans = {r["manager"] for r in c["rows"]}
        if mans.isdisjoint(used_managers):
            return c
    return None

def solution_to_row(sol, label: str, obj_name: str):
    rows = sol["rows"]
    wts = sol["weights"]
    m = sol["metrics"]
    managers = [r["manager"] for r in rows]
    funds = [r["fund_name"] for r in rows]
    sheets = [r["sheet"] for r in rows]

    combo_txt = " + ".join([f"{wts[i]}% • {funds[i]} ({sheets[i]})" for i in range(len(rows))])
    managers_txt = " + ".join([f"{wts[i]}% • {managers[i]}" for i in range(len(rows))])

    return {
        "חלופה": label,
        "אובייקטיב": obj_name,
        "יתרון": "",
        "שילוב (מסלול/קופה)": combo_txt,
        "מנהלים": managers_txt,
        "Score סטייה": sol["accuracy"],
        "חו״ל": m.get("foreign", np.nan),
        "ישראל": m.get("israel", np.nan),
        "מניות": m.get("stocks", np.nan),
        "מט״ח": m.get("fx", np.nan),
        "לא סחיר": m.get("illiquid", np.nan),
        "שארפ": m.get("sharpe", np.nan),
        "ציון שירות": sol.get("service", np.nan),
    }

def style_results_df(df: pd.DataFrame, illiquid_limit: float, dev_warn: float):
    def _style(row):
        bg = ""
        if row["חלופה"] == "חלופה 1":
            bg = "background-color: rgba(46, 204, 113, 0.20);"
        if pd.notna(illiquid_limit) and pd.notna(row["לא סחיר"]) and row["לא סחיר"] > illiquid_limit:
            bg = "background-color: rgba(231, 76, 60, 0.25);"
        if pd.notna(dev_warn) and pd.notna(row["Score סטייה"]) and row["Score סטייה"] > dev_warn:
            if "231, 76, 60" not in bg:
                bg = "background-color: rgba(243, 156, 18, 0.22);"
        return [bg] * len(row)
    return df.style.apply(_style, axis=1).format({
        "Score סטייה": "{:.2f}",
        "חו״ל": "{:.2f}",
        "ישראל": "{:.2f}",
        "מניות": "{:.2f}",
        "מט״ח": "{:.2f}",
        "לא סחיר": "{:.2f}",
        "שארפ": "{:.3f}",
        "ציון שירות": "{:.1f}",
    })

# -----------------------------
# Password gate
# -----------------------------
def check_password():
    secret = None
    try:
        secret = st.secrets.get("APP_PASSWORD", None)
    except Exception:
        secret = None
    if not secret:
        secret = "1234"  # change via secrets

    if "auth_ok" not in st.session_state:
        st.session_state.auth_ok = False
    if st.session_state.auth_ok:
        return True

    st.markdown("## 🔒 כניסה")
    st.caption("הזן סיסמה כדי להמשיך.")
    pw = st.text_input("סיסמה", type="password", placeholder="••••")
    if st.button("כניסה", use_container_width=True):
        if hmac.compare_digest(str(pw), str(secret)):
            st.session_state.auth_ok = True
            st.rerun()
        else:
            st.error("סיסמה שגויה.")
    st.stop()

check_password()

# -----------------------------
# Load embedded data
# -----------------------------
if not HOLDINGS_XLSX.exists():
    st.error("קובץ המסלולים לא נמצא בתיקיית data.")
    st.stop()

@st.cache_data(show_spinner=False)
def _cached_load_holdings(path_str: str):
    return load_holdings_excel(Path(path_str))

@st.cache_data(show_spinner=False)
def _cached_load_service(path_str: str):
    return load_service_scores_excel(Path(path_str))

with st.spinner("טוען נתונים..."):
    df_long, sheet_names = _cached_load_holdings(str(HOLDINGS_XLSX))

if df_long.empty:
    st.error("לא הצלחתי לזהות טבלאות תקינות בקובץ. ודא שבכל גיליון יש עמודת 'פרמטר' ושורות חשיפה מרכזיות.")
    st.stop()

service_scores = {}
if SERVICE_XLSX.exists():
    service_scores = _cached_load_service(str(SERVICE_XLSX))
for m in df_long["manager"].unique().tolist():
    service_scores.setdefault(m, 50.0)

# -----------------------------
# Header
# -----------------------------
st.markdown("# 📊 מנוע תמהילי קרנות השתלמות")
st.markdown(
    "האפליקציה משתמשת בנתונים **מובנים** (Single Source of Truth) מתוך קבצי אקסל פנימיים, "
    "ומחפשת שילובי קופות (1–3) עם משקלים שמסכמים ל-100% לפי היעדים שלך. "
    "כולל **כלל ישראל**: *ישראל = 100 − חו״ל*."
)

st.markdown(
    f"**סך מסלולי ההשקעה שזוהו בקובץ:** `{len(sheet_names)}`  &nbsp;|&nbsp; "
    f"**סך קופות (מנהל×מסלול) שזוהו:** `{len(df_long)}`"
)

# -----------------------------
# Tabs
# -----------------------------
tab1, tab2, tab3 = st.tabs(["הגדרות יעד", "תוצאות (3 חלופות)", "פירוט חישוב / שקיפות"])

with tab1:
    st.subheader("הגדרות יעד")
    st.caption("הגדר גם את התצורה (מספר קופות, ערבוב מנהלים, צעד משקלים) **כאן במסך הראשי**. ניתן לסמן יעד 'קשיח' (מינימום). לא-סחיר הוא תמיד מגבלה קשיחה (מקסימום).")

    cA, cB, cC, cD = st.columns([1.2, 1.2, 1.2, 1.2])
    with cA:
        n_funds = st.radio("כמה קופות לשלב?", [1, 2, 3], horizontal=True, index=1)
    with cB:
        manager_mode = st.radio("בחירת מנהלים", ["פיזור בין מנהלים", "אותו מנהל בלבד"], index=0)
    with cC:
        step = st.select_slider("צעד משקלים (%)", options=[1,2,5,10], value=5)
    with cD:
        objective = st.selectbox("דירוג ראשי", ["דיוק", "שארפ", "שירות", "מקסום מט״ח"], index=0)

    st.markdown("---")
    st.markdown("#### הגדרת היעדים")

    cols = st.columns([1.2, 1.8, 1.0, 1.2])
    cols[0].markdown("**משתנה**")
    cols[1].markdown("**ערך יעד**")
    cols[2].markdown("**כלול ביעד**")
    cols[3].markdown("**יעד קשיח (מינימום)**")

    def row_controls(label, minv, maxv, default, step=0.5, hard_allowed=True, key_prefix=""):
        c = st.columns([1.2, 1.8, 1.0, 1.2])
        c[0].markdown(f"{label}")
        val = c[1].slider(f"{label}__", minv, maxv, default, step=step, label_visibility="collapsed", key=f"{key_prefix}_val")
        inc = c[2].checkbox("כלול", value=True, label_visibility="collapsed", key=f"{key_prefix}_inc")
        hard = c[3].checkbox("קשיח", value=False, disabled=not hard_allowed, label_visibility="collapsed", key=f"{key_prefix}_hard")
        return val, inc, hard

    foreign_v, foreign_inc, foreign_hard = row_controls("יעד חו״ל (%)", 0.0, 120.0, 60.0, key_prefix="foreign")
    stocks_v, stocks_inc, stocks_hard = row_controls("יעד מניות (%)", 0.0, 120.0, 40.0, key_prefix="stocks")
    fx_v, fx_inc, fx_hard = row_controls("יעד מט״ח (%)", 0.0, 150.0, 25.0, key_prefix="fx")

    c = st.columns([1.2, 1.8, 1.0, 1.2])
    c[0].markdown("מקסימום לא סחיר (%)")
    max_illiquid = c[1].slider("מקסימום לא סחיר__", 0.0, 60.0, 20.0, step=0.5, label_visibility="collapsed", key="ill_max")
    ill_inc = c[2].checkbox("כלול", value=False, label_visibility="collapsed", key="ill_inc")
    c[3].markdown("<span style='opacity:0.7'>מגבלה קשיחה</span>", unsafe_allow_html=True)

    ill_target = np.nan
    if ill_inc:
        ill_target = st.slider("יעד לא סחיר (לדיוק בלבד)", 0.0, 60.0, 20.0, step=0.5)

    st.markdown("---")
    dev_warn = st.slider("סף אזהרת סטייה (Score)", 0.0, 30.0, 6.0, 0.5)

    targets = {
        "foreign": foreign_v if foreign_inc else np.nan,
        "stocks": stocks_v if stocks_inc else np.nan,
        "fx": fx_v if fx_inc else np.nan,
        "illiquid": ill_target if not (isinstance(ill_target, float) and np.isnan(ill_target)) else np.nan,
    }
    hard_min = {
        "foreign": bool(foreign_hard and foreign_inc),
        "stocks": bool(stocks_hard and stocks_inc),
        "fx": bool(fx_hard and fx_inc),
        "illiquid": False,
    }

    run = st.button("חשב / חפש חלופות", type="primary", use_container_width=True)

with tab2:
    st.subheader("תוצאות (3 חלופות)")
    st.caption("חלופה 1 לפי הדירוג הראשי, חלופה 2 לפי שארפ, חלופה 3 לפי שירות. נעשה מאמץ לשמור על מנהלים שונים בין חלופות.")
    placeholder_results = st.empty()

with tab3:
    st.subheader("פירוט חישוב / שקיפות")
    with st.expander("פתח פירוט", expanded=False):
        st.write("**Single Source of Truth:** הנתונים נקראים רק מקבצים פנימיים בתיקיית data.")
        st.write("**כלל ישראל:** ישראל = 100 − חו״ל.")
        st.write("**חיפוש:** כוח-גס יציב/יסודי על כל צירופי הקופות ועל כל חלוקות המשקל לפי צעד המשקלים.")
        st.write("**מגבלות:**")
        st.write("- לא-סחיר משוקלל ≤ המקסימום שהוגדר (קשיח).")
        st.write("- יעדים שסומנו 'קשיח' הם **מינימום** (לדוגמה: מט״ח 100% בלי אפשרות לפחות).")
        st.write("**Score סטייה:** סכום סטיות מוחלטות (L1) רק על המשתנים שסימנת 'כלול ביעד'.")

# -----------------------------
# Compute
# -----------------------------
if "last_results" not in st.session_state:
    st.session_state.last_results = None

def compute_all():
    with st.spinner("מחשב שילובים... (יציב/יסודי)"):
        primary_cands = find_best_solutions(
            df_long=df_long,
            n_funds=int(n_funds),
            step=int(step),
            manager_mode=manager_mode,
            objective=objective,
            targets=targets,
            hard_min=hard_min,
            max_illiquid=float(max_illiquid),
            service_scores=service_scores,
            exclude_managers=set(),
            top_k=3000,
        )
        if not primary_cands:
            return None, "לא נמצאו שילובים שעומדים במגבלות. נסה להקל יעד קשיח/מגבלת לא-סחיר או להגדיל צעד משקלים."

        sol1 = primary_cands[0]
        used = {r["manager"] for r in sol1["rows"]}

        sharpe_cands = find_best_solutions(
            df_long=df_long,
            n_funds=int(n_funds),
            step=int(step),
            manager_mode=manager_mode,
            objective="שארפ",
            targets=targets,
            hard_min=hard_min,
            max_illiquid=float(max_illiquid),
            service_scores=service_scores,
            exclude_managers=set(),
            top_k=3000,
        )
        sol2 = pick_distinct_manager_solution(sharpe_cands, used)
        note2 = None
        if sol2 is None and sharpe_cands:
            sol2 = sharpe_cands[0]
            note2 = "חלופה 2: לא נמצא פתרון ללא חפיפת מנהלים מול חלופה 1."

        used2 = used | ({r["manager"] for r in sol2["rows"]} if sol2 else set())

        service_cands = find_best_solutions(
            df_long=df_long,
            n_funds=int(n_funds),
            step=int(step),
            manager_mode=manager_mode,
            objective="שירות",
            targets=targets,
            hard_min=hard_min,
            max_illiquid=float(max_illiquid),
            service_scores=service_scores,
            exclude_managers=set(),
            top_k=3000,
        )
        sol3 = pick_distinct_manager_solution(service_cands, used2)
        note3 = None
        if sol3 is None and service_cands:
            sol3 = service_cands[0]
            note3 = "חלופה 3: לא נמצא פתרון ללא חפיפת מנהלים מול חלופות 1–2."

        rows = []
        if sol1:
            r1 = solution_to_row(sol1, "חלופה 1", objective)
            r1["יתרון"] = f"הכי מדויק ליעד, סטייה כוללת {r1['Score סטייה']:.2f}"
            rows.append(r1)
        if sol2:
            r2 = solution_to_row(sol2, "חלופה 2", "שארפ")
            r2["יתרון"] = f"שארפ גבוה יותר תוך סטייה {r2['Score סטייה']:.2f}"
            rows.append(r2)
        if sol3:
            r3 = solution_to_row(sol3, "חלופה 3", "שירות")
            r3["יתרון"] = f"ציון שירות משוקלל הגבוה ביותר תוך סטייה {r3['Score סטייה']:.2f}"
            rows.append(r3)

        out = pd.DataFrame(rows)
        notes = " | ".join([n for n in [note2, note3] if n]) if (note2 or note3) else None
        return out, notes

def render_results(df_out, notes):
    # KPI blocks
    for _, row in df_out.iterrows():
        st.markdown(f"### {row['חלופה']} • ({row['אובייקטיב']})")
        st.markdown(
            f"""
<div class="kpi-wrap">
  <div class="kpi">
    <h4>Score (סטייה מהיעד)</h4>
    <div class="big">{row['Score סטייה']:.2f}</div>
    <div class="sub">ככל שנמוך יותר — מדויק יותר</div>
  </div>
  <div class="kpi">
    <h4>חשיפות (חו״ל / מניות / מט״ח / לא סחיר)</h4>
    <div class="big">{row['חו״ל']:.1f}% • {row['מניות']:.1f}% • {row['מט״ח']:.1f}% • {row['לא סחיר']:.1f}%</div>
    <div class="sub">ישראל מחושבת: {row['ישראל']:.1f}%</div>
  </div>
  <div class="kpi">
    <h4>שארפ משוקלל</h4>
    <div class="big">{row['שארפ']:.3f}</div>
    <div class="sub">ציון שירות: {row['ציון שירות']:.1f}</div>
  </div>
</div>
""",
            unsafe_allow_html=True
        )
        st.markdown("")

    st.markdown("#### טבלה מלאה")
    col_order = [
        "חלופה","אובייקטיב","יתרון",
        "שילוב (מסלול/קופה)","מנהלים",
        "Score סטייה","חו״ל","ישראל","מניות","מט״ח","לא סחיר","שארפ","ציון שירות"
    ]
    df2 = df_out[col_order].copy()
    styled = style_results_df(df2, illiquid_limit=float(max_illiquid), dev_warn=float(dev_warn))
    st.dataframe(
        styled,
        use_container_width=True,
        height=240,
        column_config={
            "שילוב (מסלול/קופה)": st.column_config.TextColumn(width="large"),
            "מנהלים": st.column_config.TextColumn(width="large"),
            "יתרון": st.column_config.TextColumn(width="large"),
        },
    )
    if notes:
        st.info(notes)

if run:
    df_out, note = compute_all()
    if df_out is None:
        st.session_state.last_results = None
        with tab2:
            placeholder_results.error(note or "לא נמצאו תוצאות.")
    else:
        st.session_state.last_results = (df_out, note)
        st.rerun()

with tab2:
    if st.session_state.last_results is None:
        placeholder_results.info("כשתלחץ על 'חשב / חפש חלופות' בטאב הראשון — יופיעו כאן 3 חלופות.")
    else:
        df_out, note = st.session_state.last_results
        render_results(df_out, note)
