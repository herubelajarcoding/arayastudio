DB_VERSION = 'V3b_v4'

import sqlite3
import os
import re

try:
    import psycopg
except ImportError:
    psycopg = None

try:
    import psycopg2
except ImportError:
    psycopg2 = None

try:
    from psycopg_pool import ConnectionPool
except ImportError:
    ConnectionPool = None
from pathlib import Path
from datetime import date, datetime, time, timedelta
from zoneinfo import ZoneInfo
import calendar
import html
from io import BytesIO
import re

import pandas as pd
import streamlit as st

try:
    from reportlab.lib import colors
    from reportlab.lib.enums import TA_CENTER
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import mm
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    from reportlab.platypus import (
        SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, KeepTogether
    )
    REPORTLAB_AVAILABLE = True
except ImportError:
    REPORTLAB_AVAILABLE = False


# ============================================================
# CONFIG
# ============================================================

APP_DIR = Path(__file__).resolve().parent
DATA_DIR = APP_DIR / "data"
DB_PATH = DATA_DIR / "studio_control.db"
SEED_XLSM = DATA_DIR / "Studio_Control_Board_FINAL.xlsm"

ARAYA_LOGO_DATA_URI = "data:image/png;base64,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"

st.set_page_config(
    page_title="Studio Control Board",
    page_icon="📐",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Persistent database mode:
# - DATABASE_URL configured -> PostgreSQL/Supabase, persistent across redeploys.
# - no DATABASE_URL -> local SQLite fallback (development only).


# ============================================================
# STYLING
# ============================================================

st.markdown(
    """
    <style>
    .block-container {padding-top: 1.2rem; padding-bottom: 2rem;}
    /* V2e: freeze the complete Weekly Dashboard header.
       The anchor is inside the same Streamlit vertical block as:
       title + subtitle + filters + KPI cards. */
    div[data-testid="stVerticalBlock"]:has(> div > #weekly-dashboard-anchor) {
        position: -webkit-sticky !important;
        position: sticky !important;
        top: 0 !important;
        z-index: 9999 !important;
        background: rgba(255,255,255,.985) !important;
        padding: .15rem 0 .85rem !important;
        margin-bottom: .35rem !important;
        border-bottom: 1px solid #D0D5DD !important;
        box-shadow: 0 6px 14px rgba(16,24,40,.07) !important;
        backdrop-filter: blur(10px) !important;
    }
    #weekly-dashboard-anchor {
        display:block !important;
        height:0 !important;
        min-height:0 !important;
        overflow:hidden !important;
        margin:0 !important;
        padding:0 !important;
    }
    .activity-filter-note {font-size:.72rem;color:#667085;margin-top:-.35rem;margin-bottom:.25rem;}
    /* Keep the frozen block visually attached to the viewport top. */
    div[data-testid="stVerticalBlock"]:has(> div > #weekly-dashboard-anchor)
        > div:first-child { margin-top:0 !important; }

    /* Filter bar — intentionally compact and visually consistent with the mockup. */
    div[data-testid="stHorizontalBlock"] .stSelectbox > label,
    div[data-testid="stHorizontalBlock"] .stPopover > button + div {font-weight:600;}
    div[data-testid="stHorizontalBlock"] .stSelectbox > div > div,
    div[data-testid="stHorizontalBlock"] .stPopover > button {
        min-height:42px !important;
        border-radius:9px !important;
        border:1px solid #E2E8F0 !important;
        background:#F5F8FC !important;
        box-shadow:none !important;
    }
    div[data-testid="stHorizontalBlock"] .stSelectbox > div > div:hover,
    div[data-testid="stHorizontalBlock"] .stPopover > button:hover {
        border-color:#B8C7DA !important;
        background:#F0F5FA !important;
    }
    div[data-testid="stHorizontalBlock"] .stSelectbox [data-baseweb="select"] > div {
        min-height:42px !important;
        border:0 !important;
        background:transparent !important;
    }
    div[data-testid="stHorizontalBlock"] .stPopover > button {
        width:100% !important;
        min-height:42px !important;
        justify-content:flex-start !important;
        color:#172B4D !important;
        font-weight:600 !important;
        padding:0 .8rem !important;
        border-radius:9px !important;
        border:1px solid #E2E8F0 !important;
        background:#F5F8FC !important;
        box-shadow:none !important;
    }
    div[data-testid="stHorizontalBlock"] .stPopover > button p {
        font-size:.83rem !important;
        font-weight:600 !important;
        color:#172B4D !important;
    }
    .filter-label {
        font-size:.74rem; font-weight:700; color:#475467;
        margin:0 0 .28rem .05rem; letter-spacing:.02em;
        height:1.18rem !important;
        line-height:1.18rem !important;
        display:flex !important;
        align-items:flex-start !important;
    }
    /* Keep native selectboxes and checklist popovers on exactly the same
       vertical baseline. */
    div[data-testid="stHorizontalBlock"] .stSelectbox > label {
        min-height:1.18rem !important;
        margin-bottom:.28rem !important;
        line-height:1.18rem !important;
    }
    div[data-testid="stHorizontalBlock"] .stSelectbox > div > div,
    div[data-testid="stHorizontalBlock"] .stPopover > button {
        height:42px !important;
        min-height:42px !important;
        box-sizing:border-box !important;
    }
    .filter-value {font-size:.83rem;color:#172B4D;font-weight:600;}
    .app-title {display:block; font-size:2rem; line-height:1.3; font-weight:750; margin:0 0 0.1rem 0; padding-top:1.35rem; padding-bottom:.05rem; overflow:visible !important; height:auto !important; min-height:2.6rem;}
    /* ========================================================
       V3A STATIC SIDEBAR TREE
       ======================================================== */
    section[data-testid="stSidebar"] > div {
        padding-top:1.15rem;
    }
    .v3-brand {
        display:flex;
        align-items:center;
        gap:.72rem;
        padding:.35rem .15rem 1.05rem;
        margin-bottom:.25rem;
        border-bottom:1px solid #E4E7EC;
    }
    .v3-brand-mark {
        width:42px;
        height:42px;
        min-width:42px;
        display:flex;
        align-items:center;
        justify-content:center;
        background:transparent;
        overflow:hidden;
    }
    .v3-brand-mark img {
        display:block;
        width:100%;
        height:100%;
        object-fit:contain;
    }

    /* V8b — minimalist architectural landing page */
    .araya-home {
        min-height:78vh;
        display:flex;
        align-items:center;
        justify-content:center;
        padding:5vh 1.5rem;
    }
    .araya-home-inner {
        width:min(760px,92vw);
        text-align:center;
    }
    .araya-home-logo {
        width:clamp(110px,15vw,180px);
        height:auto;
        display:block;
        margin:0 auto 2.2rem;
    }
    .araya-home-rule {
        width:54px;
        height:1px;
        background:#111111;
        margin:0 auto 1.7rem;
    }
    .araya-home-title {
        margin:0;
        color:#111111;
        font-size:clamp(2.2rem,5vw,4.8rem);
        line-height:1;
        font-weight:500;
        letter-spacing:.18em;
        padding-left:.18em;
        white-space:nowrap;
    }
    @media (max-width:640px) {
        .araya-home {min-height:68vh;padding-top:2rem;}
        .araya-home-title {
            font-size:clamp(1.75rem,9vw,3rem);
            letter-spacing:.12em;
            padding-left:.12em;
        }
    }
    .v3-brand-name {
        color:#172B4D;
        font-size:1.08rem;
        font-weight:850;
        letter-spacing:.02em;
    }
    .v3-brand-sub {
        color:#667085;
        font-size:.70rem;
        margin-top:.12rem;
    }
    .v3-nav-label {
        color:#98A2B3;
        font-size:.68rem;
        font-weight:800;
        letter-spacing:.08em;
        margin:.8rem .15rem .55rem;
    }

    /* Module = section heading, not another clickable menu. */
    .v3-module-heading {
        color:#172B4D;
        font-size:1.02rem;
        font-weight:800;
        line-height:1.25;
        margin:.72rem .15rem .22rem;
        padding:.18rem .15rem;
    }

    /* Submodules = the only clickable navigation items. */
    section[data-testid="stSidebar"] .stButton > button {
        min-height:2.15rem;
        border-radius:8px;
        font-size:.84rem;
        font-weight:550;
        text-align:left;
        justify-content:flex-start;
        padding:.35rem .55rem;
        margin:.02rem 0;
    }
    section[data-testid="stSidebar"] .stButton > button:hover {
        background:#F2F4F7;
    }
    .app-subtitle {color:#667085; margin-bottom:1rem;}
    .week-title {
        font-size: 1.12rem; font-weight: 800; padding: 0.7rem 0.9rem;
        border-radius: 8px; background: #EEF2F6; margin-top: 0.75rem;
        text-align: center; color:#172B4D; letter-spacing:.01em;
    }
    .schedule-grid {
        display: grid;
        grid-template-columns: minmax(105px, .8fr) repeat(7, minmax(135px, 1fr));
        align-items: stretch;
        gap: 0;
        width: 100%;
        overflow-x: auto;
        border-left: 1px solid #EAECF0;
        border-top: 1px solid #EAECF0;
        border-radius: 0 0 8px 8px;
    }
    .schedule-head {
        min-height: 58px;
        padding: 0.45rem 0.3rem;
        font-weight: 700;
        text-align: center;
        background: #F8FAFC;
        border-right: 1px solid #EAECF0;
        border-bottom: 1px solid #D0D5DD;
    }
    .activity-head {display:flex;align-items:center;justify-content:center;}
    .schedule-head .dow {font-size: 0.72rem; color:#667085; text-transform:uppercase;}
    .schedule-head .day {font-size: 1rem; color:#101828;}
    .schedule-head.non-working-day {
        background:#ECFDF3 !important;
        border-color:#ABEFC6 !important;
    }
    .schedule-head.non-working-day .dow,
    .schedule-head.non-working-day .day {
        color:#067647 !important;
        font-weight:800 !important;
    }
    .schedule-head.non-working-day .date-detail-link:hover {
        background:#D1FADF !important;
    }
    .lane-label {
        min-height: 150px;
        padding: 0.9rem 0.45rem;
        font-weight: 850;
        font-size: 1.02rem;
        letter-spacing: .035em;
        border-right: 1px solid rgba(16,24,40,.10);
        color:#172B4D;
        display:flex;
        flex-direction:column;
        align-items:center;
        justify-content:center;
        gap:0.55rem;
    }
    .lane-icon {
        display:flex;
        width:38px;
        height:38px;
        border-radius:11px;
        align-items:center;
        justify-content:center;
        font-size:1.22rem;
        font-weight:900;
        background:rgba(255,255,255,.72);
        box-shadow:0 2px 7px rgba(16,24,40,.08);
    }
    .work-lane .lane-icon {color:#1479E9;}
    .meeting-lane .lane-icon {color:#D99A00;}
    .other-lane .lane-icon {color:#8B5CF6;}
    .lane-label.work-lane {background:#E8F3FF;}
    .lane-label.meeting-lane {background:#FFF7D6;}
    .lane-label.other-lane {background:#F2E9FF;}
    .cell.work-cell {background:#F4FAFF;}
    .cell.meeting-cell {background:#FFFBEA;}
    .meeting-cell .project {color:#1F2937;}
    .meeting-cell .task {color:#111111; font-weight:400;}
    .meeting-cell .meta {color:#475467; font-weight:500;}
    .work-cell .project, .other-cell .other-item {color:#172B4D;}
    .cell.other-cell {background:#FAF5FF;}
    .cell {
        min-height: 130px;
        padding: 0.55rem;
        border-right: 1px solid #EAECF0;
        border-bottom: 1px solid #EAECF0;
        background: white;
    }
    .project {
        font-weight: 750;
        color:#101828;
        margin: 0.10rem 0 0.28rem;
    }
    /* Project groups are separated; project header and its tasks stay tight. */
    .project.project-start {
        margin-top: 0.95rem;
        padding-top: 0.62rem;
        border-top: 1px solid rgba(16,24,40,.13);
    }
    .task {font-size: 0.88rem; line-height: 1.42; margin: 0.28rem 0; color:#111111; font-weight:400;}
    .task.submission {font-weight: 750; color:#C62828; background:#FFE7E7; border-radius:6px; padding:2px 5px;}
    .sign {color:#D92D20; font-weight: 950; margin-left: 0.2rem; letter-spacing:-.08em;}
    .meta {
        margin-left: 1.05rem;
        font-size: 0.73rem;
        color:#667085;
        line-height: 1.3;
    }
    .other-item {
        font-size: 0.88rem;
        line-height: 1.42;
        font-weight: 700;
        color:#172B4D;
        padding: 0.26rem 0;
    }
    .other-item .meta {
        margin-left: 1.05rem;
        margin-top: 0.12rem;
        font-size: 0.78rem;
        line-height: 1.35;
        font-weight: 600;
        color:#475467;
    }
    .empty {color:#98A2B3; font-size:.72rem;}
        /* V2n: visual separator between project groups in the dashboard. */
    .project {
        font-weight:700;
        color:#172B4D;
        margin-top:.28rem;
        margin-bottom:.16rem;
        padding-top:.38rem;
        border-top:1px solid #D0D5DD;
    }
    .project:first-child {
        border-top:none;
        padding-top:0;
        margin-top:0;
    }
.more-items {font-size:.75rem; font-weight:700; color:#667085; margin:.35rem 0 .15rem 1.05rem;}
    .detail-date {font-size:1rem; font-weight:700; color:#344054; margin:-.25rem 0 1rem;}
        .detail-submission {
        background:#FEE4E2 !important;
        color:#B42318 !important;
        border-radius:8px !important;
        padding:.32rem .45rem !important;
        margin:.38rem 0 !important;
        font-weight:600 !important;
    }
    .detail-submission .detail-pic {
        color:#B42318 !important;
    }
    .detail-submit-badge {
        color:#B42318 !important;
        font-weight:800 !important;
        margin-left:.2rem !important;
    }
.detail-card-grid {display:grid; grid-template-columns:repeat(3,minmax(0,1fr)); gap:1.15rem; align-items:start;}
    .detail-card-column {min-width:0;}
    .detail-card {box-sizing:border-box; width:100%; border:1px solid #D0D5DD; border-radius:14px; padding:1rem; min-height:360px; box-shadow:0 3px 10px rgba(16,24,40,.05); overflow:hidden;}

    .detail-work {background:#F3F8FF; border-color:#B9D6F8;}
    .detail-meeting {background:#FFF9E8; border-color:#F3D98A;}
    .detail-other {background:#F7F0FF; border-color:#D8C0F5;}
    .detail-card-head {font-size:1.05rem; font-weight:800; letter-spacing:.02em; margin-bottom:.9rem; color:#172B4D;}
    .detail-card-icon {display:inline-flex; align-items:center; justify-content:center; margin-right:.35rem; font-weight:900;}
    .detail-project {font-weight:800; color:#172B4D; margin:.75rem 0 .45rem; padding-bottom:.3rem; border-bottom:1px solid rgba(16,24,40,.10);}
    .detail-task {font-size:.86rem; line-height:1.45; color:#111827; margin:.35rem 0;}
    .detail-bullet {font-weight:900;}
    .detail-pic {color:#667085; font-weight:600;}
    .detail-submit {color:#DC2626; font-weight:900; margin-left:.2rem;}
    .priority-tag {display:inline-block; font-size:.65rem; font-weight:800; letter-spacing:.03em; border-radius:999px; padding:.12rem .38rem; margin-left:.35rem; vertical-align:middle;}
    .priority-high {color:#B42318; background:#FEE4E2;}
    .priority-medium {color:#9A6700; background:#FFF0C2;}
    .priority-low {color:#175CD3; background:#D1E9FF;} .priority-none {color:#667085; background:#F2F4F7;}
    .detail-meta {font-size:.73rem; line-height:1.4; color:#667085; margin-left:1.05rem;}
    .detail-meeting-item, .detail-other-item {margin:.45rem 0 .8rem; padding-bottom:.6rem; border-bottom:1px dashed rgba(16,24,40,.12);}
    .detail-empty {color:#98A2B3; font-size:.78rem; padding:1.5rem .25rem;}

    .date-detail-link {display:block; color:inherit; text-decoration:none; border-radius:10px; padding:.15rem .1rem;}
    .date-detail-link:hover {background:#EEF4FF; text-decoration:none; cursor:pointer;}
    .date-detail-link .dow, .date-detail-link .day {pointer-events:none;}
    /* KPI cards — visual enhancement only; filter/dashboard layout remains unchanged. */
    .kpi-row {
        display:grid;
        grid-template-columns:repeat(4,minmax(0,1fr));
        gap:16px;
        width:100%;
        margin:0.35rem 0 0.25rem;
    }
    .kpi {
        position:relative;
        min-height:118px;
        padding:18px 20px;
        border:1px solid rgba(16,24,40,.07);
        border-radius:15px;
        overflow:hidden;
        box-shadow:0 3px 12px rgba(16,24,40,.055);
        display:flex;
        align-items:center;
        gap:17px;
    }
    .kpi::before {
        content:"";
        position:absolute;
        width:112px;
        height:112px;
        right:-32px;
        top:-43px;
        border-radius:50%;
        background:rgba(255,255,255,.38);
    }
    .kpi::after {
        content:"";
        position:absolute;
        width:155px;
        height:155px;
        right:-48px;
        bottom:-92px;
        border-radius:50%;
        background:rgba(255,255,255,.30);
    }
    .kpi.work {background:linear-gradient(135deg,#EAF4FF 0%,#DCEEFF 100%);}
    .kpi.meeting {background:linear-gradient(135deg,#FFF9DF 0%,#FFF1B9 100%);}
    .kpi.submission {background:linear-gradient(135deg,#FFF0F1 0%,#FFE0E3 100%);}
    .kpi.other {background:linear-gradient(135deg,#F4ECFF 0%,#E9DDFF 100%);}
    .kpi-icon {
        position:relative;
        z-index:2;
        width:52px;
        height:52px;
        min-width:52px;
        border-radius:14px;
        display:flex;
        align-items:center;
        justify-content:center;
    }
    .kpi.work .kpi-icon {background:#1479E9;color:#fff;}
    .kpi.meeting .kpi-icon {background:#F5B400;color:#fff;}
    .kpi.submission .kpi-icon {background:#EF4444;color:#fff;}
    .kpi.other .kpi-icon {background:#8B5CF6;color:#fff;}
    .kpi-content {position:relative;z-index:2;}
    .kpi-label {
        font-size:.82rem;
        line-height:1.2;
        font-weight:800;
        letter-spacing:.025em;
        text-transform:uppercase;
        margin-bottom:6px;
    }
    .kpi.work .kpi-label {color:#1479E9;}
    .kpi.meeting .kpi-label {color:#C88A00;}
    .kpi.submission .kpi-label {color:#E33A40;}
    .kpi.other .kpi-label {color:#7040D8;}
    .kpi-value {
        font-size:2.42rem;
        line-height:1;
        font-weight:850;
        color:#102A56;
    }
    @media (max-width: 900px) {
        .kpi-row {grid-template-columns:repeat(2,minmax(0,1fr));}
    }
    @media (max-width: 560px) {
        .kpi-row {grid-template-columns:1fr;}
    }
    /* Setup tabs: horizontal scrolling prevents master names from being clipped. */

    /* V3 UI FIX — responsive master/input navigation */
    .stTabs [data-baseweb="tab-list"] {
        display: flex !important;
        flex-wrap: nowrap !important;
        overflow-x: auto !important;
        overflow-y: hidden !important;
        scrollbar-width: thin !important;
        gap: .25rem !important;
        max-width: 100% !important;
    }
    .stTabs [data-baseweb="tab-list"] > div {
        display: flex !important;
        flex-wrap: nowrap !important;
        min-width: max-content !important;
    }
    .stTabs [data-baseweb="tab"] {
        flex: 0 0 auto !important;
        white-space: nowrap !important;
        min-width: max-content !important;
        padding-left: .65rem !important;
        padding-right: .65rem !important;
    }

    /* Keep wide forms/tables inside the viewport instead of clipping. */
    [data-testid="stHorizontalBlock"] {
        max-width: 100% !important;
    }
    .stDataFrame, [data-testid="stDataFrame"] {
        max-width: 100% !important;
    }

    /* Horizontal scrolling for any deliberately wide navigation row. */
    .arayastd-scroll-x {
        width: 100%;
        overflow-x: auto;
        overflow-y: hidden;
    }

    /* Prevent long labels from forcing cards/forms wider than viewport. */
    .stButton button, .stSelectbox, .stTextInput, .stNumberInput,
    .stMultiSelect, .stDateInput {
        max-width: 100% !important;
    }

    .stTabs [data-baseweb="tab-list"] { overflow-x: auto !important; overflow-y: hidden !important; flex-wrap: nowrap !important; gap: .3rem !important; }
    .stTabs [data-baseweb="tab-list"] > div { flex-wrap: nowrap !important; min-width: max-content !important; }
    .stTabs [data-baseweb="tab"] { flex: 0 0 auto !important; white-space: nowrap !important; min-width: max-content !important; }
    .stTabs [data-baseweb="tab-list"] {gap: 1.25rem;}
    </style>
    """,
    unsafe_allow_html=True,
)


# ============================================================
# DATABASE
# ============================================================


def _database_url():
    """Read persistent PostgreSQL URL from Streamlit Secrets or environment."""
    try:
        url = st.secrets.get("DATABASE_URL", "")
    except Exception:
        url = ""
    return clean(url or os.getenv("DATABASE_URL", ""))


def using_postgres():
    return bool(_database_url())


@st.cache_resource(show_spinner=False)
def _postgres_pool():
    """Shared client-side connection pool for PostgreSQL/Supabase.

    Streamlit reruns the script frequently. Reusing established connections
    avoids a new TLS/database handshake for every small query.
    """
    url=_database_url()
    if not url or psycopg is None or ConnectionPool is None:
        return None

    pool=ConnectionPool(
        conninfo=url,
        min_size=1,
        max_size=6,
        timeout=15,
        max_idle=300,
        max_lifetime=1800,
        kwargs={"sslmode":"require"},
        open=True,
    )
    pool.wait(timeout=15)
    return pool


def _translate_sql_for_postgres(sql):
    """Translate the small SQLite syntax subset used by this app to PostgreSQL."""
    s = sql
    s = s.replace("INSERT OR IGNORE INTO", "INSERT INTO")
    s = re.sub(
        r"INTEGER\s+PRIMARY\s+KEY\s+AUTOINCREMENT",
        "SERIAL PRIMARY KEY",
        s,
        flags=re.I,
    )
    s = re.sub(r"\bAUTOINCREMENT\b", "", s, flags=re.I)
    # qmark -> psycopg2 format. The application SQL does not use literal '?'.
    s = s.replace("?", "%s")
    return s


class _PGCursor:
    def __init__(self, cursor):
        self._cursor = cursor

    @property
    def description(self):
        return self._cursor.description

    @property
    def rowcount(self):
        return self._cursor.rowcount

    def execute(self, sql, params=None):
        translated = _translate_sql_for_postgres(sql)
        # SQLite's INSERT OR IGNORE is used only for seed operations.
        if "INSERT OR IGNORE INTO" in sql.upper():
            translated = _translate_sql_for_postgres(
                re.sub(r"INSERT\s+OR\s+IGNORE\s+INTO", "INSERT INTO", sql, flags=re.I)
            )
            translated = translated.rstrip().rstrip(";") + " ON CONFLICT DO NOTHING"
        self._cursor.execute(translated, params or ())
        return self

    def fetchone(self):
        return self._cursor.fetchone()

    def fetchall(self):
        return self._cursor.fetchall()

    def close(self):
        self._cursor.close()


class _PGConnection:
    """DB-API compatibility wrapper with pool-aware close semantics."""
    def __init__(self, conn, pool=None):
        self._conn = conn
        self._pool = pool
        self._released = False

    def cursor(self):
        return _PGCursor(self._conn.cursor())

    def execute(self, sql, params=None):
        cur = self.cursor()
        return cur.execute(sql, params)

    def executescript(self, script):
        cur = self.cursor()
        # Schema script contains no procedural SQL; splitting on semicolons is safe.
        for stmt in script.split(";"):
            stmt = stmt.strip()
            if stmt:
                cur.execute(stmt)
        cur.close()

    def commit(self):
        self._conn.commit()

    def rollback(self):
        self._conn.rollback()

    def close(self):
        if self._released:
            return
        self._released = True

        # A SELECT can leave a transaction open. Reset it before returning the
        # connection to the shared pool so the next request starts cleanly.
        try:
            self._conn.rollback()
        except Exception:
            pass

        if self._pool is not None:
            self._pool.putconn(self._conn)
        else:
            self._conn.close()


def get_conn():
    """Use PostgreSQL when DATABASE_URL is configured; SQLite is local fallback."""
    url = _database_url()
    if url:
        pool=_postgres_pool()
        if pool is not None:
            raw=pool.getconn(timeout=15)
            return _PGConnection(raw,pool=pool)

        # Fallback only when the pool package is unavailable.
        if psycopg is not None:
            raw=psycopg.connect(url,sslmode="require")
            return _PGConnection(raw)
        if psycopg2 is not None:
            raw=psycopg2.connect(url,sslmode="require")
            return _PGConnection(raw)
        raise RuntimeError(
            "DATABASE_URL sudah diisi tetapi PostgreSQL driver belum tersedia. "
            "Pastikan requirements.txt memuat psycopg[binary,pool]."
        )

    DATA_DIR.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(DB_PATH, check_same_thread=False)
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


def table_exists(conn, name):
    if using_postgres():
        return conn.execute(
            """SELECT 1 FROM information_schema.tables
               WHERE table_schema='public' AND table_name=?""",
            (name,),
        ).fetchone() is not None
    return conn.execute(
        "SELECT 1 FROM sqlite_master WHERE type='table' AND name=?",
        (name,),
    ).fetchone() is not None


def table_columns(conn, name):
    """Return {column_name: {'notnull': bool}} for SQLite or PostgreSQL."""
    if using_postgres():
        rows = conn.execute(
            """SELECT column_name, is_nullable
               FROM information_schema.columns
               WHERE table_schema='public' AND table_name=?
               ORDER BY ordinal_position""",
            (name,),
        ).fetchall()
        return {r[0]: {"notnull": str(r[1]).upper()=="NO"} for r in rows}
    rows = conn.execute(f"PRAGMA table_info({name})").fetchall()
    return {r[1]: {"notnull": bool(r[3])} for r in rows}


DB_INTEGRITY_ERRORS = [sqlite3.IntegrityError]
if psycopg is not None:
    DB_INTEGRITY_ERRORS.append(psycopg.IntegrityError)
if psycopg2 is not None:
    DB_INTEGRITY_ERRORS.append(psycopg2.IntegrityError)
DB_INTEGRITY_ERRORS = tuple(DB_INTEGRITY_ERRORS)

def _clear_read_caches():
    """Clear cached database reads after a successful write.

    This keeps UI reads fast during normal navigation while making newly
    saved/edited/deleted data visible immediately.
    """
    for fn_name in (
        "_db_df_cached",
        "_project_name",
        "project_name_map",
        "get_projects",
        "get_staff",
        "get_refs",
        "load_activities",
        "master_values",
        "get_master_table",
    ):
        fn=globals().get(fn_name)
        try:
            if fn is not None and hasattr(fn,"clear"):
                fn.clear()
        except Exception:
            pass


def init_db():
    conn = get_conn()

    conn.executescript(
        """
        CREATE TABLE IF NOT EXISTS staff (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT UNIQUE NOT NULL,
            category TEXT,
            primary_role TEXT,
            intern_start TEXT,
            intern_end TEXT,
            active INTEGER DEFAULT 1
        );

        CREATE TABLE IF NOT EXISTS projects (
            id TEXT PRIMARY KEY,
            name TEXT NOT NULL,
            project_type TEXT,
            start_date TEXT,
            target_finish TEXT,
            duration_months REAL,
            status TEXT,
            lead TEXT,
            project_size TEXT
        );

        CREATE TABLE IF NOT EXISTS work_activity (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            project_id TEXT,
            start_date TEXT,
            end_date TEXT,
            activity_type TEXT,
            task TEXT NOT NULL,
            priority TEXT,
            pic TEXT,
            status TEXT,
            notes TEXT,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY(project_id) REFERENCES projects(id) ON UPDATE CASCADE
        );

        CREATE TABLE IF NOT EXISTS meeting_activity (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            activity_date TEXT NOT NULL,
            start_time TEXT,
            end_time TEXT,
            project_id TEXT,
            meeting_type TEXT NOT NULL,
            attendee_1 TEXT,
            attendee_2 TEXT,
            attendee_3 TEXT,
            attendee_4 TEXT,
            location TEXT,
            agenda_notes TEXT,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY(project_id) REFERENCES projects(id) ON UPDATE CASCADE
        );

        CREATE TABLE IF NOT EXISTS other_activity (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            activity_date TEXT NOT NULL,
            activity TEXT NOT NULL,
            related_staff TEXT,
            notes TEXT,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP
        );

        CREATE TABLE IF NOT EXISTS reference_values (
            category TEXT NOT NULL,
            value TEXT NOT NULL,
            PRIMARY KEY(category, value)
        );

        CREATE TABLE IF NOT EXISTS staff_allocation (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            project_id TEXT NOT NULL,
            phase TEXT NOT NULL,
            staff TEXT NOT NULL,
            role_on_project TEXT,
            notes TEXT,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP,
            UNIQUE(project_id, phase, staff)
        );

        CREATE TABLE IF NOT EXISTS freelance_project_mapping (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            freelancer TEXT NOT NULL,
            project_id TEXT NOT NULL,
            start_date TEXT,
            end_date TEXT,
            mapping_status TEXT,
            notes TEXT,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP
        );
        """
    )
    conn.commit()
    _clear_read_caches()
    conn.close()


def seed_from_workbook():
    """Seed SQLite once from the supplied workbook. The web app thereafter
    reads/writes SQLite; the workbook is the initial master/data blueprint."""
    conn = get_conn()

    if table_exists(conn, "projects") and conn.execute(
        "SELECT COUNT(*) FROM projects"
    ).fetchone()[0] > 0:
        conn.close()
        return

    if not SEED_XLSM.exists():
        conn.close()
        return

    try:
        # data_only=True reads the calculated results stored in the workbook.
        xls = pd.ExcelFile(SEED_XLSM, engine="openpyxl")

        # ---- Setup: staff
        setup = pd.read_excel(SEED_XLSM, sheet_name="Setup", header=None, engine="openpyxl")
        for r in range(2, min(len(setup), 23)):
            name = clean(setup.iloc[r, 0] if setup.shape[1] > 0 else None)
            if not name:
                continue
            category = clean(setup.iloc[r, 1])
            role = clean(setup.iloc[r, 2])
            intern_start = to_iso_date(setup.iloc[r, 3])
            intern_end = to_iso_date(setup.iloc[r, 4])
            conn.execute(
                """INSERT OR IGNORE INTO staff
                   (name, category, primary_role, intern_start, intern_end)
                   VALUES (?, ?, ?, ?, ?)""",
                (name, category, role, intern_start, intern_end),
            )

        # ---- Setup: projects H:O (zero-based 7:14)
        for r in range(2, min(len(setup), 18)):
            pid = clean(setup.iloc[r, 7])
            if not pid:
                continue
            pname = clean(setup.iloc[r, 8])
            ptype = clean(setup.iloc[r, 9])
            pstart = to_iso_date(setup.iloc[r, 10])
            pfinish = to_iso_date(setup.iloc[r, 11])
            duration = numeric_or_none(setup.iloc[r, 12])
            status = clean(setup.iloc[r, 13])
            size = clean(setup.iloc[r, 14])
            conn.execute(
                """INSERT OR IGNORE INTO projects
                   (id,name,project_type,start_date,target_finish,duration_months,status,project_size)
                   VALUES (?,?,?,?,?,?,?,?)""",
                (pid, pname, ptype, pstart, pfinish, duration, status, size),
            )

        # ---- Work Activity
        work = pd.read_excel(
            SEED_XLSM, sheet_name="Work Activity", header=1, engine="openpyxl"
        )
        for _, row in work.iterrows():
            pid = clean(row.get("Project ID"))
            task = clean(row.get("Deliverable / Task"))
            if not task:
                continue
            conn.execute(
                """INSERT INTO work_activity
                   (project_id,start_date,end_date,activity_type,task,priority,pic,status,notes)
                   VALUES (?,?,?,?,?,?,?,?,?)""",
                (
                    pid or None,
                    to_iso_date(row.get("Start Date")),
                    to_iso_date(row.get("End Date")),
                    clean(row.get("Activity Type")),
                    task,
                    clean(row.get("Priority")),
                    clean(row.get("PIC")),
                    clean(row.get("Status")),
                    clean(row.get("Notes")),
                ),
            )

        # ---- Meeting Activity
        meeting = pd.read_excel(
            SEED_XLSM, sheet_name="Meeting Activity", header=1, engine="openpyxl"
        )
        for _, row in meeting.iterrows():
            mtype = clean(row.get("Meeting Type"))
            mdate = to_iso_date(row.get("Date"))
            if not mtype or not mdate:
                continue
            conn.execute(
                """INSERT INTO meeting_activity
                   (activity_date,start_time,end_time,project_id,meeting_type,
                    attendee_1,attendee_2,attendee_3,attendee_4,location,agenda_notes)
                   VALUES (?,?,?,?,?,?,?,?,?,?,?)""",
                (
                    mdate,
                    to_iso_time(row.get("Start")),
                    to_iso_time(row.get("End")),
                    clean(row.get("Project ID")) or None,
                    mtype,
                    clean(row.get("Attendee 1")),
                    clean(row.get("Attendee 2")),
                    clean(row.get("Attendee 3")),
                    clean(row.get("Attendee 4")),
                    clean(row.get("Location")),
                    clean(row.get("Agenda / Notes")),
                ),
            )

        # ---- Other Activities
        other = pd.read_excel(
            SEED_XLSM, sheet_name="Other Activities", header=1, engine="openpyxl"
        )
        for _, row in other.iterrows():
            mdate = to_iso_date(row.get("Date"))
            activity = clean(row.get("Other Activity"))
            if not mdate or not activity:
                continue
            conn.execute(
                """INSERT INTO other_activity
                   (activity_date,activity,related_staff,notes)
                   VALUES (?,?,?,?)""",
                (
                    mdate,
                    activity,
                    clean(row.get("Related Staff")),
                    clean(row.get("Notes")),
                ),
            )

        # Reference lists from Setup.
        ref_ranges = {
            "project_type": (24, 25),
            "phase": (27, 29),
            "status": (31, 32),
            "meeting_type": (34, 35),
            "location": (37, 38),
            "activity_type": (40, 41),
            "priority": (43, 44),
            "project_size": (46, 47),
            "activity_status": (58, 59),
        }
        for cat, (col1, col2) in ref_ranges.items():
            for r in range(2, min(len(setup), 25)):
                val = clean(setup.iloc[r, col1])
                if val:
                    conn.execute(
                        "INSERT OR IGNORE INTO reference_values(category,value) VALUES (?,?)",
                        (cat, val),
                    )

        conn.commit()
        _clear_read_caches()

    except Exception as exc:
        conn.rollback()
        st.session_state["seed_error"] = str(exc)

    finally:
        conn.close()


# ============================================================
# HELPERS
# ============================================================

def _is_missing(v):
    # Handles None, NaN, NaT and pandas.NA consistently.
    try:
        result = pd.isna(v)
        return bool(result) if not hasattr(result, "__len__") else False
    except Exception:
        return False


def clean(v):
    if _is_missing(v):
        return ""
    return str(v).strip()


def numeric_or_none(v):
    if _is_missing(v):
        return None
    try:
        return float(v)
    except Exception:
        return None


def to_iso_date(v):
    if _is_missing(v):
        return None
    if isinstance(v, pd.Timestamp):
        return v.date().isoformat()
    if isinstance(v, datetime):
        return v.date().isoformat()
    if isinstance(v, date):
        return v.isoformat()
    try:
        parsed = pd.to_datetime(v, errors="coerce")
        if _is_missing(parsed):
            return None
        return parsed.date().isoformat()
    except Exception:
        return None


def safe_date(v, default=None):
    """Return a real datetime.date or default; never return pandas NaT."""
    if _is_missing(v):
        return default
    try:
        parsed = pd.to_datetime(v, errors="coerce")
        if _is_missing(parsed):
            return default
        return parsed.date()
    except Exception:
        return default


def to_iso_time(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return None
    if isinstance(v, pd.Timestamp):
        return v.strftime("%H:%M")
    if isinstance(v, datetime):
        return v.strftime("%H:%M")
    if isinstance(v, time):
        return v.strftime("%H:%M")
    s = clean(v)
    return s[:5] if len(s) >= 5 else s


def parse_date(s):
    return datetime.strptime(s, "%Y-%m-%d").date()


# ============================================================
# INDONESIA HOLIDAY CALENDAR
# ============================================================
# 2026: official national holidays + collective leave based on
# SKB 3 Menteri No. 5 Tahun 2025 / related official 2026 guidance.
# 2027: provisional public-calendar dates; update when the official
# SKB 3 Menteri for 2027 is issued.
#
# Sundays are handled separately as weekly holidays.
HOLIDAYS = {
    2026: {
        # National holidays
        "2026-01-01": "Tahun Baru Masehi",
        "2026-01-16": "Isra Mikraj Nabi Muhammad SAW",
        "2026-02-17": "Tahun Baru Imlek 2577 Kongzili",
        "2026-03-19": "Hari Suci Nyepi 1948 Saka",
        "2026-03-21": "Hari Raya Idul Fitri 1447 H",
        "2026-03-22": "Hari Raya Idul Fitri 1447 H",
        "2026-04-03": "Wafat Yesus Kristus",
        "2026-04-05": "Hari Kebangkitan Yesus Kristus",
        "2026-05-01": "Hari Buruh Internasional",
        "2026-05-14": "Kenaikan Yesus Kristus",
        "2026-05-27": "Hari Raya Idul Adha 1447 H",
        "2026-05-31": "Hari Raya Waisak 2570 BE",
        "2026-06-01": "Hari Lahir Pancasila",
        "2026-06-16": "Tahun Baru Islam 1448 H",
        "2026-08-17": "Hari Proklamasi Kemerdekaan RI",
        "2026-08-25": "Maulid Nabi Muhammad SAW",
        "2026-12-25": "Kelahiran Yesus Kristus",
        # Collective leave
        "2026-02-16": "Cuti Bersama Imlek",
        "2026-03-18": "Cuti Bersama Nyepi",
        "2026-03-20": "Cuti Bersama Idul Fitri",
        "2026-03-23": "Cuti Bersama Idul Fitri",
        "2026-03-24": "Cuti Bersama Idul Fitri",
        "2026-05-15": "Cuti Bersama Kenaikan Yesus Kristus",
        "2026-05-28": "Cuti Bersama Idul Adha",
        "2026-12-24": "Cuti Bersama Natal",
    },
    2027: {
        # Provisional public-calendar dates. These are intentionally
        # kept in code so they can be replaced when the official SKB
        # 3 Menteri 2027 is released.
        "2027-01-01": "Tahun Baru Masehi",
        "2027-01-05": "Isra Mikraj Nabi Muhammad SAW",
        "2027-02-06": "Tahun Baru Imlek 2578 Kongzili",
        "2027-03-09": "Hari Suci Nyepi / Idul Fitri (provisional)",
        "2027-03-10": "Hari Raya Idul Fitri 1448 H (provisional)",
        "2027-03-26": "Wafat Yesus Kristus",
        "2027-03-28": "Hari Kebangkitan Yesus Kristus",
        "2027-05-01": "Hari Buruh Internasional",
        "2027-05-06": "Kenaikan Yesus Kristus",
        "2027-05-17": "Hari Raya Idul Adha 1448 H (provisional)",
        "2027-05-20": "Hari Raya Waisak 2571 BE (provisional)",
        "2027-06-01": "Hari Lahir Pancasila",
        "2027-06-06": "Tahun Baru Islam 1449 H (provisional)",
        "2027-08-15": "Maulid Nabi Muhammad SAW (provisional)",
        "2027-08-17": "Hari Proklamasi Kemerdekaan RI",
        "2027-12-25": "Kelahiran Yesus Kristus",
        # Provisional collective leave
        "2027-02-05": "Cuti Bersama Imlek (provisional)",
        "2027-03-08": "Cuti Bersama Idul Fitri (provisional)",
        "2027-03-11": "Cuti Bersama Idul Fitri (provisional)",
        "2027-03-12": "Cuti Bersama Idul Fitri (provisional)",
        "2027-05-07": "Cuti Bersama Kenaikan Yesus Kristus (provisional)",
        "2027-05-18": "Cuti Bersama Idul Adha (provisional)",
        "2027-05-21": "Cuti Bersama Waisak (provisional)",
        "2027-12-24": "Cuti Bersama Natal (provisional)",
    },
}


def holiday_info(d):
    """Return (is_non_working_day, label) for Sundays or listed holidays."""
    if d.weekday() == 6:
        return True, "Minggu"
    label = HOLIDAYS.get(d.year, {}).get(d.isoformat())
    if label:
        return True, label
    return False, ""


def fmt_day(d):
    return d.strftime("%a")


@st.cache_data(ttl=30, show_spinner=False)
def project_name_map():
    conn = get_conn()
    df = pd.read_sql_query("SELECT id,name FROM projects ORDER BY id", conn)
    conn.close()
    return dict(zip(df["id"], df["name"]))


@st.cache_data(ttl=30, show_spinner=False)
def get_projects():
    conn = get_conn()
    df = pd.read_sql_query(
        "SELECT * FROM projects ORDER BY id", conn
    )
    conn.close()
    return df


@st.cache_data(ttl=30, show_spinner=False)
def get_staff():
    conn = get_conn()
    df = pd.read_sql_query(
        "SELECT * FROM staff WHERE active=1 ORDER BY name", conn
    )
    conn.close()
    return df


@st.cache_data(ttl=300, show_spinner=False)
def get_refs(category):
    conn = get_conn()
    rows = conn.execute(
        "SELECT value FROM reference_values WHERE category=? ORDER BY rowid",
        (category,),
    ).fetchall()
    conn.close()
    return [r[0] for r in rows]


@st.cache_data(ttl=20, show_spinner=False)
def load_activities(start_date, end_date, project_filter):
    conn = get_conn()

    params = [start_date.isoformat(), end_date.isoformat()]
    p_clause = ""
    if project_filter != "All Projects":
        p_clause = " AND project_id = ?"
        params.append(project_filter)

    work = pd.read_sql_query(
        f"""
        SELECT w.*, COALESCE(p.name,'') AS project_name
        FROM work_activity w
        LEFT JOIN projects p ON p.id=w.project_id
        WHERE w.end_date >= ? AND w.end_date <= ? {p_clause}
        ORDER BY w.end_date,
                 CAST(REPLACE(UPPER(COALESCE(w.project_id,'')), 'P', '') AS INTEGER),
                 CASE LOWER(COALESCE(w.priority,''))
                     WHEN 'high' THEN 1
                     WHEN 'medium' THEN 2
                     WHEN 'low' THEN 3
                     ELSE 4
                 END,
                 w.id
        """,
        conn,
        params=params,
    )

    params = [start_date.isoformat(), end_date.isoformat()]
    p_clause = ""
    if project_filter != "All Projects":
        p_clause = " AND project_id = ?"
        params.append(project_filter)

    meetings = pd.read_sql_query(
        f"""
        SELECT m.*, COALESCE(p.name,'') AS project_name
        FROM meeting_activity m
        LEFT JOIN projects p ON p.id=m.project_id
        WHERE m.activity_date >= ? AND m.activity_date <= ? {p_clause}
        ORDER BY m.activity_date,
                 CAST(REPLACE(UPPER(COALESCE(m.project_id,'')), 'P', '') AS INTEGER),
                 m.start_time, m.id
        """,
        conn,
        params=params,
    )

    # IMPORTANT: Other intentionally does NOT receive project filtering.
    others = pd.read_sql_query(
        """
        SELECT * FROM other_activity
        WHERE activity_date >= ? AND activity_date <= ?
        ORDER BY activity_date, id
        """,
        conn,
        params=[start_date.isoformat(), end_date.isoformat()],
    )

    conn.close()
    return work, meetings, others


# ============================================================
# WEEKLY DASHBOARD LOGIC
# ============================================================

def month_weeks(year, month):
    first = date(year, month, 1)
    last = date(year, month, calendar.monthrange(year, month)[1])

    # Studio calendar convention: Week 1 runs from the 1st of the month
    # through the first Sunday; subsequent weeks run Monday-Sunday.
    weeks = []
    current = first
    first_end = min(first + timedelta(days=(6 - first.weekday())), last)
    weeks.append((first, first_end))
    current = first_end + timedelta(days=1)

    while current <= last:
        week_end = min(current + timedelta(days=6), last)
        weeks.append((current, week_end))
        current = week_end + timedelta(days=1)
    return weeks


def group_items(df, date_col, project_col="project_id"):
    groups = {}
    if df.empty:
        return groups

    for _, row in df.iterrows():
        d = parse_date(row[date_col])
        pid = clean(row.get(project_col))
        key = (d, pid)
        groups.setdefault(key, []).append(row)
    return groups


def priority_rank(value):
    p = clean(value).lower()
    return {"high": 1, "medium": 2, "low": 3}.get(p, 99)


def render_work_cell(rows):
    if not rows:
        return '<div class="empty">—</div>'

    project_groups = {}
    for row in rows:
        pid = clean(row.get("project_id"))
        project_groups.setdefault(pid, []).append(row)

    project_items = sorted(project_groups.items(), key=lambda item: (item[0] == "", item[0]))
    visible_projects = project_items[:3]
    hidden_project_count = max(0, len(project_items) - 3)

    chunks = []

    for pid, project_rows in visible_projects:
        pname = clean(project_rows[0].get("project_name"))
        chunks.append(
            f'<div class="project">{html.escape(pid)} | {html.escape(pname)}</div>'
        )

        sorted_tasks = sorted(
            project_rows,
            key=lambda row: (
                priority_rank(row.get("priority")),
                clean(row.get("task")).lower(),
                int(row.get("id") or 0),
            )
        )

        visible_tasks = sorted_tasks[:3]
        hidden_task_count = max(0, len(sorted_tasks) - 3)

        for row in visible_tasks:
            task = html.escape(clean(row.get("task")))
            pic = html.escape(clean(row.get("pic")))
            activity_type = clean(row.get("activity_type")).lower()
            sign = '<span class="sign">!!</span>' if activity_type == "submission" else ""
            pic_html = f" <span style='color:#475467'>({pic.upper()})</span>" if pic else ""

            chunks.append(
                f'<div class="task {"submission" if activity_type == "submission" else ""}">'
                f'• {task}{pic_html} {sign}</div>'
            )

        if hidden_task_count:
            chunks.append(
                f'<div class="more-items">+{hidden_task_count} more task'
                f'{"s" if hidden_task_count != 1 else ""}</div>'
            )

    if hidden_project_count:
        chunks.append(
            f'<div class="more-items">+{hidden_project_count} more project'
            f'{"s" if hidden_project_count != 1 else ""}</div>'
        )

    return "".join(chunks)


def render_meeting_cell(rows):
    if not rows:
        return '<div class="empty">—</div>'

    # Same dashboard display rule as Work:
    # max 3 projects per date, max 3 meetings per project.
    project_groups = {}
    for row in rows:
        pid = clean(row.get("project_id"))
        project_groups.setdefault(pid, []).append(row)

    project_items = sorted(
        project_groups.items(),
        key=lambda item: (item[0] == "", item[0])
    )

    visible_projects = project_items[:3]
    hidden_project_count = max(0, len(project_items) - 3)

    chunks = []

    for pid, project_rows in visible_projects:
        pname = clean(project_rows[0].get("project_name"))
        chunks.append(
            f'<div class="project">{html.escape(pid)} | {html.escape(pname)}</div>'
        )

        # Keep meetings in chronological order within each project.
        sorted_meetings = sorted(
            project_rows,
            key=lambda row: (
                clean(row.get("start_time")),
                clean(row.get("meeting_type")).lower(),
                int(row.get("id") or 0),
            )
        )

        visible_meetings = sorted_meetings[:3]
        hidden_meeting_count = max(0, len(sorted_meetings) - 3)

        for row in visible_meetings:
            attendees = [
                clean(row.get("attendee_1")),
                clean(row.get("attendee_2")),
                clean(row.get("attendee_3")),
                clean(row.get("attendee_4")),
            ]
            attendees = ", ".join([a.upper() for a in attendees if a])

            chunks.append(
                '<div class="task">• '
                + html.escape(clean(row.get("meeting_type")))
                + '</div>'
            )
            chunks.append(
                f'<div class="meta">PIC: {html.escape(attendees)}</div>'
            )
            chunks.append(
                f'<div class="meta">Time: {html.escape(clean(row.get("start_time")))}'
                f' – {html.escape(clean(row.get("end_time")))}</div>'
            )
            chunks.append(
                f'<div class="meta">Location: {html.escape(clean(row.get("location")))}</div>'
            )

        if hidden_meeting_count:
            chunks.append(
                f'<div class="more-items">+{hidden_meeting_count} more meeting'
                f'{"s" if hidden_meeting_count != 1 else ""}</div>'
            )

    if hidden_project_count:
        chunks.append(
            f'<div class="more-items">+{hidden_project_count} more project'
            f'{"s" if hidden_project_count != 1 else ""}</div>'
        )

    return "".join(chunks)


def render_other_cell(rows):
    if not rows:
        return '<div class="empty">—</div>'

    # Other activities are non-project activities, so the limit is
    # max 3 activities per date.
    sorted_rows = sorted(
        rows,
        key=lambda row: (
            clean(row.get("activity")).lower(),
            clean(row.get("related_staff")).lower(),
            int(row.get("id") or 0),
        )
    )

    visible_rows = sorted_rows[:3]
    hidden_count = max(0, len(sorted_rows) - 3)

    chunks = []
    for row in visible_rows:
        activity = html.escape(clean(row.get("activity")))
        staff = html.escape(clean(row.get("related_staff")))
        chunks.append(
            f'<div class="other-item">• {activity}'
            + (f'<div class="meta">Related Staff: {staff.upper()}</div>' if staff else "")
            + "</div>"
        )

    if hidden_count:
        chunks.append(
            f'<div class="more-items">+{hidden_count} more activit'
            f'{"ies" if hidden_count != 1 else "y"}</div>'
        )

    return "".join(chunks)


def render_week(start_date, end_date, work, meetings, others, visible_activities, week_no=None):
    """Render one week as a single CSS grid so each lane row takes the
    height of its tallest populated day. This keeps empty date cells aligned.
    """
    days = []
    d = start_date
    while d <= end_date:
        days.append(d)
        d += timedelta(days=1)

    work_groups = group_items(work, "end_date")
    meeting_groups = group_items(meetings, "activity_date")

    other_groups = {}
    if not others.empty:
        for _, row in others.iterrows():
            d = parse_date(row["activity_date"])
            other_groups.setdefault(d, []).append(row)

    if week_no is None:
        week_no = 1
    st.markdown(
        f'<div class="week-title">WEEK {week_no} • {start_date.strftime("%d %b")} – '
        f'{end_date.strftime("%d %b %Y")}</div>',
        unsafe_allow_html=True,
    )

    grid = []
    # Header row
    grid.append('<div class="schedule-head activity-head">ACTIVITY</div>')
    for d in days:
        is_non_working, holiday_label = holiday_info(d)
        holiday_class = " non-working-day" if is_non_working else ""
        title = holiday_label if holiday_label else d.strftime("%d %b %Y")
        grid.append(
            f'<div class="schedule-head{holiday_class}">'
            f'<a class="date-detail-link" href="?detail_date={d.isoformat()}" target="_self" '
            f'title="Open all activities • {html.escape(title)}">'
            f'<div class="dow">{fmt_day(d)}</div>'
            f'<div class="day">{d.strftime("%d %b")}</div>'
            f'</a></div>'
        )

    lane_specs = [
        ("WORK", "work", "▣"),
        ("MEETING", "meeting", "●"),
        ("OTHER", "other", "•••"),
    ]
    lane_specs = [x for x in lane_specs if x[1] in visible_activities]

    for lane_name, lane_type, lane_icon in lane_specs:
        grid.append(
            f'<div class="lane-label {lane_type}-lane"><span class="lane-icon">{lane_icon}</span>{lane_name}</div>'
        )

        for d in days:
            if lane_type == "work":
                rows = []
                for (gd, pid), items in work_groups.items():
                    if gd == d:
                        rows.extend(items)
                content = render_work_cell(rows)
                css_class = "cell work-cell"
            elif lane_type == "meeting":
                rows = []
                for (gd, pid), items in meeting_groups.items():
                    if gd == d:
                        rows.extend(items)
                content = render_meeting_cell(rows)
                css_class = "cell meeting-cell"
            else:
                content = render_other_cell(other_groups.get(d, []))
                css_class = "cell other-cell"

            grid.append(f'<div class="{css_class}">{content}</div>')

    st.markdown(
        f'<div class="schedule-grid" style="grid-template-columns: minmax(105px, .8fr) repeat({len(days)}, minmax(135px, 1fr));">' + ''.join(grid) + '</div>',
        unsafe_allow_html=True,
    )



def checklist_filter(label, options, all_label, key_prefix, format_func=None):
    """Popover-based multi-select filter.

    UX rule: All is mutually exclusive. While All is active, individual
    choices are hidden. Clearing All reveals the individual checkboxes.
    """
    all_key = f"{key_prefix}_all"
    if all_key not in st.session_state:
        st.session_state[all_key] = True

    for opt in options:
        key = f"{key_prefix}_{opt}"
        if key not in st.session_state:
            st.session_state[key] = False

    # If all is selected, clear individual selections so the state is clean.
    if st.session_state[all_key]:
        for opt in options:
            st.session_state[f"{key_prefix}_{opt}"] = False

    chosen = [opt for opt in options if st.session_state[f"{key_prefix}_{opt}"]]
    if st.session_state[all_key]:
        summary = all_label
    elif not chosen:
        summary = "None"
    elif len(chosen) == 1:
        summary = format_func(chosen[0]) if format_func else chosen[0]
    else:
        summary = f"{len(chosen)} selected"

    # Keep the label outside the control so Project/Activity align vertically
    # with the native Month/Year/Week selectboxes.
    st.markdown(f'<div class="filter-label">{html.escape(label)}</div>', unsafe_allow_html=True)
    with st.popover(summary, use_container_width=True):
        all_selected = st.checkbox(all_label, key=all_key)
        if all_selected:
            st.caption("All selected")
        else:
            for opt in options:
                text = format_func(opt) if format_func else opt
                st.checkbox(text, key=f"{key_prefix}_{opt}")

    selected = [opt for opt in options if st.session_state[f"{key_prefix}_{opt}"]]
    return (list(options), True) if st.session_state[all_key] else (selected, False)




@st.cache_resource(show_spinner=False)
def _pdf_font_names():
    """Use DejaVu Sans when available for broader Unicode support."""
    if not REPORTLAB_AVAILABLE:
        return ("Helvetica", "Helvetica-Bold")

    regular = Path("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf")
    bold = Path("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf")
    if regular.exists() and bold.exists():
        try:
            if "ArayaSans" not in pdfmetrics.getRegisteredFontNames():
                pdfmetrics.registerFont(TTFont("ArayaSans", str(regular)))
            if "ArayaSansBold" not in pdfmetrics.getRegisteredFontNames():
                pdfmetrics.registerFont(TTFont("ArayaSansBold", str(bold)))
            return ("ArayaSans", "ArayaSansBold")
        except Exception:
            pass
    return ("Helvetica", "Helvetica-Bold")


def _pdf_safe(value):
    """PDF-friendly text while avoiding problematic Unicode dash/bullet glyphs."""
    s=clean(value)
    s=(s.replace("•","-")
         .replace("–","-")
         .replace("—","-")
         .replace("…","..."))
    return html.escape(s).replace("\n","<br/>")


def generate_date_detail_pdf(detail_date, work, meetings, others, visible_activities):
    """Generate a clean A4 PDF for the selected date using current dashboard filters."""
    if not REPORTLAB_AVAILABLE:
        return None

    regular_font,bold_font=_pdf_font_names()
    bio=BytesIO()

    doc=SimpleDocTemplate(
        bio,
        pagesize=A4,
        rightMargin=15*mm,
        leftMargin=15*mm,
        topMargin=14*mm,
        bottomMargin=14*mm,
        title=f"ARAYASTD Activity Detail {detail_date.isoformat()}",
        author="ARAYASTD Studio Control Board",
    )

    styles=getSampleStyleSheet()
    title_style=ParagraphStyle(
        "ArayaTitle", parent=styles["Title"],
        fontName=bold_font, fontSize=16, leading=20,
        textColor=colors.HexColor("#172B4D"),
        alignment=TA_CENTER, spaceAfter=4*mm,
    )
    date_style=ParagraphStyle(
        "ArayaDate", parent=styles["Normal"],
        fontName=regular_font, fontSize=9.5, leading=13,
        textColor=colors.HexColor("#475467"),
        alignment=TA_CENTER, spaceAfter=6*mm,
    )
    section_style=ParagraphStyle(
        "ArayaSection", parent=styles["Heading2"],
        fontName=bold_font, fontSize=11, leading=14,
        textColor=colors.HexColor("#172B4D"),
        spaceAfter=2.5*mm,
    )
    item_style=ParagraphStyle(
        "ArayaItem", parent=styles["BodyText"],
        fontName=regular_font, fontSize=8.6, leading=12,
        textColor=colors.HexColor("#1F2937"),
        leftIndent=2*mm, spaceAfter=1.3*mm,
    )
    meta_style=ParagraphStyle(
        "ArayaMeta", parent=item_style,
        fontSize=7.8, leading=10.5,
        textColor=colors.HexColor("#667085"),
        leftIndent=6*mm,
    )
    empty_style=ParagraphStyle(
        "ArayaEmpty", parent=item_style,
        textColor=colors.HexColor("#98A2B3"),
    )

    story=[
        Paragraph("ARAYASTD - Daily Activity Detail", title_style),
        Paragraph(detail_date.strftime("%A, %d %B %Y"), date_style),
    ]

    section_colors={
        "work": colors.HexColor("#EAF4FF"),
        "meeting": colors.HexColor("#FFF6D8"),
        "other": colors.HexColor("#F2E9FF"),
    }

    def section_header(label, kind):
        t=Table(
            [[Paragraph(label, section_style)]],
            colWidths=[180*mm],
            hAlign="LEFT",
        )
        t.setStyle(TableStyle([
            ("BACKGROUND",(0,0),(-1,-1),section_colors[kind]),
            ("BOX",(0,0),(-1,-1),0.5,colors.HexColor("#D0D5DD")),
            ("LEFTPADDING",(0,0),(-1,-1),3*mm),
            ("RIGHTPADDING",(0,0),(-1,-1),3*mm),
            ("TOPPADDING",(0,0),(-1,-1),2.2*mm),
            ("BOTTOMPADDING",(0,0),(-1,-1),1.2*mm),
        ]))
        story.append(t)
        story.append(Spacer(1,2.2*mm))

    # WORK
    if "work" in visible_activities:
        section_header("WORK", "work")
        wrows=work[work["end_date"]==detail_date.isoformat()].copy() if not work.empty else work.iloc[0:0]
        if wrows.empty:
            story.append(Paragraph("No work activity.", empty_style))
        else:
            groups={}
            for _,row in wrows.iterrows():
                groups.setdefault(clean(row.get("project_id")),[]).append(row)
            for pid,rows in sorted(groups.items(),key=lambda x:(x[0]=="",x[0])):
                pname=clean(rows[0].get("project_name"))
                story.append(Paragraph(
                    f"<b>{_pdf_safe(pid)} | {_pdf_safe(pname)}</b>", item_style
                ))
                for row in sorted(rows,key=lambda r:(priority_rank(r.get("priority")),clean(r.get("task")).lower(),int(r.get("id") or 0))):
                    task=_pdf_safe(row.get("task"))
                    pic=_pdf_safe(row.get("pic"))
                    priority=_pdf_safe(row.get("priority"))
                    atype=_pdf_safe(row.get("activity_type"))
                    story.append(Paragraph(f"- {task}", item_style))
                    meta=[]
                    if pic: meta.append(f"PIC: {pic.upper()}")
                    if atype: meta.append(f"Type: {atype}")
                    if priority: meta.append(f"Priority: {priority.upper()}")
                    if meta:
                        story.append(Paragraph(" | ".join(meta), meta_style))
        story.append(Spacer(1,4*mm))

    # MEETINGS
    if "meeting" in visible_activities:
        section_header("MEETINGS", "meeting")
        mrows=meetings[meetings["activity_date"]==detail_date.isoformat()].copy() if not meetings.empty else meetings.iloc[0:0]
        if mrows.empty:
            story.append(Paragraph("No meeting activity.", empty_style))
        else:
            for _,row in mrows.sort_values(["start_time","id"],na_position="last").iterrows():
                mtype=_pdf_safe(row.get("meeting_type"))
                story.append(Paragraph(f"- {mtype}", item_style))
                attendees=[
                    _pdf_safe(row.get("attendee_1")),
                    _pdf_safe(row.get("attendee_2")),
                    _pdf_safe(row.get("attendee_3")),
                    _pdf_safe(row.get("attendee_4")),
                ]
                attendees=", ".join([x for x in attendees if x])
                meta=[]
                start=_pdf_safe(row.get("start_time"))
                end=_pdf_safe(row.get("end_time"))
                location=_pdf_safe(row.get("location"))
                project_id=_pdf_safe(row.get("project_id"))
                project_name=_pdf_safe(row.get("project_name"))
                agenda=_pdf_safe(row.get("agenda_notes"))
                if project_id or project_name: meta.append(f"Project: {project_id} | {project_name}")
                if start or end: meta.append(f"Time: {start} - {end}")
                if attendees: meta.append(f"Attendees: {attendees}")
                if location: meta.append(f"Location: {location}")
                if agenda: meta.append(f"Notes: {agenda}")
                for line in meta:
                    story.append(Paragraph(line, meta_style))
        story.append(Spacer(1,4*mm))

    # OTHER
    if "other" in visible_activities:
        section_header("OTHER ACTIVITIES", "other")
        orows=others[others["activity_date"]==detail_date.isoformat()].copy() if not others.empty else others.iloc[0:0]
        if orows.empty:
            story.append(Paragraph("No other activity.", empty_style))
        else:
            for _,row in orows.sort_values(["activity","id"],na_position="last").iterrows():
                activity=_pdf_safe(row.get("activity"))
                story.append(Paragraph(f"- {activity}", item_style))
                staff=_pdf_safe(row.get("related_staff"))
                notes=_pdf_safe(row.get("notes"))
                if staff:
                    story.append(Paragraph(f"Related Staff: {staff.upper()}", meta_style))
                if notes:
                    story.append(Paragraph(f"Notes: {notes}", meta_style))

    doc.build(story)
    bio.seek(0)
    return bio.getvalue()


@st.dialog("Activity Detail", width="large")
def show_date_detail(detail_date, work, meetings, others, visible_activities):
    """Full-detail modal for one date. No dashboard display limits.

    The three activity cards are each emitted as one HTML block so all
    activity content stays physically inside its coloured card.
    """
    st.markdown(
        f'<div class="detail-date">{detail_date.strftime("%A, %d %B %Y")}</div>',
        unsafe_allow_html=True,
    )

    if REPORTLAB_AVAILABLE:
        pdf_bytes=generate_date_detail_pdf(
            detail_date,work,meetings,others,visible_activities
        )
        st.download_button(
            "Generate PDF",
            data=pdf_bytes,
            file_name=f"ARAYASTD_Activity_Detail_{detail_date.isoformat()}.pdf",
            mime="application/pdf",
            key=f"detail_pdf_{detail_date.isoformat()}",
            use_container_width=True,
        )
    else:
        st.caption("PDF generator belum tersedia. Pastikan reportlab terpasang.")

    def priority_class(value):
        p = clean(value).lower()
        if p == "high":
            return "priority-high"
        if p == "low":
            return "priority-low"
        if p == "medium":
            return "priority-medium"
        return "priority-none"

    def work_card_html():
        html_parts = [
            '<div class="detail-card detail-work">',
            '<div class="detail-card-head"><span class="detail-card-icon">▣</span> WORK</div>',
        ]

        wrows = (
            work[work["end_date"] == detail_date.isoformat()].copy()
            if not work.empty else work.iloc[0:0]
        )

        if wrows.empty:
            html_parts.append('<div class="detail-empty">No work activity.</div>')
        else:
            groups = {}
            for _, row in wrows.iterrows():
                groups.setdefault(clean(row.get("project_id")), []).append(row)

            for pid, rows in sorted(groups.items(), key=lambda x: (x[0] == "", x[0])):
                pname = clean(rows[0].get("project_name"))
                html_parts.append(
                    f'<div class="detail-project">{html.escape(pid)} | '
                    f'{html.escape(pname)}</div>'
                )

                rows = sorted(
                    rows,
                    key=lambda r: (
                        priority_rank(r.get("priority")),
                        clean(r.get("task")).lower(),
                        int(r.get("id") or 0),
                    ),
                )

                for row in rows:
                    task = html.escape(clean(row.get("task")))
                    pic = html.escape(clean(row.get("pic")))
                    priority = clean(row.get("priority"))
                    activity_type = clean(row.get("activity_type")).lower()

                    sign = '<span class="detail-submit">!!</span>' if activity_type == "submission" else ""
                    pic_html = f'<span class="detail-pic">({pic.upper()})</span>' if pic else ""

                    priority_html = ""
                    if priority:
                        priority_html = (
                            f'<span class="priority-tag {priority_class(priority)}">'
                            f'{html.escape(priority.upper())}</span>'
                        )

                    if activity_type == "submission":
                        html_parts.append(
                            f'<div class="detail-task detail-submission">'
                            f'<span class="detail-bullet">•</span> {task} {pic_html}'
                            f' {priority_html} '
                            f'<span class="detail-submit-badge">!!</span>'
                            f'</div>'
                        )
                    else:
                        html_parts.append(
                            f'<div class="detail-task">'
                            f'<span class="detail-bullet">•</span> {task} {pic_html}'
                            f' {priority_html}'
                            f'</div>'
                        )

        html_parts.append("</div>")
        return "".join(html_parts)

    def meeting_card_html():
        html_parts = [
            '<div class="detail-card detail-meeting">',
            '<div class="detail-card-head"><span class="detail-card-icon">●</span> MEETINGS</div>',
        ]

        mrows = (
            meetings[meetings["activity_date"] == detail_date.isoformat()].copy()
            if not meetings.empty else meetings.iloc[0:0]
        )

        if mrows.empty:
            html_parts.append('<div class="detail-empty">No meeting activity.</div>')
        else:
            groups = {}
            for _, row in mrows.iterrows():
                groups.setdefault(clean(row.get("project_id")), []).append(row)

            for pid, rows in sorted(groups.items(), key=lambda x: (x[0] == "", x[0])):
                if pid:
                    pname = clean(rows[0].get("project_name"))
                    html_parts.append(
                        f'<div class="detail-project">{html.escape(pid)} | '
                        f'{html.escape(pname)}</div>'
                    )

                rows = sorted(
                    rows,
                    key=lambda r: (
                        clean(r.get("start_time")),
                        clean(r.get("meeting_type")).lower(),
                        int(r.get("id") or 0),
                    ),
                )

                for row in rows:
                    attendees = [
                        clean(row.get("attendee_1")),
                        clean(row.get("attendee_2")),
                        clean(row.get("attendee_3")),
                        clean(row.get("attendee_4")),
                    ]
                    attendees = ", ".join(a.upper() for a in attendees if a)

                    html_parts.append(
                        f'<div class="detail-meeting-item">'
                        f'<div class="detail-task"><span class="detail-bullet">•</span> '
                        f'{html.escape(clean(row.get("meeting_type")))}</div>'
                        f'<div class="detail-meta"><b>PIC:</b> {html.escape(attendees) or "—"}</div>'
                        f'<div class="detail-meta"><b>Time:</b> '
                        f'{html.escape(clean(row.get("start_time")))} – '
                        f'{html.escape(clean(row.get("end_time")))}</div>'
                        f'<div class="detail-meta"><b>Location:</b> '
                        f'{html.escape(clean(row.get("location")))}</div>'
                        f'</div>'
                    )

        html_parts.append("</div>")
        return "".join(html_parts)

    def other_card_html():
        html_parts = [
            '<div class="detail-card detail-other">',
            '<div class="detail-card-head"><span class="detail-card-icon">•••</span> OTHER ACTIVITIES</div>',
        ]

        orows = (
            others[others["activity_date"] == detail_date.isoformat()].copy()
            if not others.empty else others.iloc[0:0]
        )

        if orows.empty:
            html_parts.append('<div class="detail-empty">No other activity.</div>')
        else:
            rows = sorted(
                [row for _, row in orows.iterrows()],
                key=lambda r: (
                    clean(r.get("activity")).lower(),
                    clean(r.get("related_staff")).lower(),
                    int(r.get("id") or 0),
                ),
            )

            for row in rows:
                activity = html.escape(clean(row.get("activity")))
                staff = html.escape(clean(row.get("related_staff")))

                html_parts.append(
                    f'<div class="detail-other-item">'
                    f'<div class="detail-task"><span class="detail-bullet">•</span> {activity}</div>'
                    + (
                        f'<div class="detail-meta"><b>Related Staff:</b> {staff.upper()}</div>'
                        if staff else ""
                    )
                    + '</div>'
                )

        html_parts.append("</div>")
        return "".join(html_parts)

    cards = []
    if "work" in visible_activities:
        cards.append(work_card_html())
    else:
        cards.append(
            '<div class="detail-card detail-work">'
            '<div class="detail-card-head"><span class="detail-card-icon">▣</span> WORK</div>'
            '<div class="detail-empty">Hidden by Activity filter.</div></div>'
        )

    if "meeting" in visible_activities:
        cards.append(meeting_card_html())
    else:
        cards.append(
            '<div class="detail-card detail-meeting">'
            '<div class="detail-card-head"><span class="detail-card-icon">●</span> MEETINGS</div>'
            '<div class="detail-empty">Hidden by Activity filter.</div></div>'
        )

    if "other" in visible_activities:
        cards.append(other_card_html())
    else:
        cards.append(
            '<div class="detail-card detail-other">'
            '<div class="detail-card-head"><span class="detail-card-icon">•••</span> OTHER ACTIVITIES</div>'
            '<div class="detail-empty">Hidden by Activity filter.</div></div>'
        )

    # One HTML wrapper keeps each card's entire content inside its coloured box.
    st.markdown(
        '<div class="detail-card-grid">'
        + "".join(f'<div class="detail-card-column">{card}</div>' for card in cards)
        + '</div>',
        unsafe_allow_html=True,
    )


def beranda_page():
    """Minimal landing page; future executive dashboard will grow from here."""
    st.markdown(
        f"""
        <div class="araya-home">
            <div class="araya-home-inner">
                <img class="araya-home-logo" src="{ARAYA_LOGO_DATA_URI}" alt="Araya Studio logo">
                <div class="araya-home-rule"></div>
                <div class="araya-home-title">ARAYA STUDIO</div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def weekly_dashboard():
    projects = get_projects()
    project_options = (["All"] + projects["id"].tolist()) if not projects.empty else ["All"]

    # First dashboard load in a Streamlit session always opens the actual
    # current month/year in Indonesia (WIB). User selections remain persistent
    # after that, including when opening a date-detail dialog.
    dashboard_today=datetime.now(ZoneInfo("Asia/Jakarta")).date()
    if not st.session_state.get("dash_calendar_initialized"):
        st.session_state["dash_month"]=dashboard_today.month
        st.session_state["dash_year"]=dashboard_today.year
        st.session_state["dash_week"]="All"
        st.session_state["dash_calendar_initialized"]=True

    # Date-detail overlay.
    #
    # The dashboard calendar itself is rendered as HTML, so clicking a date
    # changes the URL query string. Treat that query string only as a one-time
    # trigger: copy the clicked date into Session State, restore the dashboard
    # month/year from that date, clear the URL, then rerun once. On the next
    # run the dialog opens from Session State without resetting to today's
    # month and it will not reopen after the user closes it.
    detail_date_value = st.query_params.get("detail_date")
    if detail_date_value:
        try:
            clicked_date = parse_date(detail_date_value)
            st.session_state["dash_month"] = clicked_date.month
            st.session_state["dash_year"] = clicked_date.year
            st.session_state["dash_detail_pending"] = clicked_date.isoformat()
        except Exception:
            st.session_state.pop("dash_detail_pending", None)

        # Do not use pop() here. Query-param mutation can itself trigger a
        # rerun; clearing after state has been captured makes the transition
        # deterministic.
        st.query_params.clear()
        st.rerun()

    detail_date = None
    pending_detail = st.session_state.pop("dash_detail_pending", None)
    if pending_detail:
        try:
            detail_date = parse_date(pending_detail)
        except Exception:
            detail_date = None

    # Calendar range is derived from data + current year; no Setup entry needed.
    conn = get_conn()
    date_values = []
    for table, col in [
        ("work_activity", "end_date"),
        ("meeting_activity", "activity_date"),
        ("other_activity", "activity_date"),
    ]:
        rows = conn.execute(
            f"SELECT {col} FROM {table} WHERE {col} IS NOT NULL AND {col} <> ''"
        ).fetchall()
        date_values.extend([r[0] for r in rows if r[0]])
    conn.close()

    years = {dashboard_today.year}
    for value in date_values:
        try:
            years.add(parse_date(value).year)
        except Exception:
            pass
    years = sorted(years)

    month_names = [
        "January", "February", "March", "April", "May", "June",
        "July", "August", "September", "October", "November", "December"
    ]

    # ========================================================
    # STICKY DASHBOARD HEADER
    # ========================================================
    header = st.container()
    with header:
        st.markdown('<div id="weekly-dashboard-anchor"></div>', unsafe_allow_html=True)
        st.markdown('<div class="app-title">Weekly Schedule Dashboard</div>', unsafe_allow_html=True)
        st.markdown(
            '<div class="app-subtitle">Leadership view • Work & Meeting follow project filter • '
            'Other remains independent of project filter</div>',
            unsafe_allow_html=True,
        )

        c1, c2, c3, c4, c5 = st.columns([1.05, .82, 1.55, 1.15, 1.15], gap="small")
        with c1:
            selected_month_no = st.selectbox(
                "MONTH", range(1, 13), index=dashboard_today.month - 1,
                format_func=lambda x: f"📅  {month_names[x - 1]}", key="dash_month"
            )
        with c2:
            default_year_index = years.index(dashboard_today.year) if dashboard_today.year in years else len(years) - 1
            selected_year = st.selectbox("YEAR", years, index=default_year_index,
                                         format_func=lambda x: f"📅  {x}", key="dash_year")

        selected_month = date(selected_year, selected_month_no, 1)
        weeks = month_weeks(selected_month.year, selected_month.month)

        with c3:
            selected_projects, project_all = checklist_filter(
                "📁 PROJECT",
                projects["id"].tolist() if not projects.empty else [],
                "All Projects",
                "dash_project",
                format_func=lambda x: (
                    f"{x} | {projects.loc[projects['id'].eq(x), 'name'].iloc[0]}"
                    if not projects.loc[projects['id'].eq(x)].empty else x
                ),
            )
        with c4:
            week_options = ["All"] + [
                f"Week {i+1} ({w[0].strftime('%d')}–{w[1].strftime('%d %b')})"
                for i, w in enumerate(weeks)
            ]
            if st.session_state.get("dash_week") not in week_options:
                st.session_state["dash_week"]="All"
            selected_week = st.selectbox("WEEK", week_options, key="dash_week")
        with c5:
            selected_activity, activity_all = checklist_filter(
                "▱ ACTIVITY",
                ["Work", "Meeting", "Other"],
                "All",
                "dash_activity",
            )

        # Effective filters. Empty individual selection is treated as All to
        # avoid an accidental blank dashboard after clearing the checklist.
        project_filter = "All Projects" if project_all or not selected_projects else selected_projects
        if activity_all or not selected_activity:
            visible_activities = {"work", "meeting", "other"}
        else:
            visible_activities = {x.lower() for x in selected_activity}

        month_start = selected_month
        month_end = date(
            selected_month.year, selected_month.month,
            calendar.monthrange(selected_month.year, selected_month.month)[1],
        )

        # Load the month without project restriction, then apply multi-select.
        work, meetings, others = load_activities(month_start, month_end, "All Projects")
        if selected_projects and "All" not in selected_projects:
            work = work[work["project_id"].isin(selected_projects)].copy()
            meetings = meetings[meetings["project_id"].isin(selected_projects)].copy()
            # Other intentionally remains independent of project filter.

        # KPI cards follow all dashboard filters above:
        # Month/Year, Project, Week and Activity.
        kpi_work=work.copy()
        kpi_meetings=meetings.copy()
        kpi_others=others.copy()

        if selected_week != "All":
            kpi_idx=int(selected_week.split()[1])
            kpi_start,kpi_end=weeks[kpi_idx-1]
            kpi_work=kpi_work[
                (kpi_work["end_date"] >= kpi_start.isoformat())
                & (kpi_work["end_date"] <= kpi_end.isoformat())
            ].copy()
            kpi_meetings=kpi_meetings[
                (kpi_meetings["activity_date"] >= kpi_start.isoformat())
                & (kpi_meetings["activity_date"] <= kpi_end.isoformat())
            ].copy()
            kpi_others=kpi_others[
                (kpi_others["activity_date"] >= kpi_start.isoformat())
                & (kpi_others["activity_date"] <= kpi_end.isoformat())
            ].copy()

        work_count=len(kpi_work) if "work" in visible_activities else 0
        meeting_count=len(kpi_meetings) if "meeting" in visible_activities else 0
        submission_count=(
            int((kpi_work["activity_type"].str.lower()=="submission").sum())
            if "work" in visible_activities and not kpi_work.empty else 0
        )
        other_count=len(kpi_others) if "other" in visible_activities else 0

        icon_work = '''<svg width="28" height="28" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M6 2h9l3 3v17H6z"/><path d="M15 2v4h4"/><path d="M9 11h6M9 15h6M9 19h4"/></svg>'''
        icon_meeting = '''<svg width="30" height="30" viewBox="0 0 24 24" fill="currentColor"><circle cx="9" cy="8" r="4"/><circle cx="17" cy="9" r="3"/><path d="M2.5 21c.3-4 2.6-6 6.5-6s6.2 2 6.5 6z"/><path d="M14.5 15.5c3.2.1 5 1.8 5.5 4.5h-4.2c-.2-1.7-.6-3.1-1.3-4.5z"/></svg>'''
        icon_submission = '''<svg width="30" height="30" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round"><circle cx="5" cy="6" r="1.2" fill="currentColor"/><circle cx="5" cy="12" r="1.2" fill="currentColor"/><circle cx="5" cy="18" r="1.2" fill="currentColor"/><path d="M10 6h10M10 12h10M10 18h10"/></svg>'''
        icon_other = '''<svg width="30" height="30" viewBox="0 0 24 24" fill="currentColor"><circle cx="12" cy="12" r="9"/><circle cx="8" cy="12" r="1.25" fill="white"/><circle cx="12" cy="12" r="1.25" fill="white"/><circle cx="16" cy="12" r="1.25" fill="white"/></svg>'''

        cards = [
            ("work", "WORK ITEMS", work_count, icon_work),
            ("meeting", "MEETINGS", meeting_count, icon_meeting),
            ("submission", "SUBMISSIONS", submission_count, icon_submission),
            ("other", "OTHER ACTIVITIES", other_count, icon_other),
        ]

        cards_html = '<div class="kpi-row">'
        for css_class, label, value, icon in cards:
            cards_html += (
                f'<div class="kpi {css_class}">'
                f'<div class="kpi-icon">{icon}</div>'
                f'<div class="kpi-content">'
                f'<div class="kpi-label">{label}</div>'
                f'<div class="kpi-value">{value}</div>'
                f'</div></div>'
            )
        cards_html += '</div>'
        st.markdown(cards_html, unsafe_allow_html=True)

        if detail_date is not None:
            show_date_detail(
                detail_date,
                work,
                meetings,
                others,
                visible_activities,
            )

    selected_weeks = list(enumerate(weeks, start=1))
    if selected_week != "All":
        idx = int(selected_week.split()[1])
        selected_weeks = [selected_weeks[idx - 1]]

    for _, (week_start, week_end) in selected_weeks:
        w = work[(work["end_date"] >= week_start.isoformat()) & (work["end_date"] <= week_end.isoformat())].copy()
        m = meetings[(meetings["activity_date"] >= week_start.isoformat()) & (meetings["activity_date"] <= week_end.isoformat())].copy()
        o = others[(others["activity_date"] >= week_start.isoformat()) & (others["activity_date"] <= week_end.isoformat())].copy()

        if "work" not in visible_activities:
            w = w.iloc[0:0]
        if "meeting" not in visible_activities:
            m = m.iloc[0:0]
        if "other" not in visible_activities:
            o = o.iloc[0:0]

        render_week(week_start, week_end, w, m, o, visible_activities, week_no=_)


# ============================================================
# INPUT MODULES
# ============================================================



# ============================================================
# V3B DATABASE FOUNDATION
# ============================================================
# SETUP.xlsx is used only as initial seed.
# After database creation, application reads SQLite database.

DB_FILE = str(DB_PATH)
SETUP_FILE = "SETUP.xlsx"




def init_master_database():
    conn = get_conn()
    cur = conn.cursor()

    # Migration safety: remove incompatible old master schema
    try:
        cols = table_columns(conn, "master_role")
        if cols and "role" not in cols:
            for t in ["role","project_type","project_status","meeting_type","meeting_location","activity_type","priority","mapping_status","task_status","phase","project_size","workload_status"]:
                cur.execute(f"DROP TABLE IF EXISTS master_{t}")
    except Exception:
        pass

    # Master tables sesuai struktur SETUP.xlsx
    schemas = {
        "role": "role TEXT",
        "project_type": "project_type TEXT",
         "project_status": "status TEXT",
        "meeting_type": "meeting_type TEXT",
        "meeting_location": "location TEXT",
        "activity_type": "activity_type TEXT",
        "priority": "priority TEXT",
        "mapping_status": "status TEXT",
        "task_status": "status TEXT",
        "phase": "phase TEXT, sequence INTEGER, base_load REAL",
        "project_size": "project_size TEXT, multiplier REAL",
        "workload_status": "status TEXT, max_load REAL",
    }

    for name, cols in schemas.items():
        cur.execute(f"CREATE TABLE IF NOT EXISTS master_{name} (id INTEGER PRIMARY KEY AUTOINCREMENT,{cols})")

    # Import hanya jika database masih kosong
    count = cur.execute("SELECT COUNT(*) FROM master_role").fetchone()[0]

    if count == 0:
        df = pd.read_excel(SETUP_FILE, sheet_name="Setup", header=None)

        def insert_list(table, col, field):
            for val in df.iloc[2:, col]:
                if pd.notna(val):
                    cur.execute(
                        f"INSERT INTO master_{table}({field}) VALUES (?)",
                        (str(val).strip(),)
                    )

        insert_list("role",0,"role")
        insert_list("project_type",2,"project_type")
        insert_list("project_status",8,"status")
        insert_list("meeting_type",10,"meeting_type")
        insert_list("meeting_location",12,"location")
        insert_list("activity_type",14,"activity_type")
        insert_list("priority",16,"priority")
        insert_list("mapping_status",24,"status")
        insert_list("task_status",26,"status")

        for _,r in df.iloc[2:,4:7].dropna(how="all").iterrows():
            if pd.notna(r[4]):
                cur.execute("INSERT INTO master_phase(phase,sequence,base_load) VALUES (?,?,?)",
                            (str(r[4]), int(r[5]), float(r[6])))

        for _,r in df.iloc[2:,18:20].dropna(how="all").iterrows():
            if pd.notna(r[18]):
                cur.execute("INSERT INTO master_project_size(project_size,multiplier) VALUES (?,?)",
                            (str(r[18]), float(r[19])))

        for _,r in df.iloc[2:,21:23].dropna(how="all").iterrows():
            if pd.notna(r[21]):
                cur.execute("INSERT INTO master_workload_status(status,max_load) VALUES (?,?)",
                            (str(r[21]), float(r[22])))

    conn.commit()
    _clear_read_caches()
    conn.close()



@st.cache_data(ttl=300, show_spinner=False)
def get_master_table(name):
    conn = get_conn()
    df = pd.read_sql_query(f"SELECT * FROM master_{name} ORDER BY id", conn)
    conn.close()
    return df


SETUP_CRUD = {
    "Role": ("role", [("role", "Role", "text")]),
    "Project Type": ("project_type", [("project_type", "Project Type", "text")]),
    "Phase": ("phase", [("phase", "Phase", "text"), ("sequence", "Sequence", "int"), ("base_load", "Base Load", "float")]),
    "Project Status": ("project_status", [("status", "Status", "text")]),
    "Meeting Type": ("meeting_type", [("meeting_type", "Meeting Type", "text")]),
    "Meeting Location": ("meeting_location", [("location", "Location", "text")]),
    "Activity Type": ("activity_type", [("activity_type", "Activity Type", "text")]),
    "Priority": ("priority", [("priority", "Priority", "text")]),
    "Project Size": ("project_size", [("project_size", "Project Size", "text"), ("multiplier", "Multiplier", "float")]),
    "Workload Status": ("workload_status", [("status", "Status", "text"), ("max_load", "Max Load", "float")]),
    "Mapping Status": ("mapping_status", [("status", "Status", "text")]),
    "Task Status": ("task_status", [("status", "Status", "text")]),
}


def _crud_values(spec, prefix, row=None):
    values = {}
    for field, label, kind in spec:
        old = row.get(field) if row else None
        if kind == "int":
            values[field] = st.number_input(
                label, min_value=0, step=1,
                value=int(old) if old is not None and pd.notna(old) else 0,
                key=f"{prefix}_{field}",
            )
        elif kind == "float":
            values[field] = st.number_input(
                label, min_value=0.0, step=0.05, format="%.2f",
                value=float(old) if old is not None and pd.notna(old) else 0.0,
                key=f"{prefix}_{field}",
            )
        else:
            values[field] = st.text_input(
                label, value="" if old is None or pd.isna(old) else str(old),
                key=f"{prefix}_{field}",
            )
    return values


def _crud_duplicate(table, fields, values, exclude_id=None):
    conn = get_conn()
    where = " AND ".join(f"{f}=?" for f, _, _ in fields)
    params = [values[f] for f, _, _ in fields]
    sql = f"SELECT id FROM master_{table} WHERE {where}"
    if exclude_id is not None:
        sql += " AND id<>?"
        params.append(exclude_id)
    found = conn.execute(sql, params).fetchone()
    conn.close()
    return found is not None


def _clear_master_caches():
    """Invalidate Setup master caches after Add/Edit/Delete."""
    try:
        master_values.clear()
    except Exception:
        pass
    try:
        get_master_table.clear()
    except Exception:
        pass


def _crud_add(table, fields, values):
    if _crud_duplicate(table, fields, values):
        return False, "Data yang sama sudah ada."
    names = [f for f, _, _ in fields]
    conn = get_conn()
    conn.execute(
        f"INSERT INTO master_{table} ({','.join(names)}) VALUES ({','.join('?' for _ in names)})",
        [values[n] for n in names],
    )
    conn.commit()
    _clear_read_caches()
    conn.close()
    _clear_master_caches()
    return True, "Data berhasil ditambahkan."


def _crud_update(table, fields, row_id, values):
    if _crud_duplicate(table, fields, values, row_id):
        return False, "Data yang sama sudah ada."
    names = [f for f, _, _ in fields]
    conn = get_conn()
    conn.execute(
        f"UPDATE master_{table} SET {','.join(f'{n}=?' for n in names)} WHERE id=?",
        [values[n] for n in names] + [row_id],
    )
    conn.commit()
    _clear_read_caches()
    conn.close()
    _clear_master_caches()
    return True, "Data berhasil diperbarui."


def _crud_delete(table, row_id):
    conn = get_conn()
    conn.execute(f"DELETE FROM master_{table} WHERE id=?", (row_id,))
    conn.commit()
    _clear_read_caches()
    conn.close()
    _clear_master_caches()
    return True, "Data berhasil dihapus."


def setup_page():
    st.markdown('<div class="app-title">Setup Manager</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="app-subtitle">Master reference data used by Input Data and dashboards.</div>',
        unsafe_allow_html=True,
    )

    names = list(SETUP_CRUD.keys())
    tabs = st.tabs(names)

    for tab, label in zip(tabs, names):
        with tab:
            table, fields = SETUP_CRUD[label]
            df = get_master_table(table)

            st.markdown(f"### {label}")
            visible = df.drop(columns=["id"], errors="ignore")
            st.dataframe(visible, use_container_width=True, hide_index=True)

            add_col, edit_col, delete_col = st.columns(3)

            with add_col:
                with st.expander("＋ Add", expanded=False):
                    with st.form(f"add_{table}", clear_on_submit=True):
                        vals = _crud_values(fields, f"add_{table}")
                        if st.form_submit_button("Save New", type="primary", use_container_width=True):
                            if any(k=="text" and not str(vals[f]).strip() for f,_,k in fields):
                                st.error("Field wajib diisi.")
                            else:
                                ok,msg=_crud_add(table,fields,vals)
                                (st.success if ok else st.error)(msg)
                                if ok: st.rerun()

            if not df.empty:
                labels={int(r.id):" • ".join(str(r[f]) for f,_,_ in fields) for _,r in df.iterrows()}

                with edit_col:
                    with st.expander("✎ Edit", expanded=False):
                        rid=st.selectbox("Select data", list(labels), format_func=lambda x: labels[x], key=f"edit_sel_{table}")
                        row=df[df.id==rid].iloc[0].to_dict()
                        with st.form(f"edit_{table}"):
                            vals=_crud_values(fields,f"edit_{table}",row)
                            if st.form_submit_button("Save Changes",type="primary",use_container_width=True):
                                if any(k=="text" and not str(vals[f]).strip() for f,_,k in fields):
                                    st.error("Field wajib diisi.")
                                else:
                                    ok,msg=_crud_update(table,fields,int(rid),vals)
                                    (st.success if ok else st.error)(msg)
                                    if ok: st.rerun()

                with delete_col:
                    with st.expander("🗑 Delete", expanded=False):
                        rid2=st.selectbox("Select data", list(labels), format_func=lambda x: labels[x], key=f"del_sel_{table}")
                        confirm=st.checkbox("Confirm deletion",key=f"del_confirm_{table}")
                        if st.button("Delete Permanently",key=f"del_btn_{table}",disabled=not confirm,use_container_width=True):
                            ok,msg=_crud_delete(table,int(rid2))
                            (st.success if ok else st.error)(msg)
                            if ok: st.rerun()


# ============================================================
# V4 — INPUT DATA
# ============================================================
# Input Data is the operational layer. Dropdowns read directly
# from the Setup master tables, so Setup CRUD changes propagate
# automatically to these forms.

@st.cache_data(ttl=300, show_spinner=False)
def master_values(table, field):
    conn = get_conn()
    try:
        rows = conn.execute(
            f"SELECT {field} FROM master_{table} "
            f"WHERE TRIM(COALESCE({field},''))<>'' ORDER BY id"
        ).fetchall()
    except Exception:
        rows = []
    finally:
        conn.close()
    return [r[0] for r in rows]


@st.cache_data(ttl=20, show_spinner=False, max_entries=128)
def _db_df_cached(sql, params_tuple):
    conn=get_conn()
    try:
        return pd.read_sql_query(sql,conn,params=params_tuple)
    finally:
        conn.close()


def db_df(sql, params=()):
    # Streamlit cache requires stable/hashable arguments.
    return _db_df_cached(sql,tuple(params or ()))


def ensure_v4_input_schema():
    conn=get_conn()
    conn.execute("""
        CREATE TABLE IF NOT EXISTS staff_allocation (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            project_id TEXT NOT NULL,
            phase TEXT NOT NULL,
            staff TEXT NOT NULL,
            role_on_project TEXT,
            notes TEXT,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP,
            UNIQUE(project_id, phase, staff)
        )
    """)
    conn.commit()
    _clear_read_caches()
    conn.close()


def _select_or_empty(label, options, key=None, index=0, help=None):
    opts=list(options)
    if not opts:
        st.warning(f"Belum ada master data untuk {label}. Silakan isi di Setup.")
        return ""
    return st.selectbox(label, opts, index=min(index,len(opts)-1), key=key, help=help)


def _project_status_options():
    """Operational project statuses. Pipeline/Archived are not used for new project records."""
    allowed=["Active","On Hold","Cancelled","Completed"]
    master=master_values("project_status","status")
    lookup={clean(x).lower():x for x in master}
    return [lookup[x.lower()] for x in allowed if x.lower() in lookup]


def _project_options():
    # Active and On Hold projects remain available for operational recording.
    # Cancelled/Completed stay in history but are excluded from new tagging.
    # Workload is still counted only when status == Active (see allocation_calc).
    return db_df("""
        SELECT id,name
        FROM projects
        WHERE LOWER(COALESCE(status,'')) IN ('active','on hold')
        ORDER BY id
    """)


def _staff_options():
    # Active controls dropdown visibility only. Inactive staff remain stored
    # in the database and remain visible/manageable in Input Team.
    return db_df("""
        SELECT id,name,category,primary_role,active
        FROM staff
        WHERE active=1
        ORDER BY name
    """)


def _lead_options():
    """Active Permanent/Intern team members eligible to be Project Lead.
    Freelancers are assigned through Freelance Project Mapping instead."""
    return db_df("""
        SELECT id,name,category,primary_role,active
        FROM staff
        WHERE active=1 AND category <> 'Freelance'
        ORDER BY name
    """)


@st.cache_data(ttl=30, show_spinner=False)
def _project_name(pid):
    if not pid:
        return ""
    conn=get_conn()
    try:
        row=conn.execute("SELECT name FROM projects WHERE id=?",(pid,)).fetchone()
        return row[0] if row else ""
    finally:
        conn.close()


def _team_categories():
    return ["Permanent", "Intern", "Freelance"]


def _add_team():
    st.markdown("### Add Team Member")

    # Category is intentionally OUTSIDE the form so Streamlit reruns immediately
    # when the user changes Permanent / Intern / Freelance.
    cats=_team_categories()
    category=st.selectbox("Category *",cats,key="v4_add_team_category")

    with st.form("v4_add_team", clear_on_submit=True):
        c1,c2=st.columns(2)
        with c1:
            name=st.text_input("Staff Name *")
            roles=master_values("role","role")
            role=_select_or_empty("Primary Role *",roles)
        with c2:
            active=st.checkbox("Active",value=True)

        st.markdown("#### Intern Period")
        d1,d2=st.columns(2)
        with d1:
            intern_start=st.date_input(
                "Intern Start *",
                value=None,
                disabled=(category!="Intern"),
                key="v4_add_intern_start",
            )
        with d2:
            intern_end=st.date_input(
                "Intern End *",
                value=None,
                disabled=(category!="Intern"),
                key="v4_add_intern_end",
            )

        if category!="Intern":
            st.caption("Intern dates are disabled for Permanent and Freelance.")

        save=st.form_submit_button(
            "Save Team Member",type="primary",use_container_width=True
        )

    if save:
        errors=[]
        if not name.strip(): errors.append("Staff Name")
        if not role: errors.append("Primary Role")

        if category=="Intern":
            if intern_start is None: errors.append("Intern Start")
            if intern_end is None: errors.append("Intern End")
            if intern_start and intern_end and intern_end < intern_start:
                st.error("Intern End tidak boleh lebih awal dari Intern Start.")
                return

        if errors:
            st.error("Field wajib diisi: " + ", ".join(errors) + ".")
            return

        if category!="Intern":
            intern_start=None
            intern_end=None

        conn=get_conn()
        try:
            conn.execute("""INSERT INTO staff
                (name,category,primary_role,intern_start,intern_end,active)
                VALUES (?,?,?,?,?,?)""",
                (name.strip(),category,role,
                 intern_start.isoformat() if intern_start else None,
                 intern_end.isoformat() if intern_end else None,
                 1 if active else 0))
            conn.commit()
            _clear_read_caches()
            st.success("Team member berhasil ditambahkan.")
            st.rerun()
        except DB_INTEGRITY_ERRORS:
            st.error("Staff Name sudah ada.")
        finally:
            conn.close()

def _edit_team(df):
    if df.empty:
        st.info("Belum ada team member.")
        return

    st.markdown("### Edit Team Member")
    ids=df["id"].tolist()
    labels={
        int(r.id):f"{r['name']} • {r['category']} • {r['primary_role']}"
        for _,r in df.iterrows()
    }
    rid=st.selectbox(
        "Select Team Member",[None]+ids,index=0,
        format_func=lambda x:"— Select Team Member —" if x is None else labels[int(x)],
        key="v4_team_edit_id"
    )
    if rid is None:
        st.info("Pilih Team Member terlebih dahulu.")
        return
    row=df[df.id==rid].iloc[0]

    # Category is outside the form so the Intern date fields react immediately.
    cats=_team_categories()
    category=st.selectbox(
        "Category *",cats,
        index=cats.index(row["category"]) if row["category"] in cats else 0,
        key=f"v4_edit_team_category_{rid}"
    )

    with st.form("v4_edit_team"):
        c1,c2=st.columns(2)
        with c1:
            name=st.text_input("Staff Name *",value=clean(row["name"]))
            roles=master_values("role","role")
            role=_select_or_empty(
                "Primary Role *",roles,
                index=roles.index(row["primary_role"])
                if row["primary_role"] in roles else 0
            )
        with c2:
            active=st.checkbox("Active",value=bool(row["active"]))

        st.markdown("#### Intern Period")
        d1,d2=st.columns(2)
        with d1:
            sd=safe_date(row["intern_start"])
            intern_start=st.date_input(
                "Intern Start *",
                value=sd,
                disabled=(category!="Intern"),
                key=f"v4_edit_intern_start_{rid}",
            )
        with d2:
            ed=safe_date(row["intern_end"])
            intern_end=st.date_input(
                "Intern End *",
                value=ed,
                disabled=(category!="Intern"),
                key=f"v4_edit_intern_end_{rid}",
            )

        if category!="Intern":
            st.caption("Intern dates are disabled for Permanent and Freelance.")

        save=st.form_submit_button(
            "Save Changes",type="primary",use_container_width=True
        )

    if save:
        errors=[]
        if not name.strip(): errors.append("Staff Name")
        if not role: errors.append("Primary Role")

        if category=="Intern":
            if intern_start is None: errors.append("Intern Start")
            if intern_end is None: errors.append("Intern End")
            if intern_start and intern_end and intern_end < intern_start:
                st.error("Intern End tidak boleh lebih awal dari Intern Start.")
                return

        if errors:
            st.error("Field wajib diisi: " + ", ".join(errors) + ".")
            return

        if category!="Intern":
            intern_start=None
            intern_end=None

        conn=get_conn()
        try:
            conn.execute("""UPDATE staff SET name=?,category=?,primary_role=?,
                intern_start=?,intern_end=?,active=? WHERE id=?""",
                (name.strip(),category,role,
                 intern_start.isoformat() if intern_start else None,
                 intern_end.isoformat() if intern_end else None,
                 1 if active else 0,int(rid)))
            conn.commit()
            _clear_read_caches()
            st.success("Data berhasil diperbarui.")
            st.rerun()
        except DB_INTEGRITY_ERRORS:
            st.error("Staff Name sudah digunakan.")
        finally:
            conn.close()

def _delete_team(df):
    if df.empty:
        st.info("Belum ada team member.")
        return

    st.markdown("### Delete Team Member")
    labels={int(r.id):f"{r['name']} • {r['category']} • {r['primary_role']}" for _,r in df.iterrows()}
    rid=st.selectbox("Select Team Member",[None]+list(labels),index=0,
        format_func=lambda x:"— Select Team Member —" if x is None else labels[x],
        key="v4_team_del_id")
    if rid is None:
        st.info("Pilih Team Member terlebih dahulu.")
        return
    row=df[df.id==rid].iloc[0]
    st.warning(f"Delete **{row['name']}**? Data akan dihapus dari Team.")

    confirm=st.checkbox("Confirm deletion",value=False,key="v4_team_delete_confirm")
    if st.button(
        "Delete Permanently",
        key="v4_team_delete",
        type="secondary",
        use_container_width=True,
        disabled=not confirm,
    ):
        conn=get_conn()
        try:
            # DELETE means physical deletion from the database.
            # Active/Inactive is a separate visibility control for dropdowns.
            conn.execute("DELETE FROM freelance_project_mapping WHERE freelancer=?",(row["name"],))
            conn.execute("DELETE FROM staff WHERE id=?",(int(rid),))
            conn.commit()
            _clear_read_caches()
            st.success("Team member benar-benar dihapus dari database.")
            st.rerun()
        except DB_INTEGRITY_ERRORS:
            conn.rollback()
            st.error("Staff masih digunakan oleh data lain dan tidak dapat dihapus.")
        finally:
            conn.close()



def _normalize_import_header(v):
    """Normalize Excel headers for automatic matching."""
    s=clean(v).lower()
    for ch in [" ", "_", "-", "/", "\\", ".", "(", ")", "*"]:
        s=s.replace(ch,"")
    return s


def _team_import_template():
    """Create a simple Excel template matching the staff table fields."""
    sample=pd.DataFrame([{
        "Staff Name":"Example Staff",
        "Category":"Permanent",
        "Primary Role":"Architect",
        "Intern Start":"",
        "Intern End":"",
        "Active":True,
    }])
    bio=BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as writer:
        sample.to_excel(writer, index=False, sheet_name="Team Import")
    bio.seek(0)
    return bio.getvalue()


def _import_team_excel():
    st.markdown("### Import Team from Excel")
    st.caption(
        "Upload Excel untuk merekam banyak Team Member sekaligus. "
        "Kolom Excel dapat memiliki nama berbeda; pada langkah berikutnya "
        "kolom akan dipetakan ke field Team."
    )

    st.download_button(
        "Download Excel Template",
        data=_team_import_template(),
        file_name="Team_Import_Template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="v4c_team_template",
    )

    uploaded=st.file_uploader(
        "Upload Excel",
        type=["xlsx","xlsm"],
        key="v4c_team_excel",
    )

    # After a successful save, stop processing the still-selected uploader.
    # Streamlit reruns the script, but the uploaded file remains in session;
    # without this guard it would be validated a second time and appear as
    # duplicate data.
    imported_count = st.session_state.pop("v4c_team_import_success", None)
    if imported_count is not None:
        imported_file = st.session_state.pop("v4c_team_import_file", "")
        st.success(
            f"Import berhasil. {imported_count} Team Member "
            f"berhasil direkam ke database."
        )
        if imported_file:
            st.caption(f"File: {imported_file}")
        return

    if uploaded is None:
        return

    try:
        xls=pd.ExcelFile(uploaded)
        sheet=st.selectbox("Sheet",xls.sheet_names,key="v4c_team_import_sheet")
        raw=pd.read_excel(uploaded,sheet_name=sheet,engine="openpyxl")
    except Exception as exc:
        st.error(f"Excel tidak dapat dibaca: {exc}")
        return

    if raw.empty:
        st.warning("Sheet tidak memiliki data.")
        return

    raw=raw.dropna(how="all").copy()
    if raw.empty:
        st.warning("Tidak ada baris data.")
        return

    st.markdown("#### 1. Mapping Kolom Excel")
    source_cols=[str(c) for c in raw.columns]
    target_fields=[
        ("name","Staff Name *"),
        ("category","Category *"),
        ("primary_role","Primary Role *"),
        ("intern_start","Intern Start"),
        ("intern_end","Intern End"),
        ("active","Active"),
    ]

    norm={_normalize_import_header(c):c for c in source_cols}
    aliases={
        "name":["staffname","name","nama","staff"],
        "category":["category","kategori","employmentcategory","type"],
        "primary_role":["primaryrole","role","jabatan","position"],
        "intern_start":["internstart","internstartdate","startintern"],
        "intern_end":["internend","internenddate","endintern"],
        "active":["active","isactive","statusactive"],
    }

    mapping={}
    options=["— Not mapped —"]+source_cols
    for target,label in target_fields:
        guess="— Not mapped —"
        for a in aliases[target]:
            if a in norm:
                guess=norm[a]
                break
        default=options.index(guess) if guess in options else 0
        mapping[target]=st.selectbox(
            label,
            options,
            index=default,
            key=f"v4c_team_map_{target}",
        )

    required_missing=[
        label for target,label in target_fields[:3]
        if mapping[target]=="— Not mapped —"
    ]
    if required_missing:
        st.warning("Field wajib belum dipetakan: "+", ".join(required_missing))
        return

    st.markdown("#### 2. Preview Hasil Mapping")
    mapped=pd.DataFrame(index=raw.index)
    for target,label in target_fields:
        src_col=mapping[target]
        mapped[label.replace(" *","")]=(
            raw[src_col] if src_col!="— Not mapped —"
            else None
        )

    st.dataframe(mapped.head(20),use_container_width=True,hide_index=True)
    st.caption(f"{len(mapped)} baris siap divalidasi.")

    # Validation + normalization happens before anything is written to DB.
    categories=_team_categories()
    roles=master_values("role","role")
    role_lookup={clean(x).lower():x for x in roles}
    cat_lookup={clean(x).lower():x for x in categories}

    valid_rows=[]
    errors=[]
    existing_names=set(
        clean(x).lower()
        for x in db_df("SELECT name FROM staff")["name"].tolist()
    )

    for ix,row in raw.iterrows():
        name=clean(row[mapping["name"]])
        cat_raw=clean(row[mapping["category"]])
        role_raw=clean(row[mapping["primary_role"]])

        if not name:
            errors.append(f"Baris Excel {ix+2}: Staff Name kosong.")
            continue
        if not cat_raw:
            errors.append(f"Baris Excel {ix+2}: Category kosong.")
            continue
        if cat_raw.lower() not in cat_lookup:
            errors.append(
                f"Baris Excel {ix+2}: Category '{cat_raw}' tidak ada di Setup."
            )
            continue
        if not role_raw:
            errors.append(f"Baris Excel {ix+2}: Primary Role kosong.")
            continue
        if role_raw.lower() not in role_lookup:
            errors.append(
                f"Baris Excel {ix+2}: Primary Role '{role_raw}' tidak ada di Setup."
            )
            continue

        category=cat_lookup[cat_raw.lower()]
        role=role_lookup[role_raw.lower()]

        start=None
        end=None
        if mapping["intern_start"]!="— Not mapped —":
            start=to_iso_date(row[mapping["intern_start"]])
        if mapping["intern_end"]!="— Not mapped —":
            end=to_iso_date(row[mapping["intern_end"]])

        if category=="Intern":
            if not start or not end:
                errors.append(
                    f"Baris Excel {ix+2}: Intern wajib memiliki Intern Start dan Intern End."
                )
                continue
            if end<start:
                errors.append(
                    f"Baris Excel {ix+2}: Intern End lebih awal dari Intern Start."
                )
                continue
        else:
            start=None
            end=None

        active=True
        if mapping["active"]!="— Not mapped —":
            v=row[mapping["active"]]
            if pd.isna(v):
                active=True
            elif isinstance(v,bool):
                active=v
            else:
                active=str(v).strip().lower() not in {
                    "0","false","no","n","inactive","nonaktif","tidak"
                }

        duplicate_in_file=any(
            clean(r["name"]).lower()==name.lower() for r in valid_rows
        )
        if name.lower() in existing_names or duplicate_in_file:
            errors.append(
                f"Baris Excel {ix+2}: Staff Name '{name}' sudah ada."
            )
            continue

        valid_rows.append({
            "name":name,
            "category":category,
            "primary_role":role,
            "intern_start":start,
            "intern_end":end,
            "active":1 if active else 0,
        })

    if errors:
        st.error(f"Ditemukan {len(errors)} masalah. Tidak ada data yang direkam.")
        st.dataframe(
            pd.DataFrame({"Validation Error":errors}),
            use_container_width=True,
            hide_index=True,
        )
        return

    st.success(f"{len(valid_rows)} baris lolos validasi dan siap direkam.")

    if st.button(
        "Save Imported Team Data",
        type="primary",
        use_container_width=True,
        key="v4c_team_import_save",
    ):
        conn=get_conn()
        try:
            for r in valid_rows:
                conn.execute(
                    """INSERT INTO staff
                    (name,category,primary_role,intern_start,intern_end,active)
                    VALUES (?,?,?,?,?,?)""",
                    (
                        r["name"],r["category"],r["primary_role"],
                        r["intern_start"],r["intern_end"],r["active"]
                    ),
                )
            conn.commit()
            _clear_read_caches()
            st.session_state["v4c_team_import_success"] = len(valid_rows)
            st.session_state["v4c_team_import_file"] = uploaded.name
            st.rerun()
        except DB_INTEGRITY_ERRORS as exc:
            conn.rollback()
            st.error(f"Import dibatalkan karena konflik database: {exc}")
        finally:
            conn.close()


def input_team_page():
    st.markdown('<div class="app-title">Input Team</div>',unsafe_allow_html=True)
    st.markdown("Manage team members used throughout the Control Board.")
    st.caption("Active = muncul pada dropdown di modul lain. Inactive = tetap tersimpan di database, tetapi tidak muncul pada dropdown. Delete = menghapus permanen dari database.")
    df=db_df("SELECT id,name,category,primary_role,intern_start,intern_end,active FROM staff ORDER BY name")
    st.dataframe(df.drop(columns=["id"]),use_container_width=True,hide_index=True)
    a,e,d,i=st.tabs(["＋ Add","✎ Edit","🗑 Delete","⇧ Import Excel"])
    with a: _add_team()
    with e: _edit_team(df)
    with d: _delete_team(df)
    with i: _import_team_excel()


def _project_duration(start,finish):
    if not start or not finish: return None
    return (finish.year-start.year)*12+finish.month-start.month+1



def _next_project_id(conn, reserved=None):
    """Generate the next system project ID (P001, P002, ...)."""
    reserved = {str(x).upper() for x in (reserved or [])}
    rows = conn.execute("SELECT id FROM projects").fetchall()
    nums = []
    for row in rows:
        m = re.fullmatch(r"P(\d+)", str(row[0]).strip(), re.IGNORECASE)
        if m:
            nums.append(int(m.group(1)))
    n = max(nums, default=0) + 1
    while f"P{n:03d}".upper() in reserved:
        n += 1
    return f"P{n:03d}"

def _add_project():
    st.markdown("### Add Project")
    with st.form("v4_add_project",clear_on_submit=True):
        st.caption("Project ID akan digenerate otomatis oleh sistem saat data disimpan.")
        c1,c2,c3=st.columns(3)
        with c1:
            name=st.text_input("Project Name *")
            ptypes=master_values("project_type","project_type")
            ptype=_select_or_empty("Project Type",ptypes)
        with c2:
            start=st.date_input("Start Date",value=None)
            finish=st.date_input("Target Finish",value=None)
            sizes=master_values("project_size","project_size")
            size=_select_or_empty("Project Size",sizes)
        with c3:
            statuses=_project_status_options()
            status=_select_or_empty("Project Status *",statuses)
            staff=_lead_options()
            leads=staff["name"].tolist() if not staff.empty else []
            lead=_select_or_empty("Lead (Permanent / Intern)",leads)
            duration=_project_duration(start,finish)
            st.caption(f"Duration: {duration or '—'} month(s)")
        save=st.form_submit_button("Save Project",type="primary",use_container_width=True)

    if save:
        errors=[]
        if not name.strip(): errors.append("Project Name")
        if not status: errors.append("Project Status")
        if errors:
            st.error("Field wajib diisi: " + ", ".join(errors) + ".")
            return
        if start and finish and finish<start:
            st.error("Target Finish tidak boleh lebih awal dari Start Date."); return

        conn=get_conn()
        try:
            pid=_next_project_id(conn)
            conn.execute("""INSERT INTO projects
                (id,name,project_type,start_date,target_finish,duration_months,status,lead,project_size)
                VALUES (?,?,?,?,?,?,?,?,?)""",
                (pid,name.strip(),ptype,
                 start.isoformat() if start else None,
                 finish.isoformat() if finish else None,
                 _project_duration(start,finish),status,lead,size))
            conn.commit()
            _clear_read_caches()
            st.success(f"Project {pid} berhasil ditambahkan.")
            st.rerun()
        except DB_INTEGRITY_ERRORS as exc:
            conn.rollback()
            st.error(f"Project gagal ditambahkan: {exc}")
        finally:
            conn.close()


def _edit_project(df):
    if df.empty: return
    labels={str(r.id):f"{r.id} • {r['name']}" for _,r in df.iterrows()}
    pid=st.selectbox("Select Project",list(labels),format_func=lambda x:labels[x],key="v4_proj_edit_id")
    row=df[df.id==pid].iloc[0]
    with st.form("v4_edit_project"):
        c1,c2,c3=st.columns(3)
        with c1:
            st.text_input("Project ID",value=str(row["id"]),disabled=True)
            name=st.text_input("Project Name *",value=clean(row["name"]))
            ptypes=master_values("project_type","project_type")
            ptype=_select_or_empty("Project Type",ptypes,index=ptypes.index(row["project_type"]) if row["project_type"] in ptypes else 0)
        with c2:
            start=safe_date(row["start_date"], None)
            finish=safe_date(row["target_finish"], None)
            start=st.date_input("Start Date",value=start)
            finish=st.date_input("Target Finish",value=finish)
            sizes=master_values("project_size","project_size")
            size=_select_or_empty("Project Size",sizes,index=sizes.index(row["project_size"]) if row["project_size"] in sizes else 0)
        with c3:
            statuses=_project_status_options()
            status=_select_or_empty("Project Status *",statuses,index=statuses.index(row["status"]) if row["status"] in statuses else 0)
            sdf=_staff_options(); leads=sdf["name"].tolist() if not sdf.empty else []
            lead=_select_or_empty("Lead (Permanent / Intern)",leads,index=leads.index(row["lead"]) if row["lead"] in leads else 0)
            st.caption(f"Duration: {_project_duration(start,finish) or '—'} month(s)")
        save=st.form_submit_button("Save Changes",type="primary",use_container_width=True)
    if save:
        errors=[]
        if not name.strip(): errors.append("Project Name")
        if not status: errors.append("Project Status")
        if errors:
            st.error("Field wajib diisi: " + ", ".join(errors) + ".")
            return
        if start and finish and finish<start:
            st.error("Target Finish tidak boleh lebih awal dari Start Date."); return
        conn=get_conn()
        try:
            conn.execute("""UPDATE projects SET name=?,project_type=?,start_date=?,target_finish=?,
                duration_months=?,status=?,lead=?,project_size=? WHERE id=?""",
                (name.strip(),ptype,
                 start.isoformat() if start else None,
                 finish.isoformat() if finish else None,
                 _project_duration(start,finish),status,lead,size,pid))
            conn.commit()
            _clear_read_caches()
            st.success("Project berhasil diperbarui.")
            st.rerun()
        finally:
            conn.close()


def _delete_project(df):
    if df.empty:return
    labels={str(r.id):f"{r.id} • {r['name']}" for _,r in df.iterrows()}
    pid=st.selectbox("Select Project to Delete",list(labels),format_func=lambda x:labels[x],key="v4_proj_del_id")
    if st.button("Delete Permanently",key="v4_proj_delete",type="secondary"):
        conn=get_conn()
        # Do not silently remove linked activities. Mark project archived instead.
        archived=master_values("project_status","status")
        status="Archived" if "Archived" in archived else (archived[-1] if archived else "Archived")
        conn.execute("UPDATE projects SET status=? WHERE id=?",(status,pid))
        conn.commit(); _clear_read_caches(); conn.close(); st.success("Project diarsipkan agar histori aktivitas tetap aman."); st.rerun()



def _project_import_template():
    """Create an Excel template matching the project input fields."""
    sample=pd.DataFrame([{
        "Project Name":"Example Project",
        "Project Type":"",
        "Start Date":"",
        "Target Finish":"",
        "Project Status":"Active",
        "Lead":"",
        "Project Size":"",
    }])
    bio=BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as writer:
        sample.to_excel(writer, index=False, sheet_name="Project Import")
    bio.seek(0)
    return bio.getvalue()


def _import_project_excel():
    st.markdown("### Import Project from Excel")
    st.caption(
        "Upload Excel untuk merekam banyak Project sekaligus. "
        "Project ID tidak perlu diisi karena akan digenerate otomatis oleh sistem."
    )

    st.download_button(
        "Download Excel Template",
        data=_project_import_template(),
        file_name="Project_Import_Template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="v5_project_template",
    )

    uploaded=st.file_uploader(
        "Upload Excel",
        type=["xlsx","xlsm"],
        key="v5_project_excel",
    )

    imported_count=st.session_state.pop("v5_project_import_success",None)
    if imported_count is not None:
        imported_file=st.session_state.pop("v5_project_import_file","")
        st.success(f"Import berhasil. {imported_count} Project berhasil direkam ke database.")
        if imported_file:
            st.caption(f"File: {imported_file}")
        return

    if uploaded is None:
        return

    try:
        xls=pd.ExcelFile(uploaded)
        sheet=st.selectbox("Sheet",xls.sheet_names,key="v5_project_import_sheet")
        raw=pd.read_excel(uploaded,sheet_name=sheet,engine="openpyxl")
    except Exception as exc:
        st.error(f"Excel tidak dapat dibaca: {exc}")
        return

    if raw.empty:
        st.warning("Sheet tidak memiliki data.")
        return
    raw=raw.dropna(how="all").copy()
    if raw.empty:
        st.warning("Tidak ada baris data.")
        return

    st.markdown("#### 1. Mapping Kolom Excel")
    source_cols=[str(c) for c in raw.columns]
    target_fields=[
        ("name","Project Name *"),
        ("project_type","Project Type"),
        ("start_date","Start Date"),
        ("target_finish","Target Finish"),
        ("status","Project Status *"),
        ("lead","Lead"),
        ("project_size","Project Size"),
    ]

    norm={_normalize_import_header(c):c for c in source_cols}
    aliases={
        "name":["projectname","name","namaproject","namaproyek","project"],
        "project_type":["projecttype","type","tipeproject","tipeproyek"],
        "start_date":["startdate","start","tanggalmulai","projectstart"],
        "target_finish":["targetfinish","finishdate","targetdate","enddate","tanggalselesai","projectfinish"],
        "status":["projectstatus","status","statusproject","statusproyek"],
        "lead":["lead","projectlead","pic","leader"],
        "project_size":["projectsize","size","ukuranproject","ukuranproyek"],
    }

    mapping={}
    options=["— Not mapped —"]+source_cols
    for target,label in target_fields:
        guess="— Not mapped —"
        for a in aliases[target]:
            if a in norm:
                guess=norm[a]
                break
        default=options.index(guess) if guess in options else 0
        mapping[target]=st.selectbox(label,options,index=default,key=f"v5_project_map_{target}")

    required_targets={"name","status"}
    missing=[label for target,label in target_fields if target in required_targets and mapping[target]=="— Not mapped —"]
    if missing:
        st.warning("Field wajib belum dipetakan: "+", ".join(missing))
        return

    st.markdown("#### 2. Preview Hasil Mapping")
    preview=pd.DataFrame(index=raw.index)
    for target,label in target_fields:
        src_col=mapping[target]
        preview[label.replace(" *","")]=raw[src_col] if src_col!="— Not mapped —" else None
    st.dataframe(preview.head(20),use_container_width=True,hide_index=True)
    st.caption(f"{len(preview)} baris siap divalidasi.")

    project_types=master_values("project_type","project_type")
    statuses=_project_status_options()
    sizes=master_values("project_size","project_size")
    leads_df=_lead_options()
    leads=leads_df["name"].tolist() if not leads_df.empty else []

    type_lookup={clean(x).lower():x for x in project_types}
    status_lookup={clean(x).lower():x for x in statuses}
    size_lookup={clean(x).lower():x for x in sizes}
    lead_lookup={clean(x).lower():x for x in leads}

    valid_rows=[]
    errors=[]

    def cell_text(target,row):
        src_col=mapping[target]
        return "" if src_col=="— Not mapped —" else clean(row[src_col])

    for ix,row in raw.iterrows():
        name=cell_text("name",row)
        ptype_raw=cell_text("project_type",row)
        status_raw=cell_text("status",row)
        lead_raw=cell_text("lead",row)
        size_raw=cell_text("project_size",row)

        if not name:
            errors.append(f"Baris Excel {ix+2}: Project Name kosong.")
            continue

        ptype=type_lookup.get(ptype_raw.lower()) if ptype_raw else None
        if ptype_raw and not ptype:
            errors.append(f"Baris Excel {ix+2}: Project Type '{ptype_raw}' tidak ada di Setup.")
            continue

        if not status_raw:
            errors.append(f"Baris Excel {ix+2}: Project Status kosong.")
            continue
        status=status_lookup.get(status_raw.lower())
        if not status:
            errors.append(f"Baris Excel {ix+2}: Project Status '{status_raw}' tidak valid.")
            continue

        size=size_lookup.get(size_raw.lower()) if size_raw else None
        if size_raw and not size:
            errors.append(f"Baris Excel {ix+2}: Project Size '{size_raw}' tidak ada di Setup.")
            continue

        lead=lead_lookup.get(lead_raw.lower()) if lead_raw else None
        if lead_raw and not lead:
            errors.append(f"Baris Excel {ix+2}: Lead '{lead_raw}' tidak tersedia pada Team Active.")
            continue

        start_raw=cell_text("start_date",row)
        finish_raw=cell_text("target_finish",row)
        start=to_iso_date(row[mapping["start_date"]]) if mapping["start_date"]!="— Not mapped —" and start_raw else None
        finish=to_iso_date(row[mapping["target_finish"]]) if mapping["target_finish"]!="— Not mapped —" and finish_raw else None

        if start_raw and not start:
            errors.append(f"Baris Excel {ix+2}: Start Date tidak valid.")
            continue
        if finish_raw and not finish:
            errors.append(f"Baris Excel {ix+2}: Target Finish tidak valid.")
            continue
        if start and finish and finish<start:
            errors.append(f"Baris Excel {ix+2}: Target Finish lebih awal dari Start Date.")
            continue

        valid_rows.append({
            "name":name,
            "project_type":ptype,
            "start_date":start,
            "target_finish":finish,
            "duration_months":_project_duration(parse_date(start),parse_date(finish)) if start and finish else None,
            "status":status,
            "lead":lead,
            "project_size":size,
        })

    if errors:
        st.error(f"Ditemukan {len(errors)} masalah. Tidak ada data yang direkam.")
        st.dataframe(pd.DataFrame({"Validation Error":errors}),use_container_width=True,hide_index=True)
        return

    # Show the system-generated IDs before saving.
    conn=get_conn()
    reserved=[]
    for r in valid_rows:
        pid=_next_project_id(conn,reserved)
        reserved.append(pid)
        r["id"]=pid
    conn.close()

    import_preview=pd.DataFrame([{
        "Project ID":r["id"],
        "Project Name":r["name"],
        "Project Type":r["project_type"],
        "Start Date":r["start_date"],
        "Target Finish":r["target_finish"],
        "Project Status":r["status"],
        "Lead":r["lead"],
        "Project Size":r["project_size"],
    } for r in valid_rows])
    st.success(f"{len(valid_rows)} baris lolos validasi dan siap direkam.")
    st.dataframe(import_preview,use_container_width=True,hide_index=True)

    if st.button("Save Imported Project Data",type="primary",use_container_width=True,key="v5_project_import_save"):
        conn=get_conn()
        try:
            # Re-generate IDs at commit time so concurrent/previous inserts cannot collide.
            for r in valid_rows:
                r["id"]=_next_project_id(conn)
                conn.execute("""INSERT INTO projects
                    (id,name,project_type,start_date,target_finish,duration_months,status,lead,project_size)
                    VALUES (?,?,?,?,?,?,?,?,?)""",
                    (r["id"],r["name"],r["project_type"],r["start_date"],r["target_finish"],
                     r["duration_months"],r["status"],r["lead"],r["project_size"]))
            conn.commit()
            _clear_read_caches()
            st.session_state["v5_project_import_success"]=len(valid_rows)
            st.session_state["v5_project_import_file"]=uploaded.name
            st.rerun()
        except DB_INTEGRITY_ERRORS as exc:
            conn.rollback()
            st.error(f"Import dibatalkan karena konflik database: {exc}")
        finally:
            conn.close()


def input_project_page():
    st.markdown('<div class="app-title">Input Project</div>',unsafe_allow_html=True)
    st.markdown("Manage projects used by allocation, activities and dashboards.")
    df=db_df("""SELECT id,name,project_type,start_date,target_finish,duration_months,status,lead,project_size
                FROM projects ORDER BY id""")
    st.dataframe(df,use_container_width=True,hide_index=True)
    a,e,d,i=st.tabs(["＋ Add","✎ Edit","🗑 Delete","⇧ Import Excel"])
    with a:_add_project()
    with e:_edit_project(df)
    with d:_delete_project(df)
    with i:_import_project_excel()


def allocation_calc(project_id, phase, staff):
    conn=get_conn()
    prow=conn.execute("SELECT project_size,status FROM projects WHERE id=?",(project_id,)).fetchone()
    srow=conn.execute("SELECT multiplier FROM master_project_size WHERE project_size=?",(prow[0],)).fetchone() if prow else None
    prow2=conn.execute("SELECT base_load FROM master_phase WHERE phase=?",(phase,)).fetchone()
    mult=float(srow[0]) if srow else 1.0
    weight=float(prow2[0]) if prow2 else 0.0
    # Assignment remains in DB for history/re-activation, but only Active
    # projects contribute to current workload.
    total=mult*weight if prow and clean(prow[1]).lower()=="active" else 0.0
    count=conn.execute("SELECT COUNT(*) FROM staff_allocation WHERE project_id=? AND phase=?",(project_id,phase)).fetchone()[0]
    conn.close()
    return mult,weight,total,count+1,total/(count+1) if count+1 else 0


def _allocation_df():
    return db_df("""
        SELECT a.id,a.project_id,p.name AS project_name,a.phase,a.staff,
               s.category AS staff_category,s.primary_role,a.role_on_project,
               p.project_size,p.status AS project_status
        FROM staff_allocation a
        LEFT JOIN projects p ON p.id=a.project_id
        LEFT JOIN staff s ON s.name=a.staff
        ORDER BY a.project_id,a.phase,a.staff
    """)



def _allocation_import_template():
    """Excel template matching the Staff Allocation input fields."""
    sample=pd.DataFrame([{
        "Project ID":"P001",
        "Phase":"Concept",
        "Staff":"Example Staff",
        "Role on Project":"Architect",
        "Notes":"",
    }])
    bio=BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as writer:
        sample.to_excel(writer,index=False,sheet_name="Staff Allocation Import")
    bio.seek(0)
    return bio.getvalue()


def _import_allocation_excel():
    st.markdown("### Import Staff Allocation from Excel")
    st.caption(
        "Upload Excel untuk merekam banyak Staff Allocation sekaligus. "
        "Kolom Excel dapat memiliki nama berbeda dan akan dipetakan ke field Staff Allocation."
    )

    st.download_button(
        "Download Excel Template",
        data=_allocation_import_template(),
        file_name="Staff_Allocation_Import_Template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="v6_allocation_template",
    )

    uploaded=st.file_uploader(
        "Upload Excel",
        type=["xlsx","xlsm"],
        key="v6_allocation_excel",
    )

    imported_count=st.session_state.pop("v6_allocation_import_success",None)
    if imported_count is not None:
        imported_file=st.session_state.pop("v6_allocation_import_file","")
        st.success(
            f"Import berhasil. {imported_count} Staff Allocation "
            f"berhasil direkam ke database."
        )
        if imported_file:
            st.caption(f"File: {imported_file}")
        return

    if uploaded is None:
        return

    try:
        xls=pd.ExcelFile(uploaded)
        sheet=st.selectbox(
            "Sheet",xls.sheet_names,key="v6_allocation_import_sheet"
        )
        raw=pd.read_excel(uploaded,sheet_name=sheet,engine="openpyxl")
    except Exception as exc:
        st.error(f"Excel tidak dapat dibaca: {exc}")
        return

    if raw.empty:
        st.warning("Sheet tidak memiliki data.")
        return
    raw=raw.dropna(how="all").copy()
    if raw.empty:
        st.warning("Tidak ada baris data.")
        return

    st.markdown("#### 1. Mapping Kolom Excel")
    source_cols=[str(c) for c in raw.columns]
    target_fields=[
        ("project_id","Project ID *"),
        ("phase","Phase *"),
        ("staff","Staff *"),
        ("role_on_project","Role on Project"),
        ("notes","Notes"),
    ]

    norm={_normalize_import_header(c):c for c in source_cols}
    aliases={
        "project_id":["projectid","project","idproject","projectcode","kodeproject"],
        "phase":["phase","tahap","projectphase"],
        "staff":["staff","staffname","name","team","teammember","member"],
        "role_on_project":["roleonproject","projectrole","role","peran"],
        "notes":["notes","note","catatan","remark","remarks"],
    }

    mapping={}
    options=["— Not mapped —"]+source_cols
    for target,label in target_fields:
        guess="— Not mapped —"
        for a in aliases[target]:
            if a in norm:
                guess=norm[a]
                break
        default=options.index(guess) if guess in options else 0
        mapping[target]=st.selectbox(
            label,options,index=default,key=f"v6_alloc_map_{target}"
        )

    required_missing=[
        label for target,label in target_fields[:3]
        if mapping[target]=="— Not mapped —"
    ]
    if required_missing:
        st.warning("Field wajib belum dipetakan: "+", ".join(required_missing))
        return

    st.markdown("#### 2. Preview Hasil Mapping")
    mapped=pd.DataFrame(index=raw.index)
    for target,label in target_fields:
        src_col=mapping[target]
        mapped[label.replace(" *","")]=(
            raw[src_col] if src_col!="— Not mapped —" else None
        )
    st.dataframe(mapped.head(20),use_container_width=True,hide_index=True)
    st.caption(f"{len(mapped)} baris siap divalidasi.")

    projects=_project_options()
    phases=master_values("phase","phase")
    roles=master_values("role","role")
    staff_df=_staff_options()

    project_lookup={
        clean(r["id"]).lower():clean(r["id"])
        for _,r in projects.iterrows()
    }
    phase_lookup={clean(x).lower():x for x in phases}
    role_lookup={clean(x).lower():x for x in roles}
    staff_lookup={clean(r["name"]).lower():clean(r["name"])
                  for _,r in staff_df.iterrows()}

    existing=set(
        (clean(r["project_id"]).lower(),
         clean(r["phase"]).lower(),
         clean(r["staff"]).lower())
        for _,r in _allocation_df().iterrows()
    )

    valid_rows=[]
    errors=[]

    def cell_text(target,row):
        src_col=mapping[target]
        return "" if src_col=="— Not mapped —" else clean(row[src_col])

    for ix,row in raw.iterrows():
        pid_raw=cell_text("project_id",row)
        phase_raw=cell_text("phase",row)
        staff_raw=cell_text("staff",row)
        role_raw=cell_text("role_on_project",row)
        notes=cell_text("notes",row)

        if not pid_raw:
            errors.append(f"Baris Excel {ix+2}: Project ID kosong.")
            continue
        pid=project_lookup.get(pid_raw.lower())
        if not pid:
            errors.append(
                f"Baris Excel {ix+2}: Project ID '{pid_raw}' tidak tersedia "
                f"sebagai Project Active."
            )
            continue

        if not phase_raw:
            errors.append(f"Baris Excel {ix+2}: Phase kosong.")
            continue
        phase=phase_lookup.get(phase_raw.lower())
        if not phase:
            errors.append(
                f"Baris Excel {ix+2}: Phase '{phase_raw}' tidak ada di Setup."
            )
            continue

        if not staff_raw:
            errors.append(f"Baris Excel {ix+2}: Staff kosong.")
            continue
        staff=staff_lookup.get(staff_raw.lower())
        if not staff:
            errors.append(
                f"Baris Excel {ix+2}: Staff '{staff_raw}' tidak tersedia "
                f"pada Team Active."
            )
            continue

        role=None
        if role_raw:
            role=role_lookup.get(role_raw.lower())
            if not role:
                errors.append(
                    f"Baris Excel {ix+2}: Role on Project '{role_raw}' "
                    f"tidak ada di Setup."
                )
                continue

        key=(pid.lower(),phase.lower(),staff.lower())
        if key in existing or any(
            (r["project_id"].lower(),r["phase"].lower(),r["staff"].lower())==key
            for r in valid_rows
        ):
            errors.append(
                f"Baris Excel {ix+2}: Staff '{staff}' sudah dialokasikan "
                f"pada project '{pid}' dan phase '{phase}'."
            )
            continue

        valid_rows.append({
            "project_id":pid,
            "phase":phase,
            "staff":staff,
            "role_on_project":role,
            "notes":notes,
        })

    if errors:
        st.error(f"Ditemukan {len(errors)} masalah. Tidak ada data yang direkam.")
        st.dataframe(
            pd.DataFrame({"Validation Error":errors}),
            use_container_width=True,
            hide_index=True,
        )
        return

    import_preview=pd.DataFrame([{
        "Project ID":r["project_id"],
        "Project Name":_project_name(r["project_id"]),
        "Phase":r["phase"],
        "Staff":r["staff"],
        "Role on Project":r["role_on_project"],
        "Notes":r["notes"],
    } for r in valid_rows])

    st.success(f"{len(valid_rows)} baris lolos validasi dan siap direkam.")
    st.dataframe(
        import_preview,use_container_width=True,hide_index=True
    )

    if st.button(
        "Save Imported Staff Allocation Data",
        type="primary",
        use_container_width=True,
        key="v6_allocation_import_save",
    ):
        conn=get_conn()
        try:
            for r in valid_rows:
                conn.execute(
                    """INSERT INTO staff_allocation
                       (project_id,phase,staff,role_on_project,notes)
                       VALUES (?,?,?,?,?)""",
                    (r["project_id"],r["phase"],r["staff"],
                     r["role_on_project"],r["notes"])
                )
            conn.commit()
            _clear_read_caches()
            st.session_state["v6_allocation_import_success"]=len(valid_rows)
            st.session_state["v6_allocation_import_file"]=uploaded.name
            st.rerun()
        except DB_INTEGRITY_ERRORS as exc:
            conn.rollback()
            st.error(f"Import dibatalkan karena konflik database: {exc}")
        finally:
            conn.close()


def _add_allocation():
    projects=_project_options()
    staff=_staff_options()
    if projects.empty: st.warning("Buat Project terlebih dahulu."); return
    if staff.empty: st.warning("Isi Team terlebih dahulu."); return
    pids=projects["id"].tolist()
    people=staff["name"].tolist()
    phases=master_values("phase","phase")
    with st.form("v4_add_allocation",clear_on_submit=True):
        c1,c2=st.columns(2)
        with c1:
            pid=st.selectbox("Project ID *",pids)
            phase=_select_or_empty("Phase *",phases)
        with c2:
            person=st.selectbox("Staff *",people)
            roles=master_values("role","role")
            role=_select_or_empty("Role on Project",roles)
        notes=st.text_area("Notes")
        if pid and phase:
            mult,weight,total,count,individual=allocation_calc(pid,phase,person)
            st.caption(f"Size Multiplier {mult:g}  •  Phase Weight {weight:g}  •  Total Phase Load {total:g}  •  Assigned Staff {count}  •  Individual Load {individual:g}")
        save=st.form_submit_button("Save Allocation",type="primary",use_container_width=True)
    if save:
        conn=get_conn()
        try:
            conn.execute("""INSERT INTO staff_allocation(project_id,phase,staff,role_on_project,notes)
                            VALUES (?,?,?,?,?)""",(pid,phase,person,role,notes))
            conn.commit(); _clear_read_caches(); st.success("Staff allocation berhasil ditambahkan."); st.rerun()
        except DB_INTEGRITY_ERRORS:
            st.error("Staff tersebut sudah dialokasikan pada project dan phase yang sama.")
        finally: conn.close()


def _edit_allocation(df):
    if df.empty:return
    labels={int(r.id):f"{r.project_id} • {r.phase} • {r.staff}" for _,r in df.iterrows()}
    rid=st.selectbox("Select Allocation",list(labels),format_func=lambda x:labels[x],key="v4_alloc_edit_id")
    row=df[df.id==rid].iloc[0]
    projects=_project_options(); staff=_staff_options(); phases=master_values("phase","phase"); roles=master_values("role","role")
    with st.form("v4_edit_allocation"):
        pid=st.selectbox("Project ID",projects["id"].tolist(),index=projects["id"].tolist().index(row["project_id"]))
        phase=_select_or_empty("Phase",phases,index=phases.index(row["phase"]) if row["phase"] in phases else 0)
        people=staff["name"].tolist()
        person=st.selectbox("Staff",people,index=people.index(row["staff"]) if row["staff"] in people else 0)
        role=_select_or_empty("Role on Project",roles,index=roles.index(row["role_on_project"]) if row["role_on_project"] in roles else 0)
        notes=st.text_area("Notes",value=clean(row.get("notes","")))
        save=st.form_submit_button("Save Changes",type="primary",use_container_width=True)
    if save:
        conn=get_conn()
        try:
            conn.execute("""UPDATE staff_allocation SET project_id=?,phase=?,staff=?,role_on_project=?,notes=? WHERE id=?""",
                         (pid,phase,person,role,notes,int(rid)))
            conn.commit(); _clear_read_caches(); st.success("Allocation diperbarui."); st.rerun()
        except DB_INTEGRITY_ERRORS: st.error("Allocation yang sama sudah ada.")
        finally: conn.close()


def _delete_allocation(df):
    if df.empty:return
    labels={int(r.id):f"{r.project_id} • {r.phase} • {r.staff}" for _,r in df.iterrows()}
    rid=st.selectbox("Select Allocation to Delete",list(labels),format_func=lambda x:labels[x],key="v4_alloc_del_id")
    if st.button("Delete Permanently",key="v4_alloc_delete",type="secondary"):
        conn=get_conn(); conn.execute("DELETE FROM staff_allocation WHERE id=?",(int(rid),)); conn.commit(); _clear_read_caches(); conn.close()
        st.success("Allocation dihapus."); st.rerun()


def input_allocation_page():
    st.markdown('<div class="app-title">Staff Allocation</div>',unsafe_allow_html=True)
    st.markdown("Assign staff to a project phase. Load parameters are derived from Setup.")
    df=_allocation_df()
    display=df.copy()
    st.dataframe(display.drop(columns=["id"]),use_container_width=True,hide_index=True)
    a,e,d,i=st.tabs(["＋ Add","✎ Edit","🗑 Delete","⇧ Import Excel"])
    with a:_add_allocation()
    with e:_edit_allocation(df)
    with d:_delete_allocation(df)
    with i:_import_allocation_excel()


def _work_df():
    return db_df("""SELECT w.id,w.project_id,COALESCE(p.name,'') project_name,
        w.start_date,w.end_date,w.activity_type,w.task,w.priority,w.pic,w.status,w.notes
        FROM work_activity w LEFT JOIN projects p ON p.id=w.project_id
        ORDER BY w.start_date,w.project_id,w.id""")


def _add_work():
    projects=_project_options(); staff=_staff_options()
    if projects.empty: st.warning("Buat Project terlebih dahulu."); return
    pids=projects["id"].tolist()
    people=staff["name"].tolist() if not staff.empty else []
    ats=master_values("activity_type","activity_type")
    pris=master_values("priority","priority")
    statuses=master_values("task_status","status")
    with st.form("v4_add_work",clear_on_submit=True):
        c1,c2=st.columns(2)
        with c1:
            pid=st.selectbox("Project ID",pids)
            start=st.date_input("Start Date")
            end=st.date_input("End Date")
            at=_select_or_empty("Activity Type",ats)
        with c2:
            task=st.text_input("Deliverable / Task *")
            priority=_select_or_empty("Priority",pris)
            pic=_select_or_empty("PIC",people)
            status=_select_or_empty("Status",statuses)
        notes=st.text_area("Notes")
        save=st.form_submit_button("Save Activity",type="primary",use_container_width=True)
    if save:
        if not task.strip(): st.error("Deliverable / Task wajib diisi."); return
        if end<start: st.error("End Date tidak boleh lebih awal dari Start Date."); return
        conn=get_conn()
        conn.execute("""INSERT INTO work_activity
            (project_id,start_date,end_date,activity_type,task,priority,pic,status,notes)
            VALUES (?,?,?,?,?,?,?,?,?)""",
            (pid,start.isoformat(),end.isoformat(),at,task.strip(),priority,pic,status,notes))
        conn.commit(); _clear_read_caches(); conn.close(); st.success("Work activity berhasil ditambahkan."); st.rerun()


def _edit_work(df):
    if df.empty:return
    labels={int(r.id):f"{r.project_id} • {r.task}" for _,r in df.iterrows()}
    rid=st.selectbox("Select Work Activity",list(labels),format_func=lambda x:labels[x],key="v4_work_edit_id")
    row=df[df.id==rid].iloc[0]
    projects=_project_options(); staff=_staff_options()
    ats=master_values("activity_type","activity_type"); pris=master_values("priority","priority"); sts=master_values("task_status","status")
    with st.form("v4_edit_work"):
        pid=st.selectbox("Project ID",projects["id"].tolist(),index=projects["id"].tolist().index(row["project_id"]))
        sd=safe_date(row["start_date"], date.today())
        ed=safe_date(row["end_date"], sd)
        start=st.date_input("Start Date",value=sd); end=st.date_input("End Date",value=ed)
        task=st.text_input("Deliverable / Task *",value=clean(row["task"]))
        at=_select_or_empty("Activity Type",ats,index=ats.index(row["activity_type"]) if row["activity_type"] in ats else 0)
        priority=_select_or_empty("Priority",pris,index=pris.index(row["priority"]) if row["priority"] in pris else 0)
        people=staff["name"].tolist() if not staff.empty else []
        pic=_select_or_empty("PIC",people,index=people.index(row["pic"]) if row["pic"] in people else 0)
        status=_select_or_empty("Status",sts,index=sts.index(row["status"]) if row["status"] in sts else 0)
        notes=st.text_area("Notes",value=clean(row["notes"]))
        save=st.form_submit_button("Save Changes",type="primary",use_container_width=True)
    if save:
        if not task.strip():st.error("Deliverable / Task wajib diisi.");return
        if end<start:st.error("End Date tidak boleh lebih awal dari Start Date.");return
        conn=get_conn()
        conn.execute("""UPDATE work_activity SET project_id=?,start_date=?,end_date=?,activity_type=?,
            task=?,priority=?,pic=?,status=?,notes=? WHERE id=?""",
            (pid,start.isoformat(),end.isoformat(),at,task.strip(),priority,pic,status,notes,int(rid)))
        conn.commit(); _clear_read_caches();conn.close();st.success("Work activity diperbarui.");st.rerun()


def _delete_work(df):
    if df.empty:return
    labels={int(r.id):f"{r.project_id} • {r.task}" for _,r in df.iterrows()}
    rid=st.selectbox("Select Work Activity to Delete",list(labels),format_func=lambda x:labels[x],key="v4_work_del_id")
    if st.button("Delete Permanently",key="v4_work_delete",type="secondary"):
        conn=get_conn();conn.execute("DELETE FROM work_activity WHERE id=?",(int(rid),));conn.commit(); _clear_read_caches();conn.close();st.success("Activity dihapus.");st.rerun()


def _meeting_df():
    return db_df("""SELECT m.id,m.activity_date,m.start_time,m.end_time,m.project_id,
        COALESCE(p.name,'') project_name,m.meeting_type,m.attendee_1,m.attendee_2,m.attendee_3,
        m.attendee_4,m.location,m.agenda_notes
        FROM meeting_activity m LEFT JOIN projects p ON p.id=m.project_id
        ORDER BY m.activity_date,m.start_time,m.id""")


def _add_meeting():
    projects=_project_options(); pids=["No Project"]+(projects["id"].tolist() if not projects.empty else [])
    types=master_values("meeting_type","meeting_type"); locs=master_values("meeting_location","location")
    with st.form("v4_add_meeting",clear_on_submit=True):
        c1,c2=st.columns(2)
        with c1:
            d=st.date_input("Date")
            c3,c4=st.columns(2)
            with c3: start=st.time_input("Start")
            with c4: end=st.time_input("End")
            pid=st.selectbox("Project ID",pids)
            mt=_select_or_empty("Meeting Type",types)
        with c2:
            loc=_select_or_empty("Location",locs)
            a1=st.text_input("Attendee 1")
            a2=st.text_input("Attendee 2")
            a3=st.text_input("Attendee 3")
            a4=st.text_input("Attendee 4")
        notes=st.text_area("Agenda / Notes")
        save=st.form_submit_button("Save Meeting",type="primary",use_container_width=True)
    if save:
        if end<start:st.error("End time tidak boleh lebih awal dari Start.");return
        conn=get_conn()
        conn.execute("""INSERT INTO meeting_activity
            (activity_date,start_time,end_time,project_id,meeting_type,attendee_1,attendee_2,attendee_3,attendee_4,location,agenda_notes)
            VALUES (?,?,?,?,?,?,?,?,?,?,?)""",
            (d.isoformat(),start.strftime("%H:%M"),end.strftime("%H:%M"),
             None if pid=="No Project" else pid,mt,a1,a2,a3,a4,loc,notes))
        conn.commit(); _clear_read_caches();conn.close();st.success("Meeting berhasil ditambahkan.");st.rerun()


def _edit_meeting(df):
    if df.empty:return
    labels={int(r.id):f"{r.activity_date} • {r.meeting_type} • {r.project_id or 'No Project'}" for _,r in df.iterrows()}
    rid=st.selectbox("Select Meeting",list(labels),format_func=lambda x:labels[x],key="v4_meet_edit_id")
    row=df[df.id==rid].iloc[0]; projects=_project_options()
    pids=["No Project"]+(projects["id"].tolist() if not projects.empty else [])
    types=master_values("meeting_type","meeting_type");locs=master_values("meeting_location","location")
    with st.form("v4_edit_meeting"):
        d=safe_date(row["activity_date"], date.today())
        d=st.date_input("Date",value=d)
        def parse_t(v,default):
            try:return datetime.strptime(str(v),"%H:%M").time()
            except:return default
        start=st.time_input("Start",value=parse_t(row["start_time"],time(9,0)))
        end=st.time_input("End",value=parse_t(row["end_time"],time(10,0)))
        pid0=row["project_id"] or "No Project"
        pid=st.selectbox("Project ID",pids,index=pids.index(pid0) if pid0 in pids else 0)
        mt=_select_or_empty("Meeting Type",types,index=types.index(row["meeting_type"]) if row["meeting_type"] in types else 0)
        loc=_select_or_empty("Location",locs,index=locs.index(row["location"]) if row["location"] in locs else 0)
        a1=st.text_input("Attendee 1",value=clean(row["attendee_1"]))
        a2=st.text_input("Attendee 2",value=clean(row["attendee_2"]))
        a3=st.text_input("Attendee 3",value=clean(row["attendee_3"]))
        a4=st.text_input("Attendee 4",value=clean(row["attendee_4"]))
        notes=st.text_area("Agenda / Notes",value=clean(row["agenda_notes"]))
        save=st.form_submit_button("Save Changes",type="primary",use_container_width=True)
    if save:
        if end<start:st.error("End time tidak boleh lebih awal dari Start.");return
        conn=get_conn()
        conn.execute("""UPDATE meeting_activity SET activity_date=?,start_time=?,end_time=?,project_id=?,meeting_type=?,
            attendee_1=?,attendee_2=?,attendee_3=?,attendee_4=?,location=?,agenda_notes=? WHERE id=?""",
            (d.isoformat(),start.strftime("%H:%M"),end.strftime("%H:%M"),None if pid=="No Project" else pid,mt,a1,a2,a3,a4,loc,notes,int(rid)))
        conn.commit(); _clear_read_caches();conn.close();st.success("Meeting diperbarui.");st.rerun()


def _delete_meeting(df):
    if df.empty:return
    labels={int(r.id):f"{r.activity_date} • {r.meeting_type}" for _,r in df.iterrows()}
    rid=st.selectbox("Select Meeting to Delete",list(labels),format_func=lambda x:labels[x],key="v4_meet_del_id")
    if st.button("Delete Permanently",key="v4_meet_delete",type="secondary"):
        conn=get_conn();conn.execute("DELETE FROM meeting_activity WHERE id=?",(int(rid),));conn.commit(); _clear_read_caches();conn.close();st.success("Meeting dihapus.");st.rerun()


def _other_df():
    return db_df("""SELECT id,activity_date,activity,related_staff,notes
                    FROM other_activity ORDER BY activity_date,id""")


def _add_other():
    sdf=_staff_options(); people=["No Specific Staff"]+(sdf["name"].tolist() if not sdf.empty else [])
    with st.form("v4_add_other",clear_on_submit=True):
        d=st.date_input("Date")
        activity=st.text_input("Other Activity *")
        person=st.selectbox("Related Staff",people)
        notes=st.text_area("Notes")
        save=st.form_submit_button("Save Activity",type="primary",use_container_width=True)
    if save:
        if not activity.strip():st.error("Other Activity wajib diisi.");return
        conn=get_conn()
        conn.execute("INSERT INTO other_activity(activity_date,activity,related_staff,notes) VALUES (?,?,?,?)",
                     (d.isoformat(),activity.strip(),None if person=="No Specific Staff" else person,notes))
        conn.commit(); _clear_read_caches();conn.close();st.success("Other activity berhasil ditambahkan.");st.rerun()


def _edit_other(df):
    if df.empty:return
    labels={int(r.id):f"{r.activity_date} • {r.activity}" for _,r in df.iterrows()}
    rid=st.selectbox("Select Other Activity",list(labels),format_func=lambda x:labels[x],key="v4_other_edit_id")
    row=df[df.id==rid].iloc[0]
    sdf=_staff_options();people=["No Specific Staff"]+(sdf["name"].tolist() if not sdf.empty else [])
    p0=row["related_staff"] or "No Specific Staff"
    with st.form("v4_edit_other"):
        d=safe_date(row["activity_date"], date.today())
        d=st.date_input("Date",value=d)
        activity=st.text_input("Other Activity *",value=clean(row["activity"]))
        person=st.selectbox("Related Staff",people,index=people.index(p0) if p0 in people else 0)
        notes=st.text_area("Notes",value=clean(row["notes"]))
        save=st.form_submit_button("Save Changes",type="primary",use_container_width=True)
    if save:
        if not activity.strip():st.error("Other Activity wajib diisi.");return
        conn=get_conn()
        conn.execute("UPDATE other_activity SET activity_date=?,activity=?,related_staff=?,notes=? WHERE id=?",
                     (d.isoformat(),activity.strip(),None if person=="No Specific Staff" else person,notes,int(rid)))
        conn.commit(); _clear_read_caches();conn.close();st.success("Other activity diperbarui.");st.rerun()


def _delete_other(df):
    if df.empty:return
    labels={int(r.id):f"{r.activity_date} • {r.activity}" for _,r in df.iterrows()}
    rid=st.selectbox("Select Other Activity to Delete",list(labels),format_func=lambda x:labels[x],key="v4_other_del_id")
    if st.button("Delete Permanently",key="v4_other_delete",type="secondary"):
        conn=get_conn();conn.execute("DELETE FROM other_activity WHERE id=?",(int(rid),));conn.commit(); _clear_read_caches();conn.close();st.success("Other activity dihapus.");st.rerun()


def input_activities_page():
    st.markdown('<div class="app-title">Activities</div>',unsafe_allow_html=True)
    st.markdown("Operational activities feeding the Weekly Dashboard.")
    t1,t2,t3=st.tabs(["Work Activity","Meeting Activity","Other Activities"])
    with t1:
        df=_work_df();st.dataframe(df.drop(columns=["id"]),use_container_width=True,hide_index=True)
        a,e,d=st.tabs(["＋ Add","✎ Edit","🗑 Delete"])
        with a:_add_work()
        with e:_edit_work(df)
        with d:_delete_work(df)
    with t2:
        df=_meeting_df();st.dataframe(df.drop(columns=["id"]),use_container_width=True,hide_index=True)
        a,e,d=st.tabs(["＋ Add","✎ Edit","🗑 Delete"])
        with a:_add_meeting()
        with e:_edit_meeting(df)
        with d:_delete_meeting(df)
    with t3:
        df=_other_df();st.dataframe(df.drop(columns=["id"]),use_container_width=True,hide_index=True)
        a,e,d=st.tabs(["＋ Add","✎ Edit","🗑 Delete"])
        with a:_add_other()
        with e:_edit_other(df)
        with d:_delete_other(df)



def ensure_freelance_mapping_schema():
    """Ensure Freelance Project Mapping uses the current status-only model.

    Lead assignment is independent from this table. A project can therefore
    have a Lead and any number of Freelance mappings.
    """
    conn=get_conn()
    try:
        info_map=table_columns(conn, "freelance_project_mapping")
        info=list(info_map)

        if not info:
            conn.execute("""
                CREATE TABLE IF NOT EXISTS freelance_project_mapping (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    freelancer TEXT NOT NULL,
                    project_id TEXT NOT NULL,
                    mapping_status TEXT,
                    created_at TEXT DEFAULT CURRENT_TIMESTAMP
                )
            """)
        else:
            cols={name: {"notnull": meta["notnull"]} for name,meta in info_map.items()}
            # Legacy versions had start_date/end_date/notes. If any of those
            # are NOT NULL, SQLite requires a migration because the new UI
            # intentionally does not provide those fields.
            legacy_required=any(
                name in cols and bool(cols[name]["notnull"])
                for name in ("start_date","end_date","notes")
            )
            if legacy_required:
                conn.execute("""
                    CREATE TABLE freelance_project_mapping_new (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        freelancer TEXT NOT NULL,
                        project_id TEXT NOT NULL,
                        mapping_status TEXT,
                        created_at TEXT DEFAULT CURRENT_TIMESTAMP
                    )
                """)
                conn.execute("""
                    INSERT INTO freelance_project_mapping_new
                        (id,freelancer,project_id,mapping_status,created_at)
                    SELECT id,freelancer,project_id,mapping_status,created_at
                    FROM freelance_project_mapping
                """)
                conn.execute("DROP TABLE freelance_project_mapping")
                conn.execute(
                    "ALTER TABLE freelance_project_mapping_new "
                    "RENAME TO freelance_project_mapping"
                )
            elif "mapping_status" not in cols:
                conn.execute(
                    "ALTER TABLE freelance_project_mapping "
                    "ADD COLUMN mapping_status TEXT"
                )

        # Duplicate rule is only Freelance + Project. Lead is irrelevant.
        try:
            conn.execute("""
                CREATE UNIQUE INDEX IF NOT EXISTS ux_freelance_project_pair
                ON freelance_project_mapping(freelancer, project_id)
            """)
        except DB_INTEGRITY_ERRORS:
            # Keep existing duplicate legacy rows; application-level checks
            # prevent creating new duplicates.
            pass

        master=master_values("mapping_status","status")
        default_status=master[0] if master else None
        if default_status:
            conn.execute(
                "UPDATE freelance_project_mapping SET mapping_status=? "
                "WHERE mapping_status IS NULL OR TRIM(mapping_status)=''",
                (default_status,)
            )
        conn.commit()
        _clear_read_caches()
    finally:
        conn.close()

def freelance_mapping_df():
    return db_df("""
        SELECT m.id, m.freelancer, m.project_id,
               COALESCE(p.name,'') AS project_name,
               COALESCE(m.mapping_status,'') AS mapping_status
        FROM freelance_project_mapping m
        LEFT JOIN projects p ON p.id=m.project_id
        ORDER BY m.freelancer, m.id
    """)



def input_freelance_mapping_page():
    st.markdown('<div class="app-title">Freelance Project Mapping</div>',unsafe_allow_html=True)
    st.markdown("Map each freelance team member to one or more active projects.")

    freelancers_df=db_df("""
        SELECT name, primary_role
        FROM staff
        WHERE category='Freelance' AND active=1
        ORDER BY name
    """)
    projects=_project_options()
    df=freelance_mapping_df()

    if freelancers_df.empty:
        st.info("Belum ada Freelance aktif. Tambahkan Category = Freelance di Input Team terlebih dahulu.")
    if projects.empty:
        st.info("Belum ada Project Active. Tambahkan atau ubah Project di Input Project terlebih dahulu.")

    st.markdown("### Current Mapping")
    st.dataframe(
        df.drop(columns=["id"],errors="ignore"),
        use_container_width=True,
        hide_index=True
    )

    add_tab, edit_tab, delete_tab = st.tabs(["＋ Add Mapping","✎ Edit Mapping","🗑 Delete Mapping"])

    statuses=master_values("mapping_status","status")

    with add_tab:
        if not freelancers_df.empty and not projects.empty and statuses:
            with st.form("v4_add_freelance_mapping",clear_on_submit=True):
                people=freelancers_df["name"].tolist()
                pids=projects["id"].tolist()
                project_labels={
                    str(r["id"]): f"{r['id']} • {r['name']}"
                    for _,r in projects.iterrows()
                }
                c1,c2=st.columns(2)
                with c1:
                    freelancer=st.selectbox("Freelance *",people)
                    project_id=st.selectbox(
                        "Project *",pids,
                        format_func=lambda x: project_labels.get(str(x), str(x))
                    )
                with c2:
                    mapping_status=_select_or_empty("Mapping Status *",statuses)
                save=st.form_submit_button("Save Mapping",type="primary",use_container_width=True)

            if save:
                if not mapping_status:
                    st.error("Mapping Status wajib diisi.")
                else:
                    conn=get_conn()
                    try:
                        # IMPORTANT: Lead assignment in projects is completely
                        # independent. Only an existing Freelance+Project pair
                        # blocks a new mapping.
                        exists=conn.execute(
                            """SELECT 1 FROM freelance_project_mapping
                               WHERE freelancer=? AND project_id=? LIMIT 1""",
                            (freelancer,project_id)
                        ).fetchone()
                        if exists:
                            st.error("Freelance tersebut sudah memiliki mapping ke Project ini.")
                        else:
                            conn.execute("""INSERT INTO freelance_project_mapping
                                (freelancer,project_id,mapping_status)
                                VALUES (?,?,?)""",
                                (freelancer,project_id,mapping_status))
                            conn.commit()
                            _clear_read_caches()
                            st.success("Freelance project mapping berhasil ditambahkan.")
                            st.rerun()
                    except DB_INTEGRITY_ERRORS as exc:
                        conn.rollback()
                        st.error(f"Mapping gagal disimpan: {exc}")
                    finally:
                        conn.close()
        elif not statuses:
            st.warning("Mapping Status belum tersedia di Setup.")

    with edit_tab:
        if not df.empty and not freelancers_df.empty and not projects.empty and statuses:
            labels={int(r.id):f"{r.freelancer} • {r.project_id} • {r.project_name}" for _,r in df.iterrows()}
            rid=st.selectbox(
                "Select Mapping",[None]+list(labels),index=0,
                format_func=lambda x:"— Select Mapping —" if x is None else labels[x],
                key="v5c_fm_edit_id"
            )
            if rid is None:
                st.info("Pilih mapping terlebih dahulu.")
            else:
                row=df[df.id==rid].iloc[0]
                people=freelancers_df["name"].tolist()
                pids=projects["id"].tolist()
                project_labels={
                    str(r["id"]): f"{r['id']} • {r['name']}"
                    for _,r in projects.iterrows()
                }
                with st.form("v5c_edit_freelance_mapping"):
                    freelancer=st.selectbox(
                        "Freelance *",people,
                        index=people.index(row["freelancer"]) if row["freelancer"] in people else 0
                    )
                    project_id=st.selectbox(
                        "Project *",pids,
                        index=pids.index(row["project_id"]) if row["project_id"] in pids else 0,
                        format_func=lambda x: project_labels.get(str(x), str(x))
                    )
                    mapping_status=_select_or_empty(
                        "Mapping Status *",statuses,
                        index=statuses.index(row["mapping_status"])
                        if row["mapping_status"] in statuses else 0
                    )
                    save=st.form_submit_button("Save Changes",type="primary",use_container_width=True)
                if save:
                    if not mapping_status:
                        st.error("Mapping Status wajib diisi.")
                    else:
                        conn=get_conn()
                        try:
                            conn.execute("""UPDATE freelance_project_mapping
                                SET freelancer=?,project_id=?,mapping_status=?
                                WHERE id=?""",
                                (freelancer,project_id,mapping_status,int(rid)))
                            conn.commit()
                            _clear_read_caches()
                            st.success("Mapping berhasil diperbarui.")
                            st.rerun()
                        except DB_INTEGRITY_ERRORS as exc:
                            conn.rollback()
                            st.error(f"Mapping gagal diperbarui: {exc}")
                        finally:
                            conn.close()

    with delete_tab:
        if not df.empty:
            labels={int(r.id):f"{r.freelancer} • {r.project_id} • {r.project_name}" for _,r in df.iterrows()}
            rid=st.selectbox(
                "Select Mapping to Delete",[None]+list(labels),index=0,
                format_func=lambda x:"— Select Mapping to Delete —" if x is None else labels[x],
                key="v5c_fm_del_id"
            )
            if rid is not None:
                confirm=st.checkbox("Confirm deletion",value=False,key="v5c_fm_del_confirm")
                if st.button(
                    "Delete Permanently",key="v5c_fm_delete",type="secondary",
                    disabled=not confirm,use_container_width=True
                ):
                    conn=get_conn()
                    conn.execute("DELETE FROM freelance_project_mapping WHERE id=?",(int(rid),))
                    conn.commit()
                    _clear_read_caches()
                    conn.close()
                    st.success("Mapping berhasil dihapus.")
                    st.rerun()


def input_data_page(submodule):
    if submodule=="Input Team":
        input_team_page()
    elif submodule=="Input Project":
        input_project_page()
    elif submodule=="Staff Allocation":
        input_allocation_page()
    elif submodule=="Freelance Project Mapping":
        input_freelance_mapping_page()
    elif submodule=="Activities":
        input_activities_page()
    else:
        input_team_page()


# ------------------------------------------------------------
# SIDEBAR NAVIGATION — V3A STATIC TREE
# ------------------------------------------------------------
# V3a uses a clean static hierarchy:
# Module names are section headers; only submodules are clickable.

if "v3a_module" not in st.session_state:
    st.session_state.v3a_module = "Dashboard"
if "v3a_dashboard_submodule" not in st.session_state:
    st.session_state.v3a_dashboard_submodule = "Beranda"
if "v3a_input_submodule" not in st.session_state:
    st.session_state.v3a_input_submodule = "Input Team"
if "v3b_setup_submodule" not in st.session_state:
    st.session_state.v3b_setup_submodule = "Setup Manager"

st.sidebar.markdown(
    f'<div class="v3-brand">'
    f'<div class="v3-brand-mark"><img src="{ARAYA_LOGO_DATA_URI}" alt="Araya logo"></div>'
    f'<div><div class="v3-brand-name">ARAYASTD</div>'
    f'<div class="v3-brand-sub">Studio Control Board</div></div>'
    f'</div>',
    unsafe_allow_html=True,
)

st.sidebar.markdown(
    '<div class="v3-nav-label">MODULE</div>',
    unsafe_allow_html=True,
)

# ------------------------------------------------------------
# MODULE: DASHBOARD
# ------------------------------------------------------------
st.sidebar.markdown(
    '<div class="v3-module-heading">▣&nbsp;&nbsp;Dashboard</div>',
    unsafe_allow_html=True,
)

_home_selected = (
    st.session_state.v3a_module == "Dashboard"
    and st.session_state.v3a_dashboard_submodule == "Beranda"
)
if st.sidebar.button(
    f"{'●' if _home_selected else '○'}  Beranda",
    key="v3a_static_home",
    use_container_width=True,
):
    st.session_state.v3a_module = "Dashboard"
    st.session_state.v3a_dashboard_submodule = "Beranda"
    st.rerun()

_weekly_selected = (
    st.session_state.v3a_module == "Dashboard"
    and st.session_state.v3a_dashboard_submodule == "Weekly Dashboard"
)
if st.sidebar.button(
    f"{'●' if _weekly_selected else '○'}  Weekly Dashboard",
    key="v3a_static_weekly_dashboard",
    use_container_width=True,
):
    st.session_state.v3a_module = "Dashboard"
    st.session_state.v3a_dashboard_submodule = "Weekly Dashboard"
    # Opening Weekly Dashboard from the sidebar starts from the actual
    # current month/year; users can then browse other periods normally.
    _dash_today=datetime.now(ZoneInfo("Asia/Jakarta")).date()
    st.session_state["dash_month"]=_dash_today.month
    st.session_state["dash_year"]=_dash_today.year
    st.session_state["dash_week"]="All"
    st.session_state["dash_calendar_initialized"]=True
    st.rerun()

# ------------------------------------------------------------
# MODULE: INPUT DATA
# ------------------------------------------------------------
st.sidebar.markdown(
    '<div class="v3-module-heading">✎&nbsp;&nbsp;Input Data</div>',
    unsafe_allow_html=True,
)
for item in ["Input Team", "Input Project", "Staff Allocation", "Freelance Project Mapping", "Activities"]:
    selected = (
        st.session_state.v3a_module == "Input Data"
        and st.session_state.v3a_input_submodule == item
    )
    if st.sidebar.button(
        f"{'●' if selected else '○'}  {item}",
        key=f"v3a_static_input_{item.lower().replace(' ', '_')}",
        use_container_width=True,
    ):
        st.session_state.v3a_module = "Input Data"
        st.session_state.v3a_input_submodule = item
        st.rerun()

# ------------------------------------------------------------
# MODULE: SETUP
# ------------------------------------------------------------
st.sidebar.markdown(
    '<div class="v3-module-heading">⚙&nbsp;&nbsp;Setup</div>',
    unsafe_allow_html=True,
)

if st.sidebar.button(
    f"{'●' if st.session_state.v3a_module == 'Setup' else '○'}  Setup Manager",
    key="v3b_setup_manager",
    use_container_width=True,
):
    st.session_state.v3a_module = "Setup"
    st.session_state.v3b_setup_submodule = "Setup Manager"
    st.rerun()

@st.cache_resource(show_spinner=False)
def _initialize_database_once():
    """Create/migrate database objects once per app process instead of every rerun."""
    init_db()
    init_master_database()
    ensure_v4_input_schema()
    ensure_freelance_mapping_schema()
    return True


_initialize_database_once()

module = st.session_state.v3a_module

if module == "Dashboard":
    if st.session_state.v3a_dashboard_submodule == "Weekly Dashboard":
        weekly_dashboard()
    else:
        beranda_page()
elif module == "Input Data":
    input_data_page(st.session_state.v3a_input_submodule)
else:
    setup_page()
