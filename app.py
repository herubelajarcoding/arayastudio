DB_VERSION = 'V3b_v4'

import sqlite3
from pathlib import Path
from datetime import date, datetime, time, timedelta
import calendar
import html
from io import BytesIO

import pandas as pd
import streamlit as st


# ============================================================
# CONFIG
# ============================================================

APP_DIR = Path(__file__).resolve().parent
DATA_DIR = APP_DIR / "data"
DB_PATH = DATA_DIR / "studio_control.db"
SEED_XLSM = DATA_DIR / "Studio_Control_Board_FINAL.xlsm"

st.set_page_config(
    page_title="Studio Control Board",
    page_icon="📐",
    layout="wide",
    initial_sidebar_state="expanded",
)


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
        width:38px;
        height:38px;
        border-radius:11px;
        display:flex;
        align-items:center;
        justify-content:center;
        background:#172B4D;
        color:#FFFFFF;
        font-size:1.25rem;
        font-weight:800;
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
        font-size: 0.79rem;
        padding: 0.2rem 0;
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

def get_conn():
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(DB_PATH, check_same_thread=False)
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


def table_exists(conn, name):
    return conn.execute(
        "SELECT 1 FROM sqlite_master WHERE type='table' AND name=?",
        (name,),
    ).fetchone() is not None


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
            start_date TEXT NOT NULL,
            end_date TEXT NOT NULL,
            notes TEXT,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP,
            UNIQUE(freelancer, project_id, start_date, end_date)
        );
        """
    )
    conn.commit()
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


def project_name_map():
    conn = get_conn()
    df = pd.read_sql_query("SELECT id,name FROM projects ORDER BY id", conn)
    conn.close()
    return dict(zip(df["id"], df["name"]))


def get_projects():
    conn = get_conn()
    df = pd.read_sql_query(
        "SELECT * FROM projects ORDER BY id", conn
    )
    conn.close()
    return df


def get_staff():
    conn = get_conn()
    df = pd.read_sql_query(
        "SELECT * FROM staff WHERE active=1 ORDER BY name", conn
    )
    conn.close()
    return df


def get_refs(category):
    conn = get_conn()
    rows = conn.execute(
        "SELECT value FROM reference_values WHERE category=? ORDER BY rowid",
        (category,),
    ).fetchall()
    conn.close()
    return [r[0] for r in rows]


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
            f'title="{html.escape(title)}">'
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


def weekly_dashboard():
    projects = get_projects()
    project_options = (["All"] + projects["id"].tolist()) if not projects.empty else ["All"]

    # Date-detail overlay: clicking a date opens a modal on the same page.
    detail_date_value = st.query_params.get("detail_date")
    detail_date = None
    if detail_date_value:
        try:
            detail_date = parse_date(detail_date_value)
        except Exception:
            detail_date = None
        # Remove the trigger so closing the dialog does not make it reopen
        # on the next Streamlit rerun.
        st.query_params.pop("detail_date", None)

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

    years = {date.today().year}
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
                "MONTH", range(1, 13), index=date.today().month - 1,
                format_func=lambda x: f"📅  {month_names[x - 1]}", key="dash_month"
            )
        with c2:
            default_year_index = years.index(date.today().year) if date.today().year in years else len(years) - 1
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

        # KPI strip — visual only; values and filter logic are unchanged.
        work_count = len(work) if "work" in visible_activities else 0
        meeting_count = len(meetings) if "meeting" in visible_activities else 0
        submission_count = (
            int((work["activity_type"].str.lower() == "submission").sum())
            if "work" in visible_activities and not work.empty else 0
        )
        other_count = len(others) if "other" in visible_activities else 0

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
    conn = sqlite3.connect(DB_FILE)
    cur = conn.cursor()

    # Migration safety: remove incompatible old master schema
    try:
        cols = [r[1] for r in cur.execute("PRAGMA table_info(master_role)").fetchall()]
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
    conn.close()



def get_master_table(name):
    conn = sqlite3.connect(DB_FILE)
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
    conn = sqlite3.connect(DB_FILE)
    where = " AND ".join(f"{f}=?" for f, _, _ in fields)
    params = [values[f] for f, _, _ in fields]
    sql = f"SELECT id FROM master_{table} WHERE {where}"
    if exclude_id is not None:
        sql += " AND id<>?"
        params.append(exclude_id)
    found = conn.execute(sql, params).fetchone()
    conn.close()
    return found is not None


def _crud_add(table, fields, values):
    if _crud_duplicate(table, fields, values):
        return False, "Data yang sama sudah ada."
    names = [f for f, _, _ in fields]
    conn = sqlite3.connect(DB_FILE)
    conn.execute(
        f"INSERT INTO master_{table} ({','.join(names)}) VALUES ({','.join('?' for _ in names)})",
        [values[n] for n in names],
    )
    conn.commit()
    conn.close()
    return True, "Data berhasil ditambahkan."


def _crud_update(table, fields, row_id, values):
    if _crud_duplicate(table, fields, values, row_id):
        return False, "Data yang sama sudah ada."
    names = [f for f, _, _ in fields]
    conn = sqlite3.connect(DB_FILE)
    conn.execute(
        f"UPDATE master_{table} SET {','.join(f'{n}=?' for n in names)} WHERE id=?",
        [values[n] for n in names] + [row_id],
    )
    conn.commit()
    conn.close()
    return True, "Data berhasil diperbarui."


def _crud_delete(table, row_id):
    conn = sqlite3.connect(DB_FILE)
    conn.execute(f"DELETE FROM master_{table} WHERE id=?", (row_id,))
    conn.commit()
    conn.close()
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


def db_df(sql, params=()):
    conn=get_conn()
    df=pd.read_sql_query(sql,conn,params=params)
    conn.close()
    return df


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
    conn.close()


def _select_or_empty(label, options, key=None, index=0, help=None):
    opts=list(options)
    if not opts:
        st.warning(f"Belum ada master data untuk {label}. Silakan isi di Setup.")
        return ""
    return st.selectbox(label, opts, index=min(index,len(opts)-1), key=key, help=help)


def _project_options():
    df=db_df("SELECT id,name FROM projects ORDER BY id")
    return df


def _staff_options():
    # Active controls dropdown visibility only. Inactive staff remain stored
    # in the database and remain visible/manageable in Input Team.
    return db_df("""
        SELECT id,name,category,primary_role,active
        FROM staff
        WHERE active=1
        ORDER BY name
    """)


def _project_name(pid):
    if not pid: return ""
    conn=get_conn()
    row=conn.execute("SELECT name FROM projects WHERE id=?",(pid,)).fetchone()
    conn.close()
    return row[0] if row else ""


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
            st.success("Team member berhasil ditambahkan.")
            st.rerun()
        except sqlite3.IntegrityError:
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
            st.success("Data berhasil diperbarui.")
            st.rerun()
        except sqlite3.IntegrityError:
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
            st.success("Team member benar-benar dihapus dari database.")
            st.rerun()
        except sqlite3.IntegrityError:
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
            st.session_state["v4c_team_import_success"] = len(valid_rows)
            st.session_state["v4c_team_import_file"] = uploaded.name
            st.rerun()
        except sqlite3.IntegrityError as exc:
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


def _add_project():
    st.markdown("### Add Project")
    pdf=_project_options()
    with st.form("v4_add_project",clear_on_submit=True):
        c1,c2,c3=st.columns(3)
        with c1:
            pid=st.text_input("Project ID *")
            name=st.text_input("Project Name *")
            ptypes=master_values("project_type","project_type")
            ptype=_select_or_empty("Project Type",ptypes)
        with c2:
            start=st.date_input("Start Date")
            finish=st.date_input("Target Finish")
            sizes=master_values("project_size","project_size")
            size=_select_or_empty("Project Size",sizes)
        with c3:
            statuses=master_values("project_status","status")
            status=_select_or_empty("Project Status",statuses)
            staff=_staff_options()
            leads=staff["name"].tolist() if not staff.empty else []
            lead=_select_or_empty("Lead",leads)
            duration=_project_duration(start,finish)
            st.caption(f"Duration: {duration or '—'} month(s)")
        save=st.form_submit_button("Save Project",type="primary",use_container_width=True)
    if save:
        if not pid.strip() or not name.strip():
            st.error("Project ID dan Project Name wajib diisi."); return
        if finish<start:
            st.error("Target Finish tidak boleh lebih awal dari Start Date."); return
        conn=get_conn()
        try:
            conn.execute("""INSERT INTO projects
                (id,name,project_type,start_date,target_finish,duration_months,status,lead,project_size)
                VALUES (?,?,?,?,?,?,?,?,?)""",
                (pid.strip(),name.strip(),ptype,start.isoformat(),finish.isoformat(),
                 _project_duration(start,finish),status,lead,size))
            conn.commit(); st.success("Project berhasil ditambahkan."); st.rerun()
        except sqlite3.IntegrityError:
            st.error("Project ID sudah ada.")
        finally: conn.close()


def _edit_project(df):
    if df.empty: return
    labels={str(r.id):f"{r.id} • {r['name']}" for _,r in df.iterrows()}
    pid=st.selectbox("Select Project",list(labels),format_func=lambda x:labels[x],key="v4_proj_edit_id")
    row=df[df.id==pid].iloc[0]
    with st.form("v4_edit_project"):
        c1,c2,c3=st.columns(3)
        with c1:
            name=st.text_input("Project Name *",value=clean(row["name"]))
            ptypes=master_values("project_type","project_type")
            ptype=_select_or_empty("Project Type",ptypes,index=ptypes.index(row["project_type"]) if row["project_type"] in ptypes else 0)
        with c2:
            start=safe_date(row["start_date"], date.today())
            finish=safe_date(row["target_finish"], start)
            start=st.date_input("Start Date",value=start)
            finish=st.date_input("Target Finish",value=finish)
            sizes=master_values("project_size","project_size")
            size=_select_or_empty("Project Size",sizes,index=sizes.index(row["project_size"]) if row["project_size"] in sizes else 0)
        with c3:
            statuses=master_values("project_status","status")
            status=_select_or_empty("Project Status",statuses,index=statuses.index(row["status"]) if row["status"] in statuses else 0)
            sdf=_staff_options(); leads=sdf["name"].tolist() if not sdf.empty else []
            lead=_select_or_empty("Lead",leads,index=leads.index(row["lead"]) if row["lead"] in leads else 0)
            st.caption(f"Duration: {_project_duration(start,finish) or '—'} month(s)")
        save=st.form_submit_button("Save Changes",type="primary",use_container_width=True)
    if save:
        if not name.strip(): st.error("Project Name wajib diisi."); return
        if finish<start: st.error("Target Finish tidak boleh lebih awal dari Start Date."); return
        conn=get_conn()
        conn.execute("""UPDATE projects SET name=?,project_type=?,start_date=?,target_finish=?,
            duration_months=?,status=?,lead=?,project_size=? WHERE id=?""",
            (name.strip(),ptype,start.isoformat(),finish.isoformat(),_project_duration(start,finish),
             status,lead,size,pid))
        conn.commit(); conn.close(); st.success("Project berhasil diperbarui."); st.rerun()


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
        conn.commit(); conn.close(); st.success("Project diarsipkan agar histori aktivitas tetap aman."); st.rerun()



def _project_import_template():
    """Create an Excel template matching the projects table/input fields."""
    sample=pd.DataFrame([{
        "Project ID":"P001",
        "Project Name":"Example Project",
        "Project Type":"",
        "Start Date":"",
        "Target Finish":"",
        "Project Status":"",
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
        "Kolom Excel dapat memiliki nama berbeda; sistem akan mencocokkan "
        "kolom dengan field Project dan dapat disesuaikan manual."
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
        st.success(
            f"Import berhasil. {imported_count} Project berhasil direkam ke database."
        )
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
        ("project_id","Project ID *"),
        ("name","Project Name *"),
        ("project_type","Project Type"),
        ("start_date","Start Date"),
        ("target_finish","Target Finish"),
        ("status","Project Status"),
        ("lead","Lead"),
        ("project_size","Project Size"),
    ]

    norm={_normalize_import_header(c):c for c in source_cols}
    aliases={
        "project_id":["projectid","id","kodeproject","kodeproyek","projectcode","code"],
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
        mapping[target]=st.selectbox(
            label,options,index=default,key=f"v5_project_map_{target}"
        )

    missing=[label for target,label in target_fields[:2] if mapping[target]=="— Not mapped —"]
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
    statuses=master_values("project_status","status")
    sizes=master_values("project_size","project_size")
    leads_df=_staff_options()
    leads=leads_df["name"].tolist() if not leads_df.empty else []

    type_lookup={clean(x).lower():x for x in project_types}
    status_lookup={clean(x).lower():x for x in statuses}
    size_lookup={clean(x).lower():x for x in sizes}
    lead_lookup={clean(x).lower():x for x in leads}

    existing_ids=set(clean(x).lower() for x in db_df("SELECT id FROM projects")["id"].tolist())
    valid_rows=[]
    errors=[]

    def cell_text(target,row):
        src_col=mapping[target]
        return "" if src_col=="— Not mapped —" else clean(row[src_col])

    for ix,row in raw.iterrows():
        pid=cell_text("project_id",row)
        name=cell_text("name",row)
        ptype_raw=cell_text("project_type",row)
        status_raw=cell_text("status",row)
        lead_raw=cell_text("lead",row)
        size_raw=cell_text("project_size",row)

        if not pid:
            errors.append(f"Baris Excel {ix+2}: Project ID kosong.")
            continue
        if not name:
            errors.append(f"Baris Excel {ix+2}: Project Name kosong.")
            continue

        ptype=type_lookup.get(ptype_raw.lower()) if ptype_raw else None
        if ptype_raw and not ptype:
            errors.append(f"Baris Excel {ix+2}: Project Type '{ptype_raw}' tidak ada di Setup.")
            continue

        status=status_lookup.get(status_raw.lower()) if status_raw else None
        if status_raw and not status:
            errors.append(f"Baris Excel {ix+2}: Project Status '{status_raw}' tidak ada di Setup.")
            continue

        size=size_lookup.get(size_raw.lower()) if size_raw else None
        if size_raw and not size:
            errors.append(f"Baris Excel {ix+2}: Project Size '{size_raw}' tidak ada di Setup.")
            continue

        lead=lead_lookup.get(lead_raw.lower()) if lead_raw else None
        if lead_raw and not lead:
            errors.append(f"Baris Excel {ix+2}: Lead '{lead_raw}' tidak tersedia pada Team Active.")
            continue

        start=to_iso_date(row[mapping["start_date"]]) if mapping["start_date"]!="— Not mapped —" else None
        finish=to_iso_date(row[mapping["target_finish"]]) if mapping["target_finish"]!="— Not mapped —" else None
        if mapping["start_date"]!="— Not mapped —" and clean(row[mapping["start_date"]]) and not start:
            errors.append(f"Baris Excel {ix+2}: Start Date tidak valid.")
            continue
        if mapping["target_finish"]!="— Not mapped —" and clean(row[mapping["target_finish"]]) and not finish:
            errors.append(f"Baris Excel {ix+2}: Target Finish tidak valid.")
            continue
        if start and finish and finish<start:
            errors.append(f"Baris Excel {ix+2}: Target Finish lebih awal dari Start Date.")
            continue

        duplicate_file=any(clean(r["id"]).lower()==pid.lower() for r in valid_rows)
        if pid.lower() in existing_ids or duplicate_file:
            errors.append(f"Baris Excel {ix+2}: Project ID '{pid}' sudah ada.")
            continue

        duration=_project_duration(parse_date(start),parse_date(finish)) if start and finish else None
        valid_rows.append({
            "id":pid,"name":name,"project_type":ptype,"start_date":start,
            "target_finish":finish,"duration_months":duration,"status":status,
            "lead":lead,"project_size":size,
        })

    if errors:
        st.error(f"Ditemukan {len(errors)} masalah. Tidak ada data yang direkam.")
        st.dataframe(pd.DataFrame({"Validation Error":errors}),use_container_width=True,hide_index=True)
        return

    st.success(f"{len(valid_rows)} baris lolos validasi dan siap direkam.")
    if st.button("Save Imported Project Data",type="primary",use_container_width=True,key="v5_project_import_save"):
        conn=get_conn()
        try:
            for r in valid_rows:
                conn.execute("""INSERT INTO projects
                    (id,name,project_type,start_date,target_finish,duration_months,status,lead,project_size)
                    VALUES (?,?,?,?,?,?,?,?,?)""",
                    (r["id"],r["name"],r["project_type"],r["start_date"],r["target_finish"],
                     r["duration_months"],r["status"],r["lead"],r["project_size"]))
            conn.commit()
            st.session_state["v5_project_import_success"]=len(valid_rows)
            st.session_state["v5_project_import_file"]=uploaded.name
            st.rerun()
        except sqlite3.IntegrityError as exc:
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
    prow=conn.execute("SELECT project_size FROM projects WHERE id=?",(project_id,)).fetchone()
    srow=conn.execute("SELECT multiplier FROM master_project_size WHERE project_size=?",(prow[0],)).fetchone() if prow else None
    prow2=conn.execute("SELECT base_load FROM master_phase WHERE phase=?",(phase,)).fetchone()
    mult=float(srow[0]) if srow else 1.0
    weight=float(prow2[0]) if prow2 else 0.0
    total=mult*weight
    count=conn.execute("SELECT COUNT(*) FROM staff_allocation WHERE project_id=? AND phase=?",(project_id,phase)).fetchone()[0]
    conn.close()
    return mult,weight,total,count+1,total/(count+1) if count+1 else 0


def _allocation_df():
    return db_df("""
        SELECT a.id,a.project_id,p.name AS project_name,a.phase,a.staff,
               s.category AS staff_category,s.primary_role,a.role_on_project,
               p.project_size
        FROM staff_allocation a
        LEFT JOIN projects p ON p.id=a.project_id
        LEFT JOIN staff s ON s.name=a.staff
        ORDER BY a.project_id,a.phase,a.staff
    """)


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
            conn.commit(); st.success("Staff allocation berhasil ditambahkan."); st.rerun()
        except sqlite3.IntegrityError:
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
            conn.commit(); st.success("Allocation diperbarui."); st.rerun()
        except sqlite3.IntegrityError: st.error("Allocation yang sama sudah ada.")
        finally: conn.close()


def _delete_allocation(df):
    if df.empty:return
    labels={int(r.id):f"{r.project_id} • {r.phase} • {r.staff}" for _,r in df.iterrows()}
    rid=st.selectbox("Select Allocation to Delete",list(labels),format_func=lambda x:labels[x],key="v4_alloc_del_id")
    if st.button("Delete Permanently",key="v4_alloc_delete",type="secondary"):
        conn=get_conn(); conn.execute("DELETE FROM staff_allocation WHERE id=?",(int(rid),)); conn.commit(); conn.close()
        st.success("Allocation dihapus."); st.rerun()


def input_allocation_page():
    st.markdown('<div class="app-title">Staff Allocation</div>',unsafe_allow_html=True)
    st.markdown("Assign staff to a project phase. Load parameters are derived from Setup.")
    df=_allocation_df()
    display=df.copy()
    st.dataframe(display.drop(columns=["id"]),use_container_width=True,hide_index=True)
    a,e,d=st.tabs(["＋ Add","✎ Edit","🗑 Delete"])
    with a:_add_allocation()
    with e:_edit_allocation(df)
    with d:_delete_allocation(df)


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
        conn.commit(); conn.close(); st.success("Work activity berhasil ditambahkan."); st.rerun()


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
        conn.commit();conn.close();st.success("Work activity diperbarui.");st.rerun()


def _delete_work(df):
    if df.empty:return
    labels={int(r.id):f"{r.project_id} • {r.task}" for _,r in df.iterrows()}
    rid=st.selectbox("Select Work Activity to Delete",list(labels),format_func=lambda x:labels[x],key="v4_work_del_id")
    if st.button("Delete Permanently",key="v4_work_delete",type="secondary"):
        conn=get_conn();conn.execute("DELETE FROM work_activity WHERE id=?",(int(rid),));conn.commit();conn.close();st.success("Activity dihapus.");st.rerun()


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
        conn.commit();conn.close();st.success("Meeting berhasil ditambahkan.");st.rerun()


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
        conn.commit();conn.close();st.success("Meeting diperbarui.");st.rerun()


def _delete_meeting(df):
    if df.empty:return
    labels={int(r.id):f"{r.activity_date} • {r.meeting_type}" for _,r in df.iterrows()}
    rid=st.selectbox("Select Meeting to Delete",list(labels),format_func=lambda x:labels[x],key="v4_meet_del_id")
    if st.button("Delete Permanently",key="v4_meet_delete",type="secondary"):
        conn=get_conn();conn.execute("DELETE FROM meeting_activity WHERE id=?",(int(rid),));conn.commit();conn.close();st.success("Meeting dihapus.");st.rerun()


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
        conn.commit();conn.close();st.success("Other activity berhasil ditambahkan.");st.rerun()


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
        conn.commit();conn.close();st.success("Other activity diperbarui.");st.rerun()


def _delete_other(df):
    if df.empty:return
    labels={int(r.id):f"{r.activity_date} • {r.activity}" for _,r in df.iterrows()}
    rid=st.selectbox("Select Other Activity to Delete",list(labels),format_func=lambda x:labels[x],key="v4_other_del_id")
    if st.button("Delete Permanently",key="v4_other_delete",type="secondary"):
        conn=get_conn();conn.execute("DELETE FROM other_activity WHERE id=?",(int(rid),));conn.commit();conn.close();st.success("Other activity dihapus.");st.rerun()


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



def freelance_mapping_df():
    return db_df("""
        SELECT m.id, m.freelancer, m.project_id,
               COALESCE(p.name,'') AS project_name,
               m.start_date, m.end_date, m.notes
        FROM freelance_project_mapping m
        LEFT JOIN projects p ON p.id=m.project_id
        ORDER BY m.freelancer, m.start_date, m.id
    """)


def input_freelance_mapping_page():
    st.markdown('<div class="app-title">Freelance Project Mapping</div>',unsafe_allow_html=True)
    st.markdown("Map each freelance team member to one or more projects and define the assignment period.")

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
        st.info("Belum ada Project. Tambahkan Project di Input Project terlebih dahulu.")

    st.markdown("### Current Mapping")
    st.dataframe(
        df.drop(columns=["id"],errors="ignore"),
        use_container_width=True,
        hide_index=True
    )

    add_tab, edit_tab, delete_tab = st.tabs(["＋ Add Mapping","✎ Edit Mapping","🗑 Delete Mapping"])

    with add_tab:
        if not freelancers_df.empty and not projects.empty:
            with st.form("v4_add_freelance_mapping",clear_on_submit=True):
                people=freelancers_df["name"].tolist()
                pids=projects["id"].tolist()
                c1,c2=st.columns(2)
                with c1:
                    freelancer=st.selectbox("Freelance *",people)
                    project_id=st.selectbox("Project *",pids)
                with c2:
                    start=st.date_input("Assignment Start *")
                    end=st.date_input("Assignment End *")
                notes=st.text_area("Notes")
                save=st.form_submit_button("Save Mapping",type="primary",use_container_width=True)

            if save:
                if end<start:
                    st.error("Assignment End tidak boleh lebih awal dari Assignment Start.")
                else:
                    conn=get_conn()
                    try:
                        conn.execute("""INSERT INTO freelance_project_mapping
                            (freelancer,project_id,start_date,end_date,notes)
                            VALUES (?,?,?,?,?)""",
                            (freelancer,project_id,start.isoformat(),end.isoformat(),notes))
                        conn.commit()
                        st.success("Freelance project mapping berhasil ditambahkan.")
                        st.rerun()
                    except sqlite3.IntegrityError:
                        st.error("Mapping yang sama sudah ada.")
                    finally:
                        conn.close()

    with edit_tab:
        if not df.empty and not freelancers_df.empty and not projects.empty:
            labels={int(r.id):f"{r.freelancer} • {r.project_id} • {r.start_date} – {r.end_date}" for _,r in df.iterrows()}
            rid=st.selectbox("Select Mapping",list(labels),format_func=lambda x:labels[x],key="v4_fm_edit_id")
            row=df[df.id==rid].iloc[0]
            people=freelancers_df["name"].tolist()
            pids=projects["id"].tolist()
            with st.form("v4_edit_freelance_mapping"):
                freelancer=st.selectbox("Freelance *",people,index=people.index(row["freelancer"]) if row["freelancer"] in people else 0)
                project_id=st.selectbox("Project *",pids,index=pids.index(row["project_id"]) if row["project_id"] in pids else 0)
                start=st.date_input("Assignment Start *",value=safe_date(row["start_date"], date.today()))
                end=st.date_input("Assignment End *",value=safe_date(row["end_date"], date.today()))
                notes=st.text_area("Notes",value=clean(row["notes"]))
                save=st.form_submit_button("Save Changes",type="primary",use_container_width=True)
            if save:
                if end<start:
                    st.error("Assignment End tidak boleh lebih awal dari Assignment Start.")
                else:
                    conn=get_conn()
                    try:
                        conn.execute("""UPDATE freelance_project_mapping
                            SET freelancer=?,project_id=?,start_date=?,end_date=?,notes=?
                            WHERE id=?""",
                            (freelancer,project_id,start.isoformat(),end.isoformat(),notes,int(rid)))
                        conn.commit()
                        st.success("Mapping berhasil diperbarui.")
                        st.rerun()
                    except sqlite3.IntegrityError:
                        st.error("Mapping yang sama sudah ada.")
                    finally:
                        conn.close()

    with delete_tab:
        if not df.empty:
            labels={int(r.id):f"{r.freelancer} • {r.project_id} • {r.start_date} – {r.end_date}" for _,r in df.iterrows()}
            rid=st.selectbox("Select Mapping to Delete",list(labels),format_func=lambda x:labels[x],key="v4_fm_del_id")
            confirm=st.checkbox("Confirm deletion",key="v4_fm_del_confirm")
            if st.button("Delete Permanently",key="v4_fm_delete",type="secondary",disabled=not confirm,use_container_width=True):
                conn=get_conn()
                conn.execute("DELETE FROM freelance_project_mapping WHERE id=?",(int(rid),))
                conn.commit()
                conn.close()
                st.success("Mapping berhasil dihapus.")
                st.rerun()


def input_data_page(submodule):
    ensure_v4_input_schema()
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
    st.session_state.v3a_dashboard_submodule = "Weekly Dashboard"
if "v3a_input_submodule" not in st.session_state:
    st.session_state.v3a_input_submodule = "Input Team"
if "v3b_setup_submodule" not in st.session_state:
    st.session_state.v3b_setup_submodule = "Setup Manager"

st.sidebar.markdown(
    '<div class="v3-brand">'
    '<div class="v3-brand-mark">◢</div>'
    '<div><div class="v3-brand-name">ARAYASTD</div>'
    '<div class="v3-brand-sub">Studio Control Board</div></div>'
    '</div>',
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
if st.sidebar.button(
    f"{'●' if st.session_state.v3a_module == 'Dashboard' else '○'}  Weekly Dashboard",
    key="v3a_static_weekly_dashboard",
    use_container_width=True,
):
    st.session_state.v3a_module = "Dashboard"
    st.session_state.v3a_dashboard_submodule = "Weekly Dashboard"
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

init_db()
init_master_database()
ensure_v4_input_schema()

module = st.session_state.v3a_module

if module == "Dashboard":
    weekly_dashboard()
elif module == "Input Data":
    input_data_page(st.session_state.v3a_input_submodule)
else:
    setup_page()
