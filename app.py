
import sqlite3
from pathlib import Path
from datetime import date, datetime, time, timedelta
import calendar
import html

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
    div[data-testid="stVerticalBlock"]:has(#weekly-dashboard-anchor) {
        position: -webkit-sticky !important; position: sticky !important; top: 0 !important;
        z-index: 1000 !important; background: rgba(255,255,255,.98) !important;
        padding: .15rem 0 .85rem !important; margin-bottom: .35rem !important;
        border-bottom: 1px solid #D0D5DD !important;
        box-shadow: 0 6px 14px rgba(16,24,40,.07) !important;
        backdrop-filter: blur(10px) !important;
    }
    #weekly-dashboard-anchor {height: 0; overflow: hidden;}
    .activity-filter-note {font-size:.72rem;color:#667085;margin-top:-.35rem;margin-bottom:.25rem;}
    .app-title {font-size: 2rem; font-weight: 750; margin-bottom: 0.1rem;}
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
    .lane-label {
        min-height: 130px;
        padding: 0.75rem 0.55rem;
        font-weight: 800;
        font-size: 0.78rem;
        letter-spacing: .04em;
        border-right: 1px solid #D0D5DD;
        color:#344054;
    }
    .lane-label.work-lane {background:#E8F3FF;}
    .lane-label.meeting-lane {background:#FFF7D6;}
    .lane-label.other-lane {background:#F2E9FF;}
    .cell.work-cell {background:#F4FAFF;}
    .cell.meeting-cell {background:#FFFBEA;}
    .meeting-cell .project {color:#1F2937;}
    .meeting-cell .task {color:#172B4D; font-weight:650;}
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
        margin: 0.12rem 0 0.35rem;
    }
    .project + .project {
        margin-top: 0.85rem;
        padding-top: 0.55rem;
        border-top: 1px solid rgba(16,24,40,.10);
    }
    .task {font-size: 0.80rem; line-height: 1.4; margin: 0.28rem 0; color:#172B4D;}
    .task.submission {font-weight: 700; color:#C62828; background:#FFE7E7; border-radius:6px; padding:2px 5px;}
    .sign {color:#D92D20; font-weight: 900; margin-left: 0.2rem;}
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
    .kpi-row {display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:16px;margin:8px 0 12px;}
    .kpi {position:relative;min-height:126px;padding:20px 22px;border:1px solid rgba(16,24,40,.06);border-radius:15px;overflow:hidden;box-shadow:0 3px 12px rgba(16,24,40,.05);display:flex;align-items:center;gap:18px;}
    .kpi::after {content:"";position:absolute;width:150px;height:150px;right:-42px;bottom:-82px;border-radius:50%;background:rgba(255,255,255,.38);}
    .kpi::before {content:"";position:absolute;width:92px;height:92px;right:-24px;top:-35px;border-radius:50%;background:rgba(255,255,255,.25);}
    .kpi.work {background:linear-gradient(135deg,#EAF4FF,#DCEEFF);}
    .kpi.meeting {background:linear-gradient(135deg,#FFF9DF,#FFF1B9);}
    .kpi.submission {background:linear-gradient(135deg,#FFF0F1,#FFE0E3);}
    .kpi.other {background:linear-gradient(135deg,#F4ECFF,#E9DDFF);}
    .kpi-icon {width:54px;height:54px;min-width:54px;border-radius:14px;display:flex;align-items:center;justify-content:center;position:relative;z-index:2;}
    .kpi.work .kpi-icon {background:#1479E9;color:#fff;}
    .kpi.meeting .kpi-icon {background:#F5B400;color:#fff;}
    .kpi.submission .kpi-icon {background:#EF4444;color:#fff;}
    .kpi.other .kpi-icon {background:#8B5CF6;color:#fff;}
    .kpi-content {position:relative;z-index:2;}
    .kpi-label {font-size:.76rem;font-weight:800;letter-spacing:.03em;text-transform:uppercase;margin-bottom:4px;}
    .kpi.work .kpi-label {color:#1479E9;}
    .kpi.meeting .kpi-label {color:#C88A00;}
    .kpi.submission .kpi-label {color:#E33A40;}
    .kpi.other .kpi-label {color:#7040D8;}
    .kpi-value {font-size:2rem;line-height:1;font-weight:850;color:#102A56;}
    @media (max-width:900px){.kpi-row{grid-template-columns:repeat(2,minmax(0,1fr));}}
    @media (max-width:560px){.kpi-row{grid-template-columns:1fr;}}
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

def clean(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    return str(v).strip()


def numeric_or_none(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return None
    try:
        return float(v)
    except Exception:
        return None


def to_iso_date(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return None
    if isinstance(v, pd.Timestamp):
        return v.date().isoformat()
    if isinstance(v, datetime):
        return v.date().isoformat()
    if isinstance(v, date):
        return v.isoformat()
    try:
        return pd.to_datetime(v).date().isoformat()
    except Exception:
        return None


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
        ORDER BY w.end_date, w.project_id, w.id
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
        ORDER BY m.activity_date, m.start_time, m.id
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


def render_work_cell(rows):
    if not rows:
        return '<div class="empty">—</div>'

    chunks = []
    last_project = None

    for row in rows:
        pid = clean(row.get("project_id"))
        pname = clean(row.get("project_name"))
        if pid != last_project:
            if pid:
                chunks.append(
                    f'<div class="project">{html.escape(pid)} | {html.escape(pname)}</div>'
                )
            last_project = pid

        task = html.escape(clean(row.get("task")))
        pic = html.escape(clean(row.get("pic")))
        activity_type = clean(row.get("activity_type")).lower()

        sign = '<span class="sign">▲!</span>' if activity_type == "submission" else ""
        pic_html = f" <span style='color:#667085'>({pic})</span>" if pic else ""

        chunks.append(
            f'<div class="task {"submission" if activity_type == "submission" else ""}">'
            f'• {task}{pic_html} {sign}</div>'
        )

    return "".join(chunks)


def render_meeting_cell(rows):
    if not rows:
        return '<div class="empty">—</div>'

    chunks = []
    last_project = None

    for row in rows:
        pid = clean(row.get("project_id"))
        pname = clean(row.get("project_name"))
        if pid != last_project:
            if pid:
                chunks.append(
                    f'<div class="project">{html.escape(pid)} | {html.escape(pname)}</div>'
                )
            last_project = pid

        attendees = [
            clean(row.get("attendee_1")),
            clean(row.get("attendee_2")),
            clean(row.get("attendee_3")),
            clean(row.get("attendee_4")),
        ]
        attendees = ", ".join([a for a in attendees if a])

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

    return "".join(chunks)


def render_other_cell(rows):
    if not rows:
        return '<div class="empty">—</div>'

    chunks = []
    for row in rows:
        activity = html.escape(clean(row.get("activity")))
        staff = html.escape(clean(row.get("related_staff")))
        chunks.append(
            f'<div class="other-item">• {activity}'
            + (f'<div class="meta">Related Staff: {staff}</div>' if staff else "")
            + "</div>"
        )
    return "".join(chunks)


def render_week(start_date, end_date, work, meetings, others, visible_activities):
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

    week_no = ((start_date.day - 1) // 7) + 1
    st.markdown(
        f'<div class="week-title">WEEK {week_no} • {start_date.strftime("%d %b")} – '
        f'{end_date.strftime("%d %b %Y")}</div>',
        unsafe_allow_html=True,
    )

    grid = []
    # Header row
    grid.append('<div class="schedule-head activity-head">ACTIVITY</div>')
    for d in days:
        grid.append(
            f'<div class="schedule-head"><div class="dow">{fmt_day(d)}</div>'
            f'<div class="day">{d.strftime("%d %b")}</div></div>'
        )

    lane_specs = [("WORK", "work"), ("MEETING", "meeting"), ("OTHER", "other")]
    lane_specs = [x for x in lane_specs if x[1] in visible_activities]

    for lane_name, lane_type in lane_specs:
        grid.append(
            f'<div class="lane-label {lane_type}-lane">{lane_name}</div>'
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
    """Compact checklist filter. When All is selected, individual options are
    intentionally hidden. When All is cleared, individual checkboxes appear.
    Returns (effective_selection, all_selected).
    """
    all_key = f"{key_prefix}_all"
    if all_key not in st.session_state:
        st.session_state[all_key] = True

    for opt in options:
        key = f"{key_prefix}_{opt}"
        if key not in st.session_state:
            st.session_state[key] = False

    selected = []
    with st.popover(label, use_container_width=True):
        all_selected = st.checkbox(all_label, key=all_key)
        if all_selected:
            st.caption("All selected")
        else:
            for opt in options:
                key = f"{key_prefix}_{opt}"
                text = format_func(opt) if format_func else opt
                if st.checkbox(text, key=key):
                    selected.append(opt)

    if st.session_state[all_key]:
        return list(options), True
    return selected, False


def weekly_dashboard():
    projects = get_projects()
    project_options = (["All"] + projects["id"].tolist()) if not projects.empty else ["All"]

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

        c1, c2, c3, c4, c5 = st.columns([1.0, .8, 1.55, 1.05, 1.1])
        with c1:
            selected_month_no = st.selectbox(
                "MONTH", range(1, 13), index=date.today().month - 1,
                format_func=lambda x: month_names[x - 1], key="dash_month"
            )
        with c2:
            default_year_index = years.index(date.today().year) if date.today().year in years else len(years) - 1
            selected_year = st.selectbox("YEAR", years, index=default_year_index, key="dash_year")

        selected_month = date(selected_year, selected_month_no, 1)
        weeks = month_weeks(selected_month.year, selected_month.month)

        with c3:
            selected_projects, project_all = checklist_filter(
                "PROJECT",
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
                "ACTIVITY",
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

        # KPI strip — visual only; filter logic above remains unchanged.
        work_count = len(work) if "work" in visible_activities else 0
        meeting_count = len(meetings) if "meeting" in visible_activities else 0
        submission_count = int((work["activity_type"].str.lower() == "submission").sum()) if "work" in visible_activities and not work.empty else 0
        other_count = len(others) if "other" in visible_activities else 0

        icon_work = '''<svg width="29" height="29" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M6 2h9l3 3v17H6z"/><path d="M15 2v4h4"/><path d="M9 11h6M9 15h6M9 19h4"/></svg>'''
        icon_meeting = '''<svg width="31" height="31" viewBox="0 0 24 24" fill="currentColor"><circle cx="9" cy="8" r="4"/><circle cx="17" cy="9" r="3"/><path d="M2.5 21c.3-4 2.6-6 6.5-6s6.2 2 6.5 6z"/><path d="M14.5 15.5c3.2.1 5 1.8 5.5 4.5h-4.2c-.2-1.7-.6-3.1-1.3-4.5z"/></svg>'''
        icon_submission = '''<svg width="31" height="31" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round"><circle cx="5" cy="6" r="1.2" fill="currentColor"/><circle cx="5" cy="12" r="1.2" fill="currentColor"/><circle cx="5" cy="18" r="1.2" fill="currentColor"/><path d="M10 6h10M10 12h10M10 18h10"/></svg>'''
        icon_other = '''<svg width="31" height="31" viewBox="0 0 24 24" fill="currentColor"><circle cx="12" cy="12" r="9"/><circle cx="8" cy="12" r="1.25" fill="white"/><circle cx="12" cy="12" r="1.25" fill="white"/><circle cx="16" cy="12" r="1.25" fill="white"/></svg>'''

        cards = [
            ("work", "WORK ITEMS", work_count, icon_work),
            ("meeting", "MEETINGS", meeting_count, icon_meeting),
            ("submission", "SUBMISSIONS", submission_count, icon_submission),
            ("other", "OTHER ACTIVITIES", other_count, icon_other),
        ]
        cards_html = '<div class="kpi-row">'
        for css_class, label, value, icon in cards:
            cards_html += f'<div class="kpi {css_class}"><div class="kpi-icon">{icon}</div><div class="kpi-content"><div class="kpi-label">{label}</div><div class="kpi-value">{value}</div></div></div>'
        cards_html += '</div>'
        st.markdown(cards_html, unsafe_allow_html=True)
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

        render_week(week_start, week_end, w, m, o, visible_activities)


# ============================================================
# INPUT MODULES
# ============================================================

def setup_page():
    st.markdown('<div class="app-title">Setup / Master Data</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="app-subtitle">Master data is maintained separately from daily activities '
        'so future workload, SDM and Gantt modules can reuse the same data.</div>',
        unsafe_allow_html=True,
    )

    tab1, tab2 = st.tabs(["Projects", "Staff"])

    with tab1:
        projects = get_projects()
        edited = st.data_editor(
            projects[
                ["id", "name", "project_type", "start_date", "target_finish",
                 "duration_months", "status", "lead", "project_size"]
            ],
            use_container_width=True,
            hide_index=True,
            num_rows="dynamic",
            column_config={
                "id": st.column_config.TextColumn("Project ID", required=True),
                "name": st.column_config.TextColumn("Project Name", required=True),
                "start_date": st.column_config.TextColumn("Start Date"),
                "target_finish": st.column_config.TextColumn("Target Finish"),
                "duration_months": st.column_config.NumberColumn("Duration (mo)"),
            },
            key="project_editor",
        )

        if st.button("Save Project Master", type="primary"):
            conn = get_conn()
            try:
                # Replace master table from editor. V1 keeps it simple.
                conn.execute("DELETE FROM projects")
                for _, r in edited.iterrows():
                    pid = clean(r["id"])
                    pname = clean(r["name"])
                    if not pid or not pname:
                        continue
                    conn.execute(
                        """INSERT INTO projects
                        (id,name,project_type,start_date,target_finish,duration_months,status,lead,project_size)
                        VALUES (?,?,?,?,?,?,?,?,?)""",
                        (
                            pid, pname, clean(r["project_type"]),
                            clean(r["start_date"]), clean(r["target_finish"]),
                            numeric_or_none(r["duration_months"]),
                            clean(r["status"]), clean(r["lead"]), clean(r["project_size"])
                        ),
                    )
                conn.commit()
                st.success("Project master saved.")
            except Exception as exc:
                conn.rollback()
                st.error(f"Could not save projects: {exc}")
            finally:
                conn.close()

    with tab2:
        staff = get_staff()
        edited = st.data_editor(
            staff[
                ["name", "category", "primary_role", "intern_start", "intern_end", "active"]
            ],
            use_container_width=True,
            hide_index=True,
            num_rows="dynamic",
            column_config={
                "name": st.column_config.TextColumn("Staff Name", required=True),
                "active": st.column_config.CheckboxColumn("Active"),
            },
            key="staff_editor",
        )

        if st.button("Save Staff Master", type="primary"):
            conn = get_conn()
            try:
                conn.execute("DELETE FROM staff")
                for _, r in edited.iterrows():
                    name = clean(r["name"])
                    if not name:
                        continue
                    conn.execute(
                        """INSERT INTO staff
                        (name,category,primary_role,intern_start,intern_end,active)
                        VALUES (?,?,?,?,?,?)""",
                        (
                            name, clean(r["category"]), clean(r["primary_role"]),
                            clean(r["intern_start"]), clean(r["intern_end"]),
                            int(bool(r["active"]))
                        ),
                    )
                conn.commit()
                st.success("Staff master saved.")
            except Exception as exc:
                conn.rollback()
                st.error(f"Could not save staff: {exc}")
            finally:
                conn.close()


def work_input():
    st.subheader("Work Activity")

    projects = get_projects()
    staff = get_staff()
    pids = projects["id"].tolist() if not projects.empty else []
    names = staff["name"].tolist() if not staff.empty else []

    activity_types = get_refs("activity_type") or ["Submission", "Internal Review"]
    priorities = get_refs("priority") or ["High", "Medium", "Low"]
    statuses = get_refs("activity_status") or ["Not Started", "In Progress", "Review", "Submitted", "Done"]

    with st.form("work_form", clear_on_submit=True):
        c1, c2, c3 = st.columns(3)
        with c1:
            pid = st.selectbox("Project", pids)
            start = st.date_input("Start Date", value=date.today())
        with c2:
            activity_type = st.selectbox("Activity Type", activity_types)
            end = st.date_input("End Date", value=date.today())
        with c3:
            task = st.text_input("Deliverable / Task")
            priority = st.selectbox("Priority", priorities)

        c4, c5 = st.columns(2)
        with c4:
            pic = st.selectbox("PIC", [""] + names)
        with c5:
            status = st.selectbox("Status", statuses)

        notes = st.text_area("Notes")
        submitted = st.form_submit_button("Add Work Activity", type="primary")

    if submitted:
        if not task.strip():
            st.error("Deliverable / Task wajib diisi.")
            return
        conn = get_conn()
        conn.execute(
            """INSERT INTO work_activity
            (project_id,start_date,end_date,activity_type,task,priority,pic,status,notes)
            VALUES (?,?,?,?,?,?,?,?,?)""",
            (
                pid, start.isoformat(), end.isoformat(), activity_type,
                task.strip(), priority, pic, status, notes
            ),
        )
        conn.commit()
        conn.close()
        st.success("Work activity added.")


def meeting_input():
    st.subheader("Meeting Activity")

    projects = get_projects()
    staff = get_staff()
    pids = projects["id"].tolist() if not projects.empty else []
    names = [""] + staff["name"].tolist()
    meeting_types = get_refs("meeting_type") or [
        "Internal Design", "Client Meeting", "Consultant Coordination",
        "Site Meeting", "Vendor Meeting"
    ]
    locations = get_refs("location") or ["Studio", "Online", "Client Office", "Site"]

    with st.form("meeting_form", clear_on_submit=True):
        c1, c2, c3 = st.columns(3)
        with c1:
            mdate = st.date_input("Date", value=date.today(), key="meeting_date")
            start = st.time_input("Start", value=time(9, 0))
        with c2:
            pid = st.selectbox("Project", [""] + pids, key="meeting_project")
            end = st.time_input("End", value=time(10, 0))
        with c3:
            meeting_type = st.selectbox("Meeting Type", meeting_types)
            location = st.selectbox("Location", locations)

        c = st.columns(4)
        attendees = []
        for i in range(4):
            with c[i]:
                attendees.append(st.selectbox(f"Attendee {i+1}", names, key=f"att_{i}"))

        notes = st.text_area("Agenda / Notes")
        submitted = st.form_submit_button("Add Meeting Activity", type="primary")

    if submitted:
        conn = get_conn()
        conn.execute(
            """INSERT INTO meeting_activity
            (activity_date,start_time,end_time,project_id,meeting_type,
             attendee_1,attendee_2,attendee_3,attendee_4,location,agenda_notes)
            VALUES (?,?,?,?,?,?,?,?,?,?,?)""",
            (
                mdate.isoformat(), start.strftime("%H:%M"), end.strftime("%H:%M"),
                pid or None, meeting_type,
                attendees[0], attendees[1], attendees[2], attendees[3],
                location, notes
            ),
        )
        conn.commit()
        conn.close()
        st.success("Meeting activity added.")


def other_input():
    st.subheader("Other Activities")

    staff = get_staff()
    names = [""] + staff["name"].tolist()

    with st.form("other_form", clear_on_submit=True):
        c1, c2 = st.columns(2)
        with c1:
            odate = st.date_input("Date", value=date.today(), key="other_date")
            activity = st.text_input("Other Activity")
        with c2:
            related = st.selectbox("Related Staff", names, key="other_staff")
            notes = st.text_area("Notes")

        submitted = st.form_submit_button("Add Other Activity", type="primary")

    if submitted:
        if not activity.strip():
            st.error("Other Activity wajib diisi.")
            return
        conn = get_conn()
        conn.execute(
            """INSERT INTO other_activity
            (activity_date,activity,related_staff,notes)
            VALUES (?,?,?,?)""",
            (odate.isoformat(), activity.strip(), related, notes),
        )
        conn.commit()
        conn.close()
        st.success("Other activity added.")


def activities_page():
    st.markdown('<div class="app-title">Daily Activities</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="app-subtitle">Daily input layer used by the Weekly Dashboard and future modules.</div>',
        unsafe_allow_html=True,
    )

    tabs = st.tabs(["Work", "Meeting", "Other"])
    with tabs[0]:
        work_input()
    with tabs[1]:
        meeting_input()
    with tabs[2]:
        other_input()


# ============================================================
# MAIN
# ============================================================

init_db()
seed_from_workbook()

if "seed_error" in st.session_state:
    st.warning(f"Workbook seed warning: {st.session_state['seed_error']}")

st.sidebar.markdown("## 📐 Studio Control Board")
st.sidebar.caption("V1 • Setup + Activities + Weekly Dashboard")

page = st.sidebar.radio(
    "MODULE",
    ["Weekly Dashboard", "Activities", "Setup"],
)

if page == "Weekly Dashboard":
    weekly_dashboard()
elif page == "Activities":
    activities_page()
else:
    setup_page()
