
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
    .app-title {font-size: 2rem; font-weight: 750; margin-bottom: 0.1rem;}
    .app-subtitle {color:#667085; margin-bottom:1rem;}
    .week-title {
        font-size: 1rem; font-weight: 750; padding: 0.55rem 0.7rem;
        border-radius: 8px; background: #F2F4F7; margin-top: 0.8rem;
    }
    .day-head {
        border-bottom: 1px solid #D0D5DD;
        padding: 0.45rem 0.3rem;
        font-weight: 700;
        text-align: center;
        background: #F8FAFC;
        min-height: 58px;
    }
    .day-head .dow {font-size: 0.72rem; color:#667085; text-transform:uppercase;}
    .day-head .day {font-size: 1rem; color:#101828;}
    .lane-label {
        height: 100%;
        min-height: 130px;
        padding: 0.65rem 0.45rem;
        font-weight: 800;
        font-size: 0.78rem;
        letter-spacing: .04em;
        border-right: 1px solid #D0D5DD;
        color:#344054;
    }
    .cell {
        min-height: 130px;
        padding: 0.45rem;
        border-right: 1px solid #EAECF0;
        border-bottom: 1px solid #EAECF0;
        background: white;
    }
    .project {
        font-weight: 750;
        color:#101828;
        margin: 0.15rem 0 0.35rem;
    }
    .task {font-size: 0.79rem; line-height: 1.35; margin: 0.22rem 0;}
    .task.submission {font-weight: 650;}
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
    .kpi {
        padding: 0.8rem 1rem;
        border: 1px solid #EAECF0;
        border-radius: 10px;
        background: white;
    }
    .kpi-label {font-size:.72rem;color:#667085;text-transform:uppercase;}
    .kpi-value {font-size:1.35rem;font-weight:750;color:#101828;}
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

    # Same weekly model as the Excel dashboard: first week starts on
    # the first day of the month; subsequent weeks are 7-day blocks.
    weeks = []
    current = first
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


def render_week(start_date, end_date, work, meetings, others):
    days = []
    d = start_date
    while d <= end_date:
        days.append(d)
        d += timedelta(days=1)

    work_groups = group_items(work, "end_date")
    meeting_groups = group_items(meetings, "activity_date")

    # Other grouping intentionally ignores project.
    other_groups = {}
    if not others.empty:
        for _, row in others.iterrows():
            d = parse_date(row["activity_date"])
            other_groups.setdefault(d, []).append(row)

    st.markdown(
        f'<div class="week-title">WEEK • {start_date.strftime("%d %b")} – '
        f'{end_date.strftime("%d %b %Y")}</div>',
        unsafe_allow_html=True,
    )

    # CSS grid: lane + up to 7 days.
    cols = st.columns(len(days) + 1, gap="small")
    with cols[0]:
        st.markdown('<div class="day-head">LANE</div>', unsafe_allow_html=True)
    for i, d in enumerate(days, start=1):
        with cols[i]:
            st.markdown(
                f'<div class="day-head"><div class="dow">{fmt_day(d)}</div>'
                f'<div class="day">{d.strftime("%d %b")}</div></div>',
                unsafe_allow_html=True,
            )

    lane_specs = [
        ("WORK", "work"),
        ("MEETING", "meeting"),
        ("OTHER", "other"),
    ]

    for lane_name, lane_type in lane_specs:
        cols = st.columns(len(days) + 1, gap="small")
        with cols[0]:
            st.markdown(
                f'<div class="lane-label">{lane_name}</div>',
                unsafe_allow_html=True,
            )

        for i, d in enumerate(days, start=1):
            with cols[i]:
                if lane_type == "work":
                    # All project groups for the day, sorted by project.
                    rows = []
                    for (gd, pid), items in work_groups.items():
                        if gd == d:
                            rows.extend(items)
                    st.markdown(
                        f'<div class="cell">{render_work_cell(rows)}</div>',
                        unsafe_allow_html=True,
                    )

                elif lane_type == "meeting":
                    rows = []
                    for (gd, pid), items in meeting_groups.items():
                        if gd == d:
                            rows.extend(items)
                    st.markdown(
                        f'<div class="cell">{render_meeting_cell(rows)}</div>',
                        unsafe_allow_html=True,
                    )

                else:
                    rows = other_groups.get(d, [])
                    st.markdown(
                        f'<div class="cell">{render_other_cell(rows)}</div>',
                        unsafe_allow_html=True,
                    )


def weekly_dashboard():
    st.markdown('<div class="app-title">Weekly Schedule Dashboard</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="app-subtitle">Leadership view • Work & Meeting follow project filter • '
        'Other remains independent of project filter</div>',
        unsafe_allow_html=True,
    )

    projects = get_projects()
    project_options = ["All Projects"] + projects["id"].tolist() if not projects.empty else ["All Projects"]

    c1, c2, c3 = st.columns([1.1, 1.5, 1.0])
    with c1:
        month_options = []
        # Build months from projects/activity horizon + current year around it.
        conn = get_conn()
        dates = []
        for table, col in [
            ("work_activity", "end_date"),
            ("meeting_activity", "activity_date"),
            ("other_activity", "activity_date"),
        ]:
            q = f"SELECT MIN({col}), MAX({col}) FROM {table}"
            r = conn.execute(q).fetchone()
            if r[0]:
                dates.append(parse_date(r[0]))
            if r[1]:
                dates.append(parse_date(r[1]))
        conn.close()

        if dates:
            min_d, max_d = min(dates), max(dates)
        else:
            today = date.today()
            min_d, max_d = date(today.year, today.month, 1), today

        months = []
        cur = date(min_d.year, min_d.month, 1)
        while cur <= max_d:
            months.append(cur)
            if cur.month == 12:
                cur = date(cur.year + 1, 1, 1)
            else:
                cur = date(cur.year, cur.month + 1, 1)

        month_labels = [m.strftime("%b %Y") for m in months]
        default_idx = max(0, len(month_labels) - 1)
        selected_label = st.selectbox("MONTH", month_labels, index=default_idx)
        selected_month = months[month_labels.index(selected_label)]

    with c2:
        selected_project = st.selectbox("PROJECT", project_options)

    with c3:
        weeks = month_weeks(selected_month.year, selected_month.month)
        week_options = ["All"] + [f"Week {i+1}" for i in range(len(weeks))]
        selected_week = st.selectbox("WEEK", week_options)

    selected_weeks = list(enumerate(weeks, start=1))
    if selected_week != "All":
        idx = int(selected_week.split()[-1])
        selected_weeks = [selected_weeks[idx - 1]]

    month_start = selected_month
    month_end = date(
        selected_month.year,
        selected_month.month,
        calendar.monthrange(selected_month.year, selected_month.month)[1],
    )

    work, meetings, others = load_activities(
        month_start, month_end, selected_project
    )

    # KPI strip
    work_count = len(work)
    meeting_count = len(meetings)
    conflict_count = 0  # V1 UI keeps source conflict fields out; can be added later.
    submission_count = (
        int((work["activity_type"].str.lower() == "submission").sum())
        if not work.empty else 0
    )

    k = st.columns(4)
    for col, label, value in [
        (k[0], "WORK ITEMS", work_count),
        (k[1], "MEETINGS", meeting_count),
        (k[2], "SUBMISSIONS", submission_count),
        (k[3], "OTHER ACTIVITIES", len(others)),
    ]:
        with col:
            st.markdown(
                f'<div class="kpi"><div class="kpi-label">{label}</div>'
                f'<div class="kpi-value">{value}</div></div>',
                unsafe_allow_html=True,
            )

    st.write("")

    for _, (week_start, week_end) in selected_weeks:
        # Clip week data to actual week.
        w = work[
            (work["end_date"] >= week_start.isoformat())
            & (work["end_date"] <= week_end.isoformat())
        ].copy()
        m = meetings[
            (meetings["activity_date"] >= week_start.isoformat())
            & (meetings["activity_date"] <= week_end.isoformat())
        ].copy()
        o = others[
            (others["activity_date"] >= week_start.isoformat())
            & (others["activity_date"] <= week_end.isoformat())
        ].copy()

        render_week(week_start, week_end, w, m, o)


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
