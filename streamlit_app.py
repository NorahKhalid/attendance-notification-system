
import streamlit as st
import pandas as pd
import sqlite3, re, json
from pathlib import Path
from datetime import datetime, date, time
from urllib.parse import quote

APP = Path(__file__).resolve().parent
DB = APP / "attendance_system.db"

st.set_page_config(
    page_title="Attendance Notification System V12",
    page_icon="📧",
    layout="wide"
)

# -------------------- DATABASE --------------------
def conn():
    c = sqlite3.connect(DB, check_same_thread=False)
    c.row_factory = sqlite3.Row
    return c

def init_db():
    c = conn()
    c.executescript("""
    CREATE TABLE IF NOT EXISTS employees(
        emp_no TEXT PRIMARY KEY,
        name TEXT,
        email TEXT,
        category TEXT DEFAULT 'full',
        works_saturday INTEGER DEFAULT 1,
        hire_date TEXT,
        active INTEGER DEFAULT 1,
        project TEXT DEFAULT '',
        project_priority TEXT DEFAULT 'normal',
        updated_at TEXT
    );
    CREATE TABLE IF NOT EXISTS notes(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        emp_no TEXT,
        note TEXT,
        active INTEGER DEFAULT 1,
        created_at TEXT,
        updated_at TEXT
    );
    CREATE TABLE IF NOT EXISTS settings(key TEXT PRIMARY KEY,value TEXT);
    CREATE TABLE IF NOT EXISTS history(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        emp_no TEXT,
        email TEXT,
        subject TEXT,
        body TEXT,
        status TEXT,
        created_at TEXT
    );
    """)
    cols = {r["name"] for r in c.execute("PRAGMA table_info(employees)").fetchall()}
    additions = [
        ("project", "TEXT DEFAULT ''"),
        ("project_priority", "TEXT DEFAULT 'normal'"),
        ("works_saturday", "INTEGER DEFAULT 1"),
        ("hire_date", "TEXT"),
        ("category", "TEXT DEFAULT 'full'"),
        ("active", "INTEGER DEFAULT 1"),
        ("updated_at", "TEXT"),
    ]
    for n, t in additions:
        if n not in cols:
            c.execute(f"ALTER TABLE employees ADD COLUMN {n} {t}")

    c.executescript("""
    CREATE TABLE IF NOT EXISTS archive_months(
        month_key TEXT PRIMARY KEY,
        month_label TEXT,
        saved_at TEXT,
        total_cases INTEGER DEFAULT 0,
        total_employees INTEGER DEFAULT 0
    );
    CREATE TABLE IF NOT EXISTS archive_cases(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        month_key TEXT,
        emp_no TEXT, name TEXT, email TEXT, project TEXT, project_priority TEXT,
        date TEXT, case_type TEXT, duration TEXT, worked_hours REAL, expected_hours REAL,
        reason TEXT, note TEXT, excluded INTEGER DEFAULT 0, first_day INTEGER DEFAULT 0,
        sent INTEGER DEFAULT 0, sent_at TEXT, reply_note TEXT,
        subject TEXT, body TEXT, created_at TEXT,
        UNIQUE(month_key, emp_no, date, case_type)
    );
    """)

    defaults = {
        "subject_ar": "إشعار حضور - {month}",
        "subject_en": "Attendance Notification - {month}",
        "body_ar": """السلام عليكم {name}،

نود إشعاركم بوجود سجل حضور/انصراف في الفترة {period}.

التفاصيل:
{details}

في حال وجود إجازة أو استئذان أو مبرر لم يتم تسجيله، نرجو تزويد الموارد البشرية بالتوضيح.

مع الشكر،
قسم الموارد البشرية""",
        "body_en": """Dear {name},

This is to notify you of an attendance record for {period}.

Details:
{details}

If you have an approved leave, permission, or another valid justification, please provide it to HR.

Regards,
HR Department"""
    }
    for k, v in defaults.items():
        c.execute("INSERT OR IGNORE INTO settings(key,value) VALUES(?,?)", (k, v))
    c.commit()
    c.close()

init_db()

def setting(k, default=""):
    c = conn()
    r = c.execute("SELECT value FROM settings WHERE key=?", (k,)).fetchone()
    c.close()
    return r["value"] if r else default

def set_setting(k, v):
    c = conn()
    c.execute("INSERT OR REPLACE INTO settings(key,value) VALUES(?,?)", (k, str(v)))
    c.commit()
    c.close()

def employees():
    c = conn()
    d = pd.read_sql_query("SELECT * FROM employees ORDER BY name", c)
    c.close()
    return d

def active_notes():
    c = conn()
    d = pd.read_sql_query("""
        SELECT n.*, e.name
        FROM notes n
        LEFT JOIN employees e ON e.emp_no=n.emp_no
        WHERE n.active=1
        ORDER BY n.updated_at DESC
    """, c)
    c.close()
    return d

def note_for(no):
    c = conn()
    rows = c.execute("""
        SELECT note FROM notes
        WHERE emp_no=? AND active=1
        ORDER BY updated_at DESC
    """, (str(no),)).fetchall()
    c.close()
    return " | ".join(x["note"] for x in rows)

def save_emps(df):
    """Upsert employee data without overwriting user-maintained rules."""
    c = conn()
    now = datetime.now().isoformat(timespec="seconds")
    for _, r in df.iterrows():
        no = str(r.get("emp_no", "")).strip()
        if not no:
            continue
        old = c.execute("SELECT * FROM employees WHERE emp_no=?", (no,)).fetchone()

        name = str(r.get("name", "")).strip()
        email = str(r.get("email", "")).strip()
        hire = str(r.get("hire_date", "")).strip()
        project = str(r.get("project", "")).strip()

        if old:
            # Only overwrite fields that the uploaded master actually supplies.
            if not name or name.lower() in {"nan", "none"}:
                name = old["name"] or ""
            if not email or email.lower() in {"nan", "none"}:
                email = old["email"] or ""
            if not hire or hire.lower() in {"nan", "nat"}:
                hire = old["hire_date"] or ""
            if not project or project.lower() in {"nan", "none"}:
                project = old["project"] or ""

            vals = (
                no, name, email,
                old["category"] or "full",
                int(old["works_saturday"] if old["works_saturday"] is not None else 1),
                hire,
                int(old["active"] if old["active"] is not None else 1),
                project,
                old["project_priority"] or "normal",
                now
            )
            c.execute("""
                UPDATE employees SET
                    name=?, email=?, category=?, works_saturday=?, hire_date=?,
                    active=?, project=?, project_priority=?, updated_at=?
                WHERE emp_no=?
            """, vals[1:] + (no,))
        else:
            c.execute("""
                INSERT INTO employees
                (emp_no,name,email,category,works_saturday,hire_date,active,
                 project,project_priority,updated_at)
                VALUES(?,?,?,?,?,?,?,?,?,?)
            """, (
                no, name, email, "full", 1, hire, 1, project, "normal", now
            ))
    c.commit()
    c.close()

# -------------------- EXCEL HELPERS --------------------
def clean(x):
    return re.sub(r"\s+", " ", str(x)).strip().lower()

def read_excel_flexible(uploaded, candidates):
    raw = pd.read_excel(uploaded, header=None)
    best_row, best_score = 0, -1
    for i in range(min(30, len(raw))):
        vals = [clean(v) for v in raw.iloc[i].tolist()]
        score = sum(
            any(clean(c) == v or clean(c) in v for v in vals)
            for c in candidates
        )
        if score > best_score:
            best_row, best_score = i, score

    d = pd.read_excel(uploaded, header=best_row).dropna(how="all").copy()
    seen, labels = {}, []
    for c in d.columns:
        b = str(c).strip()
        n = seen.get(b, 0)
        labels.append(b if n == 0 else f"{b}__dup{n}")
        seen[b] = n + 1
    d.columns = labels
    return d

def ci(d, names):
    wanted = [clean(x) for x in names]
    for i, c in enumerate(d.columns):
        if clean(c) in wanted:
            return i
    for i, c in enumerate(d.columns):
        cc = clean(c)
        if any(w in cc or cc in w for w in wanted):
            return i
    return None

def ser(d, i):
    return pd.Series([""] * len(d), index=d.index, dtype="object") if i is None else d.iloc[:, i]

def empno(x):
    if pd.isna(x):
        return ""
    s = str(x).strip()
    return s[:-2] if re.fullmatch(r"\d+\.0", s) else s

def dt(x):
    return pd.to_datetime(x, errors="coerce")

def clean_text_series(s):
    return s.astype(str).replace({"nan": "", "None": "NaT", "NaT": ""}).str.strip()

def norm_emps(f):
    d = read_excel_flexible(f, [
        "Employment Number", "Employee Name", "Email", "Project",
        "رقم الموظف", "اسم الموظف", "البريد الإلكتروني", "المشروع"
    ])
    o = pd.DataFrame(index=d.index)
    o["emp_no"] = ser(d, ci(d, [
        "Employment Number", "Emp. No.", "Employee Number",
        "رقم الموظف", "الرقم الوظيفي"
    ])).map(empno)
    o["name"] = clean_text_series(ser(d, ci(d, [
        "Employee Name", "Emp. Name", "اسم الموظف"
    ])))
    o["email"] = clean_text_series(ser(d, ci(d, [
        "Email", "Work Email", "Company Email", "البريد الإلكتروني"
    ])))
    h = ci(d, ["Hiring Date", "Hire Date", "تاريخ المباشرة"])
    o["hire_date"] = (
        ser(d, h).map(dt).dt.strftime("%Y-%m-%d").fillna("")
        if h is not None else ""
    )
    p = ci(d, ["Project", "Project Name", "المشروع"])
    o["project"] = clean_text_series(ser(d, p)) if p is not None else ""
    return o[o.emp_no.ne("")].drop_duplicates("emp_no")

def norm_att(f):
    d = read_excel_flexible(f, [
        "Emp. No.", "Emp. Name", "Attendance Date", "Transaction Type",
        "Employee Number", "Employment Number", "تاريخ الحضور"
    ])
    o = pd.DataFrame(index=d.index)
    o["emp_no"] = ser(d, ci(d, [
        "Emp. No.", "Employment Number", "Employee Number",
        "رقم الموظف", "الرقم الوظيفي"
    ])).map(empno)
    o["name"] = clean_text_series(ser(d, ci(d, [
        "Emp. Name", "Employee Name", "اسم الموظف"
    ])))
    o["date"] = ser(d, ci(d, [
        "Attendance Date", "Date", "تاريخ الحضور", "التاريخ"
    ])).map(dt)
    o["event"] = clean_text_series(ser(d, ci(d, [
        "Transaction Type", "Type", "نوع الحركة", "نوع العملية", "الحالة"
    ])))
    o["duration"] = clean_text_series(ser(d, ci(d, [
        "Duration", "Late Duration", "مدة التأخير", "التأخير", "مدة"
    ])))
    o["entry"] = clean_text_series(ser(d, ci(d, [
        "Entry Time", "Check In", "Clock In", "وقت الدخول", "دخول"
    ])))
    o["exit"] = clean_text_series(ser(d, ci(d, [
        "Exit Time", "Check Out", "Clock Out", "وقت الخروج", "خروج"
    ])))
    return o[o.emp_no.ne("") & o.date.notna()].copy()

def norm_range(f):
    if not f:
        return pd.DataFrame()
    d = read_excel_flexible(f, [
        "Emp. No.", "Employment Number", "From Date", "To Date", "Status"
    ])
    o = pd.DataFrame(index=d.index)
    o["emp_no"] = ser(d, ci(d, [
        "Emp. No.", "Employment Number", "Employee Number", "رقم الموظف"
    ])).map(empno)
    a = ci(d, ["From Date", "Start Date", "من تاريخ"])
    b = ci(d, ["To Date", "End Date", "إلى تاريخ"])
    q = ci(d, ["Date", "Attendance Date", "التاريخ"])
    o["from"] = ser(d, a).map(dt) if a is not None else ser(d, q).map(dt)
    o["to"] = ser(d, b).map(dt) if b is not None else o["from"]
    o["status"] = clean_text_series(ser(d, ci(d, ["Status", "الحالة"]))).str.lower()
    return o[o.emp_no.ne("")].copy()

def norm_fingerprint(f):
    if not f:
        return pd.DataFrame()
    d = read_excel_flexible(f, [
        "Employment Number", "Attendance Date", "Date Requested",
        "Status", "Type", "Fingerprint"
    ])
    o = pd.DataFrame(index=d.index)
    o["emp_no"] = ser(d, ci(d, [
        "Employment Number", "Emp. No.", "Employee Number", "رقم الموظف"
    ])).map(empno)
    o["date"] = ser(d, ci(d, [
        "Attendance Date", "Date", "Date Requested", "التاريخ"
    ])).map(dt)
    o["status"] = clean_text_series(ser(d, ci(d, [
        "Status", "الحالة", "Type", "Fingerprint", "نوع"
    ])))
    return o[o.emp_no.ne("") & o.date.notna()].copy()

# -------------------- BUSINESS RULES --------------------
def classify(event):
    s = str(event).lower()
    if any(k in s for k in ["absence", "absent", "غياب", "غائب"]):
        return "غياب"
    if any(k in s for k in [
        "early leave", "early exit", "early departure",
        "early", "خروج مبكر", "مغادرة مبكرة"
    ]):
        return "خروج مبكر"
    if any(k in s for k in [
        "late", "tardy", "delay", "تأخير", "متأخر"
    ]):
        return "تأخير صباحي"
    if any(k in s for k in ["departure", "leaving", "leave", "مغادرة"]):
        return "مغادرة"
    if any(k in s for k in [
        "fingerprint", "finger print", "missing fingerprint",
        "missed fingerprint", "نسيان البصمة", "بصمة مفقودة", "بصمة"
    ]):
        return "نسيان بصمة"
    return str(event).strip() or "حالة حضور"

def fingerprint_type(status):
    s = str(status).lower()
    if any(k in s for k in ["out", "exit", "check out", "clock out", "خروج"]):
        return "نسيان بصمة خروج"
    if any(k in s for k in ["in", "entry", "check in", "clock in", "دخول"]):
        return "نسيان بصمة دخول"
    return "نسيان بصمة"

def approved(s):
    return any(k in str(s).lower() for k in [
        "approved", "موافق", "اعتمد", "مقبول", "معتمد"
    ])

def parse_clock(v):
    if v is None or str(v).strip() == "":
        return None
    s = str(v).strip()
    # Try common textual times first.
    for fmt in ("%I:%M %p", "%I:%M:%S %p", "%H:%M", "%H:%M:%S"):
        try:
            return datetime.strptime(s, fmt).time()
        except Exception:
            pass
    try:
        t = pd.to_datetime(s, errors="coerce")
        return None if pd.isna(t) else t.time()
    except Exception:
        return None

def worked_hours(entry, exit_):
    a, b = parse_clock(entry), parse_clock(exit_)
    if not a or not b:
        return None
    start = datetime.combine(date.today(), a)
    end = datetime.combine(date.today(), b)
    if end < start:
        end += pd.Timedelta(days=1)
    return round((end - start).total_seconds() / 3600, 2)

def expected_hours(works_saturday, weekday):
    # Saturday workers: 8h every working day.
    # Non-Saturday workers: 9h Sunday-Thursday.
    # Friday is treated as overtime / special day, not a normal early-leave day.
    if weekday == "Friday":
        return None
    return 8.0 if int(works_saturday or 0) else 9.0

def analyze(attf, leavef, permf, fpf):
    a = norm_att(attf)
    l = norm_range(leavef)
    p = norm_range(permf)
    f = norm_fingerprint(fpf)
    e = employees()

    if e.empty:
        z = a[["emp_no", "name"]].drop_duplicates("emp_no").copy()
        z["email"] = ""
        z["hire_date"] = ""
        z["project"] = ""
        save_emps(z)
        e = employees()

    defaults = {
        "category": "full", "works_saturday": 1, "hire_date": "",
        "active": 1, "project": "", "project_priority": "normal"
    }
    for col, default in defaults.items():
        if col not in e.columns:
            e[col] = default

    a["case_type"] = a["event"].map(classify)
    a = a.merge(
        e[[
            "emp_no", "name", "email", "category", "works_saturday",
            "hire_date", "active", "project", "project_priority"
        ]],
        on="emp_no", how="left", suffixes=("", "_m")
    )

    if "name_m" in a.columns:
        a["name"] = a["name_m"].where(
            a["name_m"].notna() & a["name_m"].ne(""), a["name"]
        )

    a = a[a["active"].fillna(1).astype(bool)].copy()

    reasons, excluded, first_day, day_flags = [], [], [], []
    work_hrs, expected_hrs = [], []

    for _, r in a.iterrows():
        no = str(r.emp_no)
        d = pd.Timestamp(r.date).normalize()
        rs = []
        fd = False

        if str(r.category) == "exclude":
            rs.append("موظف مستثنى")

        weekday = d.day_name()
        if weekday == "Saturday" and int(r.works_saturday or 0) == 0:
            rs.append("لا يعمل السبت")

        if not l.empty:
            ok = l[
                (l.emp_no.astype(str) == no) &
                l["from"].notna() & l["to"].notna() &
                (l["from"].dt.normalize() <= d) &
                (l["to"].dt.normalize() >= d) &
                l.status.map(approved)
            ]
            if not ok.empty:
                rs.append("إجازة معتمدة")

        if not p.empty:
            ok = p[
                (p.emp_no.astype(str) == no) &
                p["from"].notna() & p["to"].notna() &
                (p["from"].dt.normalize() <= d) &
                (p["to"].dt.normalize() >= d) &
                p.status.map(approved)
            ]
            if not ok.empty:
                rs.append("استئذان معتمد")

        hd = str(r.hire_date or "")
        if hd and hd not in ["nan", "NaT"]:
            try:
                fd = pd.Timestamp(hd).normalize() == d
            except Exception:
                fd = False
        if fd:
            rs.append("أول يوم عمل")

        wh = worked_hours(r.entry, r.exit)
        eh = expected_hours(r.works_saturday, weekday)

        # Friday is always visibly flagged but not treated as a normal early-leave day.
        if weekday == "Friday":
            day_flags.append("جمعة / Friday — Overtime")
        elif weekday == "Saturday":
            day_flags.append("سبت / Saturday")
        else:
            day_flags.append("")

        work_hrs.append(wh)
        expected_hrs.append(eh)

        # If the report has clock-in/out and the row is a generic departure,
        # classify it as early leave only when worked hours are below the
        # employee's schedule. This fixes the 8h/9h Saturday-worker rule.
        case = r.case_type
        if case == "مغادرة" and wh is not None and eh is not None and wh >= eh:
            case = "حضور مكتمل"
        a.loc[r.name, "case_type"] = case

        reasons.append("; ".join(rs))
        excluded.append(bool(rs))
        first_day.append(fd)

    a["reason"] = reasons
    a["excluded"] = excluded
    a["first_day"] = first_day
    a["day_flag"] = day_flags
    a["worked_hours"] = work_hrs
    a["expected_hours"] = expected_hrs
    a["note"] = a.emp_no.map(note_for)

    # Add fingerprint records that are not already represented in attendance.
    if not f.empty:
        existing = set(zip(
            a.emp_no.astype(str),
            a.date.dt.normalize()
        ))
        extra = []
        for _, r in f.iterrows():
            key = (str(r.emp_no), r.date.normalize())
            typ = fingerprint_type(r.status)
            if key in existing:
                # Improve the existing row's case type if it is a fingerprint event.
                mask = (
                    a.emp_no.astype(str).eq(str(r.emp_no)) &
                    a.date.dt.normalize().eq(r.date.normalize())
                )
                a.loc[mask, "case_type"] = typ
                continue

            emp = e[e.emp_no.astype(str) == str(r.emp_no)]
            if emp.empty:
                continue
            rr = emp.iloc[0]
            weekday = r.date.day_name()
            extra.append({
                "emp_no": str(r.emp_no),
                "name": rr["name"],
                "email": rr["email"],
                "date": r.date,
                "event": r.status,
                "duration": "",
                "entry": "",
                "exit": "",
                "case_type": typ,
                "category": rr["category"],
                "works_saturday": rr["works_saturday"],
                "hire_date": rr["hire_date"],
                "active": rr["active"],
                "project": rr["project"],
                "project_priority": rr["project_priority"],
                "reason": "نسيان البصمة",
                "excluded": False,
                "first_day": False,
                "day_flag": (
                    "جمعة / Friday — Overtime" if weekday == "Friday"
                    else "سبت / Saturday" if weekday == "Saturday" else ""
                ),
                "worked_hours": None,
                "expected_hours": expected_hours(rr["works_saturday"], weekday),
                "note": note_for(str(r.emp_no))
            })
        if extra:
            a = pd.concat([a, pd.DataFrame(extra)], ignore_index=True)

    meta = {
        "attendance_rows": len(a),
        "absence_rows": int((a.case_type == "غياب").sum()),
        "late_rows": int((a.case_type == "تأخير صباحي").sum()),
        "early_leave_rows": int((a.case_type == "خروج مبكر").sum()),
        "departure_rows": int((a.case_type == "مغادرة").sum()),
        "fingerprint_rows": int(a.case_type.str.contains("نسيان بصمة", na=False).sum()),
        "excluded": int(a.excluded.sum()),
        "first_day_rows": int(a.first_day.sum()),
        "friday_rows": int(a.day_flag.str.startswith("جمعة", na=False).sum()),
    }
    return a.reset_index(drop=True), meta

def month_key_from_df(df):
    if df.empty or "date" not in df.columns:
        return "", ""
    d = pd.to_datetime(df["date"], errors="coerce").dropna()
    if d.empty:
        return "", ""
    first = d.min()
    return first.strftime("%Y-%m"), first.strftime("%B %Y")

def archive_month(df):
    if df.empty:
        return "", "", 0
    key, label = month_key_from_df(df)
    if not key:
        return "", "", 0
    c = conn(); now=datetime.now().isoformat(timespec="seconds")
    # Preserve manual send/reply tracking before replacing the analysis snapshot.
    send_status = {}
    old = c.execute("SELECT emp_no,date,case_type,sent,sent_at,reply_note FROM archive_cases WHERE month_key=?", (key,)).fetchall()
    for r in old:
        send_status[(str(r["emp_no"]), str(r["date"]), str(r["case_type"]))] = (r["sent"],r["sent_at"],r["reply_note"])
    # Replace snapshot for this month so saving again reflects the latest corrected analysis.
    c.execute("DELETE FROM archive_cases WHERE month_key=?", (key,))
    rows=[]
    df = df.copy().sort_values(["emp_no","date"])
    df = df.drop_duplicates(subset=["emp_no","date","case_type"], keep="last")
    for _,r in df.iterrows():
        date_s=pd.Timestamp(r.date).strftime("%Y-%m-%d")
        k=(str(r.emp_no),date_s,str(r.case_type))
        sent,sent_at,reply=send_status.get(k,(0,"",""))
        rows.append((key,str(r.emp_no),str(r.name),str(r.email),str(r.project),str(r.project_priority),
                     date_s,str(r.case_type),str(r.duration),float(r.worked_hours) if pd.notna(r.worked_hours) else None,
                     float(r.expected_hours) if pd.notna(r.expected_hours) else None,str(r.reason),str(r.note),
                     int(bool(r.excluded)),int(bool(r.first_day)),sent,sent_at,reply,"","",now))
    c.executemany("""INSERT INTO archive_cases
        (month_key,emp_no,name,email,project,project_priority,date,case_type,duration,worked_hours,expected_hours,
         reason,note,excluded,first_day,sent,sent_at,reply_note,subject,body,created_at)
        VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",rows)
    c.execute("INSERT OR REPLACE INTO archive_months(month_key,month_label,saved_at,total_cases,total_employees) VALUES(?,?,?,?,?)",
              (key,label,now,len(rows),len(set(x[1] for x in rows))))
    c.commit(); c.close()
    return key,label,len(rows)

def archive_df(month_key):
    c=conn(); d=pd.read_sql_query("SELECT * FROM archive_cases WHERE month_key=? ORDER BY emp_no,date",c,params=(month_key,)); c.close(); return d

def update_archive_case(case_id, sent=None, reply_note=None):
    c=conn();
    if sent is not None:
        c.execute("UPDATE archive_cases SET sent=?, sent_at=? WHERE id=?",(int(sent),datetime.now().isoformat(timespec="seconds") if sent else "",int(case_id)))
    if reply_note is not None:
        c.execute("UPDATE archive_cases SET reply_note=? WHERE id=?",(str(reply_note),int(case_id)))
    c.commit(); c.close()

def archive_excel_bytes(df):
    from io import BytesIO
    out=BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as w:
        df.to_excel(w,index=False,sheet_name="Archive")
    return out.getvalue()

def gmail(to, subject, body):
    return (
        "https://mail.google.com/mail/?view=cm&fs=1"
        f"&to={quote(str(to))}&su={quote(str(subject))}&body={quote(str(body))}"
    )

# -------------------- SESSION STATE --------------------
if "analysis" not in st.session_state:
    st.session_state.analysis = pd.DataFrame()
if "meta" not in st.session_state:
    st.session_state.meta = {}

# -------------------- UI --------------------
st.sidebar.title("📧 Attendance Notification System")
st.sidebar.caption("V12 — Monthly Archive & Employee Grouping")

page = st.sidebar.radio("القائمة", [
    "الرئيسية", "التقارير والتحليل", "الموظفون والمشاريع",
    "الملاحظات", "القواعد والقوالب", "مراجعة المسودات",
    "السجل", "الأرشيف الشهري", "نسخة احتياطية"
])

if page == "الرئيسية":
    st.title("📊 Dashboard")
    e = employees()
    r = st.session_state.analysis
    m = st.session_state.meta

    a, b, c, d = st.columns(4)
    a.metric("الموظفون", len(e))
    b.metric("الحالات", len(r))
    c.metric("المستبعد", int(r.excluded.sum()) if not r.empty else 0)
    d.metric("آخر تشغيل", setting("last_run") or "—")

    if not r.empty:
        st.subheader("ملخص الحالات")
        cards = [
            ("غياب", "absence_rows"),
            ("تأخير", "late_rows"),
            ("خروج مبكر", "early_leave_rows"),
            ("مغادرة", "departure_rows"),
            ("نسيان بصمة", "fingerprint_rows"),
            ("أول يوم", "first_day_rows"),
        ]
        cols = st.columns(6)
        for i, (label, key) in enumerate(cards):
            cols[i].metric(label, m.get(key, 0))

        st.subheader("فلترة الحالات")
        x, y, z = st.columns(3)
        typ = x.selectbox(
            "نوع الحالة",
            ["الكل"] + sorted(r.case_type.dropna().unique().tolist())
        )
        projects = sorted([
            p for p in r.project.fillna("").astype(str).unique() if p.strip()
        ])
        proj = y.selectbox("المشروع", ["الكل"] + projects)
        priority = z.selectbox(
            "أولوية المشروع",
            ["الكل", "high", "medium", "normal"]
        )

        v = r.copy()
        if typ != "الكل":
            v = v[v.case_type == typ]
        if proj != "الكل":
            v = v[v.project == proj]
        if priority != "الكل":
            v = v[v.project_priority == priority]

        st.dataframe(
            v[[
                "emp_no", "name", "email", "project", "project_priority",
                "date", "case_type", "duration", "worked_hours",
                "expected_hours", "day_flag", "reason", "note",
                "excluded", "first_day"
            ]],
            use_container_width=True
        )
    else:
        st.info("ابدئي من صفحة التقارير وارفعِي ملفات الشهر ثم اضغطي تحليل.")

elif page == "التقارير والتحليل":
    st.title("📂 التقارير والتحليل")
    st.caption(
        "بعد التحليل تبقى النتيجة محفوظة أثناء التنقل بين الصفحات. "
        "يمكنك رفع ملف الموظفين/الإيميلات لاحقًا بدون فقد إعدادات المشاريع والقواعد."
    )

    ef = st.file_uploader(
        "1) ملف الموظفين والإيميلات والمشاريع — اختياري",
        type=["xlsx"], key="ef"
    )
    if ef and st.button("تحديث دليل الموظفين"):
        try:
            save_emps(norm_emps(ef))
            st.success("تم تحديث دليل الموظفين بدون تغيير إعدادات السبت والفئات وأولوية المشاريع.")
        except Exception as ex:
            st.error("تعذر تحديث ملف الموظفين.")
            st.exception(ex)

    af = st.file_uploader("2) تقرير الحضور — الشهر كامل", type=["xlsx"], key="af")
    lf = st.file_uploader("3) تقرير الإجازات", type=["xlsx"], key="lf")
    pf = st.file_uploader("4) تقرير الاستئذان", type=["xlsx"], key="pf")
    ff = st.file_uploader("5) تقرير نسيان البصمة", type=["xlsx"], key="ff")

    if st.button("🔎 تحليل التقارير", disabled=not af):
        try:
            r, m = analyze(af, lf, pf, ff)
            st.session_state.analysis = r
            st.session_state.meta = m
            set_setting("last_run", datetime.now().strftime("%Y-%m-%d %H:%M"))
            st.success("تم التحليل بنجاح.")
        except Exception as ex:
            st.error("تعذر تحليل التقرير. تأكدي من أن الملفات هي تقارير ZenHR نفسها.")
            st.exception(ex)

    r = st.session_state.analysis
    if not r.empty:
        st.subheader("نتيجة التحليل الحالية")
        types = st.multiselect(
            "إظهار الحالات",
            sorted(r.case_type.dropna().unique().tolist()),
            default=sorted(r.case_type.dropna().unique().tolist())
        )
        v = r[r.case_type.isin(types)].copy()
        st.dataframe(v, use_container_width=True)
        st.download_button(
            "⬇️ تحميل النتيجة CSV",
            v.to_csv(index=False).encode("utf-8-sig"),
            "attendance_analysis.csv",
            "text/csv"
        )
        st.divider()
        st.subheader("💾 حفظ الشهر")
        st.warning("احفظي الشهر بعد التأكد من أن التحليل صحيح. إذا أعدتِ الحفظ لاحقًا لنفس الشهر، سيتم تحديث نسخة الشهر بأحدث تحليل.")
        if st.button("💾 حفظ الشهر في الأرشيف",type="primary"):
            try:
                mk,label,count=archive_month(r)
                if mk:
                    st.success(f"تم حفظ {label}: {count} حالة. يمكنك الرجوع لها من «الأرشيف الشهري».")
                else:
                    st.error("لم أستطع تحديد شهر التقرير.")
            except Exception as ex:
                st.error("تعذر حفظ الأرشيف.")
                st.exception(ex)

elif page == "الموظفون والمشاريع":
    st.title("👥 الموظفون والمشاريع")
    e = employees()

    if not e.empty:
        ed = st.data_editor(
            e,
            use_container_width=True,
            num_rows="dynamic",
            column_config={
                "category": st.column_config.SelectboxColumn(
                    "الفئة",
                    options=[
                        "full", "absence_only", "exclude",
                        "half_day", "special", "trainee"
                    ]
                ),
                "works_saturday": st.column_config.CheckboxColumn(
                    "يعمل السبت"
                ),
                "active": st.column_config.CheckboxColumn("نشط"),
                "project_priority": st.column_config.SelectboxColumn(
                    "أولوية المشروع",
                    options=["high", "medium", "normal"]
                )
            }
        )
        if st.button("💾 حفظ التعديلات"):
            save_emps(ed)
            st.success("تم حفظ الموظفين والمشاريع.")

    st.subheader("إضافة موظف")
    with st.form("new"):
        a, b, c, d = st.columns(4)
        no = a.text_input("الرقم الوظيفي")
        name = b.text_input("الاسم")
        email = c.text_input("الإيميل")
        project = d.text_input("المشروع")
        hire = st.date_input("تاريخ المباشرة", date.today())
        cat = st.selectbox(
            "الفئة",
            ["full", "absence_only", "exclude", "half_day", "special", "trainee"]
        )
        sat = st.checkbox("يعمل السبت", True)
        pri = st.selectbox("أولوية المشروع", ["normal", "medium", "high"])
        if st.form_submit_button("إضافة") and no:
            save_emps(pd.DataFrame([{
                "emp_no": no,
                "name": name,
                "email": email,
                "project": project,
                "hire_date": hire.isoformat()
            }]))
            c2 = conn()
            c2.execute("""
                UPDATE employees
                SET category=?, works_saturday=?, project_priority=?
                WHERE emp_no=?
            """, (cat, int(sat), pri, no))
            c2.commit()
            c2.close()
            st.success("تمت إضافة الموظف.")

elif page == "الملاحظات":
    st.title("📝 الملاحظات")
    e = employees()
    if e.empty:
        st.warning("أضيفي الموظفين أولًا.")
    else:
        with st.form("note"):
            emp = st.selectbox(
                "الموظف",
                [f"{r.emp_no} — {r.name}" for _, r in e.iterrows()]
            )
            txt = st.text_area("الملاحظة")
            if st.form_submit_button("حفظ الملاحظة") and txt.strip():
                no = emp.split(" — ")[0]
                now = datetime.now().isoformat(timespec="seconds")
                c = conn()
                c.execute("""
                    INSERT INTO notes(emp_no,note,active,created_at,updated_at)
                    VALUES(?,?,?,?,?)
                """, (no, txt.strip(), 1, now, now))
                c.commit()
                c.close()
                st.success("تم حفظ الملاحظة.")

        n = active_notes()
        if not n.empty:
            st.dataframe(n[["id", "emp_no", "name", "note", "updated_at"]], use_container_width=True)
            nid = st.number_input("رقم الملاحظة للحذف", min_value=0, step=1)
            if st.button("حذف الملاحظة") and nid:
                c = conn()
                c.execute(
                    "UPDATE notes SET active=0,updated_at=? WHERE id=?",
                    (datetime.now().isoformat(timespec="seconds"), int(nid))
                )
                c.commit()
                c.close()
                st.success("تم حذف الملاحظة.")

elif page == "القواعد والقوالب":
    st.title("⚙️ القواعد والقوالب")
    asub = st.text_input("عنوان عربي", setting("subject_ar"))
    esub = st.text_input("English subject", setting("subject_en"))
    ab = st.text_area("القالب العربي", setting("body_ar"), height=250)
    eb = st.text_area("English template", setting("body_en"), height=250)

    st.caption(
        "المتغيرات: {name} {month} {period} {details}"
    )

    if st.button("حفظ القوالب"):
        for k, v in [
            ("subject_ar", asub), ("subject_en", esub),
            ("body_ar", ab), ("body_en", eb)
        ]:
            set_setting(k, v)
        st.success("تم حفظ القوالب.")

    st.subheader("قواعد الدوام")
    st.info(
        "من يعمل السبت = 8 ساعات في أيام العمل. "
        "من لا يعمل السبت = 9 ساعات من الأحد إلى الخميس. "
        "الجمعة يتم تمييزها كـ Overtime ولا تُحسب كمغادرة مبكرة عادية."
    )

elif page == "مراجعة المسودات":
    st.title("📧 مراجعة المسودات")
    r = st.session_state.analysis
    if r.empty:
        st.info("حللي التقارير أولًا.")
    else:
        s = r[(~r.excluded) & r.email.astype(str).str.contains("@", na=False)].copy()
        if s.empty:
            st.warning("لا توجد حالات جاهزة للمراجعة.")
        else:
            st.caption("⚠️ رسالة واحدة فقط لكل موظف عن جميع حالاته في الشهر. Gmail يفتح مسودة للمراجعة فقط ولا يرسل تلقائيًا.")
            for (no,name,email),g in s.groupby(["emp_no","name","email"],dropna=False):
                g=g.sort_values("date")
                details="\n".join(
                    f"- {x.date:%Y-%m-%d}: {x.case_type}" + (f" — مدة: {x.duration}" if str(x.duration).strip() not in {"","nan"} else "")
                    for _,x in g.iterrows()
                )
                month=g.date.min().strftime("%B %Y")
                sub=setting("subject_ar").format(month=month)
                body=setting("body_ar").format(name=name,month=month,period=f"{g.date.min():%Y-%m-%d} إلى {g.date.max():%Y-%m-%d}",details=details)
                with st.expander(f"📧 {name} — {email} | {len(g)} حالات"):
                    st.dataframe(g[["date","case_type","duration","project","reason"]],use_container_width=True,hide_index=True)
                    body=st.text_area("نص الرسالة — قابل للتعديل",body,height=260,key="body_"+str(no))
                    sub=st.text_input("العنوان — قابل للتعديل",sub,key="sub_"+str(no))
                    st.link_button("فتح Gmail كمسودة للمراجعة",gmail(email,sub,body))
                    st.caption("بعد الإرسال، اذهبي إلى الأرشيف الشهري وحددي تم الإرسال واكتبي رد الموظف إن وجد.")

elif page == "السجل":
    st.title("📜 السجل")
    c = conn()
    h = pd.read_sql_query(
        "SELECT * FROM history ORDER BY id DESC", c
    )
    c.close()
    st.dataframe(h, use_container_width=True)

elif page == "الأرشيف الشهري":
    st.title("📚 الأرشيف الشهري")
    st.caption("كل شهر محفوظ كنسخة مستقلة. يمكنك الرجوع للحالات، حالة الإرسال، ردود الموظفين، وتصدير الشهر إلى Excel.")

    c=conn(); months=pd.read_sql_query("SELECT * FROM archive_months ORDER BY month_key DESC",c); c.close()
    if not months.empty:
        labels=[f"{r.month_label} — {r.total_cases} حالة — حفظ {r.saved_at}" for _,r in months.iterrows()]
        chosen=st.selectbox("اختاري الشهر",labels)
        mk=months.iloc[labels.index(chosen)].month_key
        a=archive_df(mk)
        if not a.empty:
            st.subheader("🔎 الفلاتر")
            c1,c2,c3,c4=st.columns(4)
            emp_filter=c1.selectbox("الموظف",["الكل"]+sorted(a.name.dropna().astype(str).unique().tolist()))
            type_filter=c2.selectbox("نوع الحالة",["الكل"]+sorted(a.case_type.dropna().astype(str).unique().tolist()))
            project_filter=c3.selectbox("المشروع",["الكل"]+sorted([x for x in a.project.fillna("").astype(str).unique() if x.strip()]))
            send_filter=c4.selectbox("الإرسال",["الكل","تم الإرسال","لم يتم الإرسال"])
            v=a.copy()
            if emp_filter!="الكل": v=v[v.name==emp_filter]
            if type_filter!="الكل": v=v[v.case_type==type_filter]
            if project_filter!="الكل": v=v[v.project==project_filter]
            if send_filter=="تم الإرسال": v=v[v.sent==1]
            elif send_filter=="لم يتم الإرسال": v=v[v.sent==0]

            st.subheader("👥 الحالات مجمعة حسب الموظف")
            for no,g in v.groupby(["emp_no","name","email"],sort=False):
                emp_no_v,name_v,email_v=no
                sent_count=int(g.sent.sum()); total=len(g)
                with st.expander(f"👤 {name_v} — {email_v} | {total} حالات | أرسل: {sent_count}/{total}"):
                    show=g[["date","case_type","duration","project","project_priority","reason","note","sent","reply_note"]].copy()
                    show.columns=["التاريخ","نوع الحالة","المدة","المشروع","أولوية المشروع","السبب","الملاحظة","تم الإرسال","رد الموظف"]
                    st.dataframe(show,use_container_width=True,hide_index=True)
                    for _,row in g.iterrows():
                        cc1,cc2=st.columns([1,3])
                        sent=cc1.checkbox("تم الإرسال",value=bool(row.sent),key=f"sent_{row.id}")
                        reply=cc2.text_input("ملاحظات الرد",value=str(row.reply_note or ""),key=f"reply_{row.id}")
                        if st.button("💾 حفظ",key=f"savecase_{row.id}"):
                            update_archive_case(row.id,sent,reply)
                            st.success("تم حفظ حالة الإرسال والرد.")

            st.divider()
            export=v.copy()
            export.columns=["id","month_key","emp_no","name","email","project","project_priority","date","case_type","duration","worked_hours","expected_hours","reason","note","excluded","first_day","sent","sent_at","reply_note","subject","body","created_at"]
            st.download_button("⬇️ تنزيل الشهر Excel",archive_excel_bytes(export),file_name=f"attendance_archive_{mk}.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            st.download_button("⬇️ تنزيل الشهر CSV",v.to_csv(index=False).encode("utf-8-sig"),file_name=f"attendance_archive_{mk}.csv",mime="text/csv")
    else:
        st.info("لا يوجد أرشيف محفوظ حتى الآن. حللي الشهر ثم استخدمي زر «حفظ الشهر في الأرشيف».")

elif page == "نسخة احتياطية":
    st.title("💾 نسخة احتياطية")
    st.warning(
        "على Streamlit Cloud قد تكون الملفات المحلية مؤقتة بعد إعادة تشغيل التطبيق. "
        "حملي نسخة قاعدة البيانات دوريًا، خصوصًا بعد تحديث الموظفين أو الملاحظات."
    )
    if DB.exists():
        st.download_button(
            "⬇️ تحميل قاعدة البيانات",
            DB.read_bytes(),
            "attendance_system_backup.db"
        )

    st.subheader("استعادة قاعدة بيانات")
    restore = st.file_uploader(
        "ارفعي ملف attendance_system_backup.db",
        type=["db", "sqlite"],
        key="restore_db"
    )
    if restore and st.button("استعادة النسخة"):
        DB.write_bytes(restore.getvalue())
        st.success("تمت استعادة النسخة. أعيدي تشغيل التطبيق.")
