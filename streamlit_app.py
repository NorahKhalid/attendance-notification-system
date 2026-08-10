
import streamlit as st
import pandas as pd
import sqlite3, re
from datetime import datetime, date, time
from pathlib import Path
from urllib.parse import quote

APP_DIR = Path(__file__).resolve().parent
DB_PATH = APP_DIR / "attendance_system.db"
APP_VERSION = "V10 — Stable"

st.set_page_config(page_title="نظام إشعارات الحضور", page_icon="📧", layout="wide")

AR = {
    "title":"نظام إشعارات الحضور","dashboard":"الرئيسية","reports":"رفع التقارير",
    "employees":"الموظفون","notes":"الملاحظات","rules":"القواعد والإعدادات",
    "review":"مراجعة الإشعارات","history":"السجل","backup":"نسخة احتياطية"
}
EN = {
    "title":"Attendance Notification System","dashboard":"Dashboard","reports":"Upload Reports",
    "employees":"Employees","notes":"Notes","rules":"Rules & Settings",
    "review":"Review Notifications","history":"History","backup":"Backup"
}

if "lang" not in st.session_state:
    st.session_state.lang = "ar"

chosen_lang = st.sidebar.radio("Language / اللغة", ["العربية", "English"], horizontal=True)
st.session_state.lang = "ar" if chosen_lang == "العربية" else "en"
T = AR if st.session_state.lang == "ar" else EN

# ---------------- DB ----------------

def db():
    c = sqlite3.connect(str(DB_PATH), check_same_thread=False)
    c.row_factory = sqlite3.Row
    return c

def table_columns(c, table):
    return {r["name"] for r in c.execute(f"PRAGMA table_info({table})").fetchall()}

def init_db():
    c = db()
    c.executescript("""
    CREATE TABLE IF NOT EXISTS employees (
        emp_no TEXT PRIMARY KEY,
        name TEXT DEFAULT '',
        email TEXT DEFAULT '',
        category TEXT DEFAULT 'full',
        works_saturday INTEGER DEFAULT 1,
        hire_date TEXT DEFAULT '',
        active INTEGER DEFAULT 1,
        project TEXT DEFAULT '',
        project_priority TEXT DEFAULT 'normal',
        updated_at TEXT DEFAULT ''
    );

    CREATE TABLE IF NOT EXISTS notes (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        emp_no TEXT,
        note TEXT,
        active INTEGER DEFAULT 1,
        created_at TEXT DEFAULT '',
        updated_at TEXT DEFAULT ''
    );

    CREATE TABLE IF NOT EXISTS settings (
        key TEXT PRIMARY KEY,
        value TEXT
    );

    CREATE TABLE IF NOT EXISTS history (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        emp_no TEXT,
        email TEXT,
        subject TEXT,
        body TEXT,
        status TEXT,
        created_at TEXT
    );
    """)

    # Safe migration for old databases. Each ALTER is executed separately.
    cols = table_columns(c, "employees")
    additions = {
        "name": "TEXT DEFAULT ''",
        "email": "TEXT DEFAULT ''",
        "category": "TEXT DEFAULT 'full'",
        "works_saturday": "INTEGER DEFAULT 1",
        "hire_date": "TEXT DEFAULT ''",
        "active": "INTEGER DEFAULT 1",
        "project": "TEXT DEFAULT ''",
        "project_priority": "TEXT DEFAULT 'normal'",
        "updated_at": "TEXT DEFAULT ''",
    }
    for col, definition in additions.items():
        if col not in cols:
            c.execute(f"ALTER TABLE employees ADD COLUMN {col} {definition}")

    defaults = {
        "subject_ar": "إشعار حضور - {month}",
        "subject_en": "Attendance Notification - {month}",
        "body_ar": """السلام عليكم {name}،

نود إشعاركم بوجود سجل حضور/غياب خلال الفترة {period}.

التفاصيل:
{details}

في حال وجود إجازة أو استئذان أو مبرر لم يتم تسجيله، نرجو التوضيح لقسم الموارد البشرية خلال يومي عمل.

مع الشكر،
قسم الموارد البشرية""",
        "body_en": """Dear {name},

This is to notify you that an attendance record was identified for the period {period}.

Details:
{details}

If you have an approved leave, permission, fingerprint issue, or another valid justification that is not reflected in the attendance system, please contact HR within two working days.

Regards,
HR Department"""
    }
    for k, v in defaults.items():
        c.execute("INSERT OR IGNORE INTO settings(key,value) VALUES (?,?)", (k, v))

    c.commit()
    c.close()

init_db()

def get_setting(key, default=""):
    c = db()
    row = c.execute("SELECT value FROM settings WHERE key=?", (key,)).fetchone()
    c.close()
    return row["value"] if row else default

def set_setting(key, value):
    c = db()
    c.execute("INSERT OR REPLACE INTO settings(key,value) VALUES (?,?)", (key, str(value)))
    c.commit()
    c.close()

# ---------------- Excel helpers ----------------

def clean_col(x):
    return re.sub(r"\s+", " ", str(x)).strip().lower()

def find_header_row(raw, candidates):
    best_i, best_score = 0, -1
    for i in range(min(len(raw), 25)):
        vals = [clean_col(v) for v in raw.iloc[i].tolist()]
        score = 0
        for cand in candidates:
            cc = clean_col(cand)
            if any(cc == v or cc in v for v in vals):
                score += 1
        if score > best_score:
            best_i, best_score = i, score
    return best_i

def read_report(uploaded, candidates):
    raw = pd.read_excel(uploaded, header=None)
    h = find_header_row(raw, candidates)
    df = pd.read_excel(uploaded, header=h).dropna(how="all").copy()

    seen = {}
    labels = []
    for c in df.columns:
        base = str(c).strip()
        n = seen.get(base, 0)
        labels.append(base if n == 0 else f"{base}__dup{n}")
        seen[base] = n + 1
    df.columns = labels
    return df

def col_index(df, names):
    wanted = [clean_col(x) for x in names]
    for i, c in enumerate(df.columns):
        if clean_col(c) in wanted:
            return i
    for i, c in enumerate(df.columns):
        cc = clean_col(c)
        if any(w in cc or cc in w for w in wanted):
            return i
    return None

def series_at(df, idx):
    if idx is None:
        return pd.Series([""] * len(df), index=df.index)
    return df.iloc[:, idx]

def parse_date(v):
    if pd.isna(v):
        return pd.NaT
    return pd.to_datetime(v, errors="coerce")

def normalize_emp(x):
    if pd.isna(x):
        return ""
    s = str(x).strip()
    if re.fullmatch(r"\d+\.0", s):
        s = s[:-2]
    return s

def normalize_employees(uploaded):
    df = read_report(uploaded, [
        "Employment Number","Employee Name","Email","Hiring Date",
        "Emp. No.","Employee Number"
    ])
    no = col_index(df, ["Employment Number","Emp. No.","Employee Number","رقم الموظف"])
    name = col_index(df, ["Employee Name","Emp. Name","اسم الموظف"])
    email = col_index(df, ["Email","Work Email","Company Email","البريد الإلكتروني"])
    hire = col_index(df, ["Hiring Date","Hire Date","تاريخ المباشرة"])

    out = pd.DataFrame(index=df.index)
    out["emp_no"] = series_at(df, no).map(normalize_emp)
    out["name"] = series_at(df, name).astype(str).replace({"nan":"","None":""}).str.strip()
    out["email"] = series_at(df, email).astype(str).replace({"nan":"","None":""}).str.strip()
    if hire is not None:
        out["hire_date"] = series_at(df, hire).map(parse_date).dt.strftime("%Y-%m-%d").fillna("")
    else:
        out["hire_date"] = ""
    return out[out.emp_no.ne("")].drop_duplicates("emp_no")

def normalize_attendance(uploaded):
    df = read_report(uploaded, [
        "Emp. Name","Emp. No.","Attendance Date","Transaction Type",
        "Employee Name","Employee Number","Employment Number",
        "Date","Transaction Value","Status","Entry Time","Exit Time",
        "Check In","Check Out","وقت الدخول","وقت الخروج"
    ])
    no = col_index(df, ["Emp. No.","Employment Number","Employee Number","رقم الموظف","الرقم الوظيفي"])
    name = col_index(df, ["Emp. Name","Employee Name","اسم الموظف"])
    d = col_index(df, ["Attendance Date","Date","تاريخ الحضور","التاريخ"])
    typ = col_index(df, ["Transaction Type","Type","نوع الحركة","نوع العملية","الحالة"])
    val = col_index(df, ["Transaction Value","Value","قيمة الحركة","القيمة"])
    status = col_index(df, ["Status","الحالة"])
    entry = col_index(df, ["Entry Time","Check In","وقت الدخول","وقت الحضور"])
    exit_ = col_index(df, ["Exit Time","Check Out","وقت الخروج","وقت الانصراف"])
    duration = col_index(df, ["Duration","مدة","المدة","Duration Time"])

    out = pd.DataFrame(index=df.index)
    out["emp_no"] = series_at(df,no).map(normalize_emp)
    out["name"] = series_at(df,name).astype(str).replace({"nan":"","None":""}).str.strip()
    out["date"] = series_at(df,d).map(parse_date)
    out["event"] = series_at(df,typ).astype(str).replace({"nan":"","None":""}).str.strip()
    out["value"] = series_at(df,val).astype(str).replace({"nan":"","None":""}).str.strip()
    out["status"] = series_at(df,status).astype(str).replace({"nan":"","None":""}).str.strip()
    out["entry_time"] = series_at(df,entry).astype(str).replace({"nan":"","None":""}).str.strip()
    out["exit_time"] = series_at(df,exit_).astype(str).replace({"nan":"","None":""}).str.strip()
    out["duration"] = series_at(df,duration).astype(str).replace({"nan":"","None":""}).str.strip()
    return out[out.emp_no.ne("") & out.date.notna()].copy()

def normalize_range_report(uploaded, candidates):
    df = read_report(uploaded, candidates)
    no = col_index(df, ["Emp. No.","Employment Number","Employee Number","رقم الموظف"])
    frm = col_index(df, ["From Date","Start Date","من تاريخ","تاريخ البداية"])
    to = col_index(df, ["To Date","End Date","إلى تاريخ","تاريخ النهاية"])
    typ = col_index(df, ["Transaction Type","Type","Request Type","نوع الطلب"])
    status = col_index(df, ["Status","الحالة"])
    out = pd.DataFrame(index=df.index)
    out["emp_no"] = series_at(df,no).map(normalize_emp)
    out["from"] = series_at(df,frm).map(parse_date)
    out["to"] = series_at(df,to).map(parse_date)
    out["type"] = series_at(df,typ).astype(str).replace({"nan":"","None":""}).str.strip()
    out["status"] = series_at(df,status).astype(str).replace({"nan":"","None":""}).str.strip()
    return out[out.emp_no.ne("")].copy()

def normalize_fingerprint(uploaded):
    df = read_report(uploaded, [
        "Employment Number","Emp. No.","Attendance Date","Date","Date Requested","Status"
    ])
    no = col_index(df, ["Employment Number","Emp. No.","Employee Number","رقم الموظف"])
    d = col_index(df, ["Attendance Date","Date","Date Requested","التاريخ"])
    status = col_index(df, ["Status","الحالة","نوع البصمة"])
    out = pd.DataFrame(index=df.index)
    out["emp_no"] = series_at(df,no).map(normalize_emp)
    out["date"] = series_at(df,d).map(parse_date)
    out["status"] = series_at(df,status).astype(str).replace({"nan":"","None":""}).str.strip()
    return out[out.emp_no.ne("") & out.date.notna()].copy()

# ---------------- Employee / notes ----------------

def get_employees():
    c = db()
    df = pd.read_sql_query("SELECT * FROM employees ORDER BY name", c)
    c.close()
    return df

def save_employees(df):
    c = db()
    now = datetime.now().isoformat(timespec="seconds")
    for _, r in df.iterrows():
        emp = str(r.get("emp_no","")).strip()
        if not emp:
            continue
        old = c.execute("SELECT * FROM employees WHERE emp_no=?", (emp,)).fetchone()
        category = old["category"] if old else "full"
        sat = old["works_saturday"] if old else 1
        active = old["active"] if old else 1
        project = old["project"] if old else ""
        priority = old["project_priority"] if old else "normal"

        c.execute("""
        INSERT OR REPLACE INTO employees
        (emp_no,name,email,category,works_saturday,hire_date,active,project,project_priority,updated_at)
        VALUES (?,?,?,?,?,?,?,?,?,?)
        """, (
            emp, str(r.get("name","")), str(r.get("email","")),
            category, int(sat), str(r.get("hire_date","")),
            int(active), project, priority, now
        ))
    c.commit()
    c.close()

def note_for(emp_no):
    c = db()
    rows = c.execute("""
        SELECT note FROM notes
        WHERE emp_no=? AND active=1
        ORDER BY updated_at DESC
    """, (str(emp_no),)).fetchall()
    c.close()
    return " | ".join(r["note"] for r in rows)

def active_notes():
    c = db()
    df = pd.read_sql_query("""
        SELECT n.*, e.name
        FROM notes n LEFT JOIN employees e ON e.emp_no=n.emp_no
        WHERE n.active=1 ORDER BY n.updated_at DESC
    """, c)
    c.close()
    return df

def category_label(cat):
    return {
        "full":"إرسال كامل",
        "absence_only":"غياب فقط",
        "exclude":"مستثنى",
        "half_day":"نصف دوام",
        "special":"احتياجات خاصة",
        "trainee":"متدرب/طالب"
    }.get(cat, cat)

def priority_label(x):
    return {"normal":"عادي","important":"مهم","critical":"دقيق جدًا"}.get(str(x), str(x))

# ---------------- Classification ----------------

def classify_event(row):
    s = " ".join([
        str(row.get("event","")),
        str(row.get("value","")),
        str(row.get("status",""))
    ]).lower()

    if any(x in s for x in ["absence","absent","غياب","غائب"]):
        return "غياب"
    if any(x in s for x in ["early leave","early exit","early departure","خروج مبكر","مغادرة مبكرة"]):
        return "خروج مبكر"
    if any(x in s for x in ["late","tardy","delay","تأخير","متأخر"]):
        return "تأخير صباحي"
    if any(x in s for x in ["missing fingerprint","missed fingerprint","fingerprint missing","forgot fingerprint","نسيان البصمة","نسي البصمة","بصمة مفقودة"]):
        if any(x in s for x in ["out","exit","خروج"]):
            return "نسيان بصمة خروج"
        if any(x in s for x in ["in","entry","دخول"]):
            return "نسيان بصمة دخول"
        return "نسيان بصمة"
    if any(x in s for x in ["departure","leaving","مغادرة"]):
        return "مغادرة"
    return str(row.get("event","")).strip() or "حالة حضور"

def duration_text(v):
    s = str(v).strip()
    return "" if s.lower() in {"nan","none","nat"} else s

def time_to_minutes(v):
    s = str(v).strip()
    if not s or s.lower() in {"nan","none","nat"}:
        return None
    m = re.search(r"(\d{1,2}):(\d{2})", s)
    if not m:
        return None
    return int(m.group(1))*60 + int(m.group(2))

def required_hours(row):
    # Saturday workers: 8h every working day.
    # Non-Saturday workers: 9h Sunday-Thursday; Saturday excluded.
    d = pd.Timestamp(row["date"])
    if d.day_name() == "Saturday":
        return 8.0 if bool(row.get("works_saturday", 1)) else 0.0
    return 9.0 if not bool(row.get("works_saturday", 1)) else 8.0

def actual_hours(row):
    a = time_to_minutes(row.get("entry_time",""))
    b = time_to_minutes(row.get("exit_time",""))
    if a is None or b is None:
        return None
    diff = b - a
    if diff < 0:
        diff += 24*60
    return round(diff/60, 2)

def approved(s):
    s = str(s).lower()
    return any(x in s for x in ["approved","موافق","مقبول","معتمد"])

# ---------------- Analysis ----------------

def analyze(att_file, leave_file, permission_file, fingerprint_file):
    att = normalize_attendance(att_file)
    leaves = normalize_range_report(leave_file, ["Emp. No.","From Date","To Date","Status"]) if leave_file else pd.DataFrame()
    perms = normalize_range_report(permission_file, ["Emp. No.","From Date","To Date","Status"]) if permission_file else pd.DataFrame()
    fps = normalize_fingerprint(fingerprint_file) if fingerprint_file else pd.DataFrame()

    emps = get_employees()

    if emps.empty:
        minimal = att[["emp_no","name"]].drop_duplicates("emp_no").copy()
        minimal["email"] = ""
        minimal["hire_date"] = ""
        save_employees(minimal)
        emps = get_employees()

    att["case_type"] = att.apply(classify_event, axis=1)

    # Legacy-safe: these columns always exist.
    for col, default in [("project",""), ("project_priority","normal")]:
        if col not in emps.columns:
            emps[col] = default

    master_cols = [
        "emp_no","name","email","category","works_saturday",
        "hire_date","active","project","project_priority"
    ]
    att = att.merge(
        emps[master_cols],
        on="emp_no",
        how="left",
        suffixes=("","_master")
    )
    att["name"] = att["name_master"].where(
        att["name_master"].notna() & att["name_master"].ne(""),
        att["name"]
    )
    att["project"] = att["project"].fillna("")
    att["project_priority"] = att["project_priority"].fillna("normal")
    att = att[att["active"].fillna(1).astype(bool)].copy()

    reasons, excluded, first_flags = [], [], []
    required, actual = [], []

    for idx, r in att.iterrows():
        emp = str(r.emp_no)
        d = pd.Timestamp(r.date).normalize()
        reason = []

        if str(r.category) == "exclude":
            reason.append("موظف مستثنى")

        if d.day_name() == "Saturday" and not bool(r.works_saturday):
            reason.append("لا يعمل السبت")

        if not leaves.empty:
            x = leaves[
                (leaves.emp_no == emp) &
                leaves["from"].notna() & leaves["to"].notna() &
                (leaves["from"].dt.normalize() <= d) &
                (leaves["to"].dt.normalize() >= d) &
                leaves.status.map(approved)
            ]
            if not x.empty:
                reason.append("إجازة معتمدة")

        if not perms.empty:
            x = perms[
                (perms.emp_no == emp) &
                perms["from"].notna() & perms["to"].notna() &
                (perms["from"].dt.normalize() <= d) &
                (perms["to"].dt.normalize() >= d) &
                perms.status.map(approved)
            ]
            if not x.empty:
                reason.append("استئذان معتمد")

        if not fps.empty:
            x = fps[
                (fps.emp_no == emp) &
                (fps.date.dt.normalize() == d)
            ]
            if not x.empty:
                fs = " ".join(x.status.astype(str).tolist()).lower()
                if any(z in fs for z in ["out","exit","خروج"]):
                    reason.append("نسيان بصمة خروج")
                elif any(z in fs for z in ["in","entry","دخول"]):
                    reason.append("نسيان بصمة دخول")
                else:
                    reason.append("نسيان البصمة")

        first = False
        hd = str(r.hire_date).strip()
        if hd and hd not in {"nan","NaT"}:
            try:
                first = pd.Timestamp(hd).normalize() == d
                if first:
                    reason.append("أول يوم عمل")
            except Exception:
                pass

        req = required_hours(r)
        act = actual_hours(r)

        # If an "Early Out" row exists but actual hours reach the required hours,
        # it is informational rather than a real early departure.
        if r.case_type == "خروج مبكر" and act is not None and req > 0 and act >= req:
            att.at[idx, "case_type"] = "مكتمل ساعات الدوام"
            reason.append("مكتمل ساعات الدوام — لا يعتبر خروجًا مبكرًا")

        reasons.append("; ".join(dict.fromkeys(reason)))
        excluded.append(any(x in reason for x in [
            "موظف مستثنى","لا يعمل السبت","إجازة معتمدة","استئذان معتمد"
        ]))
        first_flags.append(first)
        required.append(req)
        actual.append(act)

    att["reason"] = reasons
    att["excluded"] = excluded
    att["first_day"] = first_flags
    att["required_hours"] = required
    att["actual_hours"] = actual
    att["note"] = att.emp_no.map(note_for)
    att["category_label"] = att.category.map(category_label).fillna("إرسال كامل")
    att["project_priority_label"] = att.project_priority.map(priority_label)

    # Add fingerprint-only cases.
    if not fps.empty:
        existing = set(zip(att.emp_no.astype(str), att.date.dt.normalize()))
        extra = []
        for _, r in fps.iterrows():
            key = (str(r.emp_no), r.date.normalize())
            if key in existing:
                continue
            emp = emps[emps.emp_no.astype(str) == str(r.emp_no)]
            if emp.empty:
                continue
            rr = emp.iloc[0]
            fs = str(r.status).lower()
            typ = (
                "نسيان بصمة خروج" if any(z in fs for z in ["out","exit","خروج"])
                else "نسيان بصمة دخول" if any(z in fs for z in ["in","entry","دخول"])
                else "نسيان البصمة"
            )
            extra.append({
                "emp_no": str(r.emp_no), "name": rr["name"], "email": rr["email"],
                "date": r.date, "entry_time": "", "exit_time": "", "duration": "",
                "event": "", "value": "", "status": r.status, "case_type": typ,
                "category": rr["category"], "works_saturday": rr["works_saturday"],
                "hire_date": rr["hire_date"], "active": rr["active"],
                "project": rr["project"], "project_priority": rr["project_priority"],
                "reason": "نسيان البصمة", "excluded": True, "first_day": False,
                "required_hours": 0, "actual_hours": None,
                "note": note_for(str(r.emp_no)),
                "category_label": category_label(rr["category"]),
                "project_priority_label": priority_label(rr["project_priority"])
            })
        if extra:
            att = pd.concat([att, pd.DataFrame(extra)], ignore_index=True)

    meta = {
        "attendance_rows": len(att),
        "total_cases": len(att),
        "unique_employees": int(att.emp_no.nunique()),
        "absence_rows": int((att.case_type == "غياب").sum()),
        "late_rows": int((att.case_type == "تأخير صباحي").sum()),
        "early_leave_rows": int((att.case_type == "خروج مبكر").sum()),
        "departure_rows": int((att.case_type == "مغادرة").sum()),
        "fingerprint_rows": int(att.case_type.str.contains("نسيان بصمة", na=False).sum()),
        "completed_hours_rows": int((att.case_type == "مكتمل ساعات الدوام").sum()),
        "excluded": int(att.excluded.sum())
    }
    return att.reset_index(drop=True), meta

# ---------------- Gmail ----------------

def make_gmail_url(to, subject, body):
    return (
        "https://mail.google.com/mail/?view=cm&fs=1"
        f"&to={quote(str(to))}"
        f"&su={quote(str(subject))}"
        f"&body={quote(str(body))}"
    )

def build_details(g):
    rows = []
    for _, x in g.iterrows():
        bits = [pd.Timestamp(x.date).strftime("%Y-%m-%d"), str(x.case_type)]
        if str(x.entry_time).strip() and str(x.entry_time).lower() != "nan":
            bits.append(f"دخول: {x.entry_time}")
        if str(x.exit_time).strip() and str(x.exit_time).lower() != "nan":
            bits.append(f"خروج: {x.exit_time}")
        if str(x.duration).strip() and str(x.duration).lower() != "nan":
            bits.append(f"المدة: {x.duration}")
        elif pd.notna(x.actual_hours):
            bits.append(f"الساعات الفعلية: {x.actual_hours}")
        rows.append("- " + " | ".join(bits))
    return "\n".join(rows)

# ---------------- UI ----------------

st.sidebar.title("📧 " + T["title"])
st.sidebar.caption(APP_VERSION)
page = st.sidebar.radio(
    "Menu",
    [T["dashboard"],T["reports"],T["employees"],T["notes"],
     T["rules"],T["review"],T["history"],T["backup"]]
)

if page == T["dashboard"]:
    st.title("📊 " + T["dashboard"])
    e = get_employees()
    n = active_notes()
    c = db()
    h = pd.read_sql_query("SELECT * FROM history ORDER BY id DESC LIMIT 100", c)
    c.close()

    a,b,c1,d = st.columns(4)
    a.metric("الموظفون" if st.session_state.lang=="ar" else "Employees", len(e))
    b.metric("الملاحظات" if st.session_state.lang=="ar" else "Notes", len(n))
    c1.metric("السجل" if st.session_state.lang=="ar" else "History", len(h))
    d.metric("آخر تشغيل", get_setting("last_run","—"))

    st.info("🔒 النظام لا يرسل أي بريد تلقائيًا. فقط يجهز مسودة Gmail للمراجعة، وأنتِ تضغطين إرسال بنفسك.")

elif page == T["reports"]:
    st.title("📂 " + T["reports"])

    emp_file = st.file_uploader(
        "1) ملف الموظفين والإيميلات — اختياري للتحديث",
        type=["xlsx"], key="emp"
    )
    if emp_file and st.button("تحديث دليل الموظفين"):
        try:
            e = normalize_employees(emp_file)
            save_employees(e)
            st.success(f"تم تحديث {len(e)} موظف.")
        except Exception as ex:
            st.error("تعذر تحديث ملف الموظفين.")
            st.exception(ex)

    att = st.file_uploader("2) تقرير الحضور — الشهر كامل", type=["xlsx"], key="att")
    leave = st.file_uploader("3) تقرير الإجازات — الشهر كامل", type=["xlsx"], key="leave")
    perm = st.file_uploader("4) تقرير الاستئذان — الشهر كامل", type=["xlsx"], key="perm")
    fp = st.file_uploader("5) تقرير نسيان البصمة — الشهر كامل", type=["xlsx"], key="fp")

    if st.button("🔎 تحليل التقارير", disabled=not att):
        try:
            res, meta = analyze(att, leave, perm, fp)
            st.session_state["analysis"] = res
            st.session_state["analysis_meta"] = meta
            set_setting("last_run", datetime.now().strftime("%Y-%m-%d %H:%M"))
            st.success("تم التحليل بنجاح.")
            st.json(meta)

            if not res.empty:
                cols = [
                    "emp_no","name","email","project","project_priority_label",
                    "date","case_type","entry_time","exit_time","duration",
                    "required_hours","actual_hours","category_label",
                    "reason","note","excluded","first_day"
                ]
                cols = [c for c in cols if c in res.columns]
                show = res[cols].copy()
                show.columns = [
                    {
                        "emp_no":"الرقم الوظيفي","name":"اسم الموظف","email":"الإيميل",
                        "project":"المشروع","project_priority_label":"أولوية المشروع",
                        "date":"التاريخ","case_type":"نوع الحالة",
                        "entry_time":"الدخول","exit_time":"الخروج","duration":"مدة الحالة",
                        "required_hours":"ساعات الدوام المطلوبة","actual_hours":"الساعات الفعلية",
                        "category_label":"نوع الإرسال","reason":"السبب","note":"ملاحظة",
                        "excluded":"مستبعد","first_day":"أول يوم"
                    }.get(c,c) for c in cols
                ]
                st.dataframe(show, use_container_width=True)
                st.download_button(
                    "⬇️ تحميل نتيجة التحليل CSV",
                    show.to_csv(index=False).encode("utf-8-sig"),
                    file_name="attendance_analysis.csv",
                    mime="text/csv"
                )
        except Exception as ex:
            st.error("تعذر تحليل التقرير.")
            st.exception(ex)

elif page == T["employees"]:
    st.title("👥 " + T["employees"])
    df = get_employees()

    if not df.empty:
        edited = st.data_editor(
            df,
            use_container_width=True,
            num_rows="dynamic",
            column_config={
                "category": st.column_config.SelectboxColumn(
                    "الفئة",
                    options=["full","absence_only","exclude","half_day","special","trainee"]
                ),
                "works_saturday": st.column_config.CheckboxColumn("يعمل السبت"),
                "project_priority": st.column_config.SelectboxColumn(
                    "أولوية المشروع",
                    options=["normal","important","critical"]
                )
            }
        )

        if st.button("حفظ تعديلات الموظفين"):
            c = db()
            for _, r in edited.iterrows():
                c.execute("""
                    UPDATE employees
                    SET name=?, email=?, category=?, works_saturday=?,
                        hire_date=?, active=?, project=?, project_priority=?, updated_at=?
                    WHERE emp_no=?
                """, (
                    r.get("name",""), r.get("email",""), r.get("category","full"),
                    int(r.get("works_saturday",1)), r.get("hire_date",""),
                    int(r.get("active",1)), r.get("project",""),
                    r.get("project_priority","normal"),
                    datetime.now().isoformat(timespec="seconds"),
                    r.get("emp_no","")
                ))
            c.commit()
            c.close()
            st.success("تم حفظ التعديلات.")

    st.subheader("إضافة موظف")
    with st.form("new_emp"):
        a,b,c1,d = st.columns(4)
        no = a.text_input("رقم الموظف")
        name = b.text_input("الاسم")
        email = c1.text_input("الإيميل")
        hire = d.date_input("تاريخ المباشرة", value=date.today())
        project = st.text_input("المشروع")
        priority = st.selectbox("أولوية المشروع", ["normal","important","critical"])
        cat = st.selectbox("الفئة", ["full","absence_only","exclude","half_day","special","trainee"])
        sat = st.checkbox("يعمل السبت", value=True)

        if st.form_submit_button("إضافة") and no:
            c = db()
            c.execute("""
                INSERT OR REPLACE INTO employees
                (emp_no,name,email,category,works_saturday,hire_date,active,project,project_priority,updated_at)
                VALUES (?,?,?,?,?,?,?,?,?,?)
            """, (
                no,name,email,cat,int(sat),hire.isoformat(),1,project,priority,
                datetime.now().isoformat(timespec="seconds")
            ))
            c.commit()
            c.close()
            st.success("تمت الإضافة.")

elif page == T["notes"]:
    st.title("📝 " + T["notes"])
    emps = get_employees()
    if emps.empty:
        st.warning("أضيفي الموظفين أولًا.")
    else:
        with st.form("note_form"):
            emp = st.selectbox("الموظف", [
                f"{r.emp_no} — {r.name}" for _,r in emps.iterrows()
            ])
            note = st.text_area("الملاحظة")
            if st.form_submit_button("حفظ الملاحظة"):
                no = emp.split(" — ")[0]
                now = datetime.now().isoformat(timespec="seconds")
                c = db()
                c.execute("""
                    INSERT INTO notes(emp_no,note,active,created_at,updated_at)
                    VALUES(?,?,?,?,?)
                """, (no,note,1,now,now))
                c.commit()
                c.close()
                st.success("تم حفظ الملاحظة.")

        n = active_notes()
        if not n.empty:
            st.dataframe(n[["id","emp_no","name","note","updated_at"]], use_container_width=True)

elif page == T["rules"]:
    st.title("⚙️ " + T["rules"])
    ar_subject = st.text_input("عنوان الإيميل عربي", get_setting("subject_ar"))
    en_subject = st.text_input("Email subject English", get_setting("subject_en"))
    ar_body = st.text_area("قالب الرسالة العربية", get_setting("body_ar"), height=240)
    en_body = st.text_area("English email template", get_setting("body_en"), height=240)
    st.caption("المتغيرات المتاحة: {name} {month} {period} {details}")

    if st.button("حفظ القوالب"):
        set_setting("subject_ar", ar_subject)
        set_setting("subject_en", en_subject)
        set_setting("body_ar", ar_body)
        set_setting("body_en", en_body)
        st.success("تم حفظ القوالب.")

elif page == T["review"]:
    st.title("📧 " + T["review"])
    st.warning("🔒 لن يتم إرسال أي رسالة تلقائيًا. النظام يفتح Gmail فقط كمسودة، والإرسال النهائي بيدك.")

    res = st.session_state.get("analysis", pd.DataFrame())
    if res.empty:
        st.info("ارفعي التقارير وشغلي التحليل أولًا.")
    else:
        sendable = res[
            (~res.excluded) &
            res.email.notna() &
            res.email.astype(str).str.contains("@")
        ].copy()

        st.write(f"جاهز للمراجعة: {sendable.emp_no.nunique()} موظف")

        if not sendable.empty:
            for (emp,name,email), g in sendable.groupby(
                ["emp_no","name","email"], dropna=False
            ):
                month = pd.Timestamp(g.date.min()).strftime("%B %Y")
                details = build_details(g)
                period = f"{g.date.min():%Y-%m-%d} إلى {g.date.max():%Y-%m-%d}"

                subject_key = "subject_ar" if st.session_state.lang=="ar" else "subject_en"
                body_key = "body_ar" if st.session_state.lang=="ar" else "body_en"

                subject = get_setting(subject_key).format(month=month)
                body = get_setting(body_key).format(
                    name=name, month=month, period=period, details=details
                )

                with st.expander(f"📧 {name} — {email}"):
                    if g.note.astype(str).str.strip().ne("").any():
                        st.warning("⚠️ توجد ملاحظة على هذا الموظف: " + note_for(emp))

                    edited_subject = st.text_input(
                        "عنوان الرسالة",
                        subject,
                        key=f"subject_{emp}"
                    )
                    edited_body = st.text_area(
                        "نص الرسالة — تقدرين تعدلينه قبل فتح Gmail",
                        body,
                        height=240,
                        key=f"body_{emp}"
                    )

                    url = make_gmail_url(email, edited_subject, edited_body)

                    st.link_button(
                        "📧 فتح مسودة Gmail — لن يتم الإرسال تلقائيًا",
                        url
                    )

                    if st.button("تسجيل كمُراجع", key=f"log_{emp}"):
                        c = db()
                        c.execute("""
                            INSERT INTO history
                            (emp_no,email,subject,body,status,created_at)
                            VALUES(?,?,?,?,?,?)
                        """, (
                            emp,email,edited_subject,edited_body,
                            "reviewed/opened",
                            datetime.now().isoformat(timespec="seconds")
                        ))
                        c.commit()
                        c.close()
                        st.success("تم تسجيل المراجعة.")

elif page == T["history"]:
    st.title("📜 " + T["history"])
    c = db()
    h = pd.read_sql_query("SELECT * FROM history ORDER BY id DESC", c)
    c.close()
    st.dataframe(h, use_container_width=True)

elif page == T["backup"]:
    st.title("💾 " + T["backup"])
    st.warning("على Streamlit Cloud، SQLite محلي. لا ترفعي قاعدة البيانات إلى GitHub.")
    if DB_PATH.exists():
        st.download_button(
            "تحميل نسخة احتياطية من قاعدة البيانات",
            DB_PATH.read_bytes(),
            file_name="attendance_system_backup.db"
        )
