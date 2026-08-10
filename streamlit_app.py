
import streamlit as st
import pandas as pd
import sqlite3, re
from datetime import datetime, date
from urllib.parse import quote
from pathlib import Path

APP_DIR = Path(__file__).resolve().parent
DB_PATH = APP_DIR / "attendance_system.db"

st.set_page_config(page_title="نظام إشعارات الغياب", page_icon="📧", layout="wide")
APP_VERSION = "V8 — Projects + Duration + Work Hours"

AR = {
    "title":"نظام إشعارات الغياب","dashboard":"الرئيسية","reports":"رفع التقارير",
    "employees":"الموظفون","notes":"الملاحظات","projects":"المشاريع",
    "rules":"القواعد والإعدادات","review":"مراجعة الإشعارات","history":"السجل",
    "backup":"نسخة احتياطية","analyze":"تحليل التقارير"
}
EN = {
    "title":"Attendance Notification System","dashboard":"Dashboard","reports":"Upload Reports",
    "employees":"Employees","notes":"Notes","projects":"Projects",
    "rules":"Rules & Settings","review":"Review Notifications","history":"History",
    "backup":"Backup","analyze":"Analyze Reports"
}
if "lang" not in st.session_state:
    st.session_state.lang = "ar"
lang = st.sidebar.radio("Language / اللغة", ["العربية", "English"], horizontal=True)
st.session_state.lang = "ar" if lang == "العربية" else "en"
T = AR if st.session_state.lang == "ar" else EN

def db():
    conn = sqlite3.connect(DB_PATH, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    c = db()
    c.executescript("""
    CREATE TABLE IF NOT EXISTS employees (
        emp_no TEXT PRIMARY KEY, name TEXT, email TEXT,
        category TEXT DEFAULT 'full', works_saturday INTEGER DEFAULT 1,
        hire_date TEXT, active INTEGER DEFAULT 1,
        project TEXT DEFAULT '', project_priority TEXT DEFAULT 'normal',
        updated_at TEXT
    );
    CREATE TABLE IF NOT EXISTS notes (
        id INTEGER PRIMARY KEY AUTOINCREMENT, emp_no TEXT, note TEXT,
        active INTEGER DEFAULT 1, created_at TEXT, updated_at TEXT
    );
    CREATE TABLE IF NOT EXISTS projects (
        project TEXT PRIMARY KEY, priority TEXT DEFAULT 'normal',
        notes TEXT DEFAULT '', updated_at TEXT
    );
    
CREATE TABLE IF NOT EXISTS settings (key TEXT PRIMARY KEY, value TEXT);

# Lightweight SQLite migration so existing deployments keep working after new fields are added.
def ensure_employee_columns():
    c = db()
    cols = {row["name"] for row in c.execute("PRAGMA table_info(employees)").fetchall()}
    if "project" not in cols:
        c.execute("ALTER TABLE employees ADD COLUMN project TEXT DEFAULT ''")
    if "project_priority" not in cols:
        c.execute("ALTER TABLE employees ADD COLUMN project_priority TEXT DEFAULT 'normal'")
    c.commit()
    c.close()

    CREATE TABLE IF NOT EXISTS history (
        id INTEGER PRIMARY KEY AUTOINCREMENT, emp_no TEXT, email TEXT,
        subject TEXT, body TEXT, status TEXT, created_at TEXT
    );
    """)
    defaults = {
        "subject_ar":"إشعار حضور - {month}",
        "subject_en":"Attendance Notification - {month}",
        "body_ar":"""السلام عليكم {name}،

نود إشعاركم بوجود سجل حضور في نظام الحضور والانصراف عن الفترة {period}.

التفاصيل:
{details}

في حال وجود إجازة أو استئذان أو مبرر لم يتم تسجيله، نرجو تزويد قسم الموارد البشرية بالتوضيح خلال يومي عمل من تاريخ استلام هذا الإشعار.

مع الشكر،
قسم الموارد البشرية""",
        "body_en":"""Dear {name},

This is to notify you that an attendance record was identified for the period {period}.

Details:
{details}

If you have an approved leave, permission, missing-fingerprint request, or another valid justification that is not reflected in the attendance system, please provide the relevant details to HR within two working days.

Regards,
HR Department"""
    }
    for k, v in defaults.items():
        c.execute("INSERT OR IGNORE INTO settings(key,value) VALUES (?,?)", (k,v))
    c.commit(); c.close()

init_db()
ensure_employee_columns()

def get_setting(k, default=""):
    c=db(); row=c.execute("SELECT value FROM settings WHERE key=?", (k,)).fetchone(); c.close()
    return row["value"] if row else default

def set_setting(k,v):
    c=db(); c.execute("INSERT OR REPLACE INTO settings(key,value) VALUES (?,?)",(k,v)); c.commit(); c.close()

def clean_col(x):
    return re.sub(r"\s+", " ", str(x)).strip().lower()

def find_header_row(raw, candidates):
    best, best_score = 0, -1
    for i in range(min(len(raw), 25)):
        vals=[clean_col(v) for v in raw.iloc[i].tolist()]
        score=sum(1 for cand in candidates if any(clean_col(cand)==v or clean_col(cand) in v for v in vals))
        if score > best_score:
            best_score, best = score, i
    return best

def read_report(uploaded, candidates):
    raw=pd.read_excel(uploaded, header=None)
    h=find_header_row(raw,candidates)
    df=pd.read_excel(uploaded,header=h).dropna(how="all").copy()
    seen, labels = {}, []
    for c in df.columns:
        base=str(c).strip()
        n=seen.get(base,0)
        labels.append(base if n==0 else f"{base}__dup{n}")
        seen[base]=n+1
    df.columns=labels
    return df

def col_index(df,names):
    wanted=[clean_col(n) for n in names]
    for i,c in enumerate(df.columns):
        if clean_col(c) in wanted: return i
    for i,c in enumerate(df.columns):
        cc=clean_col(c)
        if any(w in cc or cc in w for w in wanted): return i
    return None

def series_at(df,idx):
    return pd.Series([""]*len(df),index=df.index,dtype="object") if idx is None else df.iloc[:,idx]

def parse_date(v):
    if pd.isna(v): return pd.NaT
    return pd.to_datetime(v,errors="coerce",dayfirst=False)

def normalize_emp_no_value(x):
    if pd.isna(x): return ""
    s=str(x).strip()
    if re.fullmatch(r"\d+\.0",s): s=s[:-2]
    return s

def normalize_time_value(v):
    if pd.isna(v): return pd.NaT
    s=str(v).strip()
    if not s or s.lower() in ("nan","none","nat"): return pd.NaT
    return pd.to_datetime(s,errors="coerce")

def normalize_duration(v):
    if pd.isna(v): return ""
    s=str(v).strip()
    if not s or s.lower() in ("nan","none","nat"): return ""
    return s

def duration_minutes(v):
    if pd.isna(v): return None
    s=str(v).lower().strip()
    h=re.search(r"(\d+)\s*(?:hour|hours|ساعة|ساعات)",s)
    m=re.search(r"(\d+)\s*(?:minute|minutes|دقيقة|دقائق)",s)
    if h or m:
        return (int(h.group(1)) if h else 0)*60 + (int(m.group(1)) if m else 0)
    return None

def normalize_employees(uploaded):
    df=read_report(uploaded,["Serial","Employment Number","Employee Name","Email"])
    no=col_index(df,["Employment Number","Emp. No.","Employee Number","رقم الموظف"])
    name=col_index(df,["Employee Name","Emp. Name","اسم الموظف"])
    email=col_index(df,["Email","Work Email","Company Email","البريد الإلكتروني"])
    hire=col_index(df,["Hiring Date","Hire Date","تاريخ المباشرة"])
    out=pd.DataFrame(index=df.index)
    out["emp_no"]=series_at(df,no).map(normalize_emp_no_value)
    out["name"]=series_at(df,name).astype(str).replace({"nan":"","None":""}).str.strip()
    out["email"]=series_at(df,email).astype(str).replace({"nan":"","None":""}).str.strip()
    out["hire_date"]=series_at(df,hire).map(parse_date).dt.strftime("%Y-%m-%d").fillna("") if hire is not None else ""
    return out[out.emp_no.ne("")].drop_duplicates("emp_no")

def save_employees(df):
    c=db(); now=datetime.now().isoformat(timespec="seconds")
    for _,r in df.iterrows():
        old=c.execute("SELECT category,works_saturday,active,project,project_priority FROM employees WHERE emp_no=?",
                      (str(r.emp_no),)).fetchone()
        c.execute("""INSERT OR REPLACE INTO employees
        (emp_no,name,email,category,works_saturday,hire_date,active,project,project_priority,updated_at)
        VALUES(?,?,?,?,?,?,?,?,?,?)""",
        (str(r.emp_no),str(r.get("name","")),str(r.get("email","")),
         old["category"] if old else "full", old["works_saturday"] if old else 1,
         str(r.get("hire_date","")), old["active"] if old else 1,
         old["project"] if old else "", old["project_priority"] if old else "normal", now))
    c.commit(); c.close()

def get_employees():
    c=db(); df=pd.read_sql_query("SELECT * FROM employees ORDER BY name",c); c.close(); return df

def normalize_attendance(uploaded):
    df=read_report(uploaded,[
        "Emp. Name","Emp. No.","Attendance Date","Transaction Type",
        "Employee Name","Employee Number","Employment Number",
        "رقم الموظف","اسم الموظف","تاريخ الحضور","التاريخ",
        "Entry Time","Exit Time","Check In","Check Out",
        "وقت الدخول","وقت الخروج","Duration","مدة","مدة التأخير"
    ])
    no=col_index(df,["Emp. No.","Employment Number","Employee Number","Emp No","Employment No","رقم الموظف","الرقم الوظيفي"])
    name=col_index(df,["Emp. Name","Employee Name","Employee name","اسم الموظف"])
    d=col_index(df,["Attendance Date","Date","تاريخ الحضور","التاريخ"])
    entry=col_index(df,["Entry Time","Check In","Clock In","وقت الدخول","وقت الحضور"])
    exitc=col_index(df,["Exit Time","Check Out","Clock Out","وقت الخروج","وقت الانصراف"])
    dur=col_index(df,["Duration","Duration Time","مدة","المدة","مدة التأخير","مدة الخروج","Duration (Hours)"])
    typ=col_index(df,["Transaction Type","Type","نوع الحركة","نوع العملية","الحالة","Attendance Status"])
    value=col_index(df,["Transaction Value","Value","قيمة الحركة","القيمة"])
    status=col_index(df,["Status","الحالة","حالة الطلب"])
    out=pd.DataFrame(index=df.index)
    out["emp_no"]=series_at(df,no).map(normalize_emp_no_value)
    out["name"]=series_at(df,name).astype(str).replace({"nan":"","None":""}).str.strip()
    out["date"]=series_at(df,d).map(parse_date)
    out["entry_time"]=series_at(df,entry).map(normalize_time_value)
    out["exit_time"]=series_at(df,exitc).map(normalize_time_value)
    out["duration"]=series_at(df,dur).map(normalize_duration)
    out["event"]=series_at(df,typ).astype(str).replace({"nan":"","None":""}).str.strip()
    out["value"]=series_at(df,value).astype(str).replace({"nan":"","None":""}).str.strip()
    out["status"]=series_at(df,status).astype(str).replace({"nan":"","None":""}).str.strip()
    return out[out.emp_no.ne("") & out.date.notna()].copy()

def classify_event(row):
    s=" ".join([str(row.get("event","")),str(row.get("value","")),str(row.get("status",""))]).lower()
    if any(x in s for x in ["absence","absent","غياب","غائب"]): return "غياب"
    if any(x in s for x in ["early out","early leave","early exit","early departure","خروج مبكر","مغادرة مبكرة"]): return "خروج مبكر"
    if any(x in s for x in ["morning lateness","late","tardy","delay","تأخير","متأخر"]): return "تأخير صباحي"
    if any(x in s for x in ["departure","leaving","مغادرة"]): return "مغادرة"
    if any(x in s for x in ["missing fingerprint","missed fingerprint","fingerprint missing","forgot fingerprint","نسيان البصمة","نسي البصمة","بصمة مفقودة"]):
        if any(x in s for x in ["out","exit","خروج"]): return "نسيان بصمة خروج"
        if any(x in s for x in ["in","entry","دخول"]): return "نسيان بصمة دخول"
        return "نسيان بصمة"
    if any(x in s for x in ["overtime","over time","إضافي","اوفر تايم"]): return "عمل إضافي"
    return str(row.get("event","")).strip() or "حالة حضور"

def normalize_leave(uploaded):
    df=read_report(uploaded,["Emp. No.","Transaction Type","From Date","To Date","Status"])
    no=col_index(df,["Emp. No.","Employment Number","Employee Number","رقم الموظف"])
    frm=col_index(df,["From Date","Start Date","من تاريخ"]); to=col_index(df,["To Date","End Date","إلى تاريخ"])
    typ=col_index(df,["Transaction Type","Type","نوع الطلب"]); status=col_index(df,["Status","الحالة"])
    out=pd.DataFrame(index=df.index)
    out["emp_no"]=series_at(df,no).map(normalize_emp_no_value)
    out["from"]=series_at(df,frm).map(parse_date); out["to"]=series_at(df,to).map(parse_date)
    out["type"]=series_at(df,typ).astype(str).replace({"nan":"","None":""}).str.strip()
    out["status"]=series_at(df,status).astype(str).replace({"nan":"","None":""}).str.strip()
    return out[out.emp_no.ne("")].copy()

def normalize_permission(uploaded): return normalize_leave(uploaded)

def normalize_fingerprint(uploaded):
    df=read_report(uploaded,["Employment Number","Attendance Date","Date Requested","Status"])
    no=col_index(df,["Employment Number","Emp. No.","Employee Number","رقم الموظف"])
    d=col_index(df,["Attendance Date","Date","تاريخ"]); status=col_index(df,["Status","الحالة"])
    out=pd.DataFrame(index=df.index)
    out["emp_no"]=series_at(df,no).map(normalize_emp_no_value)
    out["date"]=series_at(df,d).map(parse_date)
    out["status"]=series_at(df,status).astype(str).replace({"nan":"","None":""}).str.strip()
    return out[out.emp_no.ne("") & out.date.notna()].copy()

def active_notes():
    c=db(); df=pd.read_sql_query("""SELECT n.*,e.name FROM notes n LEFT JOIN employees e ON e.emp_no=n.emp_no
    WHERE n.active=1 ORDER BY n.updated_at DESC""",c); c.close(); return df

def note_for(emp_no):
    c=db(); rows=c.execute("SELECT note FROM notes WHERE emp_no=? AND active=1 ORDER BY updated_at DESC",(str(emp_no),)).fetchall(); c.close()
    return " | ".join(r["note"] for r in rows)

def category_label(cat):
    return {"full":"إرسال كامل","absence_only":"غياب فقط","exclude":"مستثنى",
            "half_day":"نصف دوام","special":"احتياجات خاصة","trainee":"متدرب/طالب"}.get(cat,cat)

def project_info(emp_no):
    c=db(); r=c.execute("SELECT project,project_priority FROM employees WHERE emp_no=?",(str(emp_no),)).fetchone(); c.close()
    if not r: return "", "normal"
    return r["project"] or "", r["project_priority"] or "normal"

def project_priority_label(p):
    return {"very_high":"🔴 دقيق جدًا","high":"🟠 عالي","medium":"🟡 متوسط","normal":"🟢 عادي"}.get(p,p)

def required_hours(row):
    # Saturday workers: 8 hours on working days.
    # Non-Saturday workers: 9 hours Sunday-Thursday; Saturday is excluded.
    d=pd.Timestamp(row["date"])
    works_sat=bool(row.get("works_saturday",1))
    if d.day_name()=="Saturday":
        return 8.0 if works_sat else 0.0
    if d.day_name()=="Friday":
        return 0.0
    return 8.0 if works_sat else 9.0

def actual_hours(row):
    a=row.get("entry_time"); b=row.get("exit_time")
    if pd.isna(a) or pd.isna(b): return None
    try:
        delta=pd.Timestamp(b)-pd.Timestamp(a)
        if delta.total_seconds()<0: delta += pd.Timedelta(days=1)
        return round(delta.total_seconds()/3600,2)
    except Exception: return None

def early_out_valid(row):
    if row["case_type"]!="خروج مبكر": return True
    req=required_hours(row)
    if req==0: return False
    actual=actual_hours(row)
    if actual is None: return True
    return actual < req - 0.01

def analyze(att_file,leave_file,permission_file,fingerprint_file):
    att=normalize_attendance(att_file)
    leaves=normalize_leave(leave_file) if leave_file else pd.DataFrame()
    perms=normalize_permission(permission_file) if permission_file else pd.DataFrame()
    fps=normalize_fingerprint(fingerprint_file) if fingerprint_file else pd.DataFrame()
    emps=get_employees()

    if emps.empty:
        minimal=att[["emp_no","name"]].drop_duplicates("emp_no").copy()
        minimal["email"]=""; minimal["hire_date"]=""; minimal["category"]="full"
        minimal["works_saturday"]=1; minimal["active"]=1
        save_employees(minimal); emps=get_employees()

    att["case_type"]=att.apply(classify_event,axis=1)
    # Ensure optional project fields exist even for legacy employee tables.
    for _col, _default in [("project", ""), ("project_priority", "normal")]:
        if _col not in emps.columns:
            emps[_col] = _default

    att = att.merge(
        emps[["emp_no","name","email","category","works_saturday","hire_date","active","project","project_priority"]],
        on="emp_no", how="left", suffixes=("","_master")
    )
    att["name"] = att["name_master"].where(
        att["name_master"].notna() & att["name_master"].ne(""), att["name"]
    )
    att["project"] = att["project"].fillna("")
    att["project_priority"] = att["project_priority"].fillna("normal")
    att = att[att["active"].fillna(1).astype(bool)].copy()

    reasons=[]; excluded=[]; first_flags=[]; required=[]; actual=[]; early_valid=[]
    def is_approved(s):
        s=str(s).lower()
        return any(w in s for w in ["approved","موافق","مقبول","معتمد"])

    for _,r in att.iterrows():
        emp=str(r.emp_no); d=pd.Timestamp(r.date).normalize(); reason=[]
        if r.category=="exclude": reason.append("موظف مستثنى")
        if d.day_name()=="Saturday" and not bool(r.works_saturday): reason.append("لا يعمل السبت")
        if not leaves.empty:
            x=leaves[(leaves.emp_no==emp)&leaves["from"].notna()&leaves["to"].notna()&
                     (leaves["from"].dt.normalize()<=d)&(leaves["to"].dt.normalize()>=d)&leaves.status.map(is_approved)]
            if not x.empty: reason.append("إجازة معتمدة")
        if not perms.empty:
            x=perms[(perms.emp_no==emp)&perms["from"].notna()&perms["to"].notna()&
                     (perms["from"].dt.normalize()<=d)&(perms["to"].dt.normalize()>=d)&perms.status.map(is_approved)]
            if not x.empty: reason.append("استئذان معتمد")
        if not fps.empty:
            x=fps[(fps.emp_no==emp)&(fps.date.dt.normalize()==d)]
            if not x.empty:
                ft=" ".join(x.status.astype(str).tolist()).lower()
                reason.append("نسيان البصمة خروج" if any(z in ft for z in ["out","exit","خروج"]) else
                              "نسيان البصمة دخول" if any(z in ft for z in ["in","entry","دخول"]) else "نسيان البصمة")
        first=False
        if pd.notna(r.hire_date) and str(r.hire_date).strip() not in ("","nan","NaT"):
            try:
                first=pd.Timestamp(r.hire_date).normalize()==d
                if first: reason.append("أول يوم عمل")
            except Exception: pass
        req=required_hours(r); act=actual_hours(r)
        valid=early_out_valid(r)
        if r.case_type=="خروج مبكر" and not valid:
            reason.append("مكتمل ساعات الدوام — لا يعتبر خروجًا مبكرًا")
            # Do not exclude from the whole report; change the case to informational.
            att.loc[att.index[att.index==r.name],"case_type"]="مكتمل ساعات الدوام"
        reasons.append("; ".join(reason)); excluded.append(bool(reason and any(x in "; ".join(reason) for x in ["مستثنى","لا يعمل السبت","إجازة معتمدة","استئذان معتمد"])))
        first_flags.append(first); required.append(req); actual.append(act); early_valid.append(valid)

    att["reason"]=reasons; att["excluded"]=excluded; att["first_day"]=first_flags
    att["required_hours"]=required; att["actual_hours"]=actual; att["early_out_valid"]=early_valid
    att["note"]=att.emp_no.map(note_for)
    att["category_label"]=att.category.map(category_label).fillna("إرسال كامل")
    att["project_priority_label"]=att.project_priority.map(project_priority_label)

    if not fps.empty:
        existing=set(zip(att.emp_no.astype(str),att.date.dt.normalize()))
        extra=[]
        for _,r in fps.iterrows():
            key=(str(r.emp_no),r.date.normalize())
            if key in existing: continue
            emp=emps[emps.emp_no.astype(str)==str(r.emp_no)]
            if emp.empty: continue
            rr=emp.iloc[0]; ft=str(r.status).lower()
            typ="نسيان بصمة خروج" if any(z in ft for z in ["out","exit","خروج"]) else ("نسيان بصمة دخول" if any(z in ft for z in ["in","entry","دخول"]) else "نسيان بصمة")
            extra.append({"emp_no":str(r.emp_no),"name":rr["name"],"email":rr["email"],"date":r.date,
                          "entry_time":pd.NaT,"exit_time":pd.NaT,"duration":"","event":"","value":"","status":r.status,
                          "case_type":typ,"category":rr["category"],"works_saturday":rr["works_saturday"],
                          "hire_date":rr["hire_date"],"active":rr["active"],"project":rr["project"],
                          "project_priority":rr["project_priority"],"reason":"نسيان البصمة","excluded":True,
                          "first_day":False,"required_hours":required_hours({"date":r.date,"works_saturday":rr["works_saturday"]}),
                          "actual_hours":None,"early_out_valid":True,"note":note_for(str(r.emp_no)),
                          "category_label":category_label(rr["category"]),"project_priority_label":project_priority_label(rr["project_priority"])})
        if extra: att=pd.concat([att,pd.DataFrame(extra)],ignore_index=True)

    return att.reset_index(drop=True), {
        "attendance_rows":len(att),"total_cases":len(att),"unique_employees":int(att.emp_no.nunique()),
        "absence_rows":int((att.case_type=="غياب").sum()),
        "late_rows":int((att.case_type=="تأخير صباحي").sum()),
        "early_leave_rows":int((att.case_type=="خروج مبكر").sum()),
        "departure_rows":int((att.case_type=="مغادرة").sum()),
        "fingerprint_rows":int(att.case_type.str.contains("نسيان بصمة",na=False).sum()),
        "completed_hours_rows":int((att.case_type=="مكتمل ساعات الدوام").sum()),
        "excluded":int(att.excluded.sum())
    }

def make_gmail_url(to,subject,body):
    return f"https://mail.google.com/mail/?view=cm&fs=1&to={quote(str(to))}&su={quote(str(subject))}&body={quote(str(body))}"

st.sidebar.title("📧 "+T["title"])
st.sidebar.caption(APP_VERSION)
page=st.sidebar.radio("Menu",[T["dashboard"],T["reports"],T["employees"],T["notes"],T["projects"],T["rules"],T["review"],T["history"],T["backup"]])

if page==T["dashboard"]:
    st.title("📊 "+T["dashboard"])
    emps=get_employees(); notes=active_notes()
    c=db(); hist=pd.read_sql_query("SELECT * FROM history ORDER BY id DESC LIMIT 100",c); c.close()
    a,b,c1,d=st.columns(4)
    a.metric("الموظفون" if st.session_state.lang=="ar" else "Employees",len(emps))
    b.metric("الملاحظات" if st.session_state.lang=="ar" else "Notes",len(notes))
    c1.metric("سجل الإشعارات" if st.session_state.lang=="ar" else "History",len(hist))
    d.metric("آخر تشغيل" if st.session_state.lang=="ar" else "Last run",get_setting("last_run","—"))
    st.info("ارفعي التقارير من صفحة رفع التقارير ثم راجعي المسودات قبل فتح Gmail. لا يتم إرسال أي بريد تلقائيًا.")

elif page==T["reports"]:
    st.title("📂 "+T["reports"])
    st.caption("ملف الموظفين والإيميلات اختياري. يمكنك تحديثه بأي وقت دون إعادة رفعه كل شهر.")
    emp_file=st.file_uploader("1) ملف الموظفين والإيميلات (اختياري)",type=["xlsx"],key="emp")
    if emp_file and st.button("تحديث دليل الموظفين"):
        try:
            e=normalize_employees(emp_file); save_employees(e); st.success(f"تم تحديث {len(e)} موظف.")
        except Exception as ex: st.error(f"خطأ في ملف الموظفين: {ex}")
    att=st.file_uploader("2) تقرير الحضور والغياب — الشهر كامل",type=["xlsx"],key="att")
    leave=st.file_uploader("3) تقرير الإجازات — الشهر كامل",type=["xlsx"],key="leave")
    perm=st.file_uploader("4) تقرير الاستئذان — الشهر كامل",type=["xlsx"],key="perm")
    fp=st.file_uploader("5) تقرير نسيان البصمة — الشهر كامل",type=["xlsx"],key="fp")
    if st.button("🔎 "+T["analyze"],disabled=not att):
        try:
            res,meta=analyze(att,leave,perm,fp)
            st.session_state["analysis"]=res; st.session_state["analysis_meta"]=meta
            set_setting("last_run",datetime.now().strftime("%Y-%m-%d %H:%M"))
            st.success("تم التحليل.")
            st.json(meta)
            if not res.empty:
                m=st.columns(7)
                for col,label,key in zip(m,["غياب","تأخير صباحي","خروج مبكر","مغادرة","نسيان بصمة","مكتمل الساعات","مستبعد"],
                                         ["absence_rows","late_rows","early_leave_rows","departure_rows","fingerprint_rows","completed_hours_rows","excluded"]):
                    col.metric(label,meta.get(key,0))
                show=res[["emp_no","name","email","project","project_priority_label","date","entry_time","exit_time",
                          "duration","required_hours","actual_hours","case_type","category_label","reason","note","excluded","first_day"]].copy()
                show["entry_time"]=show["entry_time"].map(lambda x: "" if pd.isna(x) else pd.Timestamp(x).strftime("%I:%M %p"))
                show["exit_time"]=show["exit_time"].map(lambda x: "" if pd.isna(x) else pd.Timestamp(x).strftime("%I:%M %p"))
                show.columns=["الرقم الوظيفي","اسم الموظف","الإيميل","المشروع","أولوية المشروع","التاريخ",
                              "الدخول","الخروج","المدة","الساعات المطلوبة","الساعات الفعلية","نوع الحالة",
                              "نوع الإرسال","السبب","ملاحظة","مستبعد","أول يوم"]
                st.dataframe(show,use_container_width=True)
                st.download_button("⬇️ تحميل نتيجة التحليل CSV",show.to_csv(index=False).encode("utf-8-sig"),
                                   file_name="attendance_analysis.csv",mime="text/csv")
        except Exception as ex:
            st.error("تعذر تحليل التقرير. تأكدي أن الملفات هي تقارير ZenHR نفسها.")
            st.exception(ex)

elif page==T["employees"]:
    st.title("👥 "+T["employees"])
    st.write("الإيميلات والمشاريع وإعدادات الدوام قابلة للتعديل هنا، وتبقى محفوظة في قاعدة البيانات.")
    df=get_employees()
    if not df.empty:
        edited=st.data_editor(df,use_container_width=True,num_rows="dynamic",
            column_config={
                "category":st.column_config.SelectboxColumn("الفئة",options=["full","absence_only","exclude","half_day","special","trainee"]),
                "works_saturday":st.column_config.CheckboxColumn("يعمل السبت"),
                "project_priority":st.column_config.SelectboxColumn("أولوية المشروع",options=["normal","medium","high","very_high"])
            })
        if st.button("حفظ التعديلات"):
            c=db()
            for _,r in edited.iterrows():
                c.execute("""UPDATE employees SET name=?,email=?,category=?,works_saturday=?,hire_date=?,active=?,
                             project=?,project_priority=?,updated_at=? WHERE emp_no=?""",
                          (r["name"],r["email"],r["category"],int(r["works_saturday"]),r["hire_date"],int(r["active"]),
                           r["project"],r["project_priority"],datetime.now().isoformat(timespec="seconds"),r["emp_no"]))
            c.commit(); c.close(); st.success("تم الحفظ.")
    st.subheader("إضافة موظف")
    with st.form("new_emp"):
        a,b,c,d=st.columns(4)
        no=a.text_input("رقم الموظف"); name=b.text_input("الاسم"); email=c.text_input("الإيميل")
        hire=d.date_input("تاريخ المباشرة",value=date.today())
        project=st.text_input("المشروع")
        cat=st.selectbox("الفئة",["full","absence_only","exclude","half_day","special","trainee"])
        sat=st.checkbox("يعمل السبت",value=True)
        priority=st.selectbox("أولوية المشروع",["normal","medium","high","very_high"])
        if st.form_submit_button("إضافة") and no:
            save_employees(pd.DataFrame([{"emp_no":no,"name":name,"email":email,"hire_date":hire.isoformat()}]))
            c=db(); c.execute("UPDATE employees SET category=?,works_saturday=?,project=?,project_priority=? WHERE emp_no=?",
                             (cat,int(sat),project,priority,no)); c.commit(); c.close(); st.success("تمت الإضافة.")

elif page==T["notes"]:
    st.title("📝 "+T["notes"])
    emps=get_employees()
    if emps.empty: st.warning("أضيفي الموظفين أولًا.")
    else:
        with st.form("note_form"):
            emp=st.selectbox("الموظف",[f"{r.emp_no} — {r.name}" for _,r in emps.iterrows()])
            note=st.text_area("الملاحظة")
            if st.form_submit_button("حفظ الملاحظة"):
                no=emp.split(" — ")[0]; now=datetime.now().isoformat(timespec="seconds")
                c=db(); c.execute("INSERT INTO notes(emp_no,note,active,created_at,updated_at) VALUES(?,?,?,?,?)",
                                  (no,note,1,now,now)); c.commit(); c.close(); st.success("تم الحفظ.")
        n=active_notes()
        if not n.empty: st.dataframe(n[["id","emp_no","name","note","updated_at"]],use_container_width=True)
        nid=st.number_input("رقم الملاحظة للحذف",min_value=0,step=1)
        if st.button("حذف الملاحظة") and nid:
            c=db(); c.execute("UPDATE notes SET active=0,updated_at=? WHERE id=?",
                              (datetime.now().isoformat(timespec="seconds"),int(nid))); c.commit(); c.close(); st.success("تم الحذف.")

elif page==T["projects"]:
    st.title("🏢 المشاريع")
    st.caption("استخدمي هذه الصفحة لتعريف المشاريع ومستوى التدقيق. ويمكنك ربط كل موظف بالمشروع من صفحة الموظفين.")
    c=db(); projects=pd.read_sql_query("SELECT * FROM projects ORDER BY project",c); c.close()
    if not projects.empty: st.dataframe(projects,use_container_width=True)
    with st.form("project_form"):
        name=st.text_input("اسم المشروع")
        priority=st.selectbox("أولوية المشروع",["normal","medium","high","very_high"])
        notes=st.text_area("ملاحظات المشروع")
        if st.form_submit_button("حفظ المشروع") and name.strip():
            c=db(); c.execute("""INSERT OR REPLACE INTO projects(project,priority,notes,updated_at)
                                 VALUES(?,?,?,?)""",(name.strip(),priority,notes,datetime.now().isoformat(timespec="seconds")))
            c.commit(); c.close(); st.success("تم حفظ المشروع.")

elif page==T["rules"]:
    st.title("⚙️ "+T["rules"])
    st.info("قاعدة الدوام: الموظف الذي يعمل السبت = 8 ساعات في أيام عمله. الموظف الذي لا يعمل السبت = 9 ساعات من الأحد إلى الخميس. الجمعة لا تُحسب كدوام عادي.")
    ar_subject=st.text_input("عنوان الإيميل عربي",get_setting("subject_ar"))
    en_subject=st.text_input("Email subject English",get_setting("subject_en"))
    ar_body=st.text_area("قالب الرسالة العربية",get_setting("body_ar"),height=220)
    en_body=st.text_area("English email template",get_setting("body_en"),height=220)
    st.caption("المتغيرات: {name} {month} {period} {details}")
    if st.button("حفظ القواعد والقوالب"):
        for k,v in [("subject_ar",ar_subject),("subject_en",en_subject),("body_ar",ar_body),("body_en",en_body)]: set_setting(k,v)
        st.success("تم الحفظ.")

elif page==T["review"]:
    st.title("📧 "+T["review"])
    res=st.session_state.get("analysis",pd.DataFrame())
    if res.empty: st.info("ارفعي التقارير وشغلي التحليل أولًا.")
    else:
        sendable=res[(~res.excluded)&res.email.notna()&res.email.astype(str).str.contains("@")].copy()
        st.write(f"جاهز للمراجعة: {sendable.emp_no.nunique()} موظف")
        if not sendable.empty:
            grouped=[]
            for (emp,name,email),g in sendable.groupby(["emp_no","name","email"],dropna=False):
                details=[]
                for _,x in g.iterrows():
                    duration=f" — {x.duration}" if str(x.duration).strip() else ""
                    times=[]
                    if pd.notna(x.entry_time): times.append("دخول "+pd.Timestamp(x.entry_time).strftime("%I:%M %p"))
                    if pd.notna(x.exit_time): times.append("خروج "+pd.Timestamp(x.exit_time).strftime("%I:%M %p"))
                    time_text=(" | "+" | ".join(times)) if times else ""
                    details.append(f"- {pd.Timestamp(x.date).strftime('%Y-%m-%d')} — {x.case_type}{duration}{time_text}")
                details="\n".join(details)
                month=pd.Timestamp(g.date.min()).strftime("%B %Y")
                subject=get_setting("subject_ar" if st.session_state.lang=="ar" else "subject_en").format(month=month)
                body=get_setting("body_ar" if st.session_state.lang=="ar" else "body_en").format(
                    name=name,month=month,period=f"{g.date.min():%Y-%m-%d} إلى {g.date.max():%Y-%m-%d}",details=details)
                grouped.append({"emp_no":emp,"name":name,"email":email,"project":g.project.iloc[0],
                                "project_priority":project_priority_label(g.project_priority.iloc[0]),
                                "note":note_for(emp),"subject":subject,"body":body,
                                "gmail_url":make_gmail_url(email,subject,body)})
            review=pd.DataFrame(grouped)
            st.dataframe(review[["emp_no","name","email","project","project_priority","note","subject"]],use_container_width=True)
            for _,r in review.iterrows():
                with st.expander(f"📧 {r['name']} — {r['email']} — {r['project']}"):
                    if r["note"]: st.warning("ملاحظة: "+r["note"])
                    st.text_area("نص الرسالة",r["body"],height=220,key=f"body_{r.emp_no}")
                    st.link_button("فتح مسودة Gmail للمراجعة",r["gmail_url"])
                    if st.button("تسجيل كمُرسل",key=f"log_{r.emp_no}"):
                        c=db(); c.execute("""INSERT INTO history(emp_no,email,subject,body,status,created_at)
                                             VALUES(?,?,?,?,?,?)""",
                                         (r.emp_no,r.email,r.subject,r.body,"opened/marked",
                                          datetime.now().isoformat(timespec="seconds"))); c.commit(); c.close()
                        st.success("تم تسجيلها في السجل.")
        else: st.warning("لا توجد سجلات جاهزة للإرسال.")

elif page==T["history"]:
    st.title("📜 "+T["history"])
    c=db(); h=pd.read_sql_query("SELECT * FROM history ORDER BY id DESC",c); c.close()
    st.dataframe(h,use_container_width=True)

elif page==T["backup"]:
    st.title("💾 "+T["backup"])
    st.warning("على Streamlit Cloud، SQLite محلي وقد يُعاد تهيئته عند إعادة التشغيل. حملي النسخة الاحتياطية دوريًا.")
    if DB_PATH.exists():
        st.download_button("تحميل قاعدة البيانات الاحتياطية",DB_PATH.read_bytes(),file_name="attendance_system_backup.db")
    st.info("ملفات Excel الشهرية لا تُحفظ في GitHub؛ ترفعينها كل شهر من صفحة التقارير.")
