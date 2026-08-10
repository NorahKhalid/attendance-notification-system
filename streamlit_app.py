
import streamlit as st
import pandas as pd
import sqlite3, os, re
from datetime import datetime, date
from urllib.parse import quote
from pathlib import Path

APP_DIR = Path(__file__).resolve().parent
DB_PATH = APP_DIR / "attendance_system.db"

st.set_page_config(page_title="نظام إشعارات الغياب", page_icon="📧", layout="wide")

# ---------- i18n ----------
AR = {
    "title":"نظام إشعارات الغياب",
    "dashboard":"الرئيسية",
    "reports":"رفع التقارير",
    "employees":"الموظفون",
    "notes":"الملاحظات",
    "rules":"القواعد والإعدادات",
    "review":"مراجعة الإشعارات",
    "history":"السجل",
    "backup":"نسخة احتياطية",
    "analyze":"تحليل التقارير",
    "save":"حفظ",
    "delete":"حذف",
    "upload":"رفع",
}
EN = {
    "title":"Attendance Notification System",
    "dashboard":"Dashboard",
    "reports":"Upload Reports",
    "employees":"Employees",
    "notes":"Notes",
    "rules":"Rules & Settings",
    "review":"Review Notifications",
    "history":"History",
    "backup":"Backup",
    "analyze":"Analyze Reports",
    "save":"Save",
    "delete":"Delete",
    "upload":"Upload",
}

if "lang" not in st.session_state:
    st.session_state.lang = "ar"
lang = st.sidebar.radio("Language / اللغة", ["العربية", "English"], horizontal=True)
st.session_state.lang = "ar" if lang == "العربية" else "en"
T = AR if st.session_state.lang == "ar" else EN

# ---------- database ----------
def db():
    conn = sqlite3.connect(DB_PATH, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    c = db()
    c.executescript("""
    CREATE TABLE IF NOT EXISTS employees (
        emp_no TEXT PRIMARY KEY,
        name TEXT,
        email TEXT,
        category TEXT DEFAULT 'full',
        works_saturday INTEGER DEFAULT 1,
        hire_date TEXT,
        active INTEGER DEFAULT 1,
        updated_at TEXT
    );
    CREATE TABLE IF NOT EXISTS notes (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        emp_no TEXT,
        note TEXT,
        active INTEGER DEFAULT 1,
        created_at TEXT,
        updated_at TEXT
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
    CREATE TABLE IF NOT EXISTS analysis_results (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        run_at TEXT,
        emp_no TEXT,
        name TEXT,
        email TEXT,
        date TEXT,
        event TEXT,
        category TEXT,
        reason TEXT,
        note TEXT,
        gmail_url TEXT
    );
    """)
    defaults = {
        "subject_ar":"إشعار غياب - {month}",
        "subject_en":"Absence Notification - {month}",
        "body_ar":"""السلام عليكم {name}،

نود إشعاركم بوجود سجل غياب/تأخير في نظام الحضور والانصراف عن الفترة {period}.

الأيام المسجلة:
{details}

في حال وجود إجازة أو استئذان أو مبرر لم يتم تسجيله، نرجو تزويد قسم الموارد البشرية بالمستندات/التوضيح خلال يومي عمل من تاريخ استلام هذا الإشعار.

مع الشكر،
قسم الموارد البشرية""",
        "body_en":"""Dear {name},

This is to notify you that an attendance record was identified for the period {period}.

Recorded days:
{details}

If you have an approved leave, permission, fingerprint-missing request, or another valid justification that is not reflected in the attendance system, please provide the relevant details to HR within two working days.

Regards,
HR Department"""
    }
    for k,v in defaults.items():
        c.execute("INSERT OR IGNORE INTO settings(key,value) VALUES (?,?)",(k,v))
    c.commit(); c.close()

init_db()

def get_setting(k, default=""):
    c=db(); row=c.execute("SELECT value FROM settings WHERE key=?",(k,)).fetchone(); c.close()
    return row["value"] if row else default

def set_setting(k,v):
    c=db(); c.execute("INSERT OR REPLACE INTO settings(key,value) VALUES (?,?)",(k,v)); c.commit(); c.close()

# ---------- Excel helpers ----------
def clean_col(x):
    return re.sub(r"\s+"," ",str(x)).strip().lower()

def find_header_row(raw, candidates):
    best = None
    best_score = -1
    for i in range(min(len(raw), 15)):
        vals = [clean_col(v) for v in raw.iloc[i].tolist()]
        score = sum(any(c in v for v in vals) for c in candidates)
        if score > best_score:
            best_score, best = score, i
    return best if best_score > 0 else 0

def read_report(uploaded, candidates):
    raw = pd.read_excel(uploaded, header=None)
    h = find_header_row(raw, candidates)
    df = pd.read_excel(uploaded, header=h)
    df = df.dropna(how="all").copy()
    df.columns = [str(c).strip() for c in df.columns]
    return df

def col(df, names):
    mapping = {clean_col(c):c for c in df.columns}
    for n in names:
        if clean_col(n) in mapping:
            return mapping[clean_col(n)]
    # fuzzy
    for k,c in mapping.items():
        for n in names:
            if clean_col(n) in k or k in clean_col(n):
                return c
    return None

def parse_date(v):
    if pd.isna(v): return pd.NaT
    s=str(v).strip()
    # Remove weekday prefix and parse
    return pd.to_datetime(s, errors="coerce", dayfirst=False)

def employee_no_series(df):
    c=col(df,["Emp. No.","Employment Number","Employee Number","Emp No","Employment No"])
    return c

def normalize_employees(uploaded):
    df=read_report(uploaded,["Serial","Employment Number","Employee Name","Email"])
    no=col(df,["Employment Number","Emp. No.","Employee Number"])
    name=col(df,["Employee Name","Emp. Name","Employee name"])
    email=col(df,["Email","Work Email","Company Email"])
    hire=col(df,["Hiring Date","Hire Date"])
    out=pd.DataFrame()
    emp_series = df[no]
    if isinstance(emp_series, pd.DataFrame):
        emp_series = emp_series.iloc[:, 0]
    out["emp_no"] = emp_series.map(lambda x: "" if pd.isna(x) else re.sub(r"\\.0$", "", str(x).strip())) if no else ""
    out["name"]=df[name].astype(str).replace("nan","").str.strip() if name else ""
    out["email"]=df[email].astype(str).replace("nan","").str.strip() if email else ""
    out["hire_date"]=df[hire].map(parse_date).dt.strftime("%Y-%m-%d") if hire else ""
    return out[out.emp_no.ne("")].drop_duplicates("emp_no")

def save_employees(df):
    c=db()
    now=datetime.now().isoformat(timespec="seconds")
    for _,r in df.iterrows():
        old=c.execute("SELECT category,works_saturday,active FROM employees WHERE emp_no=?",(str(r.emp_no),)).fetchone()
        cat=old["category"] if old else "full"
        sat=old["works_saturday"] if old else 1
        active=old["active"] if old else 1
        c.execute("""INSERT OR REPLACE INTO employees(emp_no,name,email,category,works_saturday,hire_date,active,updated_at)
                     VALUES(?,?,?,?,?,?,?,?)""",
                  (str(r.emp_no),str(r.get("name","")),str(r.get("email","")),cat,sat,str(r.get("hire_date","")),active,now))
    c.commit(); c.close()

def get_employees():
    c=db(); df=pd.read_sql_query("SELECT * FROM employees ORDER BY name",c); c.close(); return df

# ---------- processing ----------
def normalize_attendance(uploaded):
    df=read_report(uploaded,["Emp. Name","Emp. No.","Attendance Date","Transaction Type"])
    no=col(df,["Emp. No.","Employment Number","Employee Number"])
    name=col(df,["Emp. Name","Employee Name","Employee name"])
    d=col(df,["Attendance Date","Date"])
    typ=col(df,["Transaction Type","Type"])
    value=col(df,["Transaction Value","Value"])
    out=pd.DataFrame()
    emp_series = df[no]
    if isinstance(emp_series, pd.DataFrame):
        emp_series = emp_series.iloc[:, 0]
    out["emp_no"] = emp_series.map(lambda x: "" if pd.isna(x) else re.sub(r"\\.0$", "", str(x).strip()))
    out["name"]=df[name].astype(str).replace("nan","").str.strip() if name else ""
    out["date"]=df[d].map(parse_date) if d else pd.NaT
    out["event"]=df[typ].astype(str).str.strip() if typ else ""
    out["value"]=df[value].astype(str) if value else ""
    return out[out.emp_no.ne("") & out.date.notna()].copy()

def normalize_leave(uploaded):
    df=read_report(uploaded,["Emp. No.","Transaction Type","From Date","To Date","Status"])
    no=col(df,["Emp. No.","Employment Number"])
    frm=col(df,["From Date"])
    to=col(df,["To Date"])
    typ=col(df,["Transaction Type","Type"])
    status=col(df,["Status"])
    out=pd.DataFrame()
    emp_series = df[no]
    if isinstance(emp_series, pd.DataFrame):
        emp_series = emp_series.iloc[:, 0]
    out["emp_no"] = emp_series.map(lambda x: "" if pd.isna(x) else re.sub(r"\\.0$", "", str(x).strip()))
    out["from"]=df[frm].map(parse_date)
    out["to"]=df[to].map(parse_date)
    out["type"]=df[typ].astype(str).str.strip() if typ else ""
    out["status"]=df[status].astype(str).str.strip() if status else ""
    return out[out.emp_no.ne("")].copy()

def normalize_permission(uploaded):
    return normalize_leave(uploaded)

def normalize_fingerprint(uploaded):
    df=read_report(uploaded,["Employment Number","Attendance Date","Date Requested","Status"])
    no=col(df,["Employment Number","Emp. No.","Employee Number"])
    d=col(df,["Attendance Date","Date"])
    status=col(df,["Status"])
    out=pd.DataFrame()
    emp_series = df[no]
    if isinstance(emp_series, pd.DataFrame):
        emp_series = emp_series.iloc[:, 0]
    out["emp_no"] = emp_series.map(lambda x: "" if pd.isna(x) else re.sub(r"\\.0$", "", str(x).strip()))
    out["date"]=df[d].map(parse_date)
    out["status"]=df[status].astype(str).str.strip() if status else ""
    return out[out.emp_no.ne("") & out.date.notna()].copy()

def active_notes():
    c=db(); df=pd.read_sql_query("""SELECT n.*, e.name FROM notes n LEFT JOIN employees e ON e.emp_no=n.emp_no
                                    WHERE n.active=1 ORDER BY n.updated_at DESC""",c); c.close(); return df

def note_for(emp_no):
    c=db(); rows=c.execute("SELECT note FROM notes WHERE emp_no=? AND active=1 ORDER BY updated_at DESC",(str(emp_no),)).fetchall(); c.close()
    return " | ".join(r["note"] for r in rows)

def category_label(cat):
    return {"full":"إرسال كامل","absence_only":"غياب فقط","exclude":"مستثنى","half_day":"نصف دوام",
            "special":"احتياجات خاصة","trainee":"متدرب/طالب"}.get(cat,cat)

def analyze(att_file, leave_file, permission_file, fingerprint_file):
    att=normalize_attendance(att_file)
    leaves=normalize_leave(leave_file) if leave_file else pd.DataFrame()
    perms=normalize_permission(permission_file) if permission_file else pd.DataFrame()
    fps=normalize_fingerprint(fingerprint_file) if fingerprint_file else pd.DataFrame()
    emps=get_employees()
    if emps.empty:
        return pd.DataFrame(), {"error":"لا يوجد موظفون. ارفعي ملف الموظفين أولاً."}

    absn=att[att.event.str.lower().eq("absence")].copy()
    absn["weekday"]=absn.date.dt.day_name()
    # Merge employee rules
    absn=absn.merge(emps[["emp_no","name","email","category","works_saturday","hire_date","active"]],on="emp_no",how="left",suffixes=("","_master"))
    absn["name"]=absn["name_master"].where(absn["name_master"].notna() & absn["name_master"].ne(""),absn["name"])
    absn=absn[absn["active"].fillna(1).astype(bool)].copy()

    approved_words = ["approved","موافق","approved non deduct"]
    def is_approved(s):
        s=str(s).lower()
        return any(w in s for w in approved_words)
    reasons=[]
    keep=[]
    for _,r in absn.iterrows():
        emp=str(r.emp_no); d=r.date.normalize()
        cat=str(r.category) if pd.notna(r.category) else "full"
        reason=[]
        # permanent exclusion
        if cat=="exclude":
            reason.append("موظف مستثنى")
        # Saturday
        if d.day_name()=="Saturday" and int(r.works_saturday or 0)==0:
            reason.append("لا يعمل السبت")
        # approved leave covering day
        if not leaves.empty:
            x=leaves[(leaves.emp_no==emp)&(leaves["from"].notna())&(leaves["to"].notna())&
                     (leaves["from"].dt.normalize()<=d)&(leaves["to"].dt.normalize()>=d)&leaves.status.map(is_approved)]
            if not x.empty: reason.append("إجازة معتمدة")
        # permission covering day
        if not perms.empty:
            x=perms[(perms.emp_no==emp)&(perms["from"].notna())&(perms["to"].notna())&
                    (perms["from"].dt.normalize()<=d)&(perms["to"].dt.normalize()>=d)&perms.status.map(is_approved)]
            if not x.empty: reason.append("استئذان معتمد")
        # fingerprint
        if not fps.empty:
            x=fps[(fps.emp_no==emp)&(fps.date.dt.normalize()==d)]
            if not x.empty: reason.append("طلب نسيان بصمة")
        # first day
        first_day=False
        if pd.notna(r.hire_date) and str(r.hire_date):
            try:
                first_day=pd.Timestamp(r.hire_date).normalize()==d
                if first_day: reason.append("أول يوم عمل")
            except: pass
        reasons.append("; ".join(reason))
        keep.append(len(reason)==0 or "طلب نسيان بصمة" in reason)
    absn["reason"]=reasons
    absn["excluded"]=~pd.Series(keep,index=absn.index)
    absn["first_day"]=absn["reason"].str.contains("أول يوم عمل",na=False)
    absn["note"]=absn.emp_no.map(note_for)
    # If fingerprint exists, don't send as normal absence but flag it.
    absn.loc[absn.reason.str.contains("طلب نسيان بصمة",na=False),"excluded"]=True
    absn["category_label"]=absn.category.map(category_label).fillna("إرسال كامل")
    return absn.reset_index(drop=True), {"attendance_rows":len(att),"absence_rows":len(absn),
                                         "excluded":int(absn.excluded.sum())}

def make_gmail_url(to, subject, body):
    return f"https://mail.google.com/mail/?view=cm&fs=1&to={quote(str(to))}&su={quote(str(subject))}&body={quote(str(body))}"

# ---------- sidebar ----------
st.sidebar.title("📧 " + T["title"])
page = st.sidebar.radio("Menu", [T["dashboard"],T["reports"],T["employees"],T["notes"],T["rules"],T["review"],T["history"],T["backup"]])

if page == T["dashboard"]:
    st.title("📊 " + T["dashboard"])
    emps=get_employees(); notes=active_notes()
    c=db(); hist=pd.read_sql_query("SELECT * FROM history ORDER BY id DESC LIMIT 100",c); c.close()
    a,b,c1,d=st.columns(4)
    a.metric("الموظفون" if st.session_state.lang=="ar" else "Employees",len(emps))
    b.metric("الملاحظات" if st.session_state.lang=="ar" else "Notes",len(notes))
    c1.metric("سجل الإشعارات" if st.session_state.lang=="ar" else "History",len(hist))
    d.metric("آخر تشغيل" if st.session_state.lang=="ar" else "Last run", get_setting("last_run","—"))
    st.info("ابدئي من «رفع التقارير» ثم شغلي التحليل. لن يتم إرسال أي بريد تلقائيًا.")

elif page == T["reports"]:
    st.title("📂 " + T["reports"])
    st.caption("كل ملف يمكن تغييره كل شهر. لا حاجة لتثبيت بيانات الشهر داخل الكود.")
    emp_file=st.file_uploader("1) ملف الموظفين والإيميلات (اختياري إذا كانت البيانات محفوظة)",type=["xlsx"],key="emp")
    if emp_file and st.button("تحديث دليل الموظفين"):
        try:
            e=normalize_employees(emp_file); save_employees(e)
            st.success(f"تم تحديث {len(e)} موظف.")
        except Exception as ex: st.error(f"خطأ في ملف الموظفين: {ex}")

    att=st.file_uploader("2) تقرير الحضور والغياب — الشهر كامل",type=["xlsx"],key="att")
    leave=st.file_uploader("3) تقرير الإجازات — الشهر كامل",type=["xlsx"],key="leave")
    perm=st.file_uploader("4) تقرير الاستئذان — الشهر كامل",type=["xlsx"],key="perm")
    fp=st.file_uploader("5) تقرير نسيان البصمة — الشهر كامل",type=["xlsx"],key="fp")

    if st.button("🔎 " + T["analyze"],disabled=not att):
        if not att:
            st.warning("ارفعي تقرير الحضور.")
        else:
            try:
                res,meta=analyze(att,leave,perm,fp)
                st.session_state["analysis"]=res
                st.session_state["analysis_meta"]=meta
                set_setting("last_run",datetime.now().strftime("%Y-%m-%d %H:%M"))
                st.success("تم التحليل.")
                st.write(meta)
                st.dataframe(res[["emp_no","name","email","date","event","category_label","reason","note","excluded","first_day"]],use_container_width=True)
            except Exception as ex:
                st.exception(ex)

elif page == T["employees"]:
    st.title("👥 " + T["employees"])
    st.write("يمكنك تحديث الإيميلات برفع ملف جديد، أو تعديل الموظفين مباشرة.")
    df=get_employees()
    if not df.empty:
        edited=st.data_editor(df, use_container_width=True, num_rows="dynamic",
                              column_config={"category":st.column_config.SelectboxColumn("الفئة",options=["full","absence_only","exclude","half_day","special","trainee"]),
                                             "works_saturday":st.column_config.CheckboxColumn("يعمل السبت")})
        if st.button("حفظ التعديلات"):
            c=db()
            for _,r in edited.iterrows():
                c.execute("""UPDATE employees SET name=?,email=?,category=?,works_saturday=?,hire_date=?,active=?,updated_at=?
                             WHERE emp_no=?""",(r["name"],r["email"],r["category"],int(r["works_saturday"]),r["hire_date"],int(r["active"]),datetime.now().isoformat(timespec="seconds"),r["emp_no"]))
            c.commit(); c.close(); st.success("تم الحفظ.")
    st.subheader("إضافة موظف")
    with st.form("new_emp"):
        a,b,c,d=st.columns(4)
        no=a.text_input("رقم الموظف"); name=b.text_input("الاسم"); email=c.text_input("الإيميل"); hire=d.date_input("تاريخ المباشرة",value=date.today())
        cat=st.selectbox("الفئة",["full","absence_only","exclude","half_day","special","trainee"])
        sat=st.checkbox("يعمل السبت",value=True)
        if st.form_submit_button("إضافة"):
            if no:
                save_employees(pd.DataFrame([{"emp_no":no,"name":name,"email":email,"hire_date":hire.isoformat()}]))
                c=db(); c.execute("UPDATE employees SET category=?,works_saturday=? WHERE emp_no=?",(cat,int(sat),no)); c.commit(); c.close()
                st.success("تمت الإضافة.")

elif page == T["notes"]:
    st.title("📝 " + T["notes"])
    emps=get_employees()
    if emps.empty:
        st.warning("أضيفي الموظفين أولًا.")
    else:
        with st.form("note_form"):
            emp=st.selectbox("الموظف", [f"{r.emp_no} — {r.name}" for _,r in emps.iterrows()])
            note=st.text_area("الملاحظة")
            if st.form_submit_button("حفظ الملاحظة"):
                no=emp.split(" — ")[0]
                c=db(); now=datetime.now().isoformat(timespec="seconds")
                c.execute("INSERT INTO notes(emp_no,note,active,created_at,updated_at) VALUES(?,?,?,?,?)",(no,note,1,now,now))
                c.commit(); c.close(); st.success("تم الحفظ.")
        n=active_notes()
        if not n.empty:
            st.dataframe(n[["id","emp_no","name","note","updated_at"]],use_container_width=True)
            nid=st.number_input("رقم الملاحظة للحذف",min_value=0,step=1)
            if st.button("حذف الملاحظة") and nid:
                c=db(); c.execute("UPDATE notes SET active=0,updated_at=? WHERE id=?",(datetime.now().isoformat(timespec="seconds"),int(nid))); c.commit(); c.close(); st.success("تم الحذف.")

elif page == T["rules"]:
    st.title("⚙️ " + T["rules"])
    ar_subject=st.text_input("عنوان الإيميل عربي",get_setting("subject_ar"))
    en_subject=st.text_input("Email subject English",get_setting("subject_en"))
    ar_body=st.text_area("قالب الرسالة العربية",get_setting("body_ar"),height=220)
    en_body=st.text_area("English email template",get_setting("body_en"),height=220)
    st.caption("المتغيرات: {name} {month} {period} {details}")
    if st.button("حفظ القواعد والقوالب"):
        for k,v in [("subject_ar",ar_subject),("subject_en",en_subject),("body_ar",ar_body),("body_en",en_body)]:
            set_setting(k,v)
        st.success("تم الحفظ.")

elif page == T["review"]:
    st.title("📧 " + T["review"])
    res=st.session_state.get("analysis",pd.DataFrame())
    if res.empty:
        st.info("ارفعي التقارير وشغلي التحليل أولًا.")
    else:
        sendable=res[(~res.excluded)&res.email.notna()&res.email.astype(str).str.contains("@")].copy()
        st.write(f"جاهز للمراجعة: {len(sendable)} سجل")
        if not sendable.empty:
            grouped=[]
            for (emp,name,email),g in sendable.groupby(["emp_no","name","email"],dropna=False):
                details="\n".join([f"- {pd.Timestamp(x.date).strftime('%Y-%m-%d')} ({x.event})" for _,x in g.iterrows()])
                month=pd.Timestamp(g.date.min()).strftime("%B %Y")
                subject=get_setting("subject_ar" if st.session_state.lang=="ar" else "subject_en").format(month=month)
                body=get_setting("body_ar" if st.session_state.lang=="ar" else "body_en").format(name=name,month=month,period=f"{g.date.min():%Y-%m-%d} إلى {g.date.max():%Y-%m-%d}",details=details)
                url=make_gmail_url(email,subject,body)
                note=note_for(emp)
                grouped.append({"emp_no":emp,"name":name,"email":email,"note":note,"subject":subject,"body":body,"gmail_url":url})
            review=pd.DataFrame(grouped)
            st.dataframe(review[["emp_no","name","email","note","subject"]],use_container_width=True)
            for _,r in review.iterrows():
                with st.expander(f"📧 {r['name']} — {r['email']}"):
                    if r["note"]: st.warning("ملاحظة: "+r["note"])
                    st.text_area("نص الرسالة",r["body"],height=180,key=f"body_{r.emp_no}")
                    st.link_button("فتح مسودة Gmail للمراجعة",r["gmail_url"])
                    if st.button("تسجيل كمُرسل",key=f"log_{r.emp_no}"):
                        c=db(); c.execute("INSERT INTO history(emp_no,email,subject,body,status,created_at) VALUES(?,?,?,?,?,?)",
                                          (r.emp_no,r.email,r.subject,r.body,"opened/marked",datetime.now().isoformat(timespec="seconds"))); c.commit(); c.close()
                        st.success("تم تسجيلها في السجل.")
        else:
            st.warning("لا توجد سجلات جاهزة للإرسال.")

elif page == T["history"]:
    st.title("📜 " + T["history"])
    c=db(); h=pd.read_sql_query("SELECT * FROM history ORDER BY id DESC",c); c.close()
    st.dataframe(h,use_container_width=True)

elif page == T["backup"]:
    st.title("💾 " + T["backup"])
    st.warning("على Streamlit Cloud، الملفات المحلية قد تُعاد تهيئتها عند إعادة تشغيل التطبيق. استخدمي النسخة الاحتياطية دوريًا، أو اربطي Google Sheets/Supabase لاحقًا للتخزين الدائم.")
    if DB_PATH.exists():
        st.download_button("تحميل قاعدة البيانات الاحتياطية",DB_PATH.read_bytes(),file_name="attendance_system_backup.db")
    st.info("ملفات Excel الشهرية لا تُحفظ في GitHub؛ ترفعينها كل شهر من صفحة التقارير.")
