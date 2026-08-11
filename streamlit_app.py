import streamlit as st
import pandas as pd
import sqlite3, re
from pathlib import Path
from datetime import datetime, date
from urllib.parse import quote

APP=Path(__file__).resolve().parent
DB=APP/'attendance_system.db'
st.set_page_config(page_title='Attendance Notification System V11',page_icon='📧',layout='wide')

# ---------- DB ----------
def conn():
    c=sqlite3.connect(DB,check_same_thread=False); c.row_factory=sqlite3.Row; return c

def init_db():
    c=conn(); c.executescript('''
    CREATE TABLE IF NOT EXISTS employees(emp_no TEXT PRIMARY KEY,name TEXT,email TEXT,category TEXT DEFAULT 'full',works_saturday INTEGER DEFAULT 1,hire_date TEXT,active INTEGER DEFAULT 1,project TEXT DEFAULT '',project_priority INTEGER DEFAULT 0,updated_at TEXT);
    CREATE TABLE IF NOT EXISTS notes(id INTEGER PRIMARY KEY AUTOINCREMENT,emp_no TEXT,note TEXT,active INTEGER DEFAULT 1,created_at TEXT,updated_at TEXT);
    CREATE TABLE IF NOT EXISTS settings(key TEXT PRIMARY KEY,value TEXT);
    CREATE TABLE IF NOT EXISTS history(id INTEGER PRIMARY KEY AUTOINCREMENT,emp_no TEXT,email TEXT,subject TEXT,body TEXT,status TEXT,created_at TEXT);
    ''')
    cols={r['name'] for r in c.execute('PRAGMA table_info(employees)').fetchall()}
    for n,t in [('project',"TEXT DEFAULT ''"),('project_priority','INTEGER DEFAULT 0'),('works_saturday','INTEGER DEFAULT 1'),('hire_date','TEXT'),('category',"TEXT DEFAULT 'full'"),('active','INTEGER DEFAULT 1'),('updated_at','TEXT')]:
        if n not in cols:c.execute(f'ALTER TABLE employees ADD COLUMN {n} {t}')
    defaults={'subject_ar':'إشعار غياب/حضور - {month}','subject_en':'Attendance Notification - {month}',
    'body_ar':'السلام عليكم {name}،\n\nنود إشعاركم بوجود سجل حضور في الفترة {period}.\n\nالتفاصيل:\n{details}\n\nفي حال وجود مبرر أو إجازة أو استئذان، نرجو تزويد الموارد البشرية بالتوضيح.\n\nمع الشكر،\nقسم الموارد البشرية',
    'body_en':'Dear {name},\n\nThis is to notify you of an attendance record for {period}.\n\nDetails:\n{details}\n\nIf you have a valid justification, please provide it to HR.\n\nRegards,\nHR Department'}
    for k,v in defaults.items():c.execute('INSERT OR IGNORE INTO settings VALUES(?,?)',(k,v))
    c.commit();c.close()
init_db()

def setting(k):
    c=conn();r=c.execute('SELECT value FROM settings WHERE key=?',(k,)).fetchone();c.close();return r['value'] if r else ''
def set_setting(k,v):
    c=conn();c.execute('INSERT OR REPLACE INTO settings VALUES(?,?)',(k,v));c.commit();c.close()
def employees():
    c=conn();d=pd.read_sql_query('SELECT * FROM employees ORDER BY name',c);c.close();return d
def notes():
    c=conn();d=pd.read_sql_query('SELECT n.*,e.name FROM notes n LEFT JOIN employees e ON e.emp_no=n.emp_no WHERE n.active=1 ORDER BY n.updated_at DESC',c);c.close();return d

def save_emps(df):
    c=conn();now=datetime.now().isoformat(timespec='seconds')
    for _,r in df.iterrows():
        no=str(r.get('emp_no','')).strip()
        if not no:continue
        old=c.execute('SELECT * FROM employees WHERE emp_no=?',(no,)).fetchone()
        vals=(no,str(r.get('name','')),str(r.get('email','')),old['category'] if old else 'full',old['works_saturday'] if old else 1,str(r.get('hire_date',old['hire_date'] if old else '')),old['active'] if old else 1,str(r.get('project',old['project'] if old else '')),old['project_priority'] if old else 0,now)
        c.execute('INSERT OR REPLACE INTO employees VALUES(?,?,?,?,?,?,?,?,?,?)',vals)
    c.commit();c.close()

def clean(x):return re.sub(r'\s+',' ',str(x)).strip().lower()
def header(raw,cands):
    best=0;score=-1
    for i in range(min(25,len(raw))):
        vals=[clean(x) for x in raw.iloc[i].tolist()];s=sum(any(clean(c)==v or clean(c) in v for v in vals) for c in cands)
        if s>score:best,score=i,s
    return best
def read(x,cands):
    raw=pd.read_excel(x,header=None);h=header(raw,cands);d=pd.read_excel(x,header=h).dropna(how='all').copy();
    seen={};cols=[]
    for c in d.columns:
        b=str(c).strip();n=seen.get(b,0);cols.append(b if not n else f'{b}__dup{n}');seen[b]=n+1
    d.columns=cols;return d
def ci(d,names):
    w=[clean(x) for x in names]
    for i,c in enumerate(d.columns):
        if clean(c) in w:return i
    for i,c in enumerate(d.columns):
        if any(a in clean(c) or clean(c) in a for a in w):return i
    return None
def ser(d,i):return pd.Series(['']*len(d),index=d.index) if i is None else d.iloc[:,i]
def empno(x):
    if pd.isna(x):return ''
    s=str(x).strip();return s[:-2] if re.fullmatch(r'\d+\.0',s) else s
def dt(x):return pd.to_datetime(x,errors='coerce')

def norm_emps(f):
    d=read(f,['Employment Number','Employee Name','Email','Project']);o=pd.DataFrame(index=d.index)
    o['emp_no']=ser(d,ci(d,['Employment Number','Emp. No.','Employee Number','رقم الموظف'])).map(empno)
    o['name']=ser(d,ci(d,['Employee Name','Emp. Name','اسم الموظف'])).astype(str).replace({'nan':'','None':''}).str.strip()
    o['email']=ser(d,ci(d,['Email','Work Email','Company Email','البريد الإلكتروني'])).astype(str).replace({'nan':'','None':''}).str.strip()
    h=ci(d,['Hiring Date','Hire Date','تاريخ المباشرة']);o['hire_date']=ser(d,h).map(dt).dt.strftime('%Y-%m-%d') if h is not None else ''
    p=ci(d,['Project','Project Name','المشروع']);o['project']=ser(d,p).astype(str).replace({'nan':'','None':''}).str.strip() if p is not None else ''
    return o[o.emp_no.ne('')].drop_duplicates('emp_no')

def norm_att(f):
    d=read(f,['Emp. No.','Emp. Name','Attendance Date','Transaction Type','Employee Number','Employment Number','تاريخ الحضور']);o=pd.DataFrame(index=d.index)
    o['emp_no']=ser(d,ci(d,['Emp. No.','Employment Number','Employee Number','رقم الموظف','الرقم الوظيفي'])).map(empno)
    o['name']=ser(d,ci(d,['Emp. Name','Employee Name','اسم الموظف'])).astype(str).replace({'nan':'','None':''}).str.strip();o['date']=ser(d,ci(d,['Attendance Date','Date','تاريخ الحضور','التاريخ'])).map(dt)
    o['event']=ser(d,ci(d,['Transaction Type','Type','نوع الحركة','نوع العملية','الحالة'])).astype(str).replace({'nan':'','None':''}).str.strip()
    o['duration']=ser(d,ci(d,['Duration','Late Duration','مدة التأخير','التأخير'])).astype(str).replace({'nan':'','None':''}).str.strip();return o[o.emp_no.ne('')&o.date.notna()].copy()
def norm_range(f):
    if not f:return pd.DataFrame()
    d=read(f,['Emp. No.','Employment Number','From Date','To Date','Status']);o=pd.DataFrame(index=d.index)
    o['emp_no']=ser(d,ci(d,['Emp. No.','Employment Number','Employee Number','رقم الموظف'])).map(empno);a=ci(d,['From Date','Start Date','من تاريخ']);b=ci(d,['To Date','End Date','إلى تاريخ']);q=ci(d,['Date','Attendance Date','التاريخ']);
    o['from']=ser(d,a).map(dt) if a is not None else ser(d,q).map(dt);o['to']=ser(d,b).map(dt) if b is not None else o['from'];o['status']=ser(d,ci(d,['Status','الحالة'])).astype(str).str.lower();return o[o.emp_no.ne('')]
def classify(x):
    s=str(x).lower()
    if any(k in s for k in ['absence','absent','غياب','غائب']):return 'غياب'
    if any(k in s for k in ['early','خروج مبكر','مغادرة مبكرة']):return 'خروج مبكر'
    if any(k in s for k in ['late','tardy','delay','تأخير','متأخر']):return 'تأخير صباحي'
    if any(k in s for k in ['departure','leaving','مغادرة']):return 'مغادرة'
    if any(k in s for k in ['fingerprint','بصمة','finger print']):return 'نسيان بصمة'
    return str(x).strip() or 'حالة حضور'
def approved(s):return any(k in str(s).lower() for k in ['approved','موافق','اعتمد','مقبول'])
def note_for(no):
    c=conn();r=c.execute('SELECT note FROM notes WHERE emp_no=? AND active=1 ORDER BY updated_at DESC',(str(no),)).fetchall();c.close();return ' | '.join(x['note'] for x in r)

def analyze(attf,leavef,permf,fpf):
    a=norm_att(attf);l=norm_range(leavef);p=norm_range(permf);f=norm_range(fpf);e=employees()
    if e.empty:
        z=a[['emp_no','name']].drop_duplicates('emp_no');z['email']='';z['hire_date']='';z['project']='';save_emps(z);e=employees()
    for col,default in [('category','full'),('works_saturday',1),('hire_date',''),('active',1),('project',''),('project_priority',0)]:
        if col not in e:e[col]=default
    a['case_type']=a.event.map(classify);a=a.merge(e[['emp_no','name','email','category','works_saturday','hire_date','active','project','project_priority']],on='emp_no',how='left',suffixes=('','_m'))
    if 'name_m' in a:a['name']=a.name_m.where(a.name_m.notna()&a.name_m.ne(''),a.name)
    a=a[a.active.fillna(1).astype(bool)].copy();reasons=[];ex=[];first=[]
    for _,r in a.iterrows():
        no=str(r.emp_no);d=pd.Timestamp(r.date).normalize();rs=[];fd=False
        if str(r.category)=='exclude':rs.append('موظف مستثنى')
        if d.day_name()=='Saturday' and int(r.works_saturday or 0)==0:rs.append('لا يعمل السبت')
        if not l.empty and not l[(l.emp_no==no)&(l['from'].notna())&(l['to'].notna())&(l['from'].dt.normalize()<=d)&(l['to'].dt.normalize()>=d)&l.status.map(approved)].empty:rs.append('إجازة معتمدة')
        if not p.empty and not p[(p.emp_no==no)&(p['from'].notna())&(p['to'].notna())&(p['from'].dt.normalize()<=d)&(p['to'].dt.normalize()>=d)&p.status.map(approved)].empty:rs.append('استئذان معتمد')
        hd=str(r.hire_date or '')
        if hd and hd not in ['nan','NaT']:
            try:fd=pd.Timestamp(hd).normalize()==d
            except:fd=False
        if fd:rs.append('أول يوم عمل')
        reasons.append('; '.join(rs));ex.append(bool(rs));first.append(fd)
    a['reason']=reasons;a['excluded']=ex;a['first_day']=first;a['note']=a.emp_no.map(note_for);a['day_flag']=a.date.dt.day_name().map({'Friday':'جمعة / Friday','Saturday':'سبت / Saturday'}).fillna('')
    return a,{'attendance_rows':len(a),'absence_rows':int((a.case_type=='غياب').sum()),'late_rows':int((a.case_type=='تأخير صباحي').sum()),'early_leave_rows':int((a.case_type=='خروج مبكر').sum()),'departure_rows':int((a.case_type=='مغادرة').sum()),'fingerprint_rows':int(a.case_type.str.contains('بصمة',na=False).sum()),'excluded':int(a.excluded.sum())}

def gmail(to,s,b):return f'https://mail.google.com/mail/?view=cm&fs=1&to={quote(str(to))}&su={quote(str(s))}&body={quote(str(b))}'

# ---------- UI ----------
if 'analysis' not in st.session_state:st.session_state.analysis=pd.DataFrame()
if 'meta' not in st.session_state:st.session_state.meta={}
st.sidebar.title('📧 Attendance Notification System');st.sidebar.caption('V11 — Stable Dashboard')
page=st.sidebar.radio('القائمة',['الرئيسية','التقارير والتحليل','الموظفون والمشاريع','الملاحظات','القواعد والقوالب','مراجعة المسودات','السجل','نسخة احتياطية'])

if page=='الرئيسية':
    st.title('📊 Dashboard');e=employees();r=st.session_state.analysis;m=st.session_state.meta
    a,b,c,d=st.columns(4);a.metric('الموظفون',len(e));b.metric('الحالات',len(r));c.metric('المستبعد',int(r.excluded.sum()) if not r.empty else 0);d.metric('آخر تشغيل',setting('last_run') or '—')
    if not r.empty:
        st.subheader('الحالات');cs=st.columns(6)
        for i,(x,k) in enumerate([('غياب','absence_rows'),('تأخير','late_rows'),('خروج مبكر','early_leave_rows'),('مغادرة','departure_rows'),('نسيان بصمة','fingerprint_rows'),('مستبعد','excluded')]):cs[i].metric(x,m.get(k,0))
        typ=st.selectbox('فلترة حسب الحالة',['الكل']+sorted(r.case_type.unique().tolist()));proj=st.selectbox('فلترة حسب المشروع',['الكل']+sorted([x for x in r.project.fillna('').astype(str).unique() if x]));v=r.copy()
        if typ!='الكل':v=v[v.case_type==typ]
        if proj!='الكل':v=v[v.project==proj]
        st.dataframe(v[['emp_no','name','email','project','date','case_type','duration','day_flag','reason','note','excluded','first_day']],use_container_width=True)
    else:st.info('ابدئي من صفحة التقارير وارفعِي ملفات الشهر ثم اضغطي تحليل.')

elif page=='التقارير والتحليل':
    st.title('📂 التقارير والتحليل');st.caption('بعد التحليل تبقى النتيجة محفوظة عند التنقل بين الصفحات.')
    ef=st.file_uploader('1) ملف الموظفين والإيميلات والمشاريع — اختياري',type=['xlsx'],key='ef')
    if ef and st.button('تحديث دليل الموظفين'):save_emps(norm_emps(ef));st.success('تم التحديث.')
    af=st.file_uploader('2) تقرير الحضور — الشهر كامل',type=['xlsx'],key='af');lf=st.file_uploader('3) تقرير الإجازات',type=['xlsx'],key='lf');pf=st.file_uploader('4) تقرير الاستئذان',type=['xlsx'],key='pf');ff=st.file_uploader('5) تقرير نسيان البصمة',type=['xlsx'],key='ff')
    if st.button('🔎 تحليل التقارير',disabled=not af):
        try:
            r,m=analyze(af,lf,pf,ff);st.session_state.analysis=r;st.session_state.meta=m;set_setting('last_run',datetime.now().strftime('%Y-%m-%d %H:%M'));st.success('تم التحليل بنجاح.')
        except Exception as ex:st.error('تعذر تحليل التقرير.');st.exception(ex)
    if not st.session_state.analysis.empty:
        r=st.session_state.analysis;types=st.multiselect('إظهار الحالات',sorted(r.case_type.unique()),default=sorted(r.case_type.unique()));v=r[r.case_type.isin(types)];st.dataframe(v,use_container_width=True)
        st.download_button('⬇️ تحميل النتيجة CSV',v.to_csv(index=False).encode('utf-8-sig'),'attendance_analysis.csv','text/csv')

elif page=='الموظفون والمشاريع':
    st.title('👥 الموظفون والمشاريع');e=employees()
    if not e.empty:
        ed=st.data_editor(e,use_container_width=True,num_rows='dynamic',column_config={'category':st.column_config.SelectboxColumn(options=['full','absence_only','exclude','half_day','special','trainee']),'works_saturday':st.column_config.CheckboxColumn(),'active':st.column_config.CheckboxColumn(),'project_priority':st.column_config.CheckboxColumn()})
        if st.button('💾 حفظ التعديلات'):save_emps(ed);st.success('تم الحفظ.')
    with st.form('new'):
        a,b,c,d=st.columns(4);no=a.text_input('الرقم الوظيفي');name=b.text_input('الاسم');email=c.text_input('الإيميل');project=d.text_input('المشروع');hire=st.date_input('تاريخ المباشرة',date.today());cat=st.selectbox('الفئة',['full','absence_only','exclude','half_day','special','trainee']);sat=st.checkbox('يعمل السبت',True);pri=st.checkbox('أولوية المشروع')
        if st.form_submit_button('إضافة') and no:
            save_emps(pd.DataFrame([{'emp_no':no,'name':name,'email':email,'project':project,'hire_date':hire.isoformat()}]));c=conn();c.execute('UPDATE employees SET category=?,works_saturday=?,project_priority=? WHERE emp_no=?',(cat,int(sat),int(pri),no));c.commit();c.close();st.success('تمت الإضافة.')

elif page=='الملاحظات':
    st.title('📝 الملاحظات');e=employees()
    if not e.empty:
        with st.form('note'):
            emp=st.selectbox('الموظف',[f'{r.emp_no} — {r.name}' for _,r in e.iterrows()]);txt=st.text_area('الملاحظة')
            if st.form_submit_button('حفظ الملاحظة') and txt.strip():
                no=emp.split(' — ')[0];now=datetime.now().isoformat(timespec='seconds');c=conn();c.execute('INSERT INTO notes(emp_no,note,active,created_at,updated_at) VALUES(?,?,?,?,?)',(no,txt,1,now,now));c.commit();c.close();st.success('تم الحفظ.')
        st.dataframe(notes(),use_container_width=True)

elif page=='القواعد والقوالب':
    st.title('⚙️ القواعد والقوالب');asub=st.text_input('عنوان عربي',setting('subject_ar'));esub=st.text_input('English subject',setting('subject_en'));ab=st.text_area('القالب العربي',setting('body_ar'),height=250);eb=st.text_area('English template',setting('body_en'),height=250);st.caption('المتغيرات: {name} {month} {period} {details}')
    if st.button('حفظ القوالب'):
        for k,v in [('subject_ar',asub),('subject_en',esub),('body_ar',ab),('body_en',eb)]:set_setting(k,v)
        st.success('تم الحفظ.')

elif page=='مراجعة المسودات':
    st.title('📧 مراجعة المسودات');r=st.session_state.analysis
    if r.empty:st.info('حللي التقارير أولًا.')
    else:
        s=r[(~r.excluded)&r.email.astype(str).str.contains('@',na=False)]
        for (no,name,email),g in s.groupby(['emp_no','name','email']):
            details='\n'.join(f"- {x.date:%Y-%m-%d}: {x.case_type}"+(f" — مدة: {x.duration}" if x.duration not in ['', 'nan'] else '') for _,x in g.iterrows());month=g.date.min().strftime('%B %Y');sub=setting('subject_ar').format(month=month);body=setting('body_ar').format(name=name,month=month,period=f'{g.date.min():%Y-%m-%d} إلى {g.date.max():%Y-%m-%d}',details=details)
            with st.expander(f'📧 {name} — {email}'):
                if g.project.iloc[0]:st.caption('المشروع: '+str(g.project.iloc[0]))
                if g.note.iloc[0]:st.warning('ملاحظة: '+str(g.note.iloc[0]))
                body=st.text_area('نص الرسالة — قابل للتعديل',body,height=240,key='b'+str(no));sub=st.text_input('العنوان',sub,key='s'+str(no));st.link_button('فتح Gmail كمسودة للمراجعة فقط',gmail(email,sub,body));st.caption('لن يتم الإرسال تلقائيًا.')

elif page=='السجل':
    st.title('📜 السجل');c=conn();h=pd.read_sql_query('SELECT * FROM history ORDER BY id DESC',c);c.close();st.dataframe(h,use_container_width=True)

elif page=='نسخة احتياطية':
    st.title('💾 نسخة احتياطية');st.warning('احفظي نسخة دورية من قاعدة البيانات عند استخدام Streamlit Cloud.');
    if DB.exists():st.download_button('تحميل قاعدة البيانات',DB.read_bytes(),'attendance_system_backup.db')
