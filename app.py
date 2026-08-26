
import streamlit as st
import sqlite3, json, os, re, tempfile, shutil, threading
from datetime import date, datetime
from pathlib import Path
from io import BytesIO
from difflib import SequenceMatcher
import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import html
import textwrap
import requests

try:
    from supabase import create_client
except Exception:
    create_client=None

st.set_page_config(page_title="ChapLab Teacher Hub", page_icon="📘", layout="wide")

LEGACY_DB = Path("teacher_tracker.db")
WEB_DB = Path(tempfile.gettempdir()) / "chaplab_teacher_tracker.db"
UPLOAD_DIR = Path(tempfile.gettempdir()) / "chaplab_student_work"
UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
_CLOUD_LOCK = threading.Lock()

def _secret_section(name):
    try:
        return st.secrets[name]
    except Exception:
        return None

def cloud_config():
    cfg=_secret_section("supabase")
    if not cfg:
        return None
    try:
        url=str(cfg.get("url","")).strip()
        key=str(cfg.get("service_role_key","")).strip()
        bucket=str(cfg.get("bucket","chaplab-private")).strip() or "chaplab-private"
        db_object=str(cfg.get("database_object","teacher_tracker.db")).strip() or "teacher_tracker.db"
    except Exception:
        return None
    if not url or not key:
        return None
    return {"url":url,"key":key,"bucket":bucket,"db_object":db_object}

def auth_config():
    cfg=_secret_section("auth")
    if not cfg:
        return None
    username=str(cfg.get("username","")).strip()
    password=str(cfg.get("password",""))
    if not username or not password:
        return None
    return {"username":username,"password":password}

def cloud_configured():
    return cloud_config() is not None

def require_login():
    auth=auth_config()
    if cloud_configured() and not auth:
        st.error("ChapLab Web is connected to cloud storage, but login credentials are missing.")
        st.info("Add an [auth] section to Streamlit Secrets before using student data online.")
        st.stop()

    # Local mode remains available without a login.
    if not auth:
        return

    if st.session_state.get("chaplab_authenticated"):
        return

    st.markdown("## 📘 ChapLab Teacher Hub")
    st.caption("Private teacher sign-in")
    with st.form("chaplab_login"):
        username=st.text_input("Username")
        password=st.text_input("Password",type="password")
        submitted=st.form_submit_button("Sign In",use_container_width=True)
    if submitted:
        if username==auth["username"] and password==auth["password"]:
            st.session_state["chaplab_authenticated"]=True
            st.rerun()
        else:
            st.error("Username or password is incorrect.")
    st.stop()

require_login()

@st.cache_resource
def supabase_client():
    cfg=cloud_config()
    if not cfg:
        return None
    if create_client is None:
        raise RuntimeError("The Supabase package is not installed. Install requirements.txt and restart ChapLab.")
    return create_client(cfg["url"],cfg["key"])

@st.cache_resource
def ensure_cloud_bucket():
    if not cloud_configured():
        return False
    client=supabase_client()
    cfg=cloud_config()
    try:
        client.storage.create_bucket(
            cfg["bucket"],
            options={"public":False,"file_size_limit":52428800}
        )
    except Exception:
        # The normal case after first setup is that the private bucket already exists.
        pass
    return True

def cloud_download_bytes(remote_path):
    if not cloud_configured():
        return None
    ensure_cloud_bucket()
    cfg=cloud_config()
    try:
        return supabase_client().storage.from_(cfg["bucket"]).download(remote_path)
    except Exception:
        return None

def cloud_upload_bytes(data, remote_path, content_type="application/octet-stream"):
    if not cloud_configured():
        return False
    ensure_cloud_bucket()
    cfg=cloud_config()
    with _CLOUD_LOCK:
        try:
            supabase_client().storage.from_(cfg["bucket"]).upload(
                path=remote_path,
                file=data,
                file_options={
                    "cache-control":"0",
                    "upsert":"true",
                    "content-type":content_type
                }
            )
            return True
        except Exception as e:
            st.session_state["_cloud_sync_error"]=str(e)
            return False

def cloud_upload_file(local_path, remote_path=None):
    local_path=Path(local_path)
    if not local_path.exists() or not cloud_configured():
        return False
    cfg=cloud_config()
    remote_path=remote_path or cfg["db_object"]
    return cloud_upload_bytes(local_path.read_bytes(),remote_path,"application/x-sqlite3")

def cloud_download_database(target):
    if not cloud_configured():
        return False
    cfg=cloud_config()
    data=cloud_download_bytes(cfg["db_object"])
    if not data:
        return False
    target=Path(target)
    target.parent.mkdir(parents=True,exist_ok=True)
    target.write_bytes(data)
    return True

@st.cache_resource
def prepare_database():
    if not cloud_configured():
        return str(LEGACY_DB)

    ensure_cloud_bucket()

    # Preferred: the persistent cloud copy.
    if cloud_download_database(WEB_DB):
        return str(WEB_DB)

    # First migration from the user's existing local ChapLab database.
    if LEGACY_DB.exists():
        shutil.copy2(LEGACY_DB,WEB_DB)
        cloud_upload_file(WEB_DB)
        return str(WEB_DB)

    # Brand-new web installation. init_db() will create it, then commit syncs it.
    return str(WEB_DB)

DB = prepare_database()

def cloud_status_text():
    if not cloud_configured():
        return "Local mode"
    if st.session_state.get("_cloud_sync_error"):
        return "Cloud configured — last sync needs attention"
    return "Cloud connected"

# ---------- Composition notebook style ----------
st.markdown("""
<style>
:root{
  --navy:#123f8c;--navy-dark:#0d2f6a;--ink:#253044;--muted:#6c7482;
  --blue:#9fd8f3;--green:#b9df6f;--yellow:#f7d85c;--pink:#f29ab3;
  --purple:#bda0e5;--orange:#f4a64b;--line:#e4e8ef;
}
.stApp{
  background:linear-gradient(rgba(255,255,255,.97),rgba(255,255,255,.97)),
             repeating-linear-gradient(to bottom,#fff 0,#fff 31px,#e5edf5 32px,#fff 33px);
  color:var(--ink);
}
header[data-testid="stHeader"]{
  background:rgba(255,255,255,.94);backdrop-filter:blur(8px);border-bottom:1px solid #edf0f4;
}
section[data-testid="stSidebar"],[data-testid="collapsedControl"]{display:none!important}
.block-container{
  max-width:1450px!important;margin:0 auto!important;padding:120px 34px 52px 300px!important;
  background:transparent!important;border:none!important;box-shadow:none!important;
}
.teacher-sidebar{
  position:fixed;left:0;top:0;bottom:0;width:260px;
  background:linear-gradient(180deg,var(--navy),var(--navy-dark));color:white;z-index:10000;
  padding:28px 18px 24px;box-sizing:border-box;box-shadow:5px 0 18px rgba(0,0,0,.12);overflow-y:auto;
}
.teacher-avatar{
  width:76px;height:76px;border-radius:50%;margin:8px auto 10px;display:flex;align-items:center;
  justify-content:center;background:white;color:var(--navy);border:4px solid #d7b5ff;font-size:25px;font-weight:900;
}
.teacher-brand{
  text-align:center;font-family:"Comic Sans MS","Segoe Print",cursive;font-size:24px;font-weight:800;
  line-height:1.1;margin-bottom:24px;
}
.nav-link{
  display:flex;align-items:center;gap:10px;padding:12px 14px;margin:5px 0;border-radius:10px;
  text-decoration:none!important;color:white!important;font-weight:700;font-size:14px;
}
.nav-link:hover{background:rgba(255,255,255,.13)}
.nav-link.active{background:linear-gradient(90deg,#7669eb,#5f65e8)}
.sidebar-note{
  margin-top:24px;background:linear-gradient(135deg,#f4dbff,#d9cbff);color:#2b2d3f;
  border-radius:18px;padding:16px;font-size:13px;line-height:1.4;
}
.class-bar{
  position:fixed;left:260px;right:0;top:48px;height:62px;background:rgba(255,255,255,.97);
  z-index:9000;display:flex;align-items:flex-end;gap:10px;padding:0 34px;
  border-bottom:1px solid #e6e9ef;box-shadow:0 3px 10px rgba(0,0,0,.04);
}
.class-bar-label{align-self:center;font-weight:800;color:#536072;margin-right:4px}
.class-pill{
  min-width:110px;padding:12px 18px;text-align:center;text-decoration:none!important;color:#253044!important;
  font-family:"Comic Sans MS","Segoe Print",cursive;font-weight:800;border:1px solid rgba(0,0,0,.08);
  border-radius:12px 12px 0 0;box-shadow:0 -2px 5px rgba(0,0,0,.04);
}
.class-pill:nth-of-type(2){background:var(--blue)}
.class-pill:nth-of-type(3){background:var(--green)}
.class-pill:nth-of-type(4){background:#f7f4ed}
.class-pill:nth-of-type(5){background:#d5e99a}
.class-pill.active{box-shadow:inset 0 -4px 0 #5f65e8}
.page-title{
  font-family:"Comic Sans MS","Segoe Print",cursive;font-size:32px;font-weight:900;color:#26344b;margin:0 0 6px;
}
.page-subtitle{color:var(--muted);margin-bottom:22px}
.hero-card{
  background:radial-gradient(circle at 8% 15%,rgba(255,255,255,.8),transparent 30%),
             linear-gradient(135deg,#fffaf1,#fffdf9);
  border:1px solid #e7dfd3;border-radius:22px;padding:28px 30px;box-shadow:0 7px 20px rgba(43,51,70,.08);
  margin-bottom:24px;position:relative;overflow:hidden;
}
.hero-card:after{
  content:"✦   ♡   ☆";position:absolute;right:28px;top:18px;font-size:28px;color:#b399d7;
  font-family:"Comic Sans MS",cursive;
}
.hero-card h1{
  margin:0;font-family:"Comic Sans MS","Segoe Print",cursive;font-size:38px;color:#27344a;
}
.hero-card .accent{
  display:inline-block;margin-top:8px;background:#c5a8e6;padding:6px 18px 8px;border-radius:8px;
  font-family:"Comic Sans MS","Segoe Print",cursive;font-weight:800;
}
.sticky-grid{
  display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:20px;margin:18px 0 26px;
}
.sticky-card{
  min-height:145px;padding:18px;border-radius:6px;box-shadow:7px 8px 15px rgba(0,0,0,.11);
  display:flex;flex-direction:column;justify-content:center;text-align:center;
  font-family:"Comic Sans MS","Segoe Print",cursive;color:#252a33;
}
.sticky-card strong{display:block;font-size:28px;margin-bottom:4px}
.sticky-card small{font-family:Arial,sans-serif;font-size:12px;line-height:1.3}
.s-pink{background:var(--pink)}.s-orange{background:var(--orange)}.s-yellow{background:var(--yellow)}
.s-green{background:var(--green)}.s-blue{background:var(--blue)}.s-purple{background:var(--purple)}
div[data-testid="stForm"],div[data-testid="stDataFrame"],div[data-testid="stExpander"],div[data-testid="stMetric"]{
  background:rgba(255,255,255,.97)!important;border-radius:12px!important;
}
div[data-testid="stMetric"]{border:1px solid var(--line)!important;box-shadow:0 2px 8px rgba(0,0,0,.05);padding:12px 14px}
input,textarea,[data-baseweb="select"]>div{background:#f7f9fc!important}
.profile-cover{
  height:180px;max-width:1050px;margin:0 auto;background:linear-gradient(135deg,#789dbb,#d8e8f3);
  border-radius:18px 18px 0 0;
}
.profile-header{
  max-width:1050px;margin:0 auto 18px;background:white;border:1px solid #dce2ea;border-top:0;
  border-radius:0 0 18px 18px;padding:0 24px 20px;
}
.profile-avatar{
  width:110px;height:110px;border-radius:50%;background:#31536a;color:white;display:flex;align-items:center;
  justify-content:center;font-size:36px;font-weight:900;border:5px solid white;box-shadow:0 0 0 3px #1d2830;margin-top:-58px;
}
.profile-section{
  max-width:1050px;margin:14px auto;background:white;border:1px solid #e0e5ec;border-radius:14px;
  padding:18px 20px;box-shadow:0 2px 8px rgba(0,0,0,.04);
}
@media(max-width:1050px){
  .teacher-sidebar{width:220px}.class-bar{left:220px}.block-container{padding-left:250px!important}
  .sticky-grid{grid-template-columns:repeat(2,minmax(0,1fr))}
}
@media(max-width:780px){
  .teacher-sidebar{display:none}.class-bar{left:0}.block-container{padding:120px 16px 40px!important}
  .sticky-grid{grid-template-columns:1fr}
}
</style>
""", unsafe_allow_html=True)

# ---------- Database ----------
class ChapLabConnection(sqlite3.Connection):
    def commit(self):
        result=super().commit()
        if cloud_configured():
            try:
                cloud_upload_file(DB)
                st.session_state.pop("_cloud_sync_error",None)
            except Exception as e:
                st.session_state["_cloud_sync_error"]=str(e)
        return result

def conn():
    c=sqlite3.connect(DB,factory=ChapLabConnection)
    c.row_factory=sqlite3.Row
    c.execute("PRAGMA foreign_keys=ON")
    return c

def cols(cur, table):
    return {r[1] for r in cur.execute(f"PRAGMA table_info({table})").fetchall()}

def addcol(cur, table, name, decl):
    table_exists=cur.execute(
        "SELECT 1 FROM sqlite_master WHERE type='table' AND name=?",
        (table,)
    ).fetchone()
    if not table_exists:
        return
    if name not in cols(cur, table):
        cur.execute(f"ALTER TABLE {table} ADD COLUMN {name} {decl}")

def init_db():
    c=conn(); cur=c.cursor()
    cur.executescript("""
    CREATE TABLE IF NOT EXISTS classes(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        class_name TEXT NOT NULL UNIQUE,
        subject_note TEXT DEFAULT '',
        active INTEGER DEFAULT 1
    );
    CREATE TABLE IF NOT EXISTS scholars(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        first_name TEXT NOT NULL,
        last_name TEXT NOT NULL,
        class_name TEXT DEFAULT '',
        gender TEXT DEFAULT '',
        pronouns TEXT DEFAULT '',
        active INTEGER DEFAULT 1
    );
    CREATE TABLE IF NOT EXISTS guardians(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        scholar_id INTEGER NOT NULL,
        first_name TEXT DEFAULT '',
        last_name TEXT DEFAULT '',
        relationship TEXT DEFAULT '',
        home_phone TEXT DEFAULT '',
        work_phone TEXT DEFAULT '',
        cell_phone TEXT DEFAULT '',
        email TEXT DEFAULT '',
        UNIQUE(scholar_id, first_name, last_name, relationship)
    );
    CREATE TABLE IF NOT EXISTS standards(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        subject TEXT NOT NULL,
        code TEXT NOT NULL,
        skill TEXT NOT NULL,
        description TEXT DEFAULT '',
        UNIQUE(subject, code)
    );
    CREATE TABLE IF NOT EXISTS assignments(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        title TEXT NOT NULL,
        subject TEXT NOT NULL,
        category TEXT NOT NULL,
        standard_code TEXT DEFAULT '',
        points_possible REAL DEFAULT 100,
        assignment_date TEXT
    );
    CREATE TABLE IF NOT EXISTS grades(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        scholar_id INTEGER NOT NULL,
        assignment_id INTEGER NOT NULL,
        points_earned REAL,
        UNIQUE(scholar_id, assignment_id)
    );
    CREATE TABLE IF NOT EXISTS settings(key TEXT PRIMARY KEY,value TEXT);
    CREATE TABLE IF NOT EXISTS communications(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        scholar_id INTEGER, guardian_id INTEGER,
        created_at TEXT, communication_type TEXT,
        subject TEXT, reason TEXT, details TEXT, generated_text TEXT
    );
    CREATE TABLE IF NOT EXISTS work_samples(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        scholar_id INTEGER NOT NULL, uploaded_at TEXT NOT NULL,
        subject TEXT DEFAULT '', skill_code TEXT DEFAULT '', title TEXT DEFAULT '',
        file_name TEXT DEFAULT '', file_path TEXT DEFAULT '',
        teacher_observation TEXT DEFAULT '', strengths TEXT DEFAULT '',
        needs TEXT DEFAULT '', next_steps TEXT DEFAULT ''
    );
    CREATE TABLE IF NOT EXISTS support_notes(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        scholar_id INTEGER NOT NULL, created_at TEXT NOT NULL,
        note_type TEXT DEFAULT '', area TEXT DEFAULT '', observation TEXT DEFAULT '',
        frequency TEXT DEFAULT '', intervention TEXT DEFAULT '',
        response_to_intervention TEXT DEFAULT '', impact TEXT DEFAULT ''
    );
    CREATE TABLE IF NOT EXISTS report_comments(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        scholar_id INTEGER NOT NULL, created_at TEXT NOT NULL,
        subject TEXT DEFAULT '', marking_period TEXT DEFAULT '',
        comment_text TEXT DEFAULT ''
    );
    CREATE TABLE IF NOT EXISTS contact_reminders(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        scholar_id INTEGER NOT NULL,
        guardian_id INTEGER,
        due_date TEXT DEFAULT '',
        reason TEXT DEFAULT '',
        notes TEXT DEFAULT '',
        completed INTEGER DEFAULT 0,
        completed_date TEXT DEFAULT ''
    );
    CREATE TABLE IF NOT EXISTS quarter_settings(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        academic_year TEXT DEFAULT '',
        quarter_name TEXT NOT NULL,
        start_date TEXT DEFAULT '',
        end_date TEXT DEFAULT '',
        locked INTEGER DEFAULT 0,
        UNIQUE(academic_year, quarter_name)
    );

    CREATE TABLE IF NOT EXISTS report_card_deadlines(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        academic_year TEXT DEFAULT '',
        quarter_name TEXT NOT NULL,
        due_date TEXT DEFAULT '',
        UNIQUE(academic_year, quarter_name)
    );

    CREATE TABLE IF NOT EXISTS parent_update_preferences(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        scholar_id INTEGER NOT NULL UNIQUE,
        requested_updates INTEGER DEFAULT 0,
        update_frequency TEXT DEFAULT '',
        preferred_guardian_id INTEGER,
        notes TEXT DEFAULT ''
    );

    CREATE TABLE IF NOT EXISTS quarter_closeout(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        scholar_id INTEGER NOT NULL,
        quarter_name TEXT NOT NULL,
        academic_year TEXT DEFAULT '',
        grades_reviewed INTEGER DEFAULT 0,
        data_reviewed INTEGER DEFAULT 0,
        comment_generated INTEGER DEFAULT 0,
        comment_finalized INTEGER DEFAULT 0,
        UNIQUE(scholar_id, quarter_name, academic_year)
    );
    CREATE TABLE IF NOT EXISTS benchmark_scores(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        scholar_id INTEGER NOT NULL UNIQUE,
        nwea_fall_reading TEXT DEFAULT '',
        nwea_spring_reading TEXT DEFAULT '',
        nwea_fall_math TEXT DEFAULT '',
        nwea_spring_math TEXT DEFAULT '',
        fp_fall_level TEXT DEFAULT '',
        fp_spring_level TEXT DEFAULT '',
        fp_fall_word_list TEXT DEFAULT '',
        fp_spring_word_list TEXT DEFAULT '',
        notes TEXT DEFAULT ''
    );

    CREATE TABLE IF NOT EXISTS book_levels(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        title TEXT NOT NULL,
        author TEXT DEFAULT '',
        fp_level TEXT DEFAULT '',
        notes TEXT DEFAULT '',
        UNIQUE(title, author)
    );
    CREATE TABLE IF NOT EXISTS nwea_goals(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        class_id INTEGER,
        season TEXT NOT NULL,
        subject TEXT NOT NULL,
        goal_score REAL,
        UNIQUE(class_id, season, subject)
    );
    CREATE TABLE IF NOT EXISTS grouping_settings(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        class_id INTEGER,
        low_max REAL DEFAULT 69,
        mid_max REAL DEFAULT 84,
        UNIQUE(class_id)
    );
    """)
    addcol(cur,"scholars","class_id","INTEGER")
    addcol(cur,"scholars","school_name","TEXT DEFAULT ''")
    addcol(cur,"scholars","academic_year","TEXT DEFAULT ''")
    addcol(cur,"scholars","grade_level","TEXT DEFAULT ''")
    addcol(cur,"scholars","student_id","TEXT DEFAULT ''")
    addcol(cur,"scholars","address","TEXT DEFAULT ''")
    addcol(cur,"scholars","city","TEXT DEFAULT ''")
    addcol(cur,"scholars","state_code","TEXT DEFAULT ''")
    addcol(cur,"scholars","zip_code","TEXT DEFAULT ''")
    addcol(cur,"scholars","residency","TEXT DEFAULT ''")
    addcol(cur,"scholars","gender","TEXT DEFAULT ''")
    addcol(cur,"scholars","pronouns","TEXT DEFAULT ''")
    addcol(cur,"assignments","class_id","INTEGER")
    addcol(cur,"assignments","marking_period","TEXT DEFAULT ''")
    addcol(cur,"benchmark_scores","nwea_winter_reading","TEXT DEFAULT ''")
    addcol(cur,"benchmark_scores","nwea_winter_math","TEXT DEFAULT ''")
    addcol(cur,"benchmark_scores","nwea_reading_goal","TEXT DEFAULT ''")
    addcol(cur,"benchmark_scores","nwea_math_goal","TEXT DEFAULT ''")
    cur.execute("""
        CREATE TABLE IF NOT EXISTS book_catalog(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            title TEXT NOT NULL,
            author TEXT DEFAULT '',
            fp_level TEXT DEFAULT '',
            isbn TEXT DEFAULT '',
            notes TEXT DEFAULT ''
        )
    """)
    addcol(cur,"book_catalog","author","TEXT DEFAULT ''")
    addcol(cur,"book_catalog","fp_level","TEXT DEFAULT ''")
    addcol(cur,"book_catalog","isbn","TEXT DEFAULT ''")
    addcol(cur,"book_catalog","notes","TEXT DEFAULT ''")
    addcol(cur,"communications","guardian_id","INTEGER")
    addcol(cur,"support_notes","concern_category","TEXT DEFAULT ''")
    addcol(cur,"support_notes","frequency_choice","TEXT DEFAULT ''")
    addcol(cur,"support_notes","intervention_choice","TEXT DEFAULT ''")
    addcol(cur,"support_notes","response_choice","TEXT DEFAULT ''")
    addcol(cur,"support_notes","impact_choice","TEXT DEFAULT ''")

    # migrate old class_name values
    for r in cur.execute("""SELECT DISTINCT TRIM(class_name) n FROM scholars
                            WHERE class_name IS NOT NULL AND TRIM(class_name)<>''""").fetchall():
        cur.execute("INSERT OR IGNORE INTO classes(class_name) VALUES (?)",(r["n"],))
    cur.execute("""UPDATE scholars SET class_id=(SELECT id FROM classes WHERE class_name=scholars.class_name)
                   WHERE class_id IS NULL AND TRIM(COALESCE(class_name,''))<>''""")

    defaults={
        "weights":{"Assessment":40,"Classwork":30,"Project/Lab":20,"Homework":10},
        "scale":[["A+",97,100],["A",93,96.99],["A-",90,92.99],
                 ["B+",87,89.99],["B",83,86.99],["B-",80,82.99],
                 ["C+",77,79.99],["C",73,76.99],["C-",70,72.99],
                 ["D",65,69.99],["F",0,64.99]]
    }
    for k,v in defaults.items():
        cur.execute("INSERT OR IGNORE INTO settings(key,value) VALUES (?,?)",(k,json.dumps(v)))
    c.commit(); c.close()

def get_setting(k):
    c=conn(); r=c.execute("SELECT value FROM settings WHERE key=?",(k,)).fetchone(); c.close()
    return json.loads(r["value"]) if r else None

def save_setting(k,v):
    c=conn(); c.execute("INSERT OR REPLACE INTO settings(key,value) VALUES (?,?)",(k,json.dumps(v))); c.commit(); c.close()

def classes_df(active=True):
    c=conn(); q="SELECT * FROM classes"+(" WHERE active=1" if active else "")+" ORDER BY class_name"
    df=pd.read_sql_query(q,c); c.close(); return df

def scholars_df(class_id=None, active=True):
    c=conn()
    q="""SELECT s.*, COALESCE(c.class_name,s.class_name,'') display_class
         FROM scholars s LEFT JOIN classes c ON c.id=s.class_id WHERE 1=1"""
    p=[]
    if active: q+=" AND s.active=1"
    if class_id: q+=" AND s.class_id=?"; p.append(int(class_id))
    q+=" ORDER BY s.last_name,s.first_name"
    df=pd.read_sql_query(q,c,params=p); c.close(); return df

def standards_df(subject=None):
    c=conn(); q="SELECT * FROM standards"; p=[]
    if subject: q+=" WHERE subject=?"; p=[subject]
    q+=" ORDER BY subject,code"
    df=pd.read_sql_query(q,c,params=p); c.close(); return df

def seed():
    data=[
      ("Science","3-PS2-1","Balanced and unbalanced forces","Investigate effects of balanced and unbalanced forces on motion."),
      ("Science","3-PS2-2","Patterns of motion","Use patterns of motion to make predictions."),
      ("Science","3-PS2-3","Electric and magnetic interactions","Determine cause-and-effect relationships of electric or magnetic interactions."),
      ("Science","3-PS2-4","Magnetic design problem","Apply scientific ideas about magnets to solve a problem."),
      ("Science","3-LS1-1","Life cycles","Develop models to describe unique and diverse life cycles."),
      ("Science","3-LS2-1","Social and group behavior","Construct an argument that some animals form groups that help members survive."),
      ("Science","3-LS3-1","Inheritance and variation of traits","Analyze data about inherited traits and variation."),
      ("Science","3-LS3-2","Environmental effects on traits","Use evidence that the environment can influence traits."),
      ("Science","3-LS4-1","Plant and animal extinction","Analyze fossil evidence of organisms and environments from long ago."),
      ("Science","3-LS4-2","Variation and survival","Explain how variation may provide survival advantages."),
      ("Science","3-LS4-3","Habitat and survival","Use evidence about how organisms survive in habitats."),
      ("Science","3-LS4-4","Environmental change solutions","Make a claim about the merit of a solution to environmental change."),
      ("Social Studies","SS3-A1","Develop questions","Develop questions about communities."),
      ("Social Studies","SS3-A2","Use evidence","Recognize and use primary and secondary source evidence."),
      ("Social Studies","SS3-A5","Make inferences","Identify inferences from evidence."),
      ("Social Studies","SS3-B1","Chronological reasoning","Explain relationships among events."),
      ("Social Studies","SS3-B3","Cause and effect","Identify causes and effects."),
      ("Social Studies","SS3-C1","Comparison and context","Compare communities."),
      ("Social Studies","SS3-D1","Geographic reasoning","Ask geographic questions."),
      ("Social Studies","SS3-D2","Maps and geographic tools","Use maps and geographic representations."),
      ("Social Studies","SS3-E1","Economic systems","Identify examples of scarcity, resources, production, and consumption."),
      ("Social Studies","SS3-F1","Civic participation","Participate respectfully in civic/classroom life."),
      ("Social Studies","SS3-COMM","Communities","Describe characteristics and organization of communities."),
      ("Social Studies","SS3-REG","Regions and environment","Explain how geography, climate, resources, and environment influence communities."),
      ("ELA","3R1","Questions and evidence","Develop and answer questions using relevant text evidence."),
      ("ELA","3R2","Theme or central idea","Determine theme/central idea and supporting details."),
      ("ELA","3R3","Characters, events, ideas","Describe and explain development of characters, events, or ideas."),
      ("ELA","3R4","Word meaning","Determine meaning of words and phrases."),
      ("ELA","3R5","Text structure","Explain how parts of a text build on one another."),
      ("ELA","3R6","Point of view / purpose","Discuss point of view and author's purpose."),
      ("ELA","3R7","Illustrations and text features","Explain how illustrations or text features contribute to meaning."),
      ("ELA","3W2","Informative writing","Write informative/explanatory texts."),
      ("ELA","3W3","Narrative writing","Write narratives with details and clear sequence."),
      ("ELA","3W5","Plan, revise, edit","Strengthen writing through planning, revision, and editing."),
      ("ELA","3SL1","Collaborative discussions","Participate effectively in collaborative discussions."),
      ("ELA","3L1","Grammar and usage","Use standard English grammar and usage."),
      ("ELA","3L2","Capitalization, punctuation, spelling","Use capitalization, punctuation, and spelling conventions."),
      ("Math","NY-3.OA.1","Interpret multiplication","Interpret products of whole numbers."),
      ("Math","NY-3.OA.2","Interpret division","Interpret whole-number quotients."),
      ("Math","NY-3.OA.3","Multiplication/division word problems","Solve multiplication and division word problems."),
      ("Math","NY-3.OA.4","Unknown-factor problems","Determine unknowns in multiplication/division equations."),
      ("Math","NY-3.OA.7","Multiply and divide within 100","Fluently multiply and divide within 100."),
      ("Math","NY-3.NBT.1","Rounding","Round whole numbers to nearest 10 or 100."),
      ("Math","NY-3.NBT.2","Add and subtract within 1,000","Fluently add and subtract within 1,000."),
      ("Math","NY-3.NF.1","Understand fractions","Understand fractions as quantities formed by unit fractions."),
      ("Math","NY-3.NF.2","Fractions on a number line","Represent fractions on a number line."),
      ("Math","NY-3.NF.3","Equivalent/comparable fractions","Explain equivalence and compare fractions."),
      ("Math","NY-3.MD.1","Time intervals","Tell/write time and solve interval problems."),
      ("Math","NY-3.MD.3","Graphs","Draw and interpret scaled graphs."),
      ("Math","NY-3.MD.5","Area","Understand area measurement."),
      ("Math","NY-3.MD.8","Perimeter","Solve perimeter problems."),
      ("Math","NY-3.G.1","Classify shapes","Classify shapes by attributes."),
      ("Math","NY-3.G.2","Partition shapes","Partition shapes into equal areas.")
    ]
    c=conn()
    for r in data:
        c.execute("INSERT OR IGNORE INTO standards(subject,code,skill,description) VALUES (?,?,?,?)",r)
    c.commit(); c.close()

init_db(); seed()

# ---------- Helpers ----------
def nm(row): return f"{row['first_name']} {row['last_name']}"

def pronoun_set_from_text(value):
    """Return subject/object/possessive/reflexive pronouns from a saved selection."""
    raw=str(value or "").strip().lower()
    if raw in {"she/her","she / her","female","girl"}:
        return {"subject":"she","object":"her","possessive":"her","possessive_pronoun":"hers","reflexive":"herself"}
    if raw in {"he/him","he / him","male","boy"}:
        return {"subject":"he","object":"him","possessive":"his","possessive_pronoun":"his","reflexive":"himself"}
    if raw in {"they/them","they / them","nonbinary","non-binary"}:
        return {"subject":"they","object":"them","possessive":"their","possessive_pronoun":"theirs","reflexive":"themself"}

    # Accept custom entries written as subject/object/possessive, e.g. ze/hir/hir.
    parts=[p.strip() for p in re.split(r"[/,]",raw) if p.strip()]
    if len(parts)>=2:
        subj=parts[0]
        obj=parts[1]
        poss=parts[2] if len(parts)>=3 else obj
        return {"subject":subj,"object":obj,"possessive":poss,"possessive_pronoun":poss,"reflexive":f"{obj}self"}

    # Neutral default when nothing is entered.
    return {"subject":"they","object":"them","possessive":"their","possessive_pronoun":"theirs","reflexive":"themself"}

def scholar_pronouns(sid):
    c=conn()
    r=c.execute("SELECT gender,pronouns FROM scholars WHERE id=?",(int(sid),)).fetchone()
    c.close()
    if not r:
        return pronoun_set_from_text("")
    return pronoun_set_from_text(r["pronouns"] or r["gender"] or "")

def cap_pronoun(p):
    return p[:1].upper()+p[1:] if p else p
def clean(v):
    if pd.isna(v): return ""
    s=str(v).strip()
    if s.endswith(".0") and s[:-2].isdigit(): s=s[:-2]
    return s

def norm_header(h):
    return re.sub(r'[^a-z0-9]','',str(h).lower())

ALIASES={
 "school_name":["schoolname","school"],
 "academic_year":["academicyear","schoolyear","academicy"],
 "grade_level":["gradelevel","academicgradelevel","grade"],
 "class_name":["course","coursesection","section","class","classname","studentla"],
 "student_id":["studentid","studentnumber","localid","sisid"],
 "student_first":["studentfirstname","studentfirst","studentfir","scholarfirstname","scholarfirst"],
 "student_last":["studentlastname","studentlast","studentla","scholarlastname","scholarlast"],
 "address":["address","streetaddress"],
 "city":["city"],
 "state_code":["statecode","state"],
 "zip_code":["zipcode","zip","postalcode"],
 "residency":["residency"],
 "guardian_first":["firstname","parentfirstname","guardianfirstname"],
 "guardian_last":["lastname","parentlastname","guardianlastname"],
 "relationship":["relationshiptypename","relationship","relation"],
 "home_phone":["homephone","homepho"],
 "work_phone":["workphone","workpho","workphor"],
 "cell_phone":["cellphone","mobilephone","cell"],
 "email":["email","emailaddress"]
}

def auto_mapping(columns):
    normalized={norm_header(c):c for c in columns}
    mapping={}
    for field,als in ALIASES.items():
        for a in als:
            if norm_header(a) in normalized:
                mapping[field]=normalized[norm_header(a)]
                break
    # Screenshot-style special handling:
    # Student first/last headers often start "StudentFir..." and "StudentLa..."
    for c in columns:
        n=norm_header(c)
        if "studentfir" in n and "student_first" not in mapping: mapping["student_first"]=c
        if ("studentlast" in n or n=="studentla") and "student_last" not in mapping: mapping["student_last"]=c
        if "relationship" in n and "relationship" not in mapping: mapping["relationship"]=c
    return mapping

def letter(avg):
    if avg is None: return ""
    for l,lo,hi in get_setting("scale"):
        if float(lo)<=avg<=float(hi): return l
    return ""

def summary(sid,subject):
    c=conn()
    df=pd.read_sql_query("""SELECT a.category,a.standard_code,a.points_possible,g.points_earned
                            FROM grades g JOIN assignments a ON a.id=g.assignment_id
                            WHERE g.scholar_id=? AND a.subject=? AND g.points_earned IS NOT NULL""",
                         c,params=[sid,subject]); c.close()
    if df.empty: return None,pd.DataFrame()
    df["pct"]=df.points_earned/df.points_possible*100
    weights=get_setting("weights")
    cats=df.groupby("category").pct.mean().to_dict()
    used={k:v for k,v in weights.items() if k in cats}
    tw=sum(used.values())
    avg=sum(cats[k]*(used[k]/tw) for k in used) if tw else df.pct.mean()
    skills=df[df.standard_code!=""].groupby("standard_code").pct.mean().reset_index()
    return avg,skills


def guardian_options_for_scholar(sid):
    c=conn()
    g=pd.read_sql_query("""SELECT * FROM guardians WHERE scholar_id=? ORDER BY relationship,last_name,first_name""",c,params=[sid])
    c.close()
    options={0:"— No parent/guardian selected —"}
    for _,r in g.iterrows():
        display=(" ".join(x for x in [r.first_name,r.last_name] if x)).strip() or "Unnamed Guardian"
        if r.relationship:
            display += f" ({r.relationship})"
        options[int(r.id)] = display
    return options

def delete_record(table, record_id):
    allowed={"classes","standards","assignments","work_samples","support_notes","report_comments","communications","guardians"}
    if table not in allowed:
        raise ValueError("Table not allowed.")
    c=conn()
    c.execute(f"DELETE FROM {table} WHERE id=?",(int(record_id),))
    c.commit(); c.close()

def record_select(df, label_col, key):
    if df.empty:
        return None
    ids=list(df["id"].astype(int))
    labels={int(r.id):str(r[label_col]) for _,r in df.iterrows()}
    return st.selectbox("Select record",ids,format_func=lambda x:labels.get(x,str(x)),key=key)


def comment_text(name,subject,avg,skills,next_skill):
    lookup={r.code:r.skill for _,r in standards_df(subject).iterrows()}
    high=[]; low=[]; mid=[]
    for _,r in skills.iterrows():
        item=(lookup.get(r.standard_code,r.standard_code),float(r.pct))
        if r.pct>=85:
            high.append(item)
        elif r.pct<75:
            low.append(item)
        else:
            mid.append(item)
    high=sorted(high,key=lambda x:x[1],reverse=True)[:2]
    low=sorted(low,key=lambda x:x[1])[:2]

    if high:
        strength_text=" and ".join(x[0].lower() for x in high)
        opening=f"{name} has shown particular strength in {strength_text}."
    elif avg is not None and avg>=80:
        opening=f"{name} is showing a solid overall understanding of the {subject.lower()} skills assessed this marking period."
    else:
        opening=f"{name} is continuing to build understanding of grade-level {subject.lower()} skills."

    if low:
        need_text=" and ".join(x[0].lower() for x in low)
        growth=f" An area that would benefit from additional practice is {need_text}."
        home=f" At home, you can support {name} by reviewing corrected work, asking {name} to explain answers in their own words, and completing short practice activities focused on these skills."
        teacher=f" In class, I will continue providing targeted practice, feedback, and opportunities for {name} to revisit these skills."
    else:
        growth=" Continued attention to accuracy, complete explanations, and consistent application of learned skills will help maintain this progress."
        home=f" At home, you can support {name} by discussing what was learned in class, reviewing notes or corrected work, and asking {name} to explain the thinking behind an answer."
        teacher=f" In class, I will continue challenging {name} to apply these skills independently and explain reasoning clearly."

    nxt=f" Our next instructional focus will include {next_skill.lower()}, and I encourage you to ask {name} to share what they are learning about these skills." if next_skill else ""
    return (opening+growth+home+teacher+nxt).strip()

def academic_summary_for_scholar(sid):
    strengths=[]
    needs=[]
    teacher_actions=[]
    for subj in ["ELA","Math","Science","Social Studies"]:
        avg,skills=summary(sid,subj)
        if skills.empty:
            continue
        lookup={r.code:r.skill for _,r in standards_df(subj).iterrows()}
        for _,r in skills.iterrows():
            skill=lookup.get(r.standard_code,r.standard_code)
            pct=float(r.pct)
            if pct>=85:
                strengths.append((subj,skill,pct))
            elif pct<75:
                needs.append((subj,skill,pct))

    strengths=sorted(strengths,key=lambda x:x[2],reverse=True)[:5]
    needs=sorted(needs,key=lambda x:x[2])[:5]

    for subj,skill,pct in needs[:4]:
        teacher_actions.append(f"Provide targeted {subj} practice in {skill.lower()}, review errors with the scholar, and check for understanding after reteaching.")

    if not teacher_actions:
        teacher_actions.append("Continue providing grade-level practice, feedback, and opportunities to explain thinking independently.")

    return strengths,needs,teacher_actions

def guardian_display_map(sid):
    c=conn()
    g=pd.read_sql_query("SELECT * FROM guardians WHERE scholar_id=? ORDER BY relationship,last_name,first_name",c,params=[sid])
    c.close()
    out={0:"— No parent/guardian selected —"}
    for _,r in g.iterrows():
        name=(" ".join(x for x in [r.first_name,r.last_name] if x)).strip() or "Unnamed Guardian"
        if r.relationship:
            name+=f" ({r.relationship})"
        out[int(r.id)]=name
    return out

def log_communication(sid,gid,comm_type,subject,reason,details,generated_text=""):
    c=conn()
    c.execute("""INSERT INTO communications(
        scholar_id,guardian_id,created_at,communication_type,subject,reason,details,generated_text)
        VALUES (?,?,?,?,?,?,?,?)""",
        (int(sid),int(gid) if gid else None,date.today().strftime("%m/%d/%Y"),
         comm_type,subject,reason,details,generated_text))
    c.commit(); c.close()

def reminder_rows(sid=None):
    c=conn()
    q="""SELECT contact_reminders.*, scholars.first_name||' '||scholars.last_name scholar,
         COALESCE(TRIM(guardians.first_name||' '||guardians.last_name),'') guardian,
         guardians.relationship
         FROM contact_reminders
         LEFT JOIN scholars ON scholars.id=contact_reminders.scholar_id
         LEFT JOIN guardians ON guardians.id=contact_reminders.guardian_id
         WHERE 1=1"""
    p=[]
    if sid:
        q+=" AND contact_reminders.scholar_id=?"
        p=[int(sid)]
    q+=" ORDER BY completed ASC,due_date ASC,id DESC"
    df=pd.read_sql_query(q,c,params=p)
    c.close()
    return df


def benchmark_for_scholar(sid):
    c=conn()
    r=c.execute("SELECT * FROM benchmark_scores WHERE scholar_id=?",(int(sid),)).fetchone()
    c.close()
    return r

def fp_to_num(level):
    if level is None:
        return None
    s=str(level).strip().upper()
    if not s:
        return None
    # Supports common F&P-style letter levels A-Z
    if len(s)==1 and "A" <= s <= "Z":
        return ord(s)-ord("A")+1
    return None

def compare_book_to_scholar(book_level, scholar_level):
    b=fp_to_num(book_level)
    s=fp_to_num(scholar_level)
    if b is None or s is None:
        return "I need both a verified book level and the scholar's current F&P level to compare them."
    diff=b-s
    if diff <= -2:
        return "This book is below the scholar's current F&P level and may work well for independent fluency practice."
    if diff == -1:
        return "This book is slightly below the scholar's current F&P level and is likely appropriate for independent reading."
    if diff == 0:
        return "This book matches the scholar's current F&P level and is a strong choice for independent reading."
    if diff == 1:
        return "This book is slightly above the scholar's current F&P level and may be appropriate with light teacher or family support."
    return "This book is above the scholar's current F&P level and is better suited for supported or instructional reading."

def current_fp_level(sid):
    r=benchmark_for_scholar(sid)
    if not r:
        return ""
    return (r["fp_spring_level"] or r["fp_fall_level"] or "").strip()


def parse_score(value):
    try:
        if value is None:
            return None
        s=str(value).strip()
        if not s:
            return None
        return float(s)
    except:
        return None

def get_nwea_goal(class_id, season, subject):
    c=conn()
    if class_id:
        r=c.execute("""SELECT goal_score FROM nwea_goals
                       WHERE class_id=? AND season=? AND subject=?""",
                    (int(class_id),season,subject)).fetchone()
    else:
        r=c.execute("""SELECT goal_score FROM nwea_goals
                       WHERE class_id IS NULL AND season=? AND subject=?""",
                    (season,subject)).fetchone()
    c.close()
    return float(r["goal_score"]) if r and r["goal_score"] is not None else None

def save_nwea_goal(class_id, season, subject, goal):
    c=conn()
    # SQLite NULL values do not participate in UNIQUE exactly the way a normal key does,
    # so handle All Classes/global goals manually.
    if class_id:
        existing=c.execute("""SELECT id FROM nwea_goals
                              WHERE class_id=? AND season=? AND subject=?""",
                           (int(class_id),season,subject)).fetchone()
    else:
        existing=c.execute("""SELECT id FROM nwea_goals
                              WHERE class_id IS NULL AND season=? AND subject=?""",
                           (season,subject)).fetchone()
    if existing:
        c.execute("UPDATE nwea_goals SET goal_score=? WHERE id=?",(float(goal),int(existing["id"])))
    else:
        c.execute("INSERT INTO nwea_goals(class_id,season,subject,goal_score) VALUES (?,?,?,?)",
                  (int(class_id) if class_id else None,season,subject,float(goal)))
    c.commit(); c.close()

def class_nwea_dataframe(class_id, season):
    roster=scholars_df(class_id or None)
    rows=[]
    for _,srow in roster.iterrows():
        br=benchmark_for_scholar(int(srow.id))
        reading=""
        math=""
        if br:
            if season=="Fall":
                reading=br["nwea_fall_reading"] or ""
                math=br["nwea_fall_math"] or ""
            else:
                reading=br["nwea_spring_reading"] or ""
                math=br["nwea_spring_math"] or ""
        rows.append({
            "Scholar":nm(srow),
            "Reading / ELA":parse_score(reading),
            "Math":parse_score(math)
        })
    return pd.DataFrame(rows)

def nwea_rank_summary(df, column):
    valid=df[df[column].notna()].copy()
    if valid.empty:
        return None,None
    max_score=valid[column].max()
    min_score=valid[column].min()
    highest=valid[valid[column]==max_score]["Scholar"].tolist()
    lowest=valid[valid[column]==min_score]["Scholar"].tolist()
    return (highest,max_score),(lowest,min_score)


def get_grouping_cutoffs(class_id):
    c=conn()
    if class_id:
        r=c.execute("SELECT low_max,mid_max FROM grouping_settings WHERE class_id=?",(int(class_id),)).fetchone()
    else:
        r=c.execute("SELECT low_max,mid_max FROM grouping_settings WHERE class_id IS NULL").fetchone()
    c.close()
    if r:
        return float(r["low_max"]),float(r["mid_max"])
    return 69.0,84.0

def save_grouping_cutoffs(class_id, low_max, mid_max):
    c=conn()
    if class_id:
        r=c.execute("SELECT id FROM grouping_settings WHERE class_id=?",(int(class_id),)).fetchone()
    else:
        r=c.execute("SELECT id FROM grouping_settings WHERE class_id IS NULL").fetchone()
    if r:
        c.execute("UPDATE grouping_settings SET low_max=?,mid_max=? WHERE id=?",(float(low_max),float(mid_max),int(r["id"])))
    else:
        c.execute("INSERT INTO grouping_settings(class_id,low_max,mid_max) VALUES (?,?,?)",
                  (int(class_id) if class_id else None,float(low_max),float(mid_max)))
    c.commit(); c.close()

def group_label(score, low_max, mid_max):
    if score is None or pd.isna(score):
        return "No Data"
    score=float(score)
    if score<=low_max:
        return "Low"
    if score<=mid_max:
        return "Mid"
    return "High"

def subject_grouping_df(class_id, subject, low_max, mid_max):
    roster=scholars_df(class_id or None)
    rows=[]
    for _,sr in roster.iterrows():
        avg,_=summary(int(sr.id),subject)
        rows.append({
            "Scholar":nm(sr),
            "Average":None if avg is None else round(float(avg),1),
            "Group":group_label(avg,low_max,mid_max)
        })
    if not rows:
        return pd.DataFrame(columns=["Scholar","Average","Group"])
    df=pd.DataFrame(rows)
    if "Group" not in df.columns:
        df["Group"]="No Data"
    return df

def skill_grouping_df(class_id, subject, skill_code, low_max, mid_max):
    roster=scholars_df(class_id or None)
    rows=[]
    for _,sr in roster.iterrows():
        _,skills=summary(int(sr.id),subject)
        score=None
        if not skills.empty and "standard_code" in skills.columns and skill_code in set(skills["standard_code"]):
            score=float(skills.loc[skills["standard_code"]==skill_code,"pct"].iloc[0])
        rows.append({
            "Scholar":nm(sr),
            "Skill Average":None if score is None else round(score,1),
            "Group":group_label(score,low_max,mid_max)
        })
    if not rows:
        return pd.DataFrame(columns=["Scholar","Skill Average","Group"])
    df=pd.DataFrame(rows)
    if "Group" not in df.columns:
        df["Group"]="No Data"
    return df


def get_teacher_name():
    c=conn()
    r=c.execute("SELECT value FROM settings WHERE key='teacher_name'").fetchone()
    c.close()
    if not r:
        return ""
    try:
        return json.loads(r["value"])
    except:
        return str(r["value"] or "")

def save_teacher_name(name):
    c=conn()
    c.execute("INSERT OR REPLACE INTO settings(key,value) VALUES ('teacher_name',?)",(json.dumps(name.strip()),))
    c.commit(); c.close()

def teacher_dashboard_info():
    c=conn()
    out={"homeroom":"","subjects":[]}
    for key in ["teacher_homeroom","teacher_subjects"]:
        r=c.execute("SELECT value FROM settings WHERE key=?",(key,)).fetchone()
        if r:
            try:
                val=json.loads(r["value"])
            except:
                val=r["value"]
            if key=="teacher_homeroom":
                out["homeroom"]=str(val or "")
            else:
                out["subjects"]=val if isinstance(val,list) else []
    c.close()
    return out

def save_teacher_dashboard_info(homeroom, subjects):
    c=conn()
    c.execute("INSERT OR REPLACE INTO settings(key,value) VALUES ('teacher_homeroom',?)",(json.dumps(str(homeroom or "").strip()),))
    c.execute("INSERT OR REPLACE INTO settings(key,value) VALUES ('teacher_subjects',?)",(json.dumps(list(subjects or [])),))
    c.commit(); c.close()

def class_assignments_df(class_id, subject_filter="All Subjects", sort_order="Oldest → Newest"):
    c=conn()
    q="SELECT * FROM assignments WHERE 1=1"
    p=[]
    if class_id:
        q+=" AND (class_id=? OR class_id IS NULL)"
        p.append(int(class_id))
    if subject_filter!="All Subjects":
        q+=" AND subject=?"
        p.append(subject_filter)
    if sort_order=="Newest → Oldest":
        q+=" ORDER BY assignment_date DESC,id DESC"
    elif sort_order=="Subject":
        q+=" ORDER BY subject,assignment_date,id"
    elif sort_order=="Standard Code":
        q+=" ORDER BY standard_code,assignment_date,id"
    else:
        q+=" ORDER BY assignment_date ASC,id ASC"
    df=pd.read_sql_query(q,c,params=p)
    c.close()
    return df

def gradebook_matrix(class_id, subject_filter="All Subjects", sort_order="Oldest → Newest"):
    roster=scholars_df(class_id or None)
    assignments=class_assignments_df(class_id,subject_filter,sort_order)
    if roster.empty:
        return pd.DataFrame(),assignments
    c=conn()
    rows=[]
    for _,sr in roster.iterrows():
        row={"Scholar":nm(sr)}
        for _,a in assignments.iterrows():
            label=f"{a.title}\\n{a.standard_code or 'No Standard'}"
            gr=c.execute("SELECT points_earned FROM grades WHERE scholar_id=? AND assignment_id=?",
                         (int(sr.id),int(a.id))).fetchone()
            if gr and gr["points_earned"] is not None and float(a.points_possible):
                row[label]=round(float(gr["points_earned"])/float(a.points_possible)*100,1)
            else:
                row[label]=None
        rows.append(row)
    c.close()
    return pd.DataFrame(rows),assignments

def make_grade_sheet_xlsx(class_id, class_name, subject_filter="All Subjects", sort_order="Oldest → Newest"):
    roster=scholars_df(class_id or None)
    assignments=class_assignments_df(class_id,subject_filter,sort_order)
    wb=Workbook()
    ws=wb.active
    ws.title="Grade Entry"
    ws["A1"]="ChapLab Grade Entry Sheet"
    ws["A1"].font=Font(size=16,bold=True)
    ws["A2"]="Class"; ws["B2"]=class_name
    ws["D2"]="Subject Filter"; ws["E2"]=subject_filter
    ws["A3"]="Instructions"
    ws["B3"]="Enter points earned only. Do not change scholar names or the hidden assignment ID row. Save and upload this file back into ChapLab."
    ws["A5"]="Scholar ID"; ws["B5"]="Scholar Name"
    ws["A6"]="Scholar ID"; ws["B6"]="Scholar Name"
    ws.row_dimensions[5].hidden=True
    for j,(_,a) in enumerate(assignments.iterrows(),start=3):
        col=get_column_letter(j)
        ws.cell(row=5,column=j,value=int(a.id))
        ws.cell(row=6,column=j,value=f"{a.title}\\n{a.standard_code or 'No Standard'}\\n{a.assignment_date or ''}\\n{a.points_possible:g} pts")
        ws.cell(row=6,column=j).alignment=Alignment(wrap_text=True,horizontal="center",vertical="center")
        ws.column_dimensions[col].width=18
    c=conn()
    for i,(_,sr) in enumerate(roster.iterrows(),start=7):
        ws.cell(row=i,column=1,value=int(sr.id))
        ws.cell(row=i,column=2,value=nm(sr))
        for j,(_,a) in enumerate(assignments.iterrows(),start=3):
            gr=c.execute("SELECT points_earned FROM grades WHERE scholar_id=? AND assignment_id=?",
                         (int(sr.id),int(a.id))).fetchone()
            if gr and gr["points_earned"] is not None:
                ws.cell(row=i,column=j,value=float(gr["points_earned"]))
    c.close()
    thin=Side(style="thin",color="444444")
    for cell in ws[6]:
        if cell.column<=2 or cell.value is not None:
            cell.fill=PatternFill("solid",fgColor="F4D35E")
            cell.font=Font(bold=True,color="111111")
            cell.border=Border(left=thin,right=thin,top=thin,bottom=thin)
    ws.row_dimensions[6].height=62
    max_col=max(2,2+len(assignments))
    for row in ws.iter_rows(min_row=7,max_row=6+len(roster),min_col=1,max_col=max_col):
        for cell in row:
            cell.border=Border(left=thin,right=thin,top=thin,bottom=thin)
            cell.alignment=Alignment(horizontal="left" if cell.column==2 else "center")
    ws.column_dimensions["A"].width=12
    ws.column_dimensions["B"].width=24
    ws.freeze_panes="C7"
    ws.sheet_view.showGridLines=False
    ws.page_setup.orientation="landscape"
    ws.page_setup.fitToWidth=1
    ws.page_setup.fitToHeight=0
    ws.sheet_properties.pageSetUpPr.fitToPage=True
    ws.print_title_rows="1:6"
    bio=BytesIO()
    wb.save(bio)
    bio.seek(0)
    return bio.getvalue()

def import_grade_sheet_xlsx(uploaded_file):
    wb=load_workbook(uploaded_file,data_only=True)
    ws=wb["Grade Entry"] if "Grade Entry" in wb.sheetnames else wb.active
    assignment_ids=[]
    for col in range(3,ws.max_column+1):
        aid=ws.cell(row=5,column=col).value
        if aid is not None:
            try:
                assignment_ids.append((col,int(aid)))
            except:
                pass
    saved=0; skipped=0
    c=conn()
    for row in range(7,ws.max_row+1):
        sid=ws.cell(row=row,column=1).value
        if sid is None:
            continue
        try:
            sid=int(sid)
        except:
            continue
        for col,aid in assignment_ids:
            value=ws.cell(row=row,column=col).value
            if value in (None,""):
                continue
            try:
                points=float(value)
            except:
                skipped+=1; continue
            ar=c.execute("SELECT points_possible FROM assignments WHERE id=?",(aid,)).fetchone()
            if not ar or points<0 or points>float(ar["points_possible"]):
                skipped+=1; continue
            c.execute("INSERT OR REPLACE INTO grades(scholar_id,assignment_id,points_earned) VALUES (?,?,?)",
                      (sid,aid,points))
            saved+=1
    c.commit(); c.close()
    return saved,skipped

def class_dashboard_metrics(class_id):
    roster=scholars_df(class_id or None)
    total=len(roster)
    contacts=0
    averages=[]
    c=conn()
    if not roster.empty:
        ids=[int(x) for x in roster.id]
        placeholders=",".join(["?"]*len(ids))
        contacts=c.execute(f"SELECT COUNT(*) n FROM communications WHERE scholar_id IN ({placeholders})",ids).fetchone()["n"]
    c.close()
    for _,sr in roster.iterrows():
        vals=[]
        for subj in ["ELA","Math","Science","Social Studies"]:
            av,_=summary(int(sr.id),subj)
            if av is not None:
                vals.append(av)
        if vals:
            averages.append(sum(vals)/len(vals))
    overall=sum(averages)/len(averages) if averages else None
    return total,contacts,overall


def class_name_from_id(class_id):
    if not class_id:
        return ""
    c=conn()
    r=c.execute("SELECT class_name FROM classes WHERE id=?",(int(class_id),)).fetchone()
    c.close()
    return r["class_name"] if r else ""

def scholar_full_profile(sid):
    c=conn()
    r=c.execute("""SELECT s.*, COALESCE(c.class_name,s.class_name,'') display_class
                   FROM scholars s LEFT JOIN classes c ON c.id=s.class_id
                   WHERE s.id=?""",(int(sid),)).fetchone()
    c.close()
    return r

def iep_prefill_summary(sid):
    s=scholar_full_profile(sid)
    if not s:
        return ""
    name=f"{s['first_name']} {s['last_name']}"
    parts=[
        f"Scholar: {name}",
        f"School ID: {s['student_id'] or '—'}",
        f"Grade/Class: {s['grade_level'] or '—'} / {s['display_class'] or '—'}",
        f"School Year: {s['academic_year'] or '—'}",
    ]
    # Grades
    grade_bits=[]
    for subj in ["ELA","Math","Science","Social Studies"]:
        av,_=summary(int(sid),subj)
        if av is not None:
            grade_bits.append(f"{subj} {av:.1f}% ({letter(av)})")
    if grade_bits:
        parts.append("Current grades: " + "; ".join(grade_bits))

    # Benchmarks
    br=benchmark_for_scholar(sid)
    if br:
        bench=[]
        if br["nwea_fall_reading"] or br["nwea_spring_reading"]:
            bench.append(f"NWEA Reading Fall {br['nwea_fall_reading'] or '—'} / Spring {br['nwea_spring_reading'] or '—'}")
        if br["nwea_fall_math"] or br["nwea_spring_math"]:
            bench.append(f"NWEA Math Fall {br['nwea_fall_math'] or '—'} / Spring {br['nwea_spring_math'] or '—'}")
        if br["fp_fall_level"] or br["fp_spring_level"]:
            bench.append(f"F&P Fall {br['fp_fall_level'] or '—'} / Spring {br['fp_spring_level'] or '—'}")
        if br["fp_fall_word_list"] or br["fp_spring_word_list"]:
            bench.append(f"F&P Word List Fall {br['fp_fall_word_list'] or '—'} / Spring {br['fp_spring_word_list'] or '—'}")
        if bench:
            parts.append("Benchmark data: " + "; ".join(bench))

    strengths,needs,teacher_actions=academic_summary_for_scholar(sid)
    if strengths:
        parts.append("Academic strengths: " + "; ".join(f"{subj} - {skill} ({pct:.0f}%)" for subj,skill,pct in strengths[:4]))
    if needs:
        parts.append("Areas needing support: " + "; ".join(f"{subj} - {skill} ({pct:.0f}%)" for subj,skill,pct in needs[:4]))

    c=conn()
    notes=pd.read_sql_query("SELECT * FROM support_notes WHERE scholar_id=? ORDER BY id DESC LIMIT 5",c,params=[sid])
    ws=pd.read_sql_query("SELECT * FROM work_samples WHERE scholar_id=? ORDER BY id DESC LIMIT 5",c,params=[sid])
    c.close()
    if not notes.empty:
        obs=[]
        for _,r in notes.iterrows():
            obs.append(f"{r['area']}: {r['observation']}")
        parts.append("Recent support/IEP observations: " + " | ".join(obs))
    if not ws.empty:
        evidence=[]
        for _,r in ws.iterrows():
            title=r["title"] or r["file_name"] or "Work sample"
            evidence.append(f"{title} — strengths: {r['strengths'] or '—'}; needs: {r['needs'] or '—'}")
        parts.append("Recent work-sample evidence: " + " | ".join(evidence))
    return "\n".join(parts)


def quarter_grade_data(sid, subject, marking_period):
    c=conn()
    df=pd.read_sql_query("""SELECT a.id,a.title,a.category,a.standard_code,a.points_possible,
                                   a.assignment_date,a.marking_period,g.points_earned
                            FROM grades g
                            JOIN assignments a ON a.id=g.assignment_id
                            WHERE g.scholar_id=? AND a.subject=? AND g.points_earned IS NOT NULL
                              AND COALESCE(a.marking_period,'')=?
                            ORDER BY a.assignment_date ASC,a.id ASC""",
                         c,params=[int(sid),subject,marking_period])
    c.close()
    if df.empty:
        return df
    df["pct"]=df["points_earned"]/df["points_possible"]*100
    return df

def weighted_average_from_grade_rows(df):
    if df.empty:
        return None
    weights=get_setting("weights")
    cat_avgs=df.groupby("category")["pct"].mean().to_dict()
    used={k:v for k,v in weights.items() if k in cat_avgs}
    totalw=sum(used.values())
    if totalw:
        return sum(cat_avgs[k]*(used[k]/totalw) for k in used)
    return float(df["pct"].mean())

def quarter_skill_summary(df, subject):
    if df.empty:
        return [],[]
    lookup={r.code:r.skill for _,r in standards_df(subject).iterrows()}
    skill_df=df[df["standard_code"].astype(str)!=""].groupby("standard_code")["pct"].mean().reset_index()
    strengths=[]
    needs=[]
    for _,r in skill_df.iterrows():
        skill=lookup.get(r["standard_code"],r["standard_code"])
        pct=float(r["pct"])
        if pct>=85:
            strengths.append((skill,pct))
        elif pct<75:
            needs.append((skill,pct))
    strengths=sorted(strengths,key=lambda x:x[1],reverse=True)[:3]
    needs=sorted(needs,key=lambda x:x[1])[:3]
    return strengths,needs

def quarter_grade_pattern(df):
    if df.empty:
        return {}
    result={}
    result["count"]=len(df)
    result["average"]=weighted_average_from_grade_rows(df)
    result["highest"]=df.loc[df["pct"].idxmax()]
    result["lowest"]=df.loc[df["pct"].idxmin()]
    result["first_half_avg"]=None
    result["second_half_avg"]=None
    result["trend"]=""

    if len(df)>=4:
        midpoint=max(1,len(df)//2)
        first=float(df.iloc[:midpoint]["pct"].mean())
        second=float(df.iloc[midpoint:]["pct"].mean())
        result["first_half_avg"]=first
        result["second_half_avg"]=second
        if second-first>=5:
            result["trend"]="improved over the course of the marking period"
        elif first-second>=5:
            result["trend"]="showed some decline later in the marking period"
        else:
            result["trend"]="performed fairly consistently across the marking period"

    cat_avgs=df.groupby("category")["pct"].mean().sort_values(ascending=False)
    result["category_best"]=(cat_avgs.index[0],float(cat_avgs.iloc[0])) if len(cat_avgs) else None
    result["category_need"]=(cat_avgs.index[-1],float(cat_avgs.iloc[-1])) if len(cat_avgs)>1 else None
    return result

def quarter_report_comment(name, subject, marking_period, df, next_skills, sid=None):
    pro=scholar_pronouns(sid) if sid else pronoun_set_from_text("")
    subj_pr=pro["subject"]
    poss_pr=pro["possessive"]
    if df.empty:
        return f"{name} does not yet have graded {subject.lower()} assignments tagged to {marking_period}. Add or update assignment marking periods before generating this comment."

    pattern=quarter_grade_pattern(df)
    strengths,needs=quarter_skill_summary(df,subject)
    avg=pattern["average"]
    grade_text=letter(avg) if avg is not None else ""

    opening=f"During {marking_period}, {name} earned a {avg:.1f}% ({grade_text}) in {subject} based on {pattern['count']} graded assignment"
    opening += "s." if pattern["count"]!=1 else "."

    evidence=[]
    if strengths:
        evidence.append("demonstrated strong understanding of " + " and ".join(x[0].lower() for x in strengths[:2]))
    if pattern.get("category_best"):
        cat,catavg=pattern["category_best"]
        evidence.append(f"performed especially well on {cat.lower()} work, averaging {catavg:.1f}%")
    if pattern.get("trend"):
        evidence.append(pattern["trend"])
    if evidence:
        evidence_sentence=f" {cap_pronoun(subj_pr)} " + ", and ".join(evidence) + "."
    else:
        evidence_sentence=""

    if needs:
        need_text=" and ".join(x[0].lower() for x in needs[:2])
        growth=f" An area for continued growth is {need_text}."
    elif pattern.get("category_need") and pattern["category_need"][1] < 80:
        cat,catavg=pattern["category_need"]
        growth=f" Continued practice with {cat.lower()} assignments, where the average was {catavg:.1f}%, will help strengthen overall performance."
    else:
        growth=" Continued attention to accuracy, complete explanations, and consistent application of learned skills will support continued progress."

    high=pattern["highest"]
    low=pattern["lowest"]
    assignment_evidence=""
    if len(df)>=2:
        assignment_evidence=(f" For example, {name}'s strongest recorded performance was on "
                             f"“{high['title']}” ({high['pct']:.1f}%), while “{low['title']}” "
                             f"({low['pct']:.1f}%) shows an area where additional review may be helpful.")

    if needs:
        home=f" At home, you can support {name} by reviewing corrected work connected to these skills, asking {subj_pr} to explain answers in {poss_pr} own words, and completing short practice activities that target the areas needing improvement."
    else:
        home=f" At home, you can support {name} by reviewing classwork, discussing what was learned, and asking {subj_pr} to explain the reasoning behind answers."

    teacher=f" In class, I will continue using grade and skill data to provide targeted feedback, reteaching, and opportunities for {name} to practice independently while building {poss_pr} confidence."

    nxt=""
    if next_skills:
        nxt=f" Our next instructional focus will include {next_skills.lower()}, and practicing these concepts at home will help {name} build confidence with upcoming work."

    return (opening+evidence_sentence+growth+assignment_evidence+home+teacher+nxt).strip()


def current_academic_year():
    # Prefer saved scholar academic year if available, else teacher setting
    c=conn()
    r=c.execute("SELECT academic_year FROM scholars WHERE TRIM(COALESCE(academic_year,''))<>'' LIMIT 1").fetchone()
    c.close()
    return r["academic_year"] if r else ""

def get_quarter_settings(academic_year=""):
    c=conn()
    df=pd.read_sql_query("""SELECT * FROM quarter_settings
                            WHERE academic_year=?
                            ORDER BY CASE quarter_name
                              WHEN 'Quarter 1' THEN 1
                              WHEN 'Quarter 2' THEN 2
                              WHEN 'Quarter 3' THEN 3
                              WHEN 'Quarter 4' THEN 4
                              ELSE 5 END""",c,params=[academic_year])
    c.close()
    return df

def save_quarter_setting(academic_year, quarter_name, start_date, end_date, locked=False):
    c=conn()
    c.execute("""INSERT INTO quarter_settings(academic_year,quarter_name,start_date,end_date,locked)
                 VALUES (?,?,?,?,?)
                 ON CONFLICT(academic_year,quarter_name) DO UPDATE SET
                 start_date=excluded.start_date,end_date=excluded.end_date,locked=excluded.locked""",
              (academic_year,quarter_name,start_date,end_date,1 if locked else 0))
    c.commit(); c.close()

def quarter_for_date(date_text, academic_year=""):
    if not date_text:
        return ""
    try:
        d=pd.to_datetime(date_text).date()
    except:
        return ""
    qdf=get_quarter_settings(academic_year)
    for _,r in qdf.iterrows():
        try:
            start=pd.to_datetime(r["start_date"]).date()
            end=pd.to_datetime(r["end_date"]).date()
            if start <= d <= end:
                return r["quarter_name"]
        except:
            pass
    return ""

def report_deadline(academic_year, quarter_name):
    c=conn()
    r=c.execute("""SELECT due_date FROM report_card_deadlines
                   WHERE academic_year=? AND quarter_name=?""",
                (academic_year,quarter_name)).fetchone()
    c.close()
    return r["due_date"] if r else ""

def save_report_deadline(academic_year, quarter_name, due_date):
    c=conn()
    c.execute("""INSERT INTO report_card_deadlines(academic_year,quarter_name,due_date)
                 VALUES (?,?,?)
                 ON CONFLICT(academic_year,quarter_name) DO UPDATE SET due_date=excluded.due_date""",
              (academic_year,quarter_name,due_date))
    c.commit(); c.close()

def update_preference(sid):
    c=conn()
    r=c.execute("SELECT * FROM parent_update_preferences WHERE scholar_id=?",(int(sid),)).fetchone()
    c.close()
    return r

def save_update_preference(sid, requested, frequency, guardian_id, notes):
    c=conn()
    c.execute("""INSERT INTO parent_update_preferences(
                    scholar_id,requested_updates,update_frequency,preferred_guardian_id,notes)
                 VALUES (?,?,?,?,?)
                 ON CONFLICT(scholar_id) DO UPDATE SET
                    requested_updates=excluded.requested_updates,
                    update_frequency=excluded.update_frequency,
                    preferred_guardian_id=excluded.preferred_guardian_id,
                    notes=excluded.notes""",
              (int(sid),1 if requested else 0,frequency,int(guardian_id) if guardian_id else None,notes))
    c.commit(); c.close()

def latest_subject_average(sid, subject):
    av,_=summary(sid,subject)
    return av

def subject_growth_by_quarter(sid, subject, academic_year=""):
    out=[]
    for q in ["Quarter 1","Quarter 2","Quarter 3","Quarter 4"]:
        qdf=quarter_grade_data(sid,subject,q)
        if not qdf.empty:
            out.append((q,weighted_average_from_grade_rows(qdf)))
    return out

def scholar_risk_and_growth(sid):
    strengths=[]
    risks=[]
    growth_notes=[]
    for subj in ["ELA","Math","Science","Social Studies"]:
        av,_=summary(sid,subj)
        if av is not None:
            if av>=85:
                strengths.append(f"{subj} ({av:.1f}%)")
            elif av<70:
                risks.append(f"{subj} ({av:.1f}%)")
        seq=subject_growth_by_quarter(sid,subj)
        if len(seq)>=2:
            prev=seq[-2][1]
            cur=seq[-1][1]
            delta=cur-prev
            if delta>=5:
                growth_notes.append(f"{subj} improved by {delta:.1f} points from {seq[-2][0]} to {seq[-1][0]}")
            elif delta<=-5:
                risks.append(f"{subj} dropped by {abs(delta):.1f} points from {seq[-2][0]} to {seq[-1][0]}")
    return strengths,risks,growth_notes

def generate_parent_progress_update(sid, subject="General"):
    s=scholar_full_profile(sid)
    if not s:
        return ""
    name=f"{s['first_name']} {s['last_name']}"
    strengths,risks,growth_notes=scholar_risk_and_growth(sid)
    skill_strengths,skill_needs,teacher_actions=academic_summary_for_scholar(sid)

    parts=[f"I wanted to share a progress update about {name}."]
    if subject!="General":
        av=latest_subject_average(sid,subject)
        if av is not None:
            parts.append(f"{name}'s current {subject} average is {av:.1f}% ({letter(av)}).")

    if growth_notes:
        parts.append("A positive trend I am seeing is " + "; ".join(growth_notes[:2]) + ".")
    elif strengths:
        pro=scholar_pronouns(sid)
        parts.append(f"{name} is currently showing strength in " + ", ".join(strengths[:2]) + f", and {pro['subject']} should continue building on this progress.")

    if skill_strengths:
        skilltxt=", ".join(f"{subj}: {skill}" for subj,skill,pct in skill_strengths[:2])
        parts.append(f"Specific strengths include {skilltxt}.")

    if risks:
        parts.append("An area I am watching closely is " + ", ".join(risks[:2]) + ".")
    elif skill_needs:
        needtxt=", ".join(f"{subj}: {skill}" for subj,skill,pct in skill_needs[:2])
        parts.append(f"Skills that would benefit from more practice include {needtxt}.")

    if skill_needs:
        need_names=", ".join(skill.lower() for _,skill,_ in skill_needs[:2])
        parts.append(f"At home, you can help by reviewing corrected work, asking {name} to explain answers aloud, and completing short practice focused on {need_names}.")
    else:
        parts.append(f"At home, you can help by asking {name} to explain what was learned, review classwork, and practice applying skills independently.")

    if teacher_actions:
        if teacher_actions:
            action=teacher_actions[0]
            parts.append("At school, I will continue to " + action[0].lower() + action[1:])
    return " ".join(p for p in parts if p).strip()

def quarter_missing_grade_check(sid, subject, quarter, class_id=None):
    c=conn()
    q="""SELECT a.id,a.title,a.points_possible
         FROM assignments a
         WHERE a.subject=? AND COALESCE(a.marking_period,'')=?"""
    p=[subject,quarter]
    if class_id:
        q+=" AND (a.class_id=? OR a.class_id IS NULL)"
        p.append(int(class_id))
    assignments=pd.read_sql_query(q,c,params=p)
    missing=[]
    for _,a in assignments.iterrows():
        g=c.execute("SELECT points_earned FROM grades WHERE scholar_id=? AND assignment_id=?",
                    (int(sid),int(a.id))).fetchone()
        if not g or g["points_earned"] is None:
            missing.append(a["title"])
    c.close()
    return missing, len(assignments)

def closeout_row(sid, quarter, academic_year):
    c=conn()
    r=c.execute("""SELECT * FROM quarter_closeout
                   WHERE scholar_id=? AND quarter_name=? AND academic_year=?""",
                (int(sid),quarter,academic_year)).fetchone()
    c.close()
    return r

def save_closeout(sid, quarter, academic_year, grades_reviewed, data_reviewed, comment_generated, comment_finalized):
    c=conn()
    c.execute("""INSERT INTO quarter_closeout(
                    scholar_id,quarter_name,academic_year,grades_reviewed,data_reviewed,comment_generated,comment_finalized)
                 VALUES (?,?,?,?,?,?,?)
                 ON CONFLICT(scholar_id,quarter_name,academic_year) DO UPDATE SET
                    grades_reviewed=excluded.grades_reviewed,
                    data_reviewed=excluded.data_reviewed,
                    comment_generated=excluded.comment_generated,
                    comment_finalized=excluded.comment_finalized""",
              (int(sid),quarter,academic_year,1 if grades_reviewed else 0,1 if data_reviewed else 0,
               1 if comment_generated else 0,1 if comment_finalized else 0))
    c.commit(); c.close()


def go_to_page(page_name):
    st.session_state["nav_page"]=page_name

def go_to_section(page_name, binder_tool=None, assistant_tool=None, show_class_settings=False, show_quarter_settings=False):
    st.session_state["nav_page"]=page_name
    if binder_tool is not None:
        st.session_state["class_binder_tool"]=binder_tool
    if assistant_tool is not None:
        st.session_state["assistant_tool"]=assistant_tool
    if show_class_settings:
        st.session_state["show_home_class_settings"]=True
    if show_quarter_settings:
        st.session_state["show_quarter_settings"]=True

def select_class(class_id):
    st.session_state["selected_class"]=int(class_id)
    st.session_state["nav_page"]="Class Dashboard"
    st.session_state["class_binder_tool"]="Overview"

def open_main_dashboard():
    st.session_state["nav_page"]="Home Page"

def open_side_section(section_name):
    mapping={
        "main":"Home Page",
        "scholars":"Scholars",
        "grades":"Scholar Binder",
        "books":"Book Leveler",
        "grouping":"Student Grouping",
        "reports":"Report Card Comments",
        "assistant":"Little Assistant",
        "communication":"Communication Log",
        "web":"Web & Backup",
    }
    st.session_state["nav_page"]=mapping.get(section_name,"Main Dashboard")
    if section_name=="grades":
        st.session_state["class_binder_tool"]="Overview"

def open_scholar_profile(sid):
    st.session_state["selected_profile_scholar"]=int(sid)
    st.session_state["nav_page"]="Scholar Profile"

def return_to_scholars():
    st.session_state["nav_page"]="Scholars"

def overall_scholar_average(sid):
    vals=[]
    for subj in ["ELA","Math","Science","Social Studies"]:
        av,_=summary(int(sid),subj)
        if av is not None:
            vals.append(float(av))
    return (sum(vals)/len(vals)) if vals else None

def scholar_status_summary(class_id=None):
    roster=scholars_df(class_id or None)
    low_max,mid_max=get_grouping_cutoffs(class_id)
    counts={"On Track":0,"Approaching":0,"At Risk":0,"No Data":0}
    detail=[]
    for _,sr in roster.iterrows():
        avg=overall_scholar_average(int(sr.id))
        if avg is None:
            status="No Data"
        elif avg<=low_max:
            status="At Risk"
        elif avg<=mid_max:
            status="Approaching"
        else:
            status="On Track"
        counts[status]+=1
        detail.append({"Scholar":nm(sr),"Average":None if avg is None else round(avg,1),"Status":status})
    return counts,pd.DataFrame(detail)

def recent_assignments_df(limit=5):
    c=conn()
    df=pd.read_sql_query("""SELECT a.*,c.class_name
                            FROM assignments a
                            LEFT JOIN classes c ON c.id=a.class_id
                            ORDER BY COALESCE(a.assignment_date,'') DESC,a.id DESC
                            LIMIT ?""",c,params=[int(limit)])
    c.close()
    return df

def home_announcements():
    items=[]
    ay=current_academic_year()
    for q in ["Quarter 1","Quarter 2","Quarter 3","Quarter 4"]:
        due=report_deadline(ay,q)
        if due:
            items.append(("Report Cards",f"{q} comments due",due))
    rem=reminder_rows()
    if not rem.empty:
        for _,r in rem[rem.completed==0].head(4).iterrows():
            items.append(("Parent Update",f"{r.scholar}: {r.reason}",r.due_date or ""))
    return items[:5]


def configure_local_tesseract():
    """Find a local Windows Tesseract install. Returns (ok, message)."""
    try:
        import pytesseract
        candidates=[
            r"C:\Program Files\Tesseract-OCR\tesseract.exe",
            r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
        ]
        for path in candidates:
            if os.path.exists(path):
                pytesseract.pytesseract.tesseract_cmd=path
                return True,path
        try:
            _=pytesseract.get_tesseract_version()
            return True,"PATH"
        except:
            return False,"Tesseract OCR is not installed or could not be found."
    except Exception as e:
        return False,str(e)

def normalize_ocr_text(value):
    return re.sub(r'[^a-z0-9 ]',' ',str(value).lower()).strip()

def name_similarity(a,b):
    a=normalize_ocr_text(a)
    b=normalize_ocr_text(b)
    if not a or not b:
        return 0
    return SequenceMatcher(None,a,b).ratio()

def ocr_grade_screenshot(image_file, roster, expected_assignments):
    """
    Extract likely scholar rows and numeric grades from a screenshot.
    expected_assignments is an ordered list of assignment rows/records.
    """
    from PIL import Image, ImageEnhance, ImageFilter
    import pytesseract
    from pytesseract import Output

    ok,msg=configure_local_tesseract()
    if not ok:
        raise RuntimeError(
            "Local OCR is not ready. Install Tesseract OCR for Windows, then restart ChapLab. "
            "The screenshot stays on your computer."
        )

    image=Image.open(image_file).convert("L")
    # Enlarge small screenshots and improve contrast.
    if image.width < 1800:
        scale=max(1.0,1800/image.width)
        image=image.resize((int(image.width*scale),int(image.height*scale)))
    image=ImageEnhance.Contrast(image).enhance(1.8)
    image=image.filter(ImageFilter.SHARPEN)

    data=pytesseract.image_to_data(image,output_type=Output.DATAFRAME,config="--psm 6")
    data=data.dropna(subset=["text"])
    data["text"]=data["text"].astype(str).str.strip()
    data=data[data["text"]!=""]
    if data.empty:
        return pd.DataFrame(), "No readable text was detected."

    # Group OCR words into visual lines.
    line_cols=["block_num","par_num","line_num"]
    lines=[]
    for _,group in data.groupby(line_cols,sort=False):
        group=group.sort_values("left")
        words=group["text"].tolist()
        joined=" ".join(words)
        top=float(group["top"].min())
        left=float(group["left"].min())
        right=float((group["left"]+group["width"]).max())
        lines.append({"text":joined,"words":words,"top":top,"left":left,"right":right,"group":group})

    roster_records=[]
    for _,r in roster.iterrows():
        roster_records.append({
            "Scholar ID":int(r.id),
            "Scholar":nm(r),
            "first":str(r.first_name),
            "last":str(r.last_name)
        })

    results=[]
    n_assign=len(expected_assignments)
    for line in lines:
        line_text=line["text"]
        norm=normalize_ocr_text(line_text)

        # Find best roster match using full name and last name.
        best=None
        best_score=0
        for rr in roster_records:
            full_score=name_similarity(line_text,rr["Scholar"])
            last_score=1.0 if normalize_ocr_text(rr["last"]) in norm and len(rr["last"])>=3 else name_similarity(line_text,rr["last"])*0.92
            first_score=1.0 if normalize_ocr_text(rr["first"]) in norm and len(rr["first"])>=3 else name_similarity(line_text,rr["first"])*0.85
            score=max(full_score,(last_score+first_score)/2)
            if score>best_score:
                best_score=score
                best=rr

        if not best or best_score < 0.55:
            continue

        # Extract numeric-looking tokens from the same row.
        nums=[]
        for token in line["words"]:
            clean=re.sub(r'[^0-9./-]','',token)
            # common forms: 18, 18/20, 90, 90.5
            if re.fullmatch(r'\d+(?:\.\d+)?',clean):
                nums.append(float(clean))
            elif re.fullmatch(r'\d+(?:\.\d+)?/\d+(?:\.\d+)?',clean):
                earned,total=clean.split("/")
                try:
                    nums.append(float(earned))
                except:
                    pass

        # Remove values likely coming from scholar IDs when the screenshot includes them.
        sid_text=str(best["Scholar ID"])
        nums_filtered=[]
        for num in nums:
            if str(int(num))==sid_text and num.is_integer():
                continue
            nums_filtered.append(num)

        if not nums_filtered:
            continue

        # Use the rightmost N numeric values, which most gradebooks place after the name.
        vals=nums_filtered[-n_assign:] if n_assign else nums_filtered
        row={"Scholar ID":best["Scholar ID"],"Scholar":best["Scholar"],"OCR Match %":round(best_score*100)}
        for i,a in enumerate(expected_assignments):
            col=f"{a['title']} ({a['points_possible']:g} pts)"
            row[col]=vals[i] if i < len(vals) else None
        results.append(row)

    if not results:
        return pd.DataFrame(), "OCR ran, but no scholar/grade rows could be matched. Try a clearer crop showing names and grades."

    result=pd.DataFrame(results)
    # Keep the strongest match if scholar appears twice.
    result=result.sort_values("OCR Match %",ascending=False).drop_duplicates("Scholar ID").sort_values("Scholar")
    return result, f"Matched {len(result)} scholar rows. Review every score before importing."

def import_ocr_preview(preview_df, assignments_by_col):
    saved=0
    skipped=0
    c=conn()
    for _,row in preview_df.iterrows():
        try:
            sid=int(row["Scholar ID"])
        except:
            skipped+=1
            continue
        for col,aid in assignments_by_col.items():
            value=row.get(col)
            if pd.isna(value) or value=="":
                continue
            try:
                score=float(value)
            except:
                skipped+=1
                continue
            ar=c.execute("SELECT points_possible FROM assignments WHERE id=?",(int(aid),)).fetchone()
            if not ar or score<0 or score>float(ar["points_possible"]):
                skipped+=1
                continue
            c.execute("INSERT OR REPLACE INTO grades(scholar_id,assignment_id,points_earned) VALUES (?,?,?)",
                      (sid,int(aid),score))
            saved+=1
    c.commit()
    c.close()
    return saved,skipped


def openlibrary_search_books(query, limit=8):
    """Search Open Library by title, author, ISBN, or keywords."""
    query=str(query or "").strip()
    if not query:
        return []
    try:
        params={
            "q":query,
            "limit":max(1,min(int(limit),20)),
            "fields":"key,title,author_name,first_publish_year,isbn,cover_i,edition_count,publisher,language"
        }
        r=requests.get(
            "https://openlibrary.org/search.json",
            params=params,
            timeout=10,
            headers={"User-Agent":"ChapLabTeacherHub/3.2"}
        )
        r.raise_for_status()
        docs=r.json().get("docs",[])
        results=[]
        for d in docs:
            isbns=d.get("isbn") or []
            authors=d.get("author_name") or []
            publishers=d.get("publisher") or []
            results.append({
                "key":d.get("key") or "",
                "title":d.get("title") or "Untitled",
                "author":", ".join(authors[:3]),
                "year":d.get("first_publish_year") or "",
                "isbn":isbns[0] if isbns else "",
                "all_isbns":isbns[:10],
                "cover_i":d.get("cover_i"),
                "publisher":publishers[0] if publishers else "",
                "edition_count":d.get("edition_count") or 0,
            })
        return results
    except Exception as e:
        raise RuntimeError(f"Open Library search failed: {e}")

def openlibrary_lookup_isbn(isbn):
    """Look up a specific ISBN via Open Library Books API."""
    isbn=re.sub(r"[^0-9Xx]","",str(isbn or ""))
    if not isbn:
        return None
    try:
        r=requests.get(
            "https://openlibrary.org/api/books",
            params={
                "bibkeys":f"ISBN:{isbn}",
                "format":"json",
                "jscmd":"data"
            },
            timeout=10,
            headers={"User-Agent":"ChapLabTeacherHub/3.2"}
        )
        r.raise_for_status()
        data=r.json().get(f"ISBN:{isbn}")
        if not data:
            return None
        authors=[a.get("name","") for a in data.get("authors",[]) if a.get("name")]
        publishers=[p.get("name","") for p in data.get("publishers",[]) if p.get("name")]
        cover=data.get("cover") or {}
        return {
            "title":data.get("title") or "Untitled",
            "author":", ".join(authors[:3]),
            "isbn":isbn,
            "publisher":publishers[0] if publishers else "",
            "publish_date":data.get("publish_date") or "",
            "cover_url":cover.get("medium") or cover.get("large") or cover.get("small") or "",
            "url":data.get("url") or "",
        }
    except Exception as e:
        raise RuntimeError(f"Open Library ISBN lookup failed: {e}")

def openlibrary_cover_url(book):
    cover_i=book.get("cover_i") if isinstance(book,dict) else None
    if cover_i:
        return f"https://covers.openlibrary.org/b/id/{cover_i}-M.jpg"
    return ""

def fp_level_relation(student_level, book_level):
    """Compare two F&P letter levels A-Z without pretending non-letter levels."""
    if not student_level or not book_level:
        return None
    s=str(student_level).strip().upper()
    b=str(book_level).strip().upper()
    if not re.fullmatch(r"[A-Z]",s) or not re.fullmatch(r"[A-Z]",b):
        return None
    diff=ord(b)-ord(s)
    if diff <= -2:
        return ("Below current level",diff)
    if diff == -1:
        return ("Slightly below current level",diff)
    if diff == 0:
        return ("On current level",diff)
    if diff == 1:
        return ("Slightly above current level",diff)
    return ("Above current level",diff)

# ---------- App Shell ----------
cdf=classes_df()
folder={0:"All Classes", **{int(r.id):r.class_name for _,r in cdf.iterrows()}}

if "selected_class" not in st.session_state:
    st.session_state["selected_class"]=0

preferred_names=["3-207","3-208","3-212"]
if not st.session_state["selected_class"] and not cdf.empty:
    preferred_ids=[int(cdf[cdf.class_name==n].iloc[0].id) for n in preferred_names if not cdf[cdf.class_name==n].empty]
    st.session_state["selected_class"]=preferred_ids[0] if preferred_ids else int(cdf.iloc[0].id)

selected_class=st.session_state["selected_class"]

internal_pages=["Home Page","Class Dashboard","Scholars","Scholar Profile","Scholar Binder","Book Leveler","Student Grouping","Report Card Comments","Little Assistant","Communication Log","Web & Backup"]
if "nav_page" not in st.session_state or st.session_state["nav_page"] not in internal_pages:
    st.session_state["nav_page"]="Home Page"

home_action=st.query_params.get("home")
home_class=st.query_params.get("class")
if home_class:
    try:
        st.session_state["selected_class"]=int(home_class)
        selected_class=int(home_class)
    except:
        pass
    st.query_params.clear()
if home_action:
    home_map={
        "dashboard":("Home Page",None,None,False,False),
        "classes":("Home Page",None,None,True,False),
        "scholars":("Scholars",None,None,False,False),
        "assignments":("Scholar Binder","Add Assignment",None,False,False),
        "grades":("Scholar Binder","Overview",None,False,False),
        "messages":("Little Assistant",None,"Parent Message",False,False),
        "planner":("Home Page",None,None,False,True),
    }
    if home_action in home_map:
        p,bt,at,sc,sq=home_map[home_action]
        st.session_state["nav_page"]=p
        if bt is not None:
            st.session_state["class_binder_tool"]=bt
        if at is not None:
            st.session_state["assistant_tool"]=at
        if sc:
            st.session_state["show_home_class_settings"]=True
        if sq:
            st.session_state["show_quarter_settings"]=True
    st.query_params.clear()


if side_action:
    side_map={
        "main":"Home Page",
        "scholars":"Scholars",
        "grades":"Scholar Binder",
        "books":"Book Leveler",
        "grouping":"Student Grouping",
        "reports":"Report Card Comments",
        "assistant":"Little Assistant",
        "communication":"Communication Log",
        "web":"Web & Backup",
    }
    st.session_state["nav_page"]=side_map.get(side_action,"Home Page")
    if side_action=="grades":
        st.session_state["class_binder_tool"]="Overview"
    st.query_params.clear()

quick=st.query_params.get("binder")
if quick:
    quick_map={
        "planning":("Home Page",None,None),
        "students":("Scholars",None,None),
        "data":("Student Grouping",None,None),
        "lessons":("Scholar Binder","Add Assignment",None),
        "resources":("Scholar Binder","Work Samples",None),
        "meetings":("Communication Log",None,None),
        "misc":("Little Assistant",None,"IEP / Student Support"),
    }
    if quick in quick_map:
        p,bt,at=quick_map[quick]
        st.session_state["nav_page"]=p
        if bt: st.session_state["class_binder_tool"]=bt
        if at: st.session_state["assistant_tool"]=at
    st.query_params.clear()

if st.session_state.get("chaplab_authenticated"):
    st.session_state["chaplab_last_activity"]=time.time()

page=st.session_state["nav_page"]

teacher=get_teacher_name() or "Ms. Chapman"

# Persistent teacher-hub navigation
active_classes=classes_df()
preferred_names=["3-207","3-208","3-212"]
records=[]
if not active_classes.empty:
    records=active_classes.to_dict("records")
    order={n:i for i,n in enumerate(preferred_names)}
    records=sorted(records,key=lambda r:(order.get(r["class_name"],99),r["class_name"]))

page=st.session_state["nav_page"]
selected_class=st.session_state.get("selected_class",0)

teacher=get_teacher_name() or "Ms. Chapman"
initials="".join([p[0] for p in teacher.replace(".","").split() if p])[:2].upper() or "MC"

nav_items=[
    ("🏠","Main Dashboard","Home Page"),
    ("🎓","Scholars","Scholars"),
    ("Ⓐ","Grades","Scholar Binder"),
    ("📚","Book Leveler","Book Leveler"),
    ("👥","Student Grouping","Student Grouping"),
    ("📝","Report Card Comments","Report Card Comments"),
    ("✨","Little Assistant","Little Assistant"),
    ("💬","Communication Log","Communication Log"),
    ("⚙️","Web & Backup","Web & Backup"),
]

def _native_nav_click(target):
    st.session_state["nav_page"]=target
    if target=="Scholar Binder":
        st.session_state["class_binder_tool"]="Overview"

with st.sidebar:
    st.markdown(f"### {teacher}’s Teacher Hub")
    st.caption("Navigation")
    for icon,label,target in nav_items:
        button_type="primary" if page==target else "secondary"
        if st.button(
            f"{icon} {label}",
            key=f"nav_{target}",
            use_container_width=True,
            type=button_type
        ):
            _native_nav_click(target)
            st.rerun()

    st.markdown("---")
    st.caption(f"☁️ {cloud_status_text()}")

st.markdown("**Classes**")
if records:
    class_cols=st.columns(len(records))
    for col,item in zip(class_cols,records):
        cid=int(item["id"])
        cname=item["class_name"]
        with col:
            if st.button(
                cname,
                key=f"class_nav_{cid}",
                use_container_width=True,
                type="primary" if cid==selected_class else "secondary"
            ):
                st.session_state["selected_class"]=cid
                st.session_state["nav_page"]="Class Dashboard"
                st.session_state["class_binder_tool"]="Overview"
                st.rerun()

# ---------- Pages ----------
if page=="Home Page":
    all_classes=classes_df()
    total_all=len(scholars_df())

    class_summaries=[]
    for _,cr in all_classes.iterrows():
        cid=int(cr.id)
        counts,_=scholar_status_summary(cid)
        class_summaries.append({
            "name":cr.class_name,
            "count":len(scholars_df(cid)),
            "on":counts["On Track"],
            "approach":counts["Approaching"],
            "risk":counts["At Risk"],
            "nodata":counts["No Data"]
        })

    recent=recent_assignments_df(1)
    latest="No assignments yet"
    if not recent.empty:
        rr=recent.iloc[0]
        latest=f"{rr.title} • {rr.subject}"

    anns=home_announcements()
    next_due="No saved reminder"
    if anns:
        k,l,d=anns[0]
        next_due=f"{k}: {l}" + (f" • {d}" if d else "")

    hero=f"""
    <div class="hero-card">
      <h1>{html.escape(get_teacher_name() or "Ms. Chapman")}’s Teacher Hub</h1>
      <div class="accent">Main Dashboard</div>
    </div>
    """
    if hasattr(st,"html"): st.html(hero)
    else: st.markdown(hero,unsafe_allow_html=True)

    teacher_info=teacher_dashboard_info()
    with st.expander("👩🏽‍🏫 My Teacher Info",expanded=False):
        ti1,ti2=st.columns([1,2])
        dash_homeroom=ti1.text_input(
            "Homeroom",
            value=teacher_info.get("homeroom",""),
            placeholder="Example: 3-208",
            key="dashboard_teacher_homeroom"
        )
        dash_subjects=ti2.multiselect(
            "My Subjects",
            ["ELA","Math","Science","Social Studies","Grammar","Writing"],
            default=[s for s in teacher_info.get("subjects",[]) if s in ["ELA","Math","Science","Social Studies","Grammar","Writing"]],
            key="dashboard_teacher_subjects"
        )
        if st.button("Save My Teacher Info",key="save_dashboard_teacher_info"):
            save_teacher_dashboard_info(dash_homeroom,dash_subjects)
            st.success("Teacher information saved.")

    posts=[]
    colors=["s-pink","s-orange","s-yellow"]
    for i,info in enumerate(class_summaries[:3]):
        posts.append(
            f'<div class="sticky-card {colors[i]}">'
            f'<strong>{html.escape(info["name"])}</strong>'
            f'{info["count"]} Scholars'
            f'<small>{info["on"]} on track • {info["approach"]+info["risk"]} need attention</small>'
            f'</div>'
        )
    posts.extend([
        f'<div class="sticky-card s-green"><strong>{total_all}</strong>Total Scholars<small>Across all active classes</small></div>',
        f'<div class="sticky-card s-blue"><strong>Latest</strong>{html.escape(latest)}<small>Most recent assignment entered</small></div>',
        f'<div class="sticky-card s-purple"><strong>Reminder</strong>{html.escape(next_due)}<small>Next saved deadline/update</small></div>',
    ])
    st.markdown('<div class="sticky-grid">'+"".join(posts)+'</div>',unsafe_allow_html=True)

    st.markdown('<div class="page-title">All Classes Overview</div>',unsafe_allow_html=True)
    overview_rows=[]
    for info in class_summaries:
        overview_rows.append({
            "Class":info["name"],
            "Scholars":info["count"],
            "On Track":info["on"],
            "Approaching":info["approach"],
            "At Risk":info["risk"],
            "No Data":info["nodata"]
        })
    if overview_rows:
        st.dataframe(pd.DataFrame(overview_rows),hide_index=True,use_container_width=True)

    left,right=st.columns(2)
    with left:
        st.markdown("### Recent Assignments")
        recent5=recent_assignments_df(5)
        if recent5.empty:
            st.caption("No assignments yet.")
        else:
            cols=[c for c in ["assignment_date","title","subject","class_name"] if c in recent5.columns]
            st.dataframe(recent5[cols],hide_index=True,use_container_width=True)
    with right:
        st.markdown("### Reminders & Deadlines")
        anns=home_announcements()
        if not anns:
            st.caption("No saved reminders or report-card deadlines.")
        else:
            st.dataframe(pd.DataFrame(anns,columns=["Type","Reminder","Due"]),hide_index=True,use_container_width=True)

elif page=="Class Dashboard":
    if not selected_class:
        st.warning("Choose a class from the top class bar.")
        st.stop()

    selected_name=class_name_from_id(int(selected_class)) or "Selected Class"
    counts,detail=scholar_status_summary(int(selected_class))
    roster=scholars_df(int(selected_class))
    total=len(roster)

    c=conn()
    assignment_count=c.execute("SELECT COUNT(*) n FROM assignments WHERE class_id=?",(int(selected_class),)).fetchone()["n"]
    contact_count=c.execute("""SELECT COUNT(*) n FROM communications
                               WHERE scholar_id IN (SELECT id FROM scholars WHERE class_id=?)""",
                            (int(selected_class),)).fetchone()["n"]
    c.close()

    recent=class_assignments_df(int(selected_class),"All Subjects","Newest → Oldest")
    latest="No assignments yet"
    if not recent.empty:
        rr=recent.iloc[0]
        latest=f"{rr.title} • {rr.subject}"

    hero=f"""
    <div class="hero-card">
      <h1>{html.escape(selected_name)} Dashboard</h1>
      <div class="accent">Class View</div>
    </div>
    """
    if hasattr(st,"html"): st.html(hero)
    else: st.markdown(hero,unsafe_allow_html=True)

    posts=[
        f'<div class="sticky-card s-pink"><strong>{total}</strong>Scholars<small>Active roster</small></div>',
        f'<div class="sticky-card s-orange"><strong>{counts["On Track"]}</strong>On Track<small>{counts["Approaching"]} approaching • {counts["At Risk"]} at risk</small></div>',
        f'<div class="sticky-card s-yellow"><strong>{assignment_count}</strong>Assignments<small>{html.escape(latest)}</small></div>',
        f'<div class="sticky-card s-green"><strong>{contact_count}</strong>Family Contacts<small>Logged communication</small></div>',
        f'<div class="sticky-card s-blue"><strong>{counts["No Data"]}</strong>No Data<small>Needs more grade evidence</small></div>',
        f'<div class="sticky-card s-purple"><strong>{counts["Approaching"]+counts["At Risk"]}</strong>Need Attention<small>Approaching + at risk</small></div>',
    ]
    st.markdown('<div class="sticky-grid">'+"".join(posts)+'</div>',unsafe_allow_html=True)

    left,right=st.columns(2)
    with left:
        st.markdown("### Scholars to Watch")
        if detail.empty:
            st.caption("No scholar data yet.")
        else:
            watch=detail[detail.Status.isin(["Approaching","At Risk"])].copy()
            if watch.empty:
                st.caption("No scholars are currently flagged.")
            else:
                st.dataframe(watch,hide_index=True,use_container_width=True)

    with right:
        st.markdown("### Recent Assignments")
        if recent.empty:
            st.caption("No assignments yet.")
        else:
            cols=[c for c in ["assignment_date","title","subject","standard_code"] if c in recent.columns]
            st.dataframe(recent.head(6)[cols],hide_index=True,use_container_width=True)

elif page=="Scholars":
    st.markdown('<div class="page-title">Scholars</div><div class="page-subtitle">Manage rosters and open scholar profiles.</div>',unsafe_allow_html=True)

    st.markdown("### Add Scholars")
    add_mode=st.radio(
        "How would you like to add scholars?",
        ["Manual Entry","Roster Spreadsheet"],
        horizontal=True,
        key="scholar_add_mode"
    )

    if add_mode=="Manual Entry":
        class_choices=classes_df()
        if class_choices.empty:
            st.info("Create a class from Home Page → Settings before adding scholars manually.")
        else:
            class_ids=list(class_choices.id.astype(int))
            default_class=selected_class if selected_class in class_ids else class_ids[0]
            with st.form("manual_add_scholar",clear_on_submit=True):
                r1c1,r1c2,r1c3=st.columns(3)
                first=r1c1.text_input("First Name")
                last=r1c2.text_input("Last Name")
                class_id=r1c3.selectbox(
                    "Class",
                    class_ids,
                    format_func=lambda x:class_choices[class_choices.id==x].iloc[0].class_name,
                    index=class_ids.index(default_class)
                )

                r2c1,r2c2,r2c3=st.columns(3)
                student_id=r2c1.text_input("School ID")
                grade_level=r2c2.text_input("Grade Level",value="3")
                academic_year=r2c3.text_input("Academic Year",value=current_academic_year())

                school_name=st.text_input("School Name")
                address=st.text_input("Address")
                a1,a2,a3=st.columns(3)
                city=a1.text_input("City")
                state=a2.text_input("State")
                zip_code=a3.text_input("ZIP Code")
                residency=st.text_input("Residency")

                if st.form_submit_button("Add Scholar"):
                    if not first.strip() or not last.strip():
                        st.error("First and last name are required.")
                    else:
                        cname=class_name_from_id(class_id)
                        c=conn()
                        duplicate=None
                        if student_id.strip():
                            duplicate=c.execute(
                                "SELECT id FROM scholars WHERE student_id=? AND student_id<>''",
                                (student_id.strip(),)
                            ).fetchone()
                        if not duplicate:
                            duplicate=c.execute(
                                """SELECT id FROM scholars
                                   WHERE lower(first_name)=lower(?)
                                     AND lower(last_name)=lower(?)
                                     AND class_id=?""",
                                (first.strip(),last.strip(),int(class_id))
                            ).fetchone()

                        if duplicate:
                            c.close()
                            st.warning("A matching scholar already exists in the roster.")
                        else:
                            c.execute("""INSERT INTO scholars(
                                first_name,last_name,class_name,class_id,school_name,academic_year,
                                grade_level,student_id,address,city,state_code,zip_code,residency,active)
                                VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,1)""",
                                (first.strip(),last.strip(),cname,int(class_id),school_name.strip(),
                                 academic_year.strip(),grade_level.strip(),student_id.strip(),
                                 address.strip(),city.strip(),state.strip(),zip_code.strip(),residency.strip()))
                            c.commit(); c.close()
                            st.success(f"{first.strip()} {last.strip()} added to {cname}.")
                            st.rerun()

    else:
        if st.button("📥 Import Roster Spreadsheet", key="show_roster_import"):
            st.session_state["show_roster_importer"]=not st.session_state.get("show_roster_importer",False)
        if st.session_state.get("show_roster_importer",False):
            st.markdown("### Roster Spreadsheet Importer")
            st.write("Upload an Excel or CSV roster. The importer can create scholar profiles and attach parent/guardian contact information.")
            up=st.file_uploader("Roster file",type=["xlsx","xls","csv"])
            if up:
                try:
                    if up.name.lower().endswith(".csv"):
                        df=pd.read_csv(up,dtype=str)
                    else:
                        df=pd.read_excel(up,dtype=str)
                    df=df.fillna("")
                    st.success(f"Loaded {len(df)} rows and {len(df.columns)} columns.")
                    st.dataframe(df.head(10),use_container_width=True,hide_index=True)

                    detected=auto_mapping(df.columns)
                    st.markdown("### Column matching")
                    st.caption("I matched the likely columns automatically. Change any dropdown that is incorrect.")
                    none="— Not in this file —"
                    opts=[none]+list(df.columns)
                    labels=[
                        ("school_name","School Name"),("academic_year","Academic Year"),("grade_level","Grade Level"),
                        ("class_name","Class / Course / Section"),("student_id","Student ID"),
                        ("student_first","Scholar First Name"),("student_last","Scholar Last Name"),
                        ("address","Address"),("city","City"),("state_code","State"),("zip_code","Zip Code"),
                        ("residency","Residency"),("guardian_first","Guardian First Name"),("guardian_last","Guardian Last Name"),
                        ("relationship","Relationship"),("home_phone","Home Phone"),("work_phone","Work Phone"),
                        ("cell_phone","Cell Phone"),("email","Guardian Email")
                    ]
                    chosen={}
                    left,right=st.columns(2)
                    for i,(key,label) in enumerate(labels):
                        container=left if i%2==0 else right
                        default=detected.get(key,none)
                        idx=opts.index(default) if default in opts else 0
                        chosen[key]=container.selectbox(label,opts,index=idx,key=f"map_{key}")

                    fallback_class=st.text_input("If class/section is missing, place scholars in this class folder",value="")
                    update_existing=st.checkbox("Update matching existing scholar profiles",value=True)
                    st.info("Duplicate guardian rows are combined under the same scholar. A scholar is matched primarily by Student ID when available; otherwise by first name + last name + class.")

                    if st.button("Import Scholar Profiles"):
                        if chosen["student_first"]==none or chosen["student_last"]==none:
                            st.error("Choose the Scholar First Name and Scholar Last Name columns.")
                        else:
                            c=conn(); cur=c.cursor()
                            scholar_count=0; guardian_count=0; updated=0
                            for _,r in df.iterrows():
                                def val(key):
                                    col=chosen[key]
                                    return "" if col==none else clean(r[col])
                                sf,sl=val("student_first"),val("student_last")
                                if not sf or not sl: continue
                                sid_text=val("student_id")
                                class_name=val("class_name") or fallback_class.strip()
                                class_id=None
                                if class_name:
                                    cur.execute("INSERT OR IGNORE INTO classes(class_name) VALUES (?)",(class_name,))
                                    class_id=cur.execute("SELECT id FROM classes WHERE class_name=?",(class_name,)).fetchone()["id"]

                                existing=None
                                if sid_text:
                                    existing=cur.execute("SELECT * FROM scholars WHERE student_id=? AND student_id<>''",(sid_text,)).fetchone()
                                if not existing:
                                    if class_id:
                                        existing=cur.execute("""SELECT * FROM scholars WHERE lower(first_name)=lower(?) AND lower(last_name)=lower(?) AND class_id=?""",
                                                             (sf,sl,class_id)).fetchone()
                                    else:
                                        existing=cur.execute("""SELECT * FROM scholars WHERE lower(first_name)=lower(?) AND lower(last_name)=lower(?)""",(sf,sl)).fetchone()

                                fields=(sf,sl,class_name,class_id,val("school_name"),val("academic_year"),val("grade_level"),
                                        sid_text,val("address"),val("city"),val("state_code"),val("zip_code"),val("residency"))
                                if existing:
                                    scholar_id=existing["id"]
                                    if update_existing:
                                        cur.execute("""UPDATE scholars SET first_name=?,last_name=?,class_name=?,class_id=?,
                                          school_name=?,academic_year=?,grade_level=?,student_id=?,address=?,city=?,state_code=?,zip_code=?,residency=?,active=1
                                          WHERE id=?""",fields+(scholar_id,))
                                        updated+=1
                                else:
                                    cur.execute("""INSERT INTO scholars(first_name,last_name,class_name,class_id,school_name,academic_year,grade_level,
                                      student_id,address,city,state_code,zip_code,residency) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)""",fields)
                                    scholar_id=cur.lastrowid; scholar_count+=1

                                gf,gl,rel=val("guardian_first"),val("guardian_last"),val("relationship")
                                hp,wp,cp,em=val("home_phone"),val("work_phone"),val("cell_phone"),val("email")
                                if gf or gl or hp or wp or cp or em:
                                    cur.execute("""INSERT OR IGNORE INTO guardians
                                      (scholar_id,first_name,last_name,relationship,home_phone,work_phone,cell_phone,email)
                                      VALUES (?,?,?,?,?,?,?,?)""",(scholar_id,gf,gl,rel,hp,wp,cp,em))
                                    # update contact fields if same guardian already exists
                                    cur.execute("""UPDATE guardians SET home_phone=?,work_phone=?,cell_phone=?,email=?
                                      WHERE scholar_id=? AND first_name=? AND last_name=? AND relationship=?""",
                                      (hp,wp,cp,em,scholar_id,gf,gl,rel))
                                    guardian_count+=1
                            c.commit(); c.close()
                            st.success(f"Import complete: {scholar_count} new scholars, {updated} profiles updated, and guardian/contact rows processed.")
                            st.rerun()
                except Exception as e:
                    st.error(f"I could not read that spreadsheet: {e}")

    roster=scholars_df(selected_class or None)
    st.dataframe(roster[["first_name","last_name","grade_level","display_class","student_id"]],hide_index=True,use_container_width=True)

    if not roster.empty:
        st.markdown("### Scholars")
        st.caption("Click a scholar's name to open the profile.")
        cols_per_row=3
        for i in range(0,len(roster),cols_per_row):
            rowcols=st.columns(cols_per_row)
            chunk=roster.iloc[i:i+cols_per_row]
            for j,(_,sr) in enumerate(chunk.iterrows()):
                rowcols[j].button(
                    nm(sr),
                    key=f"open_profile_{int(sr.id)}",
                    use_container_width=True,
                    on_click=open_scholar_profile,
                    args=(int(sr.id),)
                )

    if not roster.empty:
        st.markdown("### Remove from roster")
        rid=st.selectbox("Scholar",list(roster.id.astype(int)),format_func=lambda x:nm(roster[roster.id==x].iloc[0]))
        a,b=st.columns(2)
        if a.button("Archive Scholar"):
            c=conn(); c.execute("UPDATE scholars SET active=0 WHERE id=?",(rid,)); c.commit(); c.close(); st.rerun()
        ok=b.checkbox("I understand permanent delete cannot be undone")
        if b.button("Delete Permanently",disabled=not ok):
            c=conn()
            for t in ["grades","communications","work_samples","support_notes","report_comments","guardians"]:
                c.execute(f"DELETE FROM {t} WHERE scholar_id=?",(rid,))
            c.execute("DELETE FROM scholars WHERE id=?",(rid,)); c.commit(); c.close(); st.rerun()

elif page=="Scholar Profile":
    sid=st.session_state.get("selected_profile_scholar")
    if not sid:
        st.warning("Choose a scholar from the Scholars page.")
        st.button("Back to Scholars",on_click=return_to_scholars)
    else:
        s=scholar_full_profile(sid)
        if not s:
            st.warning("That scholar profile could not be found.")
            st.button("Back to Scholars",on_click=return_to_scholars)
        else:
            st.button("← Back to Scholars",key="back_to_scholars_profile",on_click=return_to_scholars)
            initials=((s["first_name"][:1] if s["first_name"] else "")+(s["last_name"][:1] if s["last_name"] else "")).upper()

            st.markdown('<div class="profile-cover"></div>',unsafe_allow_html=True)
            profile_html=(
                '<div class="profile-header">'
                f'<div class="profile-avatar">{initials}</div>'
                f'<h2 style="margin:12px 0 2px 0">{html.escape(s["first_name"])} {html.escape(s["last_name"])}</h2>'
                f'<div>School ID: {html.escape(str(s["student_id"] or "—"))} | Grade {html.escape(str(s["grade_level"] or "—"))} | {html.escape(str(s["display_class"] or "No Class"))}</div>'
                f'<div style="margin-top:5px">{html.escape(str(s["school_name"] or ""))}</div>'
                f'<div style="margin-top:5px;color:#667085;">Gender: {html.escape(str(s["gender"] or "—"))} | Pronouns: {html.escape(str(s["pronouns"] or "—"))}</div>'
                '</div>'
            )
            st.markdown(profile_html,unsafe_allow_html=True)

            # Scholar Information
            st.markdown('<div class="profile-section"><h3>Scholar Information</h3></div>',unsafe_allow_html=True)
            class_choices=classes_df()
            class_ids=[0]+list(class_choices.id.astype(int))
            class_map={0:"Unassigned",**{int(r.id):r.class_name for _,r in class_choices.iterrows()}}
            current_class=int(s["class_id"]) if s["class_id"] and int(s["class_id"]) in class_ids else 0
            with st.form("profile_edit_continuous"):
                a,b=st.columns(2)
                ef=a.text_input("First Name",value=s["first_name"] or "")
                el=b.text_input("Last Name",value=s["last_name"] or "")
                c1,c2,c3=st.columns(3)
                eid=c1.text_input("School ID",value=s["student_id"] or "")
                egrade=c2.text_input("Grade Level",value=s["grade_level"] or "")
                eclass=c3.selectbox("Class",class_ids,format_func=lambda x:class_map[x],index=class_ids.index(current_class))

                g1,g2=st.columns(2)
                gender_options=["","Female","Male","Nonbinary","Other / Prefer to self-describe","Prefer not to say"]
                current_gender=str(s["gender"] or "")
                gender_index=gender_options.index(current_gender) if current_gender in gender_options else 0
                egender=g1.selectbox("Gender",gender_options,index=gender_index)
                pronoun_options=["","she/her","he/him","they/them","Custom"]
                current_pronouns=str(s["pronouns"] or "")
                pronoun_index=pronoun_options.index(current_pronouns) if current_pronouns in pronoun_options else (4 if current_pronouns else 0)
                epronoun_choice=g2.selectbox("Pronouns",pronoun_options,index=pronoun_index)
                custom_pronouns=st.text_input(
                    "Custom pronouns",
                    value=current_pronouns if epronoun_choice=="Custom" else "",
                    placeholder="Example: ze/hir/hir"
                )
                epronouns=custom_pronouns.strip() if epronoun_choice=="Custom" else epronoun_choice

                school=st.text_input("School Name",value=s["school_name"] or "")
                year=st.text_input("Academic Year",value=s["academic_year"] or "")
                if st.form_submit_button("Save Scholar Changes"):
                    c=conn()
                    cname=class_map[eclass] if eclass else ""
                    c.execute("""UPDATE scholars SET first_name=?,last_name=?,student_id=?,grade_level=?,
                                 class_id=?,class_name=?,school_name=?,academic_year=?,gender=?,pronouns=? WHERE id=?""",
                              (ef.strip(),el.strip(),eid.strip(),egrade.strip(),
                               eclass if eclass else None,cname,school.strip(),year.strip(),egender,epronouns,int(sid)))
                    c.commit(); c.close()
                    st.success("Scholar information updated.")
                    st.rerun()

            # Benchmarks
            br=benchmark_for_scholar(sid)
            st.markdown("### Benchmark Snapshot")
            b1,b2,b3=st.columns(3)
            b1.metric("F&P", (br["fp_spring_level"] or br["fp_fall_level"] or "—") if br else "—")
            b2.metric("NWEA Reading", (br["nwea_spring_reading"] or br["nwea_winter_reading"] or br["nwea_fall_reading"] or "—") if br else "—")
            b3.metric("NWEA Math", (br["nwea_spring_math"] or br["nwea_winter_math"] or br["nwea_fall_math"] or "—") if br else "—")

            # Grades
            st.markdown("### Current Grades")
            cols=st.columns(4)
            for i,subj in enumerate(["ELA","Math","Science","Social Studies"]):
                av,_=summary(sid,subj)
                cols[i].metric(subj,"No grades" if av is None else f"{av:.1f}% ({letter(av)})")

            # Academic summary
            strengths,needs,teacher_actions=academic_summary_for_scholar(sid)
            st.markdown("### Academic Summary")
            left,right=st.columns(2)
            with left:
                st.markdown("**Strengths**")
                if strengths:
                    for subj,skill,pct in strengths:
                        st.write(f"- {subj}: {skill} — {pct:.0f}%")
                else:
                    st.caption("Not enough skill data yet.")
            with right:
                st.markdown("**Needs Support**")
                if needs:
                    for subj,skill,pct in needs:
                        st.write(f"- {subj}: {skill} — {pct:.0f}%")
                else:
                    st.caption("No below-mastery skill pattern identified.")
            st.markdown("**Teacher Support Plan**")
            for action in teacher_actions:
                st.write(f"- {action}")

            # Parent contacts
            st.markdown("### Parent / Guardian Contacts")
            c=conn()
            g=pd.read_sql_query("SELECT * FROM guardians WHERE scholar_id=? ORDER BY relationship,last_name,first_name",c,params=[sid])
            c.close()
            if g.empty:
                st.caption("No parent/guardian contacts saved yet.")
            else:
                display_g=g.copy()
                display_g["Name"]=(display_g["first_name"].fillna("")+" "+display_g["last_name"].fillna("")).str.strip()
                display_g["Phone"]=display_g["cell_phone"].where(display_g["cell_phone"].fillna("")!="",display_g["home_phone"])
                st.dataframe(display_g[["Name","relationship","Phone","email"]],hide_index=True,use_container_width=True)

            with st.expander("Add Parent / Guardian"):
                with st.form("profile_add_guardian_cont",clear_on_submit=True):
                    a,b=st.columns(2)
                    gf=a.text_input("First Name")
                    gl=b.text_input("Last Name")
                    c1,c2=st.columns(2)
                    rel=c1.text_input("Relationship")
                    phone=c2.text_input("Phone Number")
                    email=st.text_input("Email")
                    if st.form_submit_button("Add Contact"):
                        if not gf.strip() and not gl.strip():
                            st.error("Enter at least a first or last name.")
                        else:
                            c=conn()
                            c.execute("""INSERT OR IGNORE INTO guardians
                                (scholar_id,first_name,last_name,relationship,cell_phone,email)
                                VALUES (?,?,?,?,?,?)""",
                                (int(sid),gf.strip(),gl.strip(),rel.strip(),phone.strip(),email.strip()))
                            c.commit(); c.close(); st.rerun()

            if not g.empty:
                with st.expander("Edit / Delete Parent Contact"):
                    gid=st.selectbox("Choose Contact",list(g.id.astype(int)),
                                     format_func=lambda x:(f"{g[g.id==x].iloc[0].first_name} {g[g.id==x].iloc[0].last_name}".strip()),
                                     key="profile_guardian_select_cont")
                    gr=g[g.id==gid].iloc[0]
                    with st.form("profile_edit_guardian_cont"):
                        a,b=st.columns(2)
                        egf=a.text_input("First Name",value=gr.first_name or "")
                        egl=b.text_input("Last Name",value=gr.last_name or "")
                        c1,c2=st.columns(2)
                        erel=c1.text_input("Relationship",value=gr.relationship or "")
                        ephone=c2.text_input("Phone Number",value=(gr.cell_phone or gr.home_phone or ""))
                        eemail=st.text_input("Email",value=gr.email or "")
                        if st.form_submit_button("Save Contact Changes"):
                            c=conn()
                            c.execute("""UPDATE guardians SET first_name=?,last_name=?,relationship=?,
                                         cell_phone=?,home_phone='',work_phone='',email=? WHERE id=?""",
                                      (egf.strip(),egl.strip(),erel.strip(),ephone.strip(),eemail.strip(),int(gid)))
                            c.commit(); c.close(); st.rerun()
                    confirm=st.checkbox("Delete this contact",key="profile_del_guardian_confirm")
                    if st.button("Delete Parent / Guardian",disabled=not confirm,key="profile_del_guardian"):
                        c=conn()
                        c.execute("UPDATE communications SET guardian_id=NULL WHERE guardian_id=?",(int(gid),))
                        c.execute("UPDATE contact_reminders SET guardian_id=NULL WHERE guardian_id=?",(int(gid),))
                        c.execute("UPDATE parent_update_preferences SET preferred_guardian_id=NULL WHERE preferred_guardian_id=?",(int(gid),))
                        c.execute("DELETE FROM guardians WHERE id=?",(int(gid),))
                        c.commit(); c.close(); st.rerun()

            # History / reports / support / reminders
            st.markdown("### Parent Contact History")
            c=conn()
            hist=pd.read_sql_query("""SELECT communications.created_at,
                 COALESCE(TRIM(guardians.first_name||' '||guardians.last_name),'') guardian,
                 guardians.relationship,communications.communication_type,
                 communications.subject,communications.reason,communications.details
                 FROM communications LEFT JOIN guardians ON guardians.id=communications.guardian_id
                 WHERE communications.scholar_id=? ORDER BY communications.id DESC""",c,params=[sid])
            rc=pd.read_sql_query("SELECT created_at,subject,marking_period,comment_text FROM report_comments WHERE scholar_id=? ORDER BY id DESC",c,params=[sid])
            sn=pd.read_sql_query("SELECT created_at,note_type,area,observation,intervention,response_to_intervention,impact FROM support_notes WHERE scholar_id=? ORDER BY id DESC",c,params=[sid])
            c.close()

            if hist.empty: st.caption("No parent contacts logged yet.")
            else: st.dataframe(hist,hide_index=True,use_container_width=True)

            st.markdown("### Saved Report Card Comments")
            if rc.empty: st.caption("No saved report card comments yet.")
            else: st.dataframe(rc,hide_index=True,use_container_width=True)

            st.markdown("### Support / IEP Evidence")
            if sn.empty: st.caption("No support notes saved yet.")
            else: st.dataframe(sn,hide_index=True,use_container_width=True)

            st.markdown("### Parent Update Reminders")
            rem=reminder_rows(sid)
            if rem.empty: st.caption("No reminders.")
            else: st.dataframe(rem[["due_date","guardian","reason","notes","completed"]],hide_index=True,use_container_width=True)

elif page=="Scholar Binder":
    st.markdown('<div class="page-title">Grades</div><div class="page-subtitle">Assignments, gradebook, standards, NWEA/F&P, and grade settings.</div>',unsafe_allow_html=True)
    if not selected_class:
        st.info("Choose a class from the Binder Cover first.")
    else:
        class_name=folder.get(selected_class,"Selected Class")
        st.markdown(f"### 📌 {class_name}")
        binder_tool=st.radio("Binder Page",["Overview","Add Assignment","Skills & Standards","Work Samples","NWEA & F&P","Grade Settings"],horizontal=True,key="class_binder_tool")

        if binder_tool=="Overview":
            st.markdown("## Class Gradebook Overview")
            f1,f2=st.columns(2)
            subject_filter=f1.selectbox("Subject",["All Subjects","ELA","Math","Science","Social Studies"],key="grid_subject_filter")
            sort_order=f2.selectbox("Sort assignments by",["Oldest → Newest","Newest → Oldest","Subject","Standard Code"],key="grid_sort")
            matrix,assignments=gradebook_matrix(selected_class,subject_filter,sort_order)
            if matrix.empty: st.info("No scholars are available for this class.")
            else: st.dataframe(matrix,hide_index=True,use_container_width=True,height=520)

            st.markdown("### Printable / Uploadable Excel Grade Sheet")
            st.write("Download the sheet, print it or type grades into it, then upload the completed Excel file to enter grades in bulk.")
            excel_bytes=make_grade_sheet_xlsx(selected_class,class_name,subject_filter,sort_order)
            st.download_button("⬇️ Download Excel Grade Sheet",data=excel_bytes,file_name=f"{class_name.replace(' ','_')}_Grade_Entry.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",key="download_grade_sheet")
            uploaded_grade_sheet=st.file_uploader("Upload completed Excel grade sheet",type=["xlsx"],key="upload_completed_grade_sheet")
            if uploaded_grade_sheet is not None and st.button("Import Grades from Excel",key="import_grades_xlsx"):
                try:
                    saved,skipped=import_grade_sheet_xlsx(uploaded_grade_sheet)
                    st.success(f"Imported {saved} grade entries. Skipped {skipped} invalid entries.")
                    st.rerun()
                except Exception as e:
                    st.error(f"Could not import the grade sheet: {e}")

            st.markdown("---")
            st.markdown("### 📸 Import Grades from NHA Screenshot")
            st.caption(
                "Upload a screenshot that shows scholar names and their entered grades. "
                "ChapLab reads the screenshot locally, then gives you an editable preview before saving anything."
            )

            screenshot_subject=st.selectbox(
                "Subject for screenshot",
                ["ELA","Math","Science","Social Studies"],
                key="nha_ss_subject"
            )
            screenshot_assignments=class_assignments_df(selected_class,screenshot_subject,"Oldest → Newest")

            if screenshot_assignments.empty:
                st.info("Create the assignment(s) in ChapLab first so the screenshot grades have somewhere to go.")
            else:
                assign_map={
                    int(r.id):f"{r.title} | {r.standard_code or 'No Standard'} | {r.assignment_date} | {r.points_possible:g} pts"
                    for _,r in screenshot_assignments.iterrows()
                }
                selected_aids=st.multiselect(
                    "Assignments visible in the screenshot — choose them in LEFT-TO-RIGHT order",
                    list(assign_map.keys()),
                    format_func=lambda x:assign_map[x],
                    key="nha_ss_assignments"
                )

                uploaded_screens=st.file_uploader(
                    "Upload NHA gradebook screenshot(s)",
                    type=["png","jpg","jpeg"],
                    accept_multiple_files=True,
                    key="nha_grade_screens"
                )

                st.info(
                    "For best results: crop the screenshot so the scholar names and selected grade columns are visible, "
                    "keep the text at normal zoom or larger, and avoid covering rows with pop-ups."
                )

                if uploaded_screens and selected_aids:
                    selected_rows=[]
                    for aid in selected_aids:
                        ar=screenshot_assignments[screenshot_assignments.id==aid].iloc[0]
                        selected_rows.append({
                            "id":int(aid),
                            "title":ar.title,
                            "points_possible":float(ar.points_possible)
                        })

                    if st.button("Read Screenshot Grades",key="read_nha_screenshot"):
                        all_found=[]
                        messages=[]
                        roster=scholars_df(selected_class)
                        for screen in uploaded_screens:
                            try:
                                found,msg=ocr_grade_screenshot(screen,roster,selected_rows)
                                messages.append(f"{screen.name}: {msg}")
                                if not found.empty:
                                    all_found.append(found)
                            except Exception as e:
                                messages.append(f"{screen.name}: {e}")
                        if all_found:
                            combined=pd.concat(all_found,ignore_index=True)
                            combined=combined.sort_values("OCR Match %",ascending=False).drop_duplicates("Scholar ID").sort_values("Scholar")
                            st.session_state["nha_ocr_preview"]=combined
                            st.session_state["nha_ocr_assignment_ids"]=selected_aids
                        st.session_state["nha_ocr_messages"]=messages

                for msg in st.session_state.get("nha_ocr_messages",[]):
                    st.write(f"- {msg}")

                if "nha_ocr_preview" in st.session_state:
                    preview=st.session_state["nha_ocr_preview"].copy()
                    st.markdown("#### Review Before Import")
                    st.warning("OCR can make mistakes. Check scholar names and every grade before importing.")
                    edited_preview=st.data_editor(
                        preview,
                        disabled=["Scholar ID","Scholar","OCR Match %"],
                        hide_index=True,
                        use_container_width=True,
                        key="nha_ocr_editor"
                    )

                    current_ids=st.session_state.get("nha_ocr_assignment_ids",[])
                    col_to_aid={}
                    for aid in current_ids:
                        match=screenshot_assignments[screenshot_assignments.id==aid]
                        if not match.empty:
                            ar=match.iloc[0]
                            col_to_aid[f"{ar.title} ({ar.points_possible:g} pts)"]=int(aid)

                    if st.button("Import Reviewed Screenshot Grades",key="import_nha_ocr_grades"):
                        saved,skipped=import_ocr_preview(edited_preview,col_to_aid)
                        st.success(f"Imported {saved} reviewed grade entries. Skipped {skipped} invalid entries.")
                        st.session_state.pop("nha_ocr_preview",None)
                        st.session_state.pop("nha_ocr_messages",None)
                        st.session_state.pop("nha_ocr_assignment_ids",None)
                        st.rerun()

            with st.expander("Local screenshot reader setup"):
                st.write(
                    "The screenshot reader uses Tesseract OCR installed on your Windows computer. "
                    "The image is processed locally and is not sent to an outside AI service."
                )
                ok,ocr_status=configure_local_tesseract()
                if ok:
                    st.success(f"Local OCR ready: {ocr_status}")
                else:
                    st.warning(
                        "Tesseract OCR is not installed yet. Install Tesseract OCR for Windows, "
                        "then restart ChapLab. The Excel import continues to work without it."
                    )

        elif binder_tool=="Add Assignment":
            subject=st.selectbox("Subject",["ELA","Math","Science","Social Studies"],key="assignment_subject")
            st.markdown("### Assignment Information")
            with st.form("new_assignment_binder"):
                a,b=st.columns(2)
                title=a.text_input("Assignment title")
                category=b.selectbox("Grade category",list(get_setting("weights").keys()))
                sdf=standards_df(subject)
                skill_choice=st.selectbox("Skill / Standard",[""]+[f"{r.code} — {r.skill}" for _,r in sdf.iterrows()])
                c1,c2=st.columns(2)
                points=c1.number_input("Points possible",1.0,1000.0,100.0)
                adate=c2.date_input("Date",date.today())
                detected_quarter=quarter_for_date(str(adate),current_academic_year())
                marking_period=st.selectbox("Marking Period",
                                            ["Quarter 1","Quarter 2","Quarter 3","Quarter 4"],
                                            index=["Quarter 1","Quarter 2","Quarter 3","Quarter 4"].index(detected_quarter) if detected_quarter in ["Quarter 1","Quarter 2","Quarter 3","Quarter 4"] else 0,
                                            help=f"Detected from saved date ranges: {detected_quarter or 'No match'}")
                if st.form_submit_button("Create Assignment") and title.strip():
                    qsettings=get_quarter_settings(current_academic_year())
                    qrow=qsettings[qsettings.quarter_name==marking_period] if not qsettings.empty else pd.DataFrame()
                    if not qrow.empty and bool(qrow.iloc[0].locked):
                        st.error(f"{marking_period} is locked. Unlock it in Home Page settings before adding assignments.")
                        st.stop()
                    code=skill_choice.split(" — ")[0] if skill_choice else ""
                    c=conn()
                    c.execute("""INSERT INTO assignments(title,subject,category,standard_code,points_possible,assignment_date,class_id,marking_period)
                                 VALUES (?,?,?,?,?,?,?,?)""",(title.strip(),subject,category,code,float(points),str(adate),selected_class,marking_period))
                    c.commit(); c.close(); st.success("Assignment created."); st.rerun()

            assignments=class_assignments_df(selected_class,subject,"Newest → Oldest")
            if not assignments.empty:
                amap={int(r.id):f"{r.title} | {r.marking_period or 'No Quarter'} | {r.standard_code or 'No Standard'} | {r.assignment_date} | {r.points_possible:g} pts" for _,r in assignments.iterrows()}
                aid=st.selectbox("Assignment to enter grades",list(amap),format_func=lambda x:amap[x],key="assignment_grade_entry")
                ar=assignments[assignments.id==aid].iloc[0]

                with st.expander("🗑️ Delete Selected Assignment"):
                    st.warning("Deleting an assignment will also delete every scholar grade attached to that assignment.")
                    confirm_delete=st.checkbox(
                        f"I understand and want to delete: {ar.title}",
                        key="confirm_delete_assignment"
                    )
                    if st.button(
                        "Delete Assignment Permanently",
                        disabled=not confirm_delete,
                        key="delete_assignment_button"
                    ):
                        c=conn()
                        c.execute("DELETE FROM grades WHERE assignment_id=?",(int(aid),))
                        c.execute("DELETE FROM assignments WHERE id=?",(int(aid),))
                        c.commit()
                        c.close()
                        st.success("Assignment and linked grades deleted.")
                        st.rerun()

                entry_mode=st.radio("Grade Entry",["Whole Class","Search One Scholar"],horizontal=True,key="entry_mode")
                roster=scholars_df(selected_class)
                if entry_mode=="Whole Class":
                    c=conn(); current=pd.read_sql_query("SELECT scholar_id,points_earned FROM grades WHERE assignment_id=?",c,params=[aid]); c.close()
                    cmap=dict(zip(current.scholar_id,current.points_earned))
                    grid=pd.DataFrame({"Scholar ID":roster.id.astype(int),"Scholar":[nm(r) for _,r in roster.iterrows()],"Points Earned":[cmap.get(int(r.id),None) for _,r in roster.iterrows()]})
                    edited=st.data_editor(grid,disabled=["Scholar ID","Scholar"],hide_index=True,use_container_width=True)
                    if st.button("Save Class Grades",key="save_full_assignment_grades"):
                        c=conn()
                        for _,r in edited.iterrows():
                            if pd.notna(r["Points Earned"]):
                                score=float(r["Points Earned"])
                                if 0<=score<=float(ar.points_possible):
                                    c.execute("INSERT OR REPLACE INTO grades(scholar_id,assignment_id,points_earned) VALUES (?,?,?)",(int(r["Scholar ID"]),int(aid),score))
                        c.commit(); c.close(); st.success("Class grades saved."); st.rerun()
                else:
                    q=st.text_input("Search scholar",placeholder="Type first or last name")
                    filtered=roster
                    if q.strip(): filtered=filtered[(filtered.first_name+" "+filtered.last_name).str.contains(q.strip(),case=False,na=False)]
                    if not filtered.empty:
                        sid=st.selectbox("Scholar",list(filtered.id.astype(int)),format_func=lambda x:nm(filtered[filtered.id==x].iloc[0]))
                        c=conn(); old=c.execute("SELECT points_earned FROM grades WHERE scholar_id=? AND assignment_id=?",(sid,aid)).fetchone(); c.close()
                        default=float(old["points_earned"]) if old and old["points_earned"] is not None else 0.0
                        score=st.number_input("Points earned",0.0,float(ar.points_possible),default,key="single_scholar_grade")
                        if st.button("Save Scholar Grade",key="save_single_scholar_grade"):
                            c=conn(); c.execute("INSERT OR REPLACE INTO grades(scholar_id,assignment_id,points_earned) VALUES (?,?,?)",(sid,aid,float(score))); c.commit(); c.close(); st.success("Grade saved."); st.rerun()

        elif binder_tool=="Skills & Standards":
            subj=st.selectbox("Subject",["ELA","Math","Science","Social Studies"],key="binder_std_subject_v9")
            sdf=standards_df(subj)
            st.dataframe(sdf[["id","code","skill","description"]],hide_index=True,use_container_width=True)

        elif binder_tool=="Work Samples":
            roster=scholars_df(selected_class)
            if not roster.empty:
                sid=st.selectbox("Scholar",list(roster.id.astype(int)),format_func=lambda x:nm(roster[roster.id==x].iloc[0]),key="ws_student_v9")
                subj=st.selectbox("Subject",["ELA","Math","Science","Social Studies","Other"],key="ws_subject_v9")
                title=st.text_input("Work Sample Title",key="ws_title_v9")
                up=st.file_uploader("Upload PDF, image, DOCX or TXT",type=["pdf","png","jpg","jpeg","docx","txt"],key="ws_upload_v9")
                obs=st.text_area("Teacher Observation",key="ws_obs_v9")
                x,y=st.columns(2); strengths=x.text_area("Strengths",key="ws_strength_v9"); needs=y.text_area("Needs / Patterns",key="ws_needs_v9")
                nxt=st.text_area("Next Steps",key="ws_next_v9")
                if st.button("Save Work Sample",key="ws_save_v9"):
                    fn=fp=""
                    if up:
                        safe_name=datetime.now().strftime("%Y%m%d_%H%M%S_")+re.sub(r'[^A-Za-z0-9._-]','_',up.name)
                        fn=up.name
                        if cloud_configured():
                            remote_path=f"work_samples/{int(sid)}/{safe_name}"
                            content_type=getattr(up,"type",None) or "application/octet-stream"
                            if cloud_upload_bytes(bytes(up.getbuffer()),remote_path,content_type):
                                fp=f"cloud://{remote_path}"
                            else:
                                st.error("The work-sample file could not be saved to cloud storage.")
                                st.stop()
                        else:
                            folderp=UPLOAD_DIR/f"{sid}"; folderp.mkdir(parents=True,exist_ok=True)
                            dest=folderp/safe_name
                            dest.write_bytes(up.getbuffer()); fp=str(dest)
                    c=conn(); c.execute("""INSERT INTO work_samples(scholar_id,uploaded_at,subject,title,file_name,file_path,teacher_observation,strengths,needs,next_steps)
                                         VALUES (?,?,?,?,?,?,?,?,?,?)""",(sid,datetime.now().isoformat(timespec="minutes"),subj,title,fn,fp,obs,strengths,needs,nxt)); c.commit(); c.close(); st.success("Saved."); st.rerun()

        elif binder_tool=="NWEA & F&P":
            st.markdown("## 📊 NWEA & F&P Assessment Data")
            st.caption("Enter each scholar's benchmark scores here. These scores feed the scholar profile, growth summaries, parent updates, and support/IEP helper.")

            roster=scholars_df(selected_class if selected_class else None)
            if roster.empty:
                st.info("There are no scholars available. Select a class on Home Page or add scholars first.")
            else:
                sid=st.selectbox(
                    "Scholar",
                    list(roster.id.astype(int)),
                    format_func=lambda x:nm(roster[roster.id==x].iloc[0]),
                    key="assessment_student_v134"
                )
                scholar_name=nm(roster[roster.id==sid].iloc[0])
                br=benchmark_for_scholar(sid)
                def bv(k):
                    if not br or k not in br.keys():
                        return ""
                    return str(br[k] or "")

                st.markdown(f"### {scholar_name}")
                with st.form("assessment_entry_v134"):
                    st.markdown("### NWEA MAP Growth — RIT Scores")
                    r1,r2,r3=st.columns(3)
                    fall_read=r1.text_input("ELA / Reading — Fall",value=bv("nwea_fall_reading"),placeholder="RIT score")
                    winter_read=r2.text_input("ELA / Reading — Winter",value=bv("nwea_winter_reading"),placeholder="RIT score")
                    spring_read=r3.text_input("ELA / Reading — Spring",value=bv("nwea_spring_reading"),placeholder="RIT score")

                    m1,m2,m3=st.columns(3)
                    fall_math=m1.text_input("Math — Fall",value=bv("nwea_fall_math"),placeholder="RIT score")
                    winter_math=m2.text_input("Math — Winter",value=bv("nwea_winter_math"),placeholder="RIT score")
                    spring_math=m3.text_input("Math — Spring",value=bv("nwea_spring_math"),placeholder="RIT score")

                    st.markdown("### Individual NWEA Goal")
                    g1,g2=st.columns(2)
                    reading_goal=g1.text_input("ELA / Reading Goal",value=bv("nwea_reading_goal"),placeholder="Target RIT")
                    math_goal=g2.text_input("Math Goal",value=bv("nwea_math_goal"),placeholder="Target RIT")

                    st.markdown("### F&P")
                    f1,f2=st.columns(2)
                    fp_fall=f1.text_input("Fall F&P Reading Level",value=bv("fp_fall_level"),placeholder="Example: L")
                    fp_spring=f2.text_input("Spring F&P Reading Level",value=bv("fp_spring_level"),placeholder="Example: N")
                    w1,w2=st.columns(2)
                    word_fall=w1.text_input("Fall F&P Word List Level / Score",value=bv("fp_fall_word_list"))
                    word_spring=w2.text_input("Spring F&P Word List Level / Score",value=bv("fp_spring_word_list"))
                    notes=st.text_area("Assessment Notes",value=bv("notes"))

                    if st.form_submit_button("Save NWEA & F&P Data"):
                        c=conn()
                        c.execute("""INSERT INTO benchmark_scores(
                            scholar_id,nwea_fall_reading,nwea_winter_reading,nwea_spring_reading,
                            nwea_fall_math,nwea_winter_math,nwea_spring_math,nwea_reading_goal,nwea_math_goal,
                            fp_fall_level,fp_spring_level,fp_fall_word_list,fp_spring_word_list,notes)
                            VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)
                            ON CONFLICT(scholar_id) DO UPDATE SET
                            nwea_fall_reading=excluded.nwea_fall_reading,
                            nwea_winter_reading=excluded.nwea_winter_reading,
                            nwea_spring_reading=excluded.nwea_spring_reading,
                            nwea_fall_math=excluded.nwea_fall_math,
                            nwea_winter_math=excluded.nwea_winter_math,
                            nwea_spring_math=excluded.nwea_spring_math,
                            nwea_reading_goal=excluded.nwea_reading_goal,
                            nwea_math_goal=excluded.nwea_math_goal,
                            fp_fall_level=excluded.fp_fall_level,
                            fp_spring_level=excluded.fp_spring_level,
                            fp_fall_word_list=excluded.fp_fall_word_list,
                            fp_spring_word_list=excluded.fp_spring_word_list,
                            notes=excluded.notes""",
                            (sid,fall_read,winter_read,spring_read,fall_math,winter_math,spring_math,
                             reading_goal,math_goal,fp_fall,fp_spring,word_fall,word_spring,notes))
                        c.commit(); c.close()
                        st.success("Assessment data saved.")
                        st.rerun()

                br=benchmark_for_scholar(sid)
                if br:
                    st.markdown("### Progress Snapshot")
                    def n(v):
                        try: return float(v)
                        except: return None
                    rg=n(br["nwea_reading_goal"]) if "nwea_reading_goal" in br.keys() else None
                    mg=n(br["nwea_math_goal"]) if "nwea_math_goal" in br.keys() else None
                    left,right=st.columns(2)
                    with left:
                        st.markdown("**ELA / Reading**")
                        for season,key in [("Fall","nwea_fall_reading"),("Winter","nwea_winter_reading"),("Spring","nwea_spring_reading")]:
                            score=n(br[key]) if key in br.keys() else None
                            if score is not None:
                                status=""
                                if rg is not None:
                                    status=" ✅ Met/Exceeded Goal" if score>=rg else f" — {rg-score:.0f} points from goal"
                                st.write(f"- {season}: {score:g}{status}")
                    with right:
                        st.markdown("**Math**")
                        for season,key in [("Fall","nwea_fall_math"),("Winter","nwea_winter_math"),("Spring","nwea_spring_math")]:
                            score=n(br[key]) if key in br.keys() else None
                            if score is not None:
                                status=""
                                if mg is not None:
                                    status=" ✅ Met/Exceeded Goal" if score>=mg else f" — {mg-score:.0f} points from goal"
                                st.write(f"- {season}: {score:g}{status}")

            st.markdown("---")
            st.markdown("## 🎯 Class NWEA Goal Dashboard")
            if not selected_class:
                st.info("Select a class on Home Page to use the class goal dashboard.")
            else:
                season=st.radio("Testing Season",["Fall","Spring"],horizontal=True,key="class_nwea_season_v134")
                reading_goal_class=get_nwea_goal(selected_class,season,"Reading")
                math_goal_class=get_nwea_goal(selected_class,season,"Math")
                g1,g2=st.columns(2)
                rgc=g1.number_input("Class Reading / ELA Goal",0.0,400.0,float(reading_goal_class) if reading_goal_class is not None else 190.0,1.0,key="class_rg_v134")
                mgc=g2.number_input("Class Math Goal",0.0,400.0,float(math_goal_class) if math_goal_class is not None else 190.0,1.0,key="class_mg_v134")
                if st.button("Save Class NWEA Goals",key="save_class_goals_v134"):
                    save_nwea_goal(selected_class,season,"Reading",rgc)
                    save_nwea_goal(selected_class,season,"Math",mgc)
                    st.success("Class NWEA goals saved.")
                    st.rerun()
                class_scores=class_nwea_dataframe(selected_class,season)
                if not class_scores.empty:
                    class_scores["Reading Status"]=class_scores["Reading / ELA"].apply(lambda x:"✅ Met / Exceeded" if pd.notna(x) and x>=rgc else ("Below Goal" if pd.notna(x) else "No Score"))
                    class_scores["Math Status"]=class_scores["Math"].apply(lambda x:"✅ Met / Exceeded" if pd.notna(x) and x>=mgc else ("Below Goal" if pd.notna(x) else "No Score"))
                    st.dataframe(class_scores,hide_index=True,use_container_width=True)
                    rh,rl=nwea_rank_summary(class_scores,"Reading / ELA")
                    mh,ml=nwea_rank_summary(class_scores,"Math")
                    a,b=st.columns(2)
                    with a:
                        st.markdown("**Reading / ELA**")
                        if rh:
                            st.success(f"Highest: {', '.join(rh[0])} — {rh[1]:g}")
                            st.warning(f"Lowest: {', '.join(rl[0])} — {rl[1]:g}")
                    with b:
                        st.markdown("**Math**")
                        if mh:
                            st.success(f"Highest: {', '.join(mh[0])} — {mh[1]:g}")
                            st.warning(f"Lowest: {', '.join(ml[0])} — {ml[1]:g}")

        elif binder_tool=="Grade Settings":
            w=pd.DataFrame([{"Category":k,"Weight %":v} for k,v in get_setting("weights").items()])
            e=st.data_editor(w,num_rows="dynamic",hide_index=True,use_container_width=True,key="weights_v9")
            if st.button("Save Weights",key="save_weights_v9"):
                nw={str(r["Category"]):float(r["Weight %"]) for _,r in e.iterrows() if str(r["Category"]).strip()}
                if abs(sum(nw.values())-100)>.01: st.error("Weights must total 100%.")
                else: save_setting("weights",nw); st.success("Saved."); st.rerun()
            sc=pd.DataFrame(get_setting("scale"),columns=["Letter","Minimum","Maximum"])
            se=st.data_editor(sc,num_rows="dynamic",hide_index=True,use_container_width=True,key="scale_v9")
            if st.button("Save Letter Scale",key="save_scale_v9"):
                save_setting("scale",[[str(r.Letter),float(r.Minimum),float(r.Maximum)] for _,r in se.iterrows()]); st.success("Saved.")

elif page=="Book Leveler":
    st.markdown('<div class="page-title">Book Leveler</div><div class="page-subtitle">Search books online, save verified F&P levels, and check fit for a scholar.</div>',unsafe_allow_html=True)
    if not selected_class:
        st.info("Choose a class first.")
    else:
        st.markdown("## 📚 Book Level Checker")
        st.caption(
            "Search Open Library by title or ISBN, save books to your ChapLab catalog, "
            "and compare verified F&P levels with a scholar's current level."
        )

        roster=scholars_df(selected_class)
        if roster.empty:
            st.info("No scholars are available in this class.")
        else:
            sid=st.selectbox(
                "Scholar",
                list(roster.id.astype(int)),
                format_func=lambda x:nm(roster[roster.id==x].iloc[0]),
                key="book_checker_student_v32"
            )
            scholar_name=nm(roster[roster.id==sid].iloc[0])
            br=benchmark_for_scholar(sid)
            current_fp=""
            if br:
                current_fp=(br["fp_spring_level"] or br["fp_fall_level"] or "").strip().upper()

            if current_fp:
                st.success(f"{scholar_name}'s current F&P level: **{current_fp}**")
            else:
                st.warning(
                    f"No F&P level is saved for {scholar_name}. "
                    "You can still search books, but the reading-level comparison will be limited."
                )

            tabs=st.tabs(["Search Internet","My Book Catalog"])

            with tabs[0]:
                st.markdown("### Search Open Library")

                # Camera is intentionally user-activated. It is NOT rendered
                # until the teacher clicks the button below.
                if st.button("📷 Scan ISBN with Camera",key="toggle_book_camera"):
                    st.session_state["show_book_camera"]=not st.session_state.get("show_book_camera",False)
                    st.rerun()

                if st.session_state.get("show_book_camera",False):
                    with st.container(border=True):
                        cam_top1,cam_top2=st.columns([5,1])
                        cam_top1.markdown("#### 📷 ISBN Camera")
                        if cam_top2.button("✕ Close",key="close_book_camera"):
                            st.session_state["show_book_camera"]=False
                            st.rerun()
                        st.caption("Point the camera at the ISBN barcode. The camera stays off unless you open it.")
                        camera_photo=st.camera_input(
                            "Capture the ISBN barcode",
                            key="book_isbn_camera_capture"
                        )
                        if camera_photo is not None:
                            st.success("Barcode photo captured.")
                            st.caption("Type the ISBN printed above the barcode below to search for the book.")
                            captured_isbn=st.text_input(
                                "ISBN from captured book",
                                placeholder="978...",
                                key="book_camera_isbn"
                            )
                            cleaned_camera_isbn=re.sub(r"[^0-9Xx]","",captured_isbn or "")
                            if st.button("Find Captured Book",key="find_camera_book"):
                                if len(cleaned_camera_isbn) not in (10,13):
                                    st.warning("Enter a valid 10- or 13-digit ISBN.")
                                else:
                                    try:
                                        found=openlibrary_lookup_isbn(cleaned_camera_isbn)
                                        st.session_state["book_online_results"]=[found] if found else []
                                        if not found:
                                            st.warning("No matching book was found.")
                                    except Exception as e:
                                        st.error(str(e))

                search_mode=st.radio(
                    "Search by",
                    ["Title / Author / Keyword","ISBN"],
                    horizontal=True,
                    key="book_online_search_mode"
                )

                q=st.text_input(
                    "Book title, author, keywords, or ISBN",
                    placeholder="Example: Charlotte's Web or 9780064400558",
                    key="book_online_query"
                )

                if st.button("Search Internet",key="book_search_internet"):
                    try:
                        if search_mode=="ISBN":
                            result=openlibrary_lookup_isbn(q)
                            st.session_state["book_online_results"]=[result] if result else []
                        else:
                            st.session_state["book_online_results"]=openlibrary_search_books(q,limit=10)
                        if not st.session_state["book_online_results"]:
                            st.warning("No matching books were found.")
                    except Exception as e:
                        st.error(str(e))

                online_results=st.session_state.get("book_online_results",[])
                if online_results:
                    labels=[]
                    for i,b in enumerate(online_results):
                        title=b.get("title") or "Untitled"
                        author=b.get("author") or "Unknown author"
                        year=b.get("year") or b.get("publish_date") or ""
                        isbn=b.get("isbn") or ""
                        label=f"{title} — {author}"
                        if year:
                            label+=f" ({year})"
                        if isbn:
                            label+=f" | ISBN {isbn}"
                        labels.append(label)

                    idx=st.selectbox(
                        "Select a result",
                        list(range(len(online_results))),
                        format_func=lambda x:labels[x],
                        key="book_online_result_select"
                    )
                    book=online_results[idx]

                    cover_url=book.get("cover_url") or openlibrary_cover_url(book)
                    c1,c2=st.columns([1,2])
                    with c1:
                        if cover_url:
                            st.image(cover_url,width=180)
                        else:
                            st.caption("No cover image available.")
                    with c2:
                        st.markdown(f"**Title:** {book.get('title') or 'Untitled'}")
                        st.write(f"**Author:** {book.get('author') or 'Unknown'}")
                        if book.get("publisher"):
                            st.write(f"**Publisher:** {book.get('publisher')}")
                        if book.get("year") or book.get("publish_date"):
                            st.write(f"**Published:** {book.get('year') or book.get('publish_date')}")
                        if book.get("isbn"):
                            st.write(f"**ISBN:** {book.get('isbn')}")

                    st.info(
                        "Open Library does not reliably provide Fountas & Pinnell levels. "
                        "ChapLab will not guess an F&P level. Enter a verified level if you know it, "
                        "then save the book so future checks are automatic."
                    )

                    verified_fp=st.text_input(
                        "Verified F&P level (optional)",
                        placeholder="Example: M",
                        key="book_verified_fp_online"
                    ).strip().upper()

                    notes=st.text_area(
                        "Book notes",
                        placeholder="Optional notes about this edition, classroom use, or source of the verified level.",
                        height=80,
                        key="book_online_notes"
                    )

                    if verified_fp and not re.fullmatch(r"[A-Z]",verified_fp):
                        st.warning("For automatic comparison, use a single F&P letter level A–Z.")

                    if verified_fp and current_fp:
                        relation=fp_level_relation(current_fp,verified_fp)
                        if relation:
                            label,diff=relation
                            if diff==0:
                                st.success(f"This book is **on {scholar_name}'s current F&P level ({current_fp})**.")
                            elif diff<0:
                                st.info(
                                    f"This book is **{label.lower()}** for {scholar_name}: "
                                    f"book {verified_fp}, scholar {current_fp}."
                                )
                            else:
                                st.warning(
                                    f"This book is **{label.lower()}** for {scholar_name}: "
                                    f"book {verified_fp}, scholar {current_fp}."
                                )

                    if st.button("Save to My Book Catalog",key="save_online_book"):
                        c=conn()
                        # Inspect catalog columns dynamically so this remains compatible
                        # with earlier ChapLab database versions.
                        cols=[r["name"] for r in c.execute("PRAGMA table_info(book_catalog)").fetchall()]
                        title=(book.get("title") or "Untitled").strip()
                        author=(book.get("author") or "").strip()
                        isbn=(book.get("isbn") or "").strip()

                        existing=None
                        if isbn and "isbn" in cols:
                            existing=c.execute(
                                "SELECT id FROM book_catalog WHERE isbn=? LIMIT 1",
                                (isbn,)
                            ).fetchone()
                        if not existing:
                            existing=c.execute(
                                "SELECT id FROM book_catalog WHERE lower(title)=lower(?) AND lower(COALESCE(author,''))=lower(?) LIMIT 1",
                                (title,author)
                            ).fetchone()

                        if existing:
                            sets=[]
                            vals=[]
                            if "fp_level" in cols:
                                sets.append("fp_level=?"); vals.append(verified_fp)
                            if "isbn" in cols:
                                sets.append("isbn=?"); vals.append(isbn)
                            if "notes" in cols:
                                sets.append("notes=?"); vals.append(notes.strip())
                            if sets:
                                vals.append(int(existing["id"]))
                                c.execute(f"UPDATE book_catalog SET {','.join(sets)} WHERE id=?",vals)
                            st.success("Book already existed; saved information was updated.")
                        else:
                            insert_cols=["title"]
                            insert_vals=[title]
                            if "author" in cols:
                                insert_cols.append("author"); insert_vals.append(author)
                            if "fp_level" in cols:
                                insert_cols.append("fp_level"); insert_vals.append(verified_fp)
                            if "isbn" in cols:
                                insert_cols.append("isbn"); insert_vals.append(isbn)
                            if "notes" in cols:
                                insert_cols.append("notes"); insert_vals.append(notes.strip())
                            placeholders=",".join(["?"]*len(insert_cols))
                            c.execute(
                                f"INSERT INTO book_catalog({','.join(insert_cols)}) VALUES ({placeholders})",
                                insert_vals
                            )
                            st.success("Book saved to your ChapLab catalog.")
                        c.commit(); c.close()
                        st.rerun()

            with tabs[1]:
                st.markdown("### My Book Catalog")
                c=conn()
                catalog=pd.read_sql_query("SELECT * FROM book_catalog ORDER BY title",c)
                c.close()

                if catalog.empty:
                    st.caption("Your book catalog is empty. Use Search Internet to add books.")
                else:
                    search_local=st.text_input(
                        "Search saved books",
                        placeholder="Type title or author",
                        key="book_catalog_search_v32"
                    )
                    filtered=catalog.copy()
                    if search_local.strip():
                        mask=filtered["title"].astype(str).str.contains(search_local,case=False,na=False)
                        if "author" in filtered.columns:
                            mask=mask | filtered["author"].astype(str).str.contains(search_local,case=False,na=False)
                        filtered=filtered[mask]

                    if filtered.empty:
                        st.warning("No saved books match that search.")
                    else:
                        bid=st.selectbox(
                            "Saved book",
                            list(filtered.id.astype(int)),
                            format_func=lambda x:(
                                f"{filtered[filtered.id==x].iloc[0].title}"
                                + (
                                    f" — {filtered[filtered.id==x].iloc[0].author}"
                                    if "author" in filtered.columns and str(filtered[filtered.id==x].iloc[0].author or "").strip()
                                    else ""
                                )
                            ),
                            key="saved_book_select_v32"
                        )
                        row=filtered[filtered.id==bid].iloc[0]
                        saved_fp=str(row.fp_level or "").strip().upper() if "fp_level" in filtered.columns else ""

                        st.write(f"**Title:** {row.title}")
                        if "author" in filtered.columns and str(row.author or "").strip():
                            st.write(f"**Author:** {row.author}")
                        if "isbn" in filtered.columns and str(row.isbn or "").strip():
                            st.write(f"**ISBN:** {row.isbn}")
                        st.write(f"**Verified F&P:** {saved_fp or 'Not saved'}")

                        if saved_fp and current_fp:
                            relation=fp_level_relation(current_fp,saved_fp)
                            if relation:
                                label,diff=relation
                                if diff==0:
                                    st.success(f"On {scholar_name}'s current F&P level.")
                                elif diff<0:
                                    st.info(f"{label}: book {saved_fp}, scholar {current_fp}.")
                                else:
                                    st.warning(f"{label}: book {saved_fp}, scholar {current_fp}.")
                        elif not saved_fp:
                            st.info("Add a verified F&P level to compare this book automatically.")

                        st.markdown("#### Edit Saved Book")
                        with st.form("edit_saved_book_v32"):
                            etitle=st.text_input("Title",value=str(row.title or ""))
                            eauthor=st.text_input(
                                "Author",
                                value=str(row.author or "") if "author" in filtered.columns else ""
                            )
                            eisbn=st.text_input(
                                "ISBN",
                                value=str(row.isbn or "") if "isbn" in filtered.columns else ""
                            )
                            efp=st.text_input("Verified F&P",value=saved_fp)
                            enotes=st.text_area(
                                "Notes",
                                value=str(row.notes or "") if "notes" in filtered.columns else ""
                            )
                            if st.form_submit_button("Save Book Changes"):
                                c=conn()
                                cols=[r["name"] for r in c.execute("PRAGMA table_info(book_catalog)").fetchall()]
                                sets=["title=?"]; vals=[etitle.strip()]
                                if "author" in cols:
                                    sets.append("author=?"); vals.append(eauthor.strip())
                                if "isbn" in cols:
                                    sets.append("isbn=?"); vals.append(eisbn.strip())
                                if "fp_level" in cols:
                                    sets.append("fp_level=?"); vals.append(efp.strip().upper())
                                if "notes" in cols:
                                    sets.append("notes=?"); vals.append(enotes.strip())
                                vals.append(int(bid))
                                c.execute(f"UPDATE book_catalog SET {','.join(sets)} WHERE id=?",vals)
                                c.commit(); c.close()
                                st.success("Saved book updated.")
                                st.rerun()

                        confirm=st.checkbox("Delete this saved book",key="delete_saved_book_confirm_v32")
                        if st.button("Delete Book",disabled=not confirm,key="delete_saved_book_v32"):
                            c=conn()
                            c.execute("DELETE FROM book_catalog WHERE id=?",(int(bid),))
                            c.commit(); c.close()
                            st.success("Book deleted from catalog.")
                            st.rerun()



elif page=="Student Grouping":
    st.markdown('<div class="page-title">Student Grouping</div><div class="page-subtitle">Group scholars by subject performance or skill.</div>',unsafe_allow_html=True)
    st.write("Create Low, Mid, and High groups using either overall subject grades or a specific skill/standard.")

    active_class=selected_class if selected_class else None
    low_default,mid_default=get_grouping_cutoffs(active_class)

    st.markdown("### Group Cutoffs")
    c1,c2=st.columns(2)
    low_max=c1.number_input("Low group: score at or below",min_value=0.0,max_value=100.0,value=float(low_default),step=1.0)
    mid_max=c2.number_input("Mid group: score at or below",min_value=0.0,max_value=100.0,value=float(mid_default),step=1.0)

    st.caption(f"Current rule: Low ≤ {low_max:g}%, Mid = above {low_max:g}% through {mid_max:g}%, High > {mid_max:g}%.")

    if mid_max<=low_max:
        st.error("The Mid cutoff must be higher than the Low cutoff.")
    else:
        if st.button("Save Group Cutoffs"):
            save_grouping_cutoffs(active_class,low_max,mid_max)
            st.success("Grouping cutoffs saved.")
            st.rerun()

        mode=st.radio("Group students by",["Overall Subject Grade","Specific Skill / Standard"],horizontal=True)
        subject=st.selectbox("Subject",["ELA","Math","Science","Social Studies"],key="group_subject")

        if mode=="Overall Subject Grade":
            df=subject_grouping_df(active_class,subject,low_max,mid_max)
            if "Group" not in df.columns:
                df["Group"]="No Data"

            a,b,c=st.columns(3)
            a.metric("Low",int((df["Group"]=="Low").sum()))
            b.metric("Mid",int((df["Group"]=="Mid").sum()))
            c.metric("High",int((df["Group"]=="High").sum()))

            st.markdown(f"### {subject} Groups")
            tabs=st.tabs(["Low","Mid","High","No Data","All"])

            for tab,label in zip(tabs,["Low","Mid","High","No Data",None]):
                with tab:
                    show=df if label is None else df[df["Group"]==label]
                    if show.empty:
                        st.caption("No scholars in this group.")
                    else:
                        st.dataframe(show,hide_index=True,use_container_width=True)

        else:
            sdf=standards_df(subject)
            if sdf.empty:
                st.info("No skills/standards are available for this subject.")
            else:
                skill_option=st.selectbox(
                    "Skill / Standard",
                    [f"{r.code} — {r.skill}" for _,r in sdf.iterrows()]
                )
                skill_code=skill_option.split(" — ")[0]
                skill_name=skill_option.split(" — ",1)[1] if " — " in skill_option else skill_code

                df=skill_grouping_df(active_class,subject,skill_code,low_max,mid_max)
                if "Group" not in df.columns:
                    df["Group"]="No Data"

                a,b,c=st.columns(3)
                a.metric("Low",int((df["Group"]=="Low").sum()))
                b.metric("Mid",int((df["Group"]=="Mid").sum()))
                c.metric("High",int((df["Group"]=="High").sum()))

                st.markdown(f"### {subject}: {skill_name}")
                tabs=st.tabs(["Low","Mid","High","No Data","All"])
                for tab,label in zip(tabs,["Low","Mid","High","No Data",None]):
                    with tab:
                        show=df if label is None else df[df["Group"]==label]
                        if show.empty:
                            st.caption("No scholars in this group.")
                        else:
                            st.dataframe(show,hide_index=True,use_container_width=True)

                st.info("Use these groups for reteach, on-level practice, enrichment, centers, or flexible small groups. Re-group whenever new grades are entered.")

elif page=="Report Card Comments":
    st.markdown('<div class="page-title">Report Card Comments</div><div class="page-subtitle">Generate parent-friendly comments using quarter data.</div>',unsafe_allow_html=True)
    roster=scholars_df(selected_class or None)
    academic_year=current_academic_year()

    if roster.empty:
        st.info("Select a class with scholars first.")
    else:
        q=st.text_input("Find scholar",placeholder="Type first or last name",key="rc_find_scholar")
        filtered=roster
        if q.strip():
            filtered=filtered[(filtered.first_name+" "+filtered.last_name).str.contains(q.strip(),case=False,na=False)]
        if filtered.empty:
            st.warning("No matching scholar.")
        else:
            sid=st.selectbox("Scholar",list(filtered.id.astype(int)),
                             format_func=lambda x:nm(filtered[filtered.id==x].iloc[0]),
                             key="rc_scholar")
            name=nm(filtered[filtered.id==sid].iloc[0])
            subj=st.selectbox("Subject",["ELA","Math","Science","Social Studies"],key="rc_subject")
            mp=st.selectbox("Marking Period",["Quarter 1","Quarter 2","Quarter 3","Quarter 4"],key="rc_marking_period")

            due=report_deadline(academic_year,mp)
            if due:
                st.info(f"📅 {mp} report-card comments are due: {due}")

            qdf=quarter_grade_data(sid,subj,mp)
            missing,total_assignments=quarter_missing_grade_check(sid,subj,mp,selected_class)

            if total_assignments:
                complete_count=total_assignments-len(missing)
                st.write(f"**Grade completion:** {complete_count} of {total_assignments} assignments entered.")
                if missing:
                    st.warning("Missing grades: " + ", ".join(missing[:8]) + ("..." if len(missing)>8 else ""))
            else:
                st.caption("No assignments are currently tagged to this quarter/subject.")

            if not qdf.empty:
                qavg=weighted_average_from_grade_rows(qdf)
                a,b,c=st.columns(3)
                a.metric(f"{mp} Average",f"{qavg:.1f}% ({letter(qavg)})")
                b.metric("Graded Assignments",len(qdf))
                c.metric("Score Range",f"{qdf['pct'].min():.0f}%–{qdf['pct'].max():.0f}%")

                st.markdown("### Quarter Grade Evidence")
                display=qdf.copy()
                display["Grade"]=display["pct"].map(lambda x:f"{x:.1f}%")
                st.dataframe(display[["assignment_date","title","category","standard_code","Grade"]],
                             hide_index=True,use_container_width=True)

                # quarter-to-quarter growth
                quarter_order=["Quarter 1","Quarter 2","Quarter 3","Quarter 4"]
                mp_index=quarter_order.index(mp)
                if mp_index>0:
                    prev=quarter_order[mp_index-1]
                    prevdf=quarter_grade_data(sid,subj,prev)
                    if not prevdf.empty:
                        prevavg=weighted_average_from_grade_rows(prevdf)
                        delta=qavg-prevavg
                        st.write(f"**Quarter-to-quarter change:** {prev} {prevavg:.1f}% → {mp} {qavg:.1f}% ({delta:+.1f} points)")

            skill_options=[r.skill for _,r in standards_df(subj).iterrows()]
            selected_next_skills=st.multiselect("Upcoming skills to include",skill_options,key="rc_next_skills")
            next_skill_text=", ".join(selected_next_skills)

            st.markdown("### Comment Controls")
            c1,c2=st.columns(2)
            comment_length=c1.selectbox("Comment Length",["Short","Standard","Detailed"],index=1,key="rc_length")
            max_chars=c2.number_input("Maximum Characters",min_value=100,max_value=5000,value=1000,step=50,key="rc_max_chars")

            include_overall=st.checkbox("Include overall quarter grade",value=True,key="rc_include_overall")
            include_growth=st.checkbox("Include quarter-to-quarter growth",value=True,key="rc_include_growth")
            include_strength=st.checkbox("Include strongest skill",value=True,key="rc_include_strength")
            include_need=st.checkbox("Include area for improvement",value=True,key="rc_include_need")
            include_assignment=st.checkbox("Include assignment examples",value=True,key="rc_include_assignment")
            include_home=st.checkbox("Include how family can help",value=True,key="rc_include_home")
            include_teacher=st.checkbox("Include what teacher will do",value=True,key="rc_include_teacher")
            include_next=st.checkbox("Include upcoming skills",value=True,key="rc_include_next")

            if st.button("Generate Quarter-Based Comment",key="rc_generate"):
                base=quarter_report_comment(name,subj,mp,qdf,next_skill_text if include_next else "",sid=sid)
                # Simple trimming/removal based on controls
                comment=base
                if not include_assignment:
                    comment=re.sub(r' For example,.*?helpful\.', '', comment)
                if not include_home:
                    comment=re.sub(r' At home,.*?(?= In class,| Our next|$)', '', comment)
                if not include_teacher:
                    comment=re.sub(r' In class,.*?(?= Our next|$)', '', comment)
                if not include_next:
                    comment=re.sub(r' Our next instructional focus.*$', '', comment)
                if not include_growth:
                    comment=comment.replace(" improved over the course of the marking period","")
                    comment=comment.replace(" performed fairly consistently across the marking period","")
                    comment=comment.replace(" showed some decline later in the marking period","")
                if not include_overall:
                    comment=re.sub(r'During .*? graded assignments?\.', f"During {mp}, {name} completed graded {subj.lower()} work.", comment, count=1)
                if comment_length=="Short":
                    sentences=re.split(r'(?<=[.!?])\s+',comment)
                    comment=" ".join(sentences[:4])
                elif comment_length=="Standard":
                    sentences=re.split(r'(?<=[.!?])\s+',comment)
                    comment=" ".join(sentences[:6])
                if len(comment)>max_chars:
                    comment=comment[:max_chars].rsplit(" ",1)[0].rstrip(" ,;:")+"."
                st.session_state["comment"]=comment
                st.session_state["rc_comment_text"]=comment

            if "rc_comment_text" not in st.session_state:
                st.session_state["rc_comment_text"]=st.session_state.get("comment","")
            txt=st.text_area("Edit / finalize / copy",height=300,key="rc_comment_text")
            st.caption(f"Character count: {len(txt)} / {int(max_chars)}")

            if st.button("Save Final Comment",key="rc_save_final"):
                c=conn()
                c.execute("""INSERT INTO report_comments(scholar_id,created_at,subject,marking_period,comment_text)
                             VALUES (?,?,?,?,?)""",
                          (sid,datetime.now().isoformat(timespec="minutes"),subj,mp,txt))
                c.commit(); c.close()
                row=closeout_row(sid,mp,academic_year)
                save_closeout(sid,mp,academic_year,
                              bool(row["grades_reviewed"]) if row else False,
                              bool(row["data_reviewed"]) if row else False,
                              True,True)
                st.success("Saved to the scholar profile and marked finalized.")
                st.rerun()

            st.markdown("---")
            st.markdown("## Quarter Closeout")
            close=closeout_row(sid,mp,academic_year)
            g=bool(close["grades_reviewed"]) if close else False
            d=bool(close["data_reviewed"]) if close else False
            cg=bool(close["comment_generated"]) if close else False
            cf=bool(close["comment_finalized"]) if close else False

            q1,q2,q3,q4=st.columns(4)
            g_new=q1.checkbox("Grades complete/reviewed",value=g,key="close_grades")
            d_new=q2.checkbox("Data reviewed",value=d,key="close_data")
            cg_new=q3.checkbox("Comment generated",value=cg or bool(txt),key="close_generated")
            cf_new=q4.checkbox("Comment finalized",value=cf,key="close_finalized")
            if st.button("Save Closeout Status",key="save_closeout_status"):
                save_closeout(sid,mp,academic_year,g_new,d_new,cg_new,cf_new)
                st.success("Closeout status saved.")
                st.rerun()

            c=conn()
            saved_comments=pd.read_sql_query("SELECT * FROM report_comments WHERE scholar_id=? ORDER BY created_at DESC",c,params=[sid])
            c.close()
            if not saved_comments.empty:
                st.markdown("### Saved Report Card Comments")
                st.dataframe(saved_comments[["id","created_at","subject","marking_period","comment_text"]],
                             hide_index=True,use_container_width=True)

elif page=="Little Assistant":
    st.markdown('<div class="page-title">Little Assistant</div><div class="page-subtitle">Parent messages, call scripts, IAT referrals, and student-support tools.</div>',unsafe_allow_html=True)
    st.caption("Support tools for teacher paperwork and family communication.")
    tool=st.radio(
        "Choose a tool",
        ["IEP / Student Support","IAT Referral","Parent Message","Phone Call Script"],
        horizontal=True,
        key="assistant_tool"
    )

    roster=scholars_df(selected_class or None)
    if roster.empty:
        st.info("Select a class with scholars first.")
    else:
        sid=st.selectbox("Scholar",list(roster.id.astype(int)),
                         format_func=lambda x:nm(roster[roster.id==x].iloc[0]),
                         key="assistant_scholar")
        scholar=roster[roster.id==sid].iloc[0]
        name=nm(scholar)
        pro=scholar_pronouns(sid)
        subj_pr=pro["subject"]
        poss_pr=pro["possessive"]

        if tool=="IEP / Student Support":
            st.markdown("### Profile & Data Prefill")
            prefill=iep_prefill_summary(sid)
            st.text_area("Existing scholar data",value=prefill,height=260,disabled=True)

            concern_choices=[
                "Academic performance","Reading fluency","Reading comprehension","Written expression",
                "Spelling / conventions","Math computation","Math problem solving","Attention / focus",
                "Task completion","Following directions","Organization / materials","Behavior / self-management",
                "Peer interaction","Communication","Participation","Independence","Other / Custom"
            ]
            frequency_choices=[
                "Rarely","1–2 times per week","3–4 times per week","Daily","Multiple times per day",
                "Only during independent work","Only during whole group","Only during small group",
                "Across settings","Other / Custom"
            ]
            intervention_choices=[
                "Small-group reteach","1:1 teacher conference","Repeated directions","Directions broken into steps",
                "Visual reminder / checklist","Graphic organizer","Preferential seating","Extended time",
                "Reduced task length","Frequent check-ins","Model/example provided","Read directions aloud",
                "Peer support","Positive reinforcement","Behavior reminder / redirection",
                "Practice with corrected work","Home practice requested","Other / Custom"
            ]
            response_choices=[
                "Improved with support","Some improvement","Improvement was temporary","No noticeable improvement yet",
                "Needed repeated prompting","Completed with adult support","Completed independently after support",
                "Performance remained inconsistent","Other / Custom"
            ]
            impact_choices=[
                "Reduces accuracy","Slows work completion","Makes independent work difficult","Affects participation",
                "Affects comprehension","Affects written output","Affects retention / recall",
                "Affects ability to follow multi-step directions","Affects peer/classroom interactions",
                "Affects consistent academic progress","Other / Custom"
            ]

            with st.form("assistant_iep_form",clear_on_submit=True):
                note_type=st.selectbox("Documentation type",[
                    "Academic","Reading","Writing","Math","Attention / task completion",
                    "Behavior / self-management","Communication","Social interaction",
                    "Intervention response","Other"
                ])
                concern=st.selectbox("Area of concern",concern_choices)
                custom_concern=st.text_input("Custom concern (if needed)")
                area=custom_concern.strip() if concern=="Other / Custom" and custom_concern.strip() else concern
                observation=st.text_area("Objective observation")
                freq_choice=st.selectbox("Frequency / pattern",frequency_choices)
                custom_freq=st.text_input("Custom frequency/pattern")
                frequency=custom_freq.strip() if freq_choice=="Other / Custom" and custom_freq.strip() else freq_choice
                intervention_choice=st.selectbox("Primary support/intervention tried",intervention_choices)
                custom_intervention=st.text_area("Additional intervention details")
                intervention=(intervention_choice if intervention_choice!="Other / Custom" else "")
                if custom_intervention.strip():
                    intervention=(intervention+(": " if intervention else "")+custom_intervention.strip())
                response_choice=st.selectbox("Response to support",response_choices)
                response_details=st.text_area("Additional response details")
                response=(response_choice if response_choice!="Other / Custom" else "")
                if response_details.strip():
                    response=(response+(": " if response else "")+response_details.strip())
                impact_choice=st.selectbox("Educational impact",impact_choices)
                impact_details=st.text_area("Additional impact details")
                impact=(impact_choice if impact_choice!="Other / Custom" else "")
                if impact_details.strip():
                    impact=(impact+(": " if impact else "")+impact_details.strip())

                if st.form_submit_button("Save Evidence Note"):
                    c=conn()
                    c.execute("""INSERT INTO support_notes(
                        scholar_id,created_at,note_type,area,observation,frequency,intervention,
                        response_to_intervention,impact,concern_category,frequency_choice,
                        intervention_choice,response_choice,impact_choice)
                        VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                        (sid,datetime.now().isoformat(timespec="minutes"),note_type,area,observation,
                         frequency,intervention,response,impact,concern,freq_choice,
                         intervention_choice,response_choice,impact_choice))
                    c.commit(); c.close()
                    st.success("Evidence note saved.")
                    st.rerun()

            c=conn()
            notes=pd.read_sql_query("SELECT * FROM support_notes WHERE scholar_id=? ORDER BY created_at DESC",c,params=[sid])
            ws=pd.read_sql_query("SELECT * FROM work_samples WHERE scholar_id=? ORDER BY uploaded_at DESC",c,params=[sid])
            c.close()
            if st.button("Generate Teacher Support Summary",key="assistant_generate_support"):
                parts=[prefill,"\nTeacher Observations / Interventions:"]
                for _,r in notes.head(8).iterrows():
                    parts.append(f"- {r.area}: {r.observation} Frequency: {r.frequency}. Support: {r.intervention}. Response: {r.response_to_intervention}. Impact: {r.impact}.")
                if not ws.empty:
                    parts.append("\nWork-Sample Evidence:")
                    for _,r in ws.head(5).iterrows():
                        parts.append(f"- {r.title or r.file_name}: strengths {r.strengths or '—'}; needs {r.needs or '—'}; next steps {r.next_steps or '—'}.")
                parts.append("\nThis summary documents classroom evidence and teacher observations; it is not a diagnosis.")
                st.session_state["assistant_support_summary"]="\n".join(parts)
            st.text_area("Editable teacher support summary",value=st.session_state.get("assistant_support_summary",""),height=380)

        elif tool=="IAT Referral":
            st.markdown("### IAT Referral Writer")
            st.caption(
                "Build a classroom-based referral using the scholar's existing academic data, "
                "intervention evidence, work samples, and family-contact history."
            )

            # Existing scholar data
            prefill=iep_prefill_summary(sid)
            strengths,needs,teacher_actions=academic_summary_for_scholar(sid)
            br=benchmark_for_scholar(sid)

            c=conn()
            support_history=pd.read_sql_query(
                """SELECT created_at,note_type,area,observation,frequency,intervention,
                          response_to_intervention,impact
                   FROM support_notes
                   WHERE scholar_id=?
                   ORDER BY created_at DESC""",
                c,params=[sid]
            )
            work_history=pd.read_sql_query(
                """SELECT uploaded_at,title,file_name,strengths,needs,next_steps
                   FROM work_samples
                   WHERE scholar_id=?
                   ORDER BY uploaded_at DESC""",
                c,params=[sid]
            )
            contact_history=pd.read_sql_query(
                """SELECT communications.created_at,communications.communication_type,
                          communications.reason,communications.details,
                          COALESCE(TRIM(guardians.first_name||' '||guardians.last_name),'') guardian
                   FROM communications
                   LEFT JOIN guardians ON guardians.id=communications.guardian_id
                   WHERE communications.scholar_id=?
                   ORDER BY communications.id DESC""",
                c,params=[sid]
            )
            c.close()

            # Snapshot shown to teacher before writing referral
            with st.expander("Scholar Data Being Used",expanded=False):
                st.text_area("Profile / academic prefill",value=prefill,height=230,disabled=True)

                if br:
                    b1,b2,b3=st.columns(3)
                    b1.metric("F&P",br["fp_spring_level"] or br["fp_fall_level"] or "—")
                    b2.metric("NWEA Reading",br["nwea_spring_reading"] or br["nwea_winter_reading"] or br["nwea_fall_reading"] or "—")
                    b3.metric("NWEA Math",br["nwea_spring_math"] or br["nwea_winter_math"] or br["nwea_fall_math"] or "—")

                if strengths:
                    st.markdown("**Current strengths**")
                    for subj,skill,pct in strengths[:5]:
                        st.write(f"- {subj}: {skill} — {pct:.0f}%")

                if needs:
                    st.markdown("**Current areas of need**")
                    for subj,skill,pct in needs[:5]:
                        st.write(f"- {subj}: {skill} — {pct:.0f}%")

            referral_reason_choices=[
                "Academic progress below expectations",
                "Reading difficulty",
                "Writing difficulty",
                "Math difficulty",
                "Attention / focus concerns",
                "Task completion / work habits",
                "Behavior / self-management",
                "Social / peer interaction",
                "Communication concerns",
                "Attendance / missed instruction",
                "Multiple areas of concern",
                "Other / Custom"
            ]
            concern_area_choices=[
                "Reading fluency","Reading comprehension","Phonics / decoding","Vocabulary",
                "Written expression","Spelling / conventions","Math computation","Math problem solving",
                "Following directions","Attention / focus","Task initiation","Task completion",
                "Organization","Behavior regulation","Peer interaction","Participation",
                "Independence","Communication","Other / Custom"
            ]
            requested_support_choices=[
                "Problem-solving team review",
                "Additional academic intervention",
                "Reading intervention",
                "Math intervention",
                "Behavior / self-management support",
                "Observation by support staff",
                "Progress-monitoring plan",
                "Additional classroom strategies",
                "Family-school support plan",
                "Consideration for further evaluation",
                "Other / Custom"
            ]

            st.markdown("#### Referral Information")
            c1,c2=st.columns(2)
            reason_choice=c1.selectbox("Primary reason for referral",referral_reason_choices,key="iat_reason")
            reason_custom=c2.text_input("Custom reason (if needed)",key="iat_reason_custom")
            referral_reason=reason_custom.strip() if reason_choice=="Other / Custom" and reason_custom.strip() else reason_choice

            selected_concerns=st.multiselect(
                "Areas of concern",
                concern_area_choices,
                key="iat_concerns"
            )
            custom_concern=st.text_input("Additional/custom concern",key="iat_custom_concern")
            concerns=[x for x in selected_concerns if x!="Other / Custom"]
            if custom_concern.strip():
                concerns.append(custom_concern.strip())

            a,b=st.columns(2)
            onset=a.selectbox(
                "How long has the concern been observed?",
                ["Less than 1 month","1–2 months","3–4 months","Most of the school year","Since the beginning of the school year","Other"],
                key="iat_duration"
            )
            setting=b.multiselect(
                "Where is the concern most noticeable?",
                ["Whole group","Small group","Independent work","Assessments","Transitions","Unstructured time","Across settings"],
                key="iat_settings"
            )

            objective_description=st.text_area(
                "Objective description of the concern",
                placeholder="Describe what you are seeing using observable, specific language. Include examples when possible.",
                height=125,
                key="iat_objective"
            )

            impact=st.text_area(
                "How is this affecting the scholar's progress or access to instruction?",
                placeholder="Example: The scholar needs repeated prompting to begin independent work and often completes less than half of the assigned task.",
                height=105,
                key="iat_impact"
            )

            st.markdown("#### Supports Already Tried")
            selected_interventions=st.multiselect(
                "Classroom interventions / supports",
                [
                    "Small-group reteach","1:1 teacher conference","Preferential seating",
                    "Directions broken into steps","Visual checklist / reminder","Frequent check-ins",
                    "Extended time","Reduced task length","Graphic organizer","Model/example provided",
                    "Repeated practice","Read directions aloud","Peer support","Positive reinforcement",
                    "Behavior reminder / redirection","Home practice requested","Family contacted",
                    "Progress monitoring","Other"
                ],
                key="iat_interventions"
            )
            intervention_details=st.text_area(
                "Intervention details",
                placeholder="Include frequency, duration, and what the support looked like.",
                height=105,
                key="iat_intervention_details"
            )
            response_to_support=st.selectbox(
                "Overall response to interventions",
                [
                    "Improved with support",
                    "Some improvement, but concern remains",
                    "Improvement was temporary",
                    "Performance remains inconsistent",
                    "Minimal or no improvement",
                    "More data is needed"
                ],
                key="iat_response"
            )

            st.markdown("#### Strengths & Family Communication")
            strengths_note=st.text_area(
                "Scholar strengths to highlight",
                value="; ".join([f"{subj}: {skill}" for subj,skill,pct in strengths[:4]]) if strengths else "",
                height=90,
                key="iat_strengths"
            )
            family_note=st.text_area(
                "Family communication / input",
                placeholder="Summarize relevant parent/guardian communication, concerns, or strategies discussed.",
                height=90,
                key="iat_family"
            )

            request_choice=st.multiselect(
                "What support are you requesting from IAT?",
                requested_support_choices,
                key="iat_request"
            )
            custom_request=st.text_area("Additional request / question for the team",height=85,key="iat_custom_request")

            if st.button("Generate IAT Referral",key="generate_iat_referral"):
                concern_text=", ".join(concerns) if concerns else referral_reason
                setting_text=", ".join(setting) if setting else "across classroom instruction"
                intervention_text=", ".join(selected_interventions) if selected_interventions else "classroom supports documented by the teacher"
                request_text=", ".join([x for x in request_choice if x!="Other / Custom"])
                if custom_request.strip():
                    request_text=(request_text + ("; " if request_text else "") + custom_request.strip())
                if not request_text:
                    request_text="team guidance regarding appropriate next steps and additional supports"

                # Pull a concise current grade snapshot
                grade_parts=[]
                for subj in ["ELA","Math","Science","Social Studies"]:
                    av,_=summary(sid,subj)
                    if av is not None:
                        grade_parts.append(f"{subj}: {av:.1f}% ({letter(av)})")
                grades_text=", ".join(grade_parts) if grade_parts else "limited current grade data"

                benchmark_parts=[]
                if br:
                    fp=br["fp_spring_level"] or br["fp_fall_level"]
                    nr=br["nwea_spring_reading"] or br["nwea_winter_reading"] or br["nwea_fall_reading"]
                    nmth=br["nwea_spring_math"] or br["nwea_winter_math"] or br["nwea_fall_math"]
                    if fp: benchmark_parts.append(f"F&P level {fp}")
                    if nr: benchmark_parts.append(f"NWEA Reading RIT {nr}")
                    if nmth: benchmark_parts.append(f"NWEA Math RIT {nmth}")
                benchmark_text=", ".join(benchmark_parts) if benchmark_parts else "no benchmark scores currently entered"

                evidence_lines=[]
                for _,r in support_history.head(4).iterrows():
                    evidence_lines.append(
                        f"{r.area}: {r.observation or 'documented concern'}; "
                        f"support: {r.intervention or '—'}; response: {r.response_to_intervention or '—'}"
                    )
                evidence_text=" | ".join(evidence_lines) if evidence_lines else "No prior intervention notes are currently saved."

                referral=(
                    f"IAT REFERRAL – {name}\n\n"
                    f"Reason for Referral:\n"
                    f"{referral_reason}. The primary areas of concern are {concern_text}. "
                    f"The concern has been observed for {onset.lower()} and is most noticeable during {setting_text}.\n\n"
                    f"Scholar Strengths:\n"
                    f"{strengths_note.strip() or f"{name} demonstrates strengths that should continue to be built upon during intervention planning."}\n\n"
                    f"Current Academic Data:\n"
                    f"Current classroom grades: {grades_text}. Current benchmark information: {benchmark_text}.\n\n"
                    f"Description of Concern:\n"
                    f"{objective_description.strip() or 'Teacher observation details should be added here.'}\n\n"
                    f"Educational Impact:\n"
                    f"{impact.strip() or f"The concern is affecting {poss_pr} consistent progress and/or independent access to classroom instruction."}\n\n"
                    f"Interventions / Supports Attempted:\n"
                    f"{intervention_text}. {intervention_details.strip()} "
                    f"Overall response: {response_to_support}.\n\n"
                    f"Existing Classroom Evidence:\n"
                    f"{evidence_text}\n\n"
                    f"Family Communication / Input:\n"
                    f"{family_note.strip() or 'Family communication details should be added if applicable.'}\n\n"
                    f"Request to IAT:\n"
                    f"I am requesting {request_text}. I would like the team to review the available data, "
                    f"consider additional supports, and help determine appropriate next steps for {name}."
                )
                st.session_state["iat_referral_text"]=referral
                editor_key=f"iat_referral_editor_{sid}"
                st.session_state[editor_key]=referral
                st.success("IAT referral generated below. You can edit it before saving.")

            editor_key=f"iat_referral_editor_{sid}"
            if editor_key not in st.session_state:
                st.session_state[editor_key]=st.session_state.get("iat_referral_text","")
            referral_text=st.text_area(
                "Editable IAT Referral",
                height=560,
                key=editor_key
            )

            action1,action2=st.columns(2)
            if action1.button("Save IAT Referral to Scholar Record",key="save_iat_referral"):
                if not referral_text.strip():
                    st.warning("Generate or enter the referral first.")
                else:
                    c=conn()
                    c.execute(
                        """INSERT INTO support_notes(
                            scholar_id,created_at,note_type,area,observation,frequency,
                            intervention,response_to_intervention,impact,concern_category,
                            frequency_choice,intervention_choice,response_choice,impact_choice)
                            VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                        (
                            sid,
                            datetime.now().isoformat(timespec="minutes"),
                            "IAT Referral",
                            referral_reason,
                            referral_text.strip(),
                            onset,
                            ", ".join(selected_interventions),
                            response_to_support,
                            impact.strip(),
                            referral_reason,
                            onset,
                            ", ".join(selected_interventions),
                            response_to_support,
                            ", ".join(request_choice)
                        )
                    )
                    c.commit(); c.close()
                    st.success("IAT referral saved to the scholar's support history.")

            action2.download_button(
                "Download IAT Referral",
                data=referral_text or "No IAT referral generated yet.",
                file_name=f"{name.replace(' ','_')}_IAT_Referral.txt",
                mime="text/plain",
                key="download_iat_referral"
            )

            if not support_history.empty:
                with st.expander("Prior Intervention / Support Notes"):
                    st.dataframe(
                        support_history[
                            ["created_at","note_type","area","observation","frequency",
                             "intervention","response_to_intervention","impact"]
                        ],
                        hide_index=True,
                        use_container_width=True
                    )

            if not contact_history.empty:
                with st.expander("Recent Family Contact History"):
                    st.dataframe(contact_history.head(8),hide_index=True,use_container_width=True)

            if not work_history.empty:
                with st.expander("Recent Work-Sample Evidence"):
                    st.dataframe(work_history.head(8),hide_index=True,use_container_width=True)

        elif tool=="Parent Message":
            gmap=guardian_display_map(sid)
            pref=update_preference(sid)
            default_gid=0
            if pref and pref["preferred_guardian_id"] and int(pref["preferred_guardian_id"]) in gmap:
                default_gid=int(pref["preferred_guardian_id"])

            st.markdown("### Parent Update Preference")
            requested=st.checkbox("Parent requested regular updates",
                                  value=bool(pref["requested_updates"]) if pref else False,
                                  key="assistant_requested_updates")
            frequency_options=["","Weekly","Every 2 Weeks","Monthly","End of Quarter","As Needed"]
            current_freq=pref["update_frequency"] if pref and pref["update_frequency"] in frequency_options else ""
            frequency=st.selectbox("Update frequency",frequency_options,index=frequency_options.index(current_freq),key="assistant_update_frequency")
            gids=list(gmap)
            gid=st.selectbox("Parent / Guardian",gids,format_func=lambda x:gmap[x],
                             index=gids.index(default_gid) if default_gid in gids else 0,
                             key="assistant_msg_guardian")
            pref_notes=st.text_area("Update preference notes",
                                    value=pref["notes"] if pref else "",
                                    key="assistant_update_pref_notes")
            if st.button("Save Parent Update Preference",key="save_update_preference"):
                save_update_preference(sid,requested,frequency,gid,pref_notes)
                st.success("Parent update preference saved.")
                st.rerun()

            subj=st.selectbox("Subject",["General","ELA","Math","Science","Social Studies"],key="assistant_msg_subject")
            reason=st.selectbox("Message Type",[
                "Progress update","Growth update","At-risk / concern update","Positive update",
                "Homework reminder","Missing assignment","Behavior concern","Academic concern",
                "Injury / classroom incident","General update"
            ],key="assistant_msg_reason")
            details=st.text_area("Specific information to add",key="assistant_msg_details")

            if st.button("Generate Data-Informed Parent Update",key="assistant_generate_msg"):
                parent_name=gmap.get(gid,"").split(" (")[0] if gid else ""
                greeting=f"Hello {parent_name}," if parent_name else "Hello,"
                if reason in ["Progress update","Growth update","At-risk / concern update"]:
                    body=generate_parent_progress_update(sid,subj)
                    if details.strip():
                        body += " " + details.strip()
                    msg=f"{greeting} {body}"
                elif reason=="Positive update":
                    msg=f"{greeting} I wanted to share a positive update about {name}. {details} Thank you for your continued support."
                elif reason=="Homework reminder":
                    msg=f"{greeting} this is a reminder regarding {name}'s {subj} homework. {details} Please have the assignment completed and returned as soon as possible. Thank you."
                elif reason=="Behavior concern":
                    msg=f"{greeting} I am reaching out regarding {name}'s behavior today. {details} Please speak with {name} about making choices that support learning and classroom expectations. Thank you for partnering with me."
                elif reason=="Injury / classroom incident":
                    msg=f"{greeting} I am contacting you to make you aware of an incident involving {name} today. {details} I wanted to make sure you received this information directly."
                else:
                    msg=f"{greeting} I am reaching out with an update regarding {name}. {details} Thank you for your support."
                st.session_state["assistant_parent_msg"]=msg

            msg=st.text_area("Edit / copy message",value=st.session_state.get("assistant_parent_msg",""),height=260)
            if st.button("Save Message to Communication Log",key="assistant_log_msg"):
                log_communication(sid,gid,"Text / School Message",subj,reason,details,msg)
                st.success("Message saved to the selected parent's communication log.")

        elif tool=="Phone Call Script":
            gmap=guardian_display_map(sid)
            gid=st.selectbox("Parent / Guardian",list(gmap),format_func=lambda x:gmap[x],key="assistant_call_guardian")
            subj=st.selectbox("Subject",["ELA","Math","Science","Social Studies","General"],key="assistant_call_subject")
            reason=st.selectbox("Reason",["Behavior concern","Academic concern","Missing work","Positive update","Injury / classroom incident","Conference request","General update"],key="assistant_call_reason")
            details=st.text_area("Specific information",key="assistant_call_details")
            if st.button("Generate Phone Script",key="assistant_generate_call"):
                parent_label=gmap.get(gid,"the parent/guardian").split(" (")[0]
                st.session_state["assistant_phone_script"]=f"""Opening:
Hello, may I speak with {parent_label}? This is the teacher calling about {name}. Is now a good time to speak for a few minutes?

Reason:
I wanted to touch base regarding {reason.lower()} in {subj}.

Specific information:
{details or "I wanted to provide you with a clear classroom update."}

Partnership:
I will continue supporting and monitoring {name} at school. I would appreciate your help reinforcing the same expectations and/or skill at home.

Closing:
Thank you for taking the time to speak with me. Please let me know if you have any questions."""
            script=st.text_area("Edit / use during call",value=st.session_state.get("assistant_phone_script",""),height=320)
            outcome=st.text_area("Call notes / parent response",key="assistant_call_outcome")
            if st.button("Save Call to Communication Log",key="assistant_log_call"):
                log_communication(sid,gid,"Phone Call",subj,reason,outcome or details,script)
                st.success("Phone call saved to the selected parent's communication log.")

elif page=="Web & Backup":
    st.markdown('<div class="page-title">Web & Backup</div><div class="page-subtitle">Cloud connection, backups, and migration tools for the browser version of ChapLab.</div>',unsafe_allow_html=True)

    if cloud_configured():
        st.success("☁️ Cloud storage is connected.")
        cfg=cloud_config()
        st.caption(f"Private bucket: {cfg['bucket']} • Database object: {cfg['db_object']}")
        if st.session_state.get("_cloud_sync_error"):
            st.warning("Last sync message: "+st.session_state["_cloud_sync_error"])
    else:
        st.info("ChapLab is currently running in local mode.")
        st.write("To use the same data from your school browser, deploy this Web Edition and add the Supabase settings shown in the deployment guide.")

    st.markdown("### Database Backup")
    db_path=Path(DB)
    if db_path.exists():
        st.download_button(
            "⬇️ Download Current ChapLab Database",
            data=db_path.read_bytes(),
            file_name="teacher_tracker_backup.db",
            mime="application/x-sqlite3",
            use_container_width=True,
            key="download_db_backup_v4"
        )

    if cloud_configured():
        if st.button("☁️ Sync Database Now",use_container_width=True,key="sync_db_now_v4"):
            if cloud_upload_file(DB):
                st.session_state.pop("_cloud_sync_error",None)
                st.success("Database synced to private cloud storage.")
            else:
                st.error("Cloud sync did not complete. Check your Supabase secrets and bucket.")

    st.markdown("### Restore / Migrate Existing ChapLab Data")
    st.caption(
        "Use this once after deploying the web app if you want to move your current personal-computer database into the web version."
    )
    restore=st.file_uploader(
        "Upload an existing teacher_tracker.db or backup .db file",
        type=["db","sqlite","sqlite3"],
        key="restore_chaplab_db_v4"
    )
    confirm_restore=st.checkbox(
        "I understand this will replace the database currently being used by this ChapLab installation.",
        key="confirm_restore_db_v4"
    )
    if st.button(
        "Restore Uploaded Database",
        disabled=(restore is None or not confirm_restore),
        use_container_width=True,
        key="restore_db_button_v4"
    ):
        raw=bytes(restore.getbuffer())
        tmp_path=Path(tempfile.gettempdir())/"chaplab_restore_check.db"
        tmp_path.write_bytes(raw)
        try:
            check=sqlite3.connect(str(tmp_path))
            result=check.execute("PRAGMA integrity_check").fetchone()[0]
            tables={r[0] for r in check.execute("SELECT name FROM sqlite_master WHERE type='table'").fetchall()}
            check.close()
            if str(result).lower()!="ok":
                raise ValueError("SQLite integrity check did not pass.")
            if "scholars" not in tables or "settings" not in tables:
                raise ValueError("This does not appear to be a ChapLab database.")
            Path(DB).write_bytes(raw)
            if cloud_configured() and not cloud_upload_file(DB):
                raise RuntimeError("Database restored locally, but cloud upload failed.")
            st.cache_resource.clear()
            st.success("Database restored. ChapLab will reload with the migrated data.")
            st.rerun()
        except Exception as e:
            st.error(f"Could not restore this database: {e}")

    if auth_config():
        st.markdown("### Session")
        if st.button("Sign Out",key="chaplab_signout_v4"):
            st.session_state["chaplab_authenticated"]=False
            st.rerun()

    st.markdown("### Important")
    st.warning(
        "ChapLab Web is designed for one teacher account at a time. Avoid editing from two devices at the exact same time. "
        "For student information, use only cloud services approved by your school or organization."
    )

elif page=="Communication Log":
    st.markdown('<div class="page-title">Communication Log</div><div class="page-subtitle">Track family contacts and follow-up notes.</div>',unsafe_allow_html=True)
    roster=scholars_df(selected_class or None)

    st.markdown("### Add communication record")
    if roster.empty:
        st.info("No scholars in this class folder.")
    else:
        scholar_id=st.selectbox(
            "Scholar",
            list(roster.id.astype(int)),
            format_func=lambda x:nm(roster[roster.id==x].iloc[0]),
            key="comm_log_scholar"
        )
        gopts=guardian_options_for_scholar(scholar_id)
        guardian_id=st.selectbox("Parent / Guardian",list(gopts),format_func=lambda x:gopts[x],key="comm_log_guardian")
        comm_type=st.selectbox("Communication Type",[
            "Phone Call","Voicemail","Text / School Message","Email",
            "In-Person Conversation","Conference","Letter / Notice Sent","Other"
        ])
        date_text=st.text_input("Date",value=date.today().strftime("%m/%d/%Y"),placeholder="MM/DD/YYYY")
        subject=st.selectbox("Subject",["General","ELA","Math","Science","Social Studies","Behavior","Attendance","Health / Injury"])
        reason=st.selectbox("Reason",[
            "Positive update","Homework reminder","Missing assignment","Academic concern",
            "Behavior concern","Attendance / lateness","Injury / classroom incident",
            "Conference request","Progress update","Assessment / test","Supplies / materials",
            "Follow-up from previous contact","Parent question / concern","Other"
        ])
        notes_text=st.text_area("Notes / More Information",placeholder="Add what happened, what was discussed, parent response, next steps, etc.")
        if st.button("Save Communication Record"):
            c=conn()
            c.execute("""INSERT INTO communications(
                scholar_id,guardian_id,created_at,communication_type,subject,reason,details,generated_text)
                VALUES (?,?,?,?,?,?,?,?)""",
                (int(scholar_id),int(guardian_id) if guardian_id else None,date_text.strip(),
                 comm_type,subject,reason,notes_text.strip(),""))
            c.commit(); c.close()
            st.success("Communication record saved.")
            st.rerun()

    st.markdown("### Saved communication records")
    c=conn()
    q="""SELECT communications.id,communications.created_at,
         scholars.first_name||' '||scholars.last_name scholar,
         COALESCE(TRIM(guardians.first_name||' '||guardians.last_name),'') guardian,
         guardians.relationship,
         communications.communication_type,communications.subject,
         communications.reason,communications.details,communications.generated_text,
         communications.scholar_id, communications.guardian_id
         FROM communications
         LEFT JOIN scholars ON scholars.id=communications.scholar_id
         LEFT JOIN guardians ON guardians.id=communications.guardian_id
         WHERE 1=1"""
    params=[]
    if selected_class:
        q+=" AND scholars.class_id=?"
        params.append(int(selected_class))
    q+=" ORDER BY communications.id DESC"
    d=pd.read_sql_query(q,c,params=params); c.close()

    if d.empty:
        st.caption("No communication records yet.")
    else:
        show=d.copy()
        show["Parent / Guardian"]=show.apply(
            lambda r: (r["guardian"] + (f" ({r['relationship']})" if r["relationship"] else "")).strip(),
            axis=1
        )
        st.dataframe(show[["id","created_at","scholar","Parent / Guardian","communication_type","subject","reason","details"]],
                     hide_index=True,use_container_width=True)

        edit_id=st.selectbox(
            "Edit or delete record",
            list(d.id.astype(int)),
            format_func=lambda x:f"#{x} — {d[d.id==x].iloc[0].scholar} — {d[d.id==x].iloc[0].reason}",
            key="comm_edit_select"
        )
        r=d[d.id==edit_id].iloc[0]
        scholar_choices=scholars_df(selected_class or None)
        sch_ids=list(scholar_choices.id.astype(int))
        current_sch=int(r.scholar_id) if pd.notna(r.scholar_id) and int(r.scholar_id) in sch_ids else sch_ids[0]

        with st.expander("✏️ Edit selected communication"):
            with st.form("edit_comm"):
                e_sch=st.selectbox("Scholar",sch_ids,format_func=lambda x:nm(scholar_choices[scholar_choices.id==x].iloc[0]),
                                   index=sch_ids.index(current_sch))
                egopts=guardian_options_for_scholar(e_sch)
                gids=list(egopts)
                current_gid=int(r.guardian_id) if pd.notna(r.guardian_id) and int(r.guardian_id) in gids else 0
                e_guard=st.selectbox("Parent / Guardian",gids,format_func=lambda x:egopts[x],index=gids.index(current_gid))

                types=["Phone Call","Voicemail","Text / School Message","Email","In-Person Conversation","Conference","Letter / Notice Sent","Other"]
                e_type=st.selectbox("Communication Type",types,index=types.index(r.communication_type) if r.communication_type in types else 0)
                e_date=st.text_input("Date",value=str(r.created_at or ""))
                subjects=["General","ELA","Math","Science","Social Studies","Behavior","Attendance","Health / Injury"]
                e_subject=st.selectbox("Subject",subjects,index=subjects.index(r.subject) if r.subject in subjects else 0)
                reasons=["Positive update","Homework reminder","Missing assignment","Academic concern","Behavior concern","Attendance / lateness",
                         "Injury / classroom incident","Conference request","Progress update","Assessment / test","Supplies / materials",
                         "Follow-up from previous contact","Parent question / concern","Other"]
                e_reason=st.selectbox("Reason",reasons,index=reasons.index(r.reason) if r.reason in reasons else 0)
                e_notes=st.text_area("Notes / More Information",value=str(r.details or ""))

                if st.form_submit_button("Save Changes"):
                    c=conn()
                    c.execute("""UPDATE communications SET scholar_id=?,guardian_id=?,created_at=?,
                                 communication_type=?,subject=?,reason=?,details=? WHERE id=?""",
                              (int(e_sch),int(e_guard) if e_guard else None,e_date.strip(),
                               e_type,e_subject,e_reason,e_notes.strip(),int(edit_id)))
                    c.commit(); c.close()
                    st.success("Communication record updated.")
                    st.rerun()

            delete_ok=st.checkbox("I want to delete this communication record",key="comm_del_check")
            if st.button("Delete Selected Communication",disabled=not delete_ok):
                delete_record("communications",edit_id)
                st.success("Communication record deleted.")
                st.rerun()

