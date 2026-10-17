"""
Resume Extractor Job Portal — Flask conversion of the uploaded Streamlit app.
The resume parsing, keyword/domain matching, sample job dataset and report logic
are reused from the original source. UI is implemented with Flask/Jinja templates.
"""
import os
import re
import io
import sqlite3
import hashlib
from datetime import datetime
from functools import wraps
from flask import Flask, render_template, request, redirect, url_for, session, flash, send_file
from werkzeug.utils import secure_filename

try:
    import PyPDF2
    PDF_SUPPORT = True
except ImportError:
    PDF_SUPPORT = False

try:
    from docx import Document
    DOCX_SUPPORT = True
except ImportError:
    DOCX_SUPPORT = False

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "change-this-secret-before-production")
app.config["MAX_CONTENT_LENGTH"] = 10 * 1024 * 1024
ALLOWED_EXTENSIONS = {"pdf", "docx", "txt"}


DB_PATH = os.environ.get("DATABASE_PATH", os.path.join(os.path.dirname(__file__), "job_portal.db"))

def get_conn():
    conn = sqlite3.connect(DB_PATH, timeout=20)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    with get_conn() as conn:
        conn.execute("""
            CREATE TABLE IF NOT EXISTS users (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                username TEXT UNIQUE NOT NULL,
                email TEXT UNIQUE NOT NULL,
                password TEXT NOT NULL,
                profile_image BLOB,
                preferred_location TEXT DEFAULT 'All Over India',
                suggested_role TEXT DEFAULT '',
                last_login TEXT,
                created_at TEXT DEFAULT CURRENT_TIMESTAMP
            )
        """)
        conn.execute("""
            CREATE TABLE IF NOT EXISTS resumes (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                user_id INTEGER NOT NULL,
                filename TEXT,
                extracted_name TEXT,
                extracted_email TEXT,
                extracted_phone TEXT,
                skills TEXT,
                education TEXT,
                experience TEXT,
                summary TEXT,
                suggested_role TEXT,
                uploaded_at TEXT DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY(user_id) REFERENCES users(id)
            )
        """)

def hash_password(password):
    # Keep compatibility with the uploaded project's existing password hashes.
    return hashlib.sha256(password.encode()).hexdigest()

def create_user(username, email, password):
    try:
        with get_conn() as conn:
            conn.execute(
                "INSERT INTO users (username,email,password) VALUES (?,?,?)",
                (username.strip(), email.strip().lower(), hash_password(password))
            )
        return True, "Account created successfully!"
    except sqlite3.IntegrityError:
        return False, "Username or email already exists."

def verify_user(username, password):
    with get_conn() as conn:
        row = conn.execute(
            "SELECT id, username, email FROM users WHERE username=? AND password=?",
            (username.strip(), hash_password(password))
        ).fetchone()
        if row:
            conn.execute("UPDATE users SET last_login=? WHERE id=?",
                         (datetime.now().strftime("%Y-%m-%d %H:%M:%S"), row["id"]))
        return dict(row) if row else None

def get_user_profile(user_id):
    with get_conn() as conn:
        row = conn.execute("""SELECT username,email,profile_image,preferred_location,
            suggested_role,last_login,created_at FROM users WHERE id=?""", (user_id,)).fetchone()
        return dict(row) if row else {}

def update_profile(user_id, location, role, image_bytes=None):
    with get_conn() as conn:
        conn.execute("UPDATE users SET preferred_location=?, suggested_role=? WHERE id=?",
                     (location, role, user_id))
        if image_bytes is not None:
            conn.execute("UPDATE users SET profile_image=? WHERE id=?", (image_bytes, user_id))

def save_resume(user_id, filename, info):
    with get_conn() as conn:
        conn.execute("""INSERT INTO resumes
            (user_id,filename,extracted_name,extracted_email,extracted_phone,skills,
             education,experience,summary,suggested_role)
            VALUES (?,?,?,?,?,?,?,?,?,?)""",
            (user_id, filename, info.get("name"), info.get("email"), info.get("phone"),
             ", ".join(info.get("skills", [])), info.get("education"), info.get("experience"),
             info.get("summary"), info.get("suggested_role")))

def get_user_resumes(user_id):
    with get_conn() as conn:
        return conn.execute("SELECT * FROM resumes WHERE user_id=? ORDER BY uploaded_at DESC",
                            (user_id,)).fetchall()

def delete_resume(resume_id, user_id):
    with get_conn() as conn:
        conn.execute("DELETE FROM resumes WHERE id=? AND user_id=?", (resume_id, user_id))


LOCATIONS = [
    "Pune","Mumbai","Bangalore","Hyderabad","Delhi","Chennai",
    "Noida","Gurgaon","Kolkata","Ahmedabad","Remote","All Over India",
]
NAUKRI_LOC = {
    "Pune":"pune","Mumbai":"mumbai","Bangalore":"bengaluru",
    "Hyderabad":"hyderabad","Delhi":"delhi","Chennai":"chennai",
    "Noida":"noida","Gurgaon":"gurugram","Kolkata":"kolkata",
    "Ahmedabad":"ahmedabad","Remote":"work-from-home","All Over India":"india",
}
LINKEDIN_GEO = {
    "Pune":"102713980","Mumbai":"102717119","Bangalore":"105214831",
    "Hyderabad":"102691710","Delhi":"102257491","Chennai":"102741530",
    "Noida":"102713980","Gurgaon":"102713980","Kolkata":"102744304",
    "Ahmedabad":"102633910","Remote":"","All Over India":"102713980",
}


# ══════════════════════════════════════════════════════════════
#  SKILL KEYWORDS  — 300+ terms, 23 domains
# ══════════════════════════════════════════════════════════════
SKILL_KEYWORDS = [
    # DevOps / Cloud
    "aws","ec2","s3","iam","vpc","rds","lambda","cloudformation","eks","ecs","ecr",
    "route53","cloudwatch","sqs","sns","auto scaling",
    "azure","azure devops","azure pipelines","aks","arm templates","azure monitor",
    "gcp","google cloud","gke","bigquery","cloud run","cloud functions",
    "docker","kubernetes","k8s","helm","containerd","podman","docker compose",
    "jenkins","gitlab ci","github actions","circleci","argocd","teamcity",
    "terraform","ansible","puppet","chef","vagrant","packer","pulumi",
    "linux","unix","bash","shell scripting","nginx","apache",
    "git","github","gitlab","bitbucket","ci/cd","devops","devsecops","sre",
    "prometheus","grafana","datadog","elk","elasticsearch","logstash","kibana",
    "sonarqube","nexus","artifactory","trivy","istio","consul","vault",
    "microservices","serverless","infrastructure as code",
    # Full Stack / Frontend
    "html","html5","css","css3","javascript","es6","typescript",
    "react","angular","vue","next.js","nuxt.js","svelte","gatsby",
    "redux","mobx","graphql","apollo","webpack","vite","jest","cypress",
    "bootstrap","tailwind","material ui","ant design",
    "node.js","express","fastify","nestjs","mongodb","mongoose","firebase",
    "supabase","rest api","websocket",
    # Python
    "python","django","flask","fastapi","celery","sqlalchemy","pydantic",
    "pytest","asyncio","uvicorn","gunicorn",
    # Java
    "java","spring","spring boot","spring mvc","spring security","hibernate",
    "jpa","maven","gradle","junit","mockito","kafka","rabbitmq",
    "microservices","j2ee","jvm","kotlin",
    # Other Backend
    "php","laravel","ruby","rails","go","golang","rust","c#","asp.net",
    "dotnet","c++","c","scala",
    # Databases
    "sql","mysql","postgresql","sqlite","oracle","sql server","mongodb",
    "redis","cassandra","dynamodb","neo4j","mariadb",
    "snowflake","databricks","hive","hbase","redshift",
    # Mobile
    "android","kotlin android","android studio","xml layouts","retrofit",
    "ios","swift","swiftui","xcode","objective-c",
    "flutter","dart","react native","xamarin","ionic",
    "firebase","google maps api","gradle","cocoapods",
    # Data Science / AI / ML
    "machine learning","deep learning","artificial intelligence","ai","ml",
    "nlp","natural language processing","computer vision","generative ai",
    "llm","large language models","gpt","bert","transformers","hugging face",
    "tensorflow","pytorch","keras","scikit-learn","xgboost","lightgbm",
    "pandas","numpy","matplotlib","seaborn","plotly","scipy","statsmodels",
    "spark","hadoop","airflow","dbt","mlflow","kubeflow",
    "data analysis","data science","data engineering","etl","feature engineering",
    "tableau","power bi","looker","qlik","data studio",
    "statistics","probability","a/b testing","regression","classification",
    "clustering","neural networks","cnn","rnn","lstm","time series",
    "r","matlab","sas","spss","excel",
    # Cyber Security
    "cybersecurity","cyber security","information security","infosec",
    "ethical hacking","penetration testing","pentest","red team","blue team",
    "vapt","vulnerability assessment","owasp","exploit development",
    "soc","siem","splunk","qradar","arcsight",
    "nessus","burp suite","metasploit","nmap","wireshark","kali linux",
    "firewalls","ids","ips","waf","dlp","endpoint security","edr","xdr",
    "iso 27001","nist","pci dss","gdpr","incident response",
    "digital forensics","malware analysis","threat hunting","threat intelligence",
    "network security","zero trust",
    # Networking
    "networking","tcp/ip","udp","dns","dhcp","http","https",
    "routing","switching","vlan","stp","bgp","ospf","eigrp","mpls",
    "vpn","ipsec","ssl","tls","load balancing","nat",
    "cisco","juniper","aruba","fortinet","palo alto","ccna","ccnp","ccie",
    "network monitoring","snmp","sdwan","sdn",
    # HR
    "hr","human resources","recruitment","talent acquisition","sourcing",
    "interviewing","onboarding","offboarding","payroll","salary processing",
    "employee relations","performance management","appraisal","kpi","okr",
    "hris","hrms","workday","sap hr","peoplesoft","zoho people",
    "labour law","statutory compliance","pf","esi",
    "training and development","learning development",
    "organizational development","change management","employer branding",
    # Finance / Accounting
    "finance","accounting","bookkeeping","financial reporting","financial analysis",
    "tally","tally erp","quickbooks","sap fi","sap co","sap fico","oracle financials",
    "gst","taxation","income tax","tds","itr","vat",
    "auditing","internal audit","statutory audit","ifrs","ind as","gaap","us gaap",
    "budgeting","forecasting","cost accounting","management accounting",
    "accounts payable","accounts receivable","bank reconciliation",
    "financial modelling","valuation","dcf","investment banking",
    "equity research","portfolio management","mutual funds","derivatives",
    "advanced excel","vlookup","pivot tables","erp","sap",
    # Mechanical Engineering
    "autocad","solidworks","catia","catia v5","ansys","creo","nx",
    "unigraphics","solidedge","hypermesh",
    "mechanical design","product design","industrial design","dfma",
    "cad","cam","cnc","fea","cfd","thermal analysis","gd&t",
    "manufacturing","production planning","quality control","six sigma","lean",
    "iso 9001","kaizen","5s","jit","tpm",
    "hydraulics","pneumatics","robotics",
    # Civil Engineering
    "civil engineering","structural engineering","structural analysis","structural design",
    "staad pro","etabs","safe","sap2000","revit structure",
    "primavera","ms project","construction management","project planning",
    "quantity surveying","estimation","billing",
    "rcc design","steel design","foundation design","geotechnical","soil mechanics",
    "surveying","total station","levelling",
    "highway engineering","transportation","drainage","water supply","mep",
    # Electrical Engineering
    "plc","scada","dcs","hmi","automation","industrial automation",
    "embedded systems","embedded c","arm cortex","raspberry pi","arduino",
    "simulink","labview","psim","etap",
    "electrical design","eplan","autocad electrical","panel design",
    "drives","vfd","servo drives","motion control",
    "power systems","switchgear","protection relay","substation","transformer",
    "electrical installation","testing commissioning","iot",
    # Interior Design
    "interior design","interior decoration","space planning",
    "sketchup","3ds max","v-ray","lumion","enscape","d5 render",
    "revit interior","archicad","chief architect",
    "furniture design","material selection","colour theory","lighting design",
    "bim","residential design","commercial design",
    # Graphic Design
    "graphic design","visual design","branding","brand identity","logo design",
    "typography","layout design","print design","packaging design",
    "photoshop","illustrator","indesign","after effects","premiere pro",
    "canva","figma graphic","corel draw",
    "motion graphics","animation","video editing",
    # UI/UX Design
    "ui design","ux design","ui/ux","user experience","user interface",
    "figma","adobe xd","sketch","invision","zeplin","principle","framer",
    "wireframing","wireframes","prototyping","mockup","user research",
    "usability testing","information architecture",
    "design system","component library","accessibility","wcag","hci",
    "interaction design","visual hierarchy","responsive design",
    # Digital Marketing
    "digital marketing","performance marketing","growth marketing",
    "seo","search engine optimization","sem","search engine marketing",
    "google ads","google adwords","ppc","display advertising",
    "facebook ads","instagram ads","linkedin ads","youtube ads","programmatic",
    "social media marketing","social media management","community management",
    "content marketing","content strategy","content writing","copywriting",
    "email marketing","marketing automation","hubspot","mailchimp","klaviyo",
    "google analytics","ga4","adobe analytics","mixpanel","amplitude",
    "crm","salesforce","zoho crm","affiliate marketing","influencer marketing",
    "conversion rate optimization","cro","a/b testing marketing",
    # Sales
    "sales","business development","lead generation","prospecting",
    "cold calling","inside sales","field sales","b2b sales","b2c sales",
    "account management","key account","crm sales","pipeline management",
    "negotiation","objection handling","revenue generation",
    "saas sales","enterprise sales","solution selling","consultative selling",
    # Content Writing
    "content writing","blog writing","article writing","technical writing",
    "copywriting","proofreading","editing","ghostwriting","seo writing",
    "research","creative writing","script writing",
    # Customer Support
    "customer support","customer service","customer success","help desk",
    "technical support","it support","service desk",
    "zendesk","freshdesk","jira service","servicenow","intercom","chat support",
    "ticketing","sla","escalation","troubleshooting",
    # Healthcare
    "patient care","clinical care","bedside manner","icu","critical care",
    "nursing","staff nurse","clinical nursing","operation theatre","ot",
    "pharmacology","drug dispensing","pharmacy","medical billing","medical coding",
    "icd","cpt codes","health informatics","ehr","emr",
    "radiology","x-ray","mri","ct scan","lab technician",
    # Teaching / Education
    "teaching","classroom management","lesson planning","curriculum planning",
    "curriculum development","instructional design","e-learning","lms",
    "moodle","canvas lms","google classroom","online teaching",
    "assessment","evaluation","student counselling","mentoring","coaching",
    "cbse","icse","edtech","higher education","academic research",
    # Project / Agile
    "agile","scrum","kanban","safe","sprint planning",
    "pmp","prince2","project management","program management","pmo",
    "jira","confluence","trello","asana","ms project",
    # General / Soft skills
    "communication","leadership","teamwork","problem solving","time management",
    "ms office","microsoft office","word","powerpoint","presentation",
]


# ══════════════════════════════════════════════════════════════
#  ROLE CLUSTERS  — ordered most-specific to least-specific
#  DevOps checked BEFORE Security — Security requires explicit must_have
# ══════════════════════════════════════════════════════════════
ROLE_CLUSTERS = [
    # DevOps / Cloud
    {"role":"DevOps Engineer","domain":"devops",
     "must_have":["docker","kubernetes"],
     "keywords":["jenkins","ci/cd","terraform","ansible","linux","aws","azure","helm","git"],
     "min_score":3},
    {"role":"AWS Cloud Engineer","domain":"devops",
     "must_have":["aws","ec2"],
     "keywords":["s3","iam","vpc","lambda","cloudformation","eks","rds","terraform","docker"],
     "min_score":2},
    {"role":"Azure DevOps Engineer","domain":"devops",
     "must_have":["azure","azure devops"],
     "keywords":["aks","terraform","docker","kubernetes","ci/cd","git","azure pipelines"],
     "min_score":2},
    {"role":"GCP Cloud Engineer","domain":"devops",
     "must_have":["gcp","gke"],
     "keywords":["bigquery","cloud run","terraform","docker","kubernetes","python"],
     "min_score":2},
    {"role":"Site Reliability Engineer","domain":"devops",
     "must_have":["kubernetes","prometheus"],
     "keywords":["grafana","linux","docker","ci/cd","terraform","aws","sre"],
     "min_score":2},
    {"role":"Platform Engineer","domain":"devops",
     "must_have":["kubernetes","terraform"],
     "keywords":["helm","docker","ci/cd","aws","linux","argocd","istio"],
     "min_score":2},
    {"role":"DevOps / Cloud Engineer","domain":"devops",
     "must_have":["docker","terraform"],
     "keywords":["linux","aws","azure","git","ci/cd","jenkins","kubernetes","bash"],
     "min_score":2},
    {"role":"Cloud Infrastructure Engineer","domain":"devops",
     "must_have":["aws"],
     "keywords":["terraform","docker","linux","kubernetes","ci/cd","ansible","git"],
     "min_score":2},
    # Cyber Security — requires EXPLICIT security keywords
    {"role":"Ethical Hacker / Pen Tester","domain":"cybersecurity",
     "must_have":["ethical hacking","penetration testing"],
     "keywords":["vapt","owasp","metasploit","burp suite","kali linux","nmap"],
     "min_score":2},
    {"role":"SOC Analyst","domain":"cybersecurity",
     "must_have":["soc","siem"],
     "keywords":["splunk","qradar","cybersecurity","firewalls","ids","ips","incident response"],
     "min_score":2},
    {"role":"Cyber Security Analyst","domain":"cybersecurity",
     "must_have":["cybersecurity","vapt"],
     "keywords":["ethical hacking","siem","splunk","burp suite","firewalls","owasp"],
     "min_score":2},
    {"role":"Network Security Engineer","domain":"cybersecurity",
     "must_have":["network security","firewalls"],
     "keywords":["vpn","ids","ips","siem","palo alto","cisco","cybersecurity","zero trust"],
     "min_score":2},
    # AI / ML / Data
    {"role":"AI / ML Engineer","domain":"data_ai",
     "must_have":["machine learning","deep learning"],
     "keywords":["tensorflow","pytorch","python","nlp","computer vision","scikit-learn","transformers"],
     "min_score":3},
    {"role":"Data Scientist","domain":"data_ai",
     "must_have":["data science","machine learning"],
     "keywords":["python","pandas","numpy","scikit-learn","tableau","sql","statistics"],
     "min_score":2},
    {"role":"NLP / GenAI Engineer","domain":"data_ai",
     "must_have":["nlp","transformers"],
     "keywords":["bert","gpt","llm","hugging face","pytorch","python","machine learning"],
     "min_score":2},
    {"role":"Data Analyst","domain":"data_ai",
     "must_have":["data analysis"],
     "keywords":["sql","excel","power bi","tableau","python","google analytics","statistics"],
     "min_score":1},
    {"role":"Data Engineer","domain":"data_ai",
     "must_have":["etl","airflow"],
     "keywords":["spark","sql","python","hadoop","databricks","snowflake","kafka","dbt"],
     "min_score":1},
    # Full Stack / Frontend
    {"role":"Full Stack Developer","domain":"fullstack",
     "must_have":["react","node.js"],
     "keywords":["javascript","typescript","html","css","mongodb","express","sql","docker","rest api"],
     "min_score":2},
    {"role":"React Developer","domain":"fullstack",
     "must_have":["react"],
     "keywords":["javascript","typescript","html","css","tailwind","next.js","redux","graphql"],
     "min_score":1},
    {"role":"Angular Developer","domain":"fullstack",
     "must_have":["angular"],
     "keywords":["typescript","javascript","html","css","rest api"],
     "min_score":1},
    {"role":"Frontend Developer","domain":"fullstack",
     "must_have":["html","css","javascript"],
     "keywords":["react","angular","vue","bootstrap","tailwind","typescript"],
     "min_score":1},
    {"role":"Backend Developer (Node.js)","domain":"fullstack",
     "must_have":["node.js","express"],
     "keywords":["javascript","mongodb","rest api","redis","postgresql","docker","kafka"],
     "min_score":1},
    # Python
    {"role":"Python Developer","domain":"python",
     "must_have":["python"],
     "keywords":["django","flask","fastapi","postgresql","redis","rest api","docker","celery","asyncio"],
     "min_score":1},
    # Java
    {"role":"Java Developer","domain":"java",
     "must_have":["java"],
     "keywords":["spring","spring boot","microservices","sql","hibernate","kafka","maven","gradle"],
     "min_score":1},
    # Mobile
    {"role":"Android Developer","domain":"mobile",
     "must_have":["android"],
     "keywords":["kotlin android","java","firebase","retrofit","rest api","gradle","android studio"],
     "min_score":1},
    {"role":"Flutter Developer","domain":"mobile",
     "must_have":["flutter","dart"],
     "keywords":["firebase","android","ios","rest api","bloc","getx"],
     "min_score":1},
    {"role":"iOS Developer","domain":"mobile",
     "must_have":["ios","swift"],
     "keywords":["swiftui","xcode","objective-c","rest api","firebase","cocoapods"],
     "min_score":1},
    {"role":"React Native Developer","domain":"mobile",
     "must_have":["react native"],
     "keywords":["javascript","android","ios","rest api","firebase","redux","expo"],
     "min_score":1},
    # Networking
    {"role":"Network Engineer","domain":"networking",
     "must_have":["networking","ccna"],
     "keywords":["cisco","routing","switching","tcp/ip","dns","dhcp","vpn","bgp","ospf","vlan"],
     "min_score":2},
    {"role":"System Administrator","domain":"networking",
     "must_have":["linux","networking"],
     "keywords":["dns","dhcp","vpn","firewalls","tcp/ip","bash"],
     "min_score":1},
    # Marketing
    {"role":"Digital Marketing Specialist","domain":"marketing",
     "must_have":["digital marketing","seo"],
     "keywords":["google analytics","social media","content writing","google ads","sem","ppc"],
     "min_score":2},
    {"role":"Performance Marketing Executive","domain":"marketing",
     "must_have":["google ads","facebook ads"],
     "keywords":["sem","ppc","google analytics","cro","instagram ads"],
     "min_score":1},
    {"role":"SEO Specialist","domain":"marketing",
     "must_have":["seo"],
     "keywords":["google analytics","content writing","sem","content marketing"],
     "min_score":1},
    # HR
    {"role":"HR Professional","domain":"hr",
     "must_have":["hr","recruitment"],
     "keywords":["talent acquisition","payroll","employee relations","hris","onboarding","performance management"],
     "min_score":1},
    # Finance
    {"role":"Finance / Accounts Executive","domain":"finance",
     "must_have":["finance","accounting"],
     "keywords":["tally","gst","taxation","excel","financial analysis","budgeting","accounts payable"],
     "min_score":1},
    {"role":"Financial Analyst","domain":"finance",
     "must_have":["financial analysis","financial modelling"],
     "keywords":["excel","sql","python","valuation","dcf","ifrs","gaap","budgeting","forecasting"],
     "min_score":1},
    # Mechanical
    {"role":"Mechanical Design Engineer","domain":"mechanical",
     "must_have":["autocad","solidworks"],
     "keywords":["catia","ansys","creo","mechanical design","nx","fea","gd&t"],
     "min_score":1},
    {"role":"Manufacturing Engineer","domain":"mechanical",
     "must_have":["manufacturing","cnc"],
     "keywords":["autocad","solidworks","six sigma","lean","quality control","gd&t","cam"],
     "min_score":1},
    # Civil
    {"role":"Civil / Structural Engineer","domain":"civil",
     "must_have":["civil engineering","structural analysis"],
     "keywords":["autocad","staad pro","etabs","structural design","rcc design"],
     "min_score":1},
    # Electrical
    {"role":"Electrical / Automation Engineer","domain":"electrical",
     "must_have":["plc","scada"],
     "keywords":["automation","matlab","embedded systems","hmi","drives","dcs","labview"],
     "min_score":1},
    # Interior Design
    {"role":"Interior Designer","domain":"interior",
     "must_have":["sketchup","3ds max"],
     "keywords":["autocad","v-ray","revit interior","lumion","interior design","space planning"],
     "min_score":1},
    # UI/UX
    {"role":"UI/UX Designer","domain":"uiux",
     "must_have":["figma","wireframing"],
     "keywords":["adobe xd","prototyping","user research","design system","sketch","usability testing"],
     "min_score":1},
    # Graphic Design
    {"role":"Graphic Designer","domain":"design",
     "must_have":["photoshop","illustrator"],
     "keywords":["canva","figma graphic","indesign","branding","after effects","typography"],
     "min_score":1},
    # Sales
    {"role":"Sales / Business Development Executive","domain":"sales",
     "must_have":["sales","lead generation"],
     "keywords":["crm","negotiation","business development","b2b sales","account management","pipeline management"],
     "min_score":1},
    # Content
    {"role":"Content Writer","domain":"content",
     "must_have":["content writing","blog writing"],
     "keywords":["seo writing","copywriting","research","proofreading","creative writing"],
     "min_score":1},
    # Support
    {"role":"Customer Support Executive","domain":"support",
     "must_have":["customer support","customer service"],
     "keywords":["technical support","zendesk","freshdesk","crm sales","sla","ticketing","troubleshooting"],
     "min_score":1},
    # Healthcare
    {"role":"Healthcare Professional","domain":"healthcare",
     "must_have":["patient care"],
     "keywords":["nursing","icu","pharmacology","medical billing","clinical care","ehr"],
     "min_score":1},
    # Education
    {"role":"Teacher / Educator","domain":"education",
     "must_have":["teaching","curriculum planning"],
     "keywords":["lesson planning","e-learning","assessment","lms","classroom management","mentoring"],
     "min_score":1},
    # IT General (catch-all — evaluated last)
    {"role":"Software / IT Professional","domain":"it_general",
     "must_have":[],
     "keywords":["python","java","javascript","sql","linux","docker","git","api","software"],
     "min_score":2},
]

IT_DOMAINS = {
    "devops","cybersecurity","data_ai","fullstack","python","java",
    "mobile","networking","uiux","design","content","support","it_general"
}

def suggest_role(skills: list) -> tuple:
    skill_set = {s.lower() for s in skills}
    for cl in ROLE_CLUSTERS:
        must = cl.get("must_have", [])
        if must and not all(m in skill_set for m in must):
            continue
        score = sum(1 for kw in cl["keywords"] if kw in skill_set)
        if score >= cl["min_score"]:
            return cl["role"], cl["domain"]
    return "Software / IT Professional", "it_general"


# ══════════════════════════════════════════════════════════════
#  APPLY LINK GENERATORS  — proper SEO slugs
# ══════════════════════════════════════════════════════════════
def _slug(text: str) -> str:
    return re.sub(r'[^a-z0-9]+', '-', text.lower()).strip('-')

def _naukri_url(title: str, location: str) -> str:
    role_slug = _slug(title)
    loc_slug  = NAUKRI_LOC.get(location, _slug(location))
    if location == "Remote":
        return f"https://www.naukri.com/{role_slug}-jobs?workType=work+from+home"
    return f"https://www.naukri.com/{role_slug}-jobs-in-{loc_slug}"

def _linkedin_url(title: str, location: str) -> str:
    q   = title.replace(" ", "%20")
    geo = LINKEDIN_GEO.get(location, "102713980")
    if location in ("Remote","All Over India"):
        return f"https://www.linkedin.com/jobs/search/?keywords={q}&f_WT=2"
    return f"https://www.linkedin.com/jobs/search/?keywords={q}&geoId={geo}"

def _indeed_url(title: str, location: str) -> str:
    q   = title.replace(" ", "+")
    loc = "Work+from+Home" if location == "Remote" else location.replace(" ", "+")
    return f"https://in.indeed.com/jobs?q={q}&l={loc}"

def _foundit_url(title: str, location: str) -> str:
    q   = title.replace(" ", "-").lower()
    loc = location.lower().replace(" ", "-")
    return f"https://www.foundit.in/srp/results?query={q}&location={loc}"


# ══════════════════════════════════════════════════════════════
#  JOB DATASET  — 130+ entries across 23 domains
# ══════════════════════════════════════════════════════════════
def _j(title, company, domain, level, location, skills):
    return {
        "title":title,"company":company,"domain":domain,
        "level":level,"location":location,"skills":skills,
        "apply_naukri":  _naukri_url(title, location),
        "apply_linkedin":_linkedin_url(title, location),
        "apply_indeed":  _indeed_url(title, location),
        "apply_foundit": _foundit_url(title, location),
    }

JOB_DATA = [
    # ── 1. DEVOPS / CLOUD
    _j("DevOps Intern","Infosys","devops","Internship","Pune",
       ["linux","git","docker","ci/cd","bash","jenkins"]),
    _j("DevOps Engineer – Fresher","Tech Mahindra","devops","Fresher","Pune",
       ["linux","git","docker","bash","jenkins","ci/cd","ansible"]),
    _j("Junior DevOps Engineer","Wipro","devops","Junior","Bangalore",
       ["docker","jenkins","linux","git","ci/cd","ansible","terraform"]),
    _j("AWS Cloud Engineer","TCS","devops","Mid-Level","Mumbai",
       ["aws","ec2","s3","iam","vpc","lambda","cloudformation","docker","terraform","rds"]),
    _j("Kubernetes Engineer","HCL Technologies","devops","Mid-Level","Hyderabad",
       ["kubernetes","helm","docker","prometheus","grafana","linux","ci/cd","terraform"]),
    _j("Azure DevOps Engineer","Accenture","devops","Mid-Level","Noida",
       ["azure","azure devops","docker","kubernetes","terraform","git","ci/cd","aks"]),
    _j("GCP Cloud Engineer","Thoughtworks","devops","Mid-Level","Remote",
       ["gcp","gke","bigquery","cloud run","docker","kubernetes","terraform"]),
    _j("Cloud Infrastructure Engineer","Capgemini","devops","Mid-Level","Gurgaon",
       ["aws","azure","docker","kubernetes","terraform","linux","jenkins","ci/cd"]),
    _j("Platform Engineer","Flipkart","devops","Senior","Bangalore",
       ["kubernetes","helm","terraform","aws","ci/cd","python","docker","linux","argocd"]),
    _j("Site Reliability Engineer","Google India","devops","Senior","Hyderabad",
       ["kubernetes","prometheus","grafana","linux","python","ci/cd","aws","sre"]),
    _j("Senior DevOps Engineer","Amazon India","devops","Senior","Bangalore",
       ["aws","docker","kubernetes","terraform","ci/cd","linux","python","ansible","helm"]),
    _j("DevOps Engineer","Microsoft India","devops","Mid-Level","Hyderabad",
       ["azure","docker","kubernetes","terraform","ci/cd","linux","git","azure devops"]),
    _j("Cloud DevOps Engineer","Infosys","devops","Junior","Pune",
       ["aws","docker","linux","git","ci/cd","jenkins","terraform","bash"]),
    # ── 2. FULL STACK
    _j("Full Stack Intern","Freshworks","fullstack","Internship","Chennai",
       ["html","css","javascript","react","node.js","rest api"]),
    _j("Full Stack Fresher","Cognizant","fullstack","Fresher","Mumbai",
       ["html","css","javascript","react","node.js","sql","git"]),
    _j("Frontend Developer","Zoho Corp","fullstack","Junior","Chennai",
       ["html","css","javascript","react","bootstrap","git","typescript"]),
    _j("MERN Stack Developer","Razorpay","fullstack","Junior","Bangalore",
       ["react","node.js","mongodb","express","javascript","html","css","rest api"]),
    _j("React Developer","Swiggy","fullstack","Mid-Level","Bangalore",
       ["react","javascript","typescript","html","css","redux","tailwind","graphql"]),
    _j("Angular Developer","Mphasis","fullstack","Mid-Level","Pune",
       ["angular","typescript","javascript","html","css","rest api","git"]),
    _j("Backend Developer (Node.js)","Paytm","fullstack","Mid-Level","Noida",
       ["node.js","javascript","mongodb","redis","rest api","docker","kafka","express"]),
    _j("Senior Full Stack Engineer","Meesho","fullstack","Senior","Bangalore",
       ["react","node.js","javascript","postgresql","docker","aws","rest api","typescript"]),
    _j("Vue.js Developer","Zoho","fullstack","Junior","Chennai",
       ["vue","javascript","html","css","node.js","rest api","typescript"]),
    # ── 3. PYTHON
    _j("Python Intern","Jio Platforms","python","Internship","Mumbai",
       ["python","html","sql","git","flask"]),
    _j("Python Developer – Fresher","Wipro","python","Fresher","Hyderabad",
       ["python","django","sql","git","rest api","postgresql"]),
    _j("Python Developer","Freshworks","python","Junior","Chennai",
       ["python","django","rest api","postgresql","git","celery"]),
    _j("Python Backend Developer","Infosys","python","Mid-Level","Pune",
       ["python","flask","fastapi","postgresql","redis","rest api","docker","sqlalchemy"]),
    _j("Senior Python Engineer","PhonePe","python","Senior","Bangalore",
       ["python","fastapi","celery","redis","postgresql","docker","kubernetes","aws","asyncio"]),
    # ── 4. JAVA
    _j("Java Intern","TCS","java","Internship","Mumbai",
       ["java","html","sql","git","spring"]),
    _j("Java Developer – Fresher","Infosys","java","Fresher","Delhi",
       ["java","spring","sql","git","junit","maven"]),
    _j("Java Developer","Cognizant","java","Junior","Pune",
       ["java","spring","sql","rest api","git","junit","hibernate"]),
    _j("Spring Boot Developer","HCL","java","Mid-Level","Bangalore",
       ["java","spring boot","microservices","sql","hibernate","kafka","rest api","docker"]),
    _j("Senior Java Engineer","Amazon India","java","Senior","Bangalore",
       ["java","spring boot","microservices","kafka","aws","docker","kubernetes","maven"]),
    # ── 5. DATA SCIENCE / AI / ML
    _j("Data Science Intern","KPMG India","data_ai","Internship","Mumbai",
       ["python","pandas","numpy","sql","excel","statistics","matplotlib"]),
    _j("Data Analyst – Fresher","EY India","data_ai","Fresher","Pune",
       ["sql","excel","python","statistics","data analysis","tableau"]),
    _j("Data Analyst","Deloitte India","data_ai","Junior","Gurgaon",
       ["sql","excel","power bi","tableau","python","google analytics","statistics"]),
    _j("Data Scientist","Jio Platforms","data_ai","Mid-Level","Mumbai",
       ["python","machine learning","scikit-learn","pandas","numpy","sql","tableau","statistics"]),
    _j("ML Engineer","Samsung R&D","data_ai","Mid-Level","Bangalore",
       ["deep learning","computer vision","pytorch","python","nlp","machine learning","keras"]),
    _j("Business Intelligence Developer","MuSigma","data_ai","Mid-Level","Bangalore",
       ["power bi","tableau","sql","python","data analysis","excel","looker"]),
    _j("Data Engineer","Dunzo","data_ai","Mid-Level","Bangalore",
       ["python","spark","airflow","sql","aws","snowflake","etl","kafka","dbt"]),
    _j("Remote ML Engineer","Tiger Analytics","data_ai","Mid-Level","Remote",
       ["python","machine learning","deep learning","tensorflow","sql","docker","mlflow"]),
    _j("AI Engineer","Accenture AI","data_ai","Senior","Hyderabad",
       ["machine learning","deep learning","tensorflow","python","nlp","docker","aws","mlflow"]),
    _j("NLP / GenAI Engineer","Haptik (Reliance)","data_ai","Senior","Mumbai",
       ["nlp","python","pytorch","machine learning","transformers","llm","hugging face","docker"]),
    # ── 6. CYBER SECURITY
    _j("Cyber Security Intern","Tech Mahindra","cybersecurity","Internship","Pune",
       ["linux","networking","cybersecurity","wireshark","kali linux","nmap"]),
    _j("Security Analyst – Fresher","Tata Consultancy","cybersecurity","Fresher","Mumbai",
       ["cybersecurity","linux","networking","vapt","wireshark","owasp"]),
    _j("SOC Analyst L1","IBM Security","cybersecurity","Junior","Bangalore",
       ["soc","siem","splunk","linux","firewalls","ids","ips","incident response"]),
    _j("Ethical Hacker / Pen Tester","Wipro CyberDefense","cybersecurity","Mid-Level","Noida",
       ["ethical hacking","penetration testing","owasp","metasploit","burp suite","vapt","nmap","kali linux"]),
    _j("Cybersecurity Analyst","Accenture Security","cybersecurity","Mid-Level","Pune",
       ["cybersecurity","vapt","siem","splunk","firewalls","linux","owasp","iso 27001"]),
    _j("SOC Analyst L2","Infosys Security","cybersecurity","Senior","Hyderabad",
       ["soc","siem","splunk","cybersecurity","ids","ips","firewalls","incident response","threat hunting"]),
    # ── 7. NETWORKING
    _j("Network Intern","HCL Technologies","networking","Internship","Delhi",
       ["networking","tcp/ip","dns","linux","wireshark","ccna"]),
    _j("Network Admin – Fresher","Wipro","networking","Fresher","Hyderabad",
       ["networking","tcp/ip","dns","dhcp","linux","ccna","routing","switching"]),
    _j("Network Engineer","Cisco India","networking","Junior","Bangalore",
       ["networking","routing","switching","tcp/ip","dns","dhcp","cisco","ccna","vlan"]),
    _j("System Administrator","Tech Mahindra","networking","Mid-Level","Pune",
       ["linux","networking","vpn","dns","dhcp","firewalls","tcp/ip","bash"]),
    _j("Senior Network Engineer","Tata Communications","networking","Senior","Mumbai",
       ["networking","routing","switching","cisco","bgp","ospf","vpn","firewalls","mpls"]),
    # ── 8. MOBILE
    _j("Android Intern","Dream11","mobile","Internship","Mumbai",
       ["android","java","kotlin android","xml layouts","git","android studio"]),
    _j("Flutter Developer","Byju's","mobile","Junior","Bangalore",
       ["flutter","dart","android","ios","firebase","rest api","bloc"]),
    _j("Android Developer","Ola Cabs","mobile","Mid-Level","Bangalore",
       ["android","kotlin android","java","firebase","retrofit","rest api","gradle"]),
    _j("iOS Developer","PhonePe","mobile","Mid-Level","Pune",
       ["swift","ios","swiftui","xcode","rest api","firebase","cocoapods"]),
    _j("React Native Developer","Zomato","mobile","Mid-Level","Gurgaon",
       ["react native","javascript","android","ios","rest api","firebase","redux"]),
    _j("Mobile App Developer","Flipkart","mobile","Senior","Bangalore",
       ["android","kotlin android","flutter","firebase","rest api","ci/cd","gradle"]),
    # ── 9. DIGITAL MARKETING
    _j("Marketing Intern","Dentsu India","marketing","Internship","Mumbai",
       ["social media marketing","content writing","canva","google analytics","digital marketing"]),
    _j("Marketing – Fresher","MakeMyTrip","marketing","Fresher","Gurgaon",
       ["digital marketing","social media marketing","content writing","google analytics","seo"]),
    _j("SEO Specialist","HT Media","marketing","Junior","Delhi",
       ["seo","google analytics","content writing","digital marketing","sem"]),
    _j("Content Marketing Specialist","Byju's","marketing","Junior","Bangalore",
       ["content writing","seo writing","social media marketing","digital marketing","copywriting"]),
    _j("Digital Marketing Executive","Dentsu India","marketing","Mid-Level","Mumbai",
       ["digital marketing","seo","google analytics","social media marketing","content writing","sem","ppc"]),
    _j("Performance Marketing Executive","Nykaa","marketing","Mid-Level","Mumbai",
       ["google ads","facebook ads","sem","google analytics","cro","crm"]),
    _j("Growth Marketing Manager","Meesho","marketing","Senior","Bangalore",
       ["digital marketing","social media marketing","google analytics","copywriting","seo","crm","a/b testing marketing"]),
    # ── 10. HR
    _j("HR Intern","Aon Hewitt India","hr","Internship","Pune",
       ["hr","communication","ms office","recruitment","onboarding"]),
    _j("HR – Fresher","TCS","hr","Fresher","Mumbai",
       ["hr","recruitment","communication","ms office","hrms","onboarding"]),
    _j("HR Executive","Aon Hewitt India","hr","Junior","Pune",
       ["hr","recruitment","payroll","employee relations","communication","hris"]),
    _j("Recruiter","Deloitte India","hr","Mid-Level","Hyderabad",
       ["hr","talent acquisition","recruitment","communication","hris","interviewing"]),
    _j("HR Business Partner","Wipro","hr","Senior","Bangalore",
       ["hr","performance management","employee relations","hris","communication","organizational development"]),
    # ── 11. FINANCE
    _j("Finance Intern","Bajaj Finserv","finance","Internship","Pune",
       ["excel","accounting","tally","ms office","gst"]),
    _j("Accounts Executive – Fresher","Bajaj Finserv","finance","Fresher","Pune",
       ["accounting","tally","gst","excel","taxation","bookkeeping"]),
    _j("Accountant","KPMG India","finance","Junior","Mumbai",
       ["accounting","tally","gst","excel","taxation","accounts payable"]),
    _j("Financial Analyst","Goldman Sachs","finance","Mid-Level","Bangalore",
       ["finance","financial analysis","excel","sql","financial modelling","budgeting","ifrs"]),
    _j("SAP FICO Consultant","Infosys BPM","finance","Mid-Level","Hyderabad",
       ["sap fi","accounting","finance","gst","excel","erp","sap co"]),
    _j("Senior Finance Manager","HDFC Bank","finance","Senior","Mumbai",
       ["finance","financial analysis","excel","budgeting","forecasting","accounting","ifrs","financial modelling"]),
    # ── 12. MECHANICAL
    _j("Mechanical Intern","Tata Motors","mechanical","Internship","Pune",
       ["autocad","solidworks","mechanical design","gd&t"]),
    _j("Mechanical – Fresher","Bajaj Auto","mechanical","Fresher","Pune",
       ["autocad","solidworks","mechanical design","matlab","gd&t"]),
    _j("Mechanical Design Engineer","Mahindra & Mahindra","mechanical","Junior","Pune",
       ["autocad","solidworks","mechanical design","ansys","fea","gd&t"]),
    _j("Design Engineer","Bosch India","mechanical","Mid-Level","Bangalore",
       ["autocad","catia","creo","solidworks","mechanical design","nx","fea","ansys"]),
    _j("Manufacturing Engineer","Tata Motors","mechanical","Mid-Level","Pune",
       ["autocad","solidworks","mechanical design","ansys","matlab","cnc","six sigma"]),
    _j("Senior Mechanical Engineer","L&T Engineering","mechanical","Senior","Mumbai",
       ["autocad","solidworks","catia","ansys","creo","mechanical design","matlab","fea"]),
    # ── 13. CIVIL
    _j("Civil Intern","L&T Construction","civil","Internship","Hyderabad",
       ["autocad","civil engineering","ms office","structural analysis"]),
    _j("Civil – Fresher","DLF","civil","Fresher","Delhi",
       ["civil engineering","autocad","ms office","structural analysis","staad pro"]),
    _j("Civil Site Engineer","L&T Construction","civil","Junior","Hyderabad",
       ["civil engineering","autocad","structural analysis","project management","staad pro"]),
    _j("Structural Design Engineer","Shapoorji Pallonji","civil","Mid-Level","Mumbai",
       ["civil engineering","structural analysis","autocad","staad pro","etabs","rcc design"]),
    _j("Senior Civil Engineer","Gammon India","civil","Senior","Mumbai",
       ["civil engineering","structural analysis","autocad","staad pro","revit structure","project management"]),
    # ── 14. ELECTRICAL
    _j("Electrical Intern","Siemens India","electrical","Internship","Pune",
       ["plc","scada","matlab","electrical","embedded systems"]),
    _j("Electrical Engineer","ABB India","electrical","Junior","Bangalore",
       ["plc","scada","matlab","embedded systems","electrical","hmi","drives"]),
    _j("Automation Engineer","Rockwell Automation","electrical","Mid-Level","Pune",
       ["plc","scada","automation","matlab","embedded systems","electrical","dcs","hmi"]),
    _j("Senior Electrical Engineer","Siemens India","electrical","Senior","Mumbai",
       ["plc","scada","automation","matlab","embedded systems","electrical","dcs","drives"]),
    # ── 15. INTERIOR DESIGN
    _j("Interior Design Intern","Godrej Interio","interior","Internship","Mumbai",
       ["autocad","sketchup","3ds max","canva","interior design"]),
    _j("Interior Designer","Livspace","interior","Junior","Bangalore",
       ["autocad","sketchup","3ds max","v-ray","revit interior","interior design"]),
    _j("3D Interior Designer","HomeLane","interior","Mid-Level","Hyderabad",
       ["autocad","sketchup","3ds max","v-ray","revit interior","lumion","interior design","space planning"]),
    # ── 16. GRAPHIC DESIGN
    _j("Graphic Design Intern","Ogilvy India","design","Internship","Mumbai",
       ["photoshop","illustrator","canva","typography"]),
    _j("Graphic Designer","Dentsu India","design","Junior","Mumbai",
       ["photoshop","illustrator","canva","figma graphic","indesign","branding"]),
    _j("Senior Graphic Designer","Wunderman Thompson","design","Senior","Gurgaon",
       ["photoshop","illustrator","canva","figma graphic","indesign","branding","after effects"]),
    # ── 17. UI/UX
    _j("UI/UX Intern","Flipkart","uiux","Internship","Bangalore",
       ["figma","wireframing","canva","prototyping"]),
    _j("UI Designer","Swiggy","uiux","Junior","Bangalore",
       ["figma","adobe xd","wireframing","prototyping","html","css","design system"]),
    _j("UX Designer","Razorpay","uiux","Mid-Level","Bangalore",
       ["figma","adobe xd","wireframing","prototyping","user research","usability testing"]),
    _j("Product Designer","CRED","uiux","Senior","Bangalore",
       ["figma","adobe xd","wireframing","prototyping","user research","design system","accessibility"]),
    # ── 18. SALES
    _j("Sales Intern","Salesforce India","sales","Internship","Hyderabad",
       ["communication","ms office","crm sales","sales","lead generation"]),
    _j("Sales Executive","HCL Technologies","sales","Junior","Delhi",
       ["crm sales","lead generation","communication","negotiation","sales","b2b sales"]),
    _j("Business Development Executive","Infosys BPM","sales","Mid-Level","Pune",
       ["crm sales","lead generation","communication","negotiation","business development","pipeline management"]),
    _j("Senior BDE / Account Manager","TCS","sales","Senior","Mumbai",
       ["crm sales","lead generation","communication","negotiation","business development","account management","b2b sales"]),
    # ── 19. CONTENT WRITING
    _j("Content Writing Intern","HT Media","content","Internship","Delhi",
       ["content writing","blog writing","ms office","research","seo writing"]),
    _j("Content Writer","Byju's","content","Junior","Bangalore",
       ["content writing","blog writing","seo writing","copywriting","research","proofreading"]),
    _j("Senior Copywriter","Ogilvy India","content","Senior","Mumbai",
       ["content writing","blog writing","seo writing","copywriting","research","branding","creative writing"]),
    # ── 20. CUSTOMER SUPPORT
    _j("Support Intern","Freshworks","support","Internship","Chennai",
       ["communication","ms office","customer support","customer service"]),
    _j("Customer Support Executive","Amazon India","support","Junior","Hyderabad",
       ["communication","crm sales","customer support","customer service","zendesk","sla"]),
    _j("Technical Support Engineer","Cisco India","support","Mid-Level","Bangalore",
       ["technical support","networking","linux","customer support","zendesk","troubleshooting"]),
    # ── 21. HEALTHCARE
    _j("Healthcare Intern","Apollo Hospitals","healthcare","Internship","Hyderabad",
       ["patient care","communication","ms office","clinical care"]),
    _j("Staff Nurse","Fortis Healthcare","healthcare","Junior","Delhi",
       ["patient care","icu","nursing","clinical care","communication"]),
    _j("Pharmacist","Sun Pharmaceuticals","healthcare","Junior","Mumbai",
       ["pharmacology","patient care","medical billing","drug dispensing","ms office"]),
    _j("Healthcare Executive","Manipal Hospitals","healthcare","Mid-Level","Bangalore",
       ["patient care","icu","nursing","pharmacology","medical billing","ehr"]),
    # ── 22. TEACHING / EDUCATION
    _j("Teaching Intern","Byju's","education","Internship","Bangalore",
       ["teaching","communication","ms office","lesson planning"]),
    _j("Teacher","Vedantu","education","Junior","Remote",
       ["teaching","communication","curriculum planning","lesson planning","e-learning"]),
    _j("Ed-Tech Content Creator","Unacademy","education","Mid-Level","Bangalore",
       ["teaching","content writing","curriculum planning","e-learning","lms","instructional design"]),
    _j("Professor / Faculty","IIT Hyderabad","education","Senior","Hyderabad",
       ["teaching","curriculum planning","academic research","assessment","presentation","mentoring"]),
]


# ══════════════════════════════════════════════════════════════
#  MATCHING ENGINE  — weighted skill overlap
# ══════════════════════════════════════════════════════════════
DOMAIN_WEIGHTS = {
    "devops":       {"docker","kubernetes","terraform","ansible","ci/cd","aws","azure","gcp","helm","jenkins","sre"},
    "cybersecurity":{"ethical hacking","penetration testing","vapt","siem","soc","owasp","splunk","burp suite","kali linux"},
    "data_ai":      {"machine learning","deep learning","tensorflow","pytorch","nlp","scikit-learn","data science","etl","airflow"},
    "fullstack":    {"react","node.js","typescript","angular","vue","mongodb","express","graphql","next.js"},
    "python":       {"python","django","flask","fastapi","celery","asyncio","sqlalchemy"},
    "java":         {"java","spring boot","microservices","kafka","hibernate","maven","gradle"},
    "mobile":       {"android","flutter","kotlin android","swift","react native","dart","firebase","retrofit"},
    "networking":   {"cisco","routing","switching","bgp","ospf","ccna","vlan","mpls","tcp/ip"},
    "marketing":    {"seo","google ads","facebook ads","google analytics","sem","ppc","content marketing","digital marketing"},
    "hr":           {"recruitment","talent acquisition","hris","payroll","performance management","hrms"},
    "finance":      {"tally","gst","financial analysis","financial modelling","sap fi","ifrs","accounts payable"},
    "mechanical":   {"solidworks","catia","ansys","creo","fea","autocad","gd&t","mechanical design"},
    "civil":        {"staad pro","etabs","structural analysis","rcc design","revit structure","civil engineering"},
    "electrical":   {"plc","scada","automation","dcs","hmi","drives"},
    "interior":     {"sketchup","3ds max","v-ray","lumion","space planning","interior design"},
    "uiux":         {"figma","wireframing","user research","prototyping","design system","adobe xd"},
    "design":       {"photoshop","illustrator","branding","indesign","after effects"},
    "sales":        {"lead generation","crm sales","b2b sales","account management","pipeline management"},
    "content":      {"seo writing","copywriting","content writing","blog writing","creative writing"},
    "support":      {"zendesk","freshdesk","ticketing","sla","troubleshooting"},
    "healthcare":   {"patient care","nursing","icu","pharmacology","medical billing"},
    "education":    {"curriculum planning","lesson planning","e-learning","lms","instructional design"},
    "it_general":   {"python","java","javascript","sql","linux","docker","api","git"},
}

def compute_match_pct(resume_skills: set, job: dict) -> int:
    job_skills = {s.lower() for s in job["skills"]}
    domain_kw  = DOMAIN_WEIGHTS.get(job["domain"], set())
    weighted_match = 0
    max_possible   = 0
    for sk in job_skills:
        w = 2 if sk in domain_kw else 1
        max_possible += w
        if sk in resume_skills:
            weighted_match += w
    if max_possible == 0: return 0
    return min(int((weighted_match / max_possible) * 100), 99)

def match_jobs(skills: list, predicted_domain: str, location: str) -> list:
    skill_set = {s.lower() for s in skills}
    sel_loc   = location.lower().strip()
    matched   = []
    for job in JOB_DATA:
        jd = job["domain"]
        if predicted_domain == "it_general":
            if jd not in IT_DOMAINS: continue
        else:
            if jd != predicted_domain: continue
        job_loc = job["location"].lower().strip()
        if sel_loc not in ("all over india",) and job_loc not in (sel_loc, "remote"):
            continue
        pct = compute_match_pct(skill_set, job)
        if pct == 0: continue
        matched.append({**job, "match_pct": pct})
    matched.sort(key=lambda x: x["match_pct"], reverse=True)
    return matched


# ══════════════════════════════════════════════════════════════
#  TEXT EXTRACTION
# ══════════════════════════════════════════════════════════════
def extract_text_from_pdf(fb):
    if not PDF_SUPPORT: return ""
    try:
        rd = PyPDF2.PdfReader(io.BytesIO(fb))
        return "\n".join((p.extract_text() or "") for p in rd.pages)
    except Exception: return ""

def extract_text_from_docx(fb):
    if not DOCX_SUPPORT: return ""
    try:
        doc = Document(io.BytesIO(fb))
        return "\n".join(p.text for p in doc.paragraphs)
    except Exception: return ""

def extract_email(text):
    m = re.search(r'[\w.\-+]+@[\w.\-]+\.[a-zA-Z]{2,}', text or "")
    return m.group(0) if m else "Not found"

def extract_phone(text):
    m = re.search(r'(\+?\d[\d\s\-().]{8,14}\d)', text or "")
    return m.group(0).strip() if m else "Not found"

def extract_name(text):
    if not isinstance(text, str): return "Not found"
    lines = [l.strip() for l in text.split('\n') if l.strip()]
    for line in lines[:8]:
        words = line.split()
        if 2 <= len(words) <= 4 and all(re.match(r"^[A-Za-z.'\-]+$", w) for w in words):
            return line
    return lines[0] if lines else "Not found"

def extract_skills(text: str) -> list:
    if not isinstance(text, str) or not text.strip(): return []
    tl = text.lower()
    found = []
    for sk in sorted(SKILL_KEYWORDS, key=len, reverse=True):
        if re.search(r'\b' + re.escape(sk) + r'\b', tl):
            found.append(sk)
    return list(dict.fromkeys(found))

def extract_education(text):
    if not isinstance(text, str) or not text.strip(): return "Not specified"
    edu_kw = ["b.tech","b.e.","m.tech","mba","bsc","msc","bachelor","master","phd","diploma",
               "degree","university","college","institute","10th","12th","higher secondary",
               "engineering","technology","science","arts","commerce","school",
               "graduation","post graduation","pg","ug"]
    seen, lines_out = set(), []
    for line in text.split('\n'):
        cl = (line or "").strip()
        if not cl: continue
        if any(kw in cl.lower() for kw in edu_kw):
            norm = ' '.join(cl.split())
            if norm not in seen: seen.add(norm); lines_out.append(norm)
    return " | ".join(lines_out[:4]) if lines_out else "Not specified"

def extract_experience(text):
    if not isinstance(text, str) or not text.strip(): return "Not specified"
    exp_kw = ["experience","worked at","work experience","employment",
               "internship","intern","engineer","developer","analyst",
               "manager","executive","consultant","associate","senior","junior",
               "lead","architect","specialist","officer","coordinator"]
    seen, lines_out = set(), []
    for line in text.split('\n'):
        cl = (line or "").strip()
        if not cl or len(cl) < 8: continue
        if any(kw in cl.lower() for kw in exp_kw):
            norm = ' '.join(cl.split())
            if norm not in seen: seen.add(norm); lines_out.append(norm)
    return " | ".join(lines_out[:4]) if lines_out else "Not specified"

def generate_summary(name, skills, education, experience, role):
    top   = skills[:6] if skills else []
    sk    = ", ".join(top) if top else "various technologies"
    edu_s = (education[:70]+"…") if len(education) > 70 else education
    exp_s = (experience[:70]+"…") if len(experience) > 70 else experience
    return (f"{name} is a skilled {role} proficient in {sk}. "
            f"Educational background: {edu_s}. "
            f"Professional experience: {exp_s}.")

def parse_resume(file_bytes, filename):
    ext = filename.rsplit('.', 1)[-1].lower() if '.' in filename else "txt"
    if   ext == "pdf":          text = extract_text_from_pdf(file_bytes)
    elif ext in ("docx","doc"): text = extract_text_from_docx(file_bytes)
    else:                       text = file_bytes.decode("utf-8", errors="ignore")
    skills     = extract_skills(text)
    education  = extract_education(text)
    experience = extract_experience(text)
    name       = extract_name(text)
    email      = extract_email(text)
    phone      = extract_phone(text)
    role, domain = suggest_role(skills)
    summary    = generate_summary(name, skills, education, experience, role)
    return {"name":name,"email":email,"phone":phone,"skills":skills,
            "education":education,"experience":experience,"summary":summary,
            "suggested_role":role,"domain":domain,"raw_text":text}


# ══════════════════════════════════════════════════════════════
#  REPORT GENERATOR
# ══════════════════════════════════════════════════════════════
def generate_report(info, jobs):
    lines = [
        "="*64,
        "      RESUME EXTRACTOR JOB PORTAL  —  ANALYSIS REPORT",
        "="*64,
        f"Generated : {datetime.now().strftime('%d %b %Y, %I:%M %p')}",
        "",
        "── PERSONAL INFORMATION",
        f"Name   : {info.get('name','N/A')}",
        f"Email  : {info.get('email','N/A')}",
        f"Phone  : {info.get('phone','N/A')}",
        f"Role   : {info.get('suggested_role','N/A')}",
        f"Domain : {info.get('domain','N/A')}",
        "",
        "── SKILLS DETECTED",
        ", ".join(info.get('skills',[])) or "N/A",
        "",
        "── EDUCATION",
        info.get('education','N/A'),
        "",
        "── EXPERIENCE",
        info.get('experience','N/A'),
        "",
        "── RESUME SUMMARY",
        info.get('summary','N/A'),
        "",
        "── MATCHED JOBS / INTERNSHIPS",
    ]
    if jobs:
        for i, j in enumerate(jobs, 1):
            lines += [
                f"{i:>2}. [{j['level']}] {j['title']} @ {j['company']}",
                f"    Location : {j['location']}",
                f"    Match    : {j['match_pct']}%",
                f"    Naukri   : {j['apply_naukri']}",
                f"    LinkedIn : {j['apply_linkedin']}",
                f"    Indeed   : {j['apply_indeed']}",
                f"    Foundit  : {j['apply_foundit']}",
                "",
            ]
    else:
        lines.append("Currently no openings available in this location.")
    lines.append("="*64)
    return "\n".join(lines)



def login_required(view):
    @wraps(view)
    def wrapped(*args, **kwargs):
        if "user_id" not in session:
            flash("Please log in first.", "warning")
            return redirect(url_for("login"))
        return view(*args, **kwargs)
    return wrapped

@app.route("/")
def index():
    if session.get("user_id"):
        return redirect(url_for("dashboard"))
    return redirect(url_for("login"))

@app.route("/register", methods=["GET", "POST"])
def register():
    if request.method == "POST":
        username = request.form.get("username", "").strip()
        email = request.form.get("email", "").strip()
        password = request.form.get("password", "")
        if not username or not email or len(password) < 8:
            flash("Enter username, email and a password with at least 8 characters.", "danger")
        else:
            ok, message = create_user(username, email, password)
            flash(message, "success" if ok else "danger")
            if ok:
                return redirect(url_for("login"))
    return render_template("register.html")

@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        user = verify_user(request.form.get("username", ""), request.form.get("password", ""))
        if user:
            session.clear()
            session["user_id"] = user["id"]
            session["username"] = user["username"]
            session["email"] = user["email"]
            return redirect(url_for("dashboard"))
        flash("Invalid username or password.", "danger")
    return render_template("login.html")

@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))

@app.route("/dashboard")
@login_required
def dashboard():
    profile = get_user_profile(session["user_id"])
    resumes = get_user_resumes(session["user_id"])
    return render_template("dashboard.html", profile=profile, resumes=resumes)

@app.route("/profile", methods=["GET", "POST"])
@login_required
def profile():
    current = get_user_profile(session["user_id"])
    if request.method == "POST":
        location = request.form.get("preferred_location", "All Over India")
        role = request.form.get("suggested_role", "").strip()
        image = request.files.get("profile_image")
        image_bytes = image.read() if image and image.filename else None
        update_profile(session["user_id"], location, role, image_bytes)
        flash("Profile updated.", "success")
        return redirect(url_for("profile"))
    return render_template("profile.html", profile=current, locations=LOCATIONS)

@app.route("/upload", methods=["GET", "POST"])
@login_required
def upload():
    if request.method == "POST":
        file = request.files.get("resume")
        location = request.form.get("location", "All Over India")
        if not file or not file.filename:
            flash("Choose a resume file first.", "danger")
            return redirect(url_for("upload"))
        filename = secure_filename(file.filename)
        ext = filename.rsplit(".", 1)[-1].lower() if "." in filename else ""
        if ext not in ALLOWED_EXTENSIONS:
            flash("Supported formats: PDF, DOCX, TXT.", "danger")
            return redirect(url_for("upload"))
        file_bytes = file.read()
        info = parse_resume(file_bytes, filename)
        if not info.get("raw_text", "").strip():
            flash("Could not extract text. Scanned/image-only PDFs may need OCR.", "warning")
        jobs = match_jobs(info["skills"], info["domain"], location)
        save_resume(session["user_id"], filename, info)
        session["last_report"] = {
            "filename": filename, "info": {k:v for k,v in info.items() if k != "raw_text"},
            "jobs": jobs, "location": location
        }
        return redirect(url_for("results"))
    return render_template("upload.html", locations=LOCATIONS)

@app.route("/results")
@login_required
def results():
    report = session.get("last_report")
    if not report:
        flash("Upload a resume to see results.", "info")
        return redirect(url_for("upload"))
    return render_template("results.html", report=report)

@app.route("/download-report")
@login_required
def download_report():
    report = session.get("last_report")
    if not report:
        flash("No report available yet.", "warning")
        return redirect(url_for("upload"))
    content = generate_report(report["info"], report["jobs"])
    return send_file(io.BytesIO(content.encode("utf-8")), mimetype="text/plain",
                     as_attachment=True,
                     download_name="resume_report.txt")

@app.route("/reports")
@login_required
def reports():
    return render_template("reports.html", resumes=get_user_resumes(session["user_id"]))

@app.route("/reports/<int:resume_id>/delete", methods=["POST"])
@login_required
def remove_resume(resume_id):
    delete_resume(resume_id, session["user_id"])
    flash("Resume report deleted.", "success")
    return redirect(url_for("reports"))

@app.errorhandler(413)
def too_large(_error):
    flash("File is too large. Maximum upload size is 10 MB.", "danger")
    return redirect(url_for("upload"))

init_db()

if __name__ == "__main__":
