import streamlit as st
import os, time, json, re, io, requests
import pandas as pd
from dotenv import load_dotenv
from anthropic import Anthropic
from docx import Document
from pypdf import PdfReader

# ================= CONFIG =================
st.set_page_config(page_title="Enterprise QA AI", layout="wide")

st.markdown("""
<style>
.main-title {font-size:30px;font-weight:bold;color:#2c3e50;}
.sub-title {color:gray;}
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-title">🚀 Enterprise QA AI Platform</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-title">Full QA Automation + Intelligence</div>', unsafe_allow_html=True)

# ================= ENV =================
load_dotenv()
API_KEY = os.getenv("ANTHROPIC_API_KEY")
client = Anthropic(api_key=API_KEY)

# ================= LOGIN =================
USERS = {"rahul": "password123"}
if "auth" not in st.session_state:
    st.session_state.auth = False

if not st.session_state.auth:
    u = st.text_input("Username")
    p = st.text_input("Password", type="password")
    if st.button("Login"):
        if USERS.get(u) == p:
            st.session_state.auth = True
            st.rerun()
        else:
            st.error("Invalid credentials")
    st.stop()

# ================= MODELS =================
MODELS = {
    "fast": "claude-haiku-4-5-20251001",
    "balanced": "claude-sonnet-4-6",
    "powerful": "claude-opus-4-7"
}

mode = st.sidebar.selectbox("Model Mode", ["Auto", "fast", "balanced", "powerful"])

def choose_model(q, ctx):
    if mode != "Auto": return MODELS[mode]
    if len(ctx) > 4000: return MODELS["powerful"]
    if len(ctx) < 500: return MODELS["fast"]
    return MODELS["balanced"]

# ================= SECURITY =================
def mask(text):
    text = re.sub(r"\S+@\S+", "[EMAIL]", text)
    text = re.sub(r"\b\d{10}\b", "[PHONE]", text)
    return text

# ================= RATE LIMIT =================
if "last_call" not in st.session_state:
    st.session_state.last_call = 0

def rate_limit():
    if time.time() - st.session_state.last_call < 2:
        st.warning("⏳ Please wait...")
        return False
    st.session_state.last_call = time.time()
    return True

# ================= FILE =================
def extract(file):
    name = file.name.lower()
    try:
        if name.endswith((".xlsx",".xls")):
            return pd.read_excel(file, engine="openpyxl")
        if name.endswith(".csv"):
            return pd.read_csv(file)
        if name.endswith(".docx"):
            return "\n".join(p.text for p in Document(file).paragraphs)
        if name.endswith(".pdf"):
            return "\n".join(p.extract_text() for p in PdfReader(file).pages if p.extract_text())
        raw=file.read()
        return raw.decode("utf-8","ignore")
    except:
        return ""

def chunk(data):
    if isinstance(data, pd.DataFrame):
        return data.fillna("").to_dict(orient="records")
    return [data[i:i+2000] for i in range(0,len(data),2000)]

# ================= CLOUD =================
def load_file(url):
    try:
        r = requests.get(url)
        f = io.BytesIO(r.content)
        f.name="cloud"
        return f
    except:
        return None

# ================= AI =================
def call_ai(prompt, model):
    try:
        res = client.messages.create(
            model=model,
            max_tokens=2000,
            messages=[{"role":"user","content":mask(prompt)}]
        )
        return res.content[0].text
    except:
        return "AI Error"

def build_prompt(ctx):
    return f"""
Generate enterprise QA test cases JSON.

Fields:
Test ID, Story, Test Title, Steps, Expected Result,
Priority, Test Type, Automation, Status

Context:
{ctx}

Return JSON only.
"""

# ================= JSON =================
def parse_json(text):
    try:
        js = text[text.find("["):text.rfind("]")+1]
        return pd.DataFrame(json.loads(js))
    except:
        return None

# ================= DUPLICATES =================
def detect_duplicates(df):
    return df[df.duplicated("Test Title", keep=False)] if "Test Title" in df else pd.DataFrame()

# ================= MODULE =================
def detect_module(df):
    txt=" ".join(df.astype(str).values.flatten())
    m=re.findall(r"M\d{2}",txt)
    return max(set(m),key=m.count) if m else "M01"

# ================= IDS =================
def assign_ids(df,module):
    ctr={"FT":0,"AT":0,"DT":0,"IT":0,"ST":0}
    def p(t):
        if "API" in t: return "AT"
        if "DB" in t: return "DT"
        if "Integration" in t: return "IT"
        if "Functional" in t: return "FT"
        return "ST"
    ids=[]
    for _,r in df.iterrows():
        pre=p(r["Test Type"])
        ctr[pre]+=1
        ids.append(f"{pre}-{module}-{str(ctr[pre]).zfill(3)}")
    df["Test ID"]=ids
    return df

# ================= EXCEL =================
def build_excel(df):
    module=detect_module(df)
    df=assign_ids(df,module)

    output=io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        wb=writer.book
        header=wb.add_format({'bold':True,'bg_color':'#2C3E50','font_color':'white'})
        wrap=wb.add_format({'text_wrap':True})

        df.to_excel(writer,"Functional",index=False)
        ws=writer.sheets["Functional"]

        for c,col in enumerate(df.columns):
            ws.write(0,c,col,header)

        ws.set_column(0,len(df.columns)-1,25,wrap)

    return output

# ================= STATE =================
if "data" not in st.session_state: st.session_state.data=[]
if "ans" not in st.session_state: st.session_state.ans=""

# ================= INPUT =================
tab1,tab2=st.tabs(["📂 Local","🌐 Cloud"])

with tab1:
    files = st.file_uploader("Upload Files",accept_multiple_files=True)
    if files:
        st.session_state.data=[]
        for f in files:
            st.session_state.data+=chunk(extract(f))

with tab2:
    urls=st.text_area("Paste URLs")
    if st.button("Load URLs"):
        for u in urls.splitlines():
            f=load_file(u)
            if f:
                st.session_state.data+=chunk(extract(f))

# ================= ACTION =================
if st.button("✅ Generate Test Pack"):
    if not rate_limit(): st.stop()
    ctx=str(st.session_state.data)
    model=choose_model(ctx,ctx)
    st.session_state.ans=call_ai(build_prompt(ctx),model)

# ================= CHAT =================
st.subheader("💬 Chat")
q=st.text_input("Ask anything")
if st.button("Send Chat"):
    model=choose_model(q,"")
    st.session_state.ans=call_ai(q,model)

# ================= OUTPUT =================
if st.session_state.ans:
    st.write(st.session_state.ans)

    df=parse_json(st.session_state.ans)
    if df is not None:

        st.subheader("📊 Test Cases")
        st.dataframe(df)

        st.subheader("⚠️ Duplicates")
        st.dataframe(detect_duplicates(df))

        excel=build_excel(df)
        st.download_button("📥 Download Excel",excel.getvalue(),"QA_TestPack.xlsx")

# ================= COMPARISON =================
st.subheader("🔍 Excel Comparison")

c1, c2 = st.columns(2)

file1 = c1.file_uploader("Old File", key="old")
file2 = c2.file_uploader("New File", key="new")

if st.button("Compare Files"):
    if file1 and file2:
        df1=pd.read_excel(file1)
        df2=pd.read_excel(file2)

        old=set(df1["Test Title"])
        new=set(df2["Test Title"])

        st.write("Missing:", df1[df1["Test Title"].isin(old-new)])
        st.write("New:", df2[df2["Test Title"].isin(new-old)])
        st.write("Common:", df2[df2["Test Title"].isin(old&new)])
