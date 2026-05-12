import streamlit as st
import os
import time
import json
import re
from dotenv import load_dotenv
from anthropic import Anthropic
import pandas as pd
from docx import Document
from pypdf import PdfReader
from bs4 import BeautifulSoup
import chardet
import io

# ================= CONFIG =================
st.set_page_config(page_title="Enterprise QA AI", layout="wide")

# ================= ENV =================
load_dotenv()
API_KEY = os.getenv("ANTHROPIC_API_KEY")

if not API_KEY:
    st.error("Missing API key")
    st.stop()

client = Anthropic(api_key=API_KEY)

# ================= AUTH =================
USERS = {"rahul": "password123"}

def login():
    st.title("🔐 Login")
    u = st.text_input("Username")
    p = st.text_input("Password", type="password")

    if st.button("Login"):
        if USERS.get(u) == p:
            st.session_state.auth = True
            st.session_state.user = u
            st.rerun()
        else:
            st.error("Invalid")

if "auth" not in st.session_state:
    st.session_state.auth = False

if not st.session_state.auth:
    login()
    st.stop()

# ================= MODELS =================
MODELS = {
    "fast": "claude-haiku-4-5-20251001",
    "balanced": "claude-sonnet-4-6",
    "powerful": "claude-opus-4-7"
}

FALLBACK = list(MODELS.values())

# ================= SECURITY =================
def mask(text):
    text = re.sub(r"\b\d{10}\b", "[PHONE]", text)
    text = re.sub(r"\S+@\S+", "[EMAIL]", text)
    return text

def log(user, action):
    with open("audit.log", "a") as f:
        f.write(f"{time.ctime()} | {user} | {action}\n")

# ================= RATE LIMIT =================
if "last_call" not in st.session_state:
    st.session_state.last_call = 0

def rate_limit():
    now = time.time()
    if now - st.session_state.last_call < 2:
        st.warning("Slow down")
        return False
    st.session_state.last_call = now
    return True

# ================= FILE PROCESS =================
def extract(file):
    name = file.name.lower()
    try:
        if name.endswith(".docx"):
            return "\n".join(p.text for p in Document(file).paragraphs)
        if name.endswith(".pdf"):
            r = PdfReader(file)
            return "\n".join(p.extract_text() for p in r.pages if p.extract_text())
        if name.endswith(".csv"):
            return pd.read_csv(file).to_string()
        if name.endswith((".txt",".md")):
            raw = file.read()
            enc = chardet.detect(raw)["encoding"]
            return raw.decode(enc or "utf-8")
        if name.endswith(".html"):
            return BeautifulSoup(file.read(),"html.parser").get_text()
    except:
        return ""
    return ""

def chunk(text):
    return [text[i:i+2000] for i in range(0,len(text),2000)]

def retrieve(q,chunks):
    qset=set(q.lower().split())
    scored=[(len(qset & set(c.lower().split())),c) for c in chunks]
    return "\n\n".join(c for s,c in sorted(scored,reverse=True)[:5])

# ================= MODEL SELECT =================
def choose_model(q,action,ctx):
    if len(ctx)>4000 or len(q)>300:
        return MODELS["powerful"]
    if action in ["test_cases"]:
        return MODELS["balanced"]
    if len(q)<80:
        return MODELS["fast"]
    return MODELS["balanced"]

# ================= AI =================
def call_ai(prompt,model):
    for m in [model]+FALLBACK:
        try:
            res=client.messages.create(
                model=m,
                max_tokens=2000,
                messages=[{"role":"user","content":mask(prompt)}]
            )
            return res.content[0].text, m
        except:
            continue
    return "Error","none"

# ================= TEST CASE JSON =================
def parse_json(text):
    try:
        js=text[text.find("["):text.rfind("]")+1]
        return pd.DataFrame(json.loads(js))
    except:
        return None

# ================= EXCEL BUILDER =================
def build_excel(df):
    output=io.BytesIO()

    with pd.ExcelWriter(output,engine='openpyxl') as writer:

        # Summary
        summary=pd.DataFrame({
            "Metric":["Total","High","Medium","Low"],
            "Value":[len(df),
                len(df[df["Priority"].str.contains("High",na=False)]),
                len(df[df["Priority"].str.contains("Medium",na=False)]),
                len(df[df["Priority"].str.contains("Low",na=False)])
            ]
        })
        summary.to_excel(writer,sheet_name="Cover/Summary",index=False)

        # Coverage
        cov=df.groupby("Story")["Test ID"].count().reset_index()
        cov.columns=["Story","Test Count"]
        cov.to_excel(writer,"Coverage Matrix",index=False)

        # Sheets
        types={
            "01 Functional Tests":"Functional",
            "02 API Tests":"API",
            "03 DB Tests":"DB",
            "04 Integration Tests":"Integration"
        }

        for name,t in types.items():
            sub=df[df["Test Type"].str.contains(t,case=False,na=False)]
            if not sub.empty:
                sub.to_excel(writer,name,index=False)

    return output

# ================= PROMPT =================
def build_prompt(ctx):
    return f"""
Generate test cases in STRICT JSON format:

Fields:
Test ID, Story, Test Title, Pre-conditions, Steps,
Expected Result, Test Data, Priority, Test Type,
Automation, Status

Context:
{ctx}

Return ONLY JSON list.
"""

# ================= UI =================
st.title("🚀 Enterprise QA AI Assistant")

if "chat" not in st.session_state: st.session_state.chat=[]
if "chunks" not in st.session_state: st.session_state.chunks=[]
if "answer" not in st.session_state: st.session_state.answer=""

# Upload
files=st.file_uploader("Upload Docs",accept_multiple_files=True)

if files:
    st.session_state.chunks=[]
    for f in files:
        if f.size>5*1024*1024:
            st.error("File too large")
            st.stop()
        st.session_state.chunks+=chunk(extract(f))

# Button
if st.button("✅ Generate Test Pack"):
    if not rate_limit(): st.stop()

    ctx=retrieve("test cases",st.session_state.chunks)
    prompt=build_prompt(ctx)
    model=choose_model("test cases","test_cases",ctx)

    ans,m=call_ai(prompt,model)
    st.session_state.answer=ans

    # Auto save
    with open("session.json","w") as f:
        json.dump(ans,f)

# Display
if st.session_state.answer:
    st.code(st.session_state.answer)

    df=parse_json(st.session_state.answer)

    if df is not None:
        excel=build_excel(df)
        st.download_button("📥 Download QA Excel",excel.getvalue(),"QA_TestPack.xlsx")
    else:
        st.error("Parsing failed")
