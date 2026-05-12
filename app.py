import os
import re
import msal
import pandas as pd
import requests
import streamlit as st
from anthropic import Anthropic
from docx import Document
from pypdf import PdfReader

# ================= CONFIG =================
st.set_page_config(page_title="Enterprise QA Copilot", layout="wide")

st.title("🚀 Enterprise QA Copilot (Claude + AI)")
st.caption("Chat + Files + Smart AI + Memory + Analytics")

# ================= ENV =================
ANTHROPIC_API_KEY = os.getenv("ANTHROPIC_API_KEY")
CLIENT_ID = os.getenv("CLIENT_ID")
TENANT_ID = os.getenv("TENANT_ID")

# ✅ Validate AI key
if not ANTHROPIC_API_KEY:
    st.error("❌ Missing ANTHROPIC_API_KEY")
    st.stop()

# ✅ Init AI
try:
    client = Anthropic(api_key=ANTHROPIC_API_KEY)
except Exception as e:
    st.error(f"❌ AI Init Error: {e}")
    st.stop()

# ✅ Microsoft optional
graph_enabled = True
if not CLIENT_ID or not TENANT_ID:
    graph_enabled = False
    st.warning("⚠️ Microsoft integration disabled")

AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPES = ["Files.Read"]

# ================= SESSION =================
if "auth" not in st.session_state:
    st.session_state.auth = False

if "context" not in st.session_state:
    st.session_state.context = ""

if "history" not in st.session_state:
    st.session_state.history = []

# ================= LOGIN (SECURE) =================
VALID_USER = os.getenv("APP_USER", "admin")
VALID_PASS = os.getenv("APP_PASS", "admin123")

if not st.session_state.auth:
    st.subheader("🔐 Login")

    u = st.text_input("Username")
    p = st.text_input("Password", type="password")

    if st.button("Login"):
        if u == VALID_USER and p == VALID_PASS:
            st.session_state.auth = True
            st.rerun()
        else:
            st.error("Invalid credentials")

    st.stop()

# ================= MODEL =================
MODELS = {
    "fast": "claude-haiku-4-5-20251001",
    "balanced": "claude-sonnet-4-6",
    "powerful": "claude-opus-4-7"
}

def choose_model(q, ctx):
    if len(ctx) > 4000 or len(q) > 300:
        return MODELS["powerful"]
    elif len(q) < 80:
        return MODELS["fast"]
    return MODELS["balanced"]

# ================= SMART INPUT DETECTION =================
def detect_input_type(text):
    text = text.lower()

    if any(x in text for x in ["excel", "column", "row", "dataframe"]):
        return "structured"
    if any(x in text for x in ["pdf", "document", "summarize"]):
        return "document"
    if any(x in text for x in ["json", "{", "}"]):
        return "json"
    return "general"

# ================= PROMPT BUILDER =================
def build_prompt(query, context):
    input_type = detect_input_type(query)

    return f"""
You are an enterprise AI copilot.

Context:
{context}

User Query:
{query}

Input Type: {input_type}

Instructions:
- If structured → return table + insights
- If document → summarize + key points
- If general → clear explanation
- NEVER stop mid response
- Continue until fully complete
"""

# ================= FILE HANDLING =================
def extract_file(file):
    try:
        if file.name.endswith(".docx"):
            return "\n".join(p.text for p in Document(file).paragraphs)

        if file.name.endswith(".pdf"):
            return "\n".join(
                p.extract_text() for p in PdfReader(file.pages) if p.extract_text()
            )

        if file.name.endswith((".xlsx", ".xls")):
            df = pd.read_excel(file)
            return df.to_string()

        return file.read().decode("utf-8", "ignore")

    except:
        return ""

# ================= GRAPH API =================
def get_graph_token():
    try:
        app = msal.PublicClientApplication(CLIENT_ID, authority=AUTHORITY)
        flow = app.initiate_device_flow(scopes=SCOPES)

        if "user_code" not in flow:
            st.error("❌ Login start failed")
            return None

        st.info(f"👉 Open: {flow['verification_uri']}")
        st.info(f"👉 Code: {flow['user_code']}")

        result = app.acquire_token_by_device_flow(flow)

        return result.get("access_token")

    except Exception as e:
        st.error(f"Microsoft error: {e}")
        return None

# ================= AI ENGINE =================
def call_ai(prompt, model):
    output = ""
    current_prompt = prompt

    for _ in range(5):  # Continue loop
        try:
            res = client.messages.create(
                model=model,
                max_tokens=4000,
                messages=[{"role": "user", "content": current_prompt}]
            )

            chunk = res.content[0].text
            output += chunk

            # Stop if small chunk
            if len(chunk) < 200:
                break

            current_prompt = "Continue exactly from where you stopped."

        except Exception as e:
            return f"AI Error: {e}"

    return output

# ================= FILE UPLOAD =================
st.subheader("📂 Upload Files")

files = st.file_uploader("Upload", accept_multiple_files=True)

if files:
    content = ""
    for f in files:
        content += extract_file(f) + "\n"

    st.session_state.context = content
    st.success("✅ Files added to context")

# ================= GRAPH UI =================
if graph_enabled:
    st.subheader("☁️ OneDrive")

    if st.button("Connect Microsoft"):
        token = get_graph_token()
        if token:
            st.success("✅ Connected")

# ================= CHAT =================
st.subheader("💬 Chat")

for msg in st.session_state.history:
    st.chat_message(msg["role"]).write(msg["content"])

query = st.chat_input("Ask anything...")

if query:
    st.session_state.history.append({"role": "user", "content": query})

    prompt = build_prompt(query, st.session_state.context)
    model = choose_model(query, st.session_state.context)

    response = call_ai(prompt, model)

    st.session_state.history.append({"role": "assistant", "content": response})
    st.rerun()