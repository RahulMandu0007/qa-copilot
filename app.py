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

st.title("🚀 Enterprise QA Copilot (Claude + GPT + Copilot)")
st.caption("Chat + Files + Excel Export + Memory + Analytics")

# ================= ENV =================
ANTHROPIC_API_KEY = os.getenv("ANTHROPIC_API_KEY")
CLIENT_ID = os.getenv("CLIENT_ID")
TENANT_ID = os.getenv("TENANT_ID")

# ✅ Validation
if not ANTHROPIC_API_KEY:
    st.error("❌ Missing ANTHROPIC_API_KEY (check Streamlit Secrets)")
    st.stop()

# ✅ Anthropic init (safe)
try:
    client = Anthropic(api_key=ANTHROPIC_API_KEY)
except Exception as e:
    st.error(f"❌ Failed to initialize AI client: {e}")
    st.stop()

# ✅ Microsoft config validation (NON-BLOCKING)
graph_enabled = True
if not CLIENT_ID or not TENANT_ID:
    graph_enabled = False
    st.warning("⚠️ Microsoft integration disabled (CLIENT_ID / TENANT_ID missing)")

# ================= SESSION =================
if "auth" not in st.session_state:
    st.session_state.auth = False

if "context" not in st.session_state:
    st.session_state.context = ""

if "history" not in st.session_state:
    st.session_state.history = []

if "current_folder" not in st.session_state:
    st.session_state.current_folder = None

if "folder_stack" not in st.session_state:
    st.session_state.folder_stack = []

# ================= LOGIN =================
USERS = {"rahul": "password123"}

if not st.session_state.auth:
    st.subheader("🔐 Login")
    u = st.text_input("Username")
    p = st.text_input("Password", type="password")

    if st.button("Login"):
        if USERS.get(u) == p:
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

# ================= GRAPH CONFIG =================
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPES = ["Files.Read"]

# ================= SECURITY =================
def mask(text):
    text = re.sub(r"\S+@\S+", "[EMAIL]", text)
    text = re.sub(r"\b\d{10}\b", "[PHONE]", text)
    return text

# ================= FILE HANDLING =================
def extract_file(file):
    name = file.name.lower()
    try:
        if name.endswith(".docx"):
            return "\n".join(p.text for p in Document(file).paragraphs)

        if name.endswith(".pdf"):
            return "\n".join(
                p.extract_text() for p in PdfReader(file).pages if p.extract_text()
            )

        if name.endswith((".xlsx", ".xls")):
            df = pd.read_excel(file, engine="openpyxl")
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
        st.info(f"👉 Enter Code: {flow['user_code']}")

        result = app.acquire_token_by_device_flow(flow)

        if "access_token" in result:
            return result["access_token"]

        st.error("❌ Login failed")
        return None

    except Exception as e:
        st.error(f"❌ Microsoft Login Error: {e}")
        return None


def list_drive_items(token, folder_id=None):
    headers = {"Authorization": f"Bearer {token}"}
    try:
        url = (
            f"https://graph.microsoft.com/v1.0/me/drive/items/{folder_id}/children"
            if folder_id
            else "https://graph.microsoft.com/v1.0/me/drive/root/children"
        )

        res = requests.get(url, headers=headers)

        if res.status_code != 200:
            return []

        return res.json().get("value", [])

    except Exception as e:
        st.error(f"Graph error: {e}")
        return []

# ================= AI =================
def call_ai(prompt, model):
    try:
        res = client.messages.create(
            model=model,
            max_tokens=4000,
            messages=[{"role": "user", "content": mask(prompt)}]
        )
        return res.content[0].text

    except Exception as e:
        return f"AI Error: {e}"

# ================= FILE UPLOAD =================
st.subheader("📂 Upload Files")
files = st.file_uploader("Upload", accept_multiple_files=True)

if files:
    content = ""
    for f in files:
        content += extract_file(f) + "\n"
    st.session_state.context = content
    st.success("✅ Files added")

# ================= GRAPH UI =================
if graph_enabled:
    st.subheader("☁️ OneDrive Explorer")

    if st.button("Connect Microsoft"):
        token = get_graph_token()
        if token:
            st.session_state.graph_token = token

    if "graph_token" in st.session_state:

        items = list_drive_items(
            st.session_state.graph_token,
            st.session_state.current_folder
        )

        for item in items:
            name = item["name"]
            item_id = item["id"]

            if "folder" in item:
                if st.button(f"📁 {name}", key=item_id):
                    st.session_state.current_folder = item_id
                    st.rerun()

# ================= CHAT =================
st.subheader("💬 Chat")

for msg in st.session_state.history:
    st.chat_message(msg["role"]).write(msg["content"])

query = st.chat_input()

if query:
    st.session_state.history.append({"role": "user", "content": query})

    prompt = f"""
Context:
{st.session_state.context}

Query:
{query}
"""

    model = choose_model(query, st.session_state.context)
    response = call_ai(prompt, model)

    st.session_state.history.append({"role": "assistant", "content": response})
    st.rerun()