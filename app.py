import os
import re
import msal
import uuid
import pandas as pd
import requests
import streamlit as st
from anthropic import Anthropic
from docx import Document
from pypdf import PdfReader

# ================= CONFIG =================
st.set_page_config(page_title="Enterprise AI Copilot", layout="wide")

st.title("🚀 Enterprise AI Copilot (SSO + RAG + Memory)")
st.caption("Production AI System")

# ================= ENV =================
ANTHROPIC_API_KEY = os.getenv("ANTHROPIC_API_KEY")
CLIENT_ID = os.getenv("CLIENT_ID")
TENANT_ID = os.getenv("TENANT_ID")

AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPES = ["User.Read"]

if not ANTHROPIC_API_KEY:
    st.error("Missing API Key")
    st.stop()

client = Anthropic(api_key=ANTHROPIC_API_KEY)

# ================= SESSION =================
if "auth" not in st.session_state:
    st.session_state.auth = False
    st.session_state.user = None

if "history" not in st.session_state:
    st.session_state.history = []

if "memory" not in st.session_state:
    st.session_state.memory = []

if "vector_store" not in st.session_state:
    st.session_state.vector_store = []

# ================= SSO LOGIN =================
def microsoft_login():
    app = msal.PublicClientApplication(CLIENT_ID, authority=AUTHORITY)

    flow = app.initiate_device_flow(scopes=SCOPES)

    if "user_code" not in flow:
        st.error("Login failed")
        return

    st.info(f"Go to {flow['verification_uri']}")
    st.info(f"Enter code: {flow['user_code']}")

    result = app.acquire_token_by_device_flow(flow)

    if "access_token" in result:
        st.session_state.auth = True
        st.session_state.user = result["id_token_claims"]["name"]
        st.rerun()
    else:
        st.error("Login failed")

# LOGIN UI
if not st.session_state.auth:
    st.subheader("🔐 Microsoft Login")
    if st.button("Login with Microsoft"):
        microsoft_login()
    st.stop()

st.success(f"✅ Logged in as {st.session_state.user}")

# ================= FILE PARSER =================
def extract_file(file):
    try:
        if file.name.endswith(".docx"):
            return "\n".join(p.text for p in Document(file).paragraphs)

        if file.name.endswith(".pdf"):
            return "\n".join(
                p.extract_text() for p in PdfReader(file).pages if p.extract_text()
            )

        if file.name.endswith((".xlsx", ".xls")):
            df = pd.read_excel(file)
            return df.to_string()

        return file.read().decode("utf-8", "ignore")

    except:
        return ""

# ================= RAG =================
def chunk_text(text, size=500):
    return [text[i:i+size] for i in range(0, len(text), size)]

def store_chunks(text):
    chunks = chunk_text(text)
    st.session_state.vector_store.extend(chunks)

def retrieve_context(query):
    relevant = []
    for chunk in st.session_state.vector_store:
        if any(word in chunk.lower() for word in query.lower().split()):
            relevant.append(chunk)

    return "\n".join(relevant[:5])

# ================= PROMPT ENGINE =================
def build_prompt(query):
    rag_context = retrieve_context(query)
    conversation = "\n".join(
        [f"{m['role']}: {m['content']}" for m in st.session_state.memory[-5:]]
    )

    return f"""
You are an enterprise AI copilot.

Conversation Memory:
{conversation}

Retrieved Context:
{rag_context}

User Query:
{query}

Instructions:
- Answer completely
- Do not stop mid response
- If needed, continue automatically
"""

# ================= AI ENGINE =================
def call_ai(prompt):
    full_response = ""
    current_prompt = prompt

    for _ in range(5):
        try:
            res = client.messages.create(
                model="claude-sonnet-4-6",
                max_tokens=4000,
                messages=[{"role": "user", "content": current_prompt}]
            )

            chunk = res.content[0].text
            full_response += chunk

            if len(chunk) < 200:
                break

            current_prompt = "Continue from where you stopped."

        except Exception as e:
            return f"Error: {e}"

    return full_response

# ================= FILE UPLOAD =================
st.subheader("📂 Upload Knowledge")

files = st.file_uploader("Upload documents", accept_multiple_files=True)

if files:
    for f in files:
        content = extract_file(f)
        store_chunks(content)

    st.success("✅ Knowledge Base Updated")

# ================= CHAT =================
st.subheader("💬 Copilot Chat")

for msg in st.session_state.history:
    st.chat_message(msg["role"]).write(msg["content"])

query = st.chat_input("Ask anything...")

if query:
    st.session_state.history.append({"role": "user", "content": query})
    st.session_state.memory.append({"role": "user", "content": query})

    prompt = build_prompt(query)

    response = call_ai(prompt)

    st.session_state.history.append({"role": "assistant", "content": response})
    st.session_state.memory.append({"role": "assistant", "content": response})

    st.rerun()