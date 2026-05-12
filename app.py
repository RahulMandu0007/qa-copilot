import os
import re
import msal
import pandas as pd
import numpy as np
import streamlit as st
from anthropic import Anthropic
from docx import Document
from pypdf import PdfReader
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity

# ================= CONFIG =================
st.set_page_config(page_title="Enterprise Copilot Ultimate", layout="wide")
st.title("🚀 Enterprise Copilot Ultimate (FAISS + Tools + Graph)")

# ================= ENV =================
ANTHROPIC_API_KEY = os.getenv("ANTHROPIC_API_KEY")
CLIENT_ID = os.getenv("CLIENT_ID")
TENANT_ID = os.getenv("TENANT_ID")

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

if "documents" not in st.session_state:
    st.session_state.documents = []

# ================= LOGIN =================
if not st.session_state.auth:
    st.subheader("🔐 Login")
    user = st.text_input("Username")
    pwd = st.text_input("Password", type="password")

    if st.button("Login"):
        if user == os.getenv("APP_USER", "admin") and pwd == os.getenv("APP_PASS", "admin123"):
            st.session_state.auth = True
            st.session_state.user = user
            st.rerun()
        else:
            st.error("Invalid credentials")

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
            df = pd.read_excel(file, engine="openpyxl")
            st.dataframe(df)
            return df.to_csv(index=False)

        return file.read().decode("utf-8", "ignore")

    except:
        return ""

# ================= VECTOR SEARCH =================
def build_vector_store(texts):
    vectorizer = TfidfVectorizer(stop_words="english")
    vectors = vectorizer.fit_transform(texts)
    return vectorizer, vectors

def retrieve_context(query):
    if not st.session_state.documents:
        return ""

    texts = st.session_state.documents
    vectorizer, vectors = build_vector_store(texts)

    query_vec = vectorizer.transform([query])
    similarity = cosine_similarity(query_vec, vectors).flatten()

    top_idx = similarity.argsort()[-5:][::-1]
    return "\n".join([texts[i] for i in top_idx])

# ================= TOOL ENGINE =================
def detect_tool(query):
    q = query.lower()

    if any(x in q for x in ["analyze", "summary", "statistics", "excel"]):
        return "data_analysis"

    if any(x in q for x in ["table", "format"]):
        return "table"

    return "none"

def run_tool(tool, context):
    if tool == "data_analysis":
        try:
            from io import StringIO
            df = pd.read_csv(StringIO(context))
            return df.describe().to_string()
        except:
            return None

    return None

# ================= PROMPT =================
def build_prompt(query):
    context = retrieve_context(query)

    memory = "\n".join(
        [f"{m['role']}: {m['content']}" for m in st.session_state.memory[-5:]]
    )

    return f"""
You are Enterprise Copilot.

Memory:
{memory}

Context:
{context}

User:
{query}

Rules:
- Provide FULL complete answer
- Continue if needed
- Use tables when relevant
- Be structured and clear
"""

# ================= STREAMING AI =================
def call_ai(prompt):
    placeholder = st.empty()
    full_text = ""

    for _ in range(5):
        try:
            res = client.messages.create(
                model="claude-sonnet-4-6",
                max_tokens=4000,
                messages=[{"role": "user", "content": prompt}]
            )

            text = res.content[0].text
            full_text += text

            placeholder.markdown(full_text)

            if len(text) < 200:
                break

            prompt = "Continue from previous response"

        except Exception as e:
            return f"Error: {e}"

    return full_text

# ================= FILE UPLOAD =================
st.subheader("📂 Upload Knowledge")

files = st.file_uploader("Upload files", accept_multiple_files=True)

if files:
    for f in files:
        text = extract_file(f)
        chunks = [text[i:i+600] for i in range(0, len(text), 600)]
        st.session_state.documents.extend(chunks)

    st.success("✅ Knowledge indexed")

# ================= CHAT =================
st.subheader("💬 Copilot Chat")

for msg in st.session_state.history:
    st.chat_message(msg["role"]).write(msg["content"])

query = st.chat_input("Ask anything...")

if query:
    st.session_state.history.append({"role": "user", "content": query})
    st.session_state.memory.append({"role": "user", "content": query})

    tool = detect_tool(query)

    if tool != "none":
        context = retrieve_context(query)
        tool_result = run_tool(tool, context)

        if tool_result:
            response = tool_result
        else:
            prompt = build_prompt(query)
            response = call_ai(prompt)
    else:
        prompt = build_prompt(query)
        response = call_ai(prompt)

    # ✅ Structured output rendering
    try:
        if response.strip().startswith("[{"):
            import json
            df = pd.DataFrame(json.loads(response))
            st.dataframe(df)
    except:
        pass

    st.session_state.history.append({"role": "assistant", "content": response})
    st.session_state.memory.append({"role": "assistant", "content": response})

    st.rerun()
