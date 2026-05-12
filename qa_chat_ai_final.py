import streamlit as st
import os, time, re, requests, io, json
import pandas as pd
from dotenv import load_dotenv
from anthropic import Anthropic
from docx import Document
from pypdf import PdfReader
import msal

# ================= CONFIG =================
st.set_page_config(page_title="Enterprise QA Copilot", layout="wide")

st.title("🚀 Enterprise QA Copilot (Claude + GPT + Copilot)")
st.caption("Chat + Files + Excel Export + Memory + Analytics")

# ================= ENV =================
load_dotenv()
client = Anthropic(api_key=os.getenv("ANTHROPIC_API_KEY"))

# ================= SESSION =================
if "auth" not in st.session_state:
    st.session_state.auth = False

if "usage_log" not in st.session_state:
    st.session_state.usage_log = []

if "context" not in st.session_state:
    st.session_state.context = ""

if "history" not in st.session_state:
    st.session_state.history = []

# ✅ Explorer state
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
CLIENT_ID = "57a468ca-7d2c-4925-a21f-a90262fe5fc5"
TENANT_ID = "c0fcbffe-4fca-4c59-9a2c-efdcda8387fb"

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
            return "\n".join(p.extract_text() for p in PdfReader(file).pages if p.extract_text())

        if name.endswith((".xlsx", ".xls")):
            df = pd.read_excel(file, engine="openpyxl")
            return df.to_string()

        return file.read().decode("utf-8", "ignore")
    except:
        return ""

# ================= URL =================
def load_from_url(url):
    try:
        headers = {"User-Agent": "Mozilla/5.0"}
        r = requests.get(url, headers=headers)

        if r.status_code != 200:
            return ""

        if "text" in r.headers.get("Content-Type", ""):
            return r.text[:8000]

        return r.content[:8000].decode("utf-8", "ignore")
    except:
        return ""

# ================= GRAPH API =================
def get_graph_token():
    app = msal.PublicClientApplication(CLIENT_ID, authority=AUTHORITY)
    flow = app.initiate_device_flow(scopes=SCOPES)

    if "user_code" not in flow:
        st.error("❌ Failed to start login")
        return None

    st.info(f"👉 Open: {flow['verification_uri']}")
    st.info(f"👉 Enter Code: {flow['user_code']}")

    result = app.acquire_token_by_device_flow(flow)

    if "access_token" in result:
        return result["access_token"]

    st.error("Login failed")
    return None


def list_drive_items(token, folder_id=None):
    headers = {"Authorization": f"Bearer {token}"}

    if folder_id:
        url = f"https://graph.microsoft.com/v1.0/me/drive/items/{folder_id}/children"
    else:
        url = "https://graph.microsoft.com/v1.0/me/drive/root/children"

    res = requests.get(url, headers=headers)

    if res.status_code != 200:
        return None

    return res.json().get("value", [])


def download_onedrive_file(token, file_id):
    headers = {"Authorization": f"Bearer {token}"}
    url = f"https://graph.microsoft.com/v1.0/me/drive/items/{file_id}/content"

    res = requests.get(url, headers=headers)

    if res.status_code != 200:
        return None

    return res.content

# ================= MEMORY =================
def get_context_window(history, max_chars=6000):
    combined = ""
    selected = []
    for msg in reversed(history):
        block = str(msg)
        if len(combined) + len(block) < max_chars:
            selected.insert(0, msg)
            combined += block
        else:
            break
    return selected

# ================= AI ENGINE =================
def call_ai_complete(prompt, model, max_loops=5):
    final_output = ""
    current_prompt = prompt
    total_time = 0
    total_cost = 0

    for _ in range(max_loops):
        start = time.time()

        res = client.messages.create(
            model=model,
            max_tokens=4000,
            messages=[{"role": "user", "content": mask(current_prompt)}]
        )

        chunk = res.content[0].text
        final_output += chunk

        end = time.time()
        elapsed = round(end - start, 2)

        tokens = len(current_prompt.split()) + len(chunk.split())
        cost = tokens * 0.000001

        total_time += elapsed
        total_cost += cost

        st.session_state.usage_log.append({
            "model": model,
            "time": elapsed,
            "cost": round(cost, 6)
        })

        if len(chunk) < 200:
            break

        current_prompt = "Continue exactly from where you stopped. Do not repeat."

    return final_output, {
        "model": model,
        "time": round(total_time, 2),
        "cost": round(total_cost, 6)
    }

# ================= FILE UPLOAD =================
st.subheader("📂 Upload Files")
uploaded_files = st.file_uploader("Upload documents", accept_multiple_files=True)

if uploaded_files:
    content = ""
    for f in uploaded_files:
        content += extract_file(f) + "\n\n"
    st.session_state.context = content
    st.success("✅ Files added")

# ================= ENTERPRISE EXPLORER =================
st.subheader("☁️ Enterprise File Explorer")

if st.button("🔐 Connect to Microsoft"):
    token = get_graph_token()
    if token:
        st.session_state.graph_token = token
        st.success("✅ Connected")

if "graph_token" in st.session_state:

    items = list_drive_items(
        st.session_state.graph_token,
        st.session_state.current_folder
    )

    if items:

        if st.session_state.folder_stack:
            if st.button("⬅️ Back"):
                st.session_state.current_folder = st.session_state.folder_stack.pop()
                st.rerun()

        for item in items:

            name = item["name"]
            item_id = item["id"]

            if "folder" in item:

                if st.button(f"📁 {name}", key=item_id):
                    st.session_state.folder_stack.append(st.session_state.current_folder)
                    st.session_state.current_folder = item_id
                    st.rerun()

            else:
                col1, col2 = st.columns([4,1])

                with col1:
                    st.write(f"📄 {name}")

                with col2:
                    if st.button("Load", key=f"load_{item_id}"):

                        content = download_onedrive_file(
                            st.session_state.graph_token,
                            item_id
                        )

                        if content:
                            try:
                                text = content.decode("utf-8", "ignore")
                            except:
                                text = str(content)

                            st.session_state.context += text[:8000]
                            st.success(f"✅ {name} loaded")

# ================= CHAT =================
st.subheader("💬 Conversation")

for msg in st.session_state.history:
    role = "user" if msg["role"] == "user" else "assistant"
    st.chat_message(role).write(msg["content"])

query = st.chat_input("Type your message...")

if query:
    st.session_state.history.append({"role": "user", "content": query})

    trimmed_history = get_context_window(st.session_state.history)

    prompt = f"""
Context:
{st.session_state.context}

Conversation:
{trimmed_history}

User Query:
{query}
"""

    model = choose_model(query, st.session_state.context)
    answer, meta = call_ai_complete(prompt, model)

    st.session_state.history.append({"role": "assistant", "content": answer})
    st.session_state.last_answer = answer
    st.session_state.last_meta = meta

    st.rerun()

# ================= OUTPUT =================
if "last_answer" in st.session_state:

    st.subheader("🧠 Response")
    st.write(st.session_state.last_answer)

    try:
        js = st.session_state.last_answer[
            st.session_state.last_answer.find("["):
            st.session_state.last_answer.rfind("]") + 1
        ]

        df = pd.DataFrame(json.loads(js))

        st.subheader("📊 Extracted Table")
        st.dataframe(df)

        output = io.BytesIO()
        df.to_excel(output, index=False)

        st.download_button("📥 Download Excel", output.getvalue(), "output.xlsx")

    except:
        pass

# ================= DASHBOARD =================
st.subheader("📊 Usage Dashboard")

if st.session_state.usage_log:

    df = pd.DataFrame(st.session_state.usage_log)

    st.bar_chart(df["model"].value_counts())
    st.line_chart(df["cost"].cumsum())
    st.dataframe(df)