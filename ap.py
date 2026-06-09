import os
import streamlit as st
import google.generativeai as genai

from docx import Document
from langchain_core.documents import Document as LCDocument
from langchain_community.vectorstores import FAISS
from langchain_huggingface import HuggingFaceEmbeddings
from langchain_text_splitters import RecursiveCharacterTextSplitter
from docx.document import Document as _Document
from docx.oxml.text.paragraph import CT_P
from docx.oxml.table import CT_Tbl
from docx.table import _Cell, Table
from docx.text.paragraph import Paragraph

# =========================================================
# PAGE CONFIG
# =========================================================

st.set_page_config(
    page_title="HR Policy Helpdesk",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# =========================================================
# CSS — ORANGE / WHITE THEME
# =========================================================

st.markdown("""
<style>

@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600&display=swap');

:root {
    --orange:       #E85D04;
    --orange-light: #FFF0E8;
    --orange-mid:   #FFDCCA;
    --white:        #FFFFFF;
    --off-white:    #FAFAF9;
    --border:       #E8E3DE;
    --text:         #1A1A1A;
    --text-soft:    #6B6560;
    --text-muted:   #A8A09A;
    --radius:       6px;
}

html, body, .stApp {
    background-color: var(--white) !important;
    color: var(--text) !important;
    font-family: 'Inter', sans-serif !important;
    font-size: 14px !important;
}

#MainMenu, footer, header, .stDeployButton { display: none !important; }

.block-container {
    max-width: 960px !important;
    padding: 2rem 2.5rem 4rem !important;
}

/* ----- Top bar ----- */
.topbar {
    display: flex;
    align-items: center;
    gap: 1rem;
    border-bottom: 2px solid var(--orange);
    padding-bottom: 1rem;
    margin-bottom: 2.5rem;
}
.topbar-text {
    display: flex;
    flex-direction: column;
    gap: 2px;
}
.topbar-title {
    font-size: 1.25rem;
    font-weight: 600;
    color: var(--text);
    letter-spacing: -0.01em;
}
.topbar-sub {
    font-size: 12px;
    color: var(--text-muted);
    font-weight: 400;
}

/* ----- Section labels ----- */
.sec {
    font-size: 10px;
    font-weight: 600;
    letter-spacing: 0.15em;
    text-transform: uppercase;
    color: var(--orange);
    display: block;
    margin-bottom: 0.75rem;
}

/* ----- Rule line ----- */
.rule {
    height: 1px;
    background: var(--border);
    margin: 1.6rem 0;
}

/* ----- Chat input ----- */
.stChatInput textarea {
    background: var(--off-white) !important;
    border: 1px solid var(--border) !important;
    border-radius: var(--radius) !important;
    color: var(--text) !important;
    font-size: 13px !important;
    font-family: 'Inter', sans-serif !important;
}
.stChatInput textarea:focus {
    border-color: var(--orange) !important;
    box-shadow: 0 0 0 3px rgba(232,93,4,0.1) !important;
    outline: none !important;
}
.stChatInput button {
    background: var(--orange) !important;
    border-radius: var(--radius) !important;
    color: #fff !important;
}
.stChatInput button:hover {
    background: #C94E00 !important;
}

/* ----- Chat messages ----- */
.stChatMessage {
    background: var(--off-white) !important;
    border: 1px solid var(--border) !important;
    border-radius: var(--radius) !important;
    padding: 1rem 1.25rem !important;
    margin-bottom: 0.75rem !important;
}
[data-testid="stChatMessageContent"] p {
    color: var(--text) !important;
    font-size: 13px !important;
    line-height: 1.7 !important;
}

/* User bubble */
[data-testid="stChatMessage"]:has([data-testid="stChatMessageAvatarUser"]) {
    background: var(--orange-light) !important;
    border-color: var(--orange-mid) !important;
}

/* ----- Expander (Source Sections) ----- */
.streamlit-expanderHeader {
    background: var(--off-white) !important;
    border: 1px solid var(--border) !important;
    border-radius: var(--radius) !important;
    font-size: 11px !important;
    font-weight: 600 !important;
    letter-spacing: 0.1em !important;
    text-transform: uppercase !important;
    color: var(--text-soft) !important;
    padding: 0.6rem 1rem !important;
}
.streamlit-expanderHeader:hover {
    border-color: var(--orange) !important;
    color: var(--orange) !important;
}
.streamlit-expanderContent {
    border: 1px solid var(--border) !important;
    border-top: none !important;
    border-radius: 0 0 var(--radius) var(--radius) !important;
    padding: 1rem 1.25rem !important;
    background: var(--white) !important;
}

/* Source section tags */
.source-tag {
    display: inline-block;
    background: var(--orange-light);
    border: 1px solid var(--orange-mid);
    color: var(--orange);
    font-size: 10px;
    font-weight: 600;
    letter-spacing: 0.08em;
    text-transform: uppercase;
    border-radius: 4px;
    padding: 3px 8px;
    margin-bottom: 6px;
}
.source-content {
    font-size: 12px;
    color: var(--text-soft);
    line-height: 1.6;
    border-left: 2px solid var(--orange-mid);
    padding-left: 10px;
    margin-bottom: 1rem;
}

/* ----- Spinner ----- */
.stSpinner > div {
    border-top-color: var(--orange) !important;
}

/* ----- Scrollbar ----- */
::-webkit-scrollbar { width: 4px; }
::-webkit-scrollbar-track { background: transparent; }
::-webkit-scrollbar-thumb { background: var(--border); border-radius: 4px; }

/* ----- API status ----- */
.api-status-ok {
    display: flex;
    align-items: center;
    gap: 6px;
    margin-top: 0.5rem;
    font-size: 11px;
    color: #2D7A3A;
}
.api-status-empty {
    display: flex;
    align-items: center;
    gap: 6px;
    margin-top: 0.5rem;
    font-size: 11px;
    color: var(--text-muted);
}

/* ----- API key input ----- */
.stTextInput input {
    background: var(--off-white) !important;
    border: 1px solid var(--border) !important;
    border-radius: var(--radius) !important;
    color: var(--text) !important;
    font-size: 13px !important;
}
.stTextInput input:focus {
    border-color: var(--orange) !important;
    box-shadow: 0 0 0 3px rgba(232,93,4,0.1) !important;
    outline: none !important;
}


/* ----- Empty state ----- */
.empty-state {
    border: 1px dashed var(--border);
    border-radius: var(--radius);
    padding: 3.5rem 2rem;
    text-align: center;
    color: var(--text-muted);
    font-size: 13px;
    line-height: 2;
    background: var(--off-white);
    margin-top: 1rem;
}

</style>
""", unsafe_allow_html=True)

# =========================================================
# TOPBAR
# =========================================================

# Encode logo as base64 to embed inline
import base64
from pathlib import Path

logo_path = Path("LOGO_BLACK.png")
logo_b64 = base64.b64encode(logo_path.read_bytes()).decode()

st.markdown(f"""
<div class="topbar">
    <div style="margin-left:auto;">
        <img src="data:image/png;base64,{logo_b64}"
             style="height:65px; width:auto; display:block;" />
    </div>
</div>
""", unsafe_allow_html=True)

st.markdown("""
<div class="topbar">
    <div class="topbar-text">
        <span class="topbar-title">HR Policy Helpdesk</span>
        <span class="topbar-sub">Ask any question related to company HR policies</span>
    </div>
    <div style="margin-left:auto;">
        <img src="data:image/png;base64,{logo_b64}"
             style="height:65px; width:auto; display:block;" />
    </div>
</div>
""", unsafe_allow_html=True)

st.markdown('<span class="sec">00 — API Configuration</span>', unsafe_allow_html=True)
GEMINI_API_KEY = st.text_input(
    "Gemini API Key",
    type="password",
    placeholder="Enter your Gemini API key",
    label_visibility="collapsed"
)
if GEMINI_API_KEY:
    st.markdown("""
    <div class="api-status-ok">
        <span>●</span> <span>Key entered — ready to search</span>
    </div>
    """, unsafe_allow_html=True)
    genai.configure(api_key=GEMINI_API_KEY)
    model = genai.GenerativeModel("gemini-2.5-flash")
else:
    st.markdown("""
    <div class="api-status-empty">
        <span>○</span> <span>Enter your Gemini API key above &nbsp;·&nbsp;
        <a href="https://aistudio.google.com/app/apikey" target="_blank"
           style="color:#E85D04; text-decoration:none;">Get one free ↗</a></span>
    </div>
    """, unsafe_allow_html=True)

st.markdown('<div class="rule"></div>', unsafe_allow_html=True)

# =========================================================
# CONFIG
# =========================================================

DOCX_FILE = "hr_policy.docx"
VECTOR_DB = "vector_store"
# =========================================================
# EXTRACT WORD SECTIONS
# =========================================================

def iter_block_items(parent):
    if isinstance(parent, _Document):
        parent_elm = parent.element.body
    elif isinstance(parent, _Cell):
        parent_elm = parent._tc
    else:
        raise ValueError("Unsupported parent type")
    for child in parent_elm.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, parent)
        elif isinstance(child, CT_Tbl):
            yield Table(child, parent)

def is_heading(block):
    if not isinstance(block, Paragraph):
        return False
    text = block.text.strip()
    if not text or len(text) > 100:
        return False
    if "Heading" in block.style.name:
        return True
    bold_text = "".join([run.text for run in block.runs if run.bold])
    if len(bold_text.strip()) == len(text):
        return True
    return False

def extract_sections(docx_path):
    doc = Document(docx_path)
    sections = []
    current_heading = "General Introduction"
    current_content = []
    for block in iter_block_items(doc):
        if isinstance(block, Paragraph):
            text = block.text.strip()
            if not text:
                continue
            if is_heading(block):
                if current_content:
                    sections.append({"section": current_heading, "content": "\n".join(current_content)})
                current_heading = text
                current_content = []
            else:
                current_content.append(text)
        elif isinstance(block, Table):
            table_text = []
            for row in block.rows:
                row_data = [cell.text.strip().replace('\n', ' ') for cell in row.cells]
                table_text.append(" | ".join(row_data))
            if table_text:
                current_content.append("\n" + "\n".join(table_text) + "\n")
    if current_content:
        sections.append({"section": current_heading, "content": "\n".join(current_content)})
    return sections

# =========================================================
# VECTOR DB
# =========================================================

@st.cache_resource
def build_or_load_db():
    embeddings = HuggingFaceEmbeddings(model_name="sentence-transformers/all-MiniLM-L6-v2")
    if os.path.exists(VECTOR_DB):
        return FAISS.load_local(VECTOR_DB, embeddings, allow_dangerous_deserialization=True)
    sections = extract_sections(DOCX_FILE)
    text_splitter = RecursiveCharacterTextSplitter(chunk_size=800, chunk_overlap=150, separators=["\n\n", "\n", " ", ""])
    docs = []
    for sec in sections:
        for chunk in text_splitter.split_text(sec["content"]):
            docs.append(LCDocument(page_content=chunk, metadata={"section": sec["section"]}))
    db = FAISS.from_documents(docs, embeddings)
    db.save_local(VECTOR_DB)
    return db

# =========================================================
# RETRIEVAL + LLM
# =========================================================

def get_answer(question, db):
    docs = db.similarity_search(question, k=6)
    context = "\n\n".join([f"SECTION: {d.metadata['section']}\n{d.page_content}" for d in docs])
    prompt = f"""
You are an expert HR Policy Assistant.

Your task is to answer the user's question using ONLY the provided HR Policy text below.

CRITICAL INSTRUCTIONS:
1. The policy text contains tables formatted with pipe characters (|).
2. You must carefully read the rows and columns of these tables to find numerical values, thresholds, and schedules.
3. If the answer requires looking at a range (e.g., "above 20,000"), evaluate the numbers carefully.
4. If the answer is not available in the policy, say: "I could not find this information in the HR Policy."

HR POLICY:
{context}

QUESTION:
{question}
"""
    response = model.generate_content(prompt)
    return response.text, docs

# =========================================================
# UI
# =========================================================

db = build_or_load_db()

if not GEMINI_API_KEY:
    st.stop()

st.markdown('<span class="sec">Ask a Question</span>', unsafe_allow_html=True)

question = st.chat_input("Type your HR policy question here...")

if question:
    with st.spinner("Searching policy..."):
        answer, docs = get_answer(question, db)

    with st.chat_message("user"):
        st.write(question)

    with st.chat_message("assistant"):
        st.write(answer)

    st.markdown('<div class="rule"></div>', unsafe_allow_html=True)
    st.markdown('<span class="sec">Source Sections</span>', unsafe_allow_html=True)

    with st.expander("View matched policy sections"):
        for doc in docs:
            st.markdown(f'<div class="source-tag">{doc.metadata["section"]}</div>', unsafe_allow_html=True)
            st.markdown(f'<div class="source-content">{doc.page_content}</div>', unsafe_allow_html=True)

else:
    st.markdown("""
    <div class="empty-state">
        Type a question below to search the HR policy.<br>
        <strong style="color:#E85D04;">Examples:</strong> &nbsp;
        What is the leave policy? &nbsp;·&nbsp;
        How is overtime calculated? &nbsp;·&nbsp;
        What are the reimbursement limits?
    </div>
    """, unsafe_allow_html=True)
