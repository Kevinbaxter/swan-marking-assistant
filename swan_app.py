import streamlit as st
from docx import Document
import re
from collections import Counter

SHORT_PARAGRAPH_LIMIT = 260

def extract_text(doc):
    paragraphs = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
    return paragraphs

def count_headings(doc):
    return sum(1 for p in doc.paragraphs if p.style.name.startswith("Heading"))

def has_bullets(doc):
    return any("List" in p.style.name for p in doc.paragraphs)

def find_short_paragraphs(paragraphs):
    return [p for p in paragraphs if len(p) < SHORT_PARAGRAPH_LIMIT]

def find_spelling_issues(paragraphs):
    words = re.findall(r"\b[a-zA-Z]{3,}\b", " ".join(paragraphs).lower())
    counts = Counter(words)
    return [w for w, c in counts.items() if c >= 3]

st.title("🦢 SWAN Marking Assistant – Word")

uploaded = st.file_uploader("Upload a Word document (.docx)", type="docx")

if uploaded:
    doc = Document(uploaded)
    paragraphs = extract_text(doc)

    strengths = []
    weaknesses = []
    actions = []

    if count_headings(doc) >= 1:
        strengths.append("You have used headings to organise your work.")

    if has_bullets(doc):
        strengths.append("Bullet points help make your ideas clear and easy to read.")

    if not find_short_paragraphs(paragraphs):
        strengths.append("Most of your paragraphs are developed with enough detail.")

    if find_short_paragraphs(paragraphs):
        weaknesses.append("One paragraph is very short and lacks detail.")
        actions.append("Action: Rewrite the short paragraph and add one example.")

    if len(paragraphs[-1]) < SHORT_PARAGRAPH_LIMIT:
        weaknesses.append("There is no clear conclusion at the end of your work.")
        actions.append("Action: Add a short conclusion that links back to the question.")

    if find_spelling_issues(paragraphs):
        weaknesses.append("The same word is misspelt several times.")
        actions.append("Action: Find and fix the spelling errors in your work.")

    st.subheader("Strengths")
    for s in strengths[:3]:
        st.write("•", s)

    st.subheader("Weaknesses")
    for w in weaknesses[:3]:
        st.write("•", w)

    st.subheader("Next Steps")
    for a in actions[:3]:
        st.write("•", a)
