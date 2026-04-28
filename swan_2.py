def run_swan_analysis(file, ext):
    doc = None
    paragraphs = []

    # -----------------------------
    # Extract text
    # -----------------------------
    if ext == ".docx":
        doc, paragraphs = extract_text_from_docx(file)
    elif ext == ".xlsx":
        doc, paragraphs = extract_text_from_xlsx(file)
    elif ext == ".pptx":
        doc, paragraphs = extract_text_from_pptx(file)
    else:
        return [], ["Unsupported file type."], [], []

    if not paragraphs:
        return [], ["No readable content was found in the file."], [], []

    strengths = []
    weaknesses = []
    actions = []
    next_steps = []

    text = " ".join(paragraphs).lower()

    # -----------------------------
    # STRUCTURE
    # -----------------------------
    if ext == ".docx":
        if count_headings_docx(doc) >= 1:
            strengths.append("You have used headings to organise your work clearly.")
        else:
            weaknesses.append("Your work would benefit from clear headings to guide the reader.")
            actions.append("Add headings to show where each new idea or section begins.")

        if has_bullets_docx(doc):
            strengths.append("Bullet points help make your ideas clear and easy to read.")
        else:
            weaknesses.append("Some sections could be clearer with bullet points.")
            actions.append("Use bullet points for lists or key ideas.")

    # -----------------------------
    # PARAGRAPH DEVELOPMENT
    # -----------------------------
    short_paras = find_short_paragraphs(paragraphs)
    if not short_paras:
        strengths.append("Your paragraphs are well-developed with enough detail.")
    else:
        weaknesses.append("Some paragraphs are very short and lack development.")
        actions.append("Choose one short paragraph and expand it with an example or explanation.")

    # -----------------------------
    # CONCLUSION CHECK
    # -----------------------------
    last_para = paragraphs[-1].lower()
    if any(phrase in last_para for phrase in ["in conclusion", "overall", "to sum up", "in summary"]):
        strengths.append("You have included a clear concluding section.")
    else:
        weaknesses.append("Your work ends abruptly without a clear conclusion.")
        actions.append("Add a short conclusion that summarises your key points.")

    # -----------------------------
    # SPELLING / REPETITION
    # -----------------------------
    spelling_issues = find_spelling_issues(paragraphs)
    if spelling_issues:
        weaknesses.append("Some words appear repeatedly and may be misspelt.")
        actions.append("Review repeated words and check their spelling or replace them with alternatives.")

    # -----------------------------
    # SENTENCE VARIETY
    # -----------------------------
    sentences = re.split(r"[.!?]", text)
    sentence_lengths = [len(s.split()) for s in sentences if len(s.split()) > 0]

    if sentence_lengths:
        avg_len = sum(sentence_lengths) / len(sentence_lengths)

        if avg_len < 10:
            weaknesses.append("Many sentences are very short, which makes the writing feel choppy.")
            actions.append("Combine some short sentences to create more complex ones.")
        elif avg_len > 25:
            weaknesses.append("Some sentences are very long and may be hard to follow.")
            actions.append("Split long sentences into two shorter ones to improve clarity.")
        else:
            strengths.append("You use a good mix of short and longer sentences.")

    # -----------------------------
    # VOCABULARY RICHNESS
    # -----------------------------
    words = re.findall(r"\b[a-zA-Z]{3,}\b", text)
    if words:
        unique_ratio = len(set(words)) / len(words)

        if unique_ratio > 0.4:
            strengths.append("Your vocabulary is varied and precise.")
        elif unique_ratio < 0.25:
            weaknesses.append("Your vocabulary is quite limited or repetitive.")
            actions.append("Experiment with more ambitious word choices.")

    # -----------------------------
    # LINKING WORDS
    # -----------------------------
    LINKERS = [
        "however", "therefore", "in addition", "furthermore", "moreover",
        "for example", "for instance", "consequently", "as a result"
    ]

    if any(l in text for l in LINKERS):
        strengths.append("You use linking words effectively to connect ideas.")
    else:
        weaknesses.append("Your writing lacks linking words to guide the reader.")
        actions.append("Use phrases like 'however', 'in addition', or 'for example' to show connections.")

    # -----------------------------
    # ARGUMENT STRUCTURE
    # -----------------------------
    ARG_MARKERS = ["because", "this shows", "this suggests", "therefore", "as a result"]

    if any(m in text for m in ARG_MARKERS):
        strengths.append("You explain your points with reasoning or evidence.")
    else:
        weaknesses.append("Some points are stated without explanation.")
        actions.append("After making a point, add a phrase like 'this shows that…' to explain it.")

    # -----------------------------
    # NEXT STEPS (always included)
    # -----------------------------
    next_steps.append("Read your work aloud to check that it flows logically.")
    next_steps.append("Compare your structure to a model answer to see how you could improve organisation.")
    next_steps.append("Ask a peer or teacher to highlight one unclear section, then rewrite it.")

    return strengths, weaknesses, actions, next_steps
