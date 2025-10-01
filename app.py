# app.py
from __future__ import annotations

import io
import pathlib
import re
import textwrap
from datetime import datetime, timezone
from zoneinfo import ZoneInfo
from typing import List, Tuple
from zipfile import ZipFile

import streamlit as st
from docx import Document
from docx.text.run import Run
from openai import OpenAI
import difflib
from lxml import etree as ET  # lxml kommer via python-docx-avhengighet

# -----------------------------
# Build timestamp (Europe/Oslo)
# -----------------------------
def _build_time_oslo() -> str:
    try:
        ts = pathlib.Path(__file__).stat().st_mtime
        dt = datetime.fromtimestamp(ts, tz=timezone.utc).astimezone(ZoneInfo("Europe/Oslo"))
    except Exception:
        dt = datetime.now(timezone.utc).astimezone(ZoneInfo("Europe/Oslo"))
    return dt.strftime("%Y-%m-%d %H:%M (%Z)")

BUILD_TIME_LOCAL = _build_time_oslo()

# -----------------------------
# Streamlit page config & badge
# -----------------------------
st.set_page_config(page_title="MedLang Improver — ekte Track Changes", page_icon="🩺", layout="centered")

st.markdown(
    f"""
    <div style="
        position: fixed; top: 8px; left: 12px;
        padding: 4px 10px; border-radius: 8px;
        font-size: 12px; font-family: ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, monospace;
        background: rgba(0,0,0,0.06); backdrop-filter: blur(2px);
        z-index: 1000;">
        Build: {BUILD_TIME_LOCAL}
    </div>
    """,
    unsafe_allow_html=True,
)

st.title("🩺 Språkforbedrer for medisinske artikler — med Ekte Track Changes")
st.markdown(
    "Last opp Word eller lim inn tekst. Få **ren forbedret** fil og en **.docx med ekte Spor endringer** "
    "(Word: Godta/Avvis endringer)."
)

# -----------------------------
# OpenAI client
# -----------------------------
OPENAI_API_KEY = st.secrets.get("OPENAI_API_KEY", None)
if not OPENAI_API_KEY:
    st.warning("Mangler OPENAI_API_KEY i Secrets (App → ⋮ → Settings → Secrets). Appen kan ikke forbedre tekst.")
client = OpenAI(api_key=OPENAI_API_KEY) if OPENAI_API_KEY else None

# -----------------------------
# UI: input
# -----------------------------
tab1, tab2 = st.tabs(["📄 Word-fil (.docx)", "✍️ Lim inn tekst"])
uploaded_text: str | None = None
source_docx_bytes: bytes | None = None

with tab1:
    up = st.file_uploader("Last opp .docx", type=["docx"])
    if up is not None:
        try:
            source_docx_bytes = up.read()
            doc = Document(io.BytesIO(source_docx_bytes))
            uploaded_text = "\n".join(p.text for p in doc.paragraphs)
        except Exception as e:
            st.error(f"Kunne ikke lese Word-filen: {e}")

with tab2:
    pasted = st.text_area("Lim inn tekst her", height=300, placeholder="Lim inn manus her …")
    if pasted and not uploaded_text:
        uploaded_text = pasted

colA, colB = st.columns(2)
with colA:
    tone = st.selectbox(
        "Tone/retning",
        ["Nøytral faglig", "Mer konsis", "Mer formell", "For legfaglig publikum"],
        help="Velg hvordan teksten skal forbedres.",
    )
with colB:
    model_name = st.selectbox(
        "Modell",
        ["gpt-4o-mini", "gpt-4o"],
        help="Mini er rimelig og rask; gpt-4o kan gi litt høyere kvalitet.",
    )

st.caption("Tips: Del store manus i seksjoner (Introduksjon, Metode, Resultater, osv.) for bedre kontroll.")

# -----------------------------
# Prompts
# -----------------------------
GOALS = {
    "Nøytral faglig": (
        "Forbedre klarhet, flyt og grammatikk i vitenskapelig medisinsk tekst. "
        "Behold betydning, referanser og tall uendret. Unngå nye påstander."
    ),
    "Mer konsis": (
        "Forbedre klarhet og gjør teksten mer konsis uten å endre faglig innhold. "
        "Behold referanser og tall. Fjern fyllord."
    ),
    "Mer formell": (
        "Hev formalitetsnivå, vitenskapelig stil, presis terminologi og grammatikk. "
        "Ikke legg til nye påstander."
    ),
    "For legfaglig publikum": (
        "Forenkle språket lett og forklar forkortelser der det er naturlig, men behold presisjon. "
        "Ikke endre data eller resultater."
    ),
}

SYSTEM_PROMPT = (
    "Du er språkredaktør for medisinske manus. "
    "Respekter faglig innhold, data, referanser (f.eks. [12], (Smith 2020), DOI), og numerikk. "
    "Ikke legg til, fjern eller omtolk resultater. "
    "Ikke endre referanseformatering."
)

# -----------------------------
# Helpers: forbedring
# -----------------------------
def improve_text(text: str, mode: str, model_name: str) -> str:
    instr = GOALS.get(mode, GOALS["Nøytral faglig"])
    if not client:
        raise RuntimeError("OPENAI_API_KEY mangler – kan ikke kalle modellen.")
    resp = client.chat.completions.create(
        model=model_name,
        messages=[
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user", "content": f"{instr}\n\nTekst:\n{text}"},
        ],
        temperature=0.2,
    )
    return resp.choices[0].message.content.strip()

def make_docx_from_text(text: str, heading: str | None = None) -> bytes:
    d = Document()
    if heading:
        d.add_heading(heading, level=1)
    for para in text.split("\n"):
        d.add_paragraph(para)
    bio = io.BytesIO()
    d.save(bio)
    return bio.getvalue()

# -----------------------------
# Diff & ekte Track Changes (w:ins / w:del)
# -----------------------------
TOKEN_RE = re.compile(r"\s+|\w+|[^\w\s]", re.UNICODE)

def tokenize_keep_ws(s: str) -> List[str]:
    return TOKEN_RE.findall(s)

def diff_tokens(a: List[str], b: List[str]) -> List[Tuple[str, str]]:
    """
    Returnerer liste av (op, text) der op ∈ {'=', '+', '-'}
    '=': uendret, '+': innsatt i b, '-': slettet fra a
    """
    sm = difflib.SequenceMatcher(a=a, b=b, autojunk=False)
    out: List[Tuple[str, str]] = []
    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        if tag == "equal":
            txt = "".join(a[i1:i2])
            if txt:
                out.append(("=", txt))
        elif tag == "insert":
            txt = "".join(b[j1:j2])
            if txt:
                out.append(("+", txt))
        elif tag == "delete":
            txt = "".join(a[i1:i2])
            if txt:
                out.append(("-", txt))
        elif tag == "replace":
            d_txt = "".join(a[i1:i2])
            i_txt = "".join(b[j1:j2])
            if d_txt:
                out.append(("-", d_txt))
            if i_txt:
                out.append(("+", i_txt))
    return out

def make_tracked_changes_docx(original_text: str, improved_text: str, author: str = "ChatGPT") -> bytes:
    """
    Lager en .docx der endringer er merket som ekte revisjoner:
    - Innsatt: <w:ins w:author="..." w:date="..." w:id="..."><w:r><w:t>...</w:t></w:r></w:ins>
    - Slettet: <w:del w:author="..." w:date="..." w:id="..."><w:r><w:delText>...</w:delText></w:r></w:del>
    Vi bygger XML for word/document.xml og pakker det inn i et base-docx fra python-docx.
    """
    # 1) Lag base DOCX for å få standard styles, props, sectPr osv.
    base_doc = Document()
    base_bio = io.BytesIO()
    base_doc.save(base_bio)
    base_bytes = base_bio.getvalue()

    # 2) Hent eksisterende document.xml for å gjenbruke namespaces og sectPr
    with ZipFile(io.BytesIO(base_bytes), "r") as zin:
        orig_doc_xml = zin.read("word/document.xml")

    root = ET.fromstring(orig_doc_xml)
    nsmap = root.nsmap.copy()
    # Sikre at 'w' finnes
    W_NS = nsmap.get("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main")
    XML_NS = "http://www.w3.org/XML/1998/namespace"

    body = root.find(f"{{{W_NS}}}body")
    sectPr = body.find(f"{{{W_NS}}}sectPr") if body is not None else None
    sectPr_clone = ET.fromstring(ET.tostring(sectPr)) if sectPr is not None else ET.Element(f"{{{W_NS}}}sectPr")

    # 3) Bygg nytt document.xml med revisjonsmarkering
    new_root = ET.Element(f"{{{W_NS}}}document", nsmap=nsmap)
    new_body = ET.SubElement(new_root, f"{{{W_NS}}}body")

    # Forutsigbar id + tidsstempel
    now_iso = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    rev_id = 1

    def _add_text_run(parent, text: str):
        r = ET.SubElement(parent, f"{{{W_NS}}}r")
        t = ET.SubElement(r, f"{{{W_NS}}}t")
        # Bevar ledende/etterfølgende blank
        if text.startswith(" ") or text.endswith(" "):
            t.set(f"{{{XML_NS}}}space", "preserve")
        t.text = text

    def _add_del_run(parent, text: str):
        r = ET.SubElement(parent, f"{{{W_NS}}}r")
        dt = ET.SubElement(r, f"{{{W_NS}}}delText")
        if text.startswith(" ") or text.endswith(" "):
            dt.set(f"{{{XML_NS}}}space", "preserve")
        dt.text = text

    # Del i avsnitt (linjeskift)
    orig_lines = (original_text or "").split("\n")
    imp_lines = (improved_text or "").split("\n")
    max_len = max(len(orig_lines), len(imp_lines))

    for i in range(max_len):
        a = orig_lines[i] if i < len(orig_lines) else ""
        b = imp_lines[i] if i < len(imp_lines) else ""

        p = ET.SubElement(new_body, f"{{{W_NS}}}p")

        a_tok = tokenize_keep_ws(a)
        b_tok = tokenize_keep_ws(b)
        edits = diff_tokens(a_tok, b_tok)

        for op, seg in edits:
            if not seg:
                continue
            if op == "=":
                _add_text_run(p, seg)
            elif op == "+":
                ins = ET.SubElement(
                    p,
                    f"{{{W_NS}}}ins",
                    {
                        f"{{{W_NS}}}author": author,
                        f"{{{W_NS}}}date": now_iso,
                        f"{{{W_NS}}}id": str(rev_id),
                    },
                )
                rev_id += 1
                _add_text_run(ins, seg)
            elif op == "-":
                de = ET.SubElement(
                    p,
                    f"{{{W_NS}}}del",
                    {
                        f"{{{W_NS}}}author": author,
                        f"{{{W_NS}}}date": now_iso,
                        f"{{{W_NS}}}id": str(rev_id),
                    },
                )
                rev_id += 1
                _add_del_run(de, seg)

    # Legg til seksjonsegenskaper (krav i Word at body slutter med sectPr)
    new_body.append(sectPr_clone)

    new_xml = ET.tostring(new_root, xml_declaration=True, encoding="UTF-8", standalone="yes")

    # 4) Pakk nytt word/document.xml inn i base DOCX
    out_bio = io.BytesIO()
    with ZipFile(io.BytesIO(base_bytes), "r") as zin, ZipFile(out_bio, "w") as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "word/document.xml":
                data = new_xml
            zout.writestr(item, data)

    return out_bio.getvalue()

# -----------------------------
# Action
# -----------------------------
run_btn = st.button(
    "⚙️ Forbedre språk og lag Ekte Track Changes",
    type="primary",
    disabled=(not uploaded_text or not OPENAI_API_KEY),
)

if run_btn:
    if not uploaded_text:
        st.error("Ingen tekst funnet.")
        st.stop()
    if not OPENAI_API_KEY:
        st.error("OPENAI_API_KEY mangler.")
        st.stop()

    with st.spinner("Forbedrer tekst …"):
        try:
            improved = improve_text(uploaded_text, tone, model_name)
        except Exception as e:
            st.error(f"Feil fra modellen: {e}")
            st.stop()

    # Ren forbedret DOCX
    improved_docx = make_docx_from_text(improved, "Forbedret tekst")

    # Original DOCX (om bruker limte inn tekst)
    if source_docx_bytes is None:
        source_docx_bytes = make_docx_from_text(uploaded_text, "Originaltekst")

    # Ekte Track Changes DOCX
    with st.spinner("Genererer ekte Track Changes …"):
        try:
            tracked_docx = make_tracked_changes_docx(
                original_text=uploaded_text,
                improved_text=improved,
                author="ChatGPT",
            )
        except Exception as e:
            st.error(f"Klarte ikke å lage Track Changes-dokument: {e}")
            tracked_docx = None

    st.success("Ferdig!")

    st.subheader("Forhåndsvisning (ren forbedret tekst)")
    st.text_area("Forbedret tekst", improved, height=300)

    st.download_button(
        "💾 Last ned ren forbedret Word (.docx)",
        data=improved_docx,
        file_name="forbedret_ren.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )

    if tracked_docx:
        st.download_button(
            "📝 Last ned med ekte Spor endringer (.docx)",
            data=tracked_docx,
            file_name="forbedret_spor_endringer_ekte.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )

st.markdown("---")
st.markdown(
    textwrap.dedent("""
    **Om Track Changes her:** Vi genererer WordprocessingML direkte med `<w:ins>` og `<w:del>` rundt endringer.
    Microsoft Word skal da kjenne igjen forslagene slik at du kan trykke **Godta**/**Avvis**.
    """)
)
