# app.py
# Streamlit app for improving medical-scientific text and producing DOCX outputs.
# - Clean improved DOCX
# - Track Changes DOCX (via Python-Redlines, if available)
#
# NOTE: Make sure your requirements include:
#   streamlit
#   python-docx
#   openai
#   python_redlines @ git+https://github.com/JSv4/Python-Redlines@v0.0.5
#
# And add your API key in Streamlit Cloud (App → ⋮ → Settings → Secrets):
#   OPENAI_API_KEY = sk-...

from __future__ import annotations

import io
import pathlib
import textwrap
from datetime import datetime, timezone
from zoneinfo import ZoneInfo

import streamlit as st
from docx import Document
from openai import OpenAI

# Optional Track Changes engine (Python-Redlines). We import lazily inside a function
# so the app still runs even if the package or runtime is missing.


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
st.set_page_config(page_title="MedLang Improver — Track Changes", page_icon="🩺", layout="centered")

# Build badge in top-left
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

st.title("🩺 Språkforbedrer for medisinske artikler — med Track Changes")
st.markdown("Last opp Word eller lim inn tekst. Få både **ren** forbedret fil og en **.docx med Spor endringer** (når tilgjengelig).")

# -----------------------------
# OpenAI client
# -----------------------------
OPENAI_API_KEY = st.secrets.get("OPENAI_API_KEY", None)
if not OPENAI_API_KEY:
    st.warning("Mangler OPENAI_API_KEY i Secrets (App → ⋮ → Settings → Secrets). Appen vil ikke kunne forbedre tekst.")
client = OpenAI(api_key=OPENAI_API_KEY)

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
        help="Velg hvordan teksten skal forbedres."
    )
with colB:
    model_name = st.selectbox(
        "Modell",
        ["gpt-4o-mini", "gpt-4o"],
        help="Mini er rimelig og rask; gpt-4o kan gi litt høyere kvalitet."
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
# Helpers
# -----------------------------
def improve_text(text: str, mode: str, model_name: str) -> str:
    instr = GOALS.get(mode, GOALS["Nøytral faglig"])
    # Chat Completions (OpenAI SDK v1)
    resp = client.chat.completions.create(
        model=model_name,
        messages=[
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user", "content": f"{instr}\n\nTekst:\n{text}"}
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


def build_redline_docx(original_docx: bytes, improved_docx: bytes, author_tag: str = "ChatGPT") -> bytes:
    """
    Generate a real Track Changes DOCX by comparing original vs improved.
    Uses Python-Redlines (Open-XML-PowerTools under the hood).
    Returns bytes for the redlined DOCX or raises on failure.
    """
    # Lazy import so the app can run without the dependency.
    from python_redlines.engines import XmlPowerToolsEngine  # type: ignore

    engine = XmlPowerToolsEngine()
    # The wrapper accepts bytes or file paths; return value is bytes for the redline docx.
    redlined_bytes = engine.run_redline(
        author_tag=author_tag,
        original_docx_bytes=original_docx,
        modified_docx_bytes=improved_docx,
    )
    return redlined_bytes


# -----------------------------
# Action
# -----------------------------
run_btn =_
