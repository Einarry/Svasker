# app.py
from __future__ import annotations

import io
import pathlib
import re
import textwrap
from datetime import datetime, timezone
from zoneinfo import ZoneInfo
from typing import List, Tuple

import streamlit as st
from docx import Document
from docx.text.run import Run
from openai import OpenAI
import difflib

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
st.set_page_config(page_title="MedLang Improver", page_icon="🩺", layout="centered")
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

st.title("🩺 Språkforbedrer for medisinske artikler")
st.markdown("Last opp Word eller lim inn tekst. Få **ren forbedret** fil og en **visuell sammenligning** (understrek = innsatt, gjennomstreking = slettet).")

# -----------------------------
# OpenAI client
# -----------------------------
OPENAI_API_KEY = st.secrets.get("OPENAI_API_KEY", None)
if not OPENAI_API_KEY:
    st.warning("Mangler OPENAI_API_KEY i Secrets (App → ⋮ → Settings → Secrets). Appen kan ikke forbedre tekst.")
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

# --- Visual diff (no external deps) ---
# Tokeniserer på ord + mellomrom, bevarer whitespace.
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
        if tag == 'equal':
            out.append(('=', ''.join(a[i1:i2])))
        elif tag == 'insert':
            out.append(('+', ''.join(b[j1:j2])))
        elif tag == 'delete':
            out.append(('-', ''.join(a[i1:i2])))
        elif tag == 'replace':
            # representer som delete + insert
            out.append(('-', ''.join(a[i1:i2])))
            out.append(('+', ''.join(b[j1:j2])))
    return out

def make_visual_diff_docx(original_text: str, improved_text: str) -> bytes:
    """
    Lager en DOCX med visuell diff:
    - Innsatt tekst: understreket
    - Slettet tekst: gjennomstreket
    - Uendret: normal
    Avsnitt splittes på linjeskift for enkelhets skyld.
    """
    doc = Document()
    doc.add_heading("Visuell sammenligning (ikke ekte Spor endringer)", level=1)

    # Sammenlign linje for linje for å holde avsnitt strukturert
    orig_lines = original_text.split("\n")
    imp_lines = improved_text.split("\n")
    max_len = max(len(orig_lines), len(imp_lines))

    for idx in range(max_len):
        a = orig_lines[idx] if idx < len(orig_lines) else ""
        b = imp_lines[idx] if idx < len(imp_lines) else ""

        p = doc.add_paragraph()
        if not a and not b:
            continue

        a_tok = tokenize_keep_ws(a)
        b_tok = tokenize_keep_ws(b)
        edits = diff_tokens(a_tok, b_tok)

        for op, segment in edits:
            run: Run = p.add_run(segment)
            if op == '+':
                run.font.underline = True
            elif op == '-':
                run.font.strike = True
                # Gjør slettede deler litt lysegrå for synlighet
                run.font.color.rgb = None  # beholder standard; kan evt. justeres
            else:
                # lik tekst
                pass

    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()

# -----------------------------
# Action
# -----------------------------
run_btn = st.button(
    "⚙️ Forbedre språk",
    type="primary",
    disabled=(not uploaded_text or not OPENAI_API_KEY),
)

if run_btn:
    with st.spinner("Forbedrer tekst …"):
        try:
            improved = improve_text(uploaded_text, tone, model_name)
        except Exception as e:
            st.error(f"Feil fra modellen: {e}")
            st.stop()

    # Ren forbedret DOCX
    improved_docx = make_docx_from_text(improved, "Forbedret tekst")

    # Sikre original DOCX bytes for eventuell manuell sammenligning
    if source_docx_bytes is None:
        source_docx_bytes = make_docx_from_text(uploaded_text, "Originaltekst")

    # Visuell diff DOCX (lokal, ingen eksterne avhengigheter)
    with st.spinner("Lager visuell sammenligning …"):
        visual_diff_bytes = make_visual_diff_docx(
            original_text=uploaded_text,
            improved_text=improved
        )

    st.success("Ferdig!")

    st.subheader("Forhåndsvisning (ren forbedret tekst)")
    st.text_area("Forbedret tekst", improved, height=300)

    st.download_button(
        "💾 Last ned ren forbedret Word (.docx)",
        data=improved_docx,
        file_name="forbedret_ren.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )

    st.download_button(
        "📝 Last ned visuell sammenligning (.docx)",
        data=visual_diff_bytes,
        file_name="sammenligning_visuell.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )

st.markdown("---")
st.markdown(
    textwrap.dedent("""
    **Merk:** Dette dokumentet med visuell sammenligning bruker understrek (innsatt) og gjennomstreking (slettet).
    Det er ikke ekte *Spor endringer*. Hvis du trenger ekte Track Changes automatisk, kan vi koble til en
    ekstern dokumenttjeneste (f.eks. Aspose/GroupDocs Cloud) via API-nøkler, eller du kan i Word bruke
    **Se gjennom → Sammenlign** mellom original og forbedret.
    """)
)
