import io
import tempfile
import textwrap
import streamlit as st
from docx import Document
from openai import OpenAI

# NEW: Python-Redlines
from python_redlines.engines import XmlPowerToolsEngine

st.set_page_config(page_title="MedLang Improver (Track Changes)", page_icon="🩺", layout="centered")

OPENAI_API_KEY = st.secrets.get("OPENAI_API_KEY", None)
if not OPENAI_API_KEY:
    st.warning("Mangler OPENAI_API_KEY i Secrets (App → ⋮ → Settings → Secrets).")
client = OpenAI(api_key=OPENAI_API_KEY)

st.title("🩺 Språkforbedrer for medisinske artikler — med Track Changes")
st.markdown("Last opp Word eller lim inn tekst. Få både **ren** forbedret fil og en **.docx med Spor endringer**.")

tab1, tab2 = st.tabs(["📄 Word-fil (.docx)", "✍️ Lim inn tekst"])
uploaded_text = None
source_docx_bytes = None

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
    pasted = st.text_area("Lim inn tekst her", height=300, placeholder="Lim inn manus...")
    if pasted and not uploaded_text:
        uploaded_text = pasted

colA, colB = st.columns(2)
with colA:
    tone = st.selectbox("Tone/retning", ["Nøytral faglig", "Mer konsis", "Mer formell", "For legfaglig publikum"])
with colB:
    model = st.selectbox("Modell", ["gpt-4o-mini", "gpt-4o"], help="Mini er rimelig og rask. gpt-4o gir litt høyere kvalitet.")

st.caption("Tips: Del store manus i seksjoner for bedre kontroll. Track Changes genereres ved å sammenligne original og forbedret versjon.")

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

def build_redline_docx(original_docx: bytes, improved_docx: bytes, author_tag: str = "ChatGPT") -> bytes:
    """
    Bruk Python-Redlines (Open-XML-PowerTools under panseret) til å lage en .docx med track changes.
    """
    engine = XmlPowerToolsEngine()  # bruker medfølgende binær
    # run_redline kan ta bytes eller filstier og returnerer bytes for redline.docx
    redlined_bytes = engine.run_redline(
        author_tag=author_tag,
        original_docx_bytes=original_docx,
        modified_docx_bytes=improved_docx
    )
    return redlined_bytes

run = st.button("⚙️ Forbedre språk og lag Spor endringer", type="primary", disabled=(not uploaded_text or not OPENAI_API_KEY))

if run:
    with st.spinner("Forbedrer tekst..."):
        try:
            improved = improve_text(uploaded_text, tone, model)
        except Exception as e:
            st.error(f"Feil fra modellen: {e}")
            st.stop()

    # Lag 'ren' forbedret DOCX
    improved_docx = make_docx_from_text(improved, "Forbedret tekst")

    # Lag original DOCX (hvis bruker limte inn tekst i stedet for å laste opp fil)
    if source_docx_bytes is None:
        source_docx_bytes = make_docx_from_text(uploaded_text, "Originaltekst")

    # Forsøk å lage Track Changes (.docx)
    redline_bytes = None
    redline_error = None
    with st.spinner("Genererer Word med Spor endringer..."):
        try:
            redline_bytes = build_redline_docx(source_docx_bytes, improved_docx, author_tag="ChatGPT")
        except Exception as e:
            redline_error = str(e)

    st.success("Ferdig!")

    st.subheader("Forhåndsvisning (ren forbedret tekst)")
    st.text_area("Forbedret tekst", improved, height=300)

    st.download_button(
        "💾 Last ned ren forbedret Word (.docx)",
        data=improved_docx,
        file_name="forbedret_ren.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )

    if redline_bytes:
        st.download_button(
            "📝 Last ned med Spor endringer (.docx)",
            data=redline_bytes,
            file_name="forbedret_spor_endringer.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
        st.caption("Dette dokumentet inneholder ekte Track Changes du kan godta/avslå i Word.")
    else:
        st.warning(
            "Klarte ikke å generere Word med Spor endringer denne gangen. "
            "Du kan likevel bruke den rene forbedrede filen. "
            "Hvis problemet vedvarer i skyen, kan vi aktivere en fallback som legger inn kommentarer i stedet."
        )

st.markdown("---")
st.markdown(
    textwrap.dedent("""
    **Hvorfor dette virker:** Vi sammenligner original og forbedret tekst for å lage en tredje fil med Track Changes
    (w:ins/w:del). Word viser disse som endringsforslag du kan godta/avslå.
    """)
)
