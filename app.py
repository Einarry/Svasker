import io
import textwrap
import streamlit as st
from docx import Document

# OpenAI SDK v1.x
from openai import OpenAI

st.set_page_config(page_title="MedLang Improver", page_icon="🩺", layout="centered")

# ---- Auth (OpenAI) ----
# På Streamlit Cloud: legg inn OPENAI_API_KEY i app-secrets (Settings → Secrets)
OPENAI_API_KEY = st.secrets.get("OPENAI_API_KEY", None)
if not OPENAI_API_KEY:
    st.warning("Mangler OPENAI_API_KEY i Secrets (Streamlit Cloud → Settings → Secrets).")
client = OpenAI(api_key=OPENAI_API_KEY)

st.title("🩺 Språkforbedrer for medisinske artikler")
st.markdown("Lim inn tekst eller last opp en Word-fil. Få en ny Word-fil med forbedret språk.")

# ---- Input UI ----
tab1, tab2 = st.tabs(["📄 Word-fil (.docx)", "✍️ Lim inn tekst"])
uploaded_text = None

with tab1:
    up = st.file_uploader("Last opp .docx", type=["docx"])
    if up is not None:
        try:
            doc = Document(up)
            uploaded_text = "\n".join([p.text for p in doc.paragraphs])
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

st.caption("Tips: Start med korte avsnitt. Store manus kan deles i seksjoner for best kontroll.")

def improve_text(text: str, mode: str) -> str:
    goals = {
        "Nøytral faglig": (
            "Forbedre klarhet, flyt og grammatikk i vitenskapelig medisinsk tekst. "
            "Behold betydning, referanser og tall uendret. Unngå nye påstander."
        ),
        "Mer konsis": (
            "Forbedre klarhet og gjør teksten mer konsis uten å endre faglig innhold. "
            "Behold referanser og tall. Fjern fyllord."
        ),
        "Mer formell": (
            "Hev formalitetsnivået, vitenskapelig stil, presis terminologi og grammatikk. "
            "Ikke legg til nye påstander."
        ),
        "For legfaglig publikum": (
            "Forenkle språket lett og forklar forkortelser der det er naturlig, men behold presisjon. "
            "Ikke endre data eller resultater."
        ),
    }
    sys = (
        "Du er språkredaktør for medisinske manus. "
        "Respekter faglig innhold, data, referanser (f.eks. [12], (Smith 2020), DOI), og numerikk. "
        "Ikke legg til, fjern eller omtolk resultater. "
        "Ikke endre referanseformatering."
    )
    instr = goals.get(mode, goals["Nøytral faglig"])

    # Bruk Chat Completions (stabil og enkel)
    resp = client.chat.completions.create(
        model=model,
        messages=[
            {"role": "system", "content": sys},
            {"role": "user", "content": f"{instr}\n\nTekst:\n{text}"}
        ],
        temperature=0.2,
    )
    return resp.choices[0].message.content.strip()

def make_docx(improved_text: str, original_text: str | None = None) -> bytes:
    """Lag enkel .docx med forbedret tekst + valgfritt original som vedlegg i slutten."""
    new_doc = Document()
    new_doc.add_heading("Forbedret tekst", level=1)

    # Del i avsnitt for bedre lesbarhet i Word
    for para in improved_text.split("\n"):
        new_doc.add_paragraph(para)

    if original_text:
        new_doc.add_page_break()
        new_doc.add_heading("Originaltekst (referanse)", level=2)
        for para in original_text.split("\n"):
            p = new_doc.add_paragraph()
            run = p.add_run(para)
            run.italic = True

    bio = io.BytesIO()
    new_doc.save(bio)
    return bio.getvalue()

run = st.button("⚙️ Forbedre språk", type="primary", disabled=(not uploaded_text or not OPENAI_API_KEY))

if run:
    with st.spinner("Forbedrer tekst..."):
        try:
            improved = improve_text(uploaded_text, tone)
        except Exception as e:
            st.error(f"Feil fra modellen: {e}")
            st.stop()

    st.success("Ferdig!")

    st.subheader("Forhåndsvisning")
    st.text_area("Forbedret tekst", improved, height=300)

    docx_bytes = make_docx(improved_text=improved, original_text=uploaded_text)
    st.download_button(
        "💾 Last ned Word (.docx)",
        data=docx_bytes,
        file_name="forbedret_tekst.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )

st.markdown("---")
st.markdown(
    textwrap.dedent("""
    **Merk om 'Track Changes':** Ekte spor endringer i .docx er teknisk komplisert fra Python.
    Denne MVP-en leverer en forbedret versjon + original på slutten for enkel manuell sammenligning.
    Hvis du ønsker, kan vi senere:
    - legge inn kommentarer på setningsnivå (Word-kommentarer),
    - eller eksportere både original og forbedret og bruke Words innebygde *Sammenlign dokumenter*.
    """)
)