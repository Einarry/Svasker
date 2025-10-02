# app.py
from __future__ import annotations

import io
import csv
import json
import pathlib
import re
import textwrap
from datetime import datetime, timezone
from typing import Dict, List, Tuple
from zipfile import ZipFile

import streamlit as st
from docx import Document
from openai import OpenAI
import difflib
from lxml import etree as ET  # følger via python-docx

# =============================
# Build timestamp (UTC) + badge
# =============================
def _build_time_utc() -> str:
    try:
        ts = pathlib.Path(__file__).stat().st_mtime
        return datetime.fromtimestamp(ts, tz=timezone.utc).strftime("%Y-%m-%d %H:%M UTC")
    except Exception:
        return datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M UTC")

BUILD_TIME_UTC = _build_time_utc()

st.set_page_config(page_title="Einars språkvasker — Track Changes + Ordliste", page_icon="🩺", layout="centered")

# Kun én badge øverst til venstre
st.markdown(
    f"""
    <style>
      #build-badge {{
        position: fixed;
        top: 64px;
        left: 16px;
        padding: 4px 10px;
        border-radius: 8px;
        font-size: 12px;
        font-family: ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, monospace;
        background: rgba(0,0,0,0.06);
        backdrop-filter: blur(2px);
        z-index: 9998;
        pointer-events: none;
      }}
      @media (max-width: 768px) {{
        #build-badge {{
          position: sticky;
          top: 0;
          left: 0;
          margin: 6px 0;
        }}
      }}
    </style>
    <div id="build-badge">Build: {BUILD_TIME_UTC}</div>
    """,
    unsafe_allow_html=True,
)

st.title("Einars språkvasker")
st.markdown(
    "Last opp Word eller lim inn tekst. Få **ren forbedret** fil og en **.docx med ekte Spor endringer** "
    "(Word: Godta/Avvis). Du kan laste inn en **ordliste** (CSV/JSON) lokalt eller fra **Google Drive**."
)

# =============================
# OpenAI client
# =============================
OPENAI_API_KEY = st.secrets.get("OPENAI_API_KEY", None)
if not OPENAI_API_KEY:
    st.warning("Mangler OPENAI_API_KEY i Secrets (App → ⋮ → Settings → Secrets). Appen kan ikke forbedre tekst.")
client = OpenAI(api_key=OPENAI_API_KEY) if OPENAI_API_KEY else None

# =============================
# Standard-prompter (initiale verdier)
# =============================
DEFAULT_GOALS: Dict[str, str] = {
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

DEFAULT_SYSTEM_PROMPT = (
    "Du er språkredaktør for medisinske manus. "
    "Respekter faglig innhold, data, referanser (f.eks. [12], (Smith 2020), DOI), og numerikk. "
    "Ikke legg til, fjern eller omtolk resultater. "
    "Ikke endre referanseformatering."
)

# Legg i session_state hvis ikke finnes fra før
if "goals" not in st.session_state:
    st.session_state["goals"] = DEFAULT_GOALS.copy()
if "system_prompt" not in st.session_state:
    st.session_state["system_prompt"] = DEFAULT_SYSTEM_PROMPT

# =============================
# Google Drive-støtte (valgfritt)
# =============================
def drive_enabled() -> bool:
    return "GDRIVE_SERVICE_ACCOUNT_JSON" in st.secrets and "GDRIVE_FOLDER_ID" in st.secrets

def _folder_id_from_secret() -> str:
    raw = st.secrets["GDRIVE_FOLDER_ID"].strip()
    m = re.search(r"/folders/([a-zA-Z0-9_-]+)", raw)
    if m:
        return m.group(1)
    m = re.search(r"[?&]id=([a-zA-Z0-9_-]+)", raw)
    if m:
        return m.group(1)
    return raw

def drive_find_file(service, folder_id: str, name: str):
    """Finn fil i mappe ved robust navnematch (case/whitespace-insensitiv)."""
    def _norm(s: str) -> str:
        # trim, normaliser whitespace og små bokstaver
        return re.sub(r"\s+", " ", s).strip().lower()

    target_norm = _norm(name)

    # 1) List alt i mappen (paginert) og sammenlikn lokalt
    page_token = None
    while True:
        resp = service.files().list(
            q=f"'{folder_id}' in parents and trashed = false",
            fields="nextPageToken, files(id, name, mimeType)",
            pageSize=1000,
            pageToken=page_token,
        ).execute()
        for f in resp.get("files", []):
            if _norm(f["name"]) == target_norm:
                return f
        page_token = resp.get("nextPageToken")
        if not page_token:
            break

    # 2) Fallback: prøv et 'contains' på navnet uten filendelse
    base = name.rsplit(".", 1)[0]
    resp = service.files().list(
        q=f"'{folder_id}' in parents and name contains '{base}' and trashed = false",
        fields="files(id, name, mimeType)",
        pageSize=50,
    ).execute()
    files = resp.get("files", [])
    return files[0] if files else None

def drive_upload_bytes(service, folder_id: str, name: str, data: bytes, mime: str):
    from googleapiclient.http import MediaIoBaseUpload
    existing = drive_find_file(service, folder_id, name)
    media = MediaIoBaseUpload(io.BytesIO(data), mimetype=mime, resumable=False)
    if existing:
        service.files().update(fileId=existing["id"], media_body=media).execute()
        return existing["id"]
    meta = {"name": name, "parents": [folder_id]}
    file = service.files().create(body=meta, media_body=media, fields="id").execute()
    return file["id"]

def drive_download_bytes(service, file_id: str) -> bytes:
    from googleapiclient.http import MediaIoBaseDownload
    req = service.files().get_media(fileId=file_id)
    bio = io.BytesIO()
    downloader = MediaIoBaseDownload(bio, req)
    done = False
    while not done:
        _, done = downloader.next_chunk()
    return bio.getvalue()

def save_glossary_to_drive(gloss: Dict[str, List[str]], filename="ordliste.csv"):
    service = _drive_service()
    folder_id = _folder_id_from_secret()
    data = dump_glossary_csv(gloss) if filename.lower().endswith(".csv") else dump_glossary_json(gloss)
    mime = "text/csv" if filename.lower().endswith(".csv") else "application/json"
    drive_upload_bytes(service, folder_id, filename, data, mime)

def load_glossary_from_drive(filename="ordliste.csv") -> Dict[str, List[str]] | None:
    service = _drive_service()
    folder_id = _folder_id_from_secret()
    f = drive_find_file(service, folder_id, filename)
    if not f:
        return None
    data = drive_download_bytes(service, f["id"])
    if filename.lower().endswith(".json"):
        obj = json.loads(data.decode("utf-8"))
        if isinstance(obj, list):
            return {term: [] for term in obj}
        elif isinstance(obj, dict):
            return {k: list(v) if isinstance(v, (list, tuple)) else ([str(v)] if v else []) for k, v in obj.items()}
        else:
            raise ValueError("JSON må være liste eller objekt.")
    text = data.decode("utf-8")
    return _parse_glossary_csv_text(text)

# =============================
# Ordliste: parse/lagre + bruk
# =============================
TOKEN_SEP = re.compile(r"[;|]")

def _parse_glossary_csv_text(text: str) -> Dict[str, List[str]]:
    rows = list(csv.reader(io.StringIO(text)))
    if not rows:
        return {}
    header = [h.strip().lower() for h in rows[0]]
    gloss: Dict[str, List[str]] = {}

    def parse_syns(s: str) -> List[str]:
        if not s:
            return []
        parts = TOKEN_SEP.split(s)
        return [p.strip() for p in parts if p.strip()]

    if "preferred" in header:
        idx_pref = header.index("preferred")
        idx_syn = header.index("synonyms") if "synonyms" in header else None
        for r in rows[1:]:
            if not r or len(r) <= idx_pref:
                continue
            pref = r[idx_pref].strip()
            if not pref:
                continue
            syns = parse_syns(r[idx_syn]) if idx_syn is not None and len(r) > idx_syn else []
            gloss.setdefault(pref, [])
            for s in syns:
                if s and s not in gloss[pref]:
                    gloss[pref].append(s)
    else:
        for r in rows:
            if r and r[0].strip().lower() != "preferred":
                pref = r[0].strip()
                if pref:
                    gloss.setdefault(pref, [])
    return gloss

def load_glossary_file(file) -> Dict[str, List[str]]:
    name = getattr(file, "name", "").lower()
    data = file.read()
    if hasattr(file, "seek"):
        file.seek(0)
    if name.endswith(".json"):
        obj = json.loads(data.decode("utf-8"))
        if isinstance(obj, list):
            return {term: [] for term in obj}
        elif isinstance(obj, dict):
            return {k: list(v) if isinstance(v, (list, tuple)) else ([str(v)] if v else []) for k, v in obj.items()}
        else:
            raise ValueError("JSON må være liste eller objekt.")
    text = data.decode("utf-8")
    return _parse_glossary_csv_text(text)

def dump_glossary_csv(gloss: Dict[str, List[str]]) -> bytes:
    buff = io.StringIO()
    w = csv.writer(buff)
    w.writerow(["preferred", "synonyms"])
    for pref, syns in gloss.items():
        s = "; ".join(syns) if syns else ""
        w.writerow([pref, s])
    return buff.getvalue().encode("utf-8")

def dump_glossary_json(gloss: Dict[str, List[str]]) -> bytes:
    return json.dumps(gloss, ensure_ascii=False, indent=2).encode("utf-8")

def _match_case(repl: str, src: str) -> str:
    if src.isupper():
        return repl.upper()
    if len(src) > 1 and src[0].isupper() and src[1:].islower():
        return repl.capitalize()
    return repl

def apply_glossary(text: str, glossary: Dict[str, List[str]]) -> str:
    if not glossary:
        return text
    out = text
    for preferred, syns in glossary.items():
        for s in set(syns):
            if not s or s.strip().lower() == preferred.strip().lower():
                continue
            pattern = re.compile(rf"\b{re.escape(s)}\b", flags=re.IGNORECASE)
            out = pattern.sub(lambda m: _match_case(preferred, m.group(0)), out)
    return out

# =============================
# Modeller & prompts
# =============================
def improve_text(text: str, mode: str, model_name: str, glossary_note: str | None) -> str:
    if not client:
        raise RuntimeError("OPENAI_API_KEY mangler – kan ikke kalle modellen.")
    system_prompt = st.session_state["system_prompt"]
    goals_map: Dict[str, str] = st.session_state["goals"]
    instr = goals_map.get(mode, list(goals_map.values())[0] if goals_map else DEFAULT_GOALS["Nøytral faglig"])
    if glossary_note:
        instr = instr + "\n\n" + glossary_note
    resp = client.chat.completions.create(
        model=model_name,
        messages=[
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": f"{instr}\n\nTekst:\n{text}"},
        ],
        temperature=0.2,
    )
    return resp.choices[0].message.content.strip()

# =============================
# Track Changes-generator (w:ins / w:del)
# =============================
TOKEN_RE = re.compile(r"\s+|\w+|[^\w\s]", re.UNICODE)

def tokenize_keep_ws(s: str) -> List[str]:
    return TOKEN_RE.findall(s)

def diff_tokens(a: List[str], b: List[str]) -> List[Tuple[str, str]]:
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

def make_docx_from_text(text: str, heading: str | None = None) -> bytes:
    d = Document()
    if heading:
        d.add_heading(heading, level=1)
    for para in text.split("\n"):
        d.add_paragraph(para)
    bio = io.BytesIO()
    d.save(bio)
    return bio.getvalue()

def make_tracked_changes_docx(original_text: str, improved_text: str, author: str = "ChatGPT") -> bytes:
    base_doc = Document()
    base_bio = io.BytesIO()
    base_doc.save(base_bio)
    base_bytes = base_bio.getvalue()

    with ZipFile(io.BytesIO(base_bytes), "r") as zin:
        orig_doc_xml = zin.read("word/document.xml")

    root = ET.fromstring(orig_doc_xml)
    nsmap = root.nsmap.copy()
    W_NS = nsmap.get("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main")
    XML_NS = "http://www.w3.org/XML/1998/namespace"

    body = root.find(f"{{{W_NS}}}body")
    sectPr = body.find(f"{{{W_NS}}}sectPr") if body is not None else None
    sectPr_clone = ET.fromstring(ET.tostring(sectPr)) if sectPr is not None else ET.Element(f"{{{W_NS}}}sectPr")

    new_root = ET.Element(f"{{{W_NS}}}document", nsmap=nsmap)
    new_body = ET.SubElement(new_root, f"{{{W_NS}}}body")

    now_iso = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    rev_id = 1

    def _add_text_run(parent, text: str):
        r = ET.SubElement(parent, f"{{{W_NS}}}r")
        t = ET.SubElement(r, f"{{{W_NS}}}t")
        if text.startswith(" ") or text.endswith(" "):
            t.set(f"{{{XML_NS}}}space", "preserve")
        t.text = text

    def _add_del_run(parent, text: str):
        r = ET.SubElement(parent, f"{{{W_NS}}}r")
        dt = ET.SubElement(r, f"{{{W_NS}}}delText")
        if text.startswith(" ") or text.endswith(" "):
            dt.set(f"{{{XML_NS}}}space", "preserve")
        dt.text = text

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
                    p, f"{{{W_NS}}}ins",
                    {f"{{{W_NS}}}author": author, f"{{{W_NS}}}date": now_iso, f"{{{W_NS}}}id": str(rev_id)}
                )
                rev_id += 1
                _add_text_run(ins, seg)
            elif op == "-":
                de = ET.SubElement(
                    p, f"{{{W_NS}}}del",
                    {f"{{{W_NS}}}author": author, f"{{{W_NS}}}date": now_iso, f"{{{W_NS}}}id": str(rev_id)}
                )
                rev_id += 1
                _add_del_run(de, seg)

    new_body.append(sectPr_clone)
    new_xml = ET.tostring(new_root, xml_declaration=True, encoding="UTF-8", standalone="yes")

    out_bio = io.BytesIO()
    with ZipFile(io.BytesIO(base_bytes), "r") as zin, ZipFile(out_bio, "w") as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "word/document.xml":
                data = new_xml
            zout.writestr(item, data)
    return out_bio.getvalue()

# =============================
# UI: tekstinn og modellvalg
# =============================
tab1, tab2 = st.tabs(["📄 Word-fil (.docx)", "✍️ Lim inn tekst"])
uploaded_text: str | None = None

with tab1:
    up = st.file_uploader("Last opp .docx", type=["docx"])
    if up is not None:
        try:
            src_bytes = up.read()
            doc = Document(io.BytesIO(src_bytes))
            uploaded_text = "\n".join(p.text for p in doc.paragraphs)
        except Exception as e:
            st.error(f"Kunne ikke lese Word-filen: {e}")

with tab2:
    pasted = st.text_area("Lim inn tekst her", height=300, placeholder="Lim inn manus her …")
    if pasted and not uploaded_text:
        uploaded_text = pasted

# Tone + modell
tone_options = list(st.session_state["goals"].keys()) or list(DEFAULT_GOALS.keys())
colA, colB = st.columns(2)
with colA:
    tone = st.selectbox("Tone/retning", tone_options, help="Velg hvordan teksten skal forbedres.")
with colB:
    default_models = ["gpt-5", "gpt-5-mini", "gpt-4o", "gpt-4o-mini"]
    model_name = st.selectbox("Modell", default_models, index=0)
custom_model = st.text_input("Egendefinert modellnavn (valgfritt)", placeholder="f.eks. gpt-5-chat-latest")
if custom_model.strip():
    model_name = custom_model.strip()

st.caption("Tips: Del store manus i seksjoner (Introduksjon, Metode, Resultater, osv.) for bedre kontroll.")

# =============================
# ✏️ Rediger PROMPTS (SYSTEM + GOALS for valgt tone)
# =============================
with st.expander("✏️ Rediger PROMPTS (SYSTEM + GOALS for valgt tone)"):
    sys_text = st.text_area("SYSTEM PROMPT", st.session_state["system_prompt"], height=140)
    current_goal = st.session_state["goals"].get(tone, "")
    goal_text = st.text_area(f"GOALS for «{tone}»", current_goal, height=140)

    c1, c2, c3 = st.columns(3)
    with c1:
        if st.button("💾 Lagre SYSTEM PROMPT"):
            st.session_state["system_prompt"] = sys_text
            st.success("System-prompt lagret.")
    with c2:
        if st.button("💾 Lagre GOALS for valgt tone"):
            st.session_state["goals"][tone] = goal_text
            st.success(f"GOALS for «{tone}» lagret.")
    with c3:
        if st.button("↩️ Tilbakestill til standard"):
            st.session_state["system_prompt"] = DEFAULT_SYSTEM_PROMPT
            st.session_state["goals"] = DEFAULT_GOALS.copy()
            st.success("Prompter tilbakestilt til standardverdier.")

# =============================
# Ordliste-UI (kompakt – ingen innholdsvisning)
# =============================
with st.expander("📚 Ordliste (last opp / Drive / lagre)"):
    drive_ok = drive_enabled()
    st.info(("Google Drive er konfigurert." if drive_ok else
             "Google Drive er ikke konfigurert (legg inn GDRIVE_SERVICE_ACCOUNT_JSON og GDRIVE_FOLDER_ID i Secrets)."))

    default_name = st.text_input("Filnavn i Drive", value="ordliste02okt.csv",
                                 help="Brukes når du leser/lagrer mot Drive.")

    uploaded_gloss = st.file_uploader("Last opp ordliste (CSV/JSON)", type=["csv", "json"], key="gloss_uploader")
    if uploaded_gloss:
        try:
            gloss = load_glossary_file(uploaded_gloss)
            st.session_state["glossary"] = gloss
            st.success(f"Ordlistestatus: {len(gloss)} foretrukne termer lastet.")
        except Exception as e:
            st.error(f"Kunne ikke lese ordlisten: {e}")

    c1, c2, c3 = st.columns(3)
    with c1:
        if st.button("📥 Last inn fra Drive", disabled=not drive_ok):
            try:
                g = load_glossary_from_drive(default_name)
                if g is None:
                    st.warning(f"Fant ikke «{default_name}» i Drive-mappen.")
                else:
                    st.session_state["glossary"] = g
                    st.success(f"Ordlistestatus: {len(g)} termer lastet fra Drive.")
            except Exception as e:
                st.error(f"Feil ved lesing fra Drive: {e}")

    with c2:
        if st.button("💾 Lagre til Drive", disabled=not drive_ok):
            try:
                gloss = st.session_state.get("glossary", {})
                if not gloss:
                    st.warning("Ingen ordliste i minnet å lagre. Last opp eller last inn først.")
                else:
                    save_glossary_to_drive(gloss, default_name)
                    st.success(f"Lagret ordliste som «{default_name}» i Drive-mappen.")
            except Exception as e:
                st.error(f"Feil ved lagring til Drive: {e}")

    with c3:
        gloss = st.session_state.get("glossary", {})
        if gloss:
            data_csv = dump_glossary_csv(gloss)
            data_json = dump_glossary_json(gloss)
            st.download_button("⬇️ Last ned CSV", data=data_csv, file_name="ordliste.csv", mime="text/csv")
            st.download_button("⬇️ Last ned JSON", data=data_json, file_name="ordliste.json", mime="application/json")
        else:
            st.caption("Ingen ordliste lastet enda.")

# =============================
# Kjøring
# =============================
use_glossary_in_prompt = st.checkbox("Gi modellen beskjed om å bruke ordlisten (anbefalt)", value=True,
                                     help="Legger en kort instruks om prefererte termer inn i prompten.")
author_tag = st.text_input("Forfatter-tag for Track Changes", value="ChatGPT",
                           help="Vises i Word som forfatter av endringer.")

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

    glossary = st.session_state.get("glossary", {})

    glossary_note = None
    if use_glossary_in_prompt and glossary:
        preferred_list = "\n".join(f"- {p}" for p in list(glossary.keys())[:100])
        glossary_note = (
            "Bruk følgende prefererte termer konsekvent; normaliser eventuelle synonymer/varianter til eksakt skrivemåte:\n"
            f"{preferred_list}"
        )

    with st.spinner("Forbedrer tekst …"):
        try:
            improved = improve_text(uploaded_text, mode=tone, model_name=model_name, glossary_note=glossary_note)
        except Exception as e:
            st.error(f"Feil fra modellen: {e}")
            st.stop()

    final_text = apply_glossary(improved, glossary) if glossary else improved

    improved_docx = make_docx_from_text(final_text, "Forbedret tekst")

    with st.spinner("Genererer ekte Track Changes …"):
        try:
            tracked_docx = make_tracked_changes_docx(
                original_text=uploaded_text or "",
                improved_text=final_text or "",
                author=author_tag or "ChatGPT",
            )
        except Exception as e:
            st.error(f"Klarte ikke å lage Track Changes-dokument: {e}")
            tracked_docx = None

    st.success("Ferdig!")

    st.subheader("Forhåndsvisning (ren forbedret tekst)")
    st.text_area("Forbedret tekst", final_text, height=300)

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
    **Ordliste:** UI viser ikke innholdet for å spare plass. Du kan likevel laste opp, lagre til Drive,
    og laste ned som CSV/JSON.  
    **PROMPTS:** Du kan redigere både SYSTEM-prompten og GOALS-prompten for valgt tone. Endringer lagres i denne økten.  
    **Track Changes:** Dokumentet genereres med `<w:ins>`/`<w:del>` slik at Word kan Godta/Avvise endringer.
    """)
)


