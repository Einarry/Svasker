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
from lxml import etree as ET  # egen avhengighet

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

def _drive_service():
    from google.oauth2.service_account import Credentials
    from googleapiclient.discovery import build

    raw = st.secrets["GDRIVE_SERVICE_ACCOUNT_JSON"]
    try:
        info = json.loads(raw)
    except json.JSONDecodeError:
        m = re.search(r'"private_key"\s*:\s*"(.*?)"', raw, flags=re.DOTALL)
        if not m:
            raise
        key_content = m.group(1)
        fixed_key = key_content.replace("\r\n", "\\n").replace("\n", "\\n")
        raw = raw[:m.start(1)] + fixed_key + raw[m.end(1):]
        info = json.loads(raw)

    scopes = ["https://www.googleapis.com/auth/drive"]  # les + skriv
    creds = Credentials.from_service_account_info(info, scopes=scopes)
    return build("drive", "v3", credentials=creds)

def drive_list_files(service, folder_id: str, limit: int = 100) -> List[dict]:
    resp = service.files().list(
        q=f"'{folder_id}' in parents and trashed = false",
        fields="files(id,name,mimeType)",
        pageSize=limit
    ).execute()
    return resp.get("files", [])

def drive_find_file(service, folder_id: str, name: str) -> dict | None:
    def _norm(s: str) -> str:
        return re.sub(r"\s+", " ", s).strip().lower()
    target_norm = _norm(name)
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
    base = name.rsplit(".", 1)[0].replace("'", "\\'")
    resp = service.files().list(
        q=f"'{folder_id}' in parents and name contains '{base}' and trashed = false",
        fields="files(id, name, mimeType)",
        pageSize=50,
    ).execute()
    files = resp.get("files", [])
    return files[0] if files else None

def drive_upload_bytes(service, folder_id: str, name: str, data: bytes, mime: str) -> str:
    from googleapiclient.http import MediaIoBaseUpload
    existing = drive_find_file(service, folder_id, name)
    media = MediaIoBaseUpload(io.BytesIO(data), mimetype=mime, resumable=False)
    if existing:
        service.files().update(fileId=existing["id"], media_body=media).execute()
        return existing["id"]
    meta = {"name": name, "parents": [folder_id]}
    file = service.files().create(body=meta, media_body=media, fields="id").execute()
    return file["id"]

def _drive_export_bytes(service, file_id: str, mime: str) -> bytes:
    from googleapiclient.http import MediaIoBaseDownload
    req = service.files().export_media(fileId=file_id, mimeType=mime)
    bio = io.BytesIO()
    downloader = MediaIoBaseDownload(bio, req)
    done = False
    while not done:
        _, done = downloader.next_chunk()
    return bio.getvalue()

def drive_download_bytes(service, file_id: str) -> bytes:
    from googleapiclient.http import MediaIoBaseDownload
    req = service.files().get_media(fileId=file_id)
    bio = io.BytesIO()
    downloader = MediaIoBaseDownload(bio, req)
    done = False
    while not done:
        _, done = downloader.next_chunk()
    return bio.getvalue()

def save_glossary_to_drive(gloss: Dict[str, List[str]], filename: str = "ordliste.csv") -> str:
    service = _drive_service()
    folder_id = _folder_id_from_secret()
    if filename.lower().endswith(".csv"):
        data = dump_glossary_csv(gloss)
        mime = "text/csv"
    else:
        data = dump_glossary_json(gloss)
        mime = "application/json"
    return drive_upload_bytes(service, folder_id, filename, data, mime)

def load_glossary_from_drive(filename: str = "ordliste.csv") -> Dict[str, List[str]] | None:
    service = _drive_service()
    folder_id = _folder_id_from_secret()
    f = drive_find_file(service, folder_id, filename)
    if not f:
        return None
    mime = f.get("mimeType", "")
    if mime == "application/vnd.google-apps.spreadsheet":
        data = _drive_export_bytes(service, f["id"], "text/csv")
        text = data.decode("utf-8")
        return _parse_glossary_csv_text(text)
    else:
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

def apply_glossary_with_report(text: str, glossary: Dict[str, List[str]]):
    if not glossary:
        return text, {"total": 0, "by_preferred": {}, "pairs": []}
    report = {"total": 0, "by_preferred": {}, "pairs": []}
    out = text
    for preferred, syns in glossary.items():
        pref_norm = preferred.strip().lower()
        def repl(m):
            src = m.group(0)
            dst = _match_case(preferred, src)
            report["total"] += 1
            report["by_preferred"][preferred] = report["by_preferred"].get(preferred, 0) + 1
            report["pairs"].append((src, dst))
            return dst
        for s in set(syns):
            if not s:
                continue
            if s.strip().lower() == pref_norm:
                continue
            pattern = re.compile(rf"\b{re.escape(s)}\b", flags=re.IGNORECASE)
            out = pattern.sub(repl, out)
    return out, report

# =============================
# Aggressivitetsbudsjett + modellkall
# =============================
def budget_from_strength(text_len: int, strength: int) -> Tuple[int, int]:
    """
    Returner (max_changes, max_insert_chars) basert på tekstlengde og aggressivitet 1–5.
    Skalerer lineært på tekstlengde.
    """
    per_1000 = {
        1: (2, 3),    # (endringer, innsettingstegn*10) -> ca 0.3% (3 per 1000)
        2: (5, 8),    # ~0.8%
        3: (10, 15),  # ~1.5%
        4: (20, 30),  # ~3.0%
        5: (40, 60),  # ~6.0%
    }[max(1, min(5, strength))]
    per_1000_changes, per_1000_ins_chars = per_1000
    scale = max(1, int(round(text_len / 1000.0)))
    return per_1000_changes * scale, per_1000_ins_chars * scale

def build_strength_instructions(level: int) -> str:
    presets = {
        1: ("Minimal", [
            "Rett bare stavefeil, tegnsetting og klare grammatikkfeil.",
            "IKKE omformuler ellers. Ikke slå sammen/splitte setninger.",
            "Hvis setningen er akseptabel, la den stå uendret."
        ]),
        2: ("Forsiktig", [
            "Små justeringer for klarhet/flyt. Unngå synonymbytter (med mindre ordlisten krever det).",
            "Ikke endre struktur eller avsnittsrekkefølge."
        ]),
        3: ("Moderat", [
            "Forbedre klarhet og konsishet. Fjern fyllord.",
            "Behold avsnittslogikk uendret."
        ]),
        4: ("Tydelig", [
            "Konsis omskriving per setning, moderat omstrukturering tillatt.",
            "Bevar innhold/mening uendret."
        ]),
        5: ("Aggressiv", [
            "Optimaliser stil og flyt fritt så lenge meningen bevares.",
            "Moderat omstrukturering er ok."
        ]),
    }
    name, rules = presets[level]
    bullet = "\n".join(f"- {r}" for r in rules)
    return f"Aggressivitetsnivå: {name}.\n{bullet}"

def improve_text(text: str, tone: str, model_name: str, glossary_note: str | None,
                 strength: int, max_changes: int, max_insert_chars: int, strict=False) -> str:
    if not client:
        raise RuntimeError("OPENAI_API_KEY mangler – kan ikke kalle modellen.")
    system_prompt = st.session_state["system_prompt"]
    goals_map: Dict[str, str] = st.session_state["goals"]
    tone_goal = goals_map.get(tone, list(goals_map.values())[0] if goals_map else DEFAULT_GOALS["Nøytral faglig"])
    strength_instr = build_strength_instructions(strength)

    budget_text = (
        f"**Endringsbudsjett:** Sikt mot ≤ {max_changes} totale endringer (innsettinger+slettinger) "
        f"og ≤ {max_insert_chars} innsettingstegn totalt. Prioriter å rette feil med størst effekt først. "
        f"Hvis du risikerer å overskride budsjettet, la setninger stå uendret."
    )

    if strict:
        budget_text += "\n**Streng håndheving:** Du overskred budsjettet i forrige utkast. Produser en ny versjon som holder seg **innenfor** budsjettet. " \
                       "Behold kun de mest viktige forbedringene (grammatikkfeil/tvetydighet)."

    instr = tone_goal + "\n\n" + strength_instr + "\n\n" + budget_text
    if glossary_note:
        instr += "\n\n" + glossary_note

    params = {
        "model": model_name,
        "messages": [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": f"{instr}\n\nTekst:\n{text}"},
        ],
    }
    resp = client.chat.completions.create(**params)
    return resp.choices[0].message.content.strip()

# =============================
# Diff & Track Changes (to-pass, to forfattere) + måling
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

def estimate_changes(orig_text: str, new_text: str) -> Tuple[int, int]:
    a = tokenize_keep_ws(orig_text or "")
    b = tokenize_keep_ws(new_text or "")
    ops = diff_tokens(a, b)
    changes = sum(1 for op, _ in ops if op in {"+", "-"})
    ins_chars = sum(len(seg) for op, seg in ops if op == "+")
    return changes, ins_chars

def make_docx_from_text(text: str, heading: str | None = None) -> bytes:
    d = Document()
    if heading:
        d.add_heading(heading, level=1)
    for para in text.split("\n"):
        d.add_paragraph(para)
    bio = io.BytesIO()
    d.save(bio)
    return bio.getvalue()

def make_tracked_changes_docx(
    original_text: str,
    improved_text: str,   # språk-pass resultat
    final_text: str,      # etter ordliste-pass
    author_lang: str = "ChatGPT",            # forfatter for språk-pass
    author_gloss: str = "ChatGPT (ordliste)" # forfatter for ordliste-pass
) -> Tuple[bytes, int, int, int, int]:
    # Identifiser ordliste-endringer via diff(improved -> final)
    imp_tok = tokenize_keep_ws(improved_text or "")
    fin_tok = tokenize_keep_ws(final_text or "")
    gl_edits = diff_tokens(imp_tok, fin_tok)
    gl_insert_set = set(seg.strip() for op, seg in gl_edits if op == "+")
    gl_delete_set = set(seg.strip() for op, seg in gl_edits if op == "-")

    # Base DOCX for styles/sectPr
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

    # Statistikk
    total_changes = 0
    total_inserted_chars = 0
    glossary_changes = 0
    glossary_inserted_chars = 0

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

    # Diff original -> final
    orig_lines = (original_text or "").split("\n")
    fin_lines = (final_text or "").split("\n")
    max_len = max(len(orig_lines), len(fin_lines))

    for i in range(max_len):
        a = orig_lines[i] if i < len(orig_lines) else ""
        b = fin_lines[i] if i < len(fin_lines) else ""

        p = ET.SubElement(new_body, f"{{{W_NS}}}p")

        a_tok = tokenize_keep_ws(a)
        b_tok = tokenize_keep_ws(b)
        edits = diff_tokens(a_tok, b_tok)

        for op, seg in edits:
            if not seg:
                continue
            seg_norm = seg.strip()

            if op == "=":
                _add_text_run(p, seg)

            elif op == "+":
                is_gloss = seg_norm in gl_insert_set
                author = author_gloss if is_gloss else author_lang
                ins_el = ET.SubElement(
                    p, f"{{{W_NS}}}ins",
                    {f"{{{W_NS}}}author": author, f"{{{W_NS}}}date": now_iso, f"{{{W_NS}}}id": str(rev_id)}
                )
                rev_id += 1
                total_changes += 1
                total_inserted_chars += len(seg)
                if is_gloss:
                    glossary_changes += 1
                    glossary_inserted_chars += len(seg)
                _add_text_run(ins_el, seg)

            elif op == "-":
                is_gloss = seg_norm in gl_delete_set
                author = author_gloss if is_gloss else author_lang
                del_el = ET.SubElement(
                    p, f"{{{W_NS}}}del",
                    {f"{{{W_NS}}}author": author, f"{{{W_NS}}}date": now_iso, f"{{{W_NS}}}id": str(rev_id)}
                )
                rev_id += 1
                total_changes += 1
                if is_gloss:
                    glossary_changes += 1
                _add_del_run(del_el, seg)

    new_body.append(sectPr_clone)
    new_xml = ET.tostring(new_root, xml_declaration=True, encoding="UTF-8", standalone="yes")

    out_bio = io.BytesIO()
    with ZipFile(io.BytesIO(base_bytes), "r") as zin, ZipFile(out_bio, "w") as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "word/document.xml":
                data = new_xml
            zout.writestr(item, data)

    return out_bio.getvalue(), total_changes, total_inserted_chars, glossary_changes, glossary_inserted_chars

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

# Tone + modell + aggressivitet
tone_options = list(st.session_state["goals"].keys()) or list(DEFAULT_GOALS.keys())
colA, colB = st.columns(2)
with colA:
    tone = st.selectbox("Tone/retning", tone_options, help="Velg hvordan teksten skal forbedres.")
with colB:
    model_name = st.selectbox("Modell", ["gpt-5", "gpt-5-mini", "gpt-4o", "gpt-4o-mini"], index=0)

strength = st.slider(
    "Aggressivitet (1=minimal, 5=aggressiv)", min_value=1, max_value=5, value=3,
    help="Styrer hvor mange endringer som er ønsket. Appen beregner et budsjett og instruerer modellen."
)

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
use_glossary_in_prompt = st.checkbox(
    "Gi modellen beskjed om å bruke ordlisten (anbefalt)", value=True,
    help="Legger en kort instruks om prefererte termer inn i prompten."
)
author_lang = st.text_input("Forfatter-tag for språk-pass", value="ChatGPT",
                            help="Vises i Word som forfatter av språkendringer.")
author_gloss = st.text_input("Forfatter-tag for ordliste-pass", value="ChatGPT (ordliste)",
                             help="Vises i Word som forfatter av ordliste-endringer.")

stats_placeholder = st.empty()

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

    # Budsjett fra aggressivitetsnivå
    max_changes, max_ins_chars = budget_from_strength(len(uploaded_text or ""), strength)

    # Pass 1: språk-forbedring (med budsjett)
    with st.spinner("Forbedrer tekst (språk-pass) …"):
        try:
            improved = improve_text(
                text=uploaded_text, tone=tone, model_name=model_name, glossary_note=glossary_note,
                strength=strength, max_changes=max_changes, max_insert_chars=max_ins_chars, strict=False
            )
        except Exception as e:
            st.error(f"Feil fra modellen: {e}")
            st.stop()

    # Estimér mot budsjett
    est_changes, est_ins_chars = estimate_changes(uploaded_text or "", improved or "")
    exceeded = (est_changes > max_changes) or (est_ins_chars > max_ins_chars)

    # Streng re-run hvis budsjett overskredet
    if exceeded:
        with st.spinner("Strammer inn til budsjettet …"):
            try:
                improved_strict = improve_text(
                    text=uploaded_text, tone=tone, model_name=model_name, glossary_note=glossary_note,
                    strength=strength, max_changes=max_changes, max_insert_chars=max_ins_chars, strict=True
                )
                # mål på nytt, og ta den som nærmest budsjett
                ch1, ins1 = est_changes, est_ins_chars
                ch2, ins2 = estimate_changes(uploaded_text or "", improved_strict or "")
                # velg det utkastet som ikke overskrider budsjett, ellers det med færrest endringer
                if (ch2 <= max_changes and ins2 <= max_ins_chars) or (ch2 + ins2 < ch1 + ins1):
                    improved = improved_strict
                    est_changes, est_ins_chars = ch2, ins2
                    exceeded = (est_changes > max_changes) or (est_ins_chars > max_ins_chars)
            except Exception as e:
                st.warning(f"Klarte ikke å stramme inn automatisk: {e}")

    # Pass 2: deterministisk ordliste-normalisering + rapport
    if glossary:
        final_text, gloss_report = apply_glossary_with_report(improved, glossary)
    else:
        final_text, gloss_report = improved, {"total": 0, "by_preferred": {}, "pairs": []}

    improved_docx = make_docx_from_text(final_text, "Forbedret tekst")

    # Generer Track Changes (to forfattere)
    with st.spinner("Genererer ekte Track Changes …"):
        try:
            tracked_docx, changes_count, inserted_chars, gl_changes, gl_inserted_chars = make_tracked_changes_docx(
                original_text=uploaded_text or "",
                improved_text=improved or "",
                final_text=final_text or "",
                author_lang=author_lang or "ChatGPT",
                author_gloss=author_gloss or "ChatGPT (ordliste)",
            )
        except Exception as e:
            st.error(f"Klarte ikke å lage Track Changes-dokument: {e}")
            tracked_docx, changes_count, inserted_chars, gl_changes, gl_inserted_chars = None, 0, 0, 0, 0

    st.success("Ferdig!")

    # 📊 Statistikk
    with stats_placeholder.container():
        st.subheader("📊 Resultatstatistikk")
        c1, c2, c3 = st.columns(3)
        with c1:
            st.metric("Foreslåtte endringer (totalt)", f"{changes_count:,}")
        with c2:
            st.metric("Tegn i innsatte forslag (totalt)", f"{inserted_chars:,}")
        with c3:
            st.metric("Ordlistendringer (antall)", f"{gl_changes:,}")

        st.markdown(
            f"- **Budsjett (nivå {strength})**: ≤ **{max_changes}** endringer, ≤ **{max_ins_chars}** innsettingstegn\n"
            f"- **Språk-pass estimerte endringer**: {est_changes:,} / {est_ins_chars:,} tegn "
            + ("✅ innen budsjett" if not exceeded else "⚠️ over budsjett")
        )
        st.caption("Budsjettet styrer modellen via instruksjoner og en automatisk innstrammingsrunde ved behov.")

    # Detaljer om ordlistebruk (valgfritt)
    with st.expander("🔍 Ordliste-bruk (detaljer)"):
        st.write(f"Totalt ordliste-erstatninger: **{gloss_report['total']}**")
        if gloss_report["by_preferred"]:
            rows = [{"Preferert term": k, "Antall": v} for k, v in sorted(gloss_report["by_preferred"].items(), key=lambda x: (-x[1], x[0]))]
            st.table(rows)
        else:
            st.caption("Ingen ordliste-endringer i denne teksten.")

    # Nedlasting (ingen forhåndsvisningstekst i UI som ønsket)
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
    **Aggressivitet:** Slideren styrer et endringsbudsjett og klare regler i prompten. Appen måler resultatet
    og prøver en strengere ny runde hvis budsjettet sprenges.  
    **Farger i Word:** To forfattere (språk/ordliste) → to farger når Word står på “By author”.
    """)
)

