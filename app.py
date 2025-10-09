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
from openai import OpenAI
import difflib
from lxml import etree as ET  # må stå i requirements
from docx import Document     # kun brukt til "ren forbedret" (valgfritt)

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

def _toast(msg: str, icon: str = "✅"):
    try:
        st.toast(msg, icon=icon)
    except Exception:
        st.info(msg)

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

# --- App-konfig i Drive (huske sist brukt ordliste-fil) ---
CONFIG_FILENAME = ".sprakvasker_config.json"

def _load_app_config_from_drive() -> dict | None:
    if not drive_enabled():
        return None
    try:
        service = _drive_service()
        folder_id = _folder_id_from_secret()
        f = drive_find_file(service, folder_id, CONFIG_FILENAME)
        if not f:
            return None
        data = drive_download_bytes(service, f["id"])
        return json.loads(data.decode("utf-8"))
    except Exception:
        return None

def _save_app_config_to_drive(conf: dict) -> None:
    if not drive_enabled():
        return
    service = _drive_service()
    folder_id = _folder_id_from_secret()
    data = json.dumps(conf, ensure_ascii=False, indent=2).encode("utf-8")
    drive_upload_bytes(service, folder_id, CONFIG_FILENAME, data, "application/json")

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
    """
    Erstatter synonymer med foretrukket term og returnerer (ny_tekst, rapport).
    rapport = {"total": int, "by_preferred": {preferred: count, ...}, "pairs": [(src, dst), ...]}
    """
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
    Returner (max_changes, max_insert_chars) basert på tekstlengde og aggressivitet 1–5 (per avsnitt).
    """
    per_1000 = {
        1: (2, 3),
        2: (5, 8),
        3: (10, 15),
        4: (20, 30),
        5: (40, 60),
    }[max(1, min(5, strength))]
    per_1000_changes, per_1000_ins_chars = per_1000
    scale = max(1, int(round(max(1, text_len) / 1000.0)))
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
# Diff & Track Changes utils
# =============================
TOKEN_RE = re.compile(r"\s+|\w+|[^\w\s]", re.UNICODE)

def tokenize_keep_ws(s: str) -> List[str]:
    return TOKEN_RE.findall(s or "")

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

def paragraph_text(p_el, W_NS: str) -> str:
    """Hent ren tekst fra et <w:p>-element (slår sammen alle w:t)."""
    texts = []
    for t in p_el.findall(f".//{{{W_NS}}}t"):
        texts.append(t.text or "")
    return "".join(texts)

def clear_paragraph_content_keep_pPr(p_el, W_NS: str):
    """Fjerner alt i <w:p> unntatt <w:pPr> (beholder avsnittsstil/nummer)."""
    to_remove = []
    pPr = None
    for child in list(p_el):
        if child.tag == f"{{{W_NS}}}pPr":
            pPr = child
        else:
            to_remove.append(child)
    for c in to_remove:
        p_el.remove(c)
    if pPr is not None:
        # sikre at pPr ligger først
        if list(p_el) and list(p_el)[0] is not pPr:
            p_el.remove(pPr)
            p_el.insert(0, pPr)

def add_text_run(parent, W_NS: str, XML_NS: str, text: str):
    r = ET.SubElement(parent, f"{{{W_NS}}}r")
    t = ET.SubElement(r, f"{{{W_NS}}}t")
    if text.startswith(" ") or text.endswith(" "):
        t.set(f"{{{XML_NS}}}space", "preserve")
    # håndter linjeskift: erstatte '\n' med <w:br/>
    parts = text.split("\n")
    if len(parts) == 1:
        t.text = text
    else:
        # Første del i denne <w:t>
        t.text = parts[0]
        for part in parts[1:]:
            ET.SubElement(r, f"{{{W_NS}}}br")
            t2 = ET.SubElement(r, f"{{{W_NS}}}t")
            if part.startswith(" ") or part.endswith(" "):
                t2.set(f"{{{XML_NS}}}space", "preserve")
            t2.text = part

def add_del_run(parent, W_NS: str, XML_NS: str, text: str):
    r = ET.SubElement(parent, f"{{{W_NS}}}r")
    dt = ET.SubElement(r, f"{{{W_NS}}}delText")
    if text.startswith(" ") or text.endswith(" "):
        dt.set(f"{{{XML_NS}}}space", "preserve")
    parts = text.split("\n")
    if len(parts) == 1:
        dt.text = text
    else:
        # delText støtter ikke br direkte, men vi kan bare sette '\n' inn – Word viser del-sone på linjen.
        dt.text = text

def rewrite_paragraph_with_track_changes(
    p_el,
    orig_text: str,
    improved_text: str,
    final_text: str,
    authors: Tuple[str, str],
    W_NS: str,
    XML_NS: str,
    now_iso: str,
    start_rev_id: int
) -> Tuple[int, int, int, int]:
    """
    Skriver inn track changes i det opprinnelige <w:p>-elementet.
    Returnerer (ny_rev_id, total_changes, total_insert_chars, glossary_changes)
    """
    author_lang, author_gloss = authors

    # Identifiser ordliste-endringer (improved -> final)
    imp_ops = diff_tokens(tokenize_keep_ws(improved_text), tokenize_keep_ws(final_text))
    gl_ins = set(seg.strip() for op, seg in imp_ops if op == "+")
    gl_del = set(seg.strip() for op, seg in imp_ops if op == "-")

    # Diff original -> final
    ops = diff_tokens(tokenize_keep_ws(orig_text), tokenize_keep_ws(final_text))

    # Tøm innhold (behold pPr)
    clear_paragraph_content_keep_pPr(p_el, W_NS)

    rev_id = start_rev_id
    total_changes = 0
    total_insert_chars = 0
    glossary_changes = 0

    for op, seg in ops:
        if not seg:
            continue
        seg_norm = seg.strip()

        if op == "=":
            add_text_run(p_el, W_NS, XML_NS, seg)

        elif op == "+":
            is_gloss = seg_norm in gl_ins
            author = author_gloss if is_gloss else author_lang
            ins_el = ET.SubElement(
                p_el, f"{{{W_NS}}}ins",
                {f"{{{W_NS}}}author": author, f"{{{W_NS}}}date": now_iso, f"{{{W_NS}}}id": str(rev_id)}
            )
            rev_id += 1
            total_changes += 1
            total_insert_chars += len(seg)
            if is_gloss:
                glossary_changes += 1
            add_text_run(ins_el, W_NS, XML_NS, seg)

        elif op == "-":
            is_gloss = seg_norm in gl_del
            author = author_gloss if is_gloss else author_lang
            del_el = ET.SubElement(
                p_el, f"{{{W_NS}}}del",
                {f"{{{W_NS}}}author": author, f"{{{W_NS}}}date": now_iso, f"{{{W_NS}}}id": str(rev_id)}
            )
            rev_id += 1
            total_changes += 1
            if is_gloss:
                glossary_changes += 1
            add_del_run(del_el, W_NS, XML_NS, seg)

    return rev_id, total_changes, total_insert_chars, glossary_changes

# =============================
# Per-avsnitt prosess (Valg B)
# =============================
def process_docx_with_track_changes(
    src_docx_bytes: bytes,
    tone: str,
    model_name: str,
    glossary: Dict[str, List[str]],
    include_gloss_in_prompt: bool,
    strength: int,
    author_lang: str,
    author_gloss: str,
) -> Tuple[bytes, bytes, Dict[str, int], Dict[str, int]]:
    """
    Åpner original .docx og skriver endringer inn i samme document.xml.
    Returnerer (tracked_docx_bytes, clean_docx_bytes, totals, gloss_totals)
    totals: {"changes": int, "insert_chars": int}
    gloss_totals: {"changes": int, "insert_chars": int}
    """
    # For prompt: liste kun de første 100 prefererte
    glossary_note = None
    if include_gloss_in_prompt and glossary:
        preferred_list = "\n".join(f"- {p}" for p in list(glossary.keys())[:100])
        glossary_note = (
            "Bruk følgende prefererte termer konsekvent; normaliser eventuelle synonymer/varianter til eksakt skrivemåte:\n"
            f"{preferred_list}"
        )

    # Pakk ut original document.xml
    zin = ZipFile(io.BytesIO(src_docx_bytes), "r")
    doc_xml = zin.read("word/document.xml")
    root = ET.fromstring(doc_xml)

    # Namespaces
    nsmap = root.nsmap.copy()
    W_NS = nsmap.get("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main")
    XML_NS = "http://www.w3.org/XML/1998/namespace"

    # Finn alle <w:p> i dokumentet (inkluderer tabeller)
    paragraphs = root.findall(f".//{{{W_NS}}}p")

    now_iso = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    rev_id = 1

    total_changes = 0
    total_insert_chars = 0
    glossary_changes_total = 0
    glossary_insert_chars_total = 0

    # For "ren forbedret"-variant: vi lager en kopiert XML der vi skriver inn endelig tekst uten track changes
    clean_root = ET.fromstring(doc_xml)  # separat tre

    for p_idx, p in enumerate(paragraphs):
        # Hent original tekst
        orig = paragraph_text(p, W_NS)
        if not orig.strip():
            continue

        # Budsjett per avsnitt
        max_changes, max_ins_chars = budget_from_strength(len(orig), strength)

        # Språk-pass (LLM), med ev. innstrammingsrunde
        try:
            improved = improve_text(
                text=orig, tone=tone, model_name=model_name, glossary_note=glossary_note,
                strength=strength, max_changes=max_changes, max_insert_chars=max_ins_chars, strict=False
            )
        except Exception as e:
            # Hvis modellkallet feiler for ett avsnitt, fall tilbake til original
            improved = orig

        # Estimering
        a = tokenize_keep_ws(orig)
        b = tokenize_keep_ws(improved)
        ops_est = diff_tokens(a, b)
        est_changes = sum(1 for op, _ in ops_est if op in {"+", "-"})
        est_insert_chars = sum(len(seg) for op, seg in ops_est if op == "+")

        if est_changes > max_changes or est_insert_chars > max_ins_chars:
            # Streng runde
            try:
                improved_strict = improve_text(
                    text=orig, tone=tone, model_name=model_name, glossary_note=glossary_note,
                    strength=strength, max_changes=max_changes, max_insert_chars=max_ins_chars, strict=True
                )
                a2 = tokenize_keep_ws(orig)
                b2 = tokenize_keep_ws(improved_strict)
                ops2 = diff_tokens(a2, b2)
                ch2 = sum(1 for op, _ in ops2 if op in {"+", "-"})
                ins2 = sum(len(seg) for op, seg in ops2 if op == "+")
                if (ch2 <= max_changes and ins2 <= max_ins_chars) or (ch2 + ins2 < est_changes + est_insert_chars):
                    improved = improved_strict
            except Exception:
                pass  # behold improved

        # Ordliste-pass (deterministisk)
        if glossary:
            final_text, gloss_report = apply_glossary_with_report(improved, glossary)
        else:
            final_text, gloss_report = improved, {"total": 0, "by_preferred": {}, "pairs": []}

        # Skriv Track Changes inn i original-XML
        rev_id, ch, ins_chars, gl_ch = rewrite_paragraph_with_track_changes(
            p_el=p,
            orig_text=orig,
            improved_text=improved,
            final_text=final_text,
            authors=(author_lang, author_gloss),
            W_NS=W_NS,
            XML_NS=XML_NS,
            now_iso=now_iso,
            start_rev_id=rev_id
        )
        total_changes += ch
        total_insert_chars += ins_chars
        glossary_changes_total += gl_ch
        # For innsettings-tegn pga ordliste: vi kan grovt anslå via gloss_report["pairs"], men for enkelhet teller vi samme
        # sum som i track changes (innsettings-tegn er allerede i total_insert_chars). Her legger vi ikke duplikat.

        # Skriv "ren" tekst i clean_root (uten track changes)
        clean_p = clean_root.findall(f".//{{{W_NS}}}p")[p_idx]
        clear_paragraph_content_keep_pPr(clean_p, W_NS)
        add_text_run(clean_p, W_NS, XML_NS, final_text)

    # Bygg ny DOCX for tracked-varianten
    out_tracked = io.BytesIO()
    with ZipFile(out_tracked, "w") as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "word/document.xml":
                data = ET.tostring(root, xml_declaration=True, encoding="UTF-8", standalone="yes")
            zout.writestr(item, data)
    zin.close()

    # Bygg ny DOCX for clean-varianten
    out_clean = io.BytesIO()
    with ZipFile(io.BytesIO(src_docx_bytes), "r") as zin2, ZipFile(out_clean, "w") as zout2:
        for item in zin2.infolist():
            data = zin2.read(item.filename)
            if item.filename == "word/document.xml":
                data = ET.tostring(clean_root, xml_declaration=True, encoding="UTF-8", standalone="yes")
            zout2.writestr(item, data)

    totals = {"changes": total_changes, "insert_chars": total_insert_chars}
    gloss_totals = {"changes": glossary_changes_total, "insert_chars": 0}  # innsettings-tegn per ordliste deles ikke ut separat nå

    return out_tracked.getvalue(), out_clean.getvalue(), totals, gloss_totals

# =============================
# Auto-last ordliste ved oppstart + husk sist brukte filnavn
# =============================
if "drive_glossary_filename" not in st.session_state:
    conf = _load_app_config_from_drive()
    last_name = (conf or {}).get("glossary_filename")
    st.session_state["drive_glossary_filename"] = last_name or "ordliste02okt.csv"

if drive_enabled() and "glossary_autoload_done" not in st.session_state:
    try:
        g = load_glossary_from_drive(st.session_state["drive_glossary_filename"])
        if g:
            st.session_state["glossary"] = g
            st.session_state["glossary_autoload_done"] = True
            _toast(f"Ordlisten «{st.session_state['drive_glossary_filename']}» ble lastet automatisk.")
        else:
            st.session_state["glossary_autoload_done"] = True
    except Exception:
        st.session_state["glossary_autoload_done"] = True

# =============================
# UI: input
# =============================
tab1, tab2 = st.tabs(["📄 Word-fil (.docx)", "✍️ Lim inn tekst (lages som ny fil)"])
uploaded_docx_bytes: bytes | None = None
pasted_text: str | None = None

with tab1:
    up = st.file_uploader("Last opp .docx", type=["docx"])
    if up is not None:
        try:
            uploaded_docx_bytes = up.read()
        except Exception as e:
            st.error(f"Kunne ikke lese Word-filen: {e}")

with tab2:
    pasted_text = st.text_area("Lim inn tekst her", height=300, placeholder="Lim inn manus her …")
    st.caption("Hvis du limer inn tekst, bygger vi en ny Word-fil fra teksten og skriver endringer der.")

# Tone + modell + aggressivitet
tone_options = list(st.session_state["goals"].keys()) or list(DEFAULT_GOALS.keys())
colA, colB = st.columns(2)
with colA:
    tone = st.selectbox("Tone/retning", tone_options, help="Velg hvordan teksten skal forbedres.")
with colB:
    # Standard gpt-5-mini (index=1)
    model_name = st.selectbox("Modell", ["gpt-5", "gpt-5-mini", "gpt-4o", "gpt-4o-mini"], index=1)

strength = st.slider(
    "Aggressivitet (1=minimal, 5=aggressiv)", min_value=1, max_value=5, value=3,
    help="Styrer hvor mange endringer som er ønsket. Budsjett brukes per avsnitt."
)

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
# Ordliste-UI (kompakt)
# =============================
with st.expander("📚 Ordliste (last opp / Drive / lagre)"):
    drive_ok = drive_enabled()
    st.info(("Google Drive er konfigurert." if drive_ok else
             "Google Drive er ikke konfigurert (legg inn GDRIVE_SERVICE_ACCOUNT_JSON og GDRIVE_FOLDER_ID i Secrets)."))

    default_name = st.text_input(
        "Filnavn i Drive",
        key="drive_glossary_filename",
        help="Brukes når du leser/lagrer mot Drive."
    )

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
                g = load_glossary_from_drive(st.session_state["drive_glossary_filename"])
                if g is None:
                    st.warning(f"Fant ikke «{st.session_state['drive_glossary_filename']}» i Drive-mappen.")
                else:
                    st.session_state["glossary"] = g
                    _save_app_config_to_drive({"glossary_filename": st.session_state["drive_glossary_filename"]})
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
                    save_glossary_to_drive(gloss, st.session_state["drive_glossary_filename"])
                    _save_app_config_to_drive({"glossary_filename": st.session_state["drive_glossary_filename"]})
                    st.success(f"Lagret ordliste som «{st.session_state['drive_glossary_filename']}» i Drive-mappen.")
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
    "⚙️ Forbedre språk og lag Ekte Track Changes (i originalfilen)",
    type="primary",
    disabled=(not OPENAI_API_KEY or not (uploaded_docx_bytes or pasted_text)),
)

if run_btn:
    if not OPENAI_API_KEY:
        st.error("OPENAI_API_KEY mangler.")
        st.stop()

    glossary = st.session_state.get("glossary", {})

    # Kilde: enten lastet .docx eller tekst (da lager vi først en enkel .docx med teksten)
    if pasted_text and not uploaded_docx_bytes:
        # bygg en enkel docx fra limt tekst (ett avsnitt per linje)
        d = Document()
        for line in (pasted_text or "").split("\n"):
            d.add_paragraph(line)
        bio = io.BytesIO()
        d.save(bio)
        uploaded_docx_bytes = bio.getvalue()

    with st.spinner("Prosesserer dokument (per avsnitt) …"):
        try:
            tracked_bytes, clean_bytes, totals, gloss_totals = process_docx_with_track_changes(
                src_docx_bytes=uploaded_docx_bytes,
                tone=tone,
                model_name=model_name,
                glossary=glossary,
                include_gloss_in_prompt=use_glossary_in_prompt,
                strength=strength,
                author_lang=author_lang or "ChatGPT",
                author_gloss=author_gloss or "ChatGPT (ordliste)",
            )
        except Exception as e:
            st.error(f"Klarte ikke å generere Track Changes i originalfilen: {e}")
            st.stop()

    st.success("Ferdig!")

    # 📊 Statistikk
    with stats_placeholder.container():
        st.subheader("📊 Resultatstatistikk")
        c1, c2, c3 = st.columns(3)
        with c1:
            st.metric("Foreslåtte endringer (totalt)", f"{totals.get('changes', 0):,}")
        with c2:
            st.metric("Tegn i innsatte forslag (totalt)", f"{totals.get('insert_chars', 0):,}")
        with c3:
            st.metric("Ordlistendringer (antall)", f"{gloss_totals.get('changes', 0):,}")
        st.caption("Budsjett er anvendt per avsnitt. Word farger endringer “By author” (Review → Change Tracking Options).")

    st.download_button(
        "📝 Last ned med ekte Spor endringer (.docx)",
        data=tracked_bytes,
        file_name="forbedret_spor_endringer_ekte.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )

    st.download_button(
        "💾 Last ned ren forbedret Word (.docx)",
        data=clean_bytes,
        file_name="forbedret_ren.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )

st.markdown("---")
st.markdown(
    textwrap.dedent("""
    **Valg B aktiv:** Endringer skrives inn i originalens avsnitt (også i tabeller).  
    **Begrensning:** Inline-stil (fet/kursiv) på *endrede* tekstbiter kan gå tapt. Dokumentets stiler/nummerering bevares.  
    **Farger i Word:** To forfattere (språk/ordliste) → to farger når Word står på “By author”.
    """)
)
