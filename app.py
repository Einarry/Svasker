import base64
import html as html_lib
import io
import re
import subprocess
import tempfile
import zipfile
from pathlib import Path
from email import policy
from email.parser import BytesParser

import streamlit as st
from bs4 import BeautifulSoup

from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload


# --------------------------------------------------
# Streamlit setup
# --------------------------------------------------

st.set_page_config(
    page_title="Tidsskriftet PDF generator",
    page_icon="📄",
    layout="centered",
)

st.title("Tidsskriftet-style PDF generator")

st.write(
    "Upload an MHTML article and logo. Edit CSS from Google Drive, then generate "
    "a print-ready A4 PDF with two columns and page numbers."
)


# --------------------------------------------------
# Google Drive CSS setup
# --------------------------------------------------

SCOPES = ["https://www.googleapis.com/auth/drive"]


DEFAULT_CSS = """
@page {
  size: A4;
  margin: 15mm 14mm 17mm 14mm;

  @bottom-right {
    content: counter(page) " / " counter(pages);
    font-family: Georgia, "Times New Roman", serif;
    font-size: 10px;
    color: #111;
  }
}

body {
  margin: 0;
  color: #111;
  background: #fff;
  font-family: Georgia, "Times New Roman", serif;
}
"""


def get_drive_service():
    creds = service_account.Credentials.from_service_account_info(
        st.secrets["gcp_service_account"],
        scopes=SCOPES,
    )
    return build("drive", "v3", credentials=creds)


def load_css_from_drive():
    service = get_drive_service()
    file_id = st.secrets["css_file_id"]

    request = service.files().get_media(fileId=file_id)
    css_bytes = request.execute()

    return css_bytes.decode("utf-8")


def save_css_to_drive(css_text):
    service = get_drive_service()
    file_id = st.secrets["css_file_id"]

    media = MediaIoBaseUpload(
        io.BytesIO(css_text.encode("utf-8")),
        mimetype="text/css",
        resumable=False,
    )

    service.files().update(
        fileId=file_id,
        media_body=media,
    ).execute()


# --------------------------------------------------
# Helper functions
# --------------------------------------------------

def normalize_whitespace(text: str) -> str:
    return re.sub(r"\s+", " ", text or "").strip()


def data_url_from_bytes(filename: str, content: bytes) -> str:
    suffix = Path(filename).suffix.lower()

    mime = {
        ".jpg": "image/jpeg",
        ".jpeg": "image/jpeg",
        ".png": "image/png",
        ".svg": "image/svg+xml",
        ".webp": "image/webp",
    }.get(suffix, "application/octet-stream")

    encoded = base64.b64encode(content).decode("ascii")
    return f"data:{mime};base64,{encoded}"


def parse_mhtml(mhtml_bytes: bytes):
    msg = BytesParser(policy=policy.default).parsebytes(mhtml_bytes)

    html_part = next(
        part for part in msg.walk()
        if part.get_content_type() == "text/html"
    )

    source_html = html_part.get_payload(decode=True).decode(
        html_part.get_content_charset() or "utf-8",
        "replace",
    )

    resources = {}

    for part in msg.walk():
        loc = part.get("Content-Location")
        payload = part.get_payload(decode=True)
        ctype = part.get_content_type()

        if loc and payload and ctype.startswith("image/"):
            encoded = base64.b64encode(payload).decode("ascii")
            resources[loc] = f"data:{ctype};base64,{encoded}"

    return source_html, resources


# --------------------------------------------------
# Build final HTML
# --------------------------------------------------

def build_print_html(
    mhtml_bytes: bytes,
    logo_filename: str,
    logo_bytes: bytes,
    css_text: str,
) -> str:
    source_html, resources = parse_mhtml(mhtml_bytes)
    source_soup = BeautifulSoup(source_html, "html.parser")

    article = source_soup.select_one("article.scientific-article--full")
    if not article:
        raise RuntimeError("Could not find article.scientific-article--full")

    body = article.select_one(".field--name-body")
    if not body:
        raise RuntimeError("Could not find .field--name-body")

    title_el = article.find("h1")
    title = normalize_whitespace(title_el.get_text(" ", strip=True)) if title_el else "Article"

    channel_el = article.select_one(".field--name-field-channel")
    channel = normalize_whitespace(channel_el.get_text(" ", strip=True)) if channel_el else "Originalartikkel"

    byline_el = article.select_one(".field--name-field-byline")
    byline = normalize_whitespace(byline_el.get_text(" ", strip=True)) if byline_el else ""

    content = BeautifulSoup(str(body), "html.parser")

    for el in content.select(
        "button, svg, nav, .visually-hidden, .contextual, .js-contextual-links"
    ):
        el.decompose()

    for img in content.find_all("img"):
        src = img.get("src", "")

        if src in resources:
            img["src"] = resources[src]

        img.attrs.pop("loading", None)
        img.attrs.pop("srcset", None)
        img.attrs.pop("sizes", None)

    for a in content.find_all("a", href=True):
        href = a["href"]
        if href.startswith("/"):
            a["href"] = "https://tidsskriftet.no" + href

    logo_src = data_url_from_bytes(logo_filename, logo_bytes)

    final_html = f"""<!doctype html>
<html lang="no">
<head>
<meta charset="utf-8">
<title>{html_lib.escape(title)} - A4 print</title>

<style>
{css_text}
</style>
</head>

<body>
<div class="print-page">

  <header class="article-header">
    <div class="logo-wrap">
      <img src="{logo_src}" alt="Tidsskriftet">
    </div>

    <div class="kicker">{html_lib.escape(channel)}</div>

    <h1>{html_lib.escape(title)}</h1>

    <div class="meta">{html_lib.escape(byline)}</div>
  </header>

  <article class="article-columns">
    {str(content)}
  </article>

</div>
</body>
</html>
"""

    return final_html


# --------------------------------------------------
# Render PDF
# --------------------------------------------------

def render_pdf_with_weasyprint(html_text: str, output_dir: Path):
    html_path = output_dir / "article_A4_two_column.html"
    pdf_path = output_dir / "article_A4_two_column.pdf"
    zip_path = output_dir / "article_A4_two_column_package.zip"

    html_path.write_text(html_text, encoding="utf-8")

    subprocess.run(
        ["weasyprint", str(html_path), str(pdf_path)],
        check=True,
        capture_output=True,
        text=True,
    )

    with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED) as z:
        z.write(html_path, arcname=html_path.name)
        z.write(pdf_path, arcname=pdf_path.name)

    return html_path, pdf_path, zip_path


# --------------------------------------------------
# App interface
# --------------------------------------------------

st.header("1. Upload files")

mhtml_file = st.file_uploader(
    "Upload MHTML article",
    type=["mhtml", "mht"],
)

logo_file = st.file_uploader(
    "Upload logo",
    type=["jpg", "jpeg", "png", "svg", "webp"],
)


st.header("2. Edit CSS")

try:
    current_css = load_css_from_drive()
except Exception as e:
    st.warning("Could not load CSS from Google Drive. Using fallback CSS.")
    st.code(str(e))
    current_css = DEFAULT_CSS

css_text = st.text_area(
    "CSS used for PDF styling",
    value=current_css,
    height=500,
)

col1, col2 = st.columns(2)

with col1:
    if st.button("Save CSS to Google Drive"):
        try:
            save_css_to_drive(css_text)
            st.success("CSS saved to Google Drive.")
        except Exception as e:
            st.error("Could not save CSS to Google Drive.")
            st.code(str(e))

with col2:
    if st.button("Reset editor to fallback CSS"):
        css_text = DEFAULT_CSS
        st.info("Editor reset. Click save if you want to store this in Google Drive.")


st.header("3. Generate PDF")

if st.button("Generate PDF", type="primary"):
    if not mhtml_file:
        st.error("Please upload an MHTML file.")
        st.stop()

    if not logo_file:
        st.error("Please upload a logo file.")
        st.stop()

    try:
        with st.spinner("Generating PDF..."):
            html_text = build_print_html(
                mhtml_bytes=mhtml_file.read(),
                logo_filename=logo_file.name,
                logo_bytes=logo_file.read(),
                css_text=css_text,
            )

            with tempfile.TemporaryDirectory() as tmp:
                tmp_path = Path(tmp)

                html_path, pdf_path, zip_path = render_pdf_with_weasyprint(
                    html_text=html_text,
                    output_dir=tmp_path,
                )

                pdf_bytes = pdf_path.read_bytes()
                html_bytes = html_path.read_bytes()
                zip_bytes = zip_path.read_bytes()

        st.success("PDF generated.")

        st.download_button(
            label="Download PDF",
            data=pdf_bytes,
            file_name="article_A4_two_column.pdf",
            mime="application/pdf",
        )

        st.download_button(
            label="Download HTML",
            data=html_bytes,
            file_name="article_A4_two_column.html",
            mime="text/html",
        )

        st.download_button(
            label="Download ZIP package",
            data=zip_bytes,
            file_name="article_A4_two_column_package.zip",
            mime="application/zip",
        )

    except subprocess.CalledProcessError as e:
        st.error("WeasyPrint failed to generate the PDF.")
        st.code(e.stderr or str(e))

    except Exception as e:
        st.error("Could not generate PDF.")
        st.code(str(e))