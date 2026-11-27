import io
import re
import difflib
import json
from datetime import datetime

import streamlit as st
import gspread
from docxtpl import DocxTemplate
from docx.shared import Cm
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload

# =====================================================================
# CONFIG
# =====================================================================
DRIVE_FOLDER_ID_WORD = "1Iqf7sx4E104uKqqkAfTHReYVeLrFo_b3"
DRIVE_FOLDER_ID_PDF  = "1OXDu3C-RendKlRz0-HbTFHRA1XlY1JVR"

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

# =====================================================================
# CANONICAL LECTURER REFERENCE
# Source: Data_NIP_Terbaru_13Nov2026.xlsx (Data terbaru + Tambahan resmi)
# =====================================================================
CANONICAL_LECTURERS = [
    "Prof. Dr. Ir. Adrianto Ahmad, M.T., IPU, INV",
    "Prof. Ahmad Fadli, S.T., M.T., Ph.D.",
    "Prof. Amun Amri, S.T., M.T., Ph.D.",
    "Aryo Sasmita, S.T., M.Si.",
    "Prof. Dr. Ir. Bahruddin, S.T., M.T.",
    "Chairul, S.T., M.T.",
    "Dr. David Andrio, S.T., M.Si.",
    "Dra. Drastinawati, M.Si.",
    "Dr. Desi Heltina, S.T., M.T.",
    "Drs. Edward HS., M.S.",
    "Prof. Edy Saputra, S.T., M.T., Ph.D.",
    "Elvie Yenie, S.T., M.Eng.",
    "Prof. Evelyn, S.T., M.Sc., M.Eng., Ph.D.",
    "Dr. Hafidawati, S.T.P., M.T.",
    "Prof. Dr. Ida Zahrina, S.T., M.T.",
    "Idral Amri, S.T., M.T., Ph.D.",
    "Ivnaini Andesgur, S.T., M.Si.",
    "Jecky Asmura, S.T., M.T.",
    "Dr. Jhon Armedi Pinem, S.T., M.T.",
    "Dra. Khairat, M.Si.",
    "Komalasari, S.T., M.T.",
    "Dr. Maria Peratenta Sembiring, S.T., M.T.",
    "Muhammad Reza, S.T., M.Sc.",
    "Muhammad Iwan Fermi, S.T., M.T.",
    "Dra. Nirwana, M.T.",
    "Prof. Dr. Padil, S.T., M.T.",
    "Panca Setia Utama, S.T., M.T., Ph.D.",
    "Ir. Rozanna Sri Irianty, M.Si.",
    "Dr. Ir. Said Zul Amraini, S.T., M.T.",
    "Dra. Silvia Reni Yenti, M.Si.",
    "Shinta Elystia, S.T., M.Si.",
    "Sri Helianty, S.T., M.T.",
    "Sri Rezeki Muria, S.T., M.P., M.Sc.",
    "Syelvia Putri Utami, S.T., M.Eng., Ph.D.",
    "Prof. Dr. Sunarno, S.T., M.T.",
    "Prof. Syaiful Bahri, M.Si., Ph.D.",
    "Drs. Syamsu Herman, M.T.",
    "Ir. Syarfi Daud, M.T.",
    "Prof. Dr. Yelmida Azis, M.Si.",
    "Dra. Yusnimar, M.Si., M.Phil.",
    "Zuqni Meldha, S.T., M.T.",
    "Prof. Zuchra Helwani, S.T., M.T., Ph.D.",
    "Zulfansyah, S.T., M.T.",
    "Anisa Mutamima, B.Eng., M.Eng.",
    "Gunadi Priyambada, S.T., M.T.",
    "Nurfatihayati, S.T., M.T.",
    "Dewi Fitria, S.T., M.T., Ph.D.",
    "Dr. Lita Darmayanti, S.T., M.T.",
    "Hari Rionaldo, S.T., M.T., C.EIA",
    "Cory Dian Alfarisi, S.T., M.T.",
    "Salma Liska, S.Tr.T., M.T.",
    "Yogi Yolanda, S.T., M.T.",
    "Lisa Legawati, S.T., M.T.",
    "Yola Bertilsya Hendri, S.T., M.T.",
    "Esty Octiana Sari, S.T., M.T., Ph.D.",
    "Vayolla Syakirah Karil, M.T.",
    "Nesa Zafira, S.T., M.T.",
    "Suhendri, S.T., M.T.",
    "Alltop Amri Ya Habib, S.T., M.T.",
]

# =====================================================================
# NAME NORMALIZATION (preserved from desktop app)
# =====================================================================
_DEGREE_PATTERNS = [
    r'\bS\.?\s*Tr\.?\s*T\.?\b', r'\bS\.?\s*T\.?\s*P\.?\b',
    r'\bM\.?\s*Phil\.?\b', r'\bPh\.?\s*D\.?\b',
    r'\bB\.?\s*Eng\.?\b', r'\bM\.?\s*Eng\.?\b', r'\bC\.?\s*EIA\b',
    r'\bM\.?\s*Si\.?\b', r'\bS\.?\s*Si\.?\b', r'\bM\.?\s*Sc\.?\b',
    r'\bM\.?\s*P\.?\b', r'\bS\.?\s*T\.?\b', r'\bM\.?\s*T\.?\b',
    r'\bM\.?\s*S\.?\b', r'\bIPU\b', r'\bINV\b',
]
_TITLE_PATTERNS = [
    r'\bProf\b\.?', r'\bDr\b\.?', r'\bIr\b\.?', r'\bDra\b\.?', r'\bDrs\b\.?',
]

def clean_text(text):
    if not text:
        return ""
    text = re.sub(r'<[^>]*>', ' ', text)
    return " ".join(text.strip().strip('"\'').split())

def extract(pattern, text):
    m = re.search(pattern, text, re.IGNORECASE)
    return clean_text(m.group(1)) if m else ""

def extract_role(raw_text, role, num):
    if num == 1:
        num_pat = r'(?:1|I(?![I]))'
    else:
        num_pat = r'(?:2|II)'
    pat = rf'{role}\s*{num_pat}\s*[:.]?\s*(.+)'
    m = re.search(pat, raw_text, re.IGNORECASE)
    return clean_text(m.group(1)) if m else ""

def extract_name_nim(full_string):
    if not full_string:
        return "", ""
    m = re.search(r'(.*?)\s*\((.*?)\)', full_string)
    if m:
        return m.group(1).strip(), m.group(2).strip()
    return full_string.strip(), ""

def normalize_name(name):
    if not isinstance(name, str) or not name:
        return ""
    name = name.strip()
    for pat in _DEGREE_PATTERNS:
        name = re.sub(pat, ' ', name, flags=re.IGNORECASE)
    for pat in _TITLE_PATTERNS:
        name = re.sub(pat, ' ', name, flags=re.IGNORECASE)
    return re.sub(r'[^a-zA-Z]', '', name).lower()

_CANON_MAP = None

def _get_canon_map():
    global _CANON_MAP
    if _CANON_MAP is None:
        _CANON_MAP = {normalize_name(c): c for c in CANONICAL_LECTURERS}
    return _CANON_MAP

def match_lecturer_name(raw_name):
    if not raw_name:
        return raw_name
    raw_name = clean_text(raw_name)
    if not raw_name:
        return raw_name
    norm = normalize_name(raw_name)
    if not norm:
        return raw_name
    canon = _get_canon_map()
    if norm in canon:
        return canon[norm]
    for cutoff in (0.85, 0.70):
        matches = difflib.get_close_matches(norm, canon.keys(), n=1, cutoff=cutoff)
        if matches:
            return canon[matches[0]]
    return raw_name

# =====================================================================
# GOOGLE SERVICES (service account via Streamlit secrets)
# =====================================================================
@st.cache_resource(show_spinner=False)
def get_google_clients():
    sa_info = dict(st.secrets["gcp_service_account"])
    creds = Credentials.from_service_account_info(sa_info, scopes=SCOPES)
    gs_client = gspread.authorize(creds)
    drive = build("drive", "v3", credentials=creds)
    return gs_client, drive

def upload_bytes(drive, name, mime, data_bytes, parent_id):
    media = MediaIoBaseUpload(
        io.BytesIO(data_bytes), mimetype=mime, resumable=False
    )
    meta = {"name": name, "parents": [parent_id]}
    return drive.files().create(
        body=meta, media_body=media, fields="id"
    ).execute()

def convert_docx_to_pdf_via_drive(drive, docx_bytes, temp_parent_id):
    """Upload docx as Google Doc → export PDF → delete temp Google Doc."""
    temp_media = MediaIoBaseUpload(
        io.BytesIO(docx_bytes),
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        resumable=False,
    )
    temp_meta = {
        "name": "_temp_semhas_conversion",
        "mimeType": "application/vnd.google-apps.document",
        "parents": [temp_parent_id],
    }
    temp = drive.files().create(
        body=temp_meta, media_body=temp_media, fields="id"
    ).execute()
    try:
        pdf_bytes = drive.files().export_media(
            fileId=temp["id"], mimeType="application/pdf"
        ).execute()
    finally:
        drive.files().delete(fileId=temp["id"]).execute()
    return pdf_bytes

# =====================================================================
# CORE PROCESSING (same logic as desktop app)
# =====================================================================
def process(raw_text: str):
    students = []
    block_match = re.search(
        r'Mahasiswa.*?:(.*?)(?:Pembimbing|Judul|Hari)',
        raw_text, re.IGNORECASE | re.DOTALL,
    )
    block = block_match.group(1) if block_match else raw_text
    block = re.sub(r'\s+', ' ', block)

    pattern = r'(?:Nama\s*:|\d+\.)?\s*([A-Za-z\s\.,\'-]+?)\s*(?:\(\s*)?(?:NIM\s*)?(\d{9,15})\b'
    for name, nim in re.findall(pattern, block, re.IGNORECASE):
        cn = re.sub(r'^(?:Nama\s*:|Mahasiswa\s*:\s*)', '', clean_text(name), flags=re.I).strip()
        students.append((cn, nim))

    if not students:
        for n in re.findall(r'Nama[^\:]*:\s*(.*)', raw_text, re.IGNORECASE):
            students.append(extract_name_nim(n))

    nama_1 = students[0][0] if len(students) > 0 else ""
    nim_1  = students[0][1] if len(students) > 0 else ""
    nama_2 = students[1][0] if len(students) > 1 else ""
    nim_2  = students[1][1] if len(students) > 1 else ""

    nama_1_val = f"{nama_1} ({nim_1})" if nama_1 else ""
    nama_2_val = f"{nama_2} ({nim_2})" if nama_2 else ""

    nomor         = extract(r'(?:Nomor|Ns|No)\s*[:.]?\s*(\d+)', raw_text)
    judul         = extract(r'Judul[^\:]*:\s*(.*)', raw_text)
    hari_tanggal  = extract(r'Hari\s*/\s*Tanggal\s*:\s*(.*)', raw_text)
    tempat        = extract(r'Tempat\s*:\s*(.*)', raw_text)

    pembimbing_1 = match_lecturer_name(extract_role(raw_text, r'Pembimbing', 1))
    pembimbing_2 = match_lecturer_name(extract_role(raw_text, r'Pembimbing', 2))
    penguji_1    = match_lecturer_name(extract_role(raw_text, r'Penguji', 1))
    penguji_2    = match_lecturer_name(extract_role(raw_text, r'Penguji', 2))

    waktu_raw = extract(r'(?:Waktu|Pukul)\s*:\s*(.*)', raw_text)
    waktu_clean = re.sub(r'(?i)\bwib\b\.?', '', waktu_raw).strip().rstrip('.,;:').strip()
    waktu_word = re.sub(r'\s*-\s*', ' - ', " ".join(waktu_clean.split()))
    waktu_sheets = f"{waktu_word} WIB" if waktu_word else ""

    # --- Render DOCX in memory ---
    bulan = {1:"Januari",2:"Februari",3:"Maret",4:"April",5:"Mei",6:"Juni",
             7:"Juli",8:"Agustus",9:"September",10:"Oktober",11:"November",12:"Desember"}
    today = datetime.now()
    context = {
        "nomor": nomor, "nama_1": nama_1_val, "nama_2": nama_2_val,
        "pembimbing_1": pembimbing_1, "pembimbing_2": pembimbing_2,
        "penguji_1": penguji_1, "penguji_2": penguji_2,
        "judul": judul, "hari_tanggal": hari_tanggal,
        "waktu": waktu_word, "tempat": tempat,
        "tanggal_surat": f"{today.day} {bulan[today.month]} {today.year}",
    }

    doc = DocxTemplate("template_undangan.docx")
    doc.render(context)
    for section in doc.docx.sections:
        section.bottom_margin = Cm(1.0)
        section.top_margin    = Cm(1.5)
        section.left_margin   = Cm(2.0)
        section.right_margin  = Cm(2.0)

    buf = io.BytesIO()
    doc.save(buf)
    docx_bytes = buf.getvalue()

    base_filename = (
        f"Undangan Semhas {nama_1} dan {nama_2}" if nama_2_val
        else f"Undangan Semhas {nama_1}"
    )

    # --- Google services ---
    gs_client, drive = get_google_clients()

    # Upload DOCX to Word folder
    upload_bytes(
        drive, f"{base_filename}.docx",
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        docx_bytes, DRIVE_FOLDER_ID_WORD,
    )

    # Convert to PDF via Drive, upload PDF to PDF folder
    pdf_bytes = convert_docx_to_pdf_via_drive(
        drive, docx_bytes, DRIVE_FOLDER_ID_WORD
    )
    upload_bytes(
        drive, f"{base_filename}.pdf",
        "application/pdf", pdf_bytes, DRIVE_FOLDER_ID_PDF,
    )

    # ================= SHEET 1 =================
    url1 = "https://docs.google.com/spreadsheets/d/1ah-giclLV2mOlVmuCSawytTLfYUbR4vz85YLiy0CGaw/edit?gid=1220279114#gid=1220279114"
    sh1 = gs_client.open_by_url(url1).get_worksheet(0)
    if nama_2_val:
        rows1 = [
            [hari_tanggal, nama_1, nim_1, pembimbing_1, "", penguji_1],
            ["", "", "", pembimbing_2, "", penguji_2],
            [waktu_sheets, nama_2, nim_2, pembimbing_1, "", penguji_1],
            ["", "", "", pembimbing_2, "", penguji_2],
        ]
    else:
        rows1 = [
            [hari_tanggal, nama_1, nim_1, pembimbing_1, "", penguji_1],
            [waktu_sheets, "", "", pembimbing_2, "", penguji_2],
        ]
    col_d = sh1.col_values(4)
    last_row = max((i + 1 for i, v in enumerate(col_d) if v.strip()), default=1)
    next_row1 = last_row + 1
    sh1.update(values=rows1, range_name=f"A{next_row1}", value_input_option="USER_ENTERED")

    # ================= SHEET 2 =================
    url2 = "https://docs.google.com/spreadsheets/d/1fCKpGXIosV2tIOOOFjYkbQmwX9c7JeBH2VNe9Zen2h8/edit?gid=1529040550#gid=1529040550"
    sh2 = gs_client.open_by_url(url2).get_worksheet(0)
    all_data = sh2.get_all_values()
    nip_map = {}
    for row in all_data[2:]:
        if len(row) > 6 and row[5].strip() and row[6].strip():
            nip_map[normalize_name(row[5])] = row[6].replace(" ", "").strip()
        if len(row) > 8 and row[7].strip() and row[8].strip():
            nip_map[normalize_name(row[7])] = row[8].replace(" ", "").strip()

    def get_nip(raw):
        if not raw:
            return ""
        n = normalize_name(raw)
        if not n:
            return ""
        if n in nip_map:
            return nip_map[n]
        m = difflib.get_close_matches(n, nip_map.keys(), n=1, cutoff=0.7)
        return nip_map[m[0]] if m else ""

    nip_p1, nip_p2 = get_nip(pembimbing_1), get_nip(pembimbing_2)
    nip_u1, nip_u2 = get_nip(penguji_1), get_nip(penguji_2)

    if nama_2_val:
        rows2 = [
            ["", "", hari_tanggal, nama_1, nim_1, pembimbing_1, nip_p1, penguji_1, nip_u1, judul],
            ["", "", waktu_sheets, nama_2, nim_2, pembimbing_2, nip_p2, penguji_2, nip_u2, ""],
        ]
    else:
        rows2 = [
            ["", "", hari_tanggal, nama_1, nim_1, pembimbing_1, nip_p1, penguji_1, nip_u1, judul],
            ["", "", waktu_sheets, "", "", pembimbing_2, nip_p2, penguji_2, nip_u2, ""],
        ]
    col_f = sh2.col_values(6)
    last2 = max((i + 1 for i, v in enumerate(col_f) if v.strip()), default=1)
    next_row2 = last2 + 1
    sh2.update(values=rows2, range_name=f"A{next_row2}", value_input_option="USER_ENTERED")

    return {
        "docx_bytes": docx_bytes,
        "pdf_bytes": pdf_bytes,
        "base_filename": base_filename,
        "next_row1": next_row1,
        "next_row2": next_row2,
    }

# =====================================================================
# STREAMLIT UI
# =====================================================================
st.set_page_config(page_title="Undangan Semhas", page_icon="🎓", layout="centered")
st.title("🎓 Generator Undangan Seminar Hasil")
st.caption("Paste teks undangan → proses → file otomatis masuk ke Google Drive & Sheets.")

# Optional password gate
if "APP_PASSWORD" in st.secrets:
    pwd = st.text_input("Password", type="password")
    if pwd != st.secrets["APP_PASSWORD"]:
        st.info("Masukkan password untuk melanjutkan.")
        st.stop()

raw_text = st.text_area(
    "Paste informasi undangan di sini:",
    height=260,
    placeholder="Nomor: ...\nNama 1 (NIM): ...\nPembimbing 1: ...\nPenguji 1: ...",
)

if st.button("🚀 Proses", type="primary", use_container_width=True):
    if not raw_text.strip():
        st.warning("Teks masih kosong.")
    else:
        with st.spinner("Membuat dokumen, upload ke Drive, dan update Sheets..."):
            try:
                st.session_state.result = process(raw_text)
                st.success("Selesai! File sudah di Drive dan Sheets sudah diupdate.")
            except Exception as e:
                st.error(f"Terjadi kesalahan: {e}")

if "result" in st.session_state:
    r = st.session_state.result
    st.divider()
    st.subheader("📥 Download ke perangkat ini")
    c1, c2 = st.columns(2)
    with c1:
        st.download_button(
            "⬇️ Download PDF",
            data=r["pdf_bytes"],
            file_name=f"{r['base_filename']}.pdf",
            mime="application/pdf",
            use_container_width=True,
        )
    with c2:
        st.download_button(
            "⬇️ Download DOCX",
            data=r["docx_bytes"],
            file_name=f"{r['base_filename']}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
        )
    st.info(
        f"✅ DOCX & PDF sudah di-upload ke Google Drive.\n\n"
        f"📊 Nilai Semhas: baris **{r['next_row1']}**\n\n"
        f"📊 SK Final: baris **{r['next_row2']}**"
    )