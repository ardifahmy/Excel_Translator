import os
import re
import html
import time
from io import BytesIO
from typing import Dict, List, Tuple

import requests
import streamlit as st
from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell

APP_TITLE = "Excel Translator Pro"
GOOGLE_LANG_URL = "https://translation.googleapis.com/language/translate/v2/languages"
GOOGLE_TRANSLATE_URL = "https://translation.googleapis.com/language/translate/v2"
NLLB_MODEL = "facebook/nllb-200-distilled-600M"

# Nama ramah untuk bahasa yang secara khusus diminta pengguna.
# Punjabi menggunakan kode NLLB pnb_Arab (Shahmukhi), bukan pan_Guru.
FRIENDLY_NAMES = {
    "pnb_Arab": "Punjabi — Shahmukhi",
    "pan_Guru": "Punjabi — Gurmukhi",
    "tsn_Latn": "Tswana",
    "lus_Latn": "Mizo",
    "ktu_Latn": "Kituba",
    "wol_Latn": "Wolof",
    "fuv_Latn": "Fulfulde (Nigeria)",
    "ful_Latn": "Fulani / Fulfulde (kode alternatif; cek ketersediaan)",
    "fon_Latn": "Fon",
    "gaa_Latn": "Ga",
    "dyu_Latn": "Dyula",
    "sus_Latn": "Susu",
    "tiv_Latn": "Tiv",
    "nah_Latn": "Nahuatl",
    "zai_Latn": "Zapotec (kode umum; varietas perlu dipastikan)",
    "yua_Latn": "Maya Yucatec",
    "ind_Latn": "Indonesia",
    "eng_Latn": "Inggris",
    "fra_Latn": "Prancis",
    "spa_Latn": "Spanyol",
    "arb_Arab": "Arab Standar Modern",
}

# Fallback daftar pilihan jika Google API key belum dimasukkan.
# Opsi ini hanya ditampilkan sebagai informasi/pilihan awal; ketersediaan aktual
# diverifikasi kembali melalui Google Cloud API sebelum menerjemahkan.
GOOGLE_FALLBACK = {
    "af": "Afrikaans", "sq": "Albania", "am": "Amharik", "ar": "Arab",
    "hy": "Armenia", "az": "Azerbaijan", "eu": "Basque", "bn": "Bengali",
    "bs": "Bosnia", "bg": "Bulgaria", "ca": "Catalan", "zh-CN": "Mandarin (Sederhana)",
    "zh-TW": "Mandarin (Tradisional)", "hr": "Kroasia", "cs": "Ceko",
    "da": "Denmark", "nl": "Belanda", "en": "Inggris", "et": "Estonia",
    "fi": "Finlandia", "fr": "Prancis", "gl": "Galicia", "ka": "Georgia",
    "de": "Jerman", "el": "Yunani", "gu": "Gujarati", "ht": "Kreol Haiti",
    "ha": "Hausa", "he": "Ibrani", "hi": "Hindi", "hu": "Hungaria",
    "is": "Islandia", "ig": "Igbo", "id": "Indonesia", "ga": "Irlandia",
    "it": "Italia", "ja": "Jepang", "jv": "Jawa", "kn": "Kannada",
    "kk": "Kazakh", "km": "Khmer", "rw": "Kinyarwanda", "ko": "Korea",
    "lo": "Lao", "lv": "Latvia", "lt": "Lithuania", "mk": "Makedonia",
    "ms": "Melayu", "ml": "Malayalam", "mr": "Marathi", "mn": "Mongolia",
    "my": "Myanmar", "ne": "Nepal", "no": "Norwegia", "or": "Odia",
    "ps": "Pashto", "fa": "Persia", "pl": "Polandia", "pt": "Portugis",
    "pa": "Punjabi", "ro": "Rumania", "ru": "Rusia", "sr": "Serbia",
    "si": "Sinhala", "sk": "Slovakia", "sl": "Slovenia", "so": "Somali",
    "es": "Spanyol", "sw": "Swahili", "sv": "Swedia", "tl": "Tagalog",
    "ta": "Tamil", "te": "Telugu", "th": "Thai", "tr": "Turki",
    "uk": "Ukraina", "ur": "Urdu", "uz": "Uzbek", "vi": "Vietnam",
    "cy": "Welsh", "yo": "Yoruba", "zu": "Zulu",
}

def get_secret_or_env(name: str) -> str:
    try:
        value = st.secrets.get(name, "")
        if value:
            return str(value).strip()
    except Exception:
        pass
    return os.getenv(name, "").strip()

@st.cache_data(ttl=3600, show_spinner=False)
def google_supported_languages(api_key: str) -> Dict[str, str]:
    response = requests.get(
        GOOGLE_LANG_URL,
        params={"key": api_key, "target": "id"},
        timeout=30,
    )
    response.raise_for_status()
    payload = response.json()
    return {
        item["language"]: item.get("name", item["language"])
        for item in payload.get("data", {}).get("languages", [])
    }

def get_nllb_tokenizer():
    # Tokenizer relatif kecil; model 600M baru dimuat setelah tombol proses ditekan.
    from transformers import AutoTokenizer
    return AutoTokenizer.from_pretrained(NLLB_MODEL)

def get_nllb_language_options(tokenizer) -> Dict[str, str]:
    vocab = tokenizer.get_vocab()
    # NLLB language tokens have form xxx_Latn, xxx_Arab, xxx_Cyrl, etc.
    pattern = re.compile(r"^[a-z]{3}_[A-Z][a-z]{3}$")
    codes = sorted(token for token in vocab if pattern.match(token))
    options = {}
    for code in codes:
        if code == "ind_Latn":
            continue
        options[code] = FRIENDLY_NAMES.get(code, f"{code} (NLLB-200)")
    # Ensure Shahmukhi is explicit if tokenizer actually contains the token.
    if "pnb_Arab" in vocab:
        options["pnb_Arab"] = FRIENDLY_NAMES["pnb_Arab"]
    return dict(sorted(options.items(), key=lambda item: item[1].casefold()))

def is_translatable_cell(cell) -> bool:
    if isinstance(cell, MergedCell):
        return False
    value = cell.value
    if not isinstance(value, str):
        return False
    if not value.strip() or value.startswith("="):
        return False
    if cell.hyperlink:
        return False
    # Skip plain URL/email cells.
    if re.match(r"^(https?://|www\.)", value.strip(), re.I):
        return False
    if re.match(r"^[^@\s]+@[^@\s]+\.[^@\s]+$", value.strip()):
        return False
    return True

def google_translate_batch(texts: List[str], target_code: str, api_key: str) -> List[str]:
    # Google Cloud Translation Basic v2 accepts q as a repeated/list parameter.
    # Keep each request modest to avoid request-size errors.
    response = requests.post(
        GOOGLE_TRANSLATE_URL,
        params={"key": api_key},
        data=[("q", text) for text in texts] + [
            ("source", "id"),
            ("target", target_code),
            ("format", "text"),
        ],
        timeout=90,
    )
    if response.status_code >= 400:
        try:
            detail = response.json().get("error", {}).get("message", response.text)
        except Exception:
            detail = response.text
        raise RuntimeError(f"Google Cloud Translation API: {detail}")
    result = response.json().get("data", {}).get("translations", [])
    if len(result) != len(texts):
        raise RuntimeError("Jumlah hasil terjemahan tidak sama dengan jumlah teks yang dikirim.")
    return [html.unescape(item.get("translatedText", "")) for item in result]

@st.cache_resource(show_spinner="Memuat model NLLB-200. Proses pertama dapat memerlukan waktu...")
def load_nllb_model():
    import torch
    from transformers import AutoModelForSeq2SeqLM
    device = "cuda" if torch.cuda.is_available() else "cpu"
    model = AutoModelForSeq2SeqLM.from_pretrained(NLLB_MODEL)
    model.eval()
    model.to(device)
    return model, device

def nllb_translate_batch(texts: List[str], target_code: str, tokenizer, model, device) -> List[str]:
    import torch
    target_id = tokenizer.convert_tokens_to_ids(target_code)
    unk_id = tokenizer.unk_token_id
    if target_id is None or target_id == unk_id:
        raise ValueError(f"Kode bahasa NLLB tidak valid: {target_code}")
    tokenizer.src_lang = "ind_Latn"
    results = []
    for text in texts:
        encoded = tokenizer(
            text,
            return_tensors="pt",
            truncation=True,
            max_length=512,
        ).to(device)
        with torch.no_grad():
            generated = model.generate(
                **encoded,
                forced_bos_token_id=target_id,
                max_new_tokens=256,
                num_beams=4,
            )
        results.append(tokenizer.batch_decode(generated, skip_special_tokens=True)[0])
    return results

def process_workbook(uploaded_file, engine: str, target_code: str, api_key: str = "", tokenizer=None):
    raw = uploaded_file.getvalue()
    keep_vba = uploaded_file.name.lower().endswith(".xlsm")
    workbook = load_workbook(BytesIO(raw), keep_vba=keep_vba, data_only=False)

    tasks: List[Tuple[str, str, str]] = []
    for ws in workbook.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                if is_translatable_cell(cell):
                    tasks.append((ws.title, cell.coordinate, cell.value))

    if not tasks:
        raise ValueError("Tidak ditemukan sel teks yang perlu diterjemahkan.")

    # Cache per teks untuk menghindari biaya/permintaan berulang.
    unique_texts = list(dict.fromkeys(item[2] for item in tasks))
    translated_map: Dict[str, str] = {}

    progress = st.progress(0, text="Menyiapkan penerjemahan...")
    status = st.empty()

    if engine == "Google Cloud Translation API":
        # Google API menerima beberapa q dalam satu request; batasi ukuran total.
        batches, current, chars = [], [], 0
        for text in unique_texts:
            if current and (len(current) >= 50 or chars + len(text) > 18000):
                batches.append(current)
                current, chars = [], 0
            current.append(text)
            chars += len(text)
        if current:
            batches.append(current)

        for index, batch in enumerate(batches, start=1):
            status.write(f"Menerjemahkan kelompok {index}/{len(batches)}...")
            translated = google_translate_batch(batch, target_code, api_key)
            translated_map.update(dict(zip(batch, translated)))
            progress.progress(index / len(batches), text=f"Kelompok {index}/{len(batches)} selesai")

    else:
        model, device = load_nllb_model()
        batches = [unique_texts[i:i + 8] for i in range(0, len(unique_texts), 8)]
        for index, batch in enumerate(batches, start=1):
            status.write(f"NLLB-200 ({device}): kelompok {index}/{len(batches)}...")
            translated = nllb_translate_batch(batch, target_code, tokenizer, model, device)
            translated_map.update(dict(zip(batch, translated)))
            progress.progress(index / len(batches), text=f"Kelompok {index}/{len(batches)} selesai")

    for sheet_name, coordinate, original in tasks:
        workbook[sheet_name][coordinate] = translated_map.get(original, original)

    output = BytesIO()
    workbook.save(output)
    output.seek(0)
    return output.getvalue(), len(tasks), len(unique_texts), len(workbook.worksheets)

st.set_page_config(page_title=APP_TITLE, page_icon="🌐", layout="wide")
st.title("🌐 Excel Translator Pro")
st.caption("Terjemahkan isi workbook Excel dari Bahasa Indonesia ke berbagai bahasa tujuan.")

with st.sidebar:
    st.header("Pengaturan mesin")
    engine = st.radio(
        "Mesin penerjemah",
        ["Google Cloud Translation API", "NLLB-200 (lokal)"],
        help="Google Cloud cocok untuk deployment Streamlit Cloud. NLLB lokal membutuhkan RAM besar dan lebih cocok dijalankan di PC/Colab dengan GPU.",
    )
    st.divider()
    st.markdown("**Bahasa sumber**")
    st.info("Indonesia (otomatis: `id` untuk Google, `ind_Latn` untuk NLLB)")
    st.markdown("**Punjabi**")
    st.caption("Untuk Punjabi Shahmukhi, pilih `Punjabi — Shahmukhi` pada mesin NLLB jika token tersebut tersedia. Google Cloud dapat memakai kode `pa`, tetapi tidak menjamin aksara Shahmukhi.")

api_key = get_secret_or_env("GOOGLE_TRANSLATE_API_KEY")
tokenizer_for_list = None
if engine == "Google Cloud Translation API":
    if not api_key:
        api_key = st.text_input(
            "Google Cloud Translation API key",
            type="password",
            help="Simpan permanen melalui Streamlit Secrets dengan nama GOOGLE_TRANSLATE_API_KEY.",
        )
    if api_key:
        try:
            google_languages = google_supported_languages(api_key)
            target_options = {
                code: name for code, name in google_languages.items()
                if code != "id"
            }
            st.success(f"Terhubung. Google Cloud mengembalikan {len(target_options)} bahasa tujuan.")
        except Exception as exc:
            target_options = {}
            st.error(f"Tidak dapat mengambil daftar bahasa Google: {exc}")
    else:
        target_options = {}
        st.warning("Masukkan API key Google Cloud Translation untuk memuat daftar bahasa aktual dan menerjemahkan.")
else:
    target_options = {}
    tokenizer_for_list = None
    try:
        with st.spinner("Memuat daftar bahasa NLLB (tokenizer saja, tanpa model besar)..."):
            tokenizer_for_list = get_nllb_tokenizer()
            target_options = get_nllb_language_options(tokenizer_for_list)
        st.success(f"Tokenizer siap. {len(target_options)} kode bahasa tujuan terdeteksi.")
    except Exception as exc:
        st.error(f"Gagal memuat tokenizer NLLB: {exc}")
        st.info("Untuk mode NLLB lokal, pasang requirements-local.txt. Di Streamlit Cloud gratis, model 600M berpotensi melewati batas RAM.")

st.markdown("### 1. Pilih file dan bahasa tujuan")
with st.form("translation_form", clear_on_submit=False):
    uploaded = st.file_uploader(
        "Upload dokumen Excel",
        type=["xlsx", "xlsm"],
        help="Format .xlsx dan .xlsm. File asli tidak diubah.",
    )
    if target_options:
        option_codes = list(target_options.keys())
        labels = [f"{target_options[code]}  ({code})" for code in option_codes]
        selected_label = st.selectbox(
            "Bahasa tujuan",
            options=labels,
            index=0,
            help="Pilih bahasa tujuan terlebih dahulu. Penerjemahan hanya berjalan setelah tombol ditekan.",
        )
        selected_code = option_codes[labels.index(selected_label)]
    else:
        selected_code = None
        st.selectbox("Bahasa tujuan", options=["Masukkan API key / perbaiki mesin terlebih dahulu"], disabled=True)
    submitted = st.form_submit_button("🚀 Mulai Terjemahkan", type="primary", use_container_width=True)

if submitted:
    if uploaded is None:
        st.error("Upload file Excel terlebih dahulu.")
    elif not selected_code:
        st.error("Daftar bahasa tujuan belum tersedia. Periksa mesin penerjemah.")
    elif engine == "Google Cloud Translation API" and not api_key:
        st.error("API key Google Cloud belum diisi.")
    else:
        try:
            st.markdown("### 2. Proses penerjemahan")
            with st.spinner("Memproses workbook. Jangan tutup halaman..."):
                result_bytes, cell_count, unique_count, sheet_count = process_workbook(
                    uploaded,
                    engine,
                    selected_code,
                    api_key=api_key,
                    tokenizer=tokenizer_for_list,
                )
            st.success("Penerjemahan selesai.")
            c1, c2, c3 = st.columns(3)
            c1.metric("Sel teks diproses", cell_count)
            c2.metric("Teks unik diterjemahkan", unique_count)
            c3.metric("Worksheet", sheet_count)
            original_name = uploaded.name.rsplit(".", 1)[0]
            extension = uploaded.name.rsplit(".", 1)[-1]
            safe_code = re.sub(r"[^A-Za-z0-9_-]", "_", selected_code)
            output_name = f"{original_name}_translated_{safe_code}.{extension}"
            st.download_button(
                "⬇️ Download Excel hasil terjemahan",
                data=result_bytes,
                file_name=output_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
            st.caption("Rumus, angka, tanggal, sel kosong, dan hyperlink dilewati. Format sel yang didukung OpenPyXL dipertahankan.")
        except Exception as exc:
            st.error(f"Proses gagal: {exc}")
            st.exception(exc)

st.divider()
with st.expander("Catatan dukungan bahasa dan privasi"):
    st.markdown("""
- Daftar Google Cloud berasal dari endpoint resmi API dan hanya memuat bahasa yang tersedia untuk project/API key tersebut.
- Daftar NLLB dibaca langsung dari vocabulary tokenizer model, bukan daftar buatan. Bahasa yang terdeteksi belum tentu berkualitas sama untuk pasangan Indonesia → bahasa tujuan.
- Punjabi Shahmukhi memakai kode NLLB `pnb_Arab` jika benar-benar tersedia pada tokenizer.
- Maya dan Zapotec mempunyai banyak varietas. Pilih varietas spesifik jika tersedia; nama generik tidak menjamin kecocokan dialek.
- NLLB-200 adalah model riset dengan lisensi CC-BY-NC 4.0 dan bukan model produksi/terjemahan dokumen resmi. Periksa lisensi sebelum penggunaan komersial.
- File Excel dikirim ke layanan Google Cloud jika mesin Google dipilih. Jika menggunakan NLLB lokal, pemrosesan dilakukan pada mesin tempat aplikasi berjalan.
""")
