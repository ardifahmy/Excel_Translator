# Excel Translator Pro — Streamlit

Aplikasi untuk menerjemahkan isi workbook Excel dari Bahasa Indonesia ke bahasa tujuan.

## Fitur
- Form pemilihan bahasa: proses **tidak berjalan** sampai tombol `Mulai Terjemahkan` ditekan.
- Mesin Google Cloud Translation API dengan daftar bahasa aktual dari API.
- Mode NLLB-200 lokal: pilihan bahasa diambil dari vocabulary tokenizer, termasuk Punjabi Shahmukhi (`pnb_Arab`) jika tersedia.
- Mendukung `.xlsx` dan `.xlsm`.
- Melewati angka, tanggal, sel kosong, formula, hyperlink, URL, dan email.
- Menyimpan file hasil baru; file sumber tidak ditimpa.
- Menghindari request duplikat untuk teks sel yang sama.

## Penting tentang deployment gratis
Streamlit Community Cloud gratis cocok untuk antarmuka, tetapi model NLLB-200 600M berukuran besar dan dapat melewati batas memori Cloud. Karena itu:
- Untuk deploy Streamlit Cloud, gunakan **Google Cloud Translation API** sebagai mesin remote.
- Google Cloud Translation API memerlukan project Google Cloud, API aktif, dan billing. Penggunaan mungkin dikenai biaya; periksa harga dan kuota sebelum digunakan.
- Streamlit Community Cloud gratis tidak berarti mesin penerjemah/API juga gratis.
- Mode NLLB lokal lebih cocok dijalankan di komputer sendiri atau Google Colab dengan GPU. Model NLLB-200 distilled 600M berlisensi CC-BY-NC 4.0 dan ditujukan untuk riset, bukan deployment produksi/komersial.

## Menjalankan lokal — mode Google Cloud
```bash
python -m venv .venv
# Windows:
.venv\Scripts\activate
# macOS/Linux:
source .venv/bin/activate

pip install -r requirements.txt
streamlit run app.py
```
Masukkan API key Google Cloud Translation pada sidebar.

Atau atur environment variable:
```bash
GOOGLE_TRANSLATE_API_KEY=isi_api_key
```

## Deploy gratis ke Streamlit Community Cloud
1. Buat repository GitHub baru.
2. Upload `app.py` dan `requirements.txt` (boleh sertakan README).
3. Buka https://share.streamlit.io/ dan login dengan GitHub.
4. Pilih repository, branch, dan `app.py`.
5. Deploy.
6. Pada dashboard aplikasi: **Settings → Secrets**, tambahkan:
```toml
GOOGLE_TRANSLATE_API_KEY = "API_KEY_KAMU"
```
7. Restart/reboot aplikasi.

Jangan pernah mengunggah API key ke repository publik.

## Menjalankan NLLB lokal (PC/Colab)
Pasang dependensi tambahan:
```bash
pip install -r requirements-local.txt
streamlit run app.py
```
Pilih `NLLB-200 (lokal)`. Pertama kali, tokenizer dan model diunduh dari Hugging Face. Model membutuhkan RAM/VRAM besar; GPU disarankan.

## Punjabi Shahmukhi
- NLLB menggunakan kode `pnb_Arab` untuk Punjabi Shahmukhi jika token tersebut ada dalam model.
- Jangan pilih `pan_Guru`, karena itu Punjabi Gurmukhi.
- Google Cloud dapat menampilkan Punjabi sebagai `pa`, tetapi kode itu tidak menjamin aksara Shahmukhi. Untuk persyaratan Shahmukhi, gunakan opsi NLLB `pnb_Arab` jika tersedia.

## Catatan Excel
OpenPyXL mempertahankan banyak style dan struktur workbook, tetapi fitur Excel tingkat lanjut (pivot cache, slicer, beberapa objek tertanam, dan elemen proprietary) perlu pengujian khusus. Buat salinan cadangan sebelum menerjemahkan file penting.
