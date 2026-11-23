import streamlit as st
from google import genai
from google.genai import types
from pydantic import BaseModel
from PIL import Image
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import io
from datetime import datetime

# Konfigurasi Halaman Web
st.set_page_config(page_title="AI NCR Generator - QC", page_icon="📋", layout="centered")
st.title("📋 AI NCR Generator (Tim QC)")
st.write("Ubah catatan & foto temuan lapangan menjadi draf laporan NCR standar format Excel.")

# Ambil API Key dari Secrets (saat online) atau input manual (saat tes lokal)
api_key = st.secrets.get("GEMINI_API_KEY", None)
if not api_key:
    api_key = st.sidebar.text_input("Masukkan Gemini API Key:", type="password")

# Struktur Output JSON agar rapi masuk ke Excel
class LaporanNCR(BaseModel):
    kategori_temuan: str
    deskripsi_formal: str
    potensi_root_cause: str
    rekomendasi_disposition: str

# Form Input untuk Tim QC di Lapangan
with st.form("form_ncr"):
    col1, col2 = st.columns(2)
    with col1:
        nama_inspector = st.text_input("Nama Inspector", placeholder="Contoh: Alif")
        nama_unit = st.text_input("Nama Unit / Proyek", placeholder="Contoh: Tangki CPO 30KL - Unit 4")
    with col2:
        tanggal = st.date_input("Tanggal Inspeksi", datetime.today())
        area_inspeksi = st.text_input("Area / Workshop", placeholder="Contoh: Workshop 2 / Line Welding")

    catatan_kasar = st.text_area(
        "Catatan Temuan Lapangan (Ketik santai atau pakai Voice-to-Text HP)",
        placeholder="Contoh: Lasan di baffle dalam ada undercut sekitar 3mm panjang 5cm, sama lebar mainframe selisih 4mm dari drawing."
    )
    
    foto_temuan = st.file_uploader("Upload / Jepret Foto Temuan (Opsional)", type=["jpg", "jpeg", "png"])
    
    submit = st.form_submit_button("✨ Generate Laporan NCR")

# Fungsi untuk membuat file Excel (.xlsx) rapi
def buat_excel(data_dict):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Laporan NCR"

    headers = [
        "No. NCR", "Tanggal", "Nama Unit / Proyek", "Area Inspeksi",
        "Inspector", "Kategori Temuan", "Deskripsi Ketidaksesuaian",
        "Potensi Akar Masalah", "Rekomendasi Tindakan", "Status"
    ]
    ws.append(headers)

    # Styling Header
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)
    thin_border = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin')
    )

    for col_num, cell in enumerate(ws[1], 1):
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border

    # Isi Data
    row_data = [
        data_dict["no_ncr"], data_dict["tanggal"], data_dict["nama_unit"],
        data_dict["area"], data_dict["inspector"], data_dict["kategori"],
        data_dict["deskripsi"], data_dict["root_cause"], data_dict["disposition"],
        "OPEN"
    ]
    ws.append(row_data)

    # Styling Baris Data & Lebar Kolom
    column_widths = [15, 14, 25, 20, 18, 20, 45, 35, 35, 12]
    for i, width in enumerate(column_widths, 1):
        col_letter = openpyxl.utils.get_column_letter(i)
        ws.column_dimensions[col_letter].width = width
        cell = ws.cell(row=2, column=i)
        cell.alignment = Alignment(vertical="top", wrap_text=True)
        cell.border = thin_border

    ws.row_dimensions[1].height = 28
    ws.row_dimensions[2].height = 85

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# Eksekusi AI saat tombol diklik
if submit:
    if not api_key:
        st.error("API Key Gemini belum dimasukkan! Isi di menu sebelah kiri (sidebar).")
    elif not catatan_kasar:
        st.warning("Isi catatan temuan lapangannya terlebih dahulu.")
    else:
        with st.spinner("Sedang menyusun draf laporan NCR..."):
            try:
                client = genai.Client(api_key=api_key)
                
                prompt = f"""
                Kamu adalah Quality Control Engineer senior di bidang manufaktur. 
                Ubah catatan temuan lapangan berikut menjadi laporan Non-Conformance Report (NCR) yang formal, objektif, dan teknis dalam bahasa Indonesia.
                Jika dilampirkan foto, analisa juga kondisi visual pada foto tersebut untuk memperjelas deskripsi cacat.

                Data Input:
                - Unit/Proyek: {nama_unit}
                - Area: {area_inspeksi}
                - Catatan Lapangan: {catatan_kasar}

                Instruksi pengisian:
                - kategori_temuan: Pilih kategori teknis (misal: Visual Welding, Dimensional, Material, Painting, Assembly).
                - deskripsi_formal: Jelaskan ketidaksesuaian dengan kalimat pasif yang baku dan jelas.
                - potensi_root_cause: Analisis singkat kemungkinan penyebab teknis di lapangan.
                - rekomendasi_disposition: Berikan saran tindakan (misal: Rework / Repair / Use As Is / Scrap) beserta langkah teknis singkatnya.
                """

                contents = [prompt]
                if foto_temuan:
                    img = Image.open(foto_temuan)
                    contents.append(img)

                response = client.models.generate_content(
                    model="gemini-2.5-flash",
                    contents=contents,
                    config=types.GenerateContentConfig(
                        response_mime_type="application/json",
                        response_schema=LaporanNCR,
                        temperature=0.2
                    )
                )

                hasil = LaporanNCR.model_validate_json(response.text)
                no_ncr_otomatis = f"NCR-{tanggal.strftime('%Y%m%d')}-01"

                st.success("Draf NCR berhasil dibuat!")
                st.subheader("Preview Hasil:")
                st.markdown(f"**No. NCR:** {no_ncr_otomatis}")
                st.markdown(f"**Kategori:** {hasil.kategori_temuan}")
                st.markdown(f"**Deskripsi NCR:**\n{hasil.deskripsi_formal}")
                st.markdown(f"**Potensi Akar Masalah:**\n{hasil.potensi_root_cause}")
                st.markdown(f"**Rekomendasi Tindakan:**\n{hasil.rekomendasi_disposition}")

                # Siapkan data untuk Excel
                data_excel = {
                    "no_ncr": no_ncr_otomatis,
                    "tanggal": tanggal.strftime("%d-%m-%Y"),
                    "nama_unit": nama_unit or "-",
                    "area": area_inspeksi or "-",
                    "inspector": nama_inspector or "-",
                    "kategori": hasil.kategori_temuan,
                    "deskripsi": hasil.deskripsi_formal,
                    "root_cause": hasil.potensi_root_cause,
                    "disposition": hasil.rekomendasi_disposition
                }

                file_xlsx = buat_excel(data_excel)
                st.download_button(
                    label="📥 Download File Excel (.xlsx)",
                    data=file_xlsx,
                    file_name=f"{no_ncr_otomatis}_{nama_unit.replace(' ', '_')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            except Exception as e:
                st.error(f"Terjadi kesalahan: {e}")
