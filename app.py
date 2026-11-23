import io
import re
import unicodedata
from collections import defaultdict
from datetime import date

import openpyxl
import pandas as pd
import streamlit as st

st.set_page_config(page_title='Auto Input SO Harian', page_icon='📊', layout='wide')
st.title('📊 Auto Input Penjualan Promix → SO Harian')
st.caption('Khusus format Promix Table 4 dan template SO Harian TGRHAR yang diberikan.')

def normal(text):
    text = unicodedata.normalize('NFKC', str(text or '')).upper().strip()
    return re.sub(r'\s+', ' ', text)

def extract_sales(upload):
    wb = openpyxl.load_workbook(io.BytesIO(upload.getvalue()), read_only=True, data_only=True)
    if 'Table 4' not in wb.sheetnames:
        raise ValueError('Sheet Table 4 tidak ditemukan pada Promix.')
    sheet = wb['Table 4']
    sales = defaultdict(int)
    context = None
    ignored = []
    for row in sheet.iter_rows(min_row=2, values_only=True):
        desc = normal(row[0] if len(row)>0 else None)
        qty = row[1] if len(row)>1 else None
        if not desc:
            continue
        if desc.startswith('TOTAL -'):
            continue
        if qty is None:
            # Header seperti BEVERAGES DINE IN, MIE TAKE AWAY, dsb.
            if desc.endswith('DINE IN'):
                context = 'DINE IN'
            elif desc.endswith('TAKE AWAY') or desc.endswith(' TA'):
                context = 'TAKE AWAY'
            else:
                # Jangan meneruskan kategori sebelumnya ke kategori baru.
                context = None
            continue
        if isinstance(qty, (int, float)) and not isinstance(qty, bool) and context:
            if float(qty).is_integer() and qty >= 0:
                sales[(context, desc)] += int(qty)
        elif isinstance(qty, (int, float)):
            ignored.append((desc, qty))
    wb.close()
    return sales, ignored

def prepare_so(upload, sales, day, overwrite):
    wb = openpyxl.load_workbook(io.BytesIO(upload.getvalue()))
    report = []
    unmatched = []
    for category, sheet_name in [('DINE IN','INPUT PENJUALAN DINE IN'), ('TAKE AWAY','INPUT PENJUALAN TAKE AWAY')]:
        if sheet_name not in wb.sheetnames:
            raise ValueError(f'Sheet {sheet_name} tidak ditemukan.')
        ws = wb[sheet_name]
        columns = [cell.column for cell in ws[1] if cell.value == day]
        if len(columns) != 1:
            raise ValueError(f'Kolom tanggal {day} tidak ditemukan secara unik di {sheet_name}.')
        col = columns[0]
        menu_rows = defaultdict(list)
        for row in range(2, ws.max_row + 1):
            menu = normal(ws.cell(row,1).value)
            if menu and not (menu.startswith('|') or menu.startswith('TOTAL')):
                menu_rows[menu].append(row)
        for (source_category, menu), qty in sales.items():
            if source_category != category:
                continue
            matches = menu_rows.get(menu, [])
            if len(matches) != 1:
                unmatched.append({'Kategori':category,'Menu Promix':menu,'Qty':qty,'Alasan':'Tidak ada menu yang sama' if not matches else 'Nama menu duplikat di SO'})
                continue
            cell = ws.cell(matches[0], col)
            if cell.data_type == 'f':
                unmatched.append({'Kategori':category,'Menu Promix':menu,'Qty':qty,'Alasan':'Sel tujuan berisi formula'} )
                continue
            if cell.value not in (None, 0, '') and not overwrite:
                unmatched.append({'Kategori':category,'Menu Promix':menu,'Qty':qty,'Alasan':'Kolom tanggal sudah terisi (aktifkan izin timpa untuk mengganti)'})
                continue
            old = cell.value
            cell.value = qty
            report.append({'Kategori':category,'Menu':menu,'Qty':qty,'Sel':cell.coordinate,'Sebelumnya':old})
    output = io.BytesIO()
    wb.save(output)
    wb.close()
    output.seek(0)
    return output, report, unmatched

promix = st.file_uploader('1. Upload file Promix (.xlsx)', type=['xlsx'], key='promix')
so = st.file_uploader('2. Upload file SO Harian (.xlsx)', type=['xlsx'], key='so')
day = st.number_input('3. Tanggal penjualan (angka hari dalam bulan)', min_value=1, max_value=31, value=21, step=1)
overwrite = st.checkbox('Izinkan mengganti nilai tanggal yang sudah terisi', value=False)
st.warning('Pastikan bulan dan tahun file SO sesuai tanggal Promix. Aplikasi hanya memilih kolom berdasarkan angka tanggal.')
if st.button('Proses Otomatis', type='primary', disabled=not (promix and so)):
    try:
        sales, ignored = extract_sales(promix)
        output, report, unmatched = prepare_so(so, sales, int(day), overwrite)
        st.session_state['result'] = (output.getvalue(), report, unmatched, ignored)
    except Exception as exc:
        st.error(f'Gagal memproses: {exc}')
        st.session_state.pop('result', None)
if 'result' in st.session_state:
    data, report, unmatched, ignored = st.session_state['result']
    st.success(f'{len(report)} baris menu berhasil diisi. {len(unmatched)} baris perlu ditinjau.')
    st.subheader('Data berhasil diinput')
    st.dataframe(pd.DataFrame(report), use_container_width=True)
    st.subheader('Perlu diperiksa / belum diinput')
    st.dataframe(pd.DataFrame(unmatched), use_container_width=True)
    if ignored:
        st.info(f'{len(ignored)} baris ber-Qty di Promix tidak memiliki kategori Dine In / Take Away yang jelas dan tidak dimasukkan.')
    st.download_button('⬇️ Download SO Harian Hasil', data=data, file_name=f'SO_Harian_Hasil_Tgl_{day}.xlsx', mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    st.caption('Periksa kembali hasil sebelum digunakan untuk closing. Rumus dan format Excel dipertahankan sejauh didukung openpyxl; hasil rumus mungkin baru dihitung saat dibuka di Excel.')
