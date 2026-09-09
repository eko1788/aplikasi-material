import streamlit as st

st.title("This is a title")
st.title("_Streamlit_ is :blue[cool] :sunglasses:")
st.title("Dashboard", icon=":material/dashboard:")

```python
import streamlit as st
import sqlite3
import pandas as pd
import io
import os

from datetime import datetime, date
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter


# =========================================================
# KONFIGURASI
# =========================================================

st.set_page_config(
    page_title="Gudang Material",
    page_icon="🏗️",
    layout="wide"
)

DB_NAME = "gudang_material.db"
UPLOAD_DIR = "uploads"

os.makedirs(UPLOAD_DIR, exist_ok=True)


# =========================================================
# DATABASE
# =========================================================

def koneksi():
    return sqlite3.connect(DB_NAME)


def init_database():

    conn = koneksi()
    cursor = conn.cursor()

    # MASTER MATERIAL
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS material (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            kode_material TEXT UNIQUE NOT NULL,
            nama_material TEXT NOT NULL,
            kategori TEXT,
            spesifikasi TEXT,
            satuan TEXT,
            stok_minimum REAL DEFAULT 0,
            lokasi TEXT,
            created_at TEXT
        )
    """)

    # TRANSAKSI
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS transaksi (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nomor_dokumen TEXT,
            tanggal TEXT,
            jenis TEXT,
            kode_material TEXT,
            nama_material TEXT,
            spesifikasi TEXT,
            jumlah REAL,
            satuan TEXT,
            supplier_pemakai TEXT,
            proyek TEXT,
            keterangan TEXT,
            foto TEXT,
            created_at TEXT
        )
    """)

    conn.commit()
    conn.close()


init_database()


# =========================================================
# MASTER MATERIAL
# =========================================================

def ambil_material():

    conn = koneksi()

    df = pd.read_sql_query(
        """
        SELECT *
        FROM material
        ORDER BY nama_material
        """,
        conn
    )

    conn.close()

    return df


def tambah_material(
    kode,
    nama,
    kategori,
    spesifikasi,
    satuan,
    stok_minimum,
    lokasi
):

    conn = koneksi()

    try:

        conn.execute(
            """
            INSERT INTO material (
                kode_material,
                nama_material,
                kategori,
                spesifikasi,
                satuan,
                stok_minimum,
                lokasi,
                created_at
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                kode,
                nama,
                kategori,
                spesifikasi,
                satuan,
                stok_minimum,
                lokasi,
                datetime.now().strftime(
                    "%Y-%m-%d %H:%M:%S"
                )
            )
        )

        conn.commit()

        return True, "Material berhasil ditambahkan."

    except sqlite3.IntegrityError:

        return False, "Kode material sudah digunakan."

    finally:

        conn.close()


# =========================================================
# TRANSAKSI
# =========================================================

def ambil_transaksi():

    conn = koneksi()

    df = pd.read_sql_query(
        """
        SELECT *
        FROM transaksi
        ORDER BY tanggal DESC, id DESC
        """,
        conn
    )

    conn.close()

    return df


def simpan_transaksi(
    nomor_dokumen,
    tanggal,
    jenis,
    kode_material,
    nama_material,
    spesifikasi,
    jumlah,
    satuan,
    supplier_pemakai,
    proyek,
    keterangan,
    foto
):

    conn = koneksi()

    conn.execute(
        """
        INSERT INTO transaksi (
            nomor_dokumen,
            tanggal,
            jenis,
            kode_material,
            nama_material,
            spesifikasi,
            jumlah,
            satuan,
            supplier_pemakai,
            proyek,
            keterangan,
            foto,
            created_at
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        (
            nomor_dokumen,
            tanggal,
            jenis,
            kode_material,
            nama_material,
            spesifikasi,
            jumlah,
            satuan,
            supplier_pemakai,
            proyek,
            keterangan,
            foto,
            datetime.now().strftime(
                "%Y-%m-%d %H:%M:%S"
            )
        )
    )

    conn.commit()
    conn.close()


# =========================================================
# PERHITUNGAN STOK
# =========================================================

def hitung_stok():

    material = ambil_material()
    transaksi = ambil_transaksi()

    if material.empty:

        return pd.DataFrame()

    material["stok_masuk"] = 0.0
    material["stok_keluar"] = 0.0

    if not transaksi.empty:

        masuk = (
            transaksi[
                transaksi["jenis"] == "MATERIAL MASUK"
            ]
            .groupby("kode_material")["jumlah"]
            .sum()
        )

        keluar = (
            transaksi[
                transaksi["jenis"] == "MATERIAL KELUAR"
            ]
            .groupby("kode_material")["jumlah"]
            .sum()
        )

        material["stok_masuk"] = (
            material["kode_material"]
            .map(masuk)
            .fillna(0)
        )

        material["stok_keluar"] = (
            material["kode_material"]
            .map(keluar)
            .fillna(0)
        )

    material["stok_akhir"] = (
        material["stok_masuk"]
        - material["stok_keluar"]
    )

    material["status"] = material.apply(
        lambda x:
        "⚠️ STOK MINIMUM"
        if x["stok_akhir"] <= x["stok_minimum"]
        else "✅ NORMAL",
        axis=1
    )

    return material


# =========================================================
# EXPORT EXCEL
# =========================================================

def export_excel():

    stok = hitung_stok()
    material = ambil_material()
    transaksi = ambil_transaksi()

    output = io.BytesIO()

    workbook = Workbook()

    # =====================================================
    # SHEET STOK
    # =====================================================

    ws = workbook.active
    ws.title = "Stok Material"

    header_stok = [
        "Kode Material",
        "Nama Material",
        "Kategori",
        "Spesifikasi",
        "Satuan",
        "Lokasi",
        "Stok Minimum",
        "Material Masuk",
        "Material Keluar",
        "Stok Akhir",
        "Status"
    ]

    ws.append(header_stok)

    for cell in ws[1]:

        cell.font = Font(
            bold=True,
            color="FFFFFF"
        )

        cell.fill = PatternFill(
            fill_type="solid",
            fgColor="1F4E78"
        )

        cell.alignment = Alignment(
            horizontal="center",
            vertical="center"
        )

    for _, row in stok.iterrows():

        ws.append([
            row["kode_material"],
            row["nama_material"],
            row["kategori"],
            row["spesifikasi"],
            row["satuan"],
            row["lokasi"],
            row["stok_minimum"],
            row["stok_masuk"],
            row["stok_keluar"],
            row["stok_akhir"],
            row["status"]
        ])

    # =====================================================
    # SHEET TRANSAKSI
    # =====================================================

    ws2 = workbook.create_sheet(
        "Transaksi"
    )

    header_transaksi = [
        "ID",
        "Nomor Dokumen",
        "Tanggal",
        "Jenis",
        "Kode Material",
        "Nama Material",
        "Spesifikasi",
        "Jumlah",
        "Satuan",
        "Supplier / Pemakai",
        "Proyek",
        "Keterangan",
        "Foto"
    ]

    ws2.append(header_transaksi)

    for cell in ws2[1]:

        cell.font = Font(
            bold=True,
            color="FFFFFF"
        )

        cell.fill = PatternFill(
            fill_type="solid",
            fgColor="70AD47"
        )

    for _, row in transaksi.iterrows():

        ws2.append([
            row["id"],
            row["nomor_dokumen"],
            row["tanggal"],
            row["jenis"],
            row["kode_material"],
            row["nama_material"],
            row["spesifikasi"],
            row["jumlah"],
            row["satuan"],
            row["supplier_pemakai"],
            row["proyek"],
            row["keterangan"],
            row["foto"]
        ])

    # =====================================================
    # SHEET MASTER
    # =====================================================

    ws3 = workbook.create_sheet(
        "Master Material"
    )

    header_master = [
        "ID",
        "Kode Material",
        "Nama Material",
        "Kategori",
        "Spesifikasi",
        "Satuan",
        "Stok Minimum",
        "Lokasi",
        "Created At"
    ]

    ws3.append(header_master)

    for cell in ws3[1]:

        cell.font = Font(
            bold=True,
            color="FFFFFF"
        )

        cell.fill = PatternFill(
            fill_type="solid",
            fgColor="ED7D31"
        )

    for _, row in material.iterrows():

        ws3.append([
            row["id"],
            row["kode_material"],
            row["nama_material"],
            row["kategori"],
            row["spesifikasi"],
            row["satuan"],
            row["stok_minimum"],
            row["lokasi"],
            row["created_at"]
        ])

    # =====================================================
    # FORMAT EXCEL
    # =====================================================

    for sheet in workbook.worksheets:

        sheet.freeze_panes = "A2"

        for column in sheet.columns:

            max_length = 0

            column_letter = get_column_letter(
                column[0].column
            )

            for cell in column:

                if cell.value is not None:

                    max_length = max(
                        max_length,
                        len(str(cell.value))
                    )

            sheet.column_dimensions[
                column_letter
            ].width = min(
                max_length + 3,
                45
            )

    workbook.save(output)

    output.seek(0)

    return output


# =========================================================
# SIDEBAR
# =========================================================

st.sidebar.title("🏗️ GUDANG MATERIAL")

st.sidebar.caption(
    "Sistem Monitoring Stok Material"
)

menu = st.sidebar.radio(
    "MENU",
    [
        "📊 Dashboard",
        "🧱 Master Material",
        "📥 Material Masuk",
        "📤 Material Keluar",
        "📋 Riwayat Transaksi",
        "📊 Laporan Excel"
    ]
)


# =========================================================
# DASHBOARD
# =========================================================

if menu == "📊 Dashboard":

    st.title(
        "📊 Dashboard Gudang Material"
    )

    stok = hitung_stok()

    transaksi = ambil_transaksi()

    if stok.empty:

        st.info(
            "Belum ada material. "
            "Silakan input Master Material terlebih dahulu."
        )

    else:

        total_material = len(stok)

        total_stok = stok[
            "stok_akhir"
        ].sum()

        total_masuk = stok[
            "stok_masuk"
        ].sum()

        total_keluar = stok[
            "stok_keluar"
        ].sum()

        stok_minimum = len(
            stok[
                stok["stok_akhir"]
                <= stok["stok_minimum"]
            ]
        )

        col1, col2, col3, col4, col5 = st.columns(5)

        col1.metric(
            "Jenis Material",
            total_material
        )

        col2.metric(
            "Total Stok",
            f"{total_stok:,.2f}"
        )

        col3.metric(
            "Material Masuk",
            f"{total_masuk:,.2f}"
        )

        col4.metric(
            "Material Keluar",
            f"{total_keluar:,.2f}"
        )

        col5.metric(
            "Stok Minimum",
            stok_minimum
        )

        st.divider()

        # =================================================
        # GRAFIK
        # =================================================

        st.subheader(
            "📈 Grafik Stok Material"
        )

        chart_data = stok[
            [
                "nama_material",
                "stok_akhir"
            ]
        ].copy()

        chart_data = chart_data.set_index(
            "nama_material"
        )

        st.bar_chart(
            chart_data
        )

        # =================================================
        # STOK MINIMUM
        # =================================================

        st.subheader(
            "⚠️ Material dengan Stok Minimum"
        )

        minimum = stok[
            stok["stok_akhir"]
            <= stok["stok_minimum"]
        ]

        if minimum.empty:

            st.success(
                "Semua stok material masih aman."
            )

        else:

            st.dataframe(
                minimum[
                    [
                        "kode_material",
                        "nama_material",
                        "spesifikasi",
                        "satuan",
                        "stok_akhir",
                        "stok_minimum",
                        "lokasi"
                    ]
                ],
                use_container_width=True,
                hide_index=True
            )

        # =================================================
        # TABEL STOK
        # =================================================

        st.subheader(
            "📦 Posisi Stok Material"
        )

        st.dataframe(
            stok[
                [
                    "kode_material",
                    "nama_material",
                    "kategori",
                    "spesifikasi",
                    "satuan",
                    "lokasi",
                    "stok_masuk",
                    "stok_keluar",
                    "stok_akhir",
                    "status"
                ]
            ],
            use_container_width=True,
            hide_index=True
        )


# =========================================================
# MASTER MATERIAL
# =========================================================

elif menu == "🧱 Master Material":

    st.title(
        "🧱 Master Material"
    )

    with st.form(
        "form_master",
        clear_on_submit=True
    ):

        col1, col2 = st.columns(2)

        with col1:

            kode = st.text_input(
                "Kode Material *",
                placeholder="MAT-001"
            )

            nama = st.text_input(
                "Nama Material *",
                placeholder="Semen"
            )

            kategori = st.text_input(
                "Kategori",
                placeholder="Contoh: Material Bangunan"
            )

            spesifikasi = st.text_input(
                "Spesifikasi",
                placeholder="Contoh: Portland Cement 50 Kg"
            )

        with col2:

            satuan = st.selectbox(
                "Satuan",
                [
                    "PCS",
                    "UNIT",
                    "BATANG",
                    "SAK",
                    "KG",
                    "TON",
                    "M3",
                    "M2",
                    "M1",
                    "LITER",
                    "BOX",
                    "SET"
                ]
            )

            stok_minimum = st.number_input(
                "Stok Minimum",
                min_value=0.0,
                value=0.0,
                step=1.0
            )

            lokasi = st.text_input(
                "Lokasi / Rak",
                placeholder="Rak A-01"
            )

        submit = st.form_submit_button(
            "💾 Simpan Material",
            type="primary"
        )

        if submit:

            if not kode.strip():

                st.error(
                    "Kode material wajib diisi."
                )

            elif not nama.strip():

                st.error(
                    "Nama material wajib diisi."
                )

            else:

                sukses, pesan = tambah_material(
                    kode.strip(),
                    nama.strip(),
                    kategori.strip(),
                    spesifikasi.strip(),
                    satuan,
                    stok_minimum,
                    lokasi.strip()
                )

                if sukses:

                    st.success(pesan)

                    st.rerun()

                else:

                    st.error(pesan)

    st.divider()

    st.subheader(
        "📋 Daftar Material"
    )

    df = ambil_material()

    st.dataframe(
        df,
        use_container_width=True,
        hide_index=True
    )


# =========================================================
# MATERIAL MASUK / KELUAR
# =========================================================

elif menu in [
    "📥 Material Masuk",
    "📤 Material Keluar"
]:

    if menu == "📥 Material Masuk":

        jenis = "MATERIAL MASUK"

        judul = "📥 Material Masuk"

    else:

        jenis = "MATERIAL KELUAR"

        judul = "📤 Material Keluar"

    st.title(judul)

    material = ambil_material()

    if material.empty:

        st.warning(
            "Master material masih kosong."
        )

        st.stop()

    kode_list = material[
        "kode_material"
    ].tolist()

    with st.form(
        "form_transaksi",
        clear_on_submit=True
    ):

        col1, col2 = st.columns(2)

        with col1:

            nomor_dokumen = st.text_input(
                "Nomor Dokumen",
                placeholder="SJ-001/IX/2026"
            )

            tanggal = st.date_input(
                "Tanggal",
                value=date.today()
            )

            kode_material = st.selectbox(
                "Kode Material",
                kode_list
            )

            data_material = material[
                material["kode_material"]
                == kode_material
            ].iloc[0]

            nama_material = data_material[
                "nama_material"
            ]

            spesifikasi = data_material[
                "spesifikasi"
            ]

            satuan = data_material[
                "satuan"
            ]

            st.info(
                f"**{nama_material}**\n\n"
                f"Spesifikasi: {spesifikasi}\n\n"
                f"Satuan: {satuan}"
            )

        with col2:

            jumlah = st.number_input(
                "Jumlah",
                min_value=0.01,
                value=1.0,
                step=1.0
            )

            if jenis == "MATERIAL MASUK":

                pihak_label = "Supplier"

            else:

                pihak_label = "Pemakai / Departemen"

            supplier_pemakai = st.text_input(
                pihak_label
            )

            proyek = st.text_input(
                "Proyek / Lokasi Pekerjaan",
                placeholder="Proyek A"
            )

            keterangan = st.text_area(
                "Keterangan"
            )

            foto = st.file_uploader(
                "Foto Dokumen / Material",
                type=[
                    "jpg",
                    "jpeg",
                    "png",
                    "pdf"
                ]
            )

        submit = st.form_submit_button(
            "💾 Simpan Transaksi",
            type="primary"
        )

        if submit:

            # ---------------------------------------------
            # CEK STOK KELUAR
            # ---------------------------------------------

            if jenis == "MATERIAL KELUAR":

                stok = hitung_stok()

                data_stok = stok[
                    stok["kode_material"]
                    == kode_material
                ]

                if not data_stok.empty:

                    stok_tersedia = float(
                        data_stok.iloc[0][
                            "stok_akhir"
                        ]
                    )

                    if jumlah > stok_tersedia:

                        st.error(
                            f"Stok tidak mencukupi. "
                            f"Stok tersedia: "
                            f"{stok_tersedia:,.2f} "
                            f"{satuan}"
                        )

                        st.stop()

            # ---------------------------------------------
            # UPLOAD FOTO
            # ---------------------------------------------

            nama_file = ""

            if foto:

                timestamp = datetime.now().strftime(
                    "%Y%m%d_%H%M%S"
                )

                nama_file = (
                    timestamp
                    + "_"
                    + foto.name
                )

                path_file = os.path.join(
                    UPLOAD_DIR,
                    nama_file
                )

                with open(
                    path_file,
                    "wb"
                ) as file:

                    file.write(
                        foto.getbuffer()
                    )

            # ---------------------------------------------
            # SIMPAN
            # ---------------------------------------------

            simpan_transaksi(
                nomor_dokumen,
                str(tanggal),
                jenis,
                kode_material,
                nama_material,
                spesifikasi,
                jumlah,
                satuan,
                supplier_pemakai,
                proyek,
                keterangan,
                nama_file
            )

            st.success(
                f"{judul} berhasil disimpan."
            )

            st.rerun()


# =========================================================
# RIWAYAT TRANSAKSI
# =========================================================

elif menu == "📋 Riwayat Transaksi":

    st.title(
        "📋 Riwayat Transaksi"
    )

    df = ambil_transaksi()

    if df.empty:

        st.info(
            "Belum ada transaksi."
        )

    else:

        col1, col2, col3 = st.columns(3)

        with col1:

            jenis_filter = st.selectbox(
                "Jenis Transaksi",
                [
                    "SEMUA",
                    "MATERIAL MASUK",
                    "MATERIAL KELUAR"
                ]
            )

        with col2:

            tanggal_awal = st.date_input(
                "Tanggal Awal",
                value=date.today()
            )

        with col3:

            tanggal_akhir = st.date_input(
                "Tanggal Akhir",
                value=date.today()
            )

        keyword = st.text_input(
            "🔎 Cari Material / Dokumen / Proyek"
        )

        df["tanggal_date"] = pd.to_datetime(
            df["tanggal"]
        ).dt.date

        filtered = df[
            (df["tanggal_date"] >= tanggal_awal)
            &
            (df["tanggal_date"] <= tanggal_akhir)
        ]

        if jenis_filter != "SEMUA":

            filtered = filtered[
                filtered["jenis"]
                == jenis_filter
            ]

        if keyword.strip():

            keyword = keyword.strip()

            mask = (
                filtered[
                    "kode_material"
                ].astype(str)
                .str.contains(
                    keyword,
                    case=False,
                    na=False
                )
                |
                filtered[
                    "nama_material"
                ].astype(str)
                .str.contains(
                    keyword,
                    case=False,
                    na=False
                )
                |
                filtered[
                    "nomor_dokumen"
                ].astype(str)
                .str.contains(
                    keyword,
                    case=False,
                    na=False
                )
                |
                filtered[
                    "proyek"
                ].astype(str)
                .str.contains(
                    keyword,
                    case=False,
                    na=False
                )
            )

            filtered = filtered[mask]

        filtered = filtered.drop(
            columns=["tanggal_date"]
        )

        st.dataframe(
            filtered,
            use_container_width=True,
            hide_index=True
        )


# =========================================================
# LAPORAN EXCEL
# =========================================================

elif menu == "📊 Laporan Excel":

    st.title(
        "📊 Laporan Excel"
    )

    stok = hitung_stok()
    transaksi = ambil_transaksi()

    col1, col2, col3 = st.columns(3)

    col1.metric(
        "Jenis Material",
        len(stok)
    )

    col2.metric(
        "Total Transaksi",
        len(transaksi)
    )

    total_stok = (
        stok["stok_akhir"].sum()
        if not stok.empty
        else 0
    )

    col3.metric(
        "Total Stok",
        f"{total_stok:,.2f}"
    )

    st.divider()

    st.subheader(
        "📥 Download Report"
    )

    st.write(
        "Laporan Excel berisi:"
    )

    st.markdown("""
    **Sheet 1 — Stok Material**
    
    Posisi stok seluruh material.

    **Sheet 2 — Transaksi**
    
    Riwayat material masuk dan keluar.

    **Sheet 3 — Master Material**
    
    Data seluruh master material.
    """)

    file_excel = export_excel()

    nama_file = (
        "Report_Gudang_Material_"
        + datetime.now().strftime(
            "%Y%m%d_%H%M%S"
        )
        + ".xlsx"
    )

    st.download_button(
        label="📥 DOWNLOAD REPORT EXCEL",
        data=file_excel,
        file_name=nama_file,
        mime=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
        type="primary"
    )
```
