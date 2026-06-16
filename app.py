import streamlit as st
import pandas as pd

# ==================================================
# KONFIGURASI HALAMAN
# ==================================================
st.set_page_config(
    page_title="Sistem Stok Barang",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# ==================================================
# CSS MODERN DASHBOARD
# ==================================================
st.markdown("""
<style>

.stApp{
    background-color:#f4f6f9;
}

/* HEADER */
.custom-header{
    background:linear-gradient(135deg,#118d80,#0f766e);
    color:white;
    padding:25px;
    border-radius:18px;
    margin-bottom:25px;
    box-shadow:0 4px 15px rgba(0,0,0,0.15);
}

.custom-header h1{
    margin:0;
}

.custom-header h3{
    margin-top:5px;
    font-weight:400;
}

/* CARD RINGKASAN */
.card-ringkasan{
    background:white;
    border-radius:18px;
    padding:20px;
    margin-bottom:18px;
    box-shadow:0 4px 15px rgba(0,0,0,0.08);
}

.card-top{
    display:flex;
    align-items:center;
    gap:15px;
}

.icon-circle{
    width:70px;
    height:70px;
    border-radius:50%;
    display:flex;
    justify-content:center;
    align-items:center;
    font-size:35px;
    color:white;
}

.hijau{
    background:linear-gradient(135deg,#20c997,#118d80);
}

.kuning{
    background:linear-gradient(135deg,#ffb020,#ff8f00);
}

.biru{
    background:linear-gradient(135deg,#60a5fa,#2563eb);
}

.judul-card{
    font-size:16px;
    font-weight:bold;
}

.angka-card{
    font-size:48px;
    font-weight:bold;
    line-height:1;
}

.keterangan{
    margin-top:10px;
    color:#6b7280;
    font-size:14px;
}

/* INPUT */
div[data-baseweb="input"]{
    border:1px solid #cbd5e1 !important;
    border-radius:8px !important;
    background:white !important;
}

div[data-baseweb="select"] > div{
    border:1px solid #cbd5e1 !important;
    border-radius:8px !important;
    background:white !important;
}

/* TABEL */
.table-gudang{
    width:100%;
    border-collapse:collapse;
    background:white;
    border-radius:15px;
    overflow:hidden;
    box-shadow:0 4px 15px rgba(0,0,0,0.08);
}

.table-gudang th{
    background:linear-gradient(135deg,#118d80,#0f766e);
    color:white;
    padding:12px;
    text-align:left;
    border:1px solid #dfe6e9;
}

.table-gudang td{
    padding:12px;
    border:1px solid #e5e7eb;
}

.table-gudang tr:nth-child(even){
    background:#f8fafc;
}

/* BUTTON */
div.stButton > button{
    border-radius:10px !important;
    font-weight:bold !important;
}

/* WARNA TOMBOL */
div.element-container:has(.marker-hijau)
+ div.element-container button{
    background:#22c55e !important;
    color:white !important;
}

div.element-container:has(.marker-merah)
+ div.element-container button{
    background:#ef4444 !important;
    color:white !important;
}

</style>
""", unsafe_allow_html=True)

# ==================================================
# HEADER
# ==================================================
st.markdown("""
<div class="custom-header">
    <h1>📦 SISTEM STOK BARANG</h1>
    <h3>REKTORAT UNIVERSITAS WINAYA MUKTI</h3>
</div>
""", unsafe_allow_html=True)
# ==================================================
# FITUR ZOOM
# ==================================================

zoom = st.slider(
    "🔍 Zoom Tampilan",
    min_value=70,
    max_value=130,
    value=90,
    step=5
)

st.markdown(f"""
<style>
html {{
    zoom:{zoom}%;
}}
</style>
""", unsafe_allow_html=True)
# ==================================================
# DATABASE AWAL
# ==================================================
if "df_stok" not in st.session_state:

    data_awal = {
        "ID Barang":["1","2","3","4","5"],
        "Nama Barang":[
            "Kertas A4",
            "Pena Pilot",
            "Tinta Printer",
            "Buku Tulis",
            "Spidol"
        ],
        "Kategori":[
            "ATK",
            "ATK",
            "Elektronik",
            "ATK",
            "ATK"
        ],
        "Jumlah Stok":[35,15,5,20,0],
        "Satuan":[
            "Rim",
            "Pcs",
            "Pcs",
            "Pcs",
            "Pcs"
        ],
        "Pengambil":[
            "FEB",
            "FAHUTAN",
            "FAPERTA",
            "FEB",
            "FTPA"
        ]
    }

    st.session_state.df_stok = pd.DataFrame(data_awal)

# ==================================================
# STATUS STOK
# ==================================================
def status_stok(jumlah):
    if jumlah > 20:
        return "🟢 Good Stock"
    elif jumlah > 0:
        return "🟡 Low Stock"
    else:
        return "🔴 Out Of Stock"

st.session_state.df_stok["Status"] = (
    st.session_state.df_stok["Jumlah Stok"]
    .apply(status_stok)
)

# ==================================================
# LAYOUT
# ==================================================
col_kiri, col_tengah, col_kanan = st.columns(
    [1.1, 3, 1.2],
    gap="large"
)

# ==================================================
# KOLOM KIRI
# ==================================================
with col_kiri:

    st.subheader("Ringkasan Stok")

    total_produk = len(st.session_state.df_stok)

    barang_langka = len(
        st.session_state.df_stok[
            st.session_state.df_stok["Jumlah Stok"] < 10
        ]
    )

    total_stok = (
        st.session_state.df_stok["Jumlah Stok"]
        .sum()
    )

    st.markdown(f"""
    <div class="card-ringkasan">
        <div class="card-top">
            <div class="icon-circle hijau">📦</div>
            <div>
                <div class="judul-card" style="color:#118d80;">
                TOTAL PRODUK
                </div>
                <div class="angka-card" style="color:#118d80;">
                {total_produk}
                </div>
            </div>
        </div>
        <div class="keterangan">
        Total seluruh jenis produk yang tersedia
        </div>
    </div>
    """, unsafe_allow_html=True)

    st.markdown(f"""
    <div class="card-ringkasan">
        <div class="card-top">
            <div class="icon-circle kuning">⚠️</div>
            <div>
                <div class="judul-card" style="color:#ff8f00;">
                BARANG HAMPIR HABIS
                </div>
                <div class="angka-card" style="color:#ff8f00;">
                {barang_langka}
                </div>
            </div>
        </div>
        <div class="keterangan">
        Segera lakukan restock
        </div>
    </div>
    """, unsafe_allow_html=True)

    st.markdown(f"""
    <div class="card-ringkasan">
        <div class="card-top">
            <div class="icon-circle biru">📊</div>
            <div>
                <div class="judul-card" style="color:#2563eb;">
                TOTAL STOK
                </div>
                <div class="angka-card" style="color:#2563eb;">
                {total_stok}
                </div>
            </div>
        </div>
        <div class="keterangan">
        Total keseluruhan stok barang
        </div>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("### Grafik Stok")

    chart_data = (
        st.session_state.df_stok[
            ["Nama Barang","Jumlah Stok"]
        ]
        .set_index("Nama Barang")
    )

    st.bar_chart(chart_data)

# ==================================================
# KOLOM TENGAH
# ==================================================
with col_tengah:

    st.subheader("Manajemen Stok")

    c1,c2,c3,c4,c5,c6 = st.columns(
        [1.5,1.5,0.8,1.2,1.4,1]
    )

    with c1:
        input_nama = st.text_input(
            "Nama Barang",
            placeholder="Misal: Penghapus"
        )

    with c2:
        input_kategori = st.text_input(
            "Kategori",
            placeholder="Misal: ATK"
        )

    with c3:
        input_qty = st.number_input(
            "Jml",
            min_value=0,
            step=1
        )

    with c4:
        input_satuan = st.selectbox(
            "Satuan",
            ["Pcs","Rim","Lusin"]
        )

    with c5:
        input_pengambil = st.selectbox(
            "Pengambil",
            [
                "FEB",
                "FAHUTAN",
                "FAPERTA",
                "FTPA"
            ]
        )

    with c6:
        st.markdown("<br>", unsafe_allow_html=True)

        btn_add = st.button(
            "➕ Add Item",
            use_container_width=True
        )

    if btn_add and input_nama != "":

        id_baru = str(
            len(st.session_state.df_stok) + 1
        )

        data_baru = pd.DataFrame({
            "ID Barang":[id_baru],
            "Nama Barang":[input_nama],
            "Kategori":[input_kategori],
            "Jumlah Stok":[input_qty],
            "Satuan":[input_satuan],
            "Pengambil":[input_pengambil]
        })

        st.session_state.df_stok = pd.concat(
            [
                st.session_state.df_stok,
                data_baru
            ],
            ignore_index=True
        )

        st.rerun()

    tabel_html = (
        st.session_state.df_stok
        .to_html(
            classes="table-gudang",
            index=False,
            escape=False
        )
    )

    st.markdown(
        tabel_html,
        unsafe_allow_html=True
    )

# ==================================================
# KOLOM KANAN
# ==================================================
with col_kanan:

    st.subheader("Transaksi")

    input_barcode = st.text_input(
        "Scan Barcode / Input ID",
        placeholder="Ketik ID Barang"
    )

    qty = st.number_input(
        "Quantity",
        min_value=1,
        step=1
    )

    b1,b2 = st.columns(2)

    with b1:
        st.markdown(
            '<div class="marker-hijau" style="display:none;"></div>',
            unsafe_allow_html=True
        )

        btn_masuk = st.button(
            "Barang Masuk",
            use_container_width=True
        )

    with b2:
        st.markdown(
            '<div class="marker-merah" style="display:none;"></div>',
            unsafe_allow_html=True
        )

        btn_keluar = st.button(
            "Barang Keluar",
            use_container_width=True
        )

    # BARANG MASUK
    if btn_masuk:

        if input_barcode in st.session_state.df_stok["ID Barang"].values:

            st.session_state.df_stok.loc[
                st.session_state.df_stok["ID Barang"] == input_barcode,
                "Jumlah Stok"
            ] += qty

            st.success("Stok berhasil ditambahkan")
            st.rerun()

        else:
            st.error("ID Barang tidak ditemukan")

    # BARANG KELUAR
    if btn_keluar:

        if input_barcode in st.session_state.df_stok["ID Barang"].values:

            stok_sekarang = (
                st.session_state.df_stok.loc[
                    st.session_state.df_stok["ID Barang"] == input_barcode,
                    "Jumlah Stok"
                ].values[0]
            )

            if stok_sekarang >= qty:

                st.session_state.df_stok.loc[
                    st.session_state.df_stok["ID Barang"] == input_barcode,
                    "Jumlah Stok"
                ] -= qty

                st.success("Stok berhasil dikurangi")
                st.rerun()

            else:
                st.error("Stok tidak mencukupi")

        else:
            st.error("ID Barang tidak ditemukan")