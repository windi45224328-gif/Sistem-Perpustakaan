"""database.py — Skema, data statis (Excel), dan aturan bisnis SIPINTER."""
import sqlite3
import re
from pathlib import Path
from datetime import date, datetime, timedelta

DB_PATH = Path(__file__).resolve().parent / "perpustakaan.db"

# ================= DATA STATIS (DARI EXCEL) =================
DATA_JURUSAN_PRODI_KELAS = {
    "Administrasi Bisnis": [
        ("Administrasi Bisnis (AB)", "AB-1A"), ("Administrasi Bisnis Terapan (ABT)", "ABT-1A"),
        ("Manajemen Bisnis Internasional (MBI)", "MBI-1A"), ("Manajemen Pemasaran (MP)", "MP-1A")
    ],
    "Akuntansi": [
        ("Akuntansi (AK)", "AK-1A"), ("Akuntansi Manajerial (AM)", "AM-1A"),
        ("Komputerisasi Akuntansi (KA)", "KA-1A"), ("Keuangan dan Perbankan (KP)", "KP-1A"),
        ("Analisis Keuangan (KU)", "KU-1A"), ("Perbankan Syariah (PS)", "PS-1A")
    ],
    "Teknik Elektro": [
        ("Teknik Elektronika (EK)", "EK-1A"), ("Teknik Informatika (IK)", "IK-1A"),
        ("Teknologi Rekayasa Instalasi Listrik (IL)", "IL-1A"), ("Teknik Listrik (LT)", "LT-1A"),
        ("Teknologi Rekayasa Elektronika (RE)", "RE-1A"), ("Teknik Elektronika (TE)", "TE-1A"),
        ("Teknik Informatika (TI)", "TI-1A"), ("Teknologi Rekayasa Komputer (TK)", "TK-1A")
    ],
    "Teknik Mesin": [
        ("Teknik Konversi Energi (KE)", "KE-1A"), ("Teknik Mesin (ME)", "ME-1A"),
        ("Teknik Mesin Produksi dan Perawatan (MS)", "MS-1A"), ("Teknologi Rekayasa Pembangkit Energi (PE)", "PE-1A")
    ],
    "Teknik Sipil": [
        ("Konstruksi Gedung (KG)", "KG-1A"), ("Konstruksi Sipil (KS)", "KS-1A"),
        ("Perancangan Jalan dan Jembatan (PJJ)", "PJJ-1A"), ("Teknik Perawatan dan Perbaikan Gedung (PPG)", "PPG-1A")
    ]
}

KATEGORI_BUKU = {
    "Fiksi & Fantasi": "FF", "Fiksi & Romansa": "FR", "Sastra Indonesia": "SI",
    "Fiksi & Sastra": "FS", "Fiksi Remaja": "FRe", "Sejarah & Budaya": "SB",
    "Sains Anak": "SA", "Pengembangan Diri": "PD", "Bisnis & Investasi": "BI",
    "Dongeng Anak": "DA", "Aktivitas Anak": "AA", "Komik & Manga": "KM",
    "Edukasi Anak": "EA", "Kuliner": "KL", "Fiksi & Thriller": "FT",
    "Fiksi Sejarah": "FSe", "Musik": "MU", "Cerita Pendek": "CP",
    "Fiksi & Keluarga": "FK", "Perjalanan & Esai": "PE", "Biografi": "BIO",
    "Horor": "HOR", "Buku Anak": "BA", "Fiksi & Petualangan": "FP",
    "Fiksi & Pendidikan": "FPe", "Fiksi Anak & Keluarga": "FAK"
}

SKEMA = """
PRAGMA foreign_keys = ON;
CREATE TABLE IF NOT EXISTS penulis (
    id INTEGER PRIMARY KEY AUTOINCREMENT, nama_penulis TEXT NOT NULL UNIQUE);
CREATE TABLE IF NOT EXISTS penerbit (
    id INTEGER PRIMARY KEY AUTOINCREMENT, nama_penerbit TEXT NOT NULL UNIQUE);
CREATE TABLE IF NOT EXISTS buku (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    kode_buku TEXT NOT NULL UNIQUE, isbn TEXT NOT NULL UNIQUE, judul TEXT NOT NULL,
    kategori TEXT NOT NULL, id_penulis INTEGER NOT NULL, id_penerbit INTEGER NOT NULL,
    tahun_terbit INTEGER NOT NULL, stok_total INTEGER NOT NULL DEFAULT 0,
    stok_tersedia INTEGER NOT NULL DEFAULT 0,
    FOREIGN KEY (id_penulis) REFERENCES penulis(id),
    FOREIGN KEY (id_penerbit) REFERENCES penerbit(id));
CREATE TABLE IF NOT EXISTS peminjam (
    id INTEGER PRIMARY KEY AUTOINCREMENT, nim TEXT NOT NULL UNIQUE,
    nama TEXT NOT NULL, program_studi TEXT NOT NULL, 
    jurusan TEXT NOT NULL, kelas TEXT NOT NULL, no_whatsapp TEXT NOT NULL);
CREATE TABLE IF NOT EXISTS transaksi (
    id INTEGER PRIMARY KEY AUTOINCREMENT, kode_transaksi TEXT NOT NULL UNIQUE,
    id_peminjam INTEGER NOT NULL, id_buku INTEGER NOT NULL,
    tanggal_pinjam TEXT NOT NULL, tanggal_tenggat TEXT NOT NULL,
    jumlah_perpanjang INTEGER NOT NULL DEFAULT 0, tanggal_kembali TEXT,
    hari_terlambat INTEGER NOT NULL DEFAULT 0, total_denda INTEGER NOT NULL DEFAULT 0,
    denda_lunas INTEGER NOT NULL DEFAULT 0, status TEXT NOT NULL DEFAULT 'DIPINJAM',
    FOREIGN KEY (id_peminjam) REFERENCES peminjam(id),
    FOREIGN KEY (id_buku) REFERENCES buku(id));
CREATE TABLE IF NOT EXISTS log_whatsapp (
    id INTEGER PRIMARY KEY AUTOINCREMENT, id_transaksi INTEGER NOT NULL,
    jenis TEXT NOT NULL, no_tujuan TEXT NOT NULL, isi_pesan TEXT NOT NULL,
    status TEXT NOT NULL, waktu_kirim TEXT NOT NULL,
    FOREIGN KEY (id_transaksi) REFERENCES transaksi(id));
CREATE TABLE IF NOT EXISTS pengaturan (
    id INTEGER PRIMARY KEY CHECK (id = 1),
    lama_pinjam INTEGER NOT NULL DEFAULT 7,
    maks_perpanjang INTEGER NOT NULL DEFAULT 1,
    tarif_denda_per_hari INTEGER NOT NULL DEFAULT 1000,
    token_whatsapp TEXT DEFAULT '');
INSERT OR IGNORE INTO pengaturan (id) VALUES (1);
"""

def koneksi():
    conn = sqlite3.connect(DB_PATH, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    return conn

def inisialisasi():
    conn = koneksi()
    conn.executescript(SKEMA)
    conn.commit()
    if conn.execute("SELECT COUNT(*) n FROM buku").fetchone()["n"] == 0:
        _seed(conn)
    conn.close()

def q(sql, params=()):
    conn = koneksi()
    rows = [dict(r) for r in conn.execute(sql, params).fetchall()]
    conn.close()
    return rows

def x(sql, params=()):
    conn = koneksi()
    cur = conn.execute(sql, params)
    conn.commit()
    last = cur.lastrowid
    conn.close()
    return last

def _seed(conn):
    # 1. Seed Peminjam (data demo)
    data_peminjam = [
        ("0.00.00.0.01", "Mahasiswa Contoh 01", "Administrasi Bisnis (AB)", "Administrasi Bisnis", "AB-1A", "620000000001"),
        ("0.00.00.0.02", "Mahasiswa Contoh 02", "Akuntansi (AK)", "Akuntansi", "AK-1A", "620000000002"),
        ("0.00.00.0.03", "Mahasiswa Contoh 03", "Teknik Elektronika (EK)", "Teknik Elektro", "EK-1A", "620000000003"),
        ("0.00.00.0.04", "Mahasiswa Contoh 04", "Teknik Konversi Energi (KE)", "Teknik Mesin", "KE-1A", "620000000004"),
        ("0.00.00.0.05", "Mahasiswa Contoh 05", "Konstruksi Gedung (KG)", "Teknik Sipil", "KG-1A", "620000000005"),
        ("0.00.00.0.06", "Mahasiswa Contoh 06", "Teknik Informatika (TI)", "Teknik Elektro", "TI-1A", "620000000006")
    ]
    for p in data_peminjam:
        conn.execute("""INSERT INTO peminjam (nim,nama,program_studi,jurusan,kelas,no_whatsapp)
            VALUES (?,?,?,?,?,?)""", p)

    # 2. Seed Penulis & Penerbit
    penulis_list = ["J.S. Khairen", "Ra Amalia", "Aya", "Ahmad Tohari", "Paulo Coelho", "Risma Ridha Anissa", "Andrea Hirata", "Tere Liye", "Dee Lestari"]
    penerbit_list = ["Bentang Pustaka", "Sena Books", "Bukune", "Gramedia Pustaka Utama", "Elex Media Komputindo", "Kepustakaan Populer Gramedia"]
    
    for p in penulis_list: conn.execute("INSERT INTO penulis (nama_penulis) VALUES (?)", (p,))
    for p in penerbit_list: conn.execute("INSERT INTO penerbit (nama_penerbit) VALUES (?)", (p,))

    # 3. Seed Buku (Sampel 8 Buku)
    buku_sampel = [
        ("9786020523316", "Melangkah", "Fiksi & Fantasi", 1, 1, 2020, 10),
        ("9786237495543", "Titik Akhir", "Fiksi & Romansa", 2, 2, 2021, 8),
        ("9786022204534", "Mahika", "Fiksi & Romansa", 3, 3, 2019, 9),
        ("9786020320595", "Orang-Orang Proyek", "Sastra Indonesia", 4, 4, 2018, 7),
        ("9789792268362", "Zahir", "Fiksi & Sastra", 5, 4, 2005, 6),
        ("9786022916628", "Laskar Pelangi", "Fiksi & Pendidikan", 7, 1, 2005, 12),
        ("9786239726287", "Bulan", "Fiksi & Fantasi", 8, 3, 2022, 10),
        ("9786028811613", "Filosofi Kopi", "Cerita Pendek", 9, 1, 2015, 7)
    ]
    
    for b in buku_sampel:
        kat = b[2]
        kode_kat = KATEGORI_BUKU.get(kat, "UM")
        last = conn.execute("SELECT kode_buku FROM buku WHERE kode_buku LIKE ? ORDER BY id DESC LIMIT 1", (f"{kode_kat}-%",)).fetchone()
        if last:
            try: urut = int(last["kode_buku"].split("-")[1]) + 1
            except: urut = 1
        else: urut = 1
        
        kode_buku = f"{kode_kat}-{urut:03d}"
        conn.execute("""INSERT INTO buku (kode_buku,isbn,judul,kategori,id_penulis,id_penerbit,tahun_terbit,stok_total,stok_tersedia)
            VALUES (?,?,?,?,?,?,?,?,?)""", (kode_buku, b[0], b[1], b[2], b[3], b[4], b[5], b[6], b[6]))

    # 4. Seed Transaksi (Dummy Data)
    t = date.today()
    trx = [
        # Sedang Dipinjam (Tenggat future)
        ("TRX-20240501-001", 1, 1, str(t - timedelta(days=2)), str(t + timedelta(days=5)), 0, None, 0, 0, 0, "DIPINJAM"),
        # Terlambat (Tenggat lewat)
        ("TRX-20240501-002", 2, 2, str(t - timedelta(days=10)), str(t - timedelta(days=3)), 0, None, 3, 3000, 0, "TERLAMBAT"),
        # Dikembalikan tepat waktu
        ("TRX-20240501-003", 3, 3, str(t - timedelta(days=20)), str(t - timedelta(days=13)), 0, str(t - timedelta(days=15)), 0, 0, 0, "DIKEMBALIKAN"),
        # Dikembalikan terlambat (Denda Lunas)
        ("TRX-20240501-004", 4, 4, str(t - timedelta(days=30)), str(t - timedelta(days=23)), 0, str(t - timedelta(days=20)), 3, 3000, 1, "DIKEMBALIKAN"),
    ]
    for tx in trx:
        conn.execute("""INSERT INTO transaksi (kode_transaksi,id_peminjam,id_buku,
            tanggal_pinjam,tanggal_tenggat,jumlah_perpanjang,tanggal_kembali,
            hari_terlambat,total_denda,denda_lunas,status) VALUES (?,?,?,?,?,?,?,?,?,?,?)""", tx)
    
    # Sesuaikan stok
    conn.execute("UPDATE buku SET stok_tersedia=stok_tersedia-1 WHERE id IN (1,2)")
    conn.commit()

def pengaturan():
    return q("SELECT * FROM pengaturan WHERE id=1")[0]

def validasi_nim(nim):
    return isinstance(nim, str) and re.fullmatch(r"[0-9]+", nim) is not None

def validasi_wa(nomor):
    return isinstance(nomor, str) and re.fullmatch(r"62[0-9]{8,13}", nomor) is not None

def refresh_status():
    tarif = pengaturan()["tarif_denda_per_hari"]
    today = date.today()
    for t in q("SELECT * FROM transaksi WHERE status IN ('DIPINJAM','TERLAMBAT')"):
        tenggat = datetime.strptime(t["tanggal_tenggat"], "%Y-%m-%d").date()
        if today > tenggat:
            telat = (today - tenggat).days
            x("UPDATE transaksi SET status='TERLAMBAT',hari_terlambat=?,total_denda=? WHERE id=?",
              (telat, telat * tarif, t["id"]))
        else:
            x("UPDATE transaksi SET status='DIPINJAM',hari_terlambat=0,total_denda=0 WHERE id=?",
              (t["id"],))

def hitung_denda(t):
    if t["status"] == "DIKEMBALIKAN":
        return t["hari_terlambat"], t["total_denda"]
    tenggat = datetime.strptime(t["tanggal_tenggat"], "%Y-%m-%d").date()
    today = date.today()
    if today > tenggat:
        telat = (today - tenggat).days
        return telat, telat * pengaturan()["tarif_denda_per_hari"]
    return 0, 0

def cek_boleh_pinjam(id_peminjam):
    aktif = q("SELECT * FROM transaksi WHERE id_peminjam=? AND status IN ('DIPINJAM','TERLAMBAT')", (id_peminjam,))
    if aktif:
        t = aktif[0]
        return False, f"Gagal: Mahasiswa masih memiliki pinjaman aktif ({t['kode_transaksi']}). Harap kembalikan buku terlebih dahulu."
    
    denda = q("SELECT * FROM transaksi WHERE id_peminjam=? AND total_denda>0 AND denda_lunas=0", (id_peminjam,))
    if denda:
        return False, "Gagal: Mahasiswa memiliki denda yang belum dibayar."
        
    return True, ""

def kode_transaksi_baru():
    prefix = f"TRX-{date.today().strftime('%Y%m%d')}-"
    rows = q("SELECT kode_transaksi FROM transaksi WHERE kode_transaksi LIKE ? ORDER BY id DESC LIMIT 1", (prefix + "%",))
    urut = int(rows[0]["kode_transaksi"].split("-")[-1]) + 1 if rows else 1
    return f"{prefix}{urut:03d}"

def generate_kode_buku(kategori):
    kode_kat = KATEGORI_BUKU.get(kategori, "UM")
    rows = q("SELECT kode_buku FROM buku WHERE kode_buku LIKE ? ORDER BY id DESC LIMIT 1", (f"{kode_kat}-%",))
    if rows:
        try: urut = int(rows[0]["kode_buku"].split("-")[1]) + 1
        except: urut = 1
    else: urut = 1
    return f"{kode_kat}-{urut:03d}"