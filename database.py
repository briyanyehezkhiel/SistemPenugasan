import sqlite3
import os

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(BASE_DIR, "sistem.db")


def get_connection():
    return sqlite3.connect(DB_PATH)


def create_database():

    conn = get_connection()
    cursor = conn.cursor()

    # ==============================
    # DATA PEGAWAI
    # ==============================

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS pegawai (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            no INTEGER,
            nama TEXT NOT NULL,
            nip TEXT,
            pangkat TEXT,
            golongan_asli TEXT,
            golongan_utama TEXT,
            jabatan TEXT,
            tarif INTEGER NOT NULL,
            aktif INTEGER DEFAULT 1
        )
    """)

    # ==============================
    # TARIF
    # ==============================

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS tarif (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            golongan TEXT UNIQUE NOT NULL,
            nominal INTEGER NOT NULL
        )
    """)

    tarif_awal = [
        ("II", 75000),
        ("III", 80000),
        ("IV", 85000)
    ]

    for golongan, nominal in tarif_awal:

        cursor.execute("""
            INSERT OR IGNORE INTO tarif
            (golongan, nominal)
            VALUES (?, ?)
        """, (golongan, nominal))

    # ==============================
    # PERIODE BULAN
    # ==============================

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS periode (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            bulan INTEGER NOT NULL,
            tahun INTEGER NOT NULL,
            UNIQUE(bulan, tahun)
        )
    """)

    # ==============================
    # BAGIAN / KEGIATAN
    # ==============================

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS bagian (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            periode_id INTEGER NOT NULL,
            kode TEXT NOT NULL,
            kegiatan TEXT,
            lokasi TEXT,
            dana INTEGER DEFAULT 0,
            UNIQUE(periode_id, kode),
            FOREIGN KEY (periode_id)
                REFERENCES periode(id)
        )
    """)

    # ==============================
    # PEGAWAI YANG MASUK BAGIAN
    # ==============================

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS anggota_bagian (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            bagian_id INTEGER NOT NULL,
            pegawai_id INTEGER NOT NULL,
            UNIQUE(bagian_id, pegawai_id),
            FOREIGN KEY (bagian_id)
                REFERENCES bagian(id),
            FOREIGN KEY (pegawai_id)
                REFERENCES pegawai(id)
        )
    """)

    # ==============================
    # JADWAL
    # ==============================

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS jadwal (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            periode_id INTEGER NOT NULL,
            bagian_id INTEGER NOT NULL,
            pegawai_id INTEGER NOT NULL,
            hari INTEGER NOT NULL,
            UNIQUE(periode_id, pegawai_id, hari),
            FOREIGN KEY (periode_id)
                REFERENCES periode(id),
            FOREIGN KEY (bagian_id)
                REFERENCES bagian(id),
            FOREIGN KEY (pegawai_id)
                REFERENCES pegawai(id)
        )
    """)

    conn.commit()
    conn.close()