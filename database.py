import sqlite3
import os
import shutil
import tempfile


# ============================================================
# LOKASI DATABASE USER
# ============================================================

APP_NAME = "SistemPenugasan"

LOCAL_APPDATA = os.getenv(
    "LOCALAPPDATA",
    os.path.join(os.path.expanduser("~"), "AppData", "Local")
)

APP_DATA_DIR = os.path.join(LOCAL_APPDATA, APP_NAME)
os.makedirs(APP_DATA_DIR, exist_ok=True)

DB_PATH = os.path.join(APP_DATA_DIR, "sistem.db")


# ============================================================
# MIGRASI DATABASE LAMA
# ============================================================

OLD_DB_PATH = os.path.join(
    os.path.dirname(os.path.abspath(__file__)),
    "sistem.db"
)


def migrate_old_database():
    """
    Salin database lama ke AppData hanya jika database aktif belum ada.
    Kegagalan migrasi tidak disembunyikan agar pengguna mengetahui
    jika database lama gagal disalin.
    """
    if (
        os.path.isfile(OLD_DB_PATH)
        and os.path.abspath(OLD_DB_PATH) != os.path.abspath(DB_PATH)
        and not os.path.exists(DB_PATH)
    ):
        os.makedirs(os.path.dirname(DB_PATH), exist_ok=True)
        try:
            shutil.copy2(OLD_DB_PATH, DB_PATH)
        except OSError as exc:
            raise RuntimeError(
                "Database lama ditemukan, tetapi gagal dipindahkan ke "
                f"folder data aplikasi:\n{DB_PATH}\n\nDetail: {exc}"
            ) from exc


migrate_old_database()


# ============================================================
# CONNECTION
# ============================================================

def get_connection():
    """
    Membuka koneksi SQLite dengan timeout.
    Foreign key sengaja belum diaktifkan pada tahap ini karena database
    lama perlu diperiksa terlebih dahulu untuk memastikan tidak ada
    relasi yatim yang dapat mengganggu operasi aplikasi.
    """
    conn = sqlite3.connect(DB_PATH, timeout=15)
    conn.execute("PRAGMA busy_timeout = 15000")
    return conn


# ============================================================
# UTILITAS FILE SEMENTARA
# ============================================================

def _buat_path_sementara(destination_path, suffix=".tmp"):
    """Membuat nama file sementara di folder yang sama dengan tujuan."""
    folder = os.path.dirname(os.path.abspath(destination_path)) or "."
    os.makedirs(folder, exist_ok=True)

    fd, path = tempfile.mkstemp(
        prefix=".sistempenugasan_",
        suffix=suffix,
        dir=folder
    )
    os.close(fd)
    return path


def _hapus_jika_ada(path):
    try:
        if path and os.path.exists(path):
            os.remove(path)
    except OSError:
        pass


def _validasi_integritas(conn, pesan):
    result = conn.execute("PRAGMA integrity_check").fetchone()
    if not result or result[0] != "ok":
        detail = result[0] if result else "Tidak ada hasil pemeriksaan."
        raise sqlite3.DatabaseError(f"{pesan} Detail: {detail}")


# ============================================================
# BACKUP DATABASE
# ============================================================

def backup_database(destination_path):
    """
    Membuat backup konsisten menggunakan SQLite backup API.

    Backup dibuat ke file sementara terlebih dahulu. File tujuan lama
    baru diganti setelah backup berhasil dan lolos integrity_check.
    Dengan demikian, backup lama tidak terhapus jika proses baru gagal.
    """
    destination_path = os.path.abspath(destination_path)
    active_path = os.path.abspath(DB_PATH)

    if destination_path == active_path:
        raise ValueError("Lokasi backup tidak boleh sama dengan database aktif.")

    folder = os.path.dirname(destination_path) or "."
    os.makedirs(folder, exist_ok=True)

    temp_path = _buat_path_sementara(destination_path, suffix=".backup.tmp")
    source = None
    target = None

    try:
        source = get_connection()
        target = sqlite3.connect(temp_path, timeout=15)

        source.backup(target)
        _validasi_integritas(target, "Hasil backup gagal melewati integrity_check.")
        target.commit()

        target.close()
        target = None
        source.close()
        source = None

        # Penggantian dilakukan setelah file sementara selesai diverifikasi.
        os.replace(temp_path, destination_path)

    except Exception:
        if target is not None:
            target.close()
        if source is not None:
            source.close()
        _hapus_jika_ada(temp_path)
        raise

    return destination_path


# ============================================================
# RESTORE DATABASE
# ============================================================

def restore_database(backup_path):
    """
    Validasi backup, buat safety backup, lalu pulihkan database.

    File pemulihan disiapkan dan diverifikasi di folder yang sama dengan
    database aktif. Database aktif baru diganti setelah seluruh tahap
    penyalinan berhasil.
    """
    backup_path = os.path.abspath(backup_path)
    active_path = os.path.abspath(DB_PATH)

    if not os.path.isfile(backup_path):
        raise FileNotFoundError("File backup tidak ditemukan.")

    if active_path == backup_path:
        raise ValueError("File backup tidak boleh sama dengan database aktif.")

    # Validasi file sumber sebelum menyentuh database aktif.
    check_conn = sqlite3.connect(backup_path, timeout=15)
    try:
        _validasi_integritas(check_conn, "File backup rusak atau tidak valid.")

        required = {
            "pegawai",
            "tarif",
            "periode",
            "bagian",
            "anggota_bagian",
            "jadwal"
        }
        tables = {
            row[0] for row in check_conn.execute(
                "SELECT name FROM sqlite_master WHERE type='table'"
            ).fetchall()
        }
        missing = required - tables

        if missing:
            raise sqlite3.DatabaseError(
                "File bukan backup Sistem Penugasan Pegawai. "
                "Tabel hilang: " + ", ".join(sorted(missing))
            )
    finally:
        check_conn.close()

    os.makedirs(os.path.dirname(active_path), exist_ok=True)

    temp_path = _buat_path_sementara(active_path, suffix=".restore.tmp")
    safety_path = active_path + ".before-restore.bak"

    source = None
    target = None

    try:
        source = sqlite3.connect(backup_path, timeout=15)
        target = sqlite3.connect(temp_path, timeout=15)

        source.backup(target)
        target.commit()
        _validasi_integritas(
            target,
            "Salinan pemulihan gagal diverifikasi."
        )

        target.close()
        target = None
        source.close()
        source = None

        # Buat safety backup sebelum mengganti database aktif.
        if os.path.exists(active_path):
            backup_database(safety_path)

        # Jika gagal (misalnya database masih dibuka proses lain di Windows),
        # database aktif tetap dipertahankan.
        os.replace(temp_path, active_path)

    except Exception:
        if target is not None:
            target.close()
        if source is not None:
            source.close()
        _hapus_jika_ada(temp_path)
        raise

    return safety_path if os.path.exists(safety_path) else None


# ============================================================
# CREATE DATABASE
# ============================================================

def create_database():
    conn = get_connection()
    try:
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
                INSERT OR IGNORE INTO tarif (golongan, nominal)
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
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()
