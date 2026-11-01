import sys
import tkinter as tk
import datetime
from tkinter import ttk, messagebox, simpledialog, filedialog
from datetime import datetime
import calendar
import sqlite3
import math
import time
import re
import json
import os
import tempfile
import subprocess
import threading
import urllib.request

try:
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.table import Table, TableStyleInfo
except ImportError:
    Workbook = None
    load_workbook = None
    Font = PatternFill = Alignment = Border = Side = None
    get_column_letter = None
    Table = TableStyleInfo = None

from database import get_connection, create_database


# =========================================================
# WINDOWS DPI / DISPLAY SCALING
# =========================================================

def configure_windows_dpi():
    """
    Aktifkan DPI awareness agar Tkinter menggunakan koordinat fisik yang
    konsisten. Skala internal Tk dikunci ke basis 125% setelah root dibuat.
    """
    try:
        import ctypes
        user32 = ctypes.windll.user32
        try:
            user32.SetProcessDpiAwarenessContext(ctypes.c_void_p(-4))
            return
        except Exception:
            pass
        try:
            ctypes.windll.shcore.SetProcessDpiAwareness(2)
            return
        except Exception:
            pass
        try:
            user32.SetProcessDPIAware()
        except Exception:
            pass
    except Exception:
        pass


# =========================================================
# KONFIGURASI
# =========================================================
APP_VERSION = "1.1.5"
UPDATE_MANIFEST_URL = "https://raw.githubusercontent.com/briyanyehezkhiel/SistemPenugasan/main/update.json"

BAGIAN = []

NAMA_BULAN = [
    "",
    "Januari",
    "Februari",
    "Maret",
    "April",
    "Mei",
    "Juni",
    "Juli",
    "Agustus",
    "September",
    "Oktober",
    "November",
    "Desember"
 ]


# =========================================================
# MODERN UI PALETTE
# =========================================================
UI_BG = "#F5F7FA"
UI_PANEL = "#FFFFFF"
UI_BORDER = "#E2E8F0"
UI_TEXT = "#111827"
UI_MUTED = "#64748B"
UI_PRIMARY = "#0F766E"
UI_PRIMARY_HOVER = "#115E59"
UI_PRIMARY_LIGHT = "#E6F7F4"
UI_SUCCESS = "#15803D"
UI_SUCCESS_LIGHT = "#ECFDF3"
UI_DANGER = "#DC2626"
UI_DANGER_LIGHT = "#FEF2F2"
UI_HEADER = "#0F172A"
UI_HEADER_2 = "#1E293B"
UI_WHITE = "#FFFFFF"
UI_TABLE_HEAD = "#F8FAFC"
UI_TABLE_ALT = "#F8FAFC"
UI_SHADOW = "#E5E7EB"

# Palet tema aplikasi (nilai netral saja yang berubah; warna status kalender
# dan warna kegiatan dipertahankan agar makna visualnya tidak berubah).
THEME_PALETTES = {
    "light": {
        "UI_BG": "#F1F5F9", "UI_PANEL": "#FFFFFF", "UI_BORDER": "#CBD5E1",
        "UI_TEXT": "#172033", "UI_MUTED": "#526176",
        "UI_PRIMARY": "#0F766E", "UI_PRIMARY_HOVER": "#0B5F59",
        "UI_PRIMARY_LIGHT": "#CCFBF1", "UI_SUCCESS_LIGHT": "#DCFCE7",
        "UI_DANGER_LIGHT": "#FEE2E2", "UI_TABLE_HEAD": "#E8EEF5",
        "UI_TABLE_ALT": "#F3F6FA", "UI_SHADOW": "#D7E0EA",
    },
    "dark": {
        "UI_BG": "#0B1220", "UI_PANEL": "#172235", "UI_BORDER": "#34445B",
        "UI_TEXT": "#E8EEF7", "UI_MUTED": "#B0BED0",
        "UI_PRIMARY": "#2DD4BF", "UI_PRIMARY_HOVER": "#14B8A6",
        "UI_PRIMARY_LIGHT": "#164E4A", "UI_SUCCESS_LIGHT": "#153D32",
        "UI_DANGER_LIGHT": "#4A2630", "UI_TABLE_HEAD": "#26364D",
        "UI_TABLE_ALT": "#1D2B40", "UI_SHADOW": "#070D17",
    },
}


# =========================================================
# HELPER
# =========================================================

def rupiah(nilai):

    return "Rp {:,.0f}".format(
        nilai
    ).replace(",", ".")


def parse_rupiah(teks):

    teks = (
        teks
        .replace("Rp", "")
        .replace(".", "")
        .replace(",", "")
        .strip()
    )

    if not teks:
        return 0

    return int(teks)


def nilai_excel_teks(nilai):
    """Mengubah nilai Excel menjadi teks tanpa mengubah data yang sudah berupa teks."""
    if nilai is None:
        return ""
    if isinstance(nilai, bool):
        return str(nilai)
    if isinstance(nilai, int):
        return str(nilai)
    if isinstance(nilai, float):
        if nilai.is_integer():
            return str(int(nilai))
        return str(nilai).strip()
    return str(nilai).strip()


def parse_tarif_excel(nilai):
    """Membaca tarif dari kolom tarif Excel (mis. 80000 atau Rp 80.000)."""
    teks = nilai_excel_teks(nilai)
    if not teks:
        return 0
    try:
        return parse_rupiah(teks)
    except (ValueError, TypeError):
        # Ambil digit saja untuk format Excel yang tidak standar.
        digit = re.sub(r"[^0-9]", "", teks)
        return int(digit) if digit else 0


def golongan_utama_dari(teks):
    """Mengambil golongan utama dari golongan asli, mis. 'III d' -> 'III'."""
    teks = (teks or "").strip().upper()
    if not teks:
        return ""
    return teks.split()[0]


def kode_bagian_berikutnya(kode_list):
    """Menghasilkan kode pertama yang belum dipakai: A, B, C, ..."""
    dipakai = {str(k).strip().upper() for k in kode_list if k}
    nomor = 0
    while True:
        n = nomor
        hasil = ""
        while True:
            hasil = chr(ord("A") + (n % 26)) + hasil
            n = n // 26 - 1
            if n < 0:
                break
        if hasil not in dipakai:
            return hasil
        nomor += 1


def ensure_bagian_schema():
    """Migrasi ringan untuk menambahkan PJ wajib tanpa menghapus data lama."""
    conn = get_connection()
    cursor = conn.cursor()
    columns = [row[1] for row in cursor.execute("PRAGMA table_info(bagian)").fetchall()]
    if "penanggung_jawab_id" not in columns:
        cursor.execute("ALTER TABLE bagian ADD COLUMN penanggung_jawab_id INTEGER")
    if "nama_bagian" not in columns:
        cursor.execute("ALTER TABLE bagian ADD COLUMN nama_bagian TEXT")
    if "komposisi_target" not in columns:
        cursor.execute("ALTER TABLE bagian ADD COLUMN komposisi_target TEXT DEFAULT ''")

    # Hari libur disimpan per periode (bulan/tahun), sehingga setiap bulan
    # dapat memiliki tanggal merah yang berbeda. Hari Minggu tetap nonaktif
    # otomatis dan tidak perlu disimpan di tabel ini.
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS hari_libur (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            periode_id INTEGER NOT NULL,
            hari INTEGER NOT NULL,
            UNIQUE(periode_id, hari)
        )
    """)

    # Desa master dan relasi desa pada jadwal.
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS desa (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nama TEXT NOT NULL UNIQUE,
            aktif INTEGER NOT NULL DEFAULT 1
        )
    """)
    jadwal_columns = [row[1] for row in cursor.execute("PRAGMA table_info(jadwal)").fetchall()]
    if "desa_id" not in jadwal_columns:
        cursor.execute("ALTER TABLE jadwal ADD COLUMN desa_id INTEGER")

    conn.commit()
    conn.close()


def load_desa_aktif():
    conn = get_connection()
    try:
        return conn.execute(
            "SELECT id, nama FROM desa WHERE aktif = 1 ORDER BY nama"
        ).fetchall()
    finally:
        conn.close()


def load_desa_by_id(desa_id):
    conn = get_connection()
    try:
        return conn.execute("SELECT id, nama FROM desa WHERE id = ?", (desa_id,)).fetchone()
    finally:
        conn.close()


# =========================================================
# DATABASE
# =========================================================

def get_or_create_periode(bulan, tahun):

    conn = get_connection()
    cursor = conn.cursor()

    cursor.execute("""
        INSERT OR IGNORE INTO periode
        (bulan, tahun)
        VALUES (?, ?)
    """, (bulan, tahun))

    periode = cursor.execute("""
        SELECT id FROM periode
        WHERE bulan = ? AND tahun = ?
    """, (bulan, tahun)).fetchone()

    periode_id = periode[0]

    # Periode baru minimal memiliki satu kegiatan.
    # Kegiatan berikutnya dibuat lewat tombol + Tambah Kegiatan.
    jumlah = cursor.execute(
        "SELECT COUNT(*) FROM bagian WHERE periode_id = ?",
        (periode_id,)
    ).fetchone()[0]

    if jumlah == 0:
        cursor.execute("""
            INSERT INTO bagian
            (periode_id, kode, kegiatan, lokasi, dana, penanggung_jawab_id)
            VALUES (?, 'A', '', '', 0, NULL)
        """, (periode_id,))

    conn.commit()
    conn.close()

    return periode_id


def load_bagian_codes(periode_id):
    conn = get_connection()
    cursor = conn.cursor()
    rows = cursor.execute(
        "SELECT kode FROM bagian WHERE periode_id = ? ORDER BY id",
        (periode_id,)
    ).fetchall()
    conn.close()
    return [row[0] for row in rows]


def load_hari_libur(periode_id):
    """Mengambil tanggal libur manual untuk periode aktif."""
    conn = get_connection()
    rows = conn.execute(
        "SELECT hari FROM hari_libur WHERE periode_id = ? ORDER BY hari",
        (periode_id,)
    ).fetchall()
    conn.close()
    return {row[0] for row in rows}


def simpan_hari_libur(periode_id, hari_libur):
    """Menyimpan ulang daftar hari libur manual untuk satu periode."""
    conn = get_connection()
    try:
        conn.execute(
            "DELETE FROM hari_libur WHERE periode_id = ?",
            (periode_id,)
        )
        for hari in sorted(set(hari_libur)):
            conn.execute(
                "INSERT INTO hari_libur (periode_id, hari) VALUES (?, ?)",
                (periode_id, hari)
            )
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def nama_hari_libur(periode_id):
    hari = load_hari_libur(periode_id)
    if not hari:
        return "Tidak ada"
    return ", ".join(str(x) for x in sorted(hari))


def load_pegawai():

    conn = get_connection()
    cursor = conn.cursor()

    cursor.execute("""
        SELECT
            id,
            nama,
            nip,
            pangkat,
            golongan_asli,
            golongan_utama,
            jabatan,
            tarif
        FROM pegawai
        WHERE aktif = 1
        ORDER BY nama
    """)

    data = cursor.fetchall()

    conn.close()

    return data


def load_bagian(periode_id, kode):

    conn = get_connection()
    cursor = conn.cursor()

    data = cursor.execute("""
        SELECT
            id,
            kode,
            kegiatan,
            lokasi,
            dana,
            penanggung_jawab_id
        FROM bagian
        WHERE periode_id = ?
        AND kode = ?
    """, (
        periode_id,
        kode
    )).fetchone()

    conn.close()

    return data


def load_anggota(bagian_id):
    conn = get_connection()
    cursor = conn.cursor()

    cursor.execute(
        """
        SELECT
            p.id,
            p.nama
        FROM anggota_bagian ab
        JOIN pegawai p
            ON p.id = ab.pegawai_id
        WHERE ab.bagian_id = ?
        ORDER BY p.nama
        """,
        (bagian_id,)
    )

    data = cursor.fetchall()

    conn.close()

    return data
def load_riwayat_turun(periode_id):
    """Mengambil jumlah penugasan pegawai pada BULAN/periode aktif."""
    conn = get_connection()
    try:
        rows = conn.execute(
            """
            SELECT pegawai_id, COUNT(*) AS bulan_ini
            FROM jadwal
            WHERE periode_id = ?
            GROUP BY pegawai_id
            """,
            (periode_id,)
        ).fetchall()
        return {
            row[0]: {"bulan_ini": row[1] or 0}
            for row in rows
        }
    finally:
        conn.close()


def get_tahun_penugasan():
    """Daftar tahun yang mempunyai data periode, terbaru lebih dahulu."""
    conn = get_connection()
    try:
        rows = conn.execute(
            "SELECT DISTINCT tahun FROM periode ORDER BY tahun DESC"
        ).fetchall()
        return [int(row[0]) for row in rows]
    finally:
        conn.close()


def get_ringkasan_periode(periode_id):
    """Ringkasan data penugasan untuk satu bulan/periode."""
    conn = get_connection()
    try:
        kegiatan = conn.execute("SELECT COUNT(*) FROM bagian WHERE periode_id = ?", (periode_id,)).fetchone()[0]
        anggota = conn.execute(
            "SELECT COUNT(*) FROM anggota_bagian WHERE bagian_id IN (SELECT id FROM bagian WHERE periode_id = ?)",
            (periode_id,)
        ).fetchone()[0]
        jadwal = conn.execute("SELECT COUNT(*) FROM jadwal WHERE periode_id = ?", (periode_id,)).fetchone()[0]
        row = conn.execute("SELECT bulan, tahun FROM periode WHERE id = ?", (periode_id,)).fetchone()
        return {
            "periode": 1 if row else 0,
            "bulan": row[0] if row else None,
            "tahun": row[1] if row else None,
            "kegiatan": kegiatan,
            "anggota": anggota,
            "jadwal": jadwal,
        }
    finally:
        conn.close()


def hapus_data_penugasan_periode(periode_id):
    """Menghapus seluruh data penugasan satu bulan, tanpa menghapus master pegawai."""
    conn = get_connection()
    try:
        bagian_ids = [r[0] for r in conn.execute(
            "SELECT id FROM bagian WHERE periode_id = ?", (periode_id,)
        ).fetchall()]
        if bagian_ids:
            ph = ",".join("?" * len(bagian_ids))
            conn.execute(f"DELETE FROM jadwal WHERE bagian_id IN ({ph})", bagian_ids)
            conn.execute(f"DELETE FROM anggota_bagian WHERE bagian_id IN ({ph})", bagian_ids)
        conn.execute("DELETE FROM jadwal WHERE periode_id = ?", (periode_id,))
        conn.execute("DELETE FROM hari_libur WHERE periode_id = ?", (periode_id,))
        conn.execute("DELETE FROM bagian WHERE periode_id = ?", (periode_id,))
        conn.execute("DELETE FROM periode WHERE id = ?", (periode_id,))
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def get_ringkasan_tahun(tahun):
    """Ringkasan data penugasan per tahun untuk jendela pembersihan."""
    conn = get_connection()
    try:
        periode_ids = [
            r[0] for r in conn.execute(
                "SELECT id FROM periode WHERE tahun = ?", (tahun,)
            ).fetchall()
        ]
        if not periode_ids:
            return {"periode": 0, "kegiatan": 0, "anggota": 0, "jadwal": 0}
        placeholders = ",".join("?" * len(periode_ids))
        kegiatan = conn.execute(
            f"SELECT COUNT(*) FROM bagian WHERE periode_id IN ({placeholders})",
            periode_ids
        ).fetchone()[0]
        anggota = conn.execute(
            f"SELECT COUNT(*) FROM anggota_bagian WHERE bagian_id IN (SELECT id FROM bagian WHERE periode_id IN ({placeholders}))",
            periode_ids
        ).fetchone()[0]
        jadwal = conn.execute(
            f"SELECT COUNT(*) FROM jadwal WHERE periode_id IN ({placeholders})",
            periode_ids
        ).fetchone()[0]
        return {"periode": len(periode_ids), "kegiatan": kegiatan, "anggota": anggota, "jadwal": jadwal}
    finally:
        conn.close()


def hapus_data_penugasan_tahun(tahun):
    """Menghapus seluruh data penugasan satu tahun, tanpa menghapus master pegawai."""
    conn = get_connection()
    try:
        periode_ids = [
            r[0] for r in conn.execute(
                "SELECT id FROM periode WHERE tahun = ?", (tahun,)
            ).fetchall()
        ]
        if not periode_ids:
            return 0
        placeholders = ",".join("?" * len(periode_ids))
        bagian_ids = [
            r[0] for r in conn.execute(
                f"SELECT id FROM bagian WHERE periode_id IN ({placeholders})",
                periode_ids
            ).fetchall()
        ]
        if bagian_ids:
            ph_b = ",".join("?" * len(bagian_ids))
            conn.execute(f"DELETE FROM jadwal WHERE bagian_id IN ({ph_b})", bagian_ids)
            conn.execute(f"DELETE FROM anggota_bagian WHERE bagian_id IN ({ph_b})", bagian_ids)
        conn.execute(f"DELETE FROM jadwal WHERE periode_id IN ({placeholders})", periode_ids)
        conn.execute(f"DELETE FROM hari_libur WHERE periode_id IN ({placeholders})", periode_ids)
        conn.execute(f"DELETE FROM bagian WHERE periode_id IN ({placeholders})", periode_ids)
        conn.execute(f"DELETE FROM periode WHERE id IN ({placeholders})", periode_ids)
        conn.commit()
        return len(periode_ids)
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def hapus_semua_data_penugasan():
    """Menghapus seluruh histori penugasan semua tahun, master pegawai tetap."""
    conn = get_connection()
    try:
        conn.execute("DELETE FROM jadwal")
        conn.execute("DELETE FROM anggota_bagian")
        conn.execute("DELETE FROM hari_libur")
        conn.execute("DELETE FROM bagian")
        conn.execute("DELETE FROM periode")
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()
# =========================================================
# ALGORITMA KOMBINASI DANA
# =========================================================

def cari_beberapa_kombinasi(pegawai, dana, max_hasil=None):
    """
    Mencari beberapa kombinasi pegawai yang total tarifnya
    tepat sama dengan dana.

    Kombinasi dibedakan berdasarkan jumlah pegawai per golongan/tarif,
    bukan berdasarkan nama pegawai.
    """

    # Kelompokkan pegawai berdasarkan golongan utama + tarif
    kelompok = {}

    for i, p in enumerate(pegawai):
        golongan = p[5] or "Tidak diketahui"
        tarif = p[7]

        key = (golongan, tarif)

        if key not in kelompok:
            kelompok[key] = []

        kelompok[key].append(i)

    # Urutkan golongan supaya hasil konsisten
    kelompok_list = sorted(
        kelompok.items(),
        key=lambda x: (x[0][0], x[0][1])
    )

    hasil_komposisi = []

    def cari(index, sisa, pilihan):
        # Batasi hasil hanya jika pemanggil memang memberi batas.
        if max_hasil is not None and len(hasil_komposisi) >= max_hasil:
            return

        # Dana sudah tepat
        if sisa == 0:
            # Ubah pilihan jumlah menjadi signature
            signature = tuple(
                sorted(
                    (gol, tarif, jumlah)
                    for (gol, tarif), jumlah in pilihan.items()
                    if jumlah > 0
                )
            )

            if signature not in [
                h["signature"] for h in hasil_komposisi
            ]:
                hasil_komposisi.append({
                    "signature": signature,
                    "pilihan": pilihan.copy()
                })

            return

        # Semua kelompok sudah diperiksa
        if index >= len(kelompok_list):
            return

        (golongan, tarif), daftar_index = kelompok_list[index]

        maksimal = min(
            len(daftar_index),
            sisa // tarif
        )

        # Coba jumlah sedikit sampai banyak
        for jumlah in range(maksimal + 1):
            if jumlah > 0:
                pilihan[(golongan, tarif)] = jumlah

            cari(
                index + 1,
                sisa - (jumlah * tarif),
                pilihan
            )

            if jumlah > 0:
                del pilihan[(golongan, tarif)]

            if max_hasil is not None and len(hasil_komposisi) >= max_hasil:
                return

    cari(0, dana, {})

    # Prioritas:
    # 1. jumlah pegawai lebih sedikit
    # 2. jumlah golongan yang digunakan lebih sedikit
    hasil_komposisi.sort(
        key=lambda h: (
            sum(x[2] for x in h["signature"]),
            len(h["signature"])
        )
    )

    # Ubah komposisi menjadi daftar ID pegawai
    hasil = []

    for nomor, komposisi in enumerate(hasil_komposisi, start=1):
        pegawai_terpilih = []

        for (golongan, tarif), jumlah in komposisi["pilihan"].items():
            key = (golongan, tarif)
            daftar_index = kelompok[key]

            pegawai_terpilih.extend(
                daftar_index[:jumlah]
            )

        hasil.append({
            "nomor": nomor,
            "pegawai": pegawai_terpilih,
            "signature": komposisi["signature"]
        })

    return hasil



def cari_komposisi_turun(tarif_groups, dana, max_hasil=60):
    """
    Mencari komposisi TARGET JUMLAH TURUN berdasarkan dana bulanan.
    Satu 'turun' berarti satu pegawai hadir pada satu tanggal.
    Tarif dapat dipakai berulang karena pegawai yang sama dapat turun
    beberapa kali dalam satu bulan.
    """
    dana = int(dana or 0)
    groups = sorted(
        [(str(g), int(t)) for g, t in tarif_groups if int(t or 0) > 0],
        key=lambda x: (x[1], x[0])
    )
    if dana <= 0 or not groups:
        return []

    # Skala berdasarkan FPB agar DP tetap ringan.
    skala = dana
    for _, tarif in groups:
        skala = math.gcd(skala, tarif)
    if skala <= 0:
        skala = 1

    target_u = dana // skala
    vals = [(gol, tarif, tarif // skala) for gol, tarif in groups]

    # dp[total] berisi beberapa signature exact.
    dp = {0: [()]}
    for total_u in range(1, target_u + 1):
        combos = []
        for gol, tarif, unit in vals:
            if total_u < unit or (total_u - unit) not in dp:
                continue
            for prev in dp[total_u - unit]:
                counts = {(g, t): n for g, t, n in prev}
                key = (gol, tarif)
                counts[key] = counts.get(key, 0) + 1
                sig = tuple(sorted(
                    (g, t, n) for (g, t), n in counts.items() if n
                ))
                if sig not in combos:
                    combos.append(sig)
                    if len(combos) >= max_hasil:
                        break
            if len(combos) >= max_hasil:
                break
        if combos:
            combos.sort(key=lambda sig: (
                sum(x[2] for x in sig),
                len(sig),
                tuple((x[0], x[1], x[2]) for x in sig)
            ))
            dp[total_u] = combos

    hasil = []
    for nomor, sig in enumerate(dp.get(target_u, [])[:max_hasil], start=1):
        hasil.append({
            "nomor": nomor,
            "signature": sig,
            "jumlah_turun": sum(x[2] for x in sig),
            "total": sum(x[1] * x[2] for x in sig),
        })
    return hasil


def cari_saran_dana_terdekat(pegawai, target):
    """Mencari jumlah tarif yang paling dekat dengan target, di bawah dan di atas."""
    target = int(target or 0)
    if target < 0:
        return None, None
    if target == 0:
        return 0, 0

    nilai = [int(p[7] or 0) for p in pegawai if int(p[7] or 0) > 0]
    if not nilai:
        return 0, None

    # Gunakan satuan terbesar yang membagi semua nilai agar DP tetap ringan.
    skala = 0
    for v in nilai + [target]:
        skala = math.gcd(skala, v)
    skala = max(skala, 1)

    vals = [v // skala for v in nilai]
    target_u = target // skala

    # Cari sedikit ruang di atas target. Karena setiap nominal positif,
    # cukup diperluas dengan nominal terbesar untuk menemukan kandidat terdekat.
    cap = target_u + max(vals)
    reachable = [False] * (cap + 1)
    reachable[0] = True

    for v in vals:
        for total in range(cap, v - 1, -1):
            if reachable[total - v]:
                reachable[total] = True

    lower_u = None
    for total in range(min(target_u, cap), -1, -1):
        if reachable[total]:
            lower_u = total
            break

    higher_u = None
    for total in range(target_u, cap + 1):
        if reachable[total]:
            higher_u = total
            break

    lower = lower_u * skala if lower_u is not None else None
    higher = higher_u * skala if higher_u is not None else None
    return lower, higher


# =========================================================
# APLIKASI
# =========================================================

class App(tk.Tk):

    def __init__(self):

        super().__init__()

        # Kunci skala internal Tk pada basis 125% (120 DPI).
        # Seluruh widget Tk menggunakan basis ukuran yang sama, tanpa
        # mengubah implementasi kalender atau geometri tiap komponennya.
        try:
            self.tk.call("tk", "scaling", 120.0 / 72.0)
        except Exception:
            pass

        self.configure(bg=UI_BG)

        self.title(
            "Sistem Penugasan Pegawai"
        )
        self.modernize_window(self, "Sistem Penugasan Pegawai", (650, 420))

        # Ukuran awal fleksibel; Windows akan menyesuaikan dengan DPI monitor.
        self.geometry("1200x760")
        self.minsize(650, 420)
        self.state("zoomed")
        self._fullscreen = False
        self.bind("<F11>", self.toggle_fullscreen)
        self.bind("<Escape>", self.exit_fullscreen)

        create_database()
        ensure_bagian_schema()

        self.pegawai = load_pegawai()

        now = datetime.now()

        self.bulan = now.month
        self.tahun = now.year

        self.periode_id = get_or_create_periode(
            self.bulan,
            self.tahun
        )

        kode_awal = load_bagian_codes(self.periode_id)
        self.bagian_aktif = kode_awal[0] if kode_awal else "A"

        self.selected_ids = set()
        self.komposisi_target = {}
        self.komposisi_target_tarif = {}

        self.bagian_id = None

        self.current_theme = "light"
        self.build_style()
        self.build_header()
        self.build_main()

        self.load_bagian_data()
        self.load_kalender()

        # Periksa release terbaru setelah UI selesai ditampilkan.
        self.after(1800, self.check_for_updates)

    def check_for_updates(self, manual=False):
        """Periksa update melalui file manifest GitHub Raw, tanpa GitHub Releases API."""
        def worker():
            try:
                request = urllib.request.Request(
                    UPDATE_MANIFEST_URL,
                    headers={"User-Agent": "SistemPenugasanPegawai-Updater",
                             "Cache-Control": "no-cache"}
                )
                with urllib.request.urlopen(request, timeout=10) as response:
                    manifest = json.loads(response.read().decode("utf-8"))

                latest = str(manifest.get("version", "")).strip().lstrip("vV")
                if not latest or not re.fullmatch(r"\d+(?:\.\d+)*", latest):
                    raise ValueError("Format versi pada update.json tidak valid.")

                current_parts = self._version_parts(APP_VERSION)
                latest_parts = self._version_parts(latest)
                if latest_parts <= current_parts:
                    if manual:
                        self.after(0, lambda: messagebox.showinfo(
                            "Pemeriksaan update",
                            f"Aplikasi sudah menggunakan versi terbaru.\\n\\n"
                            f"Versi saat ini: {APP_VERSION}\\nVersi terbaru: {latest}",
                            parent=self))
                    return

                download_url = str(manifest.get("installer_url", "")).strip()
                release_url = str(manifest.get(
                    "release_url",
                    "https://github.com/briyanyehezkhiel/SistemPenugasan/releases/latest"
                )).strip()

                # Batasi sumber installer ke HTTPS GitHub Releases.
                from urllib.parse import urlparse
                parsed = urlparse(download_url)
                if parsed.scheme != "https" or parsed.netloc.lower() != "github.com" or "/releases/download/" not in parsed.path:
                    raise ValueError("Tautan installer pada update.json tidak valid.")

                self.after(0, lambda: self._offer_update(latest, download_url, release_url))
            except Exception as exc:
                if manual:
                    self.after(0, lambda err=str(exc): messagebox.showwarning(
                        "Pemeriksaan update gagal",
                        "Tidak dapat memeriksa informasi pembaruan. Periksa koneksi internet "
                        "atau coba beberapa saat lagi.\\n\\n" + err,
                        parent=self))

        threading.Thread(target=worker, daemon=True).start()

    @staticmethod
    def _version_parts(version):
        try:
            return tuple(int(part) for part in re.findall(r"\d+", version))
        except Exception:
            return (0,)

    def _offer_update(self, latest, download_url, release_url):
        if not self.winfo_exists():
            return
        answer = messagebox.askyesno(
            "Pembaruan tersedia",
            f"Versi yang digunakan: {APP_VERSION}\\nVersi terbaru: {latest}\\n\\n"
            "Unduh dan jalankan installer terbaru sekarang?\\n"
            "Simpan pekerjaan terlebih dahulu sebelum melanjutkan.",
            parent=self
        )
        if answer:
            self._download_update(download_url, release_url, latest)

    def _download_update(self, download_url, release_url, latest):
        progress = tk.Toplevel(self)
        progress.title("Pembaruan aplikasi")
        progress.transient(self)
        progress.resizable(False, False)
        progress.configure(bg=UI_BG)
        card = tk.Frame(
            progress, bg=UI_PANEL, padx=30, pady=24,
            highlightthickness=1, highlightbackground=UI_BORDER
        )
        card.pack(padx=18, pady=18, fill="both", expand=True)
        tk.Label(
            card, text="Memperbarui aplikasi", bg=UI_PANEL, fg=UI_TEXT,
            font=("Segoe UI", 13, "bold")
        ).pack(pady=(0, 7))
        tk.Label(
            card, text=f"Mengunduh installer versi {latest}...",
            bg=UI_PANEL, fg=UI_MUTED, font=("Segoe UI", 9)
        ).pack(pady=(0, 12))
        bar = ttk.Progressbar(card, mode="indeterminate", length=300)
        bar.pack(pady=6)
        bar.start(10)
        tk.Label(
            card, text="Mohon tunggu dan jangan matikan aplikasi.",
            bg=UI_PANEL, fg=UI_MUTED, font=("Segoe UI", 9)
        ).pack(pady=(10, 0))
        self.center_window(progress, self)

        def worker():
            path = None
            try:
                request = urllib.request.Request(
                    download_url,
                    headers={"User-Agent": "SistemPenugasanPegawai-Updater"}
                )
                with urllib.request.urlopen(request, timeout=60) as response:
                    data = response.read()
                if not data or len(data) < 100_000 or not data.startswith(b"MZ"):
                    raise ValueError("File installer yang diunduh tidak valid.")
                fd, path = tempfile.mkstemp(prefix="SistemPenugasan-Setup-", suffix=".exe")
                with os.fdopen(fd, "wb") as file:
                    file.write(data)
                self.after(0, lambda p=path: self._launch_update(p, progress))
            except Exception as exc:
                if path and os.path.exists(path):
                    try:
                        os.remove(path)
                    except OSError:
                        pass
                self.after(0, lambda err=str(exc): self._update_failed(err, progress, release_url))

        threading.Thread(target=worker, daemon=True).start()

    def _launch_update(self, installer_path, progress):
        try:
            progress.destroy()
            subprocess.Popen([installer_path], close_fds=True)
            self.destroy()
        except Exception as exc:
            messagebox.showerror("Gagal menjalankan installer", str(exc), parent=self)

    def _update_failed(self, error, progress, release_url):
        try:
            progress.destroy()
        except Exception:
            pass
        messagebox.showerror(
            "Pembaruan gagal",
            "Installer tidak berhasil diunduh. Periksa koneksi internet atau unduh manual "
            "dari GitHub Releases.\\n\\n" + error,
            parent=self
        )

    def toggle_fullscreen(self, event=None):
        self._fullscreen = not getattr(self, "_fullscreen", False)
        self.attributes("-fullscreen", self._fullscreen)
        return "break"

    def exit_fullscreen(self, event=None):
        if getattr(self, "_fullscreen", False):
            self._fullscreen = False
            self.attributes("-fullscreen", False)
        return "break"

    # =====================================================
    # MODERN UI HELPERS
    # =====================================================

    def center_window(self, win, parent=None):
        """Pusatkan window secara aman terhadap parent, termasuk pada DPI/scaling tinggi."""
        try:
            if parent is None:
                parent = self
            win.update_idletasks()

            # Ukuran aktual harus sekurang-kurangnya sebesar kebutuhan widget.
            ww = max(int(win.winfo_width()), int(win.winfo_reqwidth()), 1)
            wh = max(int(win.winfo_height()), int(win.winfo_reqheight()), 1)

            sw = int(win.winfo_screenwidth())
            sh = int(win.winfo_screenheight())

            if parent is not None and parent.winfo_exists():
                parent.update_idletasks()
                px = int(parent.winfo_rootx())
                py = int(parent.winfo_rooty())
                pw = max(int(parent.winfo_width()), int(parent.winfo_reqwidth()), 1)
                ph = max(int(parent.winfo_height()), int(parent.winfo_reqheight()), 1)

                x = px + (pw - ww) // 2
                y = py + (ph - wh) // 2
            else:
                x = (sw - ww) // 2
                y = (sh - wh) // 2

            # Jangan biarkan dialog keluar layar.
            x = max(8, min(x, max(8, sw - ww - 8)))
            y = max(8, min(y, max(8, sh - wh - 8)))

            win.geometry(f"{ww}x{wh}+{x}+{y}")
            win.lift()
        except Exception:
            pass

    def fit_popup(self, win, parent=None, min_width=520, min_height=320,
                  max_width=1000, max_height=760):
        """Sesuaikan dialog dengan isi agar tidak terpotong pada DPI/scaling tinggi."""
        try:
            if parent is None:
                parent = self
            win.update_idletasks()

            req_w = max(int(win.winfo_reqwidth()), int(min_width))
            req_h = max(int(win.winfo_reqheight()), int(min_height))

            sw = int(win.winfo_screenwidth())
            sh = int(win.winfo_screenheight())

            # Sisakan ruang untuk taskbar/title bar.
            max_w = min(int(max_width), max(420, sw - 32))
            max_h = min(int(max_height), max(300, sh - 80))

            width = min(req_w, max_w)
            height = min(req_h, max_h)

            # Kalau isi lebih tinggi daripada layar, biarkan dialog resize
            # sampai batas layar; konten yang panjang tetap dapat memakai scroll
            # jika caller menyediakannya.
            win.geometry(f"{width}x{height}")
            win.update_idletasks()
            self.center_window(win, parent)

            # Center sekali lagi setelah Windows menyelesaikan geometry/DPI.
            for delay in (40, 120, 240):
                try:
                    win.after(delay, lambda w=win, p=parent: self.center_window(w, p))
                except Exception:
                    pass
        except Exception:
            self.center_window(win, parent)

    def enable_native_window_controls(self, win, maximize=True, minimize=True):
        """Pastikan Toplevel Windows memiliki tombol minimize/maximize native.

        Tk/Windows kadang menyembunyikan tombol tersebut pada Toplevel yang
        dibuat sebagai transient/dialog. Kita aktifkan kembali style native
        Windows agar user dapat mengecilkan, membesarkan, atau maximize window
        tanpa harus menyeret sudut window secara manual.
        """
        try:
            win.resizable(True, True)
        except Exception:
            pass

        if sys.platform != "win32":
            return

        try:
            import ctypes
            from ctypes import wintypes

            hwnd = wintypes.HWND(win.winfo_id())
            user32 = ctypes.windll.user32
            GWL_STYLE = -16
            WS_MINIMIZEBOX = 0x00020000
            WS_MAXIMIZEBOX = 0x00010000
            WS_THICKFRAME = 0x00040000
            WS_SYSMENU = 0x00080000
            SWP_NOSIZE = 0x0001
            SWP_NOMOVE = 0x0002
            SWP_NOZORDER = 0x0004
            SWP_FRAMECHANGED = 0x0020

            get_style = user32.GetWindowLongW
            set_style = user32.SetWindowLongW
            get_style.argtypes = [wintypes.HWND, ctypes.c_int]
            get_style.restype = ctypes.c_long
            set_style.argtypes = [wintypes.HWND, ctypes.c_int, ctypes.c_long]
            set_style.restype = ctypes.c_long

            # Toplevel yang memakai transient() sering diberi WS_EX_TOOLWINDOW oleh
            # Tk/Windows. Style tersebut membuat title bar hanya menampilkan tombol
            # Close. Untuk aplikasi ini kita ingin window biasa: Minimize +
            # Maximize/Restore + Close tetap tersedia, termasuk pada DPI 175%.
            GWL_EXSTYLE = -20
            WS_EX_TOOLWINDOW = 0x00000080
            WS_EX_APPWINDOW = 0x00040000

            get_exstyle = user32.GetWindowLongW
            set_exstyle = user32.SetWindowLongW
            get_exstyle.argtypes = [wintypes.HWND, ctypes.c_int]
            get_exstyle.restype = ctypes.c_long
            set_exstyle.argtypes = [wintypes.HWND, ctypes.c_int, ctypes.c_long]
            set_exstyle.restype = ctypes.c_long

            style = get_style(hwnd, GWL_STYLE)
            style |= WS_SYSMENU | WS_THICKFRAME
            if maximize:
                style |= WS_MAXIMIZEBOX
            else:
                style &= ~WS_MAXIMIZEBOX
            if minimize:
                style |= WS_MINIMIZEBOX
            else:
                style &= ~WS_MINIMIZEBOX

            exstyle = get_exstyle(hwnd, GWL_EXSTYLE)
            exstyle &= ~WS_EX_TOOLWINDOW
            exstyle |= WS_EX_APPWINDOW

            set_style(hwnd, GWL_STYLE, style)
            set_exstyle(hwnd, GWL_EXSTYLE, exstyle)
            user32.SetWindowPos(
                hwnd, 0, 0, 0, 0, 0,
                SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_FRAMECHANGED
            )
            # Paksa Windows menggambar ulang title bar setelah style berubah.
            try:
                user32.DrawMenuBar(hwnd)
            except Exception:
                pass
        except Exception:
            # Native controls are an enhancement; the application must still
            # work normally if Windows API access is unavailable.
            pass

    def toggle_window_maximize(self, win):
        """Toggle maximize/restore for a Toplevel."""
        try:
            if win.state() == "zoomed":
                win.state("normal")
            else:
                win.state("zoomed")
        except Exception:
            pass

    def modernize_window(self, win, title=None, minsize=(760, 480), parent=None):
        """Gaya dasar jendela modern dan pusatkan pada parent yang benar.

        Jika caller sudah memanggil ``transient(parent)``, parent tersebut dipakai
        otomatis. Ini mencegah popup yang dibuka dari Kalender/kelas data melompat
        kembali ke jendela utama.
        """
        try:
            win.configure(bg=UI_BG)
        except Exception:
            pass
        try:
            win.minsize(*minsize)
        except Exception:
            pass
        try:
            if title:
                win.title(title)
        except Exception:
            pass
        try:
            win.option_add("*Font", ("Segoe UI", 10))
        except Exception:
            pass

        def resolve_parent():
            if parent is not None:
                return parent
            try:
                transient_name = win.transient()
                if transient_name:
                    return win.nametowidget(transient_name)
            except Exception:
                pass
            return self

        def recenter(w=win):
            try:
                if w.winfo_exists():
                    self.center_window(w, resolve_parent())
            except Exception:
                pass

        try:
            win.after_idle(lambda: self.enable_native_window_controls(win))
            win.after(40, lambda: self.enable_native_window_controls(win))
            win.after(120, lambda: self.enable_native_window_controls(win))
            win.after(250, lambda: self.enable_native_window_controls(win))
            win.after(500, lambda: self.enable_native_window_controls(win))
            win.after(80, recenter)
            win.after(160, recenter)
            win.after(320, recenter)
        except Exception:
            pass

    def modernize_tree(self, tree, striped=True):
        """Penyempurnaan kecil Treeview: zebra, hover, selection, dan scroll-wheel."""
        try:
            if striped:
                tree.tag_configure("row_even", background=UI_PANEL)
                tree.tag_configure("row_odd", background=UI_TABLE_ALT)
            tree.tag_configure("selected_row", background=UI_PRIMARY_LIGHT)
            tree.bind("<Motion>", self._tree_hover, add="+")
            tree.bind("<MouseWheel>", lambda e, t=tree: t.yview_scroll(int(-e.delta / 120), "units"), add="+")
            tree.bind("<Shift-MouseWheel>", lambda e, t=tree: t.xview_scroll(int(-e.delta / 120), "units"), add="+")
        except Exception:
            pass

    def _tree_hover(self, event):
        try:
            tree = event.widget
            row = tree.identify_row(event.y)
            old = getattr(self, "_hover_tree_item", None)
            old_tree = getattr(self, "_hover_tree", None)
            if old and old_tree is tree and tree.exists(old):
                tree.item(old, tags=tuple(t for t in tree.item(old, "tags") if t not in ("hover_row",)))
            if row:
                tags = tuple(t for t in tree.item(row, "tags") if t not in ("hover_row",))
                tree.item(row, tags=tags + ("hover_row",))
                tree.tag_configure("hover_row", background=THEME_PALETTES[getattr(self, "current_theme", "light")]["UI_PRIMARY_LIGHT"])
                self._hover_tree_item = row
                self._hover_tree = tree
            else:
                self._hover_tree_item = None
                self._hover_tree = None
        except Exception:
            pass

    # =====================================================
    # STYLE
    # =====================================================

    def build_style(self):
        style = ttk.Style(self)

        try:
            style.theme_use("clam")
        except Exception:
            pass

        # Notebook hanya menjadi pengelola halaman; navigasi visual memakai dropdown.
        try:
            style.layout("Hidden.TNotebook", [("Notebook.client", {"sticky": "nswe"})])
            style.layout("Hidden.TNotebook.Tab", [])
        except Exception:
            pass

        try:
            style.theme_use("clam")
        except Exception:
            pass

        # Halaman utama kembali memakai tab di posisi atas seperti tata letak awal.
        # Sidebar hanya untuk administrasi; Notebook tetap menjadi navigasi utama.
        try:
            style.layout("Sidebar.TNotebook", style.layout("TNotebook"))
            style.layout("Sidebar.TNotebook.Tab", style.layout("TNotebook.Tab"))
        except Exception:
            pass

        # Base typography / colors
        style.configure(".", font=("Segoe UI", 10), foreground=UI_TEXT)
        style.configure("TFrame", background=UI_BG)
        style.configure("TLabel", background=UI_BG, foreground=UI_TEXT, font=("Segoe UI", 10))

        # Modern buttons
        style.configure(
            "TButton",
            font=("Segoe UI", 10, "bold"),
            padding=(16, 9),
            foreground=UI_TEXT,
            background=UI_PANEL,
            bordercolor=UI_BORDER,
            lightcolor=UI_PANEL,
            darkcolor=UI_BORDER,
            relief="flat"
        )
        style.map(
            "TButton",
            background=[("pressed", UI_PRIMARY_LIGHT), ("active", UI_TABLE_HEAD)],
            foreground=[("disabled", "#98A2B3"), ("!disabled", UI_TEXT)]
        )

        style.configure(
            "Primary.TButton",
            foreground=UI_WHITE,
            background=UI_PRIMARY,
            bordercolor=UI_PRIMARY,
            padding=(16, 9)
        )
        style.map(
            "Primary.TButton",
            background=[("pressed", UI_PRIMARY_HOVER), ("active", UI_PRIMARY_HOVER)]
        )

        style.configure(
            "Danger.TButton",
            foreground=UI_DANGER,
            background=UI_DANGER_LIGHT,
            bordercolor="#F4C7C7",
            padding=(12, 8)
        )
        style.map(
            "Danger.TButton",
            background=[("active", UI_DANGER_LIGHT), ("pressed", UI_DANGER_LIGHT)]
        )

        # Inputs
        style.configure(
            "TEntry",
            padding=(10, 7),
            fieldbackground=UI_PANEL,
            foreground=UI_TEXT,
            bordercolor=UI_BORDER,
            lightcolor=UI_PRIMARY,
            darkcolor=UI_BORDER
        )
        style.configure(
            "TCombobox",
            padding=(8, 6),
            fieldbackground=UI_PANEL,
            foreground=UI_TEXT,
            bordercolor=UI_BORDER,
            arrowcolor=UI_MUTED
        )
        style.map(
            "TCombobox",
            fieldbackground=[("readonly", UI_PANEL)],
            selectbackground=[("readonly", UI_PRIMARY_LIGHT)],
            selectforeground=[("readonly", UI_TEXT)]
        )

        style.configure(
            "TCheckbutton",
            background=UI_BG,
            foreground=UI_TEXT,
            font=("Segoe UI", 10),
            padding=(4, 4)
        )

        # Notebook: cleaner tabs
        style.configure(
            "TNotebook",
            background=UI_BG,
            borderwidth=0,
            tabmargins=(0, 0, 0, 0)
        )
        style.configure(
            "TNotebook.Tab",
            background=UI_TABLE_HEAD,
            foreground=UI_MUTED,
            padding=(24, 13),
            font=("Segoe UI", 10, "bold"),
            borderwidth=0
        )
        style.map(
            "TNotebook.Tab",
            background=[("selected", UI_PRIMARY), ("active", UI_TABLE_ALT)],
            foreground=[("selected", "#FFFFFF"), ("active", UI_TEXT)],
            padding=[("selected", (26, 14)), ("!selected", (22, 12))]
        )

        # Menubutton
        style.configure(
            "TMenubutton",
            font=("Segoe UI", 10, "bold"),
            padding=(13, 8),
            foreground=UI_TEXT,
            background=UI_PANEL,
            bordercolor=UI_BORDER
        )

        style.configure(
            "Soft.TButton",
            foreground=UI_PRIMARY,
            background=UI_PRIMARY_LIGHT,
            bordercolor="#CFE0FF",
            padding=(13, 8)
        )
        style.map(
            "Soft.TButton",
            background=[("active", "#DCEAFF"), ("pressed", "#D2E3FF")]
        )

        # Treeview
        style.configure(
            "Treeview",
            rowheight=34,
            font=("Segoe UI", 10),
            background=UI_PANEL,
            fieldbackground=UI_PANEL,
            foreground=UI_TEXT,
            borderwidth=0,
            relief="flat"
        )
        style.map(
            "Treeview",
            background=[("selected", UI_PRIMARY_LIGHT)],
            foreground=[("selected", UI_TEXT)]
        )
        style.configure(
            "Treeview.Heading",
            font=("Segoe UI", 10, "bold"),
            background=UI_TABLE_HEAD,
            foreground=UI_TEXT,
            padding=(10, 8),
            borderwidth=0,
            relief="flat"
        )
        style.map(
            "Treeview.Heading",
            background=[("active", "#E5ECF5")]
        )

        # Scrollbars
        style.configure(
            "Vertical.TScrollbar",
            background="#CBD5E1",
            troughcolor=UI_BG,
            bordercolor=UI_BG,
            arrowcolor=UI_MUTED,
            width=11
        )
        style.configure(
            "Horizontal.TScrollbar",
            background="#CBD5E1",
            troughcolor=UI_BG,
            bordercolor=UI_BG,
            arrowcolor=UI_MUTED,
            width=10
        )

        # Modern utility styles used consistently by dialogs and secondary actions.
        style.configure(
            "Card.TFrame",
            background=UI_PANEL,
            borderwidth=0
        )
        style.configure(
            "Section.TLabel",
            background=UI_PANEL,
            foreground=UI_TEXT,
            font=("Segoe UI Semibold", 11)
        )
        style.configure(
            "Muted.TLabel",
            background=UI_PANEL,
            foreground=UI_MUTED,
            font=("Segoe UI", 9)
        )
        style.configure(
            "Accent.TButton",
            foreground=UI_WHITE,
            background=UI_PRIMARY,
            borderwidth=0,
            padding=(15, 9),
            font=("Segoe UI Semibold", 10)
        )
        style.map(
            "Accent.TButton",
            background=[("active", UI_PRIMARY_HOVER), ("pressed", UI_PRIMARY_HOVER), ("disabled", "#CBD5E1")],
            foreground=[("disabled", "#FFFFFF"), ("!disabled", UI_WHITE)]
        )
        style.configure(
            "Ghost.TButton",
            foreground=UI_TEXT,
            background=UI_PANEL,
            borderwidth=1,
            padding=(12, 8),
            font=("Segoe UI Semibold", 10)
        )
        style.map(
            "Ghost.TButton",
            background=[("active", "#F3F4F6"), ("pressed", "#E5E7EB")],
            bordercolor=[("active", "#D1D5DB")]
        )
    # =====================================================
    # HEADER
    # =====================================================

    def build_header(self):
        # Judul aplikasi sekarang menyatu dengan navbar utama.
        # Tidak lagi membuat header terpisah yang memakan tinggi layar.
        self.header_period = tk.Label(self, text="")
        self.header_period.place_forget()

    # =====================================================
    # MAIN
    # =====================================================

    def _force_canvas_scrollregion(self, canvas):
        """Pastikan area scroll mengikuti seluruh isi, bukan hanya viewport."""
        try:
            canvas.update_idletasks()
            bbox = canvas.bbox("all")
            if not bbox:
                return
            x1, y1, x2, y2 = bbox
            # Sedikit ruang agar scrollbar tetap nyaman digunakan.
            canvas.configure(scrollregion=(x1, y1, max(x2, canvas.winfo_width()), max(y2, canvas.winfo_height())))
        except (tk.TclError, AttributeError):
            pass

    def build_main(self):

        # Kerangka aplikasi dengan navbar horizontal di bagian atas.
        self.main_shell = tk.Frame(self, bg=UI_BG)
        self.main_shell.pack(fill="both", expand=True)

        self.sidebar_expanded_width = 232
        self.sidebar_collapsed_width = 68
        self._sidebar_expanded = True
        self.sidebar = tk.Frame(
            self.main_shell, bg=UI_PANEL, height=58,
            highlightthickness=1, highlightbackground=UI_BORDER, bd=0
        )
        self.sidebar.pack(side="top", fill="x")
        self.sidebar.pack_propagate(False)

        self.workspace = tk.Frame(self.main_shell, bg=UI_BG)
        self.workspace.pack(side="top", fill="both", expand=True)

        # Variabel periode harus tersedia sebelum widget sidebar dibuat.
        self.bulan_var = tk.StringVar(value=NAMA_BULAN[self.bulan])
        self.tahun_var = tk.StringVar(value=str(self.tahun))

        self._build_sidebar()

        # PEMILIH KEGIATAN: dropdown ringkas, tidak memanjang memenuhi layar.
        self.bagian_frame = tk.Frame(
            self.workspace, bg=UI_PANEL, highlightbackground=UI_BORDER,
            highlightthickness=1, padx=12, pady=8
        )
        self.bagian_frame.pack(fill="x", padx=15, pady=(10, 7))

        tk.Label(
            self.bagian_frame, text="KEGIATAN",
            bg=UI_PANEL, fg=UI_MUTED,
            font=("Segoe UI", 8, "bold")
        ).pack(side="left", padx=(2, 10))

        self.bagian_dropdown_var = tk.StringVar(value="Memuat kegiatan...")
        self.bagian_dropdown = ttk.Menubutton(
            self.bagian_frame, textvariable=self.bagian_dropdown_var,
            style="Soft.TButton"
        )
        self.bagian_dropdown.pack(side="left")
        self.bagian_dropdown_menu = tk.Menu(
            self.bagian_dropdown, tearoff=False,
            font=("Segoe UI", 10), activeborderwidth=0,
            background=UI_PANEL, foreground=UI_TEXT,
            activebackground=UI_PRIMARY, activeforeground="#FFFFFF",
            borderwidth=0, relief="flat"
        )
        self.bagian_dropdown.configure(menu=self.bagian_dropdown_menu)

        self.btn_tambah_kegiatan = ttk.Button(
            self.bagian_frame, text="＋ Tambah",
            command=self.tambah_bagian, style="Primary.TButton"
        )
        self.btn_tambah_kegiatan.pack(side="left", padx=(8, 2))

        self.btn_refresh_all = ttk.Button(
            self.bagian_frame, text="↻ Refresh Semua",
            command=self.refresh_all, style="Soft.TButton"
        )
        self.btn_refresh_all.pack(side="right", padx=(8, 2))

        self.bagian_buttons = {}
        self.refresh_bagian_buttons()

        # NOTEBOOK
        self.notebook = ttk.Notebook(
            self.workspace, style="Hidden.TNotebook"
        )

        self.notebook.pack(
            fill="both",
            expand=True,
            padx=16,
            pady=(2, 14)
        )

        # Satu dropdown untuk berpindah antar tiga halaman utama.
        self.view_bar = tk.Frame(
            self.workspace, bg=UI_PANEL, highlightbackground=UI_BORDER,
            highlightthickness=1, padx=12, pady=7
        )
        self.view_bar.pack(fill="x", padx=16, pady=(0, 4), before=self.notebook)
        tk.Label(
            self.view_bar, text="TAMPILAN",
            bg=UI_PANEL, fg=UI_MUTED, font=("Segoe UI", 8, "bold")
        ).pack(side="left", padx=(2, 10))
        self.view_dropdown_var = tk.StringVar(value="Pengaturan Kegiatan")
        self.view_dropdown = ttk.Menubutton(
            self.view_bar, textvariable=self.view_dropdown_var,
            style="Soft.TButton"
        )
        self.view_dropdown.pack(side="left")
        self.view_dropdown_menu = tk.Menu(
            self.view_dropdown, tearoff=False, font=("Segoe UI", 10),
            background=UI_PANEL, foreground=UI_TEXT,
            activebackground=UI_PRIMARY, activeforeground="#FFFFFF",
            borderwidth=0, relief="flat"
        )
        self.view_dropdown.configure(menu=self.view_dropdown_menu)
        self.view_dropdown_menu.add_command(
            label="⚙  Pengaturan Kegiatan",
            command=lambda: self.notebook.select(self.tab_pengaturan)
        )
        self.view_dropdown_menu.add_command(
            label="▦  Kalender Kegiatan",
            command=lambda: self.notebook.select(self.tab_jadwal)
        )
        self.view_dropdown_menu.add_command(
            label="▤  Data Kegiatan",
            command=lambda: self.notebook.select(self.tab_data_kegiatan)
        )

        self.tab_pengaturan = tk.Frame(
            self.notebook,
            bg=UI_BG
        )

        self.tab_jadwal = tk.Frame(
            self.notebook,
            bg=UI_BG
        )

        self.tab_data_kegiatan = tk.Frame(
            self.notebook,
            bg=UI_BG
        )

        self.tab_pegawai = tk.Frame(
            self.notebook,
            bg=UI_BG
        )

        self.notebook.add(
            self.tab_pengaturan,
            text="  Pengaturan Kegiatan  "
        )

        self.notebook.add(
            self.tab_jadwal,
            text="  Kalender Kegiatan  "
        )

        self.notebook.add(
            self.tab_data_kegiatan,
            text="  Data Kegiatan  "
        )

        self.build_pengaturan()
        self.build_jadwal()
        self.build_data_kegiatan()

        # Data Kegiatan selalu disegarkan saat tab dibuka/aktif kembali.
        # Jadi user tidak perlu menekan Refresh Semua hanya untuk melihat
        # perubahan terbaru pada kegiatan yang sedang dipilih.
        self.notebook.bind(
            "<<NotebookTabChanged>>",
            self.on_notebook_tab_changed
        )

    def toggle_theme(self):
        """Ganti tema dan bangun ulang sel kalender agar warna kegiatan ikut berubah."""
        self.current_theme = "dark" if getattr(self, "current_theme", "light") == "light" else "light"
        self.apply_theme()

        # Sel kalender kegiatan memakai warna khusus per kegiatan, bukan hanya
        # warna netral yang dapat dipetakan oleh apply_theme(). Karena itu sel
        # lama harus dibuat ulang setelah tema berubah agar warna latar dan
        # warna teks dihitung dari palet tema yang baru.
        try:
            self.load_kalender()
        except (tk.TclError, AttributeError, IndexError, ValueError):
            pass

        # Jika jendela Kalender Keseluruhan sedang terbuka, perbarui juga
        # warna selnya tanpa menutup atau mengubah posisi jendela.
        try:
            if (
                getattr(self, "overall_calendar_window", None) is not None
                and self.overall_calendar_window.winfo_exists()
            ):
                self.load_kalender_keseluruhan()
        except (tk.TclError, AttributeError, IndexError, ValueError):
            pass

    def apply_theme(self):
        """Terapkan palet ke widget yang sudah ada; warna status/kegiatan dibiarkan."""
        global UI_BG, UI_PANEL, UI_BORDER, UI_TEXT, UI_MUTED, UI_PRIMARY
        global UI_PRIMARY_HOVER, UI_PRIMARY_LIGHT, UI_SUCCESS_LIGHT
        global UI_DANGER_LIGHT, UI_TABLE_HEAD, UI_TABLE_ALT, UI_SHADOW

        theme = getattr(self, "current_theme", "light")
        palette = THEME_PALETTES[theme]
        old_palette = THEME_PALETTES["dark" if theme == "light" else "light"]
        # Simpan nilai palet aktif agar pergantian berikutnya dapat memetakan
        # kembali widget yang telah dibuat dengan warna tema sebelumnya.
        for name, value in palette.items():
            globals()[name] = value

        color_map = {
            old_palette["UI_BG"]: palette["UI_BG"],
            old_palette["UI_PANEL"]: palette["UI_PANEL"],
            old_palette["UI_BORDER"]: palette["UI_BORDER"],
            old_palette["UI_TEXT"]: palette["UI_TEXT"],
            old_palette["UI_MUTED"]: palette["UI_MUTED"],
            old_palette["UI_PRIMARY_LIGHT"]: palette["UI_PRIMARY_LIGHT"],
            old_palette["UI_SUCCESS_LIGHT"]: palette["UI_SUCCESS_LIGHT"],
            old_palette["UI_DANGER_LIGHT"]: palette["UI_DANGER_LIGHT"],
            old_palette["UI_TABLE_HEAD"]: palette["UI_TABLE_HEAD"],
            old_palette["UI_TABLE_ALT"]: palette["UI_TABLE_ALT"],
            old_palette["UI_SHADOW"]: palette["UI_SHADOW"],
            "#FFFFFF": palette["UI_PANEL"], "white": palette["UI_PANEL"],
            "#F8FAFC": palette["UI_TABLE_HEAD"], "#F1F5F9": palette["UI_BORDER"],
            "#E5E7EB": palette["UI_BORDER"], "#CBD5E1": palette["UI_BORDER"],
            "#111827": palette["UI_TEXT"], "#172033": palette["UI_TEXT"],
            "#64748B": palette["UI_MUTED"], "#6B7280": palette["UI_MUTED"],
            "#F5F7FA": palette["UI_BG"], "#F4F6FA": palette["UI_BG"],
            "#F9FAFB": palette["UI_BG"], "#F3F4F6": palette["UI_TABLE_HEAD"],
            "#F1F5F9": palette["UI_TABLE_ALT"], "#E2E8F0": palette["UI_BORDER"],
            "#D1D5DB": palette["UI_BORDER"], "#000000": palette["UI_TEXT"],
            "black": palette["UI_TEXT"],
            "#F1F4F7": palette["UI_TABLE_ALT"],
            "#F4F4F4": palette["UI_TABLE_ALT"],
            "#888888": palette["UI_MUTED"],
            "#F0FDFA": palette["UI_PRIMARY_LIGHT"],
            "#E8EEF6": palette["UI_TABLE_HEAD"],
            "#F1F5F9": palette["UI_TABLE_ALT"],
        }

        # Native Tk menus do not inherit ttk styles, so update them explicitly.
        for menu_name in ("bagian_dropdown_menu", "view_dropdown_menu", "sidebar_menu_admin"):
            menu = getattr(self, menu_name, None)
            if menu is not None:
                try:
                    menu.configure(
                        background=palette["UI_PANEL"],
                        foreground=palette["UI_TEXT"],
                        activebackground=palette["UI_PRIMARY"],
                        activeforeground="#FFFFFF",
                        borderwidth=0, relief="flat"
                    )
                except tk.TclError:
                    pass

        def recolor(widget):
            try:
                options = set(widget.keys())
                for opt in ("bg", "background", "highlightbackground", "troughcolor"):
                    if opt in options:
                        current = widget.cget(opt)
                        if current in color_map:
                            widget.configure(**{opt: color_map[current]})
                for opt in ("fg", "foreground", "insertbackground"):
                    if opt in options:
                        current = widget.cget(opt)
                        if current in color_map:
                            widget.configure(**{opt: color_map[current]})
                        elif isinstance(widget, tk.Label) and current in ("#000000", "black"):
                            widget.configure(**{opt: palette["UI_TEXT"]})
                if "selectbackground" in options:
                    widget.configure(selectbackground=palette["UI_PRIMARY"])
                if isinstance(widget, tk.Canvas):
                    # Canvas kalender memakai warna kegiatan/status khusus; hanya
                    # warna dasar netral yang dipetakan melalui color_map.
                    pass
                if isinstance(widget, ttk.Treeview):
                    # Semua tabel memakai palet yang sama. Tag khusus nonaktif
                    # tetap dibedakan, tetapi tidak lagi memakai latar putih.
                    for tag_name in widget.tag_names():
                        if tag_name in ("baris1", "row_even"):
                            widget.tag_configure(
                                tag_name, background=palette["UI_PANEL"],
                                foreground=palette["UI_TEXT"]
                            )
                        elif tag_name in ("baris2", "row_odd"):
                            widget.tag_configure(
                                tag_name, background=palette["UI_TABLE_ALT"],
                                foreground=palette["UI_TEXT"]
                            )
                        elif tag_name == "nonaktif":
                            widget.tag_configure(
                                tag_name, background=palette["UI_TABLE_HEAD"],
                                foreground=palette["UI_MUTED"]
                            )
                        elif tag_name == "hover_row":
                            widget.tag_configure(
                                tag_name, background=palette["UI_PRIMARY_LIGHT"],
                                foreground=palette["UI_TEXT"]
                            )
                        elif tag_name == "selected_row":
                            widget.tag_configure(
                                tag_name, background=palette["UI_PRIMARY_LIGHT"],
                                foreground=palette["UI_TEXT"]
                            )
                    widget.configure(
                        background=palette["UI_PANEL"],
                        foreground=palette["UI_TEXT"]
                    )
            except (tk.TclError, TypeError):
                pass
            try:
                for child in widget.winfo_children():
                    recolor(child)
            except (tk.TclError, AttributeError):
                pass

        # Tk root dan seluruh halaman utama. Toplevel yang sudah terbuka juga
        # merupakan anak root sehingga ikut diperbarui.
        recolor(self)
        try:
            self.configure(bg=palette["UI_BG"])
            self.sidebar.configure(bg=UI_HEADER)
            self.sidebar_brand.configure(bg=UI_HEADER, fg=UI_WHITE)
            self.header_period.configure(bg=UI_HEADER, fg="#D1D5DB")
            self.sidebar_version.configure(bg=UI_HEADER, fg="#D1D5DB")
            self.sidebar_period_card.configure(bg=UI_HEADER)
            self.sidebar_section_admin.configure(
                bg=UI_HEADER, fg=UI_WHITE,
                activebackground=UI_HEADER_2, activeforeground=UI_WHITE
            )
            self.theme_button.configure(
                text=("Mode Terang" if theme == "dark" else "Mode Gelap"),
                bg=UI_HEADER, fg=UI_WHITE,
                activebackground=UI_HEADER_2, activeforeground=UI_WHITE
            )
            self.workspace.configure(bg=palette["UI_BG"])
        except (tk.TclError, AttributeError):
            pass

        # Perbarui style ttk untuk combobox, entry, tombol, notebook dan treeview.
        try:
            self.build_style()
            style = ttk.Style(self)
            style.configure("TEntry", fieldbackground=palette["UI_PANEL"], foreground=palette["UI_TEXT"])
            style.configure("TCombobox", fieldbackground=palette["UI_PANEL"], foreground=palette["UI_TEXT"], background=palette["UI_PANEL"])
            style.configure("Treeview", background=palette["UI_PANEL"], fieldbackground=palette["UI_PANEL"], foreground=palette["UI_TEXT"])
            style.configure("Treeview.Heading", background=palette["UI_TABLE_HEAD"], foreground=palette["UI_TEXT"])
            style.map(
                "Treeview",
                background=[("selected", palette["UI_PRIMARY_LIGHT"])],
                foreground=[("selected", palette["UI_TEXT"])]
            )
            style.configure(
                "TScrollbar",
                background=palette["UI_TABLE_HEAD"],
                troughcolor=palette["UI_BG"],
                bordercolor=palette["UI_BG"],
                arrowcolor=palette["UI_MUTED"]
            )
            style.map(
                "TScrollbar",
                background=[("active", palette["UI_BORDER"])]
            )
            style.configure(
                "TCheckbutton", background=palette["UI_BG"],
                foreground=palette["UI_TEXT"]
            )
            style.map(
                "TCheckbutton",
                background=[("active", palette["UI_BG"])],
                foreground=[("disabled", palette["UI_MUTED"])]
            )
        except tk.TclError:
            pass

        # Re-apply the active palette to every Treeview after ttk styles are
        # rebuilt. Explicit row tags can otherwise retain their old light-theme
        # backgrounds while the surrounding interface has already gone dark.
        def _apply_tree_palette(widget):
            try:
                if isinstance(widget, ttk.Treeview):
                    widget.configure(
                        style="Treeview",
                        background=palette["UI_PANEL"],
                        fieldbackground=palette["UI_PANEL"],
                        foreground=palette["UI_TEXT"]
                    )
                    for tag_name in widget.tag_names():
                        if tag_name in ("baris1", "row_even"):
                            widget.tag_configure(
                                tag_name,
                                background=palette["UI_PANEL"],
                                foreground=palette["UI_TEXT"]
                            )
                        elif tag_name in ("baris2", "row_odd"):
                            widget.tag_configure(
                                tag_name,
                                background=palette["UI_TABLE_ALT"],
                                foreground=palette["UI_TEXT"]
                            )
                        elif tag_name in ("nonaktif",):
                            widget.tag_configure(
                                tag_name,
                                background=palette["UI_TABLE_HEAD"],
                                foreground=palette["UI_MUTED"]
                            )
                        elif tag_name in ("selected_row", "hover_row"):
                            widget.tag_configure(
                                tag_name,
                                background=palette["UI_PRIMARY_LIGHT"],
                                foreground=palette["UI_TEXT"]
                            )
            except (tk.TclError, TypeError):
                pass
            try:
                for child in widget.winfo_children():
                    _apply_tree_palette(child)
            except (tk.TclError, AttributeError):
                pass

        _apply_tree_palette(self)

    def _build_sidebar(self):
        """Navbar tunggal; menu administrasi dropdown di kanan."""
        self.sidebar.configure(bg=UI_HEADER, height=62, bd=0, highlightthickness=0)
        self.sidebar.pack_propagate(False)
        self.sidebar_brand = tk.Label(
            self.sidebar, text="Sistem Penugasan Pegawai",
            bg=UI_HEADER, fg=UI_WHITE, font=("Segoe UI Semibold", 15)
        )
        self.sidebar_brand.pack(side="left", padx=(22, 14))
        self.header_period = tk.Label(
            self.sidebar, text="", bg=UI_HEADER, fg="#D1D5DB",
            font=("Segoe UI", 9)
        )
        self.header_period.pack(side="left", padx=(0, 12))
        self.sidebar_version = tk.Label(
            self.sidebar, text=f"v{APP_VERSION}", bg=UI_HEADER,
            fg="#D1D5DB", font=("Segoe UI", 8)
        )
        self.sidebar_version.pack(side="right", padx=(4, 14))

        self.sidebar_trigger = tk.Menubutton(
            self.sidebar, text="Administrasi  ▾", bg=UI_HEADER, fg=UI_WHITE,
            activebackground=UI_HEADER_2, activeforeground=UI_WHITE,
            relief="flat", bd=0, padx=14, pady=10, cursor="hand2",
            font=("Segoe UI", 10, "bold")
        )
        self.sidebar_trigger.pack(side="right", padx=5)
        self.sidebar_section_admin = self.sidebar_trigger
        self.sidebar_menu_admin = tk.Menu(
            self.sidebar_trigger, tearoff=False,
            font=("Segoe UI", 10), bg=UI_PANEL, fg=UI_TEXT,
            activebackground=UI_PRIMARY, activeforeground="#FFFFFF",
            bd=0, relief="flat"
        )
        for label, command in [
            ("Data Pegawai", self.buka_data_pegawai),
            ("Data Lokasi", self.buka_data_desa),
            ("Kalender Keseluruhan", self.buka_kalender_keseluruhan),
            ("Kelola / Bersihkan Data", self.buka_kelola_data),
            ("Export Excel", self.export_penugasan_excel),
            ("Periksa Update", lambda: self.check_for_updates(manual=True)),
        ]:
            self.sidebar_menu_admin.add_command(label=label, command=command)
        self.sidebar_trigger.configure(menu=self.sidebar_menu_admin)

        self.theme_button = tk.Button(
            self.sidebar, text="Mode Terang" if self.current_theme == "dark" else "Mode Gelap",
            command=self.toggle_theme, bg=UI_HEADER, fg=UI_WHITE,
            activebackground=UI_HEADER_2, activeforeground=UI_WHITE,
            relief="flat", bd=0, cursor="hand2", font=("Segoe UI", 9),
            padx=12, pady=8
        )
        self.theme_button.pack(side="right", padx=5)

        self.sidebar_period_card = tk.Frame(self.sidebar, bg=UI_HEADER)
        self.sidebar_period_card.pack(side="right", padx=10)
        self.combo_bulan = ttk.Combobox(
            self.sidebar_period_card, textvariable=self.bulan_var,
            values=NAMA_BULAN[1:], state="readonly", width=11
        )
        self.combo_bulan.pack(side="left", padx=(0, 5))
        self.combo_bulan.bind("<<ComboboxSelected>>", self.change_period)
        ttk.Entry(self.sidebar_period_card, textvariable=self.tahun_var, width=6).pack(
            side="left", padx=(0, 5)
        )
        ttk.Button(self.sidebar_period_card, text="Terapkan",
                   style="Primary.TButton", command=self.change_period).pack(side="left")
        ttk.Button(self.sidebar_period_card, text="Hari Libur",
                   style="Soft.TButton", command=self.atur_hari_libur).pack(side="left", padx=5)

        self.sidebar_separator = tk.Frame(self.sidebar, bg=UI_BORDER, width=1)
        self.sidebar_nav_buttons = {}
        self._sidebar_expanded = False

    def _show_top_menu(self, event=None):
        """Menu administrasi memakai dropdown native Tk yang menutup otomatis."""
        self.sidebar_trigger.event_generate("<Button-1>")

    def _hide_top_menu(self):
        return

    def _toggle_top_menu(self):
        self._show_top_menu()

    def _run_top_menu_command(self, command):
        command()

    def _sidebar_track_pointer(self, event=None):
        """Kompatibilitas untuk binding lama; menu atas dikelola event hover sendiri."""
        return

    def _set_sidebar_expanded(self, expanded):
        self._sidebar_expanded = False

    def _set_sidebar_active(self, active_key):
        for key, btn in getattr(self, "sidebar_nav_buttons", {}).items():
            selected = key == active_key
            btn.configure(
                bg=UI_PRIMARY if selected else UI_HEADER,
                fg=UI_WHITE if selected else "#CBD5E1",
                activebackground=UI_PRIMARY_HOVER if selected else UI_HEADER_2
            )

    def _select_main_page(self, page):
        """Pilih halaman Notebook lama melalui navigasi sidebar."""
        try:
            self.notebook.select(page)
            if page == self.tab_pengaturan:
                self._set_sidebar_active("pengaturan")
            elif page == self.tab_jadwal:
                self._set_sidebar_active("jadwal")
            elif page == self.tab_data_kegiatan:
                self._set_sidebar_active("data_kegiatan")
        except (tk.TclError, AttributeError):
            pass

    def on_notebook_tab_changed(self, event=None):
        """
        Navigasi halaman tidak memaksa refresh berat.
        Data diperbarui oleh tombol Refresh atau setelah operasi yang mengubah data.
        Ini mencegah halaman tertutup overlay loading saat baru dibuka.
        """
        try:
            current = self.notebook.select()
            if current == str(self.tab_pengaturan):
                self.view_dropdown_var.set("Pengaturan Kegiatan")
            elif current == str(self.tab_jadwal):
                self.view_dropdown_var.set("Kalender Kegiatan")
            elif current == str(self.tab_data_kegiatan):
                self.view_dropdown_var.set("Data Kegiatan")
        except (tk.TclError, AttributeError):
            pass

    def _finish_tab_loading(self, action):
        try:
            action()
        except Exception as exc:
            messagebox.showerror("Gagal Memuat", "Halaman tidak dapat dimuat. Silakan coba lagi.", parent=self)
        finally:
            self.hide_loading()

    def refresh_bagian_buttons(self):
        """Perbarui dropdown kegiatan agar daftar panjang tetap ringkas."""
        conn = get_connection()
        try:
            rows = conn.execute(
                "SELECT kode, COALESCE(nama_bagian, '') FROM bagian WHERE periode_id = ? ORDER BY id",
                (self.periode_id,)
            ).fetchall()
        finally:
            conn.close()

        self.bagian_buttons = {}
        menu = getattr(self, "bagian_dropdown_menu", None)
        if menu is None:
            return

        menu.delete(0, "end")
        active_label = None
        for kode, nama_bagian in rows:
            title = (nama_bagian or "").strip() or f"Kegiatan {kode}"
            label = f"{kode}  ·  {title}"
            menu.add_command(
                label=("✓  " if kode == self.bagian_aktif else "     ") + label,
                command=lambda k=kode: self.change_bagian(k)
            )
            if kode == self.bagian_aktif:
                active_label = label

        if not rows:
            menu.add_command(label="Belum ada kegiatan", state="disabled")
        if active_label is None and rows:
            active_label = f"{rows[0][0]}  ·  {(rows[0][1] or '').strip() or 'Kegiatan ' + rows[0][0]}"
        self.bagian_dropdown_var.set(active_label or "Pilih Kegiatan")

    def update_bagian_button_state(self):
        """Refresh indikator kegiatan aktif pada dropdown."""
        try:
            self.refresh_bagian_buttons()
        except (tk.TclError, AttributeError):
            pass

    def rename_bagian(self, kode):
        """Mengganti nama tampilan kegiatan dengan dialog singkat."""
        data = load_bagian(self.periode_id, kode)
        if not data:
            return

        bagian_id = data[0]
        conn = get_connection()
        try:
            row = conn.execute(
                "SELECT COALESCE(nama_bagian, '') FROM bagian WHERE id = ?",
                (bagian_id,)
            ).fetchone()
        finally:
            conn.close()

        nama_lama = (row[0] if row else "").strip() or f"Kegiatan {kode}"

        dark = getattr(self, "current_theme", "light") == "dark"
        bg = "#0F172A" if dark else "#F4F7FB"
        panel = "#182438" if dark else "#FFFFFF"
        fg = "#F1F5F9" if dark else "#172033"
        border = "#334155" if dark else "#DCE4EF"
        accent = "#14B8A6" if dark else "#0F766E"

        dialog = tk.Toplevel(self)
        dialog.title("Ubah Nama Kegiatan")
        dialog.transient(self)
        dialog.configure(bg=bg)
        dialog.resizable(False, False)
        dialog.withdraw()

        shell = tk.Frame(
            dialog, bg=panel, padx=16, pady=14,
            highlightthickness=1, highlightbackground=border
        )
        shell.pack(fill="both", expand=True, padx=8, pady=8)

        tk.Label(
            shell, text=f"Nama kegiatan {kode}", bg=panel, fg=fg,
            font=("Segoe UI", 12, "bold")
        ).pack(anchor="w", pady=(0, 10))

        nama_var = tk.StringVar(value=nama_lama)
        entry = tk.Entry(
            shell, textvariable=nama_var, bg=bg, fg=fg,
            insertbackground=fg, relief="flat", bd=0,
            highlightthickness=1, highlightbackground=border,
            highlightcolor=accent, font=("Segoe UI", 10)
        )
        entry.pack(fill="x", ipady=7)

        error_label = tk.Label(
            shell, text="", bg=panel, fg="#DC6464",
            font=("Segoe UI", 8)
        )
        error_label.pack(anchor="w", pady=(3, 0))

        actions = tk.Frame(shell, bg=panel)
        actions.pack(fill="x", pady=(12, 0))
        hasil = {"nama": None}

        def tutup():
            try:
                dialog.grab_release()
            except tk.TclError:
                pass
            dialog.destroy()

        def simpan():
            nama = nama_var.get().strip()
            if not nama:
                error_label.config(text="Nama tidak boleh kosong.")
                entry.focus_set()
                return
            hasil["nama"] = nama
            tutup()

        tk.Button(
            actions, text="Batal", command=tutup,
            bg=bg, fg=fg, activebackground=border, activeforeground=fg,
            relief="flat", bd=0, padx=13, pady=7,
            font=("Segoe UI", 9), cursor="hand2"
        ).pack(side="right")
        tk.Button(
            actions, text="Simpan", command=simpan,
            bg=accent, fg="#FFFFFF", activebackground=accent,
            activeforeground="#FFFFFF", relief="flat", bd=0,
            padx=14, pady=7, font=("Segoe UI", 9, "bold"),
            cursor="hand2"
        ).pack(side="right", padx=(0, 6))

        dialog.bind("<Return>", lambda event: simpan())
        dialog.bind("<Escape>", lambda event: tutup())
        dialog.update_idletasks()
        dialog.geometry("360x150")
        self.center_window(dialog, self)
        dialog.deiconify()
        dialog.grab_set()
        entry.focus_set()
        entry.selection_range(0, "end")
        self.wait_window(dialog)

        nama_baru = hasil["nama"]
        if nama_baru is None:
            return

        conn = None
        try:
            conn = get_connection()
            conn.execute(
                "UPDATE bagian SET nama_bagian = ? WHERE id = ?",
                (nama_baru, bagian_id)
            )
            conn.commit()
        except Exception:
            if conn:
                conn.rollback()
            messagebox.showerror(
                "Gagal Mengganti Nama",
                "Nama kegiatan tidak dapat diubah. Silakan coba lagi.",
                parent=self
            )
            return
        finally:
            if conn:
                conn.close()

        self.refresh_bagian_buttons()
        self.load_bagian_data()

    def tambah_bagian(self):
        kode_list = load_bagian_codes(self.periode_id)
        kode_baru = kode_bagian_berikutnya(kode_list)

        conn = None
        try:
            conn = get_connection()
            cursor = conn.cursor()
            cursor.execute("""
                INSERT INTO bagian
                (periode_id, kode, nama_bagian, kegiatan, lokasi, dana, penanggung_jawab_id)
                VALUES (?, ?, ?, '', '', 0, NULL)
            """, (self.periode_id, kode_baru, f"Kegiatan {kode_baru}"))
            conn.commit()
        except sqlite3.IntegrityError:
            if conn:
                conn.rollback()
            # Ambil ulang kode karena data mungkin berubah saat proses berlangsung.
            kode_list = load_bagian_codes(self.periode_id)
            kode_baru = kode_bagian_berikutnya(kode_list)
            try:
                cursor.execute("""
                    INSERT INTO bagian
                    (periode_id, kode, nama_bagian, kegiatan, lokasi, dana, penanggung_jawab_id)
                    VALUES (?, ?, ?, '', '', 0, NULL)
                """, (self.periode_id, kode_baru, f"Kegiatan {kode_baru}"))
                conn.commit()
            except Exception:
                conn.rollback()
                messagebox.showerror("Gagal", "Kegiatan baru gagal ditambahkan. Silakan coba lagi.")
                return
        except sqlite3.OperationalError as e:
            if conn:
                conn.rollback()
            if "locked" in str(e).lower():
                messagebox.showerror(
                    "Database Sedang Dipakai",
                    "Database sedang terkunci. Tutup program/Excel lain yang sedang membuka database, lalu coba lagi."
                )
            else:
                messagebox.showerror("Gagal", "Kegiatan baru gagal ditambahkan. Silakan coba lagi.")
            return
        finally:
            if conn:
                conn.close()

        self.bagian_aktif = kode_baru

        # Perbarui layar setelah event loop kembali berjalan.
        # Hindari overlay modal: operasi render lama masih sinkron dan akan
        # membuat animasi Tkinter ikut berhenti selama proses tersebut.
        def selesai_tambah_kegiatan():
            try:
                self.refresh_bagian_buttons()
                self.load_bagian_data()
                self.load_kalender()
                self.refresh_data_kegiatan()

                if (
                    getattr(self, "overall_calendar_window", None) is not None
                    and self.overall_calendar_window.winfo_exists()
                ):
                    self.load_kalender_keseluruhan()

                self.notebook.select(self.tab_pengaturan)
            except Exception as exc:
                messagebox.showerror(
                    "Pembaruan Tampilan",
                    "Kegiatan berhasil ditambahkan, tetapi sebagian tampilan "
                    f"belum dapat diperbarui.\\n\\nDetail: {exc}\\n\\n"
                    "Gunakan Refresh Semua untuk mencoba kembali.",
                    parent=self
                )
                return

            messagebox.showinfo(
                "Kegiatan Ditambahkan",
                f"Kegiatan {kode_baru} berhasil ditambahkan. Silakan isi nama dan data kegiatannya.",
                parent=self
            )

        self.after(140, selesai_tambah_kegiatan)

    # =====================================================
    # Refresh ALL
    # =====================================================
    def show_loading(self, text="Sedang memuat data...", parent=None):
        """Indikator proses non-blocking; tidak menutupi atau mengunci area kerja."""
        self.hide_loading()
        try:
            owner = parent if parent is not None and parent.winfo_exists() else self
            palette = {
                "bg": UI_PANEL, "fg": UI_TEXT,
                "muted": UI_MUTED, "accent": UI_PRIMARY,
                "border": UI_BORDER
            }
            chip = tk.Frame(
                owner, bg=palette["bg"], padx=14, pady=9,
                highlightthickness=1, highlightbackground=palette["border"],
                bd=0
            )
            # Hanya chip kecil di pojok kanan atas. Area kerja tetap bisa dilihat
            # dan tidak ada scrim yang menghalangi klik pengguna.
            chip.place(relx=1.0, x=-18, y=14, anchor="ne")
            chip.lift()
            try:
                owner.update_idletasks()
                owner.update()
            except tk.TclError:
                pass

            dot = tk.Canvas(
                chip, width=18, height=18, bg=palette["bg"],
                highlightthickness=0, bd=0
            )
            dot.pack(side="left", padx=(0, 8))
            dot.create_oval(4, 4, 14, 14, outline=palette["border"], width=2)
            dot.create_arc(
                4, 4, 14, 14, start=25, extent=105, style="arc",
                outline=palette["accent"], width=2, tags=("arc",)
            )
            label = tk.Label(
                chip, text=text, bg=palette["bg"], fg=palette["fg"],
                font=("Segoe UI", 9, "bold"), padx=0, pady=0
            )
            label.pack(side="left")
            self._loading_bar = ttk.Progressbar(
                chip, mode="indeterminate", length=54,
                style="Horizontal.TProgressbar"
            )
            self._loading_bar.pack(side="left", padx=(10, 0))
            self._loading_bar.start(12)
            self._loading_win = chip
            self._loading_owner = owner
            self._loading_label = label
            self._loading_spinner = dot
            self._loading_spinner_index = 0

            def animate():
                try:
                    if not chip.winfo_exists():
                        return
                    idx = self._loading_spinner_index
                    dot.itemconfigure("arc", start=(25 + idx * 14) % 360)
                    self._loading_spinner_index = idx + 1
                    self._loading_spinner_job = self.after(55, animate)
                except tk.TclError:
                    self._loading_spinner_job = None

            self._loading_spinner_job = self.after(55, animate)
            owner.update_idletasks()
        except Exception:
            self._loading_win = None
            self._loading_owner = None
            self._loading_label = None
            self._loading_spinner = None
            self._loading_spinner_job = None

    def _animate_loading_spinner(self):
        """Kompatibilitas untuk pemanggil animasi lama."""
        try:
            spinner = getattr(self, "_loading_spinner", None)
            if spinner is None or not spinner.winfo_exists():
                return
            idx = getattr(self, "_loading_spinner_index", 0)
            spinner.itemconfigure("arc", start=(25 + idx * 14) % 360)
            self._loading_spinner_index = idx + 1
            self._loading_spinner_job = self.after(55, self._animate_loading_spinner)
        except (tk.TclError, AttributeError):
            self._loading_spinner_job = None

    def hide_loading(self):
        """Hentikan animasi dan hilangkan indikator proses, aman dipanggil berulang."""
        job = getattr(self, "_loading_spinner_job", None)
        if job is not None:
            try:
                self.after_cancel(job)
            except Exception:
                pass
        self._loading_spinner_job = None

        for attr in ("_loading_bar",):
            widget = getattr(self, attr, None)
            if widget is not None:
                try:
                    widget.stop()
                except Exception:
                    pass

        widget = getattr(self, "_loading_win", None)
        if widget is not None:
            try:
                if widget.winfo_exists():
                    widget.destroy()
            except Exception:
                pass

        self._loading_win = None
        self._loading_owner = None
        self._loading_bar = None
        self._loading_label = None
        self._loading_spinner = None

    def refresh_after_schedule(self, text="Menyegarkan kalender dan data...", parent=None):
        """Refresh bertahap setelah penambahan/penghapusan jadwal."""
        if getattr(self, "_schedule_refresh_in_progress", False):
            return

        self._schedule_refresh_in_progress = True
        overall_parent = (
            parent is not None
            and getattr(self, "overall_calendar_window", None) is not None
            and parent == self.overall_calendar_window
            and self.overall_calendar_window.winfo_exists()
        )
        if overall_parent:
            # Gunakan overlay di dalam Kalender Keseluruhan. Toplevel loading
            # terpisah dapat tertutup di belakang window kalender pada Windows/DPI.
            self._show_overall_loading_overlay(
                self.overall_calendar_window,
                text
            )
            try:
                self.overall_calendar_window.lift()
                self.overall_calendar_window.focus_force()
                self.overall_calendar_window.update_idletasks()
                self.overall_calendar_window.update()
            except Exception:
                pass
        else:
            self.show_loading(text, parent=parent or self)
        # Pastikan indikator sempat tampil dan animasinya berjalan sebelum
        # proses refresh dimulai, terutama jika refresh dipicu dari
        # Kalender Keseluruhan (window Toplevel).
        try:
            self.update_idletasks()
            self.update()
        except Exception:
            pass

        steps = [
            ("Memperbarui kalender kegiatan...", self.load_kalender),
            ("Memperbarui data pegawai...", self.refresh_data_pegawai),
            ("Memperbarui data kegiatan...", self.refresh_data_kegiatan),
            ("Memperbarui kalender keseluruhan...", self.load_kalender_keseluruhan),
        ]

        def jalankan(index=0):
            if index >= len(steps):
                self._schedule_refresh_in_progress = False
                if overall_parent:
                    self._hide_overall_loading_overlay()
                    try:
                        self.overall_calendar_window.lift()
                        self.overall_calendar_window.focus_force()
                    except Exception:
                        pass
                else:
                    self.hide_loading()
                return

            label, fungsi = steps[index]
            try:
                if overall_parent:
                    if getattr(self, "_overall_loading_label", None) is not None:
                        self._overall_loading_label.config(text=label)
                    self.overall_calendar_window.lift()
                    self.overall_calendar_window.update_idletasks()
                    self.overall_calendar_window.update()
                else:
                    if getattr(self, "_loading_label", None) is not None:
                        self._loading_label.config(text=label)
                    self.update()
                fungsi()
                if overall_parent:
                    self.overall_calendar_window.update_idletasks()
                    self.overall_calendar_window.update()
                else:
                    self.update()
            except Exception as e:
                self._schedule_refresh_in_progress = False
                if overall_parent:
                    self._hide_overall_loading_overlay()
                else:
                    self.hide_loading()
                messagebox.showerror(
                    "Refresh Gagal",
                    "Perubahan tersimpan, tetapi tampilan belum dapat diperbarui. Silakan coba Refresh.",
                    parent=parent or self
                )
                return

            self.after(70, lambda: jalankan(index + 1))

        # Beri jeda awal agar spinner benar-benar sempat menggambar frame
        # pertamanya di depan Kalender Keseluruhan sebelum refresh berat dimulai.
        self.after(220, jalankan)

    def refresh_all(self):
        """
        Refresh dilakukan bertahap supaya event loop Tkinter tetap hidup.
        Dengan demikian spinner benar-benar bergerak selama proses refresh,
        bukan hanya menampilkan jendela 'Memuat' yang diam/kosong.
        """
        if getattr(self, "_refresh_in_progress", False):
            return

        self._refresh_in_progress = True
        self.show_loading("Menyegarkan data dan kalender...")

        steps = [
            ("Memuat data pegawai...", self._refresh_step_pegawai),
            ("Memuat periode dan kegiatan...", self._refresh_step_periode),
            ("Memuat kalender kegiatan...", self._refresh_step_kalender),
            ("Memperbarui data pegawai...", self._refresh_step_data_pegawai),
            ("Memperbarui daftar pegawai dan PJ...", self._refresh_step_selection),
            ("Memperbarui data kegiatan...", self._refresh_step_data_kegiatan),
            ("Memperbarui kalender keseluruhan...", self._refresh_step_kalender_keseluruhan),
            ("Memperbarui tampilan...", self._refresh_step_header),
        ]

        def jalankan(index=0):
            if index >= len(steps):
                self._refresh_in_progress = False
                self.hide_loading()
                return

            label, fungsi = steps[index]
            try:
                if getattr(self, "_loading_label", None) is not None:
                    self._loading_label.config(text=label)
                # Beri kesempatan Tkinter menggambar spinner sebelum setiap
                # tahap dimulai dan setelah tahap selesai.
                self.update()
                fungsi()
                self.update()
            except Exception as e:
                self._refresh_in_progress = False
                self.hide_loading()
                messagebox.showerror(
                    "Refresh Gagal",
                    "Refresh belum dapat diselesaikan. Silakan coba lagi.",
                    parent=self
                )
                return

            # Jeda kecil ini sangat penting: event loop mendapat kesempatan
            # menjalankan animasi spinner di antara tahap-tahap refresh.
            self.after(35, lambda: jalankan(index + 1))

        self.after(180, jalankan)

    def _refresh_step_pegawai(self):
        self.pegawai = load_pegawai()

    def _refresh_step_periode(self):
        bulan = NAMA_BULAN.index(self.bulan_var.get())
        tahun = int(self.tahun_var.get())

        self.bulan = bulan
        self.tahun = tahun
        self.periode_id = get_or_create_periode(bulan, tahun)

        kode_list = load_bagian_codes(self.periode_id)
        if self.bagian_aktif not in kode_list:
            self.bagian_aktif = kode_list[0]
        self.refresh_bagian_buttons()
        self.load_bagian_data()

    def _refresh_step_kalender(self):
        self.load_kalender()

    def _refresh_step_data_pegawai(self):
        if hasattr(self, "refresh_data_pegawai"):
            self.refresh_data_pegawai()

    def _refresh_step_selection(self):
        self.refresh_pj_options()
        self.refresh_selection_table()

    def _refresh_step_data_kegiatan(self):
        self.refresh_data_kegiatan()

    def _refresh_step_kalender_keseluruhan(self):
        self.load_kalender_keseluruhan()

    def _refresh_step_header(self):
        nama_header = (
            self.nama_bagian_var.get().strip()
            if hasattr(self, "nama_bagian_var")
            else f"Kegiatan {self.bagian_aktif}"
        )
        self.header_period.config(
            text=(
                f"{NAMA_BULAN[self.bulan]} "
                f"{self.tahun}"
                f"  •  {nama_header}"
            )
        )

    # =====================================================
    # PENGATURAN
    # =====================================================

    def build_pengaturan(self):

        # Nama kegiatan hanya disimpan sebagai state internal.
        # Tidak ditampilkan sebagai field terpisah; rename dilakukan
        # dengan double-click pada tombol kegiatan di bagian atas.
        self.nama_bagian_var = tk.StringVar(value="")

        # =====================================================
        # CANVAS SCROLLABLE
        # =====================================================

        container = tk.Frame(
            self.tab_pengaturan,
            bg=UI_BG
        )
        container.pack(
            fill="both",
            expand=True
        )

        canvas = tk.Canvas(
            container,
            bg=UI_BG,
            highlightthickness=0
        )
        canvas.grid(row=0, column=0, sticky="nsew")
        scrollbar = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        hscroll = ttk.Scrollbar(container, orient="horizontal", command=canvas.xview)
        hscroll.grid(row=1, column=0, sticky="ew")
        canvas.configure(yscrollcommand=scrollbar.set, xscrollcommand=hscroll.set)
        container.grid_rowconfigure(0, weight=1)
        container.grid_columnconfigure(0, weight=1)

        # Frame isi yang bisa discroll
        scroll_frame = tk.Frame(
            canvas,
            bg=UI_BG
        )

        canvas_window = canvas.create_window(
            (0, 0),
            window=scroll_frame,
            anchor="nw"
        )

        def update_scroll_region(event=None):
            canvas.configure(
                scrollregion=canvas.bbox("all")
            )

        def resize_scroll_frame(event):
            try:
                req_w = scroll_frame.winfo_reqwidth()
                canvas.itemconfig(canvas_window, width=max(event.width, req_w))
            except Exception:
                pass

        scroll_frame.bind("<Configure>", update_scroll_region)
        canvas.bind("<Configure>", resize_scroll_frame)

        # Mouse wheel
        def scroll_mousewheel(event):
            canvas.yview_scroll(
                int(-1 * (event.delta / 120)),
                "units"
            )

        canvas.bind_all(
            "<MouseWheel>",
            scroll_mousewheel
        )

        # =====================================================
        # SATU PANEL PENGATURAN TERPADU
        # =====================================================
        frame = tk.Frame(scroll_frame, bg=UI_BG)
        frame.pack(fill="both", expand=True, padx=16, pady=12)
        frame.grid_columnconfigure(0, weight=1)
        frame.grid_rowconfigure(1, weight=1)

        config_card = tk.Frame(
            frame, bg=UI_PANEL, padx=16, pady=13,
            highlightthickness=1, highlightbackground=UI_BORDER
        )
        config_card.grid(row=0, column=0, sticky="ew")
        config_card.grid_columnconfigure(0, weight=3)
        config_card.grid_columnconfigure(1, weight=2)

        title_row = tk.Frame(config_card, bg=UI_PANEL)
        title_row.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 10))
        tk.Label(
            title_row, text="PENGATURAN KEGIATAN",
            bg=UI_PANEL, fg=UI_TEXT, font=("Segoe UI", 12, "bold")
        ).pack(side="left")
        tk.Label(
            title_row, text="Atur informasi, anggaran, dan pegawai dalam satu tempat",
            bg=UI_PANEL, fg=UI_MUTED, font=("Segoe UI", 9)
        ).pack(side="left", padx=(12, 0))

        settings = tk.Frame(config_card, bg=UI_PANEL)
        settings.grid(row=1, column=0, sticky="nsew", padx=(0, 18))
        settings.grid_columnconfigure(1, weight=1)

        tk.Label(settings, text="Deskripsi", bg=UI_PANEL, fg=UI_TEXT).grid(
            row=0, column=0, sticky="nw", padx=(0, 12), pady=5
        )
        self.kegiatan_var = tk.StringVar(value="")
        self.kegiatan_text = tk.Text(
            settings, height=3, width=34, wrap="word",
            font=("Segoe UI", 10), relief="solid", borderwidth=1
        )
        self.kegiatan_text.grid(row=0, column=1, sticky="ew", pady=5)

        tk.Label(
            settings, text="Penanggung Jawab", bg=UI_PANEL, fg=UI_TEXT,
            font=("Segoe UI", 9, "bold")
        ).grid(row=1, column=0, sticky="w", padx=(0, 12), pady=5)
        self.pj_var = tk.StringVar()
        self.pj_options = {}
        self.combo_pj = ttk.Combobox(
            settings, textvariable=self.pj_var, state="readonly", width=31
        )
        self.combo_pj.grid(row=1, column=1, sticky="ew", pady=5)

        tk.Label(
            settings, text="Dana Bulanan", bg=UI_PANEL, fg=UI_TEXT,
            font=("Segoe UI", 9, "bold")
        ).grid(row=2, column=0, sticky="w", padx=(0, 12), pady=5)
        self.dana_var = tk.StringVar()
        budget_actions = tk.Frame(settings, bg=UI_PANEL)
        budget_actions.grid(row=2, column=1, sticky="ew", pady=5)
        budget_actions.grid_columnconfigure(0, weight=1)
        ttk.Entry(budget_actions, textvariable=self.dana_var, width=20).pack(
            side="left", fill="x", expand=True, padx=(0, 10)
        )
        ttk.Button(
            budget_actions, text="Simpan Pengaturan",
            command=self.simpan_bagian, style="Primary.TButton"
        ).pack(side="left")
        tk.Label(
            settings, text="Nama kegiatan dapat diubah dengan double-click pada pilihan kegiatan di atas.",
            bg=UI_PANEL, fg=UI_MUTED, font=("Segoe UI", 8)
        ).grid(row=3, column=0, columnspan=2, sticky="w", pady=(5, 0))

        selection_panel = tk.Frame(config_card, bg=UI_PANEL)
        selection_panel.grid(row=1, column=1, sticky="nsew")
        selection_panel.grid_columnconfigure(0, weight=1)

        action_row = tk.Frame(selection_panel, bg=UI_PANEL)
        action_row.grid(row=0, column=0, sticky="ew")
        ttk.Button(
            action_row, text="✦  Tentukan Komposisi Target",
            command=self.tentukan_komposisi_target, style="Soft.TButton"
        ).pack(side="left")
        self.info_dana = tk.Label(
            selection_panel, text="Pegawai terpilih: 0 orang",
            bg=UI_PANEL, fg=UI_TEXT, font=("Segoe UI", 10, "bold")
        )
        self.info_dana.grid(row=1, column=0, sticky="w", pady=(8, 5))

        self.komposisi_text = tk.Text(
            selection_panel, height=5, wrap="word",
            bg=UI_PANEL, fg=UI_TEXT, relief="flat", borderwidth=0,
            font=("Segoe UI", 9), state="disabled",
            highlightthickness=1, highlightbackground=UI_BORDER
        )
        self.komposisi_text.grid(row=2, column=0, sticky="ew", pady=(2, 0))

        # Area tabel berada di bawah panel terpadu dan memakai seluruh lebar.
        self.build_tabel_pilih(frame)

    def build_tabel_pilih(self, parent):

        # Area pegawai dibagi dua: tabel di kiri dan ringkasan komposisi
        # yang dipilih di kanan. Ini membantu pengguna memilih manual.
        selection_area = tk.Frame(
            parent,
            bg=UI_BG,
            height=520
        )
        selection_area.grid_propagate(False)

        selection_area.grid(
            row=1,
            column=0,
            columnspan=2,
            sticky="nsew",
            pady=(12, 10)
        )
        selection_area.grid_columnconfigure(0, weight=1)
        selection_area.grid_rowconfigure(0, weight=0)
        selection_area.grid_rowconfigure(1, weight=1)

        # Pencarian ditempatkan tepat di atas daftar nama pegawai, bukan
        # di dalam panel kontrol "Pegawai untuk Kegiatan Ini".
        search_frame = tk.Frame(selection_area, bg=UI_BG)
        search_frame.grid(
            row=0, column=0, sticky="ew",
            padx=(0, 0), pady=(0, 8)
        )
        search_frame.grid_columnconfigure(1, weight=1)

        tk.Label(
            search_frame,
            text="Cari Pegawai:",
            bg=UI_BG,
            font=("Segoe UI", 10, "bold")
        ).grid(row=0, column=0, sticky="w", padx=(0, 8))

        self.search_pegawai_var = tk.StringVar()
        self.entry_search_pegawai = ttk.Entry(
            search_frame,
            textvariable=self.search_pegawai_var
        )
        self.entry_search_pegawai.grid(row=0, column=1, sticky="ew")
        self.entry_search_pegawai.bind(
            "<KeyRelease>",
            lambda event: self.refresh_selection_table()
        )
        ttk.Button(
            search_frame,
            text="✕",
            width=3,
            command=lambda: (
                self.search_pegawai_var.set(""),
                self.refresh_selection_table()
            )
        ).grid(row=0, column=2, padx=(5, 0))

        filter_row = tk.Frame(search_frame, bg=UI_BG)
        filter_row.grid(row=1, column=0, columnspan=3, sticky="ew", pady=(7, 0))
        filter_row.grid_columnconfigure(5, weight=1)

        self.filter_golongan_var = tk.StringVar(value="Semua Golongan")
        self.filter_jabatan_var = tk.StringVar(value="Semua Jabatan")
        self.filter_tarif_var = tk.StringVar(value="Semua Tarif")

        tk.Label(filter_row, text="Filter:", bg=UI_BG, fg=UI_MUTED,
                 font=("Segoe UI", 9, "bold")).pack(side="left", padx=(0, 8))
        self.filter_golongan_combo = ttk.Combobox(
            filter_row, textvariable=self.filter_golongan_var,
            state="readonly", width=19
        )
        self.filter_golongan_combo.pack(side="left", padx=(0, 7))
        self.filter_jabatan_combo = ttk.Combobox(
            filter_row, textvariable=self.filter_jabatan_var,
            state="readonly", width=25
        )
        self.filter_jabatan_combo.pack(side="left", padx=(0, 7))
        self.filter_tarif_combo = ttk.Combobox(
            filter_row, textvariable=self.filter_tarif_var,
            state="readonly", width=19
        )
        self.filter_tarif_combo.pack(side="left", padx=(0, 7))
        ttk.Button(
            filter_row, text="Reset Filter", style="Soft.TButton",
            command=self.reset_filter_pegawai
        ).pack(side="left")

        self._populate_employee_filters()
        for combo in (
            self.filter_golongan_combo,
            self.filter_jabatan_combo,
            self.filter_tarif_combo
        ):
            combo.bind("<<ComboboxSelected>>",
                       lambda event: self.refresh_selection_table())

        table_area = tk.Frame(
            selection_area,
            bg=UI_BG
        )
        table_area.grid(row=1, column=0, sticky="nsew")
        table_area.grid_columnconfigure(0, weight=1)
        table_area.grid_rowconfigure(0, weight=1)

        table_frame = tk.Frame(
            table_area,
            bg=UI_PANEL,
            height=520
        )

        table_frame.grid(row=0, column=0, sticky="nsew")
        table_frame.pack_propagate(False)

        columns = (
            "no",
            "pilih",
            "nama",
            "golongan",
            "jabatan",
            "tarif"
        )

        self.table_pilih = ttk.Treeview(
            table_frame,
            columns=columns,
            show="headings",
            selectmode="none"
        )

        heads = {
            "no": "No.",
            "pilih": "✓",
            "nama": "Nama",
            "golongan": "Golongan",
            "jabatan": "Jabatan",
            "tarif": "Tarif"
        }

        widths = {
            "no": 58,
            "pilih": 52,
            "nama": 330,
            "golongan": 110,
            "jabatan": 260,
            "tarif": 120
        }

        for col in columns:
            if col == "pilih":
                self.table_pilih.heading(
                    col,
                    text="✓",
                    command=self.toggle_semua_pegawai
                )
            else:
                self.table_pilih.heading(
                    col,
                    text=heads[col]
                )

            self.table_pilih.column(
                col,
                width=widths[col],
                anchor="center"
            )

        self.table_pilih.column("nama", anchor="w")
        self.table_pilih.column("jabatan", anchor="w")

        # Treeview dan scrollbar memakai grid agar keduanya selalu mendapat
        # ruang sendiri. Jika memakai pack berurutan, scrollbar dapat terdesak
        # oleh Treeview yang fill/expand.
        table_frame.grid_rowconfigure(0, weight=1)
        table_frame.grid_columnconfigure(0, weight=1)
        self.table_pilih.grid(row=0, column=0, sticky="nsew")

        scroll_table = ttk.Scrollbar(
            table_frame, orient="vertical", command=self.table_pilih.yview
        )
        scroll_table.grid(row=0, column=1, sticky="ns")
        scroll_table_x = ttk.Scrollbar(
            table_frame, orient="horizontal", command=self.table_pilih.xview
        )
        scroll_table_x.grid(row=1, column=0, sticky="ew")

        self.table_pilih.configure(
            yscrollcommand=scroll_table.set,
            xscrollcommand=scroll_table_x.set
        )
        self.modernize_tree(self.table_pilih)
        self.table_pilih.bind("<ButtonRelease-1>", self.toggle_pegawai)

        # Ringkasan komposisi sudah digabung ke panel pengaturan di atas.
        self.refresh_selection_table()

    def refresh_komposisi_manual(self):
        """Menampilkan rincian target turun dan nilai uangnya."""
        if not hasattr(self, "komposisi_text"):
            return
        self.komposisi_text.configure(state="normal")
        self.komposisi_text.delete("1.0", "end")

        target = getattr(self, "komposisi_target", {}) or {}
        target_tarif = getattr(self, "komposisi_target_tarif", {}) or {}

        if not target:
            text = (
                "Belum ada komposisi target.\n\n"
                "Klik 'Tentukan Komposisi Target' untuk menghitung kebutuhan "
                "turun berdasarkan dana bulanan."
            )
        else:
            lines = ["TARGET KOMPOSISI TURUN", ""]
            total_target_uang = 0
            total_target_turun = 0
            for golongan, jumlah in sorted(target.items()):
                jumlah = int(jumlah)
                tarif = int(target_tarif.get(golongan, 0))
                subtotal = jumlah * tarif
                total_target_turun += jumlah
                total_target_uang += subtotal
                lines.append(
                    f"• Gol {golongan}: {jumlah} turun × {rupiah(tarif)} = {rupiah(subtotal)}"
                )

            lines += [
                "",
                f"Total target turun : {total_target_turun} kali",
                f"Total target uang  : {rupiah(total_target_uang)}",
                "",
                "Nama pegawai tidak dipilih otomatis.",
                "Pegawai dan tanggal penugasan dipilih manual pada kalender."
            ]
            text = "\n".join(lines)

        self.komposisi_text.insert("1.0", text)
        self.komposisi_text.configure(state="disabled")

    def pilih_kombinasi_dialog(self, hasil, dana, nama_pj=""):
        """Memilih salah satu komposisi target berdasarkan jumlah TURUN."""
        win = tk.Toplevel(self)
        self.enable_native_window_controls(win)
        win.title("Pilih Komposisi Target")
        win.geometry("860x600")
        win.minsize(640, 460)
        win.resizable(True, True)
        win.configure(bg=UI_BG)
        win.transient(self)
        self.enable_native_window_controls(win)
        win.grab_set()
        self.enable_native_window_controls(win)
        self.center_window(win, self)

        tk.Label(
            win,
            text=f"Ditemukan {len(hasil)} komposisi tepat untuk {rupiah(dana)}",
            font=("Segoe UI", 13, "bold"), bg=UI_BG, fg=UI_TEXT
        ).pack(anchor="w", padx=18, pady=(14, 3))
        tk.Label(
            win,
            text="Komposisi dihitung dalam JUMLAH TURUN. Pegawai tidak dipilih otomatis.",
            fg=UI_MUTED, bg=UI_BG, font=("Segoe UI", 9)
        ).pack(anchor="w", padx=18, pady=(0, 9))

        outer = tk.Frame(win, bg=UI_BG)
        outer.pack(fill="both", expand=True, padx=16, pady=(0, 9))
        outer.grid_rowconfigure(0, weight=1)
        outer.grid_columnconfigure(0, weight=1)

        canvas = tk.Canvas(outer, highlightthickness=0, bg=UI_BG)
        canvas.grid(row=0, column=0, sticky="nsew")
        scrollbar = ttk.Scrollbar(outer, orient="vertical", command=canvas.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        canvas.configure(yscrollcommand=scrollbar.set)

        content = tk.Frame(canvas, bg=UI_BG)
        window_id = canvas.create_window((0, 0), window=content, anchor="nw")
        content.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind("<Configure>", lambda e: canvas.itemconfigure(window_id, width=e.width))

        selected = tk.IntVar(value=hasil[0]["nomor"] if hasil else 0)

        for item in hasil:
            card = tk.Frame(
                content, bg=UI_PANEL, bd=0, relief="flat",
                padx=10, pady=7, highlightthickness=1,
                highlightbackground=UI_BORDER
            )
            card.pack(fill="x", padx=5, pady=3)

            parts = []
            for golongan, tarif, jumlah in item["signature"]:
                subtotal = int(tarif) * int(jumlah)
                parts.append(
                    f"Gol {golongan}: {jumlah} turun × {rupiah(tarif)} = {rupiah(subtotal)}"
                )
            signature_text = "\n".join(parts)

            rb = ttk.Radiobutton(card, variable=selected, value=item["nomor"])
            rb.grid(row=0, column=0, rowspan=2, padx=(2, 10), sticky="n")
            tk.Label(
                card, text=f"Komposisi {item['nomor']}", bg=UI_PANEL,
                font=("Segoe UI", 9, "bold"), fg=UI_TEXT
            ).grid(row=0, column=1, sticky="w")
            tk.Label(
                card, text=signature_text, bg=UI_PANEL, justify="left",
                anchor="w", wraplength=560, font=("Segoe UI", 9), fg=UI_TEXT
            ).grid(row=1, column=1, sticky="w", pady=(2, 0))
            tk.Label(
                card,
                text=f"{item['jumlah_turun']} kali turun  •  {rupiah(item['total'])}",
                bg=UI_PANEL, fg=UI_PRIMARY, font=("Segoe UI", 9, "bold")
            ).grid(row=0, column=2, rowspan=2, padx=(10, 4), sticky="e")
            card.grid_columnconfigure(1, weight=1)
            card.bind("<Button-1>", lambda e, n=item["nomor"]: selected.set(n))

        result = {"pilihan": None}

        buttons = tk.Frame(win, bg=UI_BG)
        buttons.pack(fill="x", padx=16, pady=(0, 12))

        def pilih():
            if not selected.get():
                messagebox.showwarning(
                    "Pilih Komposisi",
                    "Pilih salah satu komposisi terlebih dahulu.",
                    parent=win
                )
                return
            result["pilihan"] = selected.get()
            win.destroy()

        ttk.Button(
            buttons, text="Batal", command=win.destroy
        ).pack(side="right", padx=(8, 0), ipadx=14, ipady=4)
        ttk.Button(
            buttons, text="Pilih Komposisi", command=pilih
        ).pack(side="right", ipadx=14, ipady=4)

        win.wait_window()
        return result["pilihan"]

    def buka_komposisi_manual(self):
        """Menentukan target jumlah pegawai per golongan tanpa memilih nama."""
        gol_values = sorted({str(p[5] or "").strip() for p in self.pegawai if str(p[5] or "").strip()})
        if not gol_values:
            messagebox.showwarning("Komposisi", "Belum ada data golongan pegawai.", parent=self)
            return

        win = tk.Toplevel(self)
        self.enable_native_window_controls(win)
        win.title("Tentukan Komposisi Target")
        win.geometry("620x500")
        win.minsize(560, 420)
        win.transient(self)
        self.enable_native_window_controls(win)
        win.grab_set()
        self.enable_native_window_controls(win)
        self.center_window(win, self)

        tk.Label(win, text="Tentukan Komposisi Target", font=("Segoe UI", 14, "bold")).pack(anchor="w", padx=20, pady=(18, 3))
        tk.Label(win, text="Tentukan target JUMLAH TURUN. Golongan dengan tarif yang sama dapat dihitung sebagai satu kelompok.", fg=UI_MUTED, wraplength=560, justify="left").pack(anchor="w", padx=20, pady=(0, 14))

        frame = tk.Frame(win)
        frame.pack(fill="both", expand=True, padx=20)
        frame.grid_columnconfigure(1, weight=1)
        rows = []

        tk.Label(frame, text="Golongan", font=("Segoe UI", 9, "bold")).grid(row=0, column=0, sticky="w", pady=(0, 5))
        tk.Label(frame, text="Jumlah", font=("Segoe UI", 9, "bold")).grid(row=0, column=1, sticky="w", pady=(0, 5))

        def add_row(golongan=None, jumlah=1):
            r = len(rows) + 1
            gvar = tk.StringVar(value=golongan or gol_values[0])
            jvar = tk.IntVar(value=jumlah)
            cb = ttk.Combobox(frame, textvariable=gvar, values=gol_values, state="readonly", width=18)
            sp = tk.Spinbox(frame, from_=1, to=999, textvariable=jvar, width=8)
            cb.grid(row=r, column=0, sticky="w", pady=5)
            sp.grid(row=r, column=1, sticky="w", pady=5)
            rows.append((gvar, jvar))

        target = getattr(self, "komposisi_target", {}) or {}
        if target:
            for g, n in target.items():
                add_row(g, n)
        else:
            add_row()

        ttk.Button(win, text="＋ Tambah Golongan", command=add_row).pack(anchor="w", padx=20, pady=10)

        def simpan():
            hasil = {}
            for gvar, jvar in rows:
                g = gvar.get().strip()
                try:
                    n = int(jvar.get())
                except Exception:
                    n = 0
                if g and n > 0:
                    hasil[g] = hasil.get(g, 0) + n
            if not hasil:
                messagebox.showwarning("Komposisi Kosong", "Masukkan minimal satu target golongan.", parent=win)
                return
            self.komposisi_target = hasil
            self.komposisi_target_tarif = getattr(self, "komposisi_target_tarif", {}) or {}
            for p in self.pegawai:
                gol = str(p[5] or "").strip()
                if gol in hasil and gol not in self.komposisi_target_tarif:
                    self.komposisi_target_tarif[gol] = int(p[7] or 0)
            if getattr(self, "bagian_id", None):
                conn = get_connection()
                try:
                    conn.execute("UPDATE bagian SET komposisi_target = ? WHERE id = ?", (json.dumps({
                        "target": hasil,
                        "tarif": self.komposisi_target_tarif
                    }, ensure_ascii=False), self.bagian_id))
                    conn.commit()
                finally:
                    conn.close()
            self.refresh_komposisi_manual()
            self.update_calendar_summary()
            win.destroy()

        bottom = tk.Frame(win)
        bottom.pack(fill="x", padx=20, pady=(0, 18))
        ttk.Button(bottom, text="Batal", command=win.destroy).pack(side="right", padx=(8, 0))
        ttk.Button(bottom, text="Simpan Target", command=simpan).pack(side="right")

    def pilih_pegawai_manual_berdasarkan_aturan(self, rules, pj_id):
        """Kompatibilitas lama: aturan komposisi tidak lagi memilih nama otomatis."""
        return None

    def _refresh_selection_table_chunked(self, chunk_size=80):
        """Render tabel pegawai per-batch supaya animasi dan UI tetap responsif."""
        tree = getattr(self, "table_pilih", None)
        if tree is None:
            self.hide_loading()
            return

        try:
            # Bersihkan isi lama dan siapkan gaya tema.
            for item in tree.get_children():
                tree.delete(item)

            palette = THEME_PALETTES.get(
                getattr(self, "current_theme", "light"),
                THEME_PALETTES["light"]
            )
            tree.tag_configure("baris1", background=palette["UI_PANEL"],
                               foreground=palette["UI_TEXT"])
            tree.tag_configure("baris2", background=palette["UI_TABLE_ALT"],
                               foreground=palette["UI_TEXT"])

            keyword = ""
            if hasattr(self, "search_pegawai_var"):
                keyword = self.search_pegawai_var.get().strip().lower()

            rows = []
            for p in self.pegawai:
                search_text = " ".join(str(p[i] or "") for i in range(1, 7)).lower()
                if keyword and keyword not in search_text:
                    continue
                rows.append(p)

            self._employee_render_rows = rows
            self._employee_render_index = 0
            self._employee_render_total = len(rows)
            self._employee_render_chunk_size = max(20, int(chunk_size))

            if getattr(self, "_loading_label", None) is not None:
                self._loading_label.config(
                    text=f"Menyiapkan {len(rows):,} data pegawai..."
                )

            self._render_employee_table_chunk()
        except Exception as exc:
            self.hide_loading()
            messagebox.showerror(
                "Gagal Memuat",
                f"Daftar pegawai tidak dapat ditampilkan.\\n\\nDetail: {exc}",
                parent=self
            )

    def _render_employee_table_chunk(self):
        """Masukkan sebagian baris, lalu beri event loop kesempatan menggambar."""
        try:
            tree = getattr(self, "table_pilih", None)
            if tree is None or not tree.winfo_exists():
                self.hide_loading()
                return

            rows = getattr(self, "_employee_render_rows", [])
            index = getattr(self, "_employee_render_index", 0)
            stop = min(index + self._employee_render_chunk_size, len(rows))

            for pos in range(index, stop):
                p = rows[pos]
                pegawai_id = p[0]
                tanda = "✓" if pegawai_id in self.selected_ids else ""
                tag = "baris1" if (pos + 1) % 2 == 0 else "baris2"
                tree.insert(
                    "", "end", iid=str(pegawai_id),
                    values=(
                        pos + 1, tanda, p[1], p[5], p[6], rupiah(p[7])
                    ),
                    tags=(tag,)
                )

            self._employee_render_index = stop
            total = len(rows)

            if getattr(self, "_loading_label", None) is not None:
                percent = 100 if total == 0 else int(stop * 100 / total)
                self._loading_label.config(
                    text=f"Memuat data pegawai... {stop:,} dari {total:,} ({percent}%)"
                )

            if stop < total:
                # 1 ms memberi kesempatan spinner, progressbar, dan input UI
                # diproses sebelum batch berikutnya.
                self.after(1, self._render_employee_table_chunk)
                return

            if hasattr(self, "info_dana"):
                self.info_dana.config(
                    text=f"Pegawai terpilih: {len(self.selected_ids)} orang"
                )
            self.refresh_komposisi_manual()
            self._employee_render_rows = []
            self.hide_loading()
        except Exception as exc:
            self._employee_render_rows = []
            self.hide_loading()
            messagebox.showerror(
                "Gagal Memuat",
                f"Terjadi kendala saat menyusun tabel pegawai.\\n\\nDetail: {exc}",
                parent=self
            )

    def _populate_employee_filters(self):
        """Isi opsi filter dari data pegawai aktif tanpa mengubah data sumber."""
        if not hasattr(self, "filter_golongan_combo"):
            return
        golongan = sorted({
            str(p[5] or "Tidak diketahui").strip() or "Tidak diketahui"
            for p in self.pegawai
        }, key=str.lower)
        jabatan = sorted({
            str(p[6] or "Tidak diketahui").strip() or "Tidak diketahui"
            for p in self.pegawai
        }, key=str.lower)
        tarif = sorted({int(p[7] or 0) for p in self.pegawai})
        self.filter_golongan_combo["values"] = ["Semua Golongan"] + golongan
        self.filter_jabatan_combo["values"] = ["Semua Jabatan"] + jabatan
        self.filter_tarif_combo["values"] = (
            ["Semua Tarif"] + [rupiah(v) for v in tarif]
        )

    def reset_filter_pegawai(self):
        self.search_pegawai_var.set("")
        self.filter_golongan_var.set("Semua Golongan")
        self.filter_jabatan_var.set("Semua Jabatan")
        self.filter_tarif_var.set("Semua Tarif")
        self.refresh_selection_table()

    def refresh_selection_table(self):

        # Tabel belum dibuat, jangan refresh dulu
        if not hasattr(self, "table_pilih"):
            return
        self._populate_employee_filters()

        # Bersihkan tabel
        for item in self.table_pilih.get_children():
            self.table_pilih.delete(item)

        # Warna baris selang-seling
        self.table_pilih.tag_configure(
            "baris1",
            background=THEME_PALETTES[getattr(self, "current_theme", "light")]["UI_PANEL"]
        )

        self.table_pilih.tag_configure(
            "baris2",
            background=THEME_PALETTES[getattr(self, "current_theme", "light")]["UI_TABLE_ALT"]
        )

        # Ambil teks pencarian
        keyword = ""

        if hasattr(
            self,
            "search_pegawai_var"
        ):
            keyword = (
                self.search_pegawai_var
                .get()
                .strip()
                .lower()
            )

        selected_golongan = getattr(self, "filter_golongan_var", tk.StringVar(value="Semua Golongan")).get()
        selected_jabatan = getattr(self, "filter_jabatan_var", tk.StringVar(value="Semua Jabatan")).get()
        selected_tarif = getattr(self, "filter_tarif_var", tk.StringVar(value="Semua Tarif")).get()

        total = 0
        nomor_baris = 1

        for p in self.pegawai:

            pegawai_id = p[0]

            # =================================================
            # FILTER SEARCH
            # =================================================

            teks_pencarian = " ".join([
                str(p[1] or ""),  # Nama
                str(p[2] or ""),  # NIP
                str(p[3] or ""),  # Pangkat
                str(p[4] or ""),  # Golongan asli
                str(p[5] or ""),  # Golongan utama
                str(p[6] or "")   # Jabatan
            ]).lower()

            if keyword and keyword not in teks_pencarian:
                continue
            if selected_golongan != "Semua Golongan" and (str(p[5] or "Tidak diketahui").strip() or "Tidak diketahui") != selected_golongan:
                continue
            if selected_jabatan != "Semua Jabatan" and (str(p[6] or "Tidak diketahui").strip() or "Tidak diketahui") != selected_jabatan:
                continue
            if selected_tarif != "Semua Tarif" and rupiah(p[7]) != selected_tarif:
                continue

            # =================================================
            # CHECKBOX
            # =================================================

            tanda = (
                "✓"
                if pegawai_id in self.selected_ids
                else ""
            )

            tag = (
                "baris1"
                if nomor_baris % 2 == 0
                else "baris2"
            )

            self.table_pilih.insert(
                "",
                "end",
                iid=str(pegawai_id),
                values=(
                    nomor_baris,
                    tanda,
                    p[1],
                    p[5],
                    p[6],
                    rupiah(p[7])
                ),
                tags=(tag,)
            )

            nomor_baris += 1

            # =================================================
            # TOTAL
            # =================================================

            if pegawai_id in self.selected_ids:
                total += p[7]

        # =====================================================
        # INFO TOTAL
        # =====================================================

        self.info_dana.config(
            text=f"Pegawai terpilih: {len(self.selected_ids)} orang"
        )

        self.refresh_komposisi_manual()

    # =====================================================
    # PILIH PEGAWAI MANUAL
    # =====================================================
    def toggle_semua_pegawai(self):

        semua_id = {
            p[0]
            for p in self.pegawai
        }

        # Kalau semuanya sudah dipilih → hapus semua, tetapi PJ tetap wajib.
        if self.selected_ids == semua_id:

            self.selected_ids.clear()
            if getattr(self, "penanggung_jawab_id", None):
                self.selected_ids.add(self.penanggung_jawab_id)

        # Kalau belum semuanya → pilih semua
        else:

            self.selected_ids = semua_id.copy()

        self.refresh_selection_table()

    def toggle_pegawai(self, event):

        item = self.table_pilih.identify_row(
            event.y
        )

        if not item:
            return

        pegawai_id = int(item)

        if pegawai_id == getattr(self, "penanggung_jawab_id", None):
            messagebox.showwarning(
                "Penanggung Jawab Wajib",
                "Penanggung jawab kegiatan wajib tetap dipilih."
            )
            return

        if pegawai_id in self.selected_ids:

            self.selected_ids.remove(
                pegawai_id
            )

        else:

            self.selected_ids.add(
                pegawai_id
            )

        self.refresh_selection_table()

    # =====================================================
    # SIMPAN BAGIAN
    # =====================================================

    def simpan_bagian(self):

        try:
            dana = parse_rupiah(self.dana_var.get())
        except Exception:
            messagebox.showerror("Dana", "Dana harus berupa angka.")
            return

        pj_id = self.get_selected_pj_id()
        if not pj_id:
            messagebox.showwarning(
                "Penanggung Jawab",
                "Pilih penanggung jawab kegiatan terlebih dahulu."
            )
            return

        self.penanggung_jawab_id = pj_id
        self.selected_ids.add(pj_id)

        conn = get_connection()
        cursor = conn.cursor()

        # Ambil anggota lama terlebih dahulu.
        # Ini penting supaya jadwal pegawai yang baru dihapus dari
        # kegiatan juga ikut dibersihkan otomatis.
        anggota_lama = {
            row[0]
            for row in cursor.execute(
                "SELECT pegawai_id FROM anggota_bagian WHERE bagian_id = ?",
                (self.bagian_id,)
            ).fetchall()
        }

        # Pegawai yang sebelumnya menjadi anggota tetapi sekarang
        # tidak dicentang lagi harus kehilangan seluruh jadwalnya
        # pada kegiatan ini.
        pegawai_dihapus = anggota_lama - self.selected_ids

        for pegawai_id in pegawai_dihapus:
            cursor.execute(
                """
                DELETE FROM jadwal
                WHERE bagian_id = ?
                AND pegawai_id = ?
                """,
                (self.bagian_id, pegawai_id)
            )

        nama_bagian = self.nama_bagian_var.get().strip() if hasattr(self, "nama_bagian_var") else ""
        if not nama_bagian:
            conn_nama = get_connection()
            row_nama = conn_nama.execute(
                "SELECT COALESCE(nama_bagian, '') FROM bagian WHERE id = ?",
                (self.bagian_id,)
            ).fetchone()
            conn_nama.close()
            nama_bagian = (row_nama[0] if row_nama else "").strip() or f"Kegiatan {self.bagian_aktif}"

        cursor.execute("""
            UPDATE bagian
            SET nama_bagian = ?,
                kegiatan = ?,
                dana = ?,
                penanggung_jawab_id = ?,
                komposisi_target = ?
            WHERE id = ?
        """, (
            nama_bagian,
            self.get_deskripsi(),
            dana,
            pj_id,
            json.dumps({
                "target": getattr(self, "komposisi_target", {}) or {},
                "tarif": getattr(self, "komposisi_target_tarif", {}) or {}
            }, ensure_ascii=False),
            self.bagian_id
        ))

        cursor.execute(
            "DELETE FROM anggota_bagian WHERE bagian_id = ?",
            (self.bagian_id,)
        )

        for pegawai_id in self.selected_ids:
            cursor.execute("""
                INSERT INTO anggota_bagian
                (bagian_id, pegawai_id)
                VALUES (?, ?)
            """, (self.bagian_id, pegawai_id))

        conn.commit()
        conn.close()

        self.refresh_selection_table()
        self.refresh_bagian_buttons()
        self.load_kalender()
        self.refresh_data_kegiatan()
        self.refresh_data_pegawai()
        self.load_kalender_keseluruhan()

        nama_pj = next(
            (p[1] for p in self.pegawai if p[0] == pj_id),
            "-"
        )

        messagebox.showinfo(
            "Berhasil",
            f"Kegiatan {self.bagian_aktif} berhasil disimpan.\n\n"
            f"Penanggung jawab: {nama_pj}"
        )

    def bersihkan_bagian(self):

        if not self.bagian_id:
            return

        konfirmasi = messagebox.askyesno(
            "Hapus Data Kegiatan",
            f"Yakin ingin menghapus seluruh data "
            f"KEGIATAN {self.bagian_aktif}?\n\n"
            f"Yang akan dihapus:\n"
            f"• Pegawai kegiatan ini\n"
            f"• Jadwal penugasan kegiatan ini\n"
            f"• Kegiatan\n"
                        f"• Dana bulanan\n\n"
            f"Data kegiatan lain tidak akan terpengaruh."
        )

        if not konfirmasi:
            return

        conn = get_connection()
        cursor = conn.cursor()

        # Hapus jadwal kegiatan ini
        cursor.execute(
            """
            DELETE FROM jadwal
            WHERE bagian_id = ?
            """,
            (self.bagian_id,)
        )

        # Hapus daftar pegawai bagian
        cursor.execute(
            """
            DELETE FROM anggota_bagian
            WHERE bagian_id = ?
            """,
            (self.bagian_id,)
        )

        # Reset data bagian
        cursor.execute(
            """
            UPDATE bagian
            SET kegiatan = '',
                lokasi = '',
                dana = 0,
                penanggung_jawab_id = NULL,
                komposisi_target = ''
            WHERE id = ?
            """,
            (self.bagian_id,)
        )

        conn.commit()
        conn.close()

        # Reset tampilan
        self.set_deskripsi("")
        self.dana_var.set("")

        self.selected_ids = set()
        self.komposisi_target = {}
        self.komposisi_target_tarif = {}

        if hasattr(self, "komposisi_aktif"):
            del self.komposisi_aktif

        self.refresh_selection_table()
        self.load_kalender()

        messagebox.showinfo(
            "Berhasil",
            f"Data KEGIATAN {self.bagian_aktif} "
            f"berhasil dibersihkan."
        )     
        
        # =====================================================
        # Acak Pegawai
        # =====================================================
    
    def acak_pegawai_dari_komposisi(self, *args, **kwargs):
        messagebox.showinfo("Fitur Dihapus", "Pemilihan otomatis/acak sudah dihapus. Pilih pegawai secara manual.", parent=self)

    def acak_lagi(self, *args, **kwargs):
        messagebox.showinfo("Fitur Dihapus", "Pemilihan otomatis/acak sudah dihapus. Pilih pegawai secara manual.", parent=self)

    def terapkan_kombinasi(self, *args, **kwargs):
        messagebox.showinfo("Fitur Dihapus", "Pemilihan otomatis/acak sudah dihapus. Pilih pegawai secara manual.", parent=self)

    # =====================================================
    # Hapus data bagian 
    # =====================================================

    def hapus_data_bagian(self, kode=None):
        """Hapus kegiatan tertentu. Jika kode=None, hapus kegiatan aktif."""
        kode_target = kode or self.bagian_aktif

        kode_list = load_bagian_codes(self.periode_id)
        if len(kode_list) <= 1:
            messagebox.showwarning(
                "Tidak Bisa Dihapus",
                "Minimal harus ada satu kegiatan dalam satu periode."
            )
            return

        data_target = load_bagian(self.periode_id, kode_target)
        if not data_target:
            self.refresh_bagian_buttons()
            return

        bagian_id_target = data_target[0]

        conn = get_connection()
        nama_row = conn.execute(
            "SELECT COALESCE(nama_bagian, '') FROM bagian WHERE id = ?",
            (bagian_id_target,)
        ).fetchone()
        conn.close()
        nama_target = (nama_row[0] if nama_row else "").strip() or f"KEGIATAN {kode_target}"

        konfirmasi = messagebox.askyesno(
            "Hapus Kegiatan",
            f'Yakin ingin menghapus "{nama_target}"?\n\n'
            "Kegiatan, pegawai, dan seluruh jadwal kegiatan ini akan dihapus."
        )
        if not konfirmasi:
            return

        conn = None
        try:
            conn = get_connection()
            cursor = conn.cursor()
            cursor.execute("DELETE FROM jadwal WHERE bagian_id = ?", (bagian_id_target,))
            cursor.execute("DELETE FROM anggota_bagian WHERE bagian_id = ?", (bagian_id_target,))
            cursor.execute("DELETE FROM bagian WHERE id = ?", (bagian_id_target,))
            conn.commit()
        except sqlite3.OperationalError as e:
            if conn:
                conn.rollback()
            if "locked" in str(e).lower():
                messagebox.showerror(
                    "Database Sedang Dipakai",
                    "Database sedang terkunci. Tutup program/Excel lain yang sedang membuka database, lalu coba lagi."
                )
            else:
                messagebox.showerror("Gagal Menghapus", "Data tidak dapat dihapus. Silakan coba lagi.")
            return
        except Exception as e:
            if conn:
                conn.rollback()
            messagebox.showerror("Gagal Menghapus", "Data tidak dapat dihapus. Silakan coba lagi.")
            return
        finally:
            if conn:
                conn.close()

        kode_list = load_bagian_codes(self.periode_id)
        if not kode_list:
            # Pengaman tambahan: periode tidak boleh kehilangan semua bagian.
            self.tambah_bagian()
            return

        if self.bagian_aktif == kode_target:
            self.bagian_aktif = kode_list[0]

        self.selected_ids = set()
        self.penanggung_jawab_id = None
        self.komposisi_aktif = None

        self.refresh_bagian_buttons()
        self.load_bagian_data()
        self.load_kalender()
        self.refresh_data_pegawai()
        self.refresh_selection_table()

        messagebox.showinfo("Berhasil", f'"{nama_target}" berhasil dihapus beserta seluruh jadwalnya.')

    # =====================================================
    # TENTUKAN KOMPOSISI TARGET
    # =====================================================

    def tentukan_komposisi_target(self):
        """
        Menentukan komposisi TARGET JUMLAH TURUN secara otomatis berdasarkan
        dana bulanan. Tidak memilih nama pegawai.
        """
        try:
            dana = parse_rupiah(self.dana_var.get())
        except Exception:
            messagebox.showerror(
                "Dana Tidak Valid",
                "Masukkan dana bulanan yang valid.",
                parent=self
            )
            return

        if dana <= 0:
            messagebox.showwarning(
                "Dana Kosong",
                "Masukkan dana bulanan terlebih dahulu.",
                parent=self
            )
            return

        # Kelompokkan berdasarkan TARIF, bukan nama golongan.
        # Jika beberapa golongan mempunyai tarif yang sama (mis. Gol II dan X
        # sama-sama Rp75.000), keduanya menjadi satu kelompok target.
        # Saat pegawai Gol II maupun Gol X dipilih di kalender, keduanya tetap
        # dihitung ke target kelompok tarif yang sama.
        tarif_to_golongan = {}
        for p in self.pegawai:
            gol = str(p[5] or "").strip() or "Tidak diketahui"
            tarif = int(p[7] or 0)
            if tarif > 0:
                tarif_to_golongan.setdefault(tarif, set()).add(gol)

        tarif_groups = []
        for tarif in sorted(tarif_to_golongan):
            golongan = sorted(tarif_to_golongan[tarif])
            if len(golongan) == 1:
                label = golongan[0]
            else:
                label = f"{golongan[0]} (atau {' / '.join(golongan[1:])})"
            tarif_groups.append((label, tarif))

        if not tarif_groups:
            messagebox.showwarning(
                "Tarif Tidak Tersedia",
                "Belum ada tarif pegawai yang dapat digunakan untuk menghitung komposisi.",
                parent=self
            )
            return

        hasil = cari_komposisi_turun(tarif_groups, dana, max_hasil=60)

        if not hasil:
            bawah, atas = cari_saran_dana_terdekat(
                [(None, None, None, None, None, gol, None, tarif) for gol, tarif in tarif_groups],
                dana
            )
            pesan = (
                f"Tidak ditemukan komposisi TURUN yang tepat untuk dana {rupiah(dana)}.\n\n"
                "Sistem mencari berdasarkan tarif yang tersedia pada Data Pegawai."
            )
            if bawah is not None or atas is not None:
                pesan += "\n\nSaran dana terdekat:"
                if bawah is not None:
                    pesan += f"\n• Di bawah: {rupiah(bawah)}"
                if atas is not None:
                    pesan += f"\n• Di atas: {rupiah(atas)}"
            messagebox.showwarning("Tidak Ada Komposisi", pesan, parent=self)
            return

        pilihan = self.pilih_kombinasi_dialog(hasil, dana)
        if pilihan is None:
            return

        komposisi = hasil[pilihan - 1]
        target = {}
        target_tarif = {}
        for golongan, tarif, jumlah in komposisi["signature"]:
            target[golongan] = target.get(golongan, 0) + int(jumlah)
            target_tarif[golongan] = int(tarif)

        self.komposisi_target = target
        self.komposisi_target_tarif = target_tarif
        self.komposisi_aktif = komposisi
        self.kombinasi_nomor = pilihan

        if getattr(self, "bagian_id", None):
            conn = get_connection()
            try:
                conn.execute(
                    "UPDATE bagian SET komposisi_target = ?, penanggung_jawab_id = COALESCE(penanggung_jawab_id, penanggung_jawab_id) WHERE id = ?",
                    (json.dumps({
                        "target": target,
                        "tarif": target_tarif
                    }, ensure_ascii=False), self.bagian_id)
                )
                conn.commit()
            finally:
                conn.close()

        self.refresh_komposisi_manual()
        self.update_calendar_summary()

        rincian = "\n".join(
            f"Gol {gol}: {jumlah} turun × {rupiah(target_tarif.get(gol, 0))} = {rupiah(int(jumlah) * int(target_tarif.get(gol, 0)))}"
            for gol, jumlah in sorted(target.items())
        )
        messagebox.showinfo(
            "Komposisi Target Berhasil",
            f"Dana bulanan: {rupiah(dana)}\n\n"
            f"{rincian}\n\n"
            f"TOTAL TARGET: {rupiah(komposisi['total'])}\n"
            f"TOTAL TURUN: {komposisi['jumlah_turun']} kali\n\n"
            "Nama pegawai tidak dipilih otomatis. Pegawai tetap dipilih manual pada kalender.",
            parent=self
        )

    def auto_tentukan(self):
        """Kompatibilitas nama lama; sekarang menentukan komposisi target tanpa memilih nama."""
        return self.tentukan_komposisi_target()

    # =====================================================
    # KALENDER
    # =====================================================

    def label_kegiatan(self, kode):
        """Nama kegiatan yang ramah untuk kalender dan pesan bentrok."""
        conn = get_connection()
        row = conn.execute(
            "SELECT COALESCE(nama_bagian, '') FROM bagian WHERE periode_id = ? AND kode = ?",
            (self.periode_id, kode)
        ).fetchone()
        conn.close()

        nama = (row[0] if row else "").strip()
        default = f"Kegiatan {kode}"
        if not nama or nama == default:
            return f"Kegiatan {kode}"
        return f"Kegiatan {kode} — {nama}"

    def ringkasan_kode_kegiatan(self):
        conn = get_connection()
        rows = conn.execute(
            "SELECT kode, COALESCE(nama_bagian, '') FROM bagian WHERE periode_id = ? ORDER BY id",
            (self.periode_id,)
        ).fetchall()
        conn.close()

        bagian = []
        for kode, nama in rows:
            nama = (nama or '').strip()
            if nama and nama != f"Kegiatan {kode}":
                bagian.append(f"{kode} = {nama}")
            else:
                bagian.append(f"{kode} = Kegiatan {kode}")
        return "  •  ".join(bagian)

    def update_notebook_title(self):
        if not hasattr(self, "notebook"):
            return
        nama = self.nama_bagian_var.get().strip() if hasattr(self, "nama_bagian_var") else f"Kegiatan {self.bagian_aktif}"
        self.notebook.tab(
            self.tab_pengaturan,
            text=f"  Pengaturan Kegiatan {self.bagian_aktif}  "
        )
        self.notebook.tab(
            self.tab_jadwal,
            text="  Kalender Kegiatan  "
        )
        self.notebook.tab(
            self.tab_data_kegiatan,
            text=f"  Data {nama}  "
        )

    def toggle_calendar_summary(self):
        self._calendar_summary_expanded = not getattr(
            self, "_calendar_summary_expanded", False
        )
        self.update_calendar_summary()

    def toggle_calendar_legend(self):
        if not hasattr(self, "calendar_legend_popup") or not self.calendar_legend_popup.winfo_exists():
            popup = tk.Toplevel(self)
            self.calendar_legend_popup = popup
            popup.overrideredirect(True)
            popup.configure(bg=UI_PANEL)
            popup.attributes("-topmost", True)
            frame = tk.Frame(
                popup, bg=UI_PANEL, padx=12, pady=10,
                highlightthickness=1, highlightbackground=UI_BORDER
            )
            frame.pack(fill="both", expand=True)
            tk.Label(
                frame, text="Keterangan warna kegiatan",
                bg=UI_PANEL, fg=UI_TEXT, font=("Segoe UI", 10, "bold")
            ).pack(anchor="w", pady=(0, 7))
            for item in self.legend_kegiatan.winfo_children():
                item.pack_configure(anchor="w", fill="x", pady=2)
                clone = tk.Label(
                    frame, text=item.cget("text"), bg=item.cget("bg"),
                    fg=item.cget("fg"), padx=9, pady=5,
                    font=("Segoe UI", 9, "bold"), anchor="w"
                )
                clone.pack(fill="x", pady=2)
            x = self.calendar_legend_button.winfo_rootx()
            y = self.calendar_legend_button.winfo_rooty() + self.calendar_legend_button.winfo_height() + 5
            popup.geometry(f"+{x}+{y}")
            popup.bind("<FocusOut>", lambda e: popup.destroy())
            popup.focus_force()
        else:
            self.calendar_legend_popup.destroy()

    def _show_calendar_cell_tooltip(self, widget, text):
        self._hide_calendar_cell_tooltip()
        tip = tk.Toplevel(self)
        self._calendar_cell_tooltip = tip
        tip.overrideredirect(True)
        tip.attributes("-topmost", True)
        frame = tk.Frame(
            tip, bg=UI_PANEL, padx=10, pady=8,
            highlightthickness=1, highlightbackground=UI_BORDER
        )
        frame.pack()
        tk.Label(
            frame, text=text, bg=UI_PANEL, fg=UI_TEXT,
            justify="left", font=("Segoe UI", 9), wraplength=360
        ).pack()
        x = widget.winfo_rootx() + widget.winfo_width() + 8
        y = widget.winfo_rooty() + 4
        tip.geometry(f"+{x}+{y}")

    def _hide_calendar_cell_tooltip(self):
        tip = getattr(self, "_calendar_cell_tooltip", None)
        if tip is not None:
            try:
                if tip.winfo_exists():
                    tip.destroy()
            except tk.TclError:
                pass
        self._calendar_cell_tooltip = None

    def update_calendar_summary(self):
        """Menampilkan rincian target TURUN dan realisasi bulanan."""
        if not hasattr(self, "calendar_summary"):
            return

        conn = get_connection()
        try:
            rows = conn.execute("""
                SELECT p.tarif, COUNT(*), COALESCE(SUM(p.tarif),0)
                FROM jadwal j JOIN pegawai p ON p.id=j.pegawai_id
                WHERE j.periode_id=? AND j.bagian_id=?
                GROUP BY p.tarif
            """, (self.periode_id, self.bagian_id)).fetchall()
            total_uang = conn.execute("""
                SELECT COALESCE(SUM(p.tarif),0)
                FROM jadwal j JOIN pegawai p ON p.id=j.pegawai_id
                WHERE j.periode_id=? AND j.bagian_id=?
            """, (self.periode_id, self.bagian_id)).fetchone()[0]
            desa_rows = conn.execute("""
                SELECT j.hari, COALESCE(d.nama, 'Belum dipilih')
                FROM jadwal j LEFT JOIN desa d ON d.id=j.desa_id
                WHERE j.periode_id=? AND j.bagian_id=?
                GROUP BY j.hari, d.nama ORDER BY j.hari
            """, (self.periode_id, self.bagian_id)).fetchall()
        finally:
            conn.close()

        actual = {int(tarif or 0): (int(n), int(uang or 0)) for tarif,n,uang in rows}
        target = getattr(self, "komposisi_target", {}) or {}
        target_tarif = getattr(self, "komposisi_target_tarif", {}) or {}

        # Jika data lama hanya menyimpan target angka, ambil tarif dari master.
        if target and not target_tarif:
            for p in self.pegawai:
                gol = str(p[5] or "").strip() or "Tidak diketahui"
                if gol in target:
                    target_tarif.setdefault(gol, int(p[7] or 0))

        for item in self.calendar_summary.get_children():
            self.calendar_summary.delete(item)

        total_target_uang = 0
        total_target_turun = 0
        if target:
            for g, n in sorted(target.items()):
                n = int(n)
                tarif = int(target_tarif.get(g, 0))
                subtotal = n * tarif
                got, uang = actual.get(tarif, (0, 0))
                selisih = got - n
                total_target_turun += n
                total_target_uang += subtotal

                if selisih < 0:
                    status = f"Kurang {abs(selisih)}"
                elif selisih > 0:
                    status = f"Lebih {selisih}"
                else:
                    status = "Terpenuhi"

                self.calendar_summary.insert(
                    "", "end",
                    values=(
                        f"Golongan {g}", f"{n} turun", rupiah(tarif),
                        rupiah(subtotal), f"{got} turun",
                        f"{selisih:+d}" if selisih else "0",
                        status
                    )
                )
        else:
            self.calendar_summary.insert(
                "", "end",
                values=("Belum ditentukan", "—", "—", "—", "—", "—", "—")
            )

        total_actual = sum(v[0] for v in actual.values())
        self.calendar_summary_totals.config(
            text=(
                f"TOTAL TARGET  {total_target_turun} turun  •  {rupiah(total_target_uang)}"
                f"        |        TOTAL REALISASI  {total_actual} turun"
                f"  •  {rupiah(total_uang)}"
            )
        )

        if desa_rows:
            village_text = "Lokasi penugasan:  " + "     •     ".join(
                f"Tgl {h}: {d}" for h, d in desa_rows
            )
        else:
            village_text = "Lokasi penugasan: belum ada."
        self.calendar_summary_villages.config(text=village_text)

        total_target = sum(int(v) for v in target.values()) if target else 0
        self.calendar_summary_status.config(
            text=f"{total_actual}/{total_target} turun terealisasi" if target
            else f"{total_actual} turun terealisasi"
        )
        if self._calendar_summary_expanded:
            self.calendar_summary.pack(fill="x", pady=(7, 2))
            self.calendar_summary_toggle.configure(text="▾  Ringkasan Komposisi & Realisasi")
        else:
            self.calendar_summary.pack_forget()
            self.calendar_summary_toggle.configure(text="▸  Ringkasan Komposisi & Realisasi")

    def peringatan_komposisi_tanggal(self, hari):
        """Menghitung sisa target bulanan setelah tanggal yang sedang dipilih."""
        target = getattr(self, "komposisi_target", {}) or {}
        if not target or not self.bagian_id:
            return ""

        conn = get_connection()
        try:
            rows = conn.execute("""
                SELECT p.tarif, COUNT(*)
                FROM jadwal j JOIN pegawai p ON p.id=j.pegawai_id
                WHERE j.periode_id=? AND j.bagian_id=?
                GROUP BY p.tarif
            """, (self.periode_id, self.bagian_id)).fetchall()
        finally:
            conn.close()

        actual = {int(tarif or 0): int(n) for tarif,n in rows}
        target_tarif = getattr(self, "komposisi_target_tarif", {}) or {}
        messages = []
        for g, n in sorted(target.items()):
            tarif = int(target_tarif.get(g, 0))
            got = actual.get(tarif, 0)
            if got < int(n):
                messages.append(
                    f"Gol {g}: masih kurang {int(n)-got} turun ({got}/{int(n)})"
                )
            elif got > int(n):
                messages.append(
                    f"Gol {g}: melebihi target {got-int(n)} turun ({got}/{int(n)})"
                )
        return " • ".join(messages)

    def load_kalender(self):
        """Kalender kegiatan dengan header tanggal tetap saat scroll vertikal.

        Nama/no pegawai tetap di kiri; panel nama dan jadwal bergerak vertikal
        bersama. Header tanggal terpisah dari badan tabel, tetapi horizontalnya
        tersinkron agar tanggal tetap tepat di atas kolom yang sesuai.
        """
        for target in (
            self.calendar_frame, self.calendar_names_frame,
            self.calendar_header_frame, self.calendar_names_header_frame
        ):
            for widget in target.winfo_children():
                widget.destroy()

        bulan = NAMA_BULAN.index(self.bulan_var.get())
        tahun = int(self.tahun_var.get())
        jumlah_hari = calendar.monthrange(tahun, bulan)[1]
        hari_libur = load_hari_libur(self.periode_id)

        conn = get_connection()
        kegiatan_rows = conn.execute(
            "SELECT kode, COALESCE(nama_bagian,''), COALESCE(kegiatan,'') FROM bagian WHERE periode_id=? ORDER BY id",
            (self.periode_id,)
        ).fetchall()
        jadwal_rows = conn.execute("""
            SELECT j.pegawai_id,j.hari,b.kode,COALESCE(b.nama_bagian,''),COALESCE(b.kegiatan,''),j.desa_id,COALESCE(d.nama,'')
            FROM jadwal j JOIN bagian b ON b.id=j.bagian_id LEFT JOIN desa d ON d.id=j.desa_id
            WHERE j.periode_id=?
        """, (self.periode_id,)).fetchall()
        anggota_rows = conn.execute(
            "SELECT pegawai_id FROM anggota_bagian WHERE bagian_id=?", (self.bagian_id,)
        ).fetchall() if self.bagian_id else []
        conn.close()

        if getattr(self, "current_theme", "light") == "dark":
            palette = [
                "#285B78", "#28634F", "#76552B", "#594477",
                "#743F4C", "#28666A", "#66552C", "#486044"
            ]
        else:
            palette = [
                "#D7EAFE", "#D9F2E4", "#FCECCB", "#E9DDF8",
                "#F9DDE2", "#D7F0F0", "#F3E8C9", "#E2EED8"
            ]
        self.calendar_colors = {r[0]: palette[i % len(palette)] for i, r in enumerate(kegiatan_rows)}
        jadwal_map = {
            (r[0], r[1]): {"kode": r[2], "nama": (r[3] or "").strip() or f"Kegiatan {r[2]}", "desa_id": r[5], "desa": r[6]}
            for r in jadwal_rows
        }

        self.info_libur.config(
            text=f"Libur manual: {', '.join(map(str, sorted(hari_libur))) if hari_libur else 'tidak ada'}  ·  Minggu otomatis libur"
        )
        for w in self.legend_kegiatan.winfo_children():
            w.destroy()
        for kode, nama, _ in kegiatan_rows:
            tk.Label(
                self.legend_kegiatan,
                text=f"{kode} · {(nama or '').strip() or f'Kegiatan {kode}'}",
                bg=self.calendar_colors[kode],
                fg="#172033" if getattr(self, "current_theme", "light") != "dark" else "#FFFFFF",
                padx=9, pady=6, font=("Segoe UI", 9, "bold")
            ).pack(anchor="w", fill="x", pady=2)
        self.calendar_legend_button.configure(
            command=self.toggle_calendar_legend
        )

        # Header kiri tetap di tempat; header tanggal terpisah dari area scroll vertikal.
        ttk.Label(self.calendar_names_header_frame, text="No.", font=("Segoe UI", 9, "bold"), anchor="center").grid(row=0, column=0, sticky="nsew", padx=1, pady=1)
        ttk.Label(self.calendar_names_header_frame, text="Nama Pegawai", font=("Segoe UI", 9, "bold"), anchor="center").grid(row=0, column=1, sticky="nsew", padx=1, pady=1)
        self.calendar_names_header_frame.grid_columnconfigure(0, minsize=42, weight=0)
        self.calendar_names_header_frame.grid_columnconfigure(1, minsize=245, weight=1)

        nama_hari = ["Sen", "Sel", "Rab", "Kam", "Jum", "Sab", "Min"]
        for hari in range(1, jumlah_hari + 1):
            tanggal = datetime(tahun, bulan, hari)
            nonkerja = tanggal.weekday() == 6 or hari in hari_libur
            ttk.Label(self.calendar_header_frame, text=str(hari), font=("Segoe UI", 9, "bold"), anchor="center", style="Holiday.TLabel" if nonkerja else "TLabel").grid(row=0, column=hari-1, sticky="nsew", padx=1, pady=(1, 0))
            ttk.Label(self.calendar_header_frame, text="LIBUR" if nonkerja else nama_hari[tanggal.weekday()], font=("Segoe UI", 7 if nonkerja else 8, "bold" if nonkerja else "normal"), anchor="center", style="Holiday.TLabel" if nonkerja else "TLabel").grid(row=1, column=hari-1, sticky="nsew", padx=1, pady=(0, 1))
            self.calendar_header_frame.grid_columnconfigure(hari-1, minsize=78, weight=0, uniform="activity_day_columns")
            self.calendar_frame.grid_columnconfigure(hari-1, minsize=78, weight=0, uniform="activity_day_columns")

        anggota_ids = {r[0] for r in anggota_rows}
        semua_pegawai = sorted([p for p in self.pegawai if p[0] in anggota_ids], key=lambda p: str(p[1] or "").lower())
        if not semua_pegawai:
            ttk.Label(self.calendar_names_frame, text="—", anchor="center").grid(row=0, column=0, sticky="nsew", padx=1, pady=1)
            ttk.Label(self.calendar_names_frame, text="Belum ada pegawai yang dipilih untuk kegiatan ini.", anchor="w").grid(row=0, column=1, sticky="nsew", padx=1, pady=1)
            ttk.Label(self.calendar_frame, text="Belum ada pegawai yang dipilih untuk kegiatan ini.", anchor="w").grid(row=0, column=0, sticky="nsew", padx=1, pady=1)
            self.calendar_names_frame.grid_rowconfigure(0, minsize=47)
            self.calendar_frame.grid_rowconfigure(0, minsize=47)
        else:
            for nomor, p in enumerate(semua_pegawai, start=1):
                row = nomor - 1
                pid, nama = p[0], p[1] or "Tanpa Nama"
                ttk.Label(self.calendar_names_frame, text=str(nomor), anchor="center", padding=(2, 0)).grid(row=row, column=0, sticky="nsew", padx=1, pady=1)
                ttk.Label(self.calendar_names_frame, text=nama, anchor="w", padding=(5, 0)).grid(row=row, column=1, sticky="nsew", padx=1, pady=1)
                for hari in range(1, jumlah_hari + 1):
                    tanggal = datetime(tahun, bulan, hari)
                    nonkerja = tanggal.weekday() == 6 or hari in hari_libur
                    detail = jadwal_map.get((pid, hari))
                    if nonkerja:
                        cell = tk.Button(
                            self.calendar_frame, text="LIBUR", width=6, state="disabled",
                            disabledforeground=UI_MUTED,
                            bg=UI_TABLE_HEAD, activebackground=UI_TABLE_HEAD,
                            relief="solid", bd=1, font=("Segoe UI", 8),
                            highlightthickness=0
                        )
                    else:
                        kode = detail["kode"] if detail else ""
                        desa = detail["desa"] if detail else ""
                        text = kode + (f"\n{desa}" if desa else "")
                        cell = tk.Button(
                            self.calendar_frame, text=(kode + (f"\n{desa[:10] + ('…' if len(desa) > 10 else '')}" if desa else "")),
                            width=6, height=2, wraplength=64,
                            bg=self.calendar_colors.get(kode, UI_PANEL) if detail else UI_PANEL,
                            fg="#F8FAFC" if getattr(self, "current_theme", "light") == "dark" and detail else UI_TEXT,
                            activebackground=self.calendar_colors.get(kode, UI_TABLE_ALT) if detail else UI_TABLE_ALT,
                            activeforeground="#FFFFFF" if getattr(self, "current_theme", "light") == "dark" and detail else UI_TEXT,
                            relief="solid", bd=1, highlightthickness=0,
                            highlightbackground=UI_BORDER,
                            font=("Segoe UI", 8, "bold") if kode else ("Segoe UI", 8),
                            command=(lambda pid=pid, h=hari, d=detail: self.popup_detail_penugasan(
                                pid, h, self._bagian_id_from_kode(d["kode"]), parent=self
                            ) if d else self.toggle_jadwal(pid, h))
                        )
                    if detail:
                        def show_cell_detail(event, n=nama, h=hari, d=detail, widget=cell):
                            warning = self.peringatan_komposisi_tanggal(h)
                            detail_text = (
                                f"{n}  ·  Tanggal {h}  ·  {d['kode']}\n"
                                f"Lokasi: {d['desa'] or 'Belum dipilih'}"
                            )
                            if warning:
                                detail_text += f"\n⚠ {warning}"
                            self._show_calendar_cell_tooltip(widget, detail_text)
                        cell.bind("<Enter>", show_cell_detail)
                        cell.bind("<Leave>", lambda e: self._hide_calendar_cell_tooltip())
                    cell.grid(row=row, column=hari-1, sticky="nsew", padx=1, pady=1)
                self.calendar_names_frame.grid_rowconfigure(row, minsize=47)
                self.calendar_frame.grid_rowconfigure(row, minsize=47)

        self.calendar_names_frame.grid_columnconfigure(0, minsize=42, weight=0)
        self.calendar_names_frame.grid_columnconfigure(1, minsize=245, weight=1)
        self.calendar_header_frame.grid_rowconfigure(0, minsize=26)
        self.calendar_header_frame.grid_rowconfigure(1, minsize=22)
        self.calendar_names_header_frame.grid_rowconfigure(0, minsize=48)
        self.calendar_header_frame.update_idletasks()
        self.calendar_names_header_frame.update_idletasks()
        try:
            if hasattr(self, "_refresh_activity_calendar_regions"):
                self._refresh_activity_calendar_regions()
                self.after_idle(self._refresh_activity_calendar_regions)
        except Exception:
            pass
        self.update_calendar_summary()
        # Kalender Keseluruhan tidak diubah oleh revisi ini.
        self.load_kalender_keseluruhan()

    def pilih_penugasan_keseluruhan(self, pegawai_id, hari):
        """Menambahkan penugasan langsung dari Kalender Keseluruhan."""
        # Tampilkan spinner SEBELUM query dan pembuatan popup supaya user
        # langsung mendapat umpan balik bahwa klik sedang diproses.
        self.show_loading("Menyiapkan pilihan penugasan...", parent=self.overall_calendar_window)
        self.update_idletasks()

        bulan = NAMA_BULAN.index(self.bulan_var.get())
        tahun = int(self.tahun_var.get())
        if datetime(tahun, bulan, hari).weekday() == 6 or hari in load_hari_libur(self.periode_id):
            self.hide_loading()
            return

        conn = get_connection()
        try:
            existing = conn.execute(
                "SELECT b.kode, COALESCE(b.nama_bagian,''), COALESCE(d.nama,'') "
                "FROM jadwal j JOIN bagian b ON b.id=j.bagian_id "
                "LEFT JOIN desa d ON d.id=j.desa_id "
                "WHERE j.periode_id=? AND j.pegawai_id=? AND j.hari=?",
                (self.periode_id, pegawai_id, hari)
            ).fetchone()
            if existing:
                messagebox.showinfo(
                    "Penugasan Sudah Ada",
                    f"Pegawai ini sudah memiliki penugasan pada tanggal {hari}:\n"
                    f"{existing[0]} — {existing[1] or ('Kegiatan ' + existing[0])}\n"
                    f"Lokasi: {existing[2] or '-'}",
                    parent=self.overall_calendar_window
                )
                self.hide_loading()
                return

            kegiatan_rows = conn.execute(
                "SELECT id, kode, COALESCE(nama_bagian,''), penanggung_jawab_id "
                "FROM bagian WHERE periode_id=? ORDER BY id",
                (self.periode_id,)
            ).fetchall()
            desa_rows = load_desa_aktif()
        finally:
            conn.close()

        if not kegiatan_rows:
            self.hide_loading()
            messagebox.showwarning("Kegiatan", "Belum ada kegiatan pada periode ini.", parent=self.overall_calendar_window)
            return
        if not desa_rows:
            self.hide_loading()
            messagebox.showwarning("Data Lokasi", "Belum ada lokasi aktif. Tambahkan lokasi terlebih dahulu.", parent=self.overall_calendar_window)
            return

        # Query selesai; tutup spinner tepat sebelum popup tampil.
        self.hide_loading()
        win = tk.Toplevel(self.overall_calendar_window)
        self.enable_native_window_controls(win)
        win.title(f"Tambah Penugasan — Tanggal {hari}")
        win.geometry("600x360")
        win.minsize(560, 340)
        win.transient(self.overall_calendar_window)
        self.enable_native_window_controls(win)
        win.grab_set()
        self.enable_native_window_controls(win)

        win.configure(bg=UI_BG)
        tk.Label(
            win, text=f"Penugasan tanggal {hari}",
            bg=UI_BG, fg=UI_TEXT, font=("Segoe UI Semibold", 16)
        ).pack(anchor="w", padx=24, pady=(22, 4))
        tk.Label(
            win,
            text="Pilih kegiatan dan desa. Penanggung Jawab kegiatan akan otomatis ikut turun.",
            bg=UI_BG, fg=UI_MUTED, font=("Segoe UI", 9)
        ).pack(anchor="w", padx=24, pady=(0, 16))

        form = tk.Frame(win, bg=UI_PANEL, highlightbackground=UI_BORDER, highlightthickness=1)
        form.pack(fill="x", padx=24, pady=(0, 6))
        form.columnconfigure(1, weight=1)

        tk.Label(form, text="Kegiatan:").grid(row=0, column=0, sticky="w", pady=7)
        kegiatan_var = tk.StringVar()
        kegiatan_values = [
            f"{r[1]} — {(r[2] or '').strip() or ('Kegiatan ' + r[1])}" + ("" if r[3] else "  • PJ belum ditentukan")
            for r in kegiatan_rows
        ]
        kegiatan_combo = ttk.Combobox(form, textvariable=kegiatan_var, values=kegiatan_values, state="readonly")
        kegiatan_combo.grid(row=0, column=1, sticky="ew", padx=(12, 0), pady=7)
        kegiatan_combo.current(0)

        tk.Label(form, text="Lokasi:").grid(row=1, column=0, sticky="w", pady=7)
        desa_var = tk.StringVar()
        desa_combo = ttk.Combobox(form, textvariable=desa_var, values=[r[1] for r in desa_rows], state="readonly")
        desa_combo.grid(row=1, column=1, sticky="ew", padx=(12, 0), pady=7)
        desa_combo.current(0)

        info = tk.Label(
            win, text="", bg=UI_BG, fg=UI_MUTED,
            justify="left", anchor="w", font=("Segoe UI", 9)
        )
        info.pack(fill="x", padx=24, pady=(8, 0))

        def update_desa_default(*_):
            idx = kegiatan_combo.current()
            if idx < 0:
                return
            bagian_id = kegiatan_rows[idx][0]
            conn2 = get_connection()
            try:
                row = conn2.execute(
                    "SELECT d.id, d.nama FROM jadwal j JOIN desa d ON d.id=j.desa_id "
                    "WHERE j.periode_id=? AND j.bagian_id=? AND j.hari=? LIMIT 1",
                    (self.periode_id, bagian_id, hari)
                ).fetchone()
            finally:
                conn2.close()
            if row and row[1] in [x[1] for x in desa_rows]:
                desa_var.set(row[1])
                info.config(text=f"Kegiatan ini sudah memakai {row[1]} pada tanggal {hari}; desa disamakan otomatis.")
            else:
                if not desa_var.get():
                    desa_combo.current(0)
                info.config(text="")

        kegiatan_combo.bind("<<ComboboxSelected>>", update_desa_default)
        update_desa_default()

        result = {"saved": False}
        bottom = tk.Frame(win, bg=UI_BG)
        bottom.pack(fill="x", padx=24, pady=(14, 22))

        def batal():
            win.destroy()

        def simpan():
            idx = kegiatan_combo.current()
            desa_nama = desa_var.get().strip()
            if idx < 0 or not desa_nama:
                messagebox.showwarning("Penugasan", "Pilih kegiatan dan desa terlebih dahulu.", parent=win)
                return
            bagian_id, kode, nama_bagian, pj_id = kegiatan_rows[idx]
            if not pj_id:
                nama_tampil = (nama_bagian or '').strip() or f"Kegiatan {kode}"
                messagebox.showwarning(
                    "Penanggung Jawab Belum Ditentukan",
                    f"{kode} — {nama_tampil} belum memiliki Penanggung Jawab.\n\n"
                    "Tentukan PJ terlebih dahulu pada Pengaturan Kegiatan sebelum membuat penugasan.",
                    parent=win
                )
                return
            desa_row = next((r for r in desa_rows if r[1] == desa_nama), None)
            if not desa_row:
                return
            desa_id = desa_row[0]

            conn2 = get_connection()
            try:
                # Satu desa tidak boleh dipakai dua kegiatan pada tanggal yang sama.
                konflik_desa = conn2.execute(
                    "SELECT b.kode, COALESCE(b.nama_bagian,'') FROM jadwal j "
                    "JOIN bagian b ON b.id=j.bagian_id WHERE j.periode_id=? AND j.hari=? "
                    "AND j.desa_id=? AND j.bagian_id<>? LIMIT 1",
                    (self.periode_id, hari, desa_id, bagian_id)
                ).fetchone()
                if konflik_desa:
                    messagebox.showwarning(
                        "Desa Bentrok",
                        f"{desa_nama} sudah digunakan oleh Kegiatan {konflik_desa[0]} — "
                        f"{konflik_desa[1] or ('Kegiatan ' + konflik_desa[0])} pada tanggal {hari}.",
                        parent=win
                    )
                    return

                # PJ juga harus bebas pada tanggal tersebut karena setiap pegawai hanya boleh satu kegiatan/hari.
                if pj_id and int(pj_id) != int(pegawai_id):
                    pj_conflict = conn2.execute(
                        "SELECT b.kode, COALESCE(b.nama_bagian,'') FROM jadwal j "
                        "JOIN bagian b ON b.id=j.bagian_id WHERE j.periode_id=? AND j.pegawai_id=? "
                        "AND j.hari=? AND j.bagian_id<>? LIMIT 1",
                        (self.periode_id, pj_id, hari, bagian_id)
                    ).fetchone()
                    if pj_conflict:
                        messagebox.showwarning(
                            "PJ Bentrok",
                            f"PJ untuk Kegiatan {kode} sudah bertugas pada Kegiatan {pj_conflict[0]} "
                            f"pada tanggal {hari}. Penugasan dari Kalender Keseluruhan dibatalkan.",
                            parent=win
                        )
                        return

                conn2.execute(
                    "INSERT INTO jadwal (periode_id, bagian_id, pegawai_id, hari, desa_id) VALUES (?, ?, ?, ?, ?)",
                    (self.periode_id, bagian_id, pegawai_id, hari, desa_id)
                )
                if pj_id and int(pj_id) != int(pegawai_id):
                    pj_existing = conn2.execute(
                        "SELECT id FROM jadwal WHERE periode_id=? AND bagian_id=? AND pegawai_id=? AND hari=?",
                        (self.periode_id, bagian_id, pj_id, hari)
                    ).fetchone()
                    if not pj_existing:
                        conn2.execute(
                            "INSERT INTO jadwal (periode_id, bagian_id, pegawai_id, hari, desa_id) VALUES (?, ?, ?, ?, ?)",
                            (self.periode_id, bagian_id, pj_id, hari, desa_id)
                        )
                conn2.commit()
            except sqlite3.IntegrityError as exc:
                conn2.rollback()
                messagebox.showwarning("Penugasan", "Penugasan tidak dapat disimpan. Periksa kembali pilihan kegiatan, desa, dan jadwal pegawai.", parent=win)
                return
            except Exception as exc:
                conn2.rollback()
                messagebox.showerror("Penugasan", "Penugasan tidak dapat disimpan. Silakan coba lagi.", parent=win)
                return
            finally:
                conn2.close()

            result["saved"] = True
            win.destroy()

        ttk.Button(bottom, text="Batal", command=batal).pack(side="right", padx=(8, 0))
        ttk.Button(bottom, text="Simpan", command=simpan, style="Primary.TButton").pack(side="right")
        self.fit_popup(
            win, self.overall_calendar_window,
            min_width=560, min_height=340, max_width=760, max_height=520
        )
        win.wait_window()

        if result["saved"]:
            self.refresh_after_schedule(
                "Menyimpan penugasan dan menyegarkan semua kalender...",
                parent=self.overall_calendar_window
            )

    def _cek_pj_kegiatan(self, bagian_id, parent=None):
        """Pastikan kegiatan memiliki PJ sebelum penugasan dibuat."""
        conn = get_connection()
        try:
            row = conn.execute(
                "SELECT kode, COALESCE(nama_bagian,''), penanggung_jawab_id FROM bagian WHERE id=?",
                (bagian_id,)
            ).fetchone()
        finally:
            conn.close()

        if not row:
            messagebox.showwarning(
                "Kegiatan Tidak Ditemukan",
                "Kegiatan yang dipilih tidak tersedia. Silakan muat ulang halaman.",
                parent=parent or self
            )
            return False

        kode, nama, pj_id = row
        if not pj_id:
            nama_tampil = (nama or '').strip() or f"Kegiatan {kode}"
            messagebox.showwarning(
                "Penanggung Jawab Belum Ditentukan",
                f"{kode} — {nama_tampil} belum memiliki Penanggung Jawab.\n\n"
                "Tentukan PJ terlebih dahulu pada Pengaturan Kegiatan sebelum membuat penugasan.",
                parent=parent or self
            )
            return False
        return True

    def _bagian_id_from_kode(self, kode):
        conn = get_connection()
        try:
            row = conn.execute(
                "SELECT id FROM bagian WHERE periode_id=? AND kode=?",
                (self.periode_id, kode)
            ).fetchone()
            return row[0] if row else None
        finally:
            conn.close()

    def popup_detail_penugasan(self, pegawai_id, hari, bagian_id, parent=None):
        """Menampilkan detail penugasan yang dipilih; hapus hanya setelah konfirmasi."""
        from_overall_calendar = (
            parent is not None
            and parent is getattr(self, "overall_calendar_window", None)
        )
        parent = parent or getattr(self, "overall_calendar_window", None) or self
        bulan = NAMA_BULAN.index(self.bulan_var.get())
        tahun = int(self.tahun_var.get())
        conn = get_connection()
        try:
            row = conn.execute("""
                SELECT p.nama, b.kode, COALESCE(b.nama_bagian,''), COALESCE(b.kegiatan,''),
                       COALESCE(b.lokasi,''), COALESCE(b.dana,0), b.penanggung_jawab_id,
                       COALESCE(d.nama,'')
                FROM jadwal j
                JOIN pegawai p ON p.id=j.pegawai_id
                JOIN bagian b ON b.id=j.bagian_id
                LEFT JOIN desa d ON d.id=j.desa_id
                WHERE j.periode_id=? AND j.pegawai_id=? AND j.hari=? AND j.bagian_id=?
            """, (self.periode_id, pegawai_id, hari, bagian_id)).fetchone()
            if not row:
                messagebox.showinfo("Penugasan", "Penugasan ini sudah tidak tersedia.", parent=parent)
                return
            nama_pegawai, kode, nama_bagian, kegiatan, lokasi, dana, pj_id, desa = row
            pj_row = conn.execute("SELECT nama FROM pegawai WHERE id=?", (pj_id,)).fetchone() if pj_id else None
            pj_nama = pj_row[0] if pj_row else "Belum ditentukan"
            anggota = conn.execute("""
                SELECT p.nama FROM jadwal j JOIN pegawai p ON p.id=j.pegawai_id
                WHERE j.periode_id=? AND j.bagian_id=? AND j.hari=? ORDER BY p.nama
            """, (self.periode_id, bagian_id, hari)).fetchall()
            anggota_text = "\n".join(f"• {x[0]}" for x in anggota) or "Belum ada pegawai"
        finally:
            conn.close()

        win = tk.Toplevel(parent)
        self.enable_native_window_controls(win)
        win.title("Detail Penugasan")
        win.transient(parent)
        win.configure(bg=UI_BG)
        win.resizable(True, True)
        tk.Label(win, text=f"Detail Penugasan — {kode}", bg=UI_BG, fg=UI_TEXT,
                 font=("Segoe UI", 15, "bold")).pack(anchor="w", padx=22, pady=(18, 8))
        body = tk.Frame(win, bg=UI_PANEL, padx=18, pady=14, highlightthickness=1, highlightbackground="#D7DEE8")
        body.pack(fill="both", expand=True, padx=18, pady=4)
        fields = [
            ("Tanggal", f"{hari} {NAMA_BULAN[bulan]} {tahun}"),
            ("Pegawai", nama_pegawai or "Tanpa Nama"),
            ("Kode kegiatan", kode),
            ("Nama kegiatan", nama_bagian.strip() or f"Kegiatan {kode}"),
            ("Deskripsi", kegiatan.strip() or "-"),
            ("Lokasi", lokasi.strip() or desa.strip() or "Belum ditentukan"),
            ("Penanggung jawab", pj_nama or "Belum ditentukan"),
            ("Dana kegiatan", rupiah(int(dana or 0))),
            ("Pegawai pada tanggal ini", anggota_text),
        ]
        for label, value in fields:
            line = tk.Frame(body, bg=UI_PANEL)
            line.pack(fill="x", anchor="w", pady=2)
            tk.Label(line, text=f"{label}:", width=21, anchor="nw", bg=UI_PANEL, fg=UI_MUTED,
                     font=("Segoe UI", 9, "bold")).pack(side="left")
            tk.Label(line, text=str(value), anchor="nw", justify="left", wraplength=460,
                     bg=UI_PANEL, fg=UI_TEXT, font=("Segoe UI", 9)).pack(side="left", fill="x", expand=True)

        bottom = tk.Frame(win, bg=UI_BG)
        bottom.pack(fill="x", padx=18, pady=(8, 16))
        def delete_assignment():
            if not messagebox.askyesno(
                "Konfirmasi Hapus", 
                f"Hapus penugasan {nama_pegawai} pada tanggal {hari}?\n\n"
                "Penugasan pegawai lain pada kegiatan ini tidak akan dihapus.", parent=win
            ):
                return
            win.destroy()
            self.hapus_penugasan_keseluruhan(
                pegawai_id, hari, bagian_id,
                from_overall_calendar=from_overall_calendar
            )
        ttk.Button(bottom, text="Tutup", command=win.destroy).pack(side="right", padx=(8, 0))
        ttk.Button(bottom, text="Hapus Penugasan Ini", command=delete_assignment).pack(side="right")
        self.fit_popup(win, parent, min_width=560, min_height=430, max_width=760, max_height=760)

    def hapus_penugasan_keseluruhan(
        self, pegawai_id, hari, bagian_id, from_overall_calendar=False
    ):
        """Hapus penugasan; aturan pindah kegiatan hanya berlaku di Kalender Kegiatan."""
        # Kalender Keseluruhan boleh menghapus penugasan kegiatan mana pun.
        # Dari Kalender Kegiatan, pengguna wajib berada di kegiatan pemilik jadwal.
        bagian_aktif_id = getattr(self, "bagian_id", None)
        if (
            not from_overall_calendar
            and (bagian_aktif_id is None or int(bagian_aktif_id) != int(bagian_id))
        ):
            conn_info = None
            try:
                conn_info = get_connection()
                target = conn_info.execute(
                    "SELECT kode, COALESCE(nama_bagian, '') FROM bagian WHERE id=?",
                    (bagian_id,)
                ).fetchone()
            finally:
                if conn_info:
                    conn_info.close()

            kode_target = target[0] if target else "?"
            nama_target = ((target[1] if target else "") or "").strip() or f"Kegiatan {kode_target}"
            dialog = tk.Toplevel(self)
            dialog.title("Buka Kegiatan yang Tepat")
            dialog.transient(self)
            dialog.resizable(False, False)
            dialog.configure(bg=UI_BG)

            card = tk.Frame(dialog, bg=UI_PANEL, padx=22, pady=18)
            card.pack(fill="both", expand=True)
            tk.Label(
                card, text="Buka kegiatan yang sesuai dulu",
                bg=UI_PANEL, fg=UI_TEXT, font=("Segoe UI", 13, "bold"), anchor="w"
            ).pack(fill="x", pady=(0, 8))
            tk.Label(
                card,
                text="Penugasan ini milik kegiatan lain. Pilih kegiatan tersebut terlebih dahulu "
                     "agar penugasan tidak terhapus dari tab yang salah.",
                bg=UI_PANEL, fg=UI_MUTED, font=("Segoe UI", 10),
                justify="left", wraplength=410, anchor="w"
            ).pack(fill="x", pady=(0, 14))
            target_box = tk.Frame(card, bg="#E5F4F1", padx=12, pady=10)
            target_box.pack(fill="x", pady=(0, 16))
            tk.Label(
                target_box, text="KEGIATAN YANG DITUJU",
                bg="#E5F4F1", fg="#47716D", font=("Segoe UI", 9, "bold"), anchor="w"
            ).pack(fill="x")
            tk.Label(
                target_box, text=f"{nama_target} ({kode_target})",
                bg="#E5F4F1", fg="#145B55", font=("Segoe UI", 12, "bold"),
                anchor="w", wraplength=410, justify="left"
            ).pack(fill="x", pady=(3, 0))
            actions = tk.Frame(card, bg=UI_PANEL)
            actions.pack(fill="x")

            def buka_kegiatan():
                dialog.destroy()
                try:
                    self.change_bagian(kode_target)
                    self.notebook.select(self.tab_jadwal)
                except Exception as exc:
                    messagebox.showerror(
                        "Tidak Bisa Membuka Kegiatan",
                        f"Tidak berhasil membuka {nama_target}. Pilih kegiatan tersebut "
                        "dari daftar kegiatan di bagian atas.\\n\\n" + str(exc),
                        parent=self
                    )

            ttk.Button(actions, text="Nanti Saja", command=dialog.destroy).pack(side="right", padx=(8, 0))
            ttk.Button(actions, text=f"Buka {nama_target}", command=buka_kegiatan).pack(side="right")
            dialog.update_idletasks()
            width, height = max(470, dialog.winfo_reqwidth()), dialog.winfo_reqheight()
            x = self.winfo_rootx() + max(0, (self.winfo_width() - width) // 2)
            y = self.winfo_rooty() + max(0, (self.winfo_height() - height) // 2)
            dialog.geometry(f"{width}x{height}+{x}+{y}")
            dialog.grab_set()
            dialog.focus_force()
            return

        # Only show loading after any required activity-tab validation has passed.
        self.show_loading("Menghapus penugasan dan menyegarkan kalender...")
        self.update_idletasks()
        conn = None
        berhasil_dihapus = False
        try:
            conn = get_connection()
            row = conn.execute(
                "SELECT id, bagian_id FROM jadwal WHERE periode_id=? AND pegawai_id=? AND hari=?",
                (self.periode_id, pegawai_id, hari)
            ).fetchone()
            if not row or int(row[1]) != int(bagian_id):
                messagebox.showwarning(
                    "Jadwal Berubah",
                    "Penugasan ini sudah berubah atau tidak ditemukan. Kalender akan diperbarui.",
                    parent=getattr(self, "overall_calendar_window", self)
                )
                return

            pj_row = conn.execute(
                "SELECT penanggung_jawab_id, kode, COALESCE(nama_bagian,'') FROM bagian WHERE id=?",
                (bagian_id,)
            ).fetchone()
            pj_id = pj_row[0] if pj_row else None
            if pj_id and int(pegawai_id) == int(pj_id):
                non_pj_count = conn.execute(
                    "SELECT COUNT(DISTINCT pegawai_id) FROM jadwal "
                    "WHERE periode_id=? AND bagian_id=? AND hari=? AND pegawai_id<>?",
                    (self.periode_id, bagian_id, hari, pj_id)
                ).fetchone()[0]
                if non_pj_count > 0:
                    messagebox.showinfo(
                        "Penanggung Jawab",
                        "Masih ada pegawai lain yang bertugas pada kegiatan ini di tanggal tersebut. "
                        "Hapus penugasan pegawai lain terlebih dahulu; jika PJ menjadi satu-satunya "
                        "yang tersisa, penugasannya bisa dihapus.",
                        parent=getattr(self, "overall_calendar_window", self)
                    )
                    return

            conn.execute("DELETE FROM jadwal WHERE id=?", (row[0],))
            conn.commit()
            berhasil_dihapus = True
        except Exception as exc:
            if conn:
                try:
                    conn.rollback()
                except Exception:
                    pass
            messagebox.showerror(
                "Gagal Menghapus",
                f"Penugasan tidak berhasil dihapus. Silakan coba lagi.\\n\\nDetail: {exc}",
                parent=getattr(self, "overall_calendar_window", self)
            )
        finally:
            if conn:
                conn.close()
            self.hide_loading()

        if berhasil_dihapus:
            self.refresh_after_schedule(
                "Menghapus penugasan dan menyegarkan semua kalender..."
            )

    def load_kalender_keseluruhan(self):
        """Kalender dashboard bulanan: semua pegawai dan semua kegiatan, bersifat baca-saja."""
        if (
            getattr(self, "overall_calendar_window", None) is None
            or not self.overall_calendar_window.winfo_exists()
            or not hasattr(self, "overall_calendar_frame")
        ):
            return

        for target_frame in (self.overall_calendar_frame, self.overall_header_frame, self.overall_names_frame):
            for widget in target_frame.winfo_children():
                widget.destroy()

        bulan = NAMA_BULAN.index(self.bulan_var.get())
        tahun = int(self.tahun_var.get())
        jumlah_hari = calendar.monthrange(tahun, bulan)[1]
        hari_libur = load_hari_libur(self.periode_id)

        conn = get_connection()
        kegiatan_rows = conn.execute(
            """
            SELECT id, kode, COALESCE(nama_bagian, ''), COALESCE(kegiatan, '')
            FROM bagian
            WHERE periode_id = ?
            ORDER BY id
            """,
            (self.periode_id,)
        ).fetchall()

        jadwal_rows = conn.execute(
            """
            SELECT j.pegawai_id, j.hari, j.bagian_id, b.kode, COALESCE(b.nama_bagian, ''),
                   COALESCE(b.kegiatan, ''), COALESCE(d.nama, '')
            FROM jadwal j
            JOIN bagian b ON b.id = j.bagian_id
            LEFT JOIN desa d ON d.id = j.desa_id
            WHERE j.periode_id = ?
            """,
            (self.periode_id,)
        ).fetchall()
        conn.close()

        if getattr(self, "current_theme", "light") == "dark":
            palette = [
                "#285B78", "#28634F", "#76552B", "#594477",
                "#743F4C", "#28666A", "#66552C", "#486044",
                "#66513C", "#405A82", "#704454", "#45654C"
            ]
        else:
            palette = [
                "#D7EAFE", "#D9F2E4", "#FCECCB", "#E9DDF8",
                "#F9DDE2", "#D7F0F0", "#F3E8C9", "#E2EED8",
                "#EEE1D1", "#DDE7FA", "#F5DDE7", "#E2F0D9"
            ]
        colors = {row[1]: palette[i % len(palette)] for i, row in enumerate(kegiatan_rows)}
        names = {row[1]: ((row[2] or '').strip() or f"Kegiatan {row[1]}") for row in kegiatan_rows}
        jadwal_map = {
            (row[0], row[1]): {
                "bagian_id": row[2],
                "kode": row[3],
                "nama": (row[4] or '').strip() or f"Kegiatan {row[3]}",
                "deskripsi": (row[5] or '').strip(),
                "desa": (row[6] or '').strip(),
            }
            for row in jadwal_rows
        }

        self.overall_calendar_detail.config(text="")
        self.overall_info_kalender.config(
            text=f"Semua pegawai • {NAMA_BULAN[bulan]} {tahun} • Semua kegiatan"
        )
        if hari_libur:
            self.overall_info_libur.config(
                text=f"Libur manual: {', '.join(str(x) for x in sorted(hari_libur))} • Minggu selalu libur"
            )
        else:
            self.overall_info_libur.config(text="Libur manual: tidak ada • Minggu selalu libur")

        for widget in self.overall_legend_kegiatan.winfo_children():
            widget.destroy()
        tk.Label(
            self.overall_legend_kegiatan,
            text="Keterangan kegiatan:",
            bg=UI_BG,
            font=("Segoe UI", 9, "bold")
        ).pack(side="left", padx=(0, 8))
        for bagian_id, kode, nama, deskripsi in kegiatan_rows:
            tk.Label(
                self.overall_legend_kegiatan,
                text=f"{kode} = {names[kode]}",
                bg=colors[kode],
                relief="solid", bd=1, padx=8, pady=3,
                font=("Segoe UI", 9, "bold")
            ).pack(side="left", padx=3)

        # Pastikan scrollbar legend menghitung ulang lebar A-Z dan mulai
        # dari kiri setiap kali kalender keseluruhan direfresh.
        try:
            self.overall_legend_kegiatan.update_idletasks()
            self.overall_legend_canvas.configure(
                scrollregion=self.overall_legend_canvas.bbox("all")
            )
            self.overall_legend_canvas.xview_moveto(0)
        except Exception:
            pass

        nama_hari = ["Sen", "Sel", "Rab", "Kam", "Jum", "Sab", "Min"]
        for hari in range(1, jumlah_hari + 1):
            tanggal = datetime(tahun, bulan, hari)
            nonkerja = tanggal.weekday() == 6 or hari in hari_libur
            ttk.Label(
                self.overall_header_frame, text=str(hari),
                font=("Segoe UI", 9, "bold"), anchor="center",
                style="Holiday.TLabel" if nonkerja else "TLabel"
            ).grid(row=0, column=hari-1, sticky="nsew", padx=1, pady=(1, 0))
            ttk.Label(
                self.overall_header_frame,
                text="LIBUR" if nonkerja else nama_hari[tanggal.weekday()],
                font=("Segoe UI", 7 if nonkerja else 8, "bold" if nonkerja else "normal"),
                anchor="center",
                style="Holiday.TLabel" if nonkerja else "TLabel"
            ).grid(row=1, column=hari-1, sticky="nsew", padx=1, pady=(0, 1))

        semua_pegawai = sorted(self.pegawai, key=lambda p: str(p[1] or '').lower())
        if not semua_pegawai:
            ttk.Label(
                self.overall_names_frame, text="Belum ada data pegawai aktif.",
                font=("Segoe UI", 10)
            ).grid(row=0, column=0, columnspan=2, pady=20)
            return

        for nomor, pegawai in enumerate(semua_pegawai, start=1):
            row = nomor - 1
            pegawai_id = pegawai[0]
            nama = pegawai[1] or "Tanpa Nama"
            ttk.Label(
                self.overall_names_frame, text=str(nomor), anchor="center", padding=(2, 0)
            ).grid(row=row, column=0, sticky="nsew", padx=1, pady=1)
            ttk.Label(
                self.overall_names_frame, text=nama, anchor="w", padding=(5, 0)
            ).grid(row=row, column=1, sticky="nsew", padx=1, pady=1)

            for hari in range(1, jumlah_hari + 1):
                tanggal = datetime(tahun, bulan, hari)
                nonkerja = tanggal.weekday() == 6 or hari in hari_libur
                detail = jadwal_map.get((pegawai_id, hari))
                kode = detail["kode"] if detail else ""
                warna = colors.get(kode, UI_PANEL) if detail else UI_PANEL

                if nonkerja:
                    cell = tk.Label(
                        self.overall_calendar_frame, text="LIBUR", width=6, height=1,
                        relief="solid", bd=1, bg=UI_TABLE_HEAD, fg=UI_MUTED,
                        highlightthickness=0, font=("Segoe UI", 8)
                    )
                else:
                    cell = tk.Label(
                        self.overall_calendar_frame, text=kode, width=6, height=1,
                        relief="solid", bd=1, bg=warna,
                        fg="#F8FAFC" if getattr(self, "current_theme", "light") == "dark" and kode else UI_TEXT,
                        highlightthickness=0, font=("Segoe UI", 9, "bold") if kode else ("Segoe UI", 9)
                    )

                if detail:
                    desa = detail.get("desa") or "Lokasi belum ditentukan"
                    desa_ringkas = desa if len(desa) <= 10 else desa[:9] + "…"
                    cell.configure(text=f"{kode}\n{desa_ringkas}", justify="center", font=("Segoe UI", 7, "bold"))
                    detail_text = f"{nama} • Tanggal {hari} • {kode} — {detail['nama']} • Lokasi: {desa}"
                    if detail.get("deskripsi"):
                        detail_text += f" • {detail['deskripsi']}"
                    detail_text += " • Klik untuk melihat detail"
                    cell.bind(
                        "<Enter>",
                        lambda event, txt=detail_text: self.overall_calendar_detail.config(text=txt)
                    )
                    cell.bind(
                        "<Button-1>",
                        lambda event, pid=pegawai_id, h=hari, bid=detail["bagian_id"]: self.popup_detail_penugasan(
                            pid, h, bid, parent=self.overall_calendar_window
                        )
                    )
                    cell.configure(cursor="hand2")
                else:
                    cell.bind(
                        "<Enter>",
                        lambda event, n=nama, h=hari: self.overall_calendar_detail.config(
                            text=f"{n} • Tanggal {h} • Belum ada penugasan • Klik untuk menambahkan kegiatan"
                        )
                    )
                    cell.bind(
                        "<Button-1>",
                        lambda event, pid=pegawai_id, h=hari: self.pilih_penugasan_keseluruhan(pid, h)
                    )
                    cell.configure(cursor="hand2")
                cell.bind("<Leave>", lambda event: self.overall_calendar_detail.config(text=""))
                cell.grid(row=row, column=hari-1, sticky="nsew", padx=1, pady=1)

        self.overall_names_frame.grid_columnconfigure(0, minsize=42)
        self.overall_names_frame.grid_columnconfigure(1, minsize=220)
        for col in range(jumlah_hari):
            self.overall_header_frame.grid_columnconfigure(col, minsize=72)
            self.overall_calendar_frame.grid_columnconfigure(col, minsize=72)
        try:
            self.overall_header_frame.update_idletasks()
            self.overall_names_frame.update_idletasks()
            self.overall_calendar_frame.update_idletasks()
            self._refresh_overall_calendar_regions()
        except Exception:
            pass

    def atur_hari_libur(self):
        """Dialog kalender modern untuk memilih hari libur manual."""
        bulan = NAMA_BULAN.index(self.bulan_var.get())
        tahun = int(self.tahun_var.get())
        jumlah_hari = calendar.monthrange(tahun, bulan)[1]
        tersimpan = load_hari_libur(self.periode_id)

        dark = getattr(self, "current_theme", "light") == "dark"
        bg = "#0F172A" if dark else "#F3F6FB"
        panel = "#182438" if dark else "#FFFFFF"
        text = "#E7EEF8" if dark else "#17243A"
        muted = "#9AAAC0" if dark else "#64748B"
        border = "#33445D" if dark else "#DCE4EF"
        accent = "#14B8A6" if dark else "#0F8B83"
        selected_bg = "#155E59" if dark else "#D7F4EF"
        sunday_bg = "#263449" if dark else "#EEF2F7"

        dialog = tk.Toplevel(self)
        self.modernize_window(dialog, "Atur Hari Libur", (620, 520))
        dialog.configure(bg=bg)
        dialog.transient(self)
        dialog.resizable(False, False)
        self.enable_native_window_controls(dialog)
        self.center_window(dialog, self)

        shell = tk.Frame(dialog, bg=bg, padx=22, pady=18)
        shell.pack(fill="both", expand=True)
        heading = tk.Frame(shell, bg=bg)
        heading.pack(fill="x", pady=(0, 14))
        tk.Label(
            heading, text="Atur Hari Libur", bg=bg, fg=text,
            font=("Segoe UI", 17, "bold")
        ).pack(anchor="w")
        tk.Label(
            heading, text=f"{NAMA_BULAN[bulan]} {tahun}  ·  Pilih tanggal libur tambahan",
            bg=bg, fg=muted, font=("Segoe UI", 9)
        ).pack(anchor="w", pady=(3, 0))

        card = tk.Frame(shell, bg=panel, padx=15, pady=14,
                        highlightthickness=1, highlightbackground=border)
        card.pack(fill="both", expand=True)
        summary = tk.Label(card, text="", bg=panel, fg=accent,
                           font=("Segoe UI", 9, "bold"))
        summary.pack(anchor="w", pady=(0, 10))

        grid = tk.Frame(card, bg=panel)
        grid.pack(fill="both", expand=True)
        for col, nama in enumerate(["Sen", "Sel", "Rab", "Kam", "Jum", "Sab", "Min"]):
            tk.Label(grid, text=nama, bg=panel, fg=muted,
                     font=("Segoe UI", 9, "bold"), width=7).grid(
                         row=0, column=col, padx=3, pady=(0, 7))

        vars_hari, day_widgets = {}, {}

        def update_day_style(hari):
            var, widget = vars_hari[hari], day_widgets[hari]
            weekday = datetime(tahun, bulan, hari).weekday()
            if weekday == 6:
                widget.configure(bg=sunday_bg, fg=muted, state="disabled",
                                 activebackground=sunday_bg, activeforeground=muted)
            elif var.get():
                widget.configure(bg=selected_bg, fg=accent, activebackground=selected_bg,
                                 activeforeground=accent, relief="solid",
                                 highlightbackground=accent)
            else:
                widget.configure(bg=bg, fg=text, activebackground=selected_bg,
                                 activeforeground=accent, relief="flat",
                                 highlightbackground=border)

        def update_summary():
            count = sum(1 for h, v in vars_hari.items()
                        if v.get() and datetime(tahun, bulan, h).weekday() != 6)
            summary.configure(text=f"{count} tanggal dipilih  ·  Minggu otomatis libur")

        offset = datetime(tahun, bulan, 1).weekday()
        for hari in range(1, jumlah_hari + 1):
            tanggal = datetime(tahun, bulan, hari)
            row = 1 + (offset + hari - 1) // 7
            col = tanggal.weekday()
            var = tk.BooleanVar(value=hari in tersimpan)
            vars_hari[hari] = var
            cb = tk.Checkbutton(
                grid, text=str(hari), variable=var, indicatoron=False,
                width=5, height=2, font=("Segoe UI", 10, "bold"),
                bd=0, relief="flat", highlightthickness=1,
                command=lambda h=hari: (update_day_style(h), update_summary())
            )
            cb.grid(row=row, column=col, padx=3, pady=3, sticky="nsew")
            day_widgets[hari] = cb
            update_day_style(hari)

        for col in range(7):
            grid.grid_columnconfigure(col, weight=1, uniform="day")

        tk.Label(shell,
                 text="Minggu ditandai otomatis sebagai hari libur dan tidak dapat diubah.",
                 bg=bg, fg=muted, font=("Segoe UI", 8)).pack(anchor="w", pady=(10, 12))
        actions = tk.Frame(shell, bg=bg)
        actions.pack(fill="x")

        def pilih_semua():
            for h, var in vars_hari.items():
                if datetime(tahun, bulan, h).weekday() != 6:
                    var.set(True)
                    update_day_style(h)
            update_summary()

        def kosongkan():
            for h, var in vars_hari.items():
                var.set(False)
                update_day_style(h)
            update_summary()

        ttk.Button(actions, text="Pilih Semua", command=pilih_semua,
                   style="Soft.TButton").pack(side="left")
        ttk.Button(actions, text="Kosongkan", command=kosongkan,
                   style="Soft.TButton").pack(side="left", padx=7)

        def simpan():
            baru = {h for h, var in vars_hari.items()
                    if var.get() and datetime(tahun, bulan, h).weekday() != 6}
            conn = None
            try:
                simpan_hari_libur(self.periode_id, baru)
                nonkerja = set(baru)
                nonkerja.update(h for h in range(1, jumlah_hari + 1)
                                if datetime(tahun, bulan, h).weekday() == 6)
                conn = get_connection()
                for hari in sorted(nonkerja):
                    conn.execute("DELETE FROM jadwal WHERE periode_id = ? AND hari = ?",
                                 (self.periode_id, hari))
                conn.commit()
            except Exception:
                if conn:
                    conn.rollback()
                messagebox.showerror("Gagal Menyimpan",
                                     "Pengaturan hari libur tidak dapat disimpan. Silakan coba lagi.",
                                     parent=dialog)
                return
            finally:
                if conn:
                    conn.close()

            dialog.destroy()
            self.load_kalender()
            self.refresh_data_pegawai()
            if (getattr(self, "overall_calendar_window", None) is not None
                    and self.overall_calendar_window.winfo_exists()):
                self.load_kalender_keseluruhan()
            messagebox.showinfo("Hari Libur",
                                f"{len(baru)} tanggal libur tambahan berhasil disimpan.",
                                parent=self)

        ttk.Button(actions, text="Batal", command=dialog.destroy,
                   style="Soft.TButton").pack(side="right", padx=(0, 8))
        ttk.Button(actions, text="Simpan Perubahan", command=simpan,
                   style="Primary.TButton").pack(side="right")
        update_summary()
        self.fit_popup(dialog, self, min_width=620, min_height=520,
                       max_width=760, max_height=700)
        dialog.grab_set()
        dialog.focus_force()

    # CELL KALENDER
    # =====================================================

    def create_calendar_cell(
        self,
        row,
        hari,
        pegawai_id
    ):

        # Cek jadwal
        conn = get_connection()
        cursor = conn.cursor()

        jadwal = cursor.execute("""
            SELECT
                j.bagian_id,
                b.kode
            FROM jadwal j
            JOIN bagian b
                ON b.id = j.bagian_id
            WHERE j.periode_id = ?
            AND j.pegawai_id = ?
            AND j.hari = ?
        """, (
            self.periode_id,
            pegawai_id,
            hari
        )).fetchone()

        conn.close()

        kode = jadwal[1] if jadwal else ""

        warna = "#DDE7F0"

        if kode == "A":
            warna = "#B8D4EF"

        elif kode == "B":
            warna = "#C7E6C5"

        elif kode == "C":
            warna = "#F3D6A5"

        button = tk.Button(
            self.calendar_frame,
            text=kode,
            width=3,
            height=1,
            bg=warna,
            relief="solid",
            bd=1,
            command=lambda
                p=pegawai_id,
                h=hari:
                self.toggle_jadwal(p, h)
        )

        button.grid(
            row=row,
            column=hari + 1,
            padx=1,
            pady=1
        )

    # =====================================================
    # TOGGLE JADWAL
    # =====================================================

    def pilih_desa_dialog(self, hari, desa_default=None):
        """Meminta desa untuk tanggal penugasan melalui dropdown master desa."""
        desa = load_desa_aktif()
        if not desa:
            messagebox.showwarning(
                "Data Lokasi Kosong",
                "Belum ada lokasi aktif di database. Tambahkan lokasi terlebih dahulu pada menu Data Keseluruhan → Data Lokasi.",
                parent=self
            )
            return None

        win = tk.Toplevel(self)
        self.enable_native_window_controls(win)
        win.title(f"Pilih Lokasi — Tanggal {hari}")
        win.geometry("460x210")
        win.transient(self)
        self.enable_native_window_controls(win)
        win.grab_set()
        self.enable_native_window_controls(win)
        self.center_window(win, self)
        tk.Label(win, text=f"Lokasi untuk penugasan tanggal {hari}", font=("Segoe UI", 12, "bold")).pack(anchor="w", padx=20, pady=(18, 8))
        tk.Label(win, text="Setiap tanggal kegiatan wajib memiliki satu lokasi.", fg=UI_MUTED).pack(anchor="w", padx=20, pady=(0, 10))
        var = tk.StringVar()
        values = [x[1] for x in desa]
        combo = ttk.Combobox(win, textvariable=var, values=values, state="readonly", width=42)
        combo.pack(padx=20, pady=5)
        if desa_default:
            row = next((x for x in desa if x[0] == desa_default), None)
            if row:
                var.set(row[1])
        elif values:
            var.set(values[0])
        result = {"id": None}
        def simpan():
            pilihan = var.get().strip()
            row = next((x for x in desa if x[1] == pilihan), None)
            if not row:
                messagebox.showwarning("Lokasi", "Pilih lokasi terlebih dahulu.", parent=win)
                return
            result["id"] = row[0]
            win.destroy()
        # Saat pertama memilih desa, pengguna hanya perlu memilih Simpan atau Batal.
        # Penghapusan jadwal dilakukan langsung dengan klik kedua pada sel kalender.
        bottom = tk.Frame(win)
        bottom.pack(fill="x", padx=20, pady=15)
        ttk.Button(bottom, text="Batal", command=win.destroy).pack(side="right", padx=(8,0))
        ttk.Button(bottom, text="Simpan", command=simpan, style="Primary.TButton").pack(side="right")
        self.fit_popup(
            win, self,
            min_width=460, min_height=230, max_width=680, max_height=420
        )
        win.wait_window()
        return result["id"]

    def toggle_jadwal(self, pegawai_id, hari):
        self.show_loading("Memproses penugasan...")
        self.update_idletasks()
        bulan = NAMA_BULAN.index(self.bulan_var.get())
        tahun = int(self.tahun_var.get())
        if datetime(tahun, bulan, hari).weekday() == 6 or hari in load_hari_libur(self.periode_id):
            self.hide_loading()
            return

        conn = get_connection()
        try:
            existing = conn.execute("""
                SELECT j.id, j.bagian_id, j.desa_id, COALESCE(d.nama, '')
                FROM jadwal j LEFT JOIN desa d ON d.id = j.desa_id
                WHERE j.periode_id = ? AND j.pegawai_id = ? AND j.hari = ?
            """, (self.periode_id, pegawai_id, hari)).fetchone()

            if existing and existing[1] != self.bagian_id:
                bagian_lain = conn.execute("SELECT kode, COALESCE(nama_bagian, '') FROM bagian WHERE id = ?", (existing[1],)).fetchone()
                kode_lain = bagian_lain[0] if bagian_lain else "?"
                nama_lain = (bagian_lain[1] if bagian_lain else '').strip() or f"Kegiatan {kode_lain}"
                messagebox.showwarning("Jadwal Bentrok", f"Pegawai ini sudah bertugas pada Kegiatan {kode_lain} — {nama_lain} pada tanggal {hari}.", parent=self)
                self.hide_loading()
                return

            if existing and existing[1] == self.bagian_id:
                pj_id = getattr(self, "penanggung_jawab_id", None)
                if pegawai_id == pj_id:
                    non_pj_count = conn.execute("SELECT COUNT(*) FROM jadwal WHERE periode_id=? AND bagian_id=? AND hari=? AND pegawai_id<>?", (self.periode_id, self.bagian_id, hari, pj_id)).fetchone()[0] if pj_id else 0
                    if non_pj_count > 0:
                        messagebox.showinfo("Penanggung Jawab", "PJ otomatis mengikuti tanggal kegiatan dan tidak dapat dilepas selama masih ada pegawai lain pada tanggal tersebut.", parent=self)
                        self.hide_loading()
                        return
                conn.execute("DELETE FROM jadwal WHERE id = ?", (existing[0],))
                conn.commit()

                # Klik kedua pada pegawai yang sudah terpilih = hapus penugasan.
                # Langsung refresh seluruh tampilan agar perubahan terlihat tanpa
                # harus menekan tombol Refresh manual.
                self.hide_loading()
                self.refresh_after_schedule(
                    "Menghapus penugasan dan menyegarkan kalender..."
                )
                return

            # Baru pada jalur penambahan penugasan kita wajib memastikan PJ sudah ada.
            if not self._cek_pj_kegiatan(self.bagian_id, parent=self):
                self.hide_loading()
                return

            existing_activity_day = conn.execute("SELECT desa_id FROM jadwal WHERE periode_id=? AND bagian_id=? AND hari=? LIMIT 1", (self.periode_id, self.bagian_id, hari)).fetchone()
            self.hide_loading()
            desa_id = existing_activity_day[0] if existing_activity_day and existing_activity_day[0] else self.pilih_desa_dialog(hari)
            if not desa_id:
                self.hide_loading()
                return

            # Satu desa tidak boleh dipakai dua kegiatan pada tanggal yang sama.
            konflik_desa = conn.execute("""
                SELECT b.kode, COALESCE(b.nama_bagian, '')
                FROM jadwal j JOIN bagian b ON b.id=j.bagian_id
                WHERE j.periode_id=? AND j.hari=? AND j.desa_id=? AND j.bagian_id<>?
                LIMIT 1
            """, (self.periode_id, hari, desa_id, self.bagian_id)).fetchone()
            if konflik_desa:
                kode, nama = konflik_desa
                nama = nama.strip() or f"Kegiatan {kode}"
                desa_row = load_desa_by_id(desa_id)
                nama_desa = desa_row[1] if desa_row else "desa tersebut"
                messagebox.showwarning("Desa Bentrok", f"{nama_desa} sudah digunakan oleh Kegiatan {kode} — {nama} pada tanggal {hari}.", parent=self)
                self.hide_loading()
                return

            conn.execute("INSERT INTO jadwal (periode_id, bagian_id, pegawai_id, hari, desa_id) VALUES (?, ?, ?, ?, ?)", (self.periode_id, self.bagian_id, pegawai_id, hari, desa_id))

            # PJ otomatis turun pada tanggal kegiatan yang dipilih.
            pj_id = getattr(self, "penanggung_jawab_id", None)
            if pj_id and pj_id != pegawai_id:
                pj_existing = conn.execute("SELECT id FROM jadwal WHERE periode_id=? AND bagian_id=? AND pegawai_id=? AND hari=?", (self.periode_id, self.bagian_id, pj_id, hari)).fetchone()
                if not pj_existing:
                    conn.execute("INSERT INTO jadwal (periode_id, bagian_id, pegawai_id, hari, desa_id) VALUES (?, ?, ?, ?, ?)", (self.periode_id, self.bagian_id, pj_id, hari, desa_id))

            conn.commit()
        except sqlite3.IntegrityError as exc:
            conn.rollback()
            self.hide_loading()
            messagebox.showwarning("Jadwal", "Penugasan tidak dapat disimpan. Periksa kembali pilihan kegiatan, desa, dan jadwal pegawai.", parent=self)
        except Exception as exc:
            conn.rollback()
            self.hide_loading()
            messagebox.showerror("Jadwal", "Penugasan tidak dapat disimpan. Silakan coba lagi.", parent=self)
        finally:
            conn.close()

        self.hide_loading()
        self.refresh_after_schedule(
            "Menyimpan penugasan dan menyegarkan kalender..."
        )

    def build_data_kegiatan(self):
        """Tabel rekap penugasan khusus kegiatan aktif pada bulan aktif."""
        for w in self.tab_data_kegiatan.winfo_children():
            w.destroy()

        # Seluruh halaman Data Kegiatan dapat digeser dua arah. Tabel-tabel di
        # dalamnya tetap memiliki scrollbar masing-masing untuk data yang panjang.
        container = tk.Frame(self.tab_data_kegiatan, bg=UI_BG)
        container.pack(fill="both", expand=True)
        page_canvas = tk.Canvas(container, bg=UI_BG, highlightthickness=0)
        page_v = ttk.Scrollbar(container, orient="vertical", command=page_canvas.yview)
        page_canvas.configure(yscrollcommand=page_v.set)
        page_canvas.grid(row=0, column=0, sticky="nsew")
        page_v.grid(row=0, column=1, sticky="ns")
        container.grid_rowconfigure(0, weight=1)
        container.grid_columnconfigure(0, weight=1)
        content = tk.Frame(page_canvas, bg=UI_BG)
        page_window = page_canvas.create_window((0, 0), window=content, anchor="nw")
        self.data_kegiatan_page_canvas = page_canvas

        def _sync_data_page_scrollregion(event=None):
            try:
                content.update_idletasks()
                page_canvas.configure(
                    scrollregion=(0, 0, max(page_canvas.winfo_width(), 1),
                                  max(content.winfo_reqheight(), page_canvas.winfo_height(), 1))
                )
            except tk.TclError:
                pass

        content.bind("<Configure>", _sync_data_page_scrollregion, add="+")
        page_canvas.bind(
            "<Configure>",
            lambda e: (
                page_canvas.itemconfigure(page_window, width=max(e.width, 1)),
                _sync_data_page_scrollregion()
            ),
            add="+"
        )

        tk.Label(
            content, text="DATA KEGIATAN",
            bg=UI_BG, font=("Segoe UI Semibold", 14)
        ).pack(anchor="w", padx=20, pady=(15, 3))

        self.data_kegiatan_info = tk.Label(
            content, text="", bg=UI_BG, fg=UI_MUTED
        )
        self.data_kegiatan_info.pack(anchor="w", padx=20, pady=(0, 8))

        frame = tk.Frame(
            content,
            bg=UI_PANEL,
            highlightbackground=UI_BORDER,
            highlightthickness=1
        )
        frame.pack(fill="both", expand=True, padx=20, pady=10)
        # Data Kegiatan hanya menampilkan JUMLAH TURUN pada kegiatan ini.
        # Tanggal/hari penugasan sudah ditampilkan pada Kalender Kegiatan,
        # sehingga tidak perlu diduplikasi di tabel data.
        cols = ("no", "nama", "nip", "pangkat", "golongan", "jabatan", "tarif", "pj", "jumlah")
        self.table_data_kegiatan = ttk.Treeview(frame, columns=cols, show="headings", selectmode="browse")
        heads = {
            "no":"No.", "nama":"Nama", "nip":"NIP", "pangkat":"Pangkat", "golongan":"Golongan",
            "jabatan":"Jabatan", "tarif":"Tarif", "pj":"PJ",
            "jumlah":"Turun"
        }
        widths = {"no":58,"nama":280,"nip":185,"pangkat":170,"golongan":105,"jabatan":230,"tarif":125,"pj":70,"jumlah":120}
        for c in cols:
            self.table_data_kegiatan.heading(c, text=heads[c])
            self.table_data_kegiatan.column(c, width=widths[c], anchor="center")
        self.table_data_kegiatan.column("nama", anchor="w")
        self.table_data_kegiatan.column("jabatan", anchor="w")
        frame.grid_rowconfigure(0, weight=1)
        frame.grid_columnconfigure(0, weight=1)
        scroll = ttk.Scrollbar(
            frame, orient="vertical", command=self.table_data_kegiatan.yview
        )
        scroll_x = ttk.Scrollbar(
            frame, orient="horizontal", command=self.table_data_kegiatan.xview
        )
        self.table_data_kegiatan.grid(row=0, column=0, sticky="nsew")
        scroll.grid(row=0, column=1, sticky="ns")
        scroll_x.grid(row=1, column=0, sticky="ew")
        self.table_data_kegiatan.configure(
            yscrollcommand=scroll.set, xscrollcommand=scroll_x.set
        )
        self.modernize_tree(self.table_data_kegiatan)

        # Ringkasan desa yang dipakai pada kegiatan ini.
        desa_title = tk.Label(
            content,
            text="REKAP DESA PADA KEGIATAN INI",
            bg=UI_BG,
            font=("Segoe UI", 11, "bold")
        )
        desa_title.pack(anchor="w", padx=20, pady=(4, 4))

        desa_frame = tk.Frame(
            content,
            bg=UI_PANEL,
            highlightbackground=UI_BORDER,
            highlightthickness=1
        )
        desa_frame.pack(fill="both", expand=False, padx=20, pady=(0, 10))
        desa_frame.grid_rowconfigure(0, weight=1)
        desa_frame.grid_columnconfigure(0, weight=1)

        desa_cols = ("no", "desa", "turun", "tanggal", "kegiatan")
        self.table_desa_kegiatan = ttk.Treeview(
            desa_frame, columns=desa_cols, show="headings", height=6
        )
        self.table_desa_kegiatan.heading("no", text="No.")
        self.table_desa_kegiatan.heading("desa", text="Lokasi")
        self.table_desa_kegiatan.heading("turun", text="Dipakai")
        self.table_desa_kegiatan.heading("tanggal", text="Tanggal")
        self.table_desa_kegiatan.heading("kegiatan", text="Kegiatan")
        self.table_desa_kegiatan.column("no", width=58, anchor="center", stretch=False)
        self.table_desa_kegiatan.column("desa", width=280, anchor="w")
        self.table_desa_kegiatan.column("turun", width=100, anchor="center")
        self.table_desa_kegiatan.column("tanggal", width=180, anchor="w")
        self.table_desa_kegiatan.column("kegiatan", width=360, anchor="w")
        self.table_desa_kegiatan.grid(row=0, column=0, sticky="nsew")

        desa_vscroll = ttk.Scrollbar(
            desa_frame, orient="vertical", command=self.table_desa_kegiatan.yview
        )
        desa_vscroll.grid(row=0, column=1, sticky="ns")
        desa_hscroll = ttk.Scrollbar(
            desa_frame, orient="horizontal", command=self.table_desa_kegiatan.xview
        )
        desa_hscroll.grid(row=1, column=0, sticky="ew")
        self.table_desa_kegiatan.configure(
            yscrollcommand=desa_vscroll.set,
            xscrollcommand=desa_hscroll.set
        )
        self.modernize_tree(self.table_desa_kegiatan)

        toolbar = tk.Frame(content, bg=UI_BG)
        toolbar.pack(fill="x", padx=20, pady=(0, 12))
        ttk.Button(
            toolbar,
            text="🗑 Hapus Keanggotaan Permanen",
            command=self.hapus_keanggotaan_kegiatan_terpilih
        ).pack(side="left")
        tk.Label(
            toolbar,
            text="Hapus pegawai dari kegiatan ini pada bulan aktif beserta jadwalnya.",
            bg=UI_BG, fg=UI_MUTED, font=("Segoe UI", 9)
        ).pack(side="left", padx=10)
        self.table_data_kegiatan.bind(
            "<Delete>", lambda event: self.hapus_keanggotaan_kegiatan_terpilih()
        )
        self.refresh_data_kegiatan()

    def hapus_keanggotaan_kegiatan_terpilih(self):
        """Menghapus keanggotaan pegawai secara permanen dari kegiatan/periode aktif."""
        if not hasattr(self, "table_data_kegiatan") or self.table_data_kegiatan is None:
            return
        pilihan = self.table_data_kegiatan.selection()
        if not pilihan:
            messagebox.showinfo("Keanggotaan Kegiatan", "Pilih pegawai yang ingin dihapus dari kegiatan terlebih dahulu.", parent=self)
            return
        try:
            pegawai_id = int(pilihan[0])
        except (ValueError, TypeError):
            return

        if pegawai_id == getattr(self, "penanggung_jawab_id", None):
            messagebox.showwarning(
                "Penanggung Jawab",
                "Pegawai ini adalah Penanggung Jawab kegiatan.\n\nGanti Penanggung Jawab terlebih dahulu sebelum menghapus keanggotaannya.",
                parent=self
            )
            return

        conn = get_connection()
        try:
            row = conn.execute("SELECT nama FROM pegawai WHERE id = ?", (pegawai_id,)).fetchone()
            if not row:
                return
            nama = row[0] or "Tanpa Nama"
            jumlah_jadwal = conn.execute(
                "SELECT COUNT(*) FROM jadwal WHERE periode_id = ? AND bagian_id = ? AND pegawai_id = ?",
                (self.periode_id, self.bagian_id, pegawai_id)
            ).fetchone()[0]
        finally:
            conn.close()

        if not messagebox.askyesno(
            "Hapus Keanggotaan Permanen",
            f"Hapus {nama} dari Kegiatan {self.bagian_aktif} pada {NAMA_BULAN[self.bulan]} {self.tahun}?\n\n"
            f"• Keanggotaan kegiatan dihapus\n• {jumlah_jadwal} jadwal kegiatan ikut dihapus\n\n"
            "Pegawai tetap ada di Data Pegawai dan bisa ditambahkan kembali kapan saja.\n\nLanjutkan?",
            parent=self
        ):
            return

        conn = get_connection()
        try:
            conn.execute(
                "DELETE FROM jadwal WHERE periode_id = ? AND bagian_id = ? AND pegawai_id = ?",
                (self.periode_id, self.bagian_id, pegawai_id)
            )
            conn.execute(
                "DELETE FROM anggota_bagian WHERE bagian_id = ? AND pegawai_id = ?",
                (self.bagian_id, pegawai_id)
            )
            conn.commit()
        except Exception as exc:
            conn.rollback()
            messagebox.showerror("Gagal Menghapus", "Keanggotaan tidak dapat dihapus. Silakan coba lagi.", parent=self)
            return
        finally:
            conn.close()

        self.selected_ids.discard(pegawai_id)
        self.refresh_selection_table()
        self.refresh_data_kegiatan()
        self.refresh_data_pegawai()
        self.load_kalender()
        self.load_kalender_keseluruhan()

    def refresh_data_kegiatan(self):
        if not hasattr(self, "table_data_kegiatan") or self.table_data_kegiatan is None:
            return
        try:
            for item in self.table_data_kegiatan.get_children():
                self.table_data_kegiatan.delete(item)
        except tk.TclError:
            return

        data = load_bagian(self.periode_id, self.bagian_aktif)
        if not data:
            return
        bagian_id, kode, kegiatan, lokasi, dana, pj_id = data
        nama = getattr(self, "nama_bagian_var", tk.StringVar(value=f"Kegiatan {kode}")).get().strip() or f"Kegiatan {kode}"
        if hasattr(self, "data_kegiatan_info"):
            self.data_kegiatan_info.config(
                text=f"{nama} • {NAMA_BULAN[self.bulan]} {self.tahun} • hanya menghitung penugasan pada kegiatan ini"
            )

        conn=get_connection()
        try:
            anggota=conn.execute("""
                SELECT p.id,p.nama,p.nip,p.pangkat,p.golongan_asli,p.jabatan,p.tarif
                FROM anggota_bagian ab JOIN pegawai p ON p.id=ab.pegawai_id
                WHERE ab.bagian_id=? ORDER BY p.nama
            """, (bagian_id,)).fetchall()
            for nomor, p in enumerate(anggota, start=1):
                # Satu-satunya angka rekap di Data Kegiatan adalah
                # jumlah turun pegawai pada kegiatan ini di bulan aktif.
                jumlah_turun = conn.execute(
                    "SELECT COUNT(*) FROM jadwal WHERE periode_id=? AND bagian_id=? AND pegawai_id=?",
                    (self.periode_id, bagian_id, p[0])
                ).fetchone()[0]
                self.table_data_kegiatan.insert(
                    "", "end", iid=str(p[0]),
                    values=(
                        nomor, p[1],p[2],p[3],p[4],p[5],rupiah(p[6]),
                        "Ya" if p[0]==pj_id else "",
                        f"{jumlah_turun} kali"
                    )
                )

            if hasattr(self, "table_desa_kegiatan"):
                for item in self.table_desa_kegiatan.get_children():
                    self.table_desa_kegiatan.delete(item)

                desa_rekap = conn.execute("""
                    SELECT d.nama,
                           COUNT(DISTINCT j.hari) AS jumlah_dipakai,
                           GROUP_CONCAT(DISTINCT j.hari) AS tanggal,
                           GROUP_CONCAT(DISTINCT
                               CASE
                                   WHEN COALESCE(b.nama_bagian,'') <> '' THEN
                                       b.kode || ' - ' || b.nama_bagian
                                   ELSE
                                       'Kegiatan ' || b.kode
                               END
                           ) AS kegiatan
                    FROM jadwal j
                    JOIN desa d ON d.id=j.desa_id
                    JOIN bagian b ON b.id=j.bagian_id
                    WHERE j.periode_id=? AND j.bagian_id=? AND j.desa_id IS NOT NULL
                    GROUP BY d.id, d.nama
                    ORDER BY d.nama COLLATE NOCASE
                """, (self.periode_id, bagian_id)).fetchall()

                for nomor_desa, (nama_desa, jumlah_dipakai, tanggal, kegiatan_desa) in enumerate(desa_rekap, start=1):
                    try:
                        tanggal_list = sorted({int(x) for x in (tanggal or "").split(",") if str(x).strip()})
                        tanggal_text = ", ".join(str(x) for x in tanggal_list) if tanggal_list else "-"
                    except Exception:
                        tanggal_text = tanggal or "-"
                    self.table_desa_kegiatan.insert(
                        "", "end",
                        values=(
                            nomor_desa,
                            nama_desa,
                            f"{jumlah_dipakai} kali",
                            tanggal_text,
                            kegiatan_desa or f"Kegiatan {kode}"
                        )
                    )
        finally:
            conn.close()

    def build_jadwal(self):
        """Kalender Kegiatan dengan header tanggal tetap saat scroll vertikal."""
        style = ttk.Style(self)
        style.configure("Holiday.TLabel", foreground="#888888", background="#ECECEC")

        root = tk.Frame(self.tab_jadwal, bg=UI_BG)
        root.pack(fill="both", expand=True)
        root.grid_rowconfigure(1, weight=1)
        root.grid_columnconfigure(0, weight=1)

        header = tk.Frame(root, bg=UI_BG)
        header.grid(row=0, column=0, sticky="ew", padx=16, pady=(8, 5))

        # Toolbar ringkas: informasi bulan sudah tersedia di bagian atas aplikasi.
        kalender_tools = tk.Frame(
            header, bg=UI_PANEL, padx=10, pady=7,
            highlightthickness=1, highlightbackground=UI_BORDER
        )
        kalender_tools.pack(fill="x")
        ttk.Button(
            kalender_tools, text="▦  Atur Hari Libur",
            command=self.atur_hari_libur, style="Soft.TButton"
        ).pack(side="left")
        self.info_libur = tk.Label(
            kalender_tools, text="", bg=UI_PANEL, fg=UI_MUTED,
            font=("Segoe UI", 9)
        )
        self.info_libur.pack(side="left", padx=(12, 0))

        # Keterangan warna disembunyikan agar tidak memenuhi header.
        self.legend_kegiatan = tk.Frame(header, bg=UI_BG)
        self.calendar_legend_button = ttk.Button(
            header, text="ⓘ  Keterangan kegiatan",
            style="Soft.TButton"
        )
        self.calendar_legend_button.pack(anchor="w", pady=(5, 0))
        self.calendar_detail = tk.Label(
            header, text="", bg=UI_BG, fg=UI_MUTED, anchor="w",
            font=("Segoe UI", 9)
        )
        # Detail dipindahkan ke tooltip pada sel kalender, bukan baris permanen.

        # Ringkasan target dibuat sebagai panel ringkas yang bisa dibuka/tutup.
        self.calendar_summary_card = tk.Frame(
            header, bg=UI_PANEL, padx=10, pady=7,
            highlightthickness=1, highlightbackground=UI_BORDER
        )
        self.calendar_summary_card.pack(fill="x", pady=(6, 0))
        summary_head = tk.Frame(self.calendar_summary_card, bg=UI_PANEL)
        summary_head.pack(fill="x")
        self.calendar_summary_toggle = ttk.Button(
            summary_head, text="▸  Ringkasan Komposisi & Realisasi",
            style="Soft.TButton", command=self.toggle_calendar_summary
        )
        self.calendar_summary_toggle.pack(side="left")
        self.calendar_summary_status = tk.Label(
            summary_head, text="", bg=UI_PANEL, fg=UI_MUTED,
            font=("Segoe UI", 9)
        )
        self.calendar_summary_status.pack(side="right")
        # Ringkasan dibuat dalam tabel agar angka target dan realisasi sejajar.
        summary_body = tk.Frame(self.calendar_summary_card, bg=UI_PANEL)
        summary_body.pack(fill="x", pady=(7, 0))
        summary_columns = ("golongan", "target", "tarif", "subtotal", "realisasi", "selisih", "status")
        # Baris ringkasan diberi tinggi khusus agar teks tidak terpotong
        # pada Windows dengan DPI/scaling 125% atau lebih tinggi.
        summary_style = ttk.Style(self)
        summary_style.configure(
            "CalendarSummary.Treeview",
            rowheight=42,
            font=("Segoe UI", 10)
        )
        summary_style.configure(
            "CalendarSummary.Treeview.Heading",
            font=("Segoe UI", 9, "bold"),
            padding=(8, 10)
        )
        self.calendar_summary = ttk.Treeview(
            summary_body, columns=summary_columns, show="headings",
            height=3, style="CalendarSummary.Treeview"
        )
        summary_headers = {
            "golongan": "Golongan", "target": "Target Turun",
            "tarif": "Tarif", "subtotal": "Total Target",
            "realisasi": "Realisasi", "selisih": "Selisih", "status": "Status"
        }
        summary_widths = {
            "golongan": 190, "target": 105, "tarif": 125,
            "subtotal": 135, "realisasi": 105, "selisih": 105, "status": 115
        }
        for col in summary_columns:
            self.calendar_summary.heading(col, text=summary_headers[col])
            self.calendar_summary.column(
                col, width=summary_widths[col], minwidth=75,
                anchor="center", stretch=(col == "golongan")
            )
        self.calendar_summary.pack(fill="x", expand=True)

        self.calendar_summary_totals = tk.Label(
            summary_body, text="", bg=UI_PANEL, fg=UI_TEXT,
            anchor="w", justify="left", font=("Segoe UI", 9, "bold")
        )
        self.calendar_summary_totals.pack(fill="x", pady=(6, 0))
        self.calendar_summary_villages = tk.Label(
            summary_body, text="", bg=UI_PANEL, fg=UI_MUTED,
            anchor="w", justify="left", font=("Segoe UI", 9),
            wraplength=1400
        )
        self.calendar_summary_villages.pack(fill="x", pady=(3, 0))
        self._calendar_summary_expanded = False

        viewport = tk.Frame(root, bg=UI_PANEL, highlightbackground=UI_BORDER, highlightthickness=1)
        viewport.grid(row=1, column=0, sticky="nsew", padx=14, pady=(2, 12))
        viewport.grid_columnconfigure(1, weight=1)
        viewport.grid_rowconfigure(1, weight=1)

        # Header pegawai dan header tanggal berada di baris terpisah dari badan tabel.
        # Dengan begitu header tidak ikut tergulung vertikal.
        names_header = tk.Frame(viewport, bg=UI_TABLE_HEAD, width=310, height=48)
        names_header.grid(row=0, column=0, sticky="nsew")
        names_header.grid_propagate(False)
        names_header.grid_columnconfigure(0, minsize=42, weight=0)
        names_header.grid_columnconfigure(1, minsize=245, weight=1)

        date_header_canvas = tk.Canvas(viewport, bg=UI_TABLE_HEAD, highlightthickness=0,
                                       height=48, xscrollincrement=20)
        date_header_canvas.grid(row=0, column=1, sticky="ew")
        date_header_frame = tk.Frame(date_header_canvas, bg=UI_TABLE_HEAD)
        date_header_window = date_header_canvas.create_window((0, 0), window=date_header_frame, anchor="nw")

        left_canvas = tk.Canvas(viewport, bg=UI_PANEL, highlightthickness=0, width=310, yscrollincrement=20)
        left_canvas.grid(row=1, column=0, sticky="nsew")
        left_frame = tk.Frame(left_canvas, bg=UI_PANEL)
        left_window = left_canvas.create_window((0, 0), window=left_frame, anchor="nw")
        left_frame.grid_columnconfigure(0, minsize=42, weight=0)
        left_frame.grid_columnconfigure(1, minsize=245, weight=1)

        right_canvas = tk.Canvas(viewport, bg=UI_PANEL, highlightthickness=0,
                                 xscrollincrement=20, yscrollincrement=20)
        right_canvas.grid(row=1, column=1, sticky="nsew")
        right_frame = tk.Frame(right_canvas, bg=UI_PANEL)
        right_window = right_canvas.create_window((0, 0), window=right_frame, anchor="nw")

        vscroll = ttk.Scrollbar(viewport, orient="vertical",
                                command=lambda *args: (left_canvas.yview(*args), right_canvas.yview(*args)))
        vscroll.grid(row=1, column=2, sticky="ns")
        hscroll = ttk.Scrollbar(viewport, orient="horizontal")
        hscroll.grid(row=2, column=1, sticky="ew")

        def horizontal_scroll(*args):
            date_header_canvas.xview(*args)
            right_canvas.xview(*args)

        hscroll.configure(command=horizontal_scroll)
        date_header_canvas.configure(xscrollcommand=hscroll.set)
        right_canvas.configure(xscrollcommand=hscroll.set, yscrollcommand=vscroll.set)
        left_canvas.configure(yscrollcommand=vscroll.set)

        self.calendar_names_header_frame = names_header
        self.calendar_header_frame = date_header_frame
        self.calendar_names_canvas = left_canvas
        self.calendar_canvas = right_canvas
        self.calendar_names_frame = left_frame
        self.calendar_frame = right_frame

        def refresh_regions(event=None):
            try:
                for frame in (names_header, date_header_frame, left_frame, right_frame):
                    frame.update_idletasks()
                date_width = max(date_header_frame.winfo_reqwidth(), date_header_canvas.winfo_width())
                body_width = max(right_frame.winfo_reqwidth(), right_canvas.winfo_width())
                left_width = max(left_frame.winfo_reqwidth(), left_canvas.winfo_width())
                date_header_canvas.itemconfigure(date_header_window, width=date_width)
                right_canvas.itemconfigure(right_window, width=body_width)
                left_canvas.itemconfigure(left_window, width=left_width)
                date_header_canvas.configure(scrollregion=(0, 0, date_width, max(date_header_frame.winfo_reqheight(), 1)))
                right_canvas.configure(scrollregion=(0, 0, body_width, max(right_frame.winfo_reqheight(), 1)))
                left_canvas.configure(scrollregion=(0, 0, left_width, max(left_frame.winfo_reqheight(), 1)))
                # Pastikan dua header dan dua panel punya lebar kolom yang sama.
                date_header_canvas.xview_moveto(right_canvas.xview()[0])
            except tk.TclError:
                pass

        for widget in (date_header_frame, left_frame, right_frame, date_header_canvas, left_canvas, right_canvas):
            widget.bind("<Configure>", refresh_regions, add="+")

        def sync_mousewheel(event):
            delta = int(-event.delta / 120)
            left_canvas.yview_scroll(delta, "units")
            right_canvas.yview_scroll(delta, "units")
            return "break"

        def horizontal_mousewheel(event):
            right_canvas.xview_scroll(int(-event.delta / 120), "units")
            date_header_canvas.xview_scroll(int(-event.delta / 120), "units")
            return "break"

        for widget in (left_canvas, right_canvas, left_frame, right_frame):
            widget.bind("<MouseWheel>", sync_mousewheel, add="+")
            widget.bind("<Shift-MouseWheel>", horizontal_mousewheel, add="+")
        for widget in (date_header_canvas, date_header_frame):
            widget.bind("<Shift-MouseWheel>", horizontal_mousewheel, add="+")
        self._refresh_activity_calendar_regions = refresh_regions

    def build_kalender_keseluruhan(self, parent):
        """Kalender Keseluruhan dengan header tanggal tetap dan nama pegawai beku."""
        ttk.Style(self).configure("Holiday.TLabel", foreground="#888888", background="#ECECEC")
        root = tk.Frame(parent, bg=UI_BG)
        root.pack(fill="both", expand=True)
        header = tk.Frame(root, bg=UI_BG)
        header.pack(fill="x", padx=20, pady=(18, 8))
        tk.Label(header, text="Kalender Keseluruhan", bg=UI_BG,
                 font=("Segoe UI", 18, "bold"), fg=UI_TEXT).pack(anchor="w", pady=(0, 3))
        self.overall_info_kalender = tk.Label(header, text="", bg=UI_BG, fg=UI_MUTED)
        self.overall_info_kalender.pack(anchor="w")
        tools = tk.Frame(header, bg=UI_BG)
        tools.pack(fill="x", pady=(10, 0))
        ttk.Button(tools, text="⚙ Atur Hari Libur", command=self.atur_hari_libur).pack(side="left")
        self.overall_info_libur = tk.Label(tools, text="", bg=UI_BG, fg=UI_MUTED)
        self.overall_info_libur.pack(side="left", padx=12)

        legend_holder = tk.Frame(header, bg=UI_BG)
        legend_holder.pack(fill="x", pady=(10, 0))
        self.overall_legend_canvas = tk.Canvas(legend_holder, bg=UI_BG, highlightthickness=0, height=36)
        self.overall_legend_canvas.pack(fill="x", expand=True)
        self.overall_legend_scroll = tk.Scrollbar(legend_holder, orient="horizontal",
            command=self.overall_legend_canvas.xview, width=14, bg="#94A3B8",
            troughcolor="#E2E8F0", activebackground="#64748B", relief="flat", bd=0)
        self.overall_legend_scroll.pack(fill="x", pady=(2, 0))
        self.overall_legend_canvas.configure(xscrollcommand=self.overall_legend_scroll.set)
        self.overall_legend_kegiatan = tk.Frame(self.overall_legend_canvas, bg=UI_BG)
        self._overall_legend_window = self.overall_legend_canvas.create_window(
            (0, 0), window=self.overall_legend_kegiatan, anchor="nw")
        self.overall_legend_kegiatan.bind("<Configure>", lambda e:
            self.overall_legend_canvas.configure(scrollregion=self.overall_legend_canvas.bbox("all")))
        self.overall_legend_canvas.bind("<Configure>", lambda e:
            self.overall_legend_canvas.itemconfigure(self._overall_legend_window, height=max(34, e.height)))
        self.overall_calendar_detail = tk.Label(header, text="", bg=UI_BG, fg=UI_MUTED, anchor="w")
        self.overall_calendar_detail.pack(fill="x", pady=(5, 0))

        viewport = tk.Frame(root, bg=UI_PANEL, highlightbackground=UI_BORDER, highlightthickness=1)
        viewport.pack(fill="both", expand=True, padx=20, pady=(4, 16))
        viewport.grid_rowconfigure(1, weight=1)
        viewport.grid_columnconfigure(1, weight=1)

        names_header = tk.Frame(viewport, bg=UI_TABLE_HEAD, width=270, height=48)
        names_header.grid(row=0, column=0, sticky="nsew")
        names_header.grid_propagate(False)
        names_header.grid_columnconfigure(0, minsize=42, weight=0)
        names_header.grid_columnconfigure(1, minsize=220, weight=1)
        ttk.Label(names_header, text="No.", anchor="center", font=("Segoe UI", 9, "bold")).grid(
            row=0, column=0, sticky="nsew", padx=1, pady=1)
        ttk.Label(names_header, text="Nama Pegawai", anchor="center", font=("Segoe UI", 9, "bold")).grid(
            row=0, column=1, sticky="nsew", padx=1, pady=1)

        date_header_canvas = tk.Canvas(viewport, bg=UI_TABLE_HEAD, highlightthickness=0, height=48, xscrollincrement=20)
        date_header_canvas.grid(row=0, column=1, sticky="ew")
        date_header_frame = tk.Frame(date_header_canvas, bg=UI_TABLE_HEAD)
        date_header_window = date_header_canvas.create_window((0, 0), window=date_header_frame, anchor="nw")

        names_canvas = tk.Canvas(viewport, bg=UI_PANEL, highlightthickness=0, width=270, yscrollincrement=20)
        names_canvas.grid(row=1, column=0, sticky="nsew")
        names_frame = tk.Frame(names_canvas, bg=UI_PANEL)
        names_window = names_canvas.create_window((0, 0), window=names_frame, anchor="nw")
        names_frame.grid_columnconfigure(0, minsize=42, weight=0)
        names_frame.grid_columnconfigure(1, minsize=220, weight=1)

        dates_canvas = tk.Canvas(viewport, bg=UI_PANEL, highlightthickness=0, xscrollincrement=20, yscrollincrement=20)
        dates_canvas.grid(row=1, column=1, sticky="nsew")
        dates_frame = tk.Frame(dates_canvas, bg=UI_PANEL)
        dates_window = dates_canvas.create_window((0, 0), window=dates_frame, anchor="nw")

        vscroll = tk.Scrollbar(viewport, orient="vertical",
            command=lambda *a: (names_canvas.yview(*a), dates_canvas.yview(*a)), width=18)
        vscroll.grid(row=1, column=2, sticky="ns")
        hscroll = tk.Scrollbar(viewport, orient="horizontal", width=18)
        hscroll.grid(row=2, column=1, sticky="ew")
        def xscroll(*args):
            date_header_canvas.xview(*args)
            dates_canvas.xview(*args)
        hscroll.configure(command=xscroll)
        date_header_canvas.configure(xscrollcommand=hscroll.set)
        dates_canvas.configure(xscrollcommand=hscroll.set)
        names_canvas.configure(yscrollcommand=vscroll.set)

        self.overall_names_header = names_header
        self.overall_header_frame = date_header_frame
        self.overall_names_canvas = names_canvas
        self.overall_names_frame = names_frame
        self.overall_calendar_canvas = dates_canvas
        self.overall_calendar_frame = dates_frame
        self.overall_page_canvas = dates_canvas

        def refresh_regions(event=None):
            try:
                for frame in (date_header_frame, names_frame, dates_frame):
                    frame.update_idletasks()

                dw = max(date_header_frame.winfo_reqwidth(), date_header_canvas.winfo_width())
                nw = max(names_frame.winfo_reqwidth(), names_canvas.winfo_width())
                bw = max(dates_frame.winfo_reqwidth(), dates_canvas.winfo_width())

                # Samakan tinggi setiap baris berdasarkan kebutuhan terbesar dari
                # panel nama dan panel kalender. Dengan begitu, scrollbar tidak
                # mencapai baris terakhir nama sementara kalender berhenti lebih awal.
                row_count = max(
                    len(names_frame.grid_slaves()),
                    len(dates_frame.grid_slaves())
                )
                row_count = max(
                    max((int(w.grid_info().get("row", 0)) for w in names_frame.grid_slaves()), default=-1),
                    max((int(w.grid_info().get("row", 0)) for w in dates_frame.grid_slaves()), default=-1)
                ) + 1

                for row in range(row_count):
                    try:
                        name_h = names_frame.grid_bbox(0, row)[3]
                    except Exception:
                        name_h = 0
                    try:
                        date_h = dates_frame.grid_bbox(0, row)[3]
                    except Exception:
                        date_h = 0
                    shared_row_h = max(name_h, date_h, 1)
                    names_frame.grid_rowconfigure(row, minsize=shared_row_h)
                    dates_frame.grid_rowconfigure(row, minsize=shared_row_h)

                names_frame.update_idletasks()
                dates_frame.update_idletasks()
                shared_h = max(
                    names_frame.winfo_reqheight(),
                    dates_frame.winfo_reqheight(),
                    1
                )

                date_header_canvas.itemconfigure(date_header_window, width=dw)
                names_canvas.itemconfigure(names_window, width=nw, height=shared_h)
                dates_canvas.itemconfigure(dates_window, width=bw, height=shared_h)

                date_header_canvas.configure(
                    scrollregion=(0, 0, dw, max(date_header_frame.winfo_reqheight(), 1))
                )
                names_canvas.configure(scrollregion=(0, 0, nw, shared_h))
                dates_canvas.configure(scrollregion=(0, 0, bw, shared_h))

                date_header_canvas.xview_moveto(dates_canvas.xview()[0])
                # Posisi y disamakan setelah kedua scrollregion memiliki tinggi sama.
                dates_canvas.yview_moveto(names_canvas.yview()[0])
            except tk.TclError:
                pass

        for w in (date_header_frame, names_frame, dates_frame, date_header_canvas, names_canvas, dates_canvas):
            w.bind("<Configure>", refresh_regions, add="+")
        def wheel(event):
            delta = int(-event.delta / 120) * 3
            names_canvas.yview_scroll(delta, "units")
            dates_canvas.yview_scroll(delta, "units")
            return "break"
        def shift_wheel(event):
            delta = int(-event.delta / 120) * 3
            dates_canvas.xview_scroll(delta, "units")
            date_header_canvas.xview_scroll(delta, "units")
            return "break"
        for w in (names_canvas, dates_canvas, names_frame, dates_frame):
            w.bind("<MouseWheel>", wheel, add="+")
            w.bind("<Shift-MouseWheel>", shift_wheel, add="+")
        for w in (date_header_canvas, date_header_frame):
            w.bind("<Shift-MouseWheel>", shift_wheel, add="+")
        self._refresh_overall_calendar_regions = refresh_regions

    def buka_kalender_keseluruhan(self):
        """Buka Kalender Keseluruhan tepat di tengah aplikasi utama."""
        if getattr(self, "overall_calendar_window", None) is not None and self.overall_calendar_window.winfo_exists():
            self.overall_calendar_window.lift()
            self.overall_calendar_window.focus_force()
            self.center_window(self.overall_calendar_window, self)
            try:
                self.overall_calendar_window.lift()
                self.overall_calendar_window.focus_force()
            except Exception:
                pass
            return

        win = tk.Toplevel(self)
        self.overall_calendar_window = win
        self.modernize_window(win, "Kalender Keseluruhan", (700, 450))
        win.title("Kalender Keseluruhan")
        # Ukuran cukup besar untuk desktop, tetapi tetap aman untuk layar kecil/DPI tinggi.
        sw, sh = win.winfo_screenwidth(), win.winfo_screenheight()
        width = min(1500, max(900, int(sw * 0.86)))
        height = min(850, max(560, int(sh * 0.82)))
        win.geometry(f"{width}x{height}")
        win.minsize(700, 450)
        self.enable_native_window_controls(win)
        win.protocol("WM_DELETE_WINDOW", lambda: self.tutup_kalender_keseluruhan(win))
        self.center_window(win, self)

        # Jangan menutupi seluruh jendela dengan spinner saat render sinkron.
        # Overlay lama ikut membeku karena render kalender berjalan di event loop
        # utama. Tampilkan shell lebih dahulu, lalu render setelah Tk idle.
        self.after_idle(lambda w=win: self._build_overall_calendar_without_overlay(w))

    def _show_overall_loading_overlay(self, win, text="Memuat Kalender Keseluruhan..."):
        """Indikator kecil non-blocking pada Kalender Keseluruhan."""
        try:
            self._hide_overall_loading_overlay()
            if win is None or not win.winfo_exists():
                return
            chip = tk.Frame(
                win, bg=UI_PANEL, padx=13, pady=8,
                highlightthickness=1, highlightbackground=UI_BORDER, bd=0
            )
            chip.place(relx=1.0, x=-16, y=14, anchor="ne")
            chip.lift()
            spinner = tk.Canvas(
                chip, width=18, height=18, bg=UI_PANEL,
                highlightthickness=0, bd=0
            )
            spinner.pack(side="left", padx=(0, 8))
            spinner.create_oval(4, 4, 14, 14, outline=UI_BORDER, width=2)
            spinner.create_arc(
                4, 4, 14, 14, start=25, extent=105, style="arc",
                outline=UI_PRIMARY, width=2, tags=("arc",)
            )
            label = tk.Label(
                chip, text=text, bg=UI_PANEL, fg=UI_TEXT,
                font=("Segoe UI", 9, "bold")
            )
            label.pack(side="left")
            self._overall_loading_overlay = chip
            self._overall_loading_spinner = spinner
            self._overall_loading_label = label
            self._overall_loading_index = 0

            def animate():
                try:
                    if not chip.winfo_exists():
                        return
                    idx = self._overall_loading_index
                    spinner.itemconfigure("arc", start=(25 + idx * 14) % 360)
                    self._overall_loading_index = idx + 1
                    self._overall_loading_job = self.after(55, animate)
                except tk.TclError:
                    self._overall_loading_job = None

            self._overall_loading_job = self.after(55, animate)
            win.update_idletasks()
        except Exception:
            self._overall_loading_overlay = None
            self._overall_loading_spinner = None
            self._overall_loading_label = None
            self._overall_loading_job = None

    def _hide_overall_loading_overlay(self):
        job = getattr(self, "_overall_loading_job", None)
        if job is not None:
            try:
                self.after_cancel(job)
            except Exception:
                pass
        self._overall_loading_job = None
        for attr in ("_overall_loading_bar",):
            bar = getattr(self, attr, None)
            if bar is not None:
                try:
                    bar.stop()
                except Exception:
                    pass
        overlay = getattr(self, "_overall_loading_overlay", None)
        if overlay is not None:
            try:
                if overlay.winfo_exists():
                    overlay.destroy()
            except Exception:
                pass
        self._overall_loading_overlay = None
        self._overall_loading_spinner = None
        self._overall_loading_label = None
        self._overall_loading_bar = None

    def _build_overall_calendar_without_overlay(self, win):
        """Render kalender tanpa overlay animasi yang membeku pada proses sinkron."""
        try:
            if not win.winfo_exists():
                return
            self.build_kalender_keseluruhan(win)
            win.update_idletasks()
            self.center_window(win, self)
            self.load_kalender_keseluruhan()
            if win.winfo_exists():
                win.lift()
                win.focus_force()
        except Exception as exc:
            messagebox.showerror(
                "Kalender Keseluruhan",
                f"Kalender tidak dapat dimuat.\\n\\nDetail: {exc}",
                parent=win
            )

    def _build_overall_calendar_with_overlay(self, win):
        """Alias kompatibilitas; render kalender tanpa overlay layar penuh."""
        self._build_overall_calendar_without_overlay(win)

    def _render_overall_calendar_after_overlay(self, win):
        """Alias kompatibilitas; selalu bersihkan indikator setelah render."""
        try:
            if win is not None and win.winfo_exists():
                self.load_kalender_keseluruhan()
        except Exception as exc:
            messagebox.showerror(
                "Kalender Keseluruhan",
                f"Kalender tidak dapat dimuat.\\n\\nDetail: {exc}",
                parent=win if win is not None and win.winfo_exists() else self
            )
        finally:
            self._hide_overall_loading_overlay()

    def _finish_overall_calendar_loading(self):
        """Kompatibilitas untuk caller lama; render dan bersihkan indikator."""
        win = getattr(self, "overall_calendar_window", None)
        if win is not None and win.winfo_exists():
            self._render_overall_calendar_after_overlay(win)

    def tutup_kalender_keseluruhan(self, win):
        try:
            if getattr(self, "_overall_loading_overlay", None) is not None:
                self._hide_overall_loading_overlay()
            if win.winfo_exists():
                win.destroy()
        finally:
            self.overall_calendar_window = None
            for attr in (
                "overall_calendar_frame",
                "overall_calendar_canvas",
                "overall_header_frame",
                "overall_names_frame",
                "overall_names_canvas",
                "overall_names_header",
                "overall_info_kalender",
                "overall_info_libur",
                "overall_legend_kegiatan",
                "overall_legend_canvas",
                "overall_legend_scroll",
                "overall_calendar_detail",
            ):
                if hasattr(self, attr):
                    try:
                        delattr(self, attr)
                    except Exception:
                        pass

    def buka_data_pegawai(self):
        """Buka data pegawai sebagai satu jendela global."""
        if getattr(self, "data_pegawai_window", None) is not None and self.data_pegawai_window.winfo_exists():
            self.data_pegawai_window.lift()
            self.data_pegawai_window.focus_force()
            self.refresh_data_pegawai()
            return

        win = tk.Toplevel(self)
        self.data_pegawai_window = win
        self.modernize_window(win, "Data Pegawai", (1000, 600))
        win.title("Data Pegawai")
        win.geometry("1450x750")
        win.minsize(1100, 600)
        self.enable_native_window_controls(win)
        win.transient(self)
        self.enable_native_window_controls(win)
        win.protocol("WM_DELETE_WINDOW", lambda: self.tutup_data_pegawai(win))
        self.center_window(win, self)
        self.build_pegawai(win)
        self.refresh_data_pegawai()

    def tutup_data_pegawai(self, win):
        try:
            if win.winfo_exists():
                win.destroy()
        finally:
            self.data_pegawai_window = None
            self.table_pegawai = None
            self.search_data_pegawai_var = None
            self.label_jumlah_pegawai = None
            self.tampilkan_nonaktif_var = None

    def build_pegawai(self, parent):

        tk.Label(
            parent,
            text="DATA PEGAWAI",
            bg=UI_BG,
            font=("Segoe UI", 17, "bold")
        ).pack(
            anchor="w",
            padx=20,
            pady=(15, 5)
        )

        self.label_jumlah_pegawai = tk.Label(
            parent,
            text="",
            bg=UI_BG,
            fg=UI_MUTED,
            font=("Segoe UI", 10)
        )
        self.label_jumlah_pegawai.pack(
            anchor="w",
            padx=20
        )

        # =====================================================
        # TOOLBAR DATA PEGAWAI
        # =====================================================
        toolbar = tk.Frame(parent, bg=UI_BG)
        toolbar.pack(fill="x", padx=20, pady=(12, 5))

        ttk.Button(
            toolbar,
            text="＋ Tambah Pegawai",
            command=lambda: self.buka_form_pegawai()
        ).pack(side="left", padx=(0, 6))

        ttk.Button(
            toolbar,
            text="✎ Edit",
            command=self.edit_pegawai_terpilih
        ).pack(side="left", padx=6)

        ttk.Button(
            toolbar,
            text="🗑 Nonaktifkan",
            command=self.hapus_pegawai_terpilih
        ).pack(side="left", padx=6)

        ttk.Button(
            toolbar,
            text="🗑 Hapus Permanen",
            command=self.hapus_pegawai_permanen_terpilih
        ).pack(side="left", padx=6)

        ttk.Button(
            toolbar,
            text="↥ Impor dari Excel",
            command=self.impor_pegawai_dari_excel
        ).pack(side="left", padx=6)

        ttk.Button(
            toolbar,
            text="✓ Aktifkan",
            command=self.aktifkan_pegawai_terpilih
        ).pack(side="left", padx=6)

        ttk.Button(
            toolbar,
            text="↻ Refresh Semua",
            command=self.refresh_all
        ).pack(side="left", padx=6)

        self.tampilkan_nonaktif_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            toolbar,
            text="Tampilkan pegawai nonaktif",
            variable=self.tampilkan_nonaktif_var,
            command=self.refresh_data_pegawai
        ).pack(side="right", padx=(10, 0))

        # =====================================================
        # SEARCH DATA PEGAWAI
        # =====================================================
        search_frame = tk.Frame(parent, bg=UI_BG)
        search_frame.pack(fill="x", padx=20, pady=(5, 5))

        tk.Label(
            search_frame,
            text="Cari Pegawai:",
            bg=UI_BG,
            font=("Segoe UI", 10, "bold")
        ).pack(side="left", padx=(0, 8))

        self.search_data_pegawai_var = tk.StringVar()
        self.entry_search_data_pegawai = ttk.Entry(
            search_frame,
            textvariable=self.search_data_pegawai_var,
            width=45
        )
        self.entry_search_data_pegawai.pack(side="left")
        self.entry_search_data_pegawai.bind(
            "<KeyRelease>",
            lambda event: self.refresh_data_pegawai()
        )

        ttk.Button(
            search_frame,
            text="✕",
            width=3,
            command=self.clear_search_data_pegawai
        ).pack(side="left", padx=5)

        # =====================================================
        # TABLE
        # =====================================================
        frame = tk.Frame(parent, bg=UI_PANEL, highlightbackground=UI_BORDER, highlightthickness=1)
        frame.pack(fill="both", expand=True, padx=20, pady=15)

        columns = (
            "no",
            "nama",
            "nip",
            "pangkat",
            "golongan",
            "jabatan",
            "tarif",
            "bulan_ini",
            "status"
        )

        self.table_pegawai = ttk.Treeview(
            frame,
            columns=columns,
            show="headings",
            selectmode="browse"
        )

        heads = {
            "no": "No.",
            "nama": "Nama",
            "nip": "NIP",
            "pangkat": "Pangkat",
            "golongan": "Golongan",
            "jabatan": "Jabatan",
            "tarif": "Tarif",
            "bulan_ini": "Turun Bulan Ini",
            "status": "Status"
        }

        widths = {
            "no": 58,
            "nama": 280,
            "nip": 185,
            "pangkat": 170,
            "golongan": 105,
            "jabatan": 230,
            "tarif": 135,
            "bulan_ini": 145,
            "status": 100
        }

        for col in columns:
            self.table_pegawai.heading(col, text=heads[col])
            self.table_pegawai.column(
                col,
                width=widths[col],
                anchor="center"
            )

        self.table_pegawai.column("nama", anchor="w")
        self.table_pegawai.column("jabatan", anchor="w")

        # Gunakan grid untuk memastikan scrollbar tidak pernah tertutup oleh
        # Treeview (pack + expand sebelumnya dapat membuat scrollbar tampak hilang
        # pada DPI tinggi). Scrollbar dibuat lebih tebal agar mudah terlihat.
        frame.grid_rowconfigure(0, weight=1)
        frame.grid_columnconfigure(0, weight=1)
        self.table_pegawai.grid(row=0, column=0, sticky="nsew")

        scroll = tk.Scrollbar(
            frame, orient="vertical", command=self.table_pegawai.yview,
            width=18, bg="#94A3B8", troughcolor="#E2E8F0",
            activebackground="#64748B", relief="flat", bd=0
        )
        scroll.grid(row=0, column=1, sticky="ns")
        scroll_x = tk.Scrollbar(
            frame, orient="horizontal", command=self.table_pegawai.xview,
            width=18, bg="#94A3B8", troughcolor="#E2E8F0",
            activebackground="#64748B", relief="flat", bd=0
        )
        scroll_x.grid(row=1, column=0, sticky="ew")
        self.table_pegawai.configure(yscrollcommand=scroll.set, xscrollcommand=scroll_x.set)
        self.modernize_tree(self.table_pegawai)
        # Scroll tabel dengan wheel / Shift+wheel tanpa harus tepat di scrollbar.
        self.table_pegawai.bind("<MouseWheel>", lambda e: self.table_pegawai.yview_scroll(int(-e.delta / 120), "units"), add="+")
        self.table_pegawai.bind("<Shift-MouseWheel>", lambda e: self.table_pegawai.xview_scroll(int(-e.delta / 120), "units"), add="+")

        # Double-click baris = edit.
        self.table_pegawai.bind(
            "<Double-1>",
            lambda event: self.edit_pegawai_terpilih()
        )
        self.table_pegawai.bind(
            "<Return>",
            lambda event: self.edit_pegawai_terpilih()
        )
        self.table_pegawai.bind(
            "<Delete>",
            lambda event: self.hapus_pegawai_terpilih()
        )

        self.refresh_data_pegawai()

    def _pegawai_terpilih_id(self):
        """Mengambil ID pegawai dari baris yang sedang dipilih."""
        if not hasattr(self, "table_pegawai") or self.table_pegawai is None:
            return None
        try:
            pilihan = self.table_pegawai.selection()
        except tk.TclError:
            return None
        if not pilihan:
            return None
        try:
            return int(pilihan[0])
        except (ValueError, TypeError):
            return None

    def _ambil_pegawai_by_id(self, pegawai_id):
        conn = get_connection()
        try:
            return conn.execute(
                """
                SELECT id, nama, nip, pangkat, golongan_asli,
                       golongan_utama, jabatan, tarif, aktif
                FROM pegawai
                WHERE id = ?
                """,
                (pegawai_id,)
            ).fetchone()
        finally:
            conn.close()

    def edit_pegawai_terpilih(self):
        pegawai_id = self._pegawai_terpilih_id()
        if pegawai_id is None:
            messagebox.showinfo(
                "Data Pegawai",
                "Pilih satu pegawai yang ingin diedit terlebih dahulu.",
                parent=getattr(self, "data_pegawai_window", self)
            )
            return
        self.buka_form_pegawai(pegawai_id)

    def buka_form_pegawai(self, pegawai_id=None):
        """Form tambah/edit pegawai."""
        parent = getattr(self, "data_pegawai_window", self)
        edit_mode = pegawai_id is not None

        data = None
        if edit_mode:
            data = self._ambil_pegawai_by_id(pegawai_id)
            if not data:
                messagebox.showerror(
                    "Data Pegawai",
                    "Data pegawai tidak ditemukan.",
                    parent=parent
                )
                self.refresh_data_pegawai()
                return

        win = tk.Toplevel(parent)
        self.enable_native_window_controls(win)
        win.title("Edit Pegawai" if edit_mode else "Tambah Pegawai")
        win.geometry("560x560")
        win.minsize(500, 500)
        win.transient(parent)
        self.enable_native_window_controls(win)
        win.grab_set()
        self.enable_native_window_controls(win)
        self.center_window(win, parent)

        outer = tk.Frame(win, bg=UI_BG)
        outer.pack(fill="both", expand=True)

        tk.Label(
            outer,
            text="EDIT DATA PEGAWAI" if edit_mode else "TAMBAH DATA PEGAWAI",
            bg=UI_BG,
            font=("Segoe UI Semibold", 14)
        ).pack(anchor="w", padx=25, pady=(20, 5))

        tk.Label(
            outer,
            text="Isi data pegawai. Tarif disimpan sebagai tarif aktual pegawai.",
            bg=UI_BG,
            fg=UI_MUTED,
            font=("Segoe UI", 9)
        ).pack(anchor="w", padx=25, pady=(0, 15))

        form = tk.Frame(outer, bg=UI_PANEL, bd=1, relief="solid")
        form.pack(fill="both", expand=True, padx=25, pady=(0, 15))

        values = {
            "nama": data[1] if data else "",
            "nip": data[2] if data else "",
            "pangkat": data[3] if data else "",
            "golongan": data[4] if data else "",
            "jabatan": data[6] if data else "",
            "tarif": str(data[7]) if data else "",
        }

        vars_ = {key: tk.StringVar(value=value or "") for key, value in values.items()}

        fields = [
            ("Nama *", "nama"),
            ("NIP", "nip"),
            ("Pangkat", "pangkat"),
            ("Golongan", "golongan"),
            ("Jabatan", "jabatan"),
            ("Tarif *", "tarif"),
        ]

        entries = {}
        for row, (label, key) in enumerate(fields):
            tk.Label(
                form,
                text=label,
                bg=UI_PANEL,
                font=("Segoe UI", 10, "bold")
            ).grid(row=row, column=0, sticky="w", padx=18, pady=9)

            ent = ttk.Entry(form, textvariable=vars_[key], width=45)
            ent.grid(row=row, column=1, sticky="ew", padx=(0, 18), pady=9)
            entries[key] = ent

        form.columnconfigure(1, weight=1)

        # Golongan utama hanya informasi turunan; pengguna tetap mengisi
        # golongan asli seperti III d, III c, VII, X, dst.
        tk.Label(
            form,
            text="Golongan Utama",
            bg=UI_PANEL,
            fg=UI_MUTED,
            font=("Segoe UI", 9)
        ).grid(row=len(fields), column=0, sticky="w", padx=18, pady=(5, 14))

        gol_utama_var = tk.StringVar(value=golongan_utama_dari(vars_["golongan"].get()))
        tk.Label(
            form,
            textvariable=gol_utama_var,
            bg=UI_PANEL,
            fg=UI_TEXT,
            font=("Segoe UI", 10, "bold")
        ).grid(row=len(fields), column=1, sticky="w", padx=(0, 18), pady=(5, 14))

        def update_golongan_utama(*_):
            gol_utama_var.set(golongan_utama_dari(vars_["golongan"].get()))

        vars_["golongan"].trace_add("write", update_golongan_utama)

        def simpan():
            nama = vars_["nama"].get().strip()
            nip = vars_["nip"].get().strip()
            pangkat = vars_["pangkat"].get().strip()
            golongan = vars_["golongan"].get().strip()
            jabatan = vars_["jabatan"].get().strip()
            tarif_text = vars_["tarif"].get().strip()

            if not nama:
                messagebox.showwarning(
                    "Data Pegawai",
                    "Nama pegawai wajib diisi.",
                    parent=win
                )
                entries["nama"].focus_set()
                return

            try:
                tarif = parse_rupiah(tarif_text)
            except (ValueError, TypeError):
                messagebox.showwarning(
                    "Data Pegawai",
                    "Tarif harus berupa angka, misalnya 80000 atau Rp 80.000.",
                    parent=win
                )
                entries["tarif"].focus_set()
                return

            if tarif <= 0:
                messagebox.showwarning(
                    "Data Pegawai",
                    "Tarif harus lebih besar dari 0.",
                    parent=win
                )
                entries["tarif"].focus_set()
                return

            # NIP tidak boleh dobel jika diisi.
            conn = get_connection()
            try:
                if nip:
                    if edit_mode:
                        duplikat = conn.execute(
                            "SELECT id FROM pegawai WHERE nip = ? AND id <> ? LIMIT 1",
                            (nip, pegawai_id)
                        ).fetchone()
                    else:
                        duplikat = conn.execute(
                            "SELECT id FROM pegawai WHERE nip = ? LIMIT 1",
                            (nip,)
                        ).fetchone()
                    if duplikat:
                        messagebox.showwarning(
                            "Data Pegawai",
                            f"NIP {nip} sudah digunakan oleh pegawai lain.",
                            parent=win
                        )
                        entries["nip"].focus_set()
                        return

                golongan_utama = golongan_utama_dari(golongan)

                if edit_mode:
                    conn.execute(
                        """
                        UPDATE pegawai
                        SET nama = ?, nip = ?, pangkat = ?,
                            golongan_asli = ?, golongan_utama = ?,
                            jabatan = ?, tarif = ?
                        WHERE id = ?
                        """,
                        (
                            nama, nip, pangkat,
                            golongan, golongan_utama,
                            jabatan, tarif, pegawai_id
                        )
                    )
                else:
                    no_row = conn.execute(
                        "SELECT COALESCE(MAX(no), 0) + 1 FROM pegawai"
                    ).fetchone()
                    no = int(no_row[0] or 1)
                    conn.execute(
                        """
                        INSERT INTO pegawai
                        (no, nama, nip, pangkat, golongan_asli,
                         golongan_utama, jabatan, tarif, aktif)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, 1)
                        """,
                        (
                            no, nama, nip, pangkat,
                            golongan, golongan_utama,
                            jabatan, tarif
                        )
                    )

                conn.commit()
            except Exception as exc:
                conn.rollback()
                messagebox.showerror(
                    "Data Pegawai",
                    "Data pegawai gagal disimpan. Periksa kembali isian lalu coba lagi.",
                    parent=win
                )
                return
            finally:
                conn.close()

            win.grab_release()
            win.destroy()
            self.pegawai = load_pegawai()
            self.refresh_pj_options()
            self.refresh_selection_table()
            self.load_bagian_data()
            self.refresh_data_pegawai()
            self.load_kalender()
            self.load_kalender_keseluruhan()

        button_frame = tk.Frame(outer, bg=UI_BG)
        button_frame.pack(fill="x", padx=25, pady=(0, 20))

        ttk.Button(
            button_frame,
            text="Batal",
            command=lambda: (win.grab_release(), win.destroy())
        ).pack(side="right", padx=(8, 0))
        ttk.Button(
            button_frame,
            text="Simpan",
            command=simpan
        ).pack(side="right")

        entries["nama"].focus_set()
        win.bind("<Escape>", lambda event: (win.grab_release(), win.destroy()))
        win.bind("<Return>", lambda event: simpan())
        self.fit_popup(
            win, parent,
            min_width=520, min_height=520, max_width=760, max_height=760
        )

    def aktifkan_pegawai_terpilih(self):
        pegawai_id = self._pegawai_terpilih_id()
        if pegawai_id is None:
            messagebox.showinfo(
                "Data Pegawai",
                "Pilih satu pegawai yang ingin diaktifkan kembali.",
                parent=getattr(self, "data_pegawai_window", self)
            )
            return

        data = self._ambil_pegawai_by_id(pegawai_id)
        if not data:
            return
        if int(data[8] or 0) == 1:
            messagebox.showinfo(
                "Data Pegawai",
                f"{data[1]} sudah berstatus aktif.",
                parent=getattr(self, "data_pegawai_window", self)
            )
            return

        jawab = messagebox.askyesno(
            "Aktifkan Pegawai",
            f"Aktifkan kembali pegawai {data[1]}?",
            parent=getattr(self, "data_pegawai_window", self)
        )
        if not jawab:
            return

        conn = get_connection()
        try:
            conn.execute("UPDATE pegawai SET aktif = 1 WHERE id = ?", (pegawai_id,))
            conn.commit()
        except Exception as exc:
            conn.rollback()
            messagebox.showerror(
                "Aktifkan Pegawai",
                "Pegawai gagal diaktifkan. Silakan coba lagi.",
                parent=getattr(self, "data_pegawai_window", self)
            )
            return
        finally:
            conn.close()

        self.pegawai = load_pegawai()
        self.refresh_pj_options()
        self.refresh_selection_table()
        self.load_bagian_data()
        self.refresh_data_pegawai()
        self.load_kalender()
        self.load_kalender_keseluruhan()

    def hapus_pegawai_terpilih(self):
        pegawai_id = self._pegawai_terpilih_id()
        if pegawai_id is None:
            messagebox.showinfo(
                "Data Pegawai",
                "Pilih satu pegawai yang ingin dihapus terlebih dahulu.",
                parent=getattr(self, "data_pegawai_window", self)
            )
            return

        data = self._ambil_pegawai_by_id(pegawai_id)
        if not data:
            return

        nama = data[1] or "Tanpa Nama"

        # Jangan menonaktifkan pegawai yang masih dipakai pada kegiatan
        # periode aktif. Pengguna perlu melepas/mengganti pegawai tersebut
        # dari kegiatan bulan berjalan terlebih dahulu agar jadwal tetap valid.
        conn = get_connection()
        try:
            dipakai_periode_ini = conn.execute(
                """
                SELECT COUNT(*)
                FROM anggota_bagian ab
                JOIN bagian b ON b.id = ab.bagian_id
                WHERE b.periode_id = ? AND ab.pegawai_id = ?
                """,
                (self.periode_id, pegawai_id)
            ).fetchone()[0]
        finally:
            conn.close()

        if dipakai_periode_ini:
            messagebox.showwarning(
                "Pegawai Masih Dipakai",
                f"{nama} masih terdaftar pada kegiatan di {NAMA_BULAN[self.bulan]} {self.tahun}.\n\n"
                "Lepaskan pegawai dari kegiatan bulan tersebut terlebih dahulu, "
                "baru pegawai dapat dinonaktifkan.",
                parent=getattr(self, "data_pegawai_window", self)
            )
            return

        # Hapus secara soft-delete supaya riwayat penugasan lama tetap aman.
        jawab = messagebox.askyesno(
            "Hapus Pegawai",
            f"Yakin ingin menghapus/nonaktifkan pegawai berikut?\n\n"
            f"Nama : {nama}\n"
            f"NIP  : {data[2] or '-'}\n\n"
            f"Data tidak dihapus permanen. Pegawai akan dibuat nonaktif "
            f"agar riwayat penugasan tetap tersimpan.",
            parent=getattr(self, "data_pegawai_window", self)
        )
        if not jawab:
            return

        conn = get_connection()
        try:
            conn.execute(
                "UPDATE pegawai SET aktif = 0 WHERE id = ?",
                (pegawai_id,)
            )
            conn.commit()
        except Exception as exc:
            conn.rollback()
            messagebox.showerror(
                "Hapus Pegawai",
                "Pegawai gagal dinonaktifkan. Silakan coba lagi.",
                parent=getattr(self, "data_pegawai_window", self)
            )
            return
        finally:
            conn.close()

        self.pegawai = load_pegawai()
        self.refresh_pj_options()
        self.refresh_selection_table()
        self.load_bagian_data()
        self.refresh_data_pegawai()
        self.load_kalender()
        self.load_kalender_keseluruhan()


    def hapus_pegawai_permanen_terpilih(self):
        """Menghapus pegawai secara permanen beserta seluruh riwayat yang terkait."""
        pegawai_id = self._pegawai_terpilih_id()
        parent = getattr(self, "data_pegawai_window", self)
        if pegawai_id is None:
            messagebox.showinfo(
                "Hapus Permanen",
                "Pilih satu pegawai yang ingin dihapus permanen terlebih dahulu.",
                parent=parent
            )
            return

        data = self._ambil_pegawai_by_id(pegawai_id)
        if not data:
            return

        nama = data[1] or "Tanpa Nama"
        nip = data[2] or "-"

        conn = get_connection()
        try:
            jumlah_jadwal = conn.execute(
                "SELECT COUNT(*) FROM jadwal WHERE pegawai_id = ?",
                (pegawai_id,)
            ).fetchone()[0]
            jumlah_anggota = conn.execute(
                "SELECT COUNT(*) FROM anggota_bagian WHERE pegawai_id = ?",
                (pegawai_id,)
            ).fetchone()[0]
        finally:
            conn.close()

        pesan = (
            f"PERINGATAN: tindakan ini PERMANEN.\n\n"
            f"Pegawai:\n{nama}\nNIP: {nip}\n\n"
            f"Yang akan dihapus:\n"
            f"• Data master pegawai\n"
            f"• {jumlah_anggota} keanggotaan kegiatan\n"
            f"• {jumlah_jadwal} jadwal/rekam turun\n\n"
            f"Riwayat pegawai ini tidak dapat dipulihkan dari aplikasi.\n"
            f"Pastikan sudah melakukan backup/export terlebih dahulu.\n\n"
            f"Lanjutkan?"
        )
        if not messagebox.askyesno("Hapus Pegawai Permanen", pesan, parent=parent):
            return

        konfirmasi = simpledialog.askstring(
            "Konfirmasi Terakhir",
            f"Untuk benar-benar menghapus {nama}, ketik:\n\nHAPUS PERMANEN",
            parent=parent
        )
        if konfirmasi != "HAPUS PERMANEN":
            messagebox.showinfo(
                "Dibatalkan",
                "Penghapusan permanen dibatalkan. Data pegawai tetap aman.",
                parent=parent
            )
            return

        conn = get_connection()
        try:
            # Lepaskan dulu referensi PJ agar aman jika foreign key aktif.
            conn.execute(
                "UPDATE bagian SET penanggung_jawab_id = NULL WHERE penanggung_jawab_id = ?",
                (pegawai_id,)
            )
            conn.execute("DELETE FROM jadwal WHERE pegawai_id = ?", (pegawai_id,))
            conn.execute("DELETE FROM anggota_bagian WHERE pegawai_id = ?", (pegawai_id,))
            conn.execute("DELETE FROM pegawai WHERE id = ?", (pegawai_id,))
            conn.commit()
        except Exception as exc:
            conn.rollback()
            messagebox.showerror(
                "Hapus Permanen Gagal",
                "Pegawai gagal dihapus permanen. Silakan coba lagi.",
                parent=parent
            )
            return
        finally:
            conn.close()

        self.pegawai = load_pegawai()
        self.selected_ids.discard(pegawai_id)
        self.refresh_pj_options()
        self.refresh_selection_table()
        self.load_bagian_data()
        self.refresh_data_pegawai()
        self.load_kalender()
        self.load_kalender_keseluruhan()
        self.refresh_data_kegiatan()
        messagebox.showinfo(
            "Berhasil",
            f"Pegawai {nama} sudah dihapus permanen beserta seluruh riwayat terkait.",
            parent=parent
        )

    def impor_pegawai_dari_excel(self):
        """
        Mengganti seluruh master pegawai dari file Excel yang dipilih user.

        Nama file bebas. Syarat satu-satunya: workbook memiliki sheet
        bernama persis ``DaftarPegawai``. Kolom dicari berdasarkan nama
        header, bukan posisi kolom, sehingga sheet tetap bisa memiliki
        kolom tambahan.
        """
        parent = getattr(self, "data_pegawai_window", self)
        if load_workbook is None:
            messagebox.showerror(
                "Import Excel",
                "Library openpyxl belum terpasang. Jalankan: pip install openpyxl",
                parent=parent
            )
            return

        path = filedialog.askopenfilename(
            parent=parent,
            title="Pilih File Excel Pegawai",
            filetypes=[
                ("Excel Macro-Enabled (*.xlsm)", "*.xlsm"),
                ("Excel Workbook (*.xlsx)", "*.xlsx"),
                ("Semua File Excel", "*.xlsm *.xlsx"),
            ]
        )
        if not path:
            return

        # Nama file TIDAK diperiksa. Yang diperiksa adalah nama sheet.
        try:
            wb = load_workbook(
                filename=path,
                read_only=True,
                data_only=True,
                keep_vba=path.lower().endswith(".xlsm")
            )
        except Exception as exc:
            messagebox.showerror(
                "Import Excel",
                "File Excel tidak dapat dibaca.\n\n"
                "Proses tidak dapat diselesaikan. Silakan coba lagi.",
                parent=parent
            )
            return

        sheet_target = "DaftarPegawai"
        if sheet_target not in wb.sheetnames:
            tersedia = "\n".join(f"• {nama}" for nama in wb.sheetnames[:30])
            if len(wb.sheetnames) > 30:
                tersedia += "\n• ..."
            wb.close()
            messagebox.showerror(
                "Import Excel",
                f"Sheet '{sheet_target}' tidak ditemukan.\n\n"
                "Nama file bebas, tetapi workbook wajib memiliki sheet "
                "bernama persis 'DaftarPegawai'.\n\n"
                "Sheet yang ditemukan:\n"
                f"{tersedia}",
                parent=parent
            )
            return

        ws = wb[sheet_target]

        # ---------------------------------------------------------
        # Cari header berdasarkan NAMA KOLOM, bukan posisi kolom.
        # Contoh file kamu:
        # NO | NAMA | NIP | PANGKAT | GOL | JABATAN | Angka | ...
        # ---------------------------------------------------------
        def norm_header(value):
            if value is None:
                return ""
            return re.sub(r"\s+", " ", str(value).strip().upper())

        header_row_no = None
        header_map = {}
        nama_alias = {"NAMA", "NAMA PEGAWAI"}
        nip_alias = {"NIP"}
        pangkat_alias = {"PANGKAT"}
        gol_alias = {"GOL", "GOLONGAN", "GOLONGAN ASLI"}
        jabatan_alias = {"JABATAN"}
        tarif_alias = {"ANGKA", "TARIF", "NOMINAL", "ANGKA SESUAI INPUT"}

        max_scan_rows = min(ws.max_row or 1, 30)
        for r in range(1, max_scan_rows + 1):
            kandidat = {}
            for c in range(1, (ws.max_column or 1) + 1):
                h = norm_header(ws.cell(r, c).value)
                if h:
                    kandidat[h] = c

            if (
                any(h in kandidat for h in nama_alias)
                and any(h in kandidat for h in nip_alias)
                and any(h in kandidat for h in pangkat_alias)
                and any(h in kandidat for h in gol_alias)
                and any(h in kandidat for h in jabatan_alias)
            ):
                header_row_no = r
                header_map = kandidat
                break

        if header_row_no is None:
            wb.close()
            messagebox.showerror(
                "Import Excel",
                "Sheet 'DaftarPegawai' ditemukan, tetapi header pegawai tidak dikenali.\n\n"
                "Header minimal yang harus ada:\n"
                "NAMA, NIP, PANGKAT, GOL, JABATAN\n\n"
                "Tarif boleh menggunakan salah satu nama: Angka, Tarif, Nominal, "
                "atau Angka Sesuai Input.",
                parent=parent
            )
            return

        def find_col(aliases):
            for alias in aliases:
                if alias in header_map:
                    return header_map[alias]
            return None

        col_nama = find_col(nama_alias)
        col_nip = find_col(nip_alias)
        col_pangkat = find_col(pangkat_alias)
        col_gol = find_col(gol_alias)
        col_jabatan = find_col(jabatan_alias)
        # Prioritaskan "ANGKA" seperti file DaftarPegawai kamu.
        col_tarif = (
            header_map.get("ANGKA")
            or header_map.get("TARIF")
            or header_map.get("NOMINAL")
            or header_map.get("ANGKA SESUAI INPUT")
        )
        col_no = header_map.get("NO")

        rows = []
        nip_seen = set()
        errors = []

        for row_no in range(header_row_no + 1, (ws.max_row or 0) + 1):
            def cell_value(col):
                return ws.cell(row_no, col).value if col else None

            values = [cell_value(c) for c in [col_no, col_nama, col_nip, col_pangkat, col_gol, col_jabatan, col_tarif]]
            if not any(v is not None and str(v).strip() != "" for v in values):
                continue

            no = nilai_excel_teks(cell_value(col_no))
            nama = nilai_excel_teks(cell_value(col_nama))
            nip = nilai_excel_teks(cell_value(col_nip))
            pangkat = nilai_excel_teks(cell_value(col_pangkat))
            golongan = nilai_excel_teks(cell_value(col_gol))
            jabatan = nilai_excel_teks(cell_value(col_jabatan))
            tarif = parse_tarif_excel(cell_value(col_tarif))

            # Jika baris tidak punya nama sama sekali, abaikan baris kosong/teks tambahan.
            if not nama:
                continue

            if nip and nip in nip_seen:
                errors.append(f"Baris {row_no}: NIP {nip} duplikat")
                continue
            if nip:
                nip_seen.add(nip)

            rows.append({
                "no": no,
                "nama": nama,
                "nip": nip,
                "pangkat": pangkat,
                "golongan": golongan,
                "golongan_utama": golongan_utama_dari(golongan),
                "jabatan": jabatan,
                "tarif": tarif,
            })

        wb.close()

        if errors:
            messagebox.showerror(
                "Import Excel Dibatalkan",
                "Ada masalah pada data Excel. Tidak ada data yang diubah.\n\n"
                + "\n".join(errors[:20])
                + ("\n..." if len(errors) > 20 else ""),
                parent=parent
            )
            return

        if not rows:
            messagebox.showwarning(
                "Import Excel",
                "Sheet 'DaftarPegawai' ditemukan, tetapi tidak ada data pegawai yang terbaca.",
                parent=parent
            )
            return

        tanpa_tarif = sum(1 for r in rows if r["tarif"] <= 0)
        peringatan_tarif = (
            f"\n\nPERHATIAN: {tanpa_tarif} pegawai tidak memiliki tarif valid. "
            "Tarif master berdasarkan golongan akan digunakan jika tersedia."
            if tanpa_tarif else ""
        )

        jawab = messagebox.askyesno(
            "Ganti Seluruh Data Pegawai",
            f"File: {path}\n"
            f"Sheet: {sheet_target}\n"
            f"Header ditemukan di baris: {header_row_no}\n"
            f"Pegawai terbaca: {len(rows)} orang\n\n"
            "IMPORT INI AKAN MENGGANTIKAN SELURUH DATA PEGAWAI LAMA.\n"
            "Semua keanggotaan dan jadwal yang terkait pegawai lama juga akan dihapus.\n"
            "Data kegiatan/periode tetap ada, tetapi anggotanya akan kosong.\n\n"
            "Pastikan data lama sudah di-backup/export sebelum melanjutkan."
            + peringatan_tarif
            + "\n\nLanjutkan?",
            parent=parent
        )
        if not jawab:
            return

        konfirmasi = simpledialog.askstring(
            "Konfirmasi Penggantian Data",
            "Ketik GANTI untuk menghapus data pegawai lama dan memasukkan data Excel baru:",
            parent=parent
        )
        if konfirmasi != "GANTI":
            messagebox.showinfo(
                "Dibatalkan",
                "Import dibatalkan. Tidak ada data yang diubah.",
                parent=parent
            )
            return

        conn = get_connection()
        try:
            # Hapus relasi lama terlebih dahulu. Master pegawai baru kemudian dimasukkan.
            conn.execute("UPDATE bagian SET penanggung_jawab_id = NULL")
            conn.execute("DELETE FROM jadwal")
            conn.execute("DELETE FROM anggota_bagian")
            conn.execute("DELETE FROM pegawai")

            tarif_master = {
                str(r[0]).strip().upper(): int(r[1] or 0)
                for r in conn.execute("SELECT golongan, nominal FROM tarif").fetchall()
            }

            for idx, item in enumerate(rows, start=1):
                tarif = item["tarif"]
                if tarif <= 0:
                    tarif = tarif_master.get(item["golongan_utama"], 0)

                conn.execute(
                    """
                    INSERT INTO pegawai
                    (no, nama, nip, pangkat, golongan_asli, golongan_utama,
                     jabatan, tarif, aktif)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, 1)
                    """,
                    (
                        idx,
                        item["nama"],
                        item["nip"],
                        item["pangkat"],
                        item["golongan"],
                        item["golongan_utama"],
                        item["jabatan"],
                        tarif
                    )
                )
            conn.commit()
        except Exception as exc:
            conn.rollback()
            messagebox.showerror(
                "Import Excel Gagal",
                "Import gagal dan tidak ada perubahan yang disimpan. Periksa file lalu coba lagi.",
                parent=parent
            )
            return
        finally:
            conn.close()

        # Semua tampilan di-refresh dari database setelah import.
        # Jangan memanggil refresh satu per satu di sini karena itu membuat
        # kalender/tabel dibangun berulang kali dan terasa berat.
        self.selected_ids.clear()
        self.refresh_all()

        messagebox.showinfo(
            "Import Berhasil",
            f"Berhasil mengimpor {len(rows)} pegawai dari sheet '{sheet_target}'.\n\n"
            "Nama file bebas; yang diwajibkan adalah nama sheet 'DaftarPegawai'.\n"
            "Keanggotaan dan jadwal lama sudah dikosongkan. Data kegiatan/periode tetap ada.",
            parent=parent
        )

    def refresh_data_pegawai(self):

        try:
            self.pegawai = load_pegawai()
        except Exception:
            pass

        if not hasattr(self, "table_pegawai") or self.table_pegawai is None:
            return

        try:
            if not self.table_pegawai.winfo_exists():
                self.table_pegawai = None
                return
        except tk.TclError:
            self.table_pegawai = None
            return

        try:
            for item in self.table_pegawai.get_children():
                self.table_pegawai.delete(item)
        except tk.TclError:
            self.table_pegawai = None
            return

        if hasattr(self, "label_jumlah_pegawai"):
            try:
                self.label_jumlah_pegawai.config(
                    text=f"{len(self.pegawai)} pegawai aktif"
                )
            except tk.TclError:
                pass

        riwayat = load_riwayat_turun(self.periode_id)

        keyword = ""
        if hasattr(self, "search_data_pegawai_var") and self.search_data_pegawai_var:
            keyword = self.search_data_pegawai_var.get().strip().lower()

        tampil_nonaktif = bool(
            getattr(self, "tampilkan_nonaktif_var", tk.BooleanVar(value=False)).get()
        )

        self.table_pegawai.tag_configure("baris1", background=THEME_PALETTES[getattr(self, "current_theme", "light")]["UI_PANEL"])
        self.table_pegawai.tag_configure("baris2", background=THEME_PALETTES[getattr(self, "current_theme", "light")]["UI_TABLE_ALT"])
        self.table_pegawai.tag_configure("nonaktif", foreground="#888888", background=THEME_PALETTES[getattr(self, "current_theme", "light")]["UI_TABLE_HEAD"])

        # Ambil semua pegawai agar checkbox "Tampilkan nonaktif" benar-benar
        # dapat menampilkan pegawai yang sudah dinonaktifkan.
        conn = get_connection()
        try:
            semua_pegawai = conn.execute(
                """
                SELECT id, nama, nip, pangkat, golongan_asli,
                       golongan_utama, jabatan, tarif, aktif
                FROM pegawai
                ORDER BY nama
                """
            ).fetchall()
        finally:
            conn.close()

        nomor_baris = 1
        for p in semua_pegawai:
            aktif = int(p[8] or 0) == 1
            if not aktif and not tampil_nonaktif:
                continue

            teks_pencarian = " ".join([
                str(p[1] or ""),
                str(p[2] or ""),
                str(p[3] or ""),
                str(p[4] or ""),
                str(p[5] or ""),
                str(p[6] or "")
            ]).lower()

            if keyword and keyword not in teks_pencarian:
                continue

            pegawai_id = p[0]
            data = riwayat.get(
                pegawai_id,
                {"bulan_ini": 0, "total": 0}
            )

            tag = "baris1" if nomor_baris % 2 == 1 else "baris2"
            if not aktif:
                tag = "nonaktif"

            self.table_pegawai.insert(
                "",
                "end",
                iid=str(pegawai_id),
                values=(
                    nomor_baris,
                    p[1],
                    p[2],
                    p[3],
                    p[4],
                    p[6],
                    rupiah(p[7]),
                    f"{data['bulan_ini']} kali",
                    "Aktif" if aktif else "Nonaktif"
                ),
                tags=(tag,)
            )
            nomor_baris += 1

    # =====================================================
    # EXPORT EXCEL
    # =====================================================

    @staticmethod
    def _excel_sheet_name(name, fallback="Sheet"):
        """Membersihkan nama sheet Excel (maks. 31 karakter)."""
        name = str(name or "").strip()
        name = re.sub(r'[\\/*?:\[\]]', '-', name)
        name = name or fallback
        return name[:31]

    @staticmethod
    def _excel_unique_sheet_name(wb, name):
        """Memastikan nama worksheet unik dan tetap <= 31 karakter."""
        base = App._excel_sheet_name(name)
        if base not in wb.sheetnames:
            return base
        nomor = 2
        while True:
            suffix = f" ({nomor})"
            kandidat = base[:31 - len(suffix)] + suffix
            if kandidat not in wb.sheetnames:
                return kandidat
            nomor += 1

    @staticmethod
    def _format_excel_sheet(ws, header_row=None, widths=None):
        """Format umum workbook hasil export."""
        thin = Side(style="thin", color="D9D9D9")
        border = Border(left=thin, right=thin, top=thin, bottom=thin)
        ws.sheet_view.showGridLines = False

        if header_row:
            for cell in ws[header_row]:
                if cell.value is not None:
                    cell.font = Font(bold=True, color="FFFFFF")
                    cell.fill = PatternFill("solid", fgColor="17324D")
                    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                    cell.border = border
            ws.row_dimensions[header_row].height = 28
            ws.freeze_panes = None
            ws.auto_filter.ref = f"A{header_row}:{get_column_letter(ws.max_column)}{ws.max_row}"

        for row in ws.iter_rows():
            for cell in row:
                if cell.value is not None:
                    cell.border = border
                    if cell.row != header_row:
                        cell.alignment = Alignment(vertical="center", wrap_text=True)

        if widths:
            for col, width in widths.items():
                ws.column_dimensions[col].width = width

    def _add_excel_table(self, ws, ref, name):
        """Tambahkan Excel Table dengan referensi yang valid dan nama unik.

        Tabel hanya dibuat jika memiliki minimal satu baris data selain header.
        Ini mencegah Excel menampilkan pesan Repair pada workbook tertentu.
        """
        if Table is None or TableStyleInfo is None:
            return
        try:
            # Ref berbentuk A7:M20 -> pastikan baris akhir lebih besar dari header.
            m = re.match(r"^[A-Z]+(\d+):[A-Z]+(\d+)$", str(ref))
            if not m or int(m.group(2)) <= int(m.group(1)):
                return

            # Nama tabel Excel harus unik dalam satu workbook dan hanya memakai
            # karakter yang aman.
            safe_name = re.sub(r"[^A-Za-z0-9_]", "_", str(name))
            if not safe_name or safe_name[0].isdigit():
                safe_name = "tbl_" + safe_name
            existing = {t.name for ws0 in ws.parent.worksheets for t in getattr(ws0, "tables", {}).values()}
            base = safe_name
            n = 2
            while safe_name in existing:
                safe_name = f"{base}_{n}"
                n += 1

            table = Table(displayName=safe_name, ref=ref)
            table.tableStyleInfo = TableStyleInfo(
                name="TableStyleMedium2",
                showFirstColumn=False,
                showLastColumn=False,
                showRowStripes=True,
                showColumnStripes=False,
            )
            # openpyxl akan membuat AutoFilter tabel sendiri; jangan membuat
            # AutoFilter sheet kedua pada rentang yang sama.
            ws.add_table(table)
        except Exception:
            # Export tetap dapat dibuka walaupun Excel Table tidak dapat dibuat.
            pass

    def _get_export_target_summary(self, conn, bagian_id, periode_id):
        """Ambil target turun + realisasi + uang untuk satu kegiatan."""
        row = conn.execute(
            "SELECT COALESCE(komposisi_target, '') FROM bagian WHERE id=?",
            (bagian_id,)
        ).fetchone()
        target = {}
        target_tarif = {}
        try:
            stored = json.loads(row[0]) if row and row[0] else {}
            if isinstance(stored, dict) and "target" in stored:
                target = {str(k): int(v) for k, v in (stored.get("target") or {}).items()}
                target_tarif = {str(k): int(v) for k, v in (stored.get("tarif") or {}).items()}
            elif isinstance(stored, dict):
                target = {str(k): int(v) for k, v in stored.items()}
        except Exception:
            target, target_tarif = {}, {}

        # Kelompok realisasi berdasarkan tarif. Golongan dengan tarif sama
        # memang harus dihitung sebagai satu kelompok target.
        actual_rows = conn.execute(
            """SELECT p.tarif, COUNT(*), COALESCE(SUM(p.tarif),0)
               FROM jadwal j JOIN pegawai p ON p.id=j.pegawai_id
               WHERE j.periode_id=? AND j.bagian_id=?
               GROUP BY p.tarif""",
            (periode_id, bagian_id)
        ).fetchall()
        actual = {int(t or 0): (int(n), int(u or 0)) for t, n, u in actual_rows}

        details = []
        total_target_uang = 0
        total_target_turun = 0
        total_realisasi_turun = 0
        total_realisasi_uang = sum(v[1] for v in actual.values())
        for label, jumlah in sorted(target.items()):
            tarif = int(target_tarif.get(label, 0))
            subtotal = int(jumlah) * tarif
            got, uang = actual.get(tarif, (0, 0))
            kurang = max(0, int(jumlah) - got)
            lebih = max(0, got - int(jumlah))
            details.append({
                "label": label, "target": int(jumlah), "tarif": tarif,
                "target_uang": subtotal, "realisasi": got,
                "realisasi_uang": uang, "kurang": kurang, "lebih": lebih
            })
            total_target_turun += int(jumlah)
            total_target_uang += subtotal

        total_realisasi_turun = sum(v[0] for v in actual.values())
        return {
            "details": details,
            "target_turun": total_target_turun,
            "target_uang": total_target_uang,
            "realisasi_turun": total_realisasi_turun,
            "realisasi_uang": total_realisasi_uang,
        }

    def _export_calendar_sheet(self, wb, conn, bulan, tahun, periode_id):
        """Export kalender keseluruhan lengkap dengan desa + ringkasan uang/target."""
        ws = wb.active
        ws.title = "Kalender Keseluruhan"
        ws.sheet_view.zoomScale = 60
        ws.sheet_view.showGridLines = False

        jumlah_hari = calendar.monthrange(tahun, bulan)[1]
        hari_libur = load_hari_libur(periode_id)

        kegiatan_rows = conn.execute(
            """SELECT id, kode, COALESCE(nama_bagian,''), COALESCE(kegiatan,''),
                      COALESCE(dana,0)
               FROM bagian WHERE periode_id=? ORDER BY id""", (periode_id,)
        ).fetchall()
        pegawai_rows = conn.execute(
            "SELECT id,nama,nip,pangkat,golongan_asli,jabatan,tarif FROM pegawai WHERE aktif=1 ORDER BY nama"
        ).fetchall()
        jadwal_rows = conn.execute(
            """SELECT j.pegawai_id,j.hari,j.bagian_id,b.kode,COALESCE(b.nama_bagian,''),COALESCE(d.nama,'')
               FROM jadwal j JOIN bagian b ON b.id=j.bagian_id
               LEFT JOIN desa d ON d.id=j.desa_id
               WHERE j.periode_id=? ORDER BY j.pegawai_id,j.hari""", (periode_id,)
        ).fetchall()

        nama_kegiatan = {r[1]: (r[2] or '').strip() or f"Kegiatan {r[1]}" for r in kegiatan_rows}
        jadwal_map = {(r[0],r[1]): {"kode":r[3],"nama":nama_kegiatan.get(r[3],f"Kegiatan {r[3]}"),"desa":(r[5] or '').strip()} for r in jadwal_rows}

        total_dana = sum(int(r[4] or 0) for r in kegiatan_rows)
        total_target_turun = total_target_uang = total_real_turun = total_real_uang = 0
        target_blocks=[]
        for r in kegiatan_rows:
            summary=self._get_export_target_summary(conn,r[0],periode_id)
            total_target_turun += summary["target_turun"]
            total_target_uang += summary["target_uang"]
            total_real_turun += summary["realisasi_turun"]
            total_real_uang += summary["realisasi_uang"]
            target_blocks.append((r[1], (r[2] or '').strip() or f"Kegiatan {r[1]}", summary))

        # Ringkasan sebelum tabel kalender.
        ws.merge_cells(start_row=1,start_column=1,end_row=1,end_column=jumlah_hari+1)
        ws.cell(1,1,f"KALENDER PENUGASAN {NAMA_BULAN[bulan].upper()} {tahun}")
        ws.cell(1,1).font=Font(size=16,bold=True,color="FFFFFF")
        ws.cell(1,1).fill=PatternFill("solid",fgColor="17324D")
        ws.cell(1,1).alignment=Alignment(horizontal="center",vertical="center")
        ws.row_dimensions[1].height=32

        summary_lines=[
            ("TOTAL DANA SELURUH KEGIATAN", total_dana),
            ("TOTAL TARGET TURUN", total_target_turun),
            ("TOTAL TARGET UANG", total_target_uang),
            ("TOTAL REALISASI TURUN", total_real_turun),
            ("TOTAL UANG REALISASI", total_real_uang),
            ("SELISIH DANA", total_dana-total_real_uang),
        ]
        for i,(label,val) in enumerate(summary_lines,start=2):
            ws.cell(i,1,label).font=Font(bold=True)
            ws.cell(i,1).fill=PatternFill("solid",fgColor="D9E6F2")
            ws.cell(i,2,val)
            if "TURUN" not in label:
                ws.cell(i,2).number_format='#,##0'
            if "DANA" in label or "UANG" in label or "SELISIH" in label:
                ws.cell(i,2).number_format='"Rp "#,##0'

        # Rincian target tiap kegiatan.
        row=9
        ws.cell(row,1,"RINGKASAN TARGET KOMPOSISI PER KEGIATAN").font=Font(bold=True,size=11)
        row+=1
        for kode,nama,summary in target_blocks:
            ws.cell(row,1,f"{kode} - {nama}").font=Font(bold=True)
            row+=1
            if summary["details"]:
                for d in summary["details"]:
                    status = "terpenuhi" if d["kurang"]==0 and d["lebih"]==0 else (f"kurang {d['kurang']} turun" if d["kurang"] else f"lebih {d['lebih']} turun")
                    ws.cell(row,1,f"Gol {d['label']}")
                    ws.cell(row,2,d["target"]); ws.cell(row,3,d["tarif"]); ws.cell(row,4,d["target_uang"])
                    ws.cell(row,5,d["realisasi"]); ws.cell(row,6,d["realisasi_uang"]); ws.cell(row,7,status)
                    ws.cell(row,3).number_format='"Rp "#,##0'; ws.cell(row,4).number_format='"Rp "#,##0'; ws.cell(row,6).number_format='"Rp "#,##0'
                    row+=1
            else:
                ws.cell(row,1,"Belum ada target komposisi")
                row+=1
            row+=1

        table_header = max(row+1, 12)
        ws.cell(table_header,1,"Nama Pegawai")
        nama_hari=["Sen","Sel","Rab","Kam","Jum","Sab","Min"]
        for hari in range(1,jumlah_hari+1):
            dt=datetime(tahun,bulan,hari)
            ws.cell(table_header,hari+1,f"{hari}\n{nama_hari[dt.weekday()]}")
        for cell in ws[table_header]:
            if cell.value is not None:
                cell.font=Font(bold=True,color="FFFFFF"); cell.fill=PatternFill("solid",fgColor="17324D")
                cell.alignment=Alignment(horizontal="center",vertical="center",wrap_text=True)
        ws.row_dimensions[table_header].height=32

        palette=["B8D4EF","C7E6C5","F3D6A5","D9C2E9","F4B8B8","BFE3E3","E7D6A8","C9D7B8"]
        fills={r[1]:PatternFill("solid",fgColor=palette[i%len(palette)]) for i,r in enumerate(kegiatan_rows)}
        holiday_fill=PatternFill("solid",fgColor="ECECEC")
        thin=Side(style="thin",color="D9D9D9"); border=Border(left=thin,right=thin,top=thin,bottom=thin)
        for ri,p in enumerate(pegawai_rows,start=table_header+1):
            ws.cell(ri,1,p[1] or "Tanpa Nama")
            for hari in range(1,jumlah_hari+1):
                c=ws.cell(ri,hari+1); dt=datetime(tahun,bulan,hari); detail=jadwal_map.get((p[0],hari))
                if dt.weekday()==6 or hari in hari_libur:
                    c.value="LIBUR"; c.fill=holiday_fill
                elif detail:
                    c.value=f"{detail['kode']} - {detail['nama']}\nLokasi: {detail['desa'] or 'Lokasi belum ditentukan'}"; c.fill=fills.get(detail['kode'],PatternFill("solid",fgColor="FFFFFF"))
                c.alignment=Alignment(horizontal="center",vertical="center",wrap_text=True); c.border=border
            ws.cell(ri,1).border=border; ws.cell(ri,1).alignment=Alignment(vertical="center",wrap_text=True)

        end=max(table_header,table_header+len(pegawai_rows))
        self._add_excel_table(ws,f"A{table_header}:{get_column_letter(jumlah_hari+1)}{end}","tblKalenderKeseluruhan")
        ws.freeze_panes=None
        ws.column_dimensions["A"].width=28
        for col in range(2,jumlah_hari+2): ws.column_dimensions[get_column_letter(col)].width=14
        ws.page_setup.orientation=ws.ORIENTATION_LANDSCAPE
        ws.page_setup.fitToPage=True; ws.page_setup.fitToWidth=1; ws.page_setup.fitToHeight=0
        ws.sheet_properties.pageSetUpPr.fitToPage=True

    def _export_data_pegawai_sheet(self, wb, conn):
        """Export data pegawai lengkap dengan total turun dan rincian kegiatan/desa."""
        ws = wb.create_sheet("Data Pegawai")
        ws.sheet_view.zoomScale = 75
        headers = [
            "Nama", "NIP", "Pangkat", "Golongan", "Golongan Utama", "Jabatan",
            "Tarif", "Status", "Turun Bulan Ini", "Total Uang Bulan Ini", "Kegiatan", "Tanggal", "Lokasi"
        ]
        ws.append(headers)

        rows = conn.execute(
            """
            SELECT id, nama, nip, pangkat, golongan_asli, golongan_utama, jabatan, tarif, aktif
            FROM pegawai ORDER BY nama
            """
        ).fetchall()
        for p in rows:
            detail = conn.execute(
                """
                SELECT b.kode, COALESCE(b.nama_bagian,''), j.hari, COALESCE(d.nama,'')
                FROM jadwal j
                JOIN bagian b ON b.id=j.bagian_id
                LEFT JOIN desa d ON d.id=j.desa_id
                WHERE j.pegawai_id=? AND j.periode_id=?
                ORDER BY j.hari, b.id
                """, (p[0], self.periode_id)
            ).fetchall()
            kegiatan = "; ".join(
                f"{r[0]} - {(r[1] or f'Kegiatan {r[0]}')}" for r in detail
            )
            tanggal = ", ".join(str(r[2]) for r in detail) or "-"
            desa = "; ".join(
                f"Tgl {r[2]}: {r[3] or 'Belum ditentukan'}" for r in detail
            ) or "-"
            ws.append([
                p[1], p[2], p[3], p[4], p[5], p[6], p[7],
                "Aktif" if p[8] else "Nonaktif", len(detail), len(detail) * int(p[7] or 0), kegiatan, tanggal, desa
            ])

        for cell in ws[1]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor="17324D")
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        for row in ws.iter_rows(min_row=2, min_col=7, max_col=7):
            row[0].number_format = '#,##0'
        for row in ws.iter_rows(min_row=2, min_col=10, max_col=10):
            row[0].number_format = '"Rp "#,##0'
        self._format_excel_sheet(
            ws, header_row=1,
            widths={"A":30,"B":20,"C":18,"D":12,"E":16,"F":28,"G":14,"H":12,"I":18,"J":20,"K":35,"L":18,"M":45}
        )
        ws.freeze_panes = None
        ws.auto_filter.ref = None
        self._add_excel_table(ws, f"A1:{get_column_letter(ws.max_column)}{ws.max_row}", "tblDataPegawai")

    def _export_data_desa_sheet(self, wb, conn, bulan, tahun, periode_id):
        """Export master desa + penggunaan desa pada periode aktif."""
        ws = wb.create_sheet("Data Lokasi")
        ws.sheet_view.zoomScale = 80
        headers = ["Nama Lokasi", "Status", "Dipakai Bulan Ini", "Tanggal", "Kegiatan", "Pegawai"]
        ws.append(headers)
        rows = conn.execute(
            """
            SELECT d.id, d.nama, d.aktif,
                   COUNT(DISTINCT j.hari) AS jumlah,
                   GROUP_CONCAT(DISTINCT j.hari) AS tanggal
            FROM desa d
            LEFT JOIN jadwal j ON j.desa_id=d.id AND j.periode_id=?
            GROUP BY d.id, d.nama, d.aktif
            ORDER BY d.nama COLLATE NOCASE
            """, (periode_id,)
        ).fetchall()
        for did, nama, aktif, jumlah, tanggal in rows:
            detail = conn.execute(
                """
                SELECT j.hari, b.kode, COALESCE(b.nama_bagian,''), p.nama
                FROM jadwal j
                JOIN bagian b ON b.id=j.bagian_id
                JOIN pegawai p ON p.id=j.pegawai_id
                WHERE j.desa_id=? AND j.periode_id=?
                ORDER BY j.hari, b.id, p.nama
                """, (did, periode_id)
            ).fetchall()
            kegiatan = "; ".join(
                f"Tgl {r[0]}: {r[1]} - {(r[2] or f'Kegiatan {r[1]}')}" for r in detail
            ) or "-"
            pegawai = "; ".join(
                f"Tgl {r[0]}: {r[3]}" for r in detail
            ) or "-"
            ws.append([
                nama, "Aktif" if aktif else "Nonaktif", jumlah or 0,
                ", ".join(sorted({str(r[0]) for r in detail}, key=lambda x:int(x))) or "-",
                kegiatan, pegawai
            ])
        self._format_excel_sheet(
            ws, header_row=1,
            widths={"A":30,"B":12,"C":20,"D":18,"E":55,"F":55}
        )
        ws.freeze_panes = None
        ws.auto_filter.ref = ws.dimensions

    def _export_kegiatan_gabungan_sheet(self, wb, conn, bagian, bulan, tahun, periode_id):
        """Satu sheet per kegiatan: data kegiatan di atas, kalender kegiatan di bawah."""
        bagian_id, kode, nama_bagian, deskripsi, dana, pj_id = bagian
        judul = (nama_bagian or '').strip() or f"Kegiatan {kode}"
        sheet_name = self._excel_unique_sheet_name(wb, judul)
        ws = wb.create_sheet(sheet_name)
        ws.sheet_view.zoomScale = 70
        ws.sheet_view.showGridLines = False

        jumlah_hari = calendar.monthrange(tahun, bulan)[1]
        hari_libur = load_hari_libur(periode_id)
        pj = conn.execute("SELECT nama FROM pegawai WHERE id=?", (pj_id,)).fetchone() if pj_id else None
        pj_nama = pj[0] if pj else "-"

        # =====================================================
        # BAGIAN 1 — DATA KEGIATAN
        # =====================================================
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=13)
        ws.cell(1, 1, f"KEGIATAN {judul.upper()} — {NAMA_BULAN[bulan]} {tahun}")
        ws.cell(1, 1).font = Font(size=16, bold=True, color="FFFFFF")
        ws.cell(1, 1).fill = PatternFill("solid", fgColor="17324D")
        ws.cell(1, 1).alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[1].height = 30

        summary = self._get_export_target_summary(conn, bagian_id, periode_id)
        metadata = [
            ("Penanggung Jawab", pj_nama),
            ("Deskripsi", deskripsi or ""),
            ("Dana Bulanan", dana or 0),
            ("Target Turun", summary["target_turun"]),
            ("Target Uang", summary["target_uang"]),
            ("Realisasi Turun", summary["realisasi_turun"]),
            ("Uang Realisasi", summary["realisasi_uang"]),
            ("Kekurangan Turun", max(0, summary["target_turun"]-summary["realisasi_turun"])),
            ("Sisa Dana", int(dana or 0)-summary["realisasi_uang"]),
            ("Periode", f"{NAMA_BULAN[bulan]} {tahun}"),
        ]
        for r, (label, value) in enumerate(metadata, start=2):
            ws.cell(r, 1, label).font = Font(bold=True)
            ws.cell(r, 1).fill = PatternFill("solid", fgColor="D9E6F2")
            ws.cell(r, 2, value)
            if any(k in label for k in ("Dana", "Uang", "Sisa")):
                ws.cell(r, 2).number_format = '"Rp "#,##0'

        target_start = 2
        target_col = 4
        ws.cell(target_start, target_col, "TARGET KOMPOSISI TURUN").font = Font(bold=True, size=11)
        for idx, d in enumerate(summary["details"], start=target_start+1):
            status = "terpenuhi" if d["kurang"] == 0 and d["lebih"] == 0 else (f"kurang {d['kurang']} turun" if d["kurang"] else f"lebih {d['lebih']} turun")
            ws.cell(idx, target_col, f"Gol {d['label']}: {d['target']} turun × {d['tarif']}")
            ws.cell(idx, target_col+1, d["target_uang"])
            ws.cell(idx, target_col+2, f"Realisasi {d['realisasi']} turun — {status}")
            ws.cell(idx, target_col+1).number_format = '"Rp "#,##0'
        anggota = conn.execute(
            """SELECT p.id,p.nama,p.nip,p.pangkat,p.golongan_asli,p.golongan_utama,p.jabatan,p.tarif
               FROM anggota_bagian ab JOIN pegawai p ON p.id=ab.pegawai_id
               WHERE ab.bagian_id=? ORDER BY p.nama""", (bagian_id,)
        ).fetchall()

        header_data = max(13, 5 + len(summary["details"]))
        headers = ["No", "Nama", "NIP", "Pangkat", "Golongan", "Golongan Utama", "Jabatan", "Tarif", "PJ", "Turun", "Total Uang", "Tanggal", "Lokasi"]
        for c, h in enumerate(headers, start=1):
            ws.cell(header_data, c, h)
        self._format_excel_sheet(
            ws, header_row=header_data,
            widths={"A":6,"B":27,"C":20,"D":16,"E":12,"F":16,"G":25,"H":13,"I":8,"J":9,"K":15,"L":18,"M":42}
        )

        for nomor, p in enumerate(anggota, start=1):
            detail = conn.execute(
                """SELECT j.hari,COALESCE(d.nama,'') FROM jadwal j
                   LEFT JOIN desa d ON d.id=j.desa_id
                   WHERE j.bagian_id=? AND j.pegawai_id=? AND j.periode_id=? ORDER BY j.hari""",
                (bagian_id, p[0], periode_id)
            ).fetchall()
            tanggal = ", ".join(str(r[0]) for r in detail) or "-"
            desa = "; ".join(f"Tgl {r[0]}: {r[1] or 'Belum ditentukan'}" for r in detail) or "-"
            turun = len(detail)
            total = turun * (p[7] or 0)
            ws.append([nomor, p[1], p[2], p[3], p[4], p[5], p[6], p[7], "Ya" if p[0] == pj_id else "", turun, total, tanggal, desa])

        data_end = max(header_data, ws.max_row)
        ws.freeze_panes = None
        for row in ws.iter_rows(min_row=header_data + 1, max_row=data_end):
            for cell in row:
                cell.border = Border(left=Side(style="thin", color="D9D9D9"), right=Side(style="thin", color="D9D9D9"), top=Side(style="thin", color="D9D9D9"), bottom=Side(style="thin", color="D9D9D9"))
                cell.alignment = Alignment(vertical="center", wrap_text=True)
        for row in ws.iter_rows(min_row=header_data + 1, min_col=8, max_col=8):
            row[0].number_format = '#,##0'
        for row in ws.iter_rows(min_row=header_data + 1, min_col=11, max_col=11):
            row[0].number_format = '#,##0'
        ws.auto_filter.ref = None
        self._add_excel_table(ws, f"A{header_data}:M{data_end}", f"tblKegiatan{bagian_id}Data")

        # Ringkasan akhir data kegiatan: mudah dibaca tanpa membuka aplikasi.
        recap_row = data_end + 2
        ws.cell(recap_row, 1, "RINGKASAN TARGET & REALISASI").font = Font(bold=True, size=11)
        recap_headers = ["Kelompok Tarif/Golongan", "Target Turun", "Tarif", "Target Uang", "Realisasi Turun", "Realisasi Uang", "Status"]
        for c,h in enumerate(recap_headers, start=1):
            ws.cell(recap_row+1,c,h)
        for rr,d in enumerate(summary["details"], start=recap_row+2):
            status = "Terpenuhi" if d["kurang"]==0 and d["lebih"]==0 else (f"Kurang {d['kurang']} turun" if d["kurang"] else f"Lebih {d['lebih']} turun")
            vals=[f"Gol {d['label']}",d["target"],d["tarif"],d["target_uang"],d["realisasi"],d["realisasi_uang"],status]
            for c,v in enumerate(vals,start=1): ws.cell(rr,c,v)
            ws.cell(rr,3).number_format='"Rp "#,##0'; ws.cell(rr,4).number_format='"Rp "#,##0'; ws.cell(rr,6).number_format='"Rp "#,##0'
        total_row=recap_row+2+len(summary["details"])
        ws.cell(total_row,1,"TOTAL")
        ws.cell(total_row,2,summary["target_turun"]); ws.cell(total_row,4,summary["target_uang"])
        ws.cell(total_row,5,summary["realisasi_turun"]); ws.cell(total_row,6,summary["realisasi_uang"])
        ws.cell(total_row,4).number_format='"Rp "#,##0'; ws.cell(total_row,6).number_format='"Rp "#,##0'
        ws.cell(total_row+1,1,"DANA BULANAN"); ws.cell(total_row+1,2,dana or 0); ws.cell(total_row+1,2).number_format='"Rp "#,##0'
        ws.cell(total_row+2,1,"SISA DANA"); ws.cell(total_row+2,2,int(dana or 0)-summary["realisasi_uang"]); ws.cell(total_row+2,2).number_format='"Rp "#,##0'
        for r in range(recap_row+1,total_row+3):
            for c in range(1,8):
                ws.cell(r,c).border=Border(left=Side(style="thin",color="D9D9D9"),right=Side(style="thin",color="D9D9D9"),top=Side(style="thin",color="D9D9D9"),bottom=Side(style="thin",color="D9D9D9"))
                ws.cell(r,c).alignment=Alignment(vertical="center",wrap_text=True)

        # =====================================================
        # BAGIAN 2 — KALENDER KEGIATAN, DI BAWAH DATA KEGIATAN
        # =====================================================
        gap = 3
        title_row = total_row + 3
        ws.merge_cells(start_row=title_row, start_column=1, end_row=title_row, end_column=jumlah_hari + 1)
        ws.cell(title_row, 1, f"KALENDER {judul.upper()} — {NAMA_BULAN[bulan]} {tahun}")
        ws.cell(title_row, 1).font = Font(size=15, bold=True, color="FFFFFF")
        ws.cell(title_row, 1).fill = PatternFill("solid", fgColor="17324D")
        ws.cell(title_row, 1).alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[title_row].height = 30

        meta_row = title_row + 1
        ws.cell(meta_row, 1, "Penanggung Jawab").font = Font(bold=True)
        ws.cell(meta_row, 2, pj_nama)
        ws.cell(meta_row + 1, 1, "Deskripsi").font = Font(bold=True)
        ws.cell(meta_row + 1, 2, deskripsi or "")

        header_cal = title_row + 3
        ws.cell(header_cal, 1, "Nama Pegawai")
        nama_hari = ["Sen", "Sel", "Rab", "Kam", "Jum", "Sab", "Min"]
        for hari in range(1, jumlah_hari + 1):
            dt = datetime(tahun, bulan, hari)
            c = ws.cell(header_cal, hari + 1, f"{hari}\n{nama_hari[dt.weekday()]}")
            c.font = Font(bold=True, color="FFFFFF")
            c.fill = PatternFill("solid", fgColor="17324D")
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.row_dimensions[header_cal].height = 30

        jadwal = conn.execute(
            """SELECT j.pegawai_id,j.hari,COALESCE(d.nama,'') FROM jadwal j
               LEFT JOIN desa d ON d.id=j.desa_id
               WHERE j.bagian_id=? AND j.periode_id=?""", (bagian_id, periode_id)
        ).fetchall()
        jmap = {(r[0], r[1]): r[2] for r in jadwal}

        holiday_fill = PatternFill("solid", fgColor="ECECEC")
        fill = PatternFill("solid", fgColor="B8D4EF")
        thin = Side(style="thin", color="D9D9D9")
        border = Border(left=thin, right=thin, top=thin, bottom=thin)
        for rr, p in enumerate(anggota, start=header_cal + 1):
            ws.cell(rr, 1, p[1])
            ws.cell(rr, 1).border = border
            ws.cell(rr, 1).alignment = Alignment(vertical="center", wrap_text=True)
            for hari in range(1, jumlah_hari + 1):
                c = ws.cell(rr, hari + 1)
                dt = datetime(tahun, bulan, hari)
                desa = jmap.get((p[0], hari))
                if dt.weekday() == 6 or hari in hari_libur:
                    c.value = "LIBUR"
                    c.fill = holiday_fill
                elif desa is not None:
                    c.value = f"{kode}\nLokasi: {desa or 'Belum ditentukan'}"
                    c.fill = fill
                else:
                    c.value = ""
                c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                c.border = border

        calendar_end = max(header_cal, header_cal + len(anggota))
        self._add_excel_table(ws, f"A{header_cal}:{get_column_letter(jumlah_hari + 1)}{calendar_end}", f"tblKegiatan{bagian_id}Kal")
        # Tidak memakai freeze panes di sheet gabungan: setiap tabel memiliki headernya sendiri
        # dan pengguna dapat menggulir dari atas sampai bawah tanpa terkunci pada tabel pertama.
        ws.column_dimensions["A"].width = 27
        for col in range(2, jumlah_hari + 2):
            ws.column_dimensions[get_column_letter(col)].width = 12
        ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
        ws.page_setup.fitToPage = True
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
        ws.sheet_properties.pageSetUpPr.fitToPage = True

    
    def export_penugasan_excel(self):
        """Export satu periode secara lengkap: kalender, pegawai, desa, dan setiap kegiatan."""
        if Workbook is None:
            messagebox.showerror("Export Excel","Library openpyxl belum terpasang. Jalankan: pip install openpyxl",parent=self)
            return
        try:
            bulan=NAMA_BULAN.index(self.bulan_var.get()); tahun=int(self.tahun_var.get())
        except Exception:
            messagebox.showerror("Export Excel","Bulan atau tahun tidak valid.",parent=self); return
        periode_id=get_or_create_periode(bulan,tahun)
        default_name=f"Penugasan_{NAMA_BULAN[bulan]}_{tahun}.xlsx"
        path=filedialog.asksaveasfilename(parent=self,title="Simpan Export Penugasan",defaultextension=".xlsx",initialfile=default_name,filetypes=[("Excel Workbook","*.xlsx")])
        if not path: return
        try:
            wb=Workbook(); conn=get_connection()
            try:
                self._export_calendar_sheet(wb,conn,bulan,tahun,periode_id)
                self._export_data_pegawai_sheet(wb,conn)
                self._export_data_desa_sheet(wb,conn,bulan,tahun,periode_id)
                bagian_rows=conn.execute(
                    """SELECT id,kode,COALESCE(nama_bagian,''),COALESCE(kegiatan,''),COALESCE(dana,0),penanggung_jawab_id
                       FROM bagian WHERE periode_id=? ORDER BY id""",(periode_id,)
                ).fetchall()
                for bagian in bagian_rows:
                    self._export_kegiatan_gabungan_sheet(wb, conn, bagian, bulan, tahun, periode_id)
            finally:
                conn.close()
            wb.save(path); wb.close()
            jumlah_sheet=3+len(bagian_rows)
            messagebox.showinfo("Export Berhasil",f"File penugasan berhasil dibuat.\n\n{path}\n\nIsi: Kalender Keseluruhan + Data Pegawai + Data Lokasi + satu sheet per kegiatan dengan Data Kegiatan + Kalender Kegiatan ({jumlah_sheet} sheet).",parent=self)
        except Exception as exc:
            try: wb.close()
            except Exception: pass
            messagebox.showerror("Export Gagal","File Excel tidak dapat dibuat. Silakan coba lagi.",parent=self)

    # =====================================================
    # KELOLA / BERSIHKAN DATA PENUGASAN
    # =====================================================

    def buka_data_desa(self):
        """Master Data Desa dengan fungsi yang setara dengan Data Pegawai."""
        if getattr(self, "data_desa_window", None) is not None:
            try:
                if self.data_desa_window.winfo_exists():
                    self.data_desa_window.lift()
                    self.data_desa_window.focus_force()
                    self.refresh_data_desa()
                    return
            except tk.TclError:
                pass

        win = tk.Toplevel(self)
        self.data_desa_window = win
        self.modernize_window(win, "Data Lokasi", (760, 500))
        win.title("Data Lokasi")
        # Toplevel ikut menyesuaikan ukuran monitor/DPI dan tetap bisa di-resize.
        sw = max(900, int(win.winfo_screenwidth() * 0.78))
        sh = max(560, int(win.winfo_screenheight() * 0.78))
        win.geometry(f"{sw}x{sh}")
        win.minsize(700, 450)
        self.enable_native_window_controls(win)
        win.protocol("WM_DELETE_WINDOW", lambda: self.tutup_data_desa(win))
        self.center_window(win, self)

        tk.Label(
            win, text="DATA LOKASI", bg=UI_BG,
            font=("Segoe UI", 17, "bold")
        ).pack(anchor="w", padx=20, pady=(15, 5))

        self.label_jumlah_desa = tk.Label(
            win, text="", bg=UI_BG, fg=UI_MUTED,
            font=("Segoe UI", 10)
        )
        self.label_jumlah_desa.pack(anchor="w", padx=20)

        toolbar = tk.Frame(win, bg=UI_BG)
        toolbar.pack(fill="x", padx=20, pady=(12, 5))
        ttk.Button(toolbar, text="＋ Tambah Lokasi", command=self.tambah_desa).pack(side="left", padx=(0, 6))
        ttk.Button(toolbar, text="✎ Edit", command=self.edit_desa_terpilih).pack(side="left", padx=6)
        ttk.Button(toolbar, text="🗑 Nonaktifkan", command=self.nonaktifkan_desa_terpilih).pack(side="left", padx=6)
        ttk.Button(toolbar, text="🗑 Hapus Permanen", command=self.hapus_desa_permanen_terpilih).pack(side="left", padx=6)
        ttk.Button(toolbar, text="↥ Impor dari Excel", command=self.impor_desa_dari_excel).pack(side="left", padx=6)
        ttk.Button(toolbar, text="✓ Aktifkan", command=self.aktifkan_desa_terpilih).pack(side="left", padx=6)
        ttk.Button(toolbar, text="↻ Refresh", command=self.refresh_data_desa).pack(side="left", padx=6)

        self.tampilkan_nonaktif_desa_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            toolbar, text="Tampilkan lokasi nonaktif",
            variable=self.tampilkan_nonaktif_desa_var,
            command=self.refresh_data_desa
        ).pack(side="right", padx=(10, 0))

        search = tk.Frame(win, bg=UI_BG)
        search.pack(fill="x", padx=20, pady=(5, 5))
        tk.Label(search, text="Cari Lokasi:", bg=UI_BG, font=("Segoe UI", 10, "bold")).pack(side="left", padx=(0, 8))
        self.search_data_desa_var = tk.StringVar()
        ent = ttk.Entry(search, textvariable=self.search_data_desa_var, width=45)
        ent.pack(side="left")
        ent.bind("<KeyRelease>", lambda event: self.refresh_data_desa())
        ttk.Button(search, text="✕", width=3, command=lambda: (self.search_data_desa_var.set(""), self.refresh_data_desa())).pack(side="left", padx=5)

        frame = tk.Frame(win, bg=UI_PANEL)
        frame.pack(fill="both", expand=True, padx=20, pady=15)
        columns = ("no", "nama", "status", "turun", "tanggal", "kegiatan")
        self.table_desa = ttk.Treeview(frame, columns=columns, show="headings", selectmode="browse")
        self.table_desa.heading("no", text="No.")
        self.table_desa.heading("nama", text="Nama Lokasi")
        self.table_desa.heading("status", text="Status")
        self.table_desa.heading("turun", text="Dipakai Bulan Ini")
        self.table_desa.heading("tanggal", text="Tanggal")
        self.table_desa.heading("kegiatan", text="Kegiatan")
        self.table_desa.column("no", width=58, anchor="center", stretch=False)
        self.table_desa.column("nama", width=300, anchor="w")
        self.table_desa.column("status", width=100, anchor="center")
        self.table_desa.column("turun", width=140, anchor="center")
        self.table_desa.column("tanggal", width=220, anchor="w")
        self.table_desa.column("kegiatan", width=360, anchor="w")
        frame.grid_rowconfigure(0, weight=1)
        frame.grid_columnconfigure(0, weight=1)
        self.table_desa.grid(row=0, column=0, sticky="nsew")
        sc = tk.Scrollbar(
            frame, orient="vertical", command=self.table_desa.yview,
            width=18, bg="#94A3B8", troughcolor="#E2E8F0",
            activebackground="#64748B", relief="flat", bd=0
        )
        sc.grid(row=0, column=1, sticky="ns")
        scx = tk.Scrollbar(
            frame, orient="horizontal", command=self.table_desa.xview,
            width=18, bg="#94A3B8", troughcolor="#E2E8F0",
            activebackground="#64748B", relief="flat", bd=0
        )
        scx.grid(row=1, column=0, sticky="ew")
        self.table_desa.configure(yscrollcommand=sc.set, xscrollcommand=scx.set)
        self.modernize_tree(self.table_desa)
        self.table_desa.bind("<MouseWheel>", lambda e: self.table_desa.yview_scroll(int(-e.delta / 120), "units"), add="+")
        self.table_desa.bind("<Shift-MouseWheel>", lambda e: self.table_desa.xview_scroll(int(-e.delta / 120), "units"), add="+")
        self.table_desa.bind("<Double-1>", lambda event: self.edit_desa_terpilih())

        self.refresh_data_desa()

    def tutup_data_desa(self, win):
        try:
            if win.winfo_exists():
                win.destroy()
        finally:
            self.data_desa_window = None
            self.table_desa = None
            self.search_data_desa_var = None
            self.label_jumlah_desa = None

    def refresh_setelah_perubahan_desa(self):
        """Sinkronkan seluruh tampilan setelah master desa berubah."""
        # Refresh tabel master desa yang sedang terbuka.
        self.refresh_data_desa()

        # Refresh semua tampilan yang bergantung pada data desa/jadwal.
        # Ini juga memperbarui kalender kegiatan dan kalender keseluruhan.
        try:
            self.refresh_all()
        except Exception:
            # refresh_all menampilkan error sendiri; kalender aktif tetap
            # dicoba dimuat ulang agar perubahan desa langsung terlihat.
            try:
                self.load_kalender()
            except Exception:
                pass
            try:
                self.load_kalender_keseluruhan()
            except Exception:
                pass

    def refresh_data_desa(self):
        if not hasattr(self, "table_desa") or self.table_desa is None:
            return
        try:
            if not self.table_desa.winfo_exists():
                self.table_desa = None
                return
        except tk.TclError:
            self.table_desa = None
            return

        for item in self.table_desa.get_children():
            self.table_desa.delete(item)

        keyword = ""
        if getattr(self, "search_data_desa_var", None) is not None:
            keyword = self.search_data_desa_var.get().strip()
        tampil_nonaktif = bool(getattr(self, "tampilkan_nonaktif_desa_var", tk.BooleanVar(value=False)).get())

        conn = get_connection()
        try:
            sql = """
                SELECT d.id, d.nama, d.aktif,
                       COUNT(DISTINCT j.hari) AS jumlah_dipakai,
                       GROUP_CONCAT(DISTINCT j.hari) AS tanggal,
                       GROUP_CONCAT(DISTINCT
                           CASE
                               WHEN COALESCE(b.nama_bagian,'') <> '' THEN
                                   b.kode || ' - ' || b.nama_bagian
                               ELSE
                                   'Kegiatan ' || b.kode
                           END
                       ) AS kegiatan
                FROM desa d
                LEFT JOIN jadwal j
                    ON j.desa_id=d.id AND j.periode_id=?
                LEFT JOIN bagian b
                    ON b.id=j.bagian_id
                WHERE 1=1
            """
            # Jangan menghapus parameter periode_id. Bug sebelumnya membuat
            # query menerima jumlah parameter yang tidak sesuai sehingga rekap
            # desa kosong.
            params = [self.periode_id]
            if not tampil_nonaktif:
                sql += " AND d.aktif = 1"
            if keyword:
                sql += " AND d.nama LIKE ?"
                params.append(f"%{keyword}%")
            sql += " GROUP BY d.id, d.nama, d.aktif ORDER BY d.nama COLLATE NOCASE"
            rows = conn.execute(sql, params).fetchall()
        finally:
            conn.close()

        aktif = 0
        for nomor_desa, (did, nama, status, jumlah_dipakai, tanggal, kegiatan) in enumerate(rows, start=1):
            if status:
                aktif += 1
            try:
                tanggal_list = sorted({int(x) for x in (tanggal or "").split(",") if str(x).strip()})
                tanggal_text = ", ".join(str(x) for x in tanggal_list) if tanggal_list else "-"
            except Exception:
                tanggal_text = tanggal or "-"
            self.table_desa.insert(
                "", "end", iid=str(did),
                values=(
                    nomor_desa,
                    nama,
                    "Aktif" if status else "Nonaktif",
                    f"{jumlah_dipakai} kali",
                    tanggal_text,
                    kegiatan or "-"
                )
            )

        total = len(rows)
        if getattr(self, "label_jumlah_desa", None) is not None:
            self.label_jumlah_desa.config(text=f"{total} desa ditampilkan  •  {aktif} aktif")

    def _desa_terpilih(self):
        if not hasattr(self, "table_desa") or self.table_desa is None:
            return None
        sel = self.table_desa.selection()
        return int(sel[0]) if sel else None

    def tambah_desa(self):
        parent = getattr(self, "data_desa_window", self)
        nama = simpledialog.askstring("Tambah Lokasi", "Nama lokasi:", parent=parent)
        if not nama or not nama.strip():
            return
        nama = nama.strip()
        conn = get_connection()
        try:
            conn.execute("INSERT INTO desa(nama, aktif) VALUES(?, 1)", (nama,))
            conn.commit()
        except sqlite3.IntegrityError:
            messagebox.showwarning("Data Lokasi", "Nama lokasi sudah ada.", parent=parent)
            return
        finally:
            conn.close()
        self.refresh_setelah_perubahan_desa()

    def edit_desa_terpilih(self):
        parent = getattr(self, "data_desa_window", self)
        desa_id = self._desa_terpilih()
        if not desa_id:
            messagebox.showwarning("Data Lokasi", "Pilih lokasi terlebih dahulu.", parent=parent)
            return
        conn = get_connection()
        row = conn.execute("SELECT nama FROM desa WHERE id=?", (desa_id,)).fetchone()
        conn.close()
        if not row:
            return
        nama_lama = row[0]
        nama = simpledialog.askstring("Edit Lokasi", "Nama lokasi:", initialvalue=nama_lama, parent=parent)
        if nama is None or not nama.strip():
            return
        conn = get_connection()
        try:
            conn.execute("UPDATE desa SET nama=? WHERE id=?", (nama.strip(), desa_id))
            conn.commit()
        except sqlite3.IntegrityError:
            messagebox.showwarning("Data Lokasi", "Nama lokasi sudah digunakan oleh lokasi lain.", parent=parent)
            return
        finally:
            conn.close()
        self.refresh_setelah_perubahan_desa()

    def nonaktifkan_desa_terpilih(self):
        parent = getattr(self, "data_desa_window", self)
        desa_id = self._desa_terpilih()
        if not desa_id:
            messagebox.showwarning("Data Lokasi", "Pilih lokasi terlebih dahulu.", parent=parent)
            return
        if not messagebox.askyesno("Nonaktifkan Lokasi", "Nonaktifkan desa terpilih? Desa tidak akan muncul sebagai pilihan baru pada kalender.", parent=parent):
            return
        conn = get_connection()
        conn.execute("UPDATE desa SET aktif=0 WHERE id=?", (desa_id,))
        conn.commit(); conn.close()
        self.refresh_setelah_perubahan_desa()

    def aktifkan_desa_terpilih(self):
        parent = getattr(self, "data_desa_window", self)
        desa_id = self._desa_terpilih()
        if not desa_id:
            messagebox.showwarning("Data Lokasi", "Pilih lokasi terlebih dahulu.", parent=parent)
            return
        conn = get_connection()
        conn.execute("UPDATE desa SET aktif=1 WHERE id=?", (desa_id,))
        conn.commit(); conn.close()
        self.refresh_setelah_perubahan_desa()

    def hapus_desa_permanen_terpilih(self):
        parent = getattr(self, "data_desa_window", self)
        desa_id = self._desa_terpilih()
        if not desa_id:
            messagebox.showwarning("Data Lokasi", "Pilih lokasi terlebih dahulu.", parent=parent)
            return
        conn = get_connection()
        row = conn.execute("SELECT nama FROM desa WHERE id=?", (desa_id,)).fetchone()
        dipakai = conn.execute("SELECT COUNT(*) FROM jadwal WHERE desa_id=?", (desa_id,)).fetchone()[0]
        conn.close()
        if not row:
            return
        nama = row[0]
        if dipakai:
            messagebox.showwarning(
                "Tidak Dapat Dihapus",
                f"Desa '{nama}' sudah digunakan oleh {dipakai} data penugasan.\n\nNonaktifkan saja agar desa tidak muncul pada pilihan penugasan baru.",
                parent=parent
            )
            return
        if not messagebox.askyesno(
            "Hapus Permanen Lokasi",
            f"Hapus permanen desa '{nama}'?\n\nData yang sudah dihapus tidak dapat dipulihkan dari aplikasi.",
            parent=parent
        ):
            return
        konfirmasi = simpledialog.askstring(
            "Konfirmasi Hapus Permanen",
            "Ketik HAPUS PERMANEN untuk melanjutkan:",
            parent=parent
        )
        if konfirmasi != "HAPUS PERMANEN":
            return
        conn = get_connection()
        try:
            conn.execute("DELETE FROM desa WHERE id=?", (desa_id,))
            conn.commit()
        finally:
            conn.close()
        self.refresh_setelah_perubahan_desa()

    def impor_desa_dari_excel(self):
        """Import master desa dari sheet Nama Desa/NamaDesa dengan data mulai A1."""
        parent = getattr(self, "data_desa_window", self)
        if load_workbook is None:
            messagebox.showerror("Import Excel", "Library openpyxl belum terpasang.", parent=parent)
            return

        path = filedialog.askopenfilename(
            parent=parent,
            title="Pilih File Excel Nama Desa",
            filetypes=[
                ("Excel Macro-Enabled (*.xlsm)", "*.xlsm"),
                ("Excel Workbook (*.xlsx)", "*.xlsx"),
                ("Semua File Excel", "*.xlsm *.xlsx"),
            ]
        )
        if not path:
            return

        try:
            wb = load_workbook(
                filename=path, read_only=True, data_only=True,
                keep_vba=path.lower().endswith(".xlsm")
            )
        except Exception as exc:
            messagebox.showerror("Import Excel", "File Excel tidak dapat dibaca. Pastikan file valid lalu coba lagi.", parent=parent)
            return

        def norm_sheet(value):
            return re.sub(r"\s+", "", str(value or "").strip().lower())

        target_sheet = next((x for x in wb.sheetnames if norm_sheet(x) in {"namadesa", "nama_desa"}), None)
        if target_sheet is None:
            tersedia = "\n".join(f"• {x}" for x in wb.sheetnames[:30])
            wb.close()
            messagebox.showerror(
                "Import Excel",
                "Sheet nama desa tidak ditemukan.\n\n"
                "Workbook wajib memiliki sheet bernama 'Nama Desa' atau 'NamaDesa'.\n"
                "Berbeda dengan sheet DaftarPegawai, sheet ini TIDAK memakai header: "
                "sel A1 sudah merupakan nama desa pertama, A2 desa kedua, dan seterusnya.\n\n"
                f"Sheet yang ditemukan:\n{tersedia}",
                parent=parent
            )
            return

        ws = wb[target_sheet]
        nama_desa = []
        seen = set()
        for r in range(1, (ws.max_row or 0) + 1):
            value = ws.cell(r, 1).value
            nama = str(value).strip() if value is not None else ""
            if not nama:
                continue
            key = re.sub(r"\s+", " ", nama).strip().casefold()
            if key in seen:
                continue
            seen.add(key)
            nama_desa.append(nama)
        wb.close()

        if not nama_desa:
            messagebox.showwarning("Import Excel", f"Sheet '{target_sheet}' ditemukan, tetapi kolom A tidak berisi nama desa.", parent=parent)
            return

        jawab = messagebox.askyesno(
            "Ganti Seluruh Data Desa",
            f"File: {path}\nSheet: {target_sheet}\nNama desa terbaca: {len(nama_desa)}\n\n"
            "IMPORT INI AKAN MENGGANTIKAN SELURUH DATA DESA MASTER LAMA.\n"
            "Desa yang sedang dipakai pada data penugasan tidak boleh dihapus.\n"
            "Pastikan data sudah di-backup/export sebelum melanjutkan.\n\nLanjutkan?",
            parent=parent
        )
        if not jawab:
            return
        konfirmasi = simpledialog.askstring(
            "Konfirmasi Penggantian Data Desa",
            "Ketik GANTI untuk mengganti master desa:",
            parent=parent
        )
        if konfirmasi != "GANTI":
            return

        conn = get_connection()
        try:
            dipakai = conn.execute("SELECT COUNT(*) FROM jadwal WHERE desa_id IS NOT NULL").fetchone()[0]
            if dipakai:
                messagebox.showwarning(
                    "Import Dibatalkan",
                    f"Masih ada {dipakai} jadwal yang menggunakan desa.\n\n"
                    "Bersihkan/export data penugasan lama terlebih dahulu sebelum mengganti master desa.",
                    parent=parent
                )
                return
            conn.execute("DELETE FROM desa")
            for nama in nama_desa:
                conn.execute("INSERT INTO desa(nama, aktif) VALUES(?,1)", (nama,))
            conn.commit()
        except sqlite3.IntegrityError as exc:
            conn.rollback()
            messagebox.showerror("Import Desa Gagal", "Ada nama desa yang tidak valid atau duplikat. Periksa file lalu coba lagi.", parent=parent)
            return
        except Exception as exc:
            conn.rollback()
            messagebox.showerror("Import Desa Gagal", "Import gagal dan perubahan dibatalkan. Silakan periksa file lalu coba lagi.", parent=parent)
            return
        finally:
            conn.close()

        self.refresh_setelah_perubahan_desa()
        messagebox.showinfo(
            "Import Berhasil",
            f"Berhasil mengimpor {len(nama_desa)} nama desa dari sheet '{target_sheet}'.\n\n"
            "Data diambil langsung dari kolom A mulai A1 (A1, A2, A3, dst.).",
            parent=parent
        )

    def buka_kelola_data(self):
        """Jendela untuk menghapus histori penugasan per bulan tanpa menghapus master pegawai."""
        if getattr(self, "kelola_data_window", None) is not None:
            try:
                if self.kelola_data_window.winfo_exists():
                    self.kelola_data_window.lift()
                    self.kelola_data_window.focus_force()
                    self.refresh_kelola_data()
                    return
            except tk.TclError:
                pass

        win = tk.Toplevel(self)
        self.kelola_data_window = win
        self.modernize_window(win, "Kelola / Bersihkan Data", (760, 500))
        win.title("Kelola / Bersihkan Data Penugasan")
        win.geometry("980x650")
        win.minsize(850, 560)
        self.enable_native_window_controls(win)
        win.transient(self)
        self.enable_native_window_controls(win)
        win.protocol("WM_DELETE_WINDOW", lambda: self.tutup_kelola_data(win))
        self.center_window(win, self)

        outer = tk.Frame(win, bg=UI_BG)
        outer.pack(fill="both", expand=True)
        tk.Label(outer, text="KELOLA / BERSIHKAN DATA PENUGASAN", bg=UI_BG, font=("Segoe UI", 17, "bold")).pack(anchor="w", padx=25, pady=(20, 5))
        tk.Label(
            outer,
            text=(
                "Data penugasan disimpan per bulan. Pilih satu bulan yang sudah tidak diperlukan "
                "untuk menghapus seluruh kegiatan, keanggotaan, jadwal, dan hari liburnya.\n"
                "Master Data Pegawai TIDAK ikut dihapus. Data yang sudah dihapus tidak dapat dipulihkan dari aplikasi."
            ),
            bg=UI_BG, fg=UI_MUTED, justify="left", font=("Segoe UI", 10)
        ).pack(anchor="w", padx=25, pady=(0, 15))

        table_frame = tk.Frame(outer, bg=UI_PANEL, highlightbackground=UI_BORDER, highlightthickness=1)
        table_frame.pack(fill="both", expand=True, padx=25, pady=5)
        self.kelola_data_tree = ttk.Treeview(
            table_frame, columns=("no", "periode_id", "periode", "kegiatan", "anggota", "jadwal"),
            show="headings", selectmode="browse"
        )
        heads = {"no":"No.", "periode_id":"ID", "periode":"Periode", "kegiatan":"Kegiatan", "anggota":"Keanggotaan", "jadwal":"Turun"}
        widths = {"no":58, "periode_id":70, "periode":180, "kegiatan":130, "anggota":150, "jadwal":120}
        for c in heads:
            self.kelola_data_tree.heading(c, text=heads[c])
            self.kelola_data_tree.column(c, width=widths[c], anchor="center")
        self.kelola_data_tree.column("no", width=58, stretch=False)
        self.kelola_data_tree.column("periode_id", width=70, stretch=False)
        self.kelola_data_tree.pack(side="left", fill="both", expand=True)
        sb = ttk.Scrollbar(table_frame, orient="vertical", command=self.kelola_data_tree.yview)
        sb.pack(side="right", fill="y")
        self.kelola_data_tree.configure(yscrollcommand=sb.set)
        self.modernize_tree(self.kelola_data_tree)
        self.kelola_data_tree.bind("<Delete>", lambda event: self.bersihkan_periode_terpilih())

        btn_frame = tk.Frame(outer, bg=UI_BG)
        btn_frame.pack(fill="x", padx=25, pady=(12, 20))
        ttk.Button(btn_frame, text="↻ Refresh", command=self.refresh_kelola_data).pack(side="left")
        ttk.Button(btn_frame, text="🗑 Bersihkan Bulan Terpilih", command=self.bersihkan_periode_terpilih).pack(side="left", padx=8)
        ttk.Button(btn_frame, text="⚠ Bersihkan SEMUA Histori", command=self.bersihkan_semua_histori).pack(side="right")
        self.refresh_kelola_data()

    def tutup_kelola_data(self, win=None):
        try:
            target = win or getattr(self, "kelola_data_window", None)
            if target is not None and target.winfo_exists():
                target.destroy()
        finally:
            self.kelola_data_window = None
            if hasattr(self, "kelola_data_tree"):
                self.kelola_data_tree = None

    def refresh_kelola_data(self):
        tree = getattr(self, "kelola_data_tree", None)
        if tree is None:
            return
        try:
            for item in tree.get_children():
                tree.delete(item)
        except tk.TclError:
            return
        conn = get_connection()
        try:
            rows = conn.execute("SELECT id, bulan, tahun FROM periode ORDER BY tahun DESC, bulan DESC").fetchall()
            for nomor, (periode_id, bulan, tahun) in enumerate(rows, start=1):
                ringkas = get_ringkasan_periode(periode_id)
                tree.insert(
                    "", "end", iid=str(periode_id),
                    values=(nomor, periode_id, f"{NAMA_BULAN[int(bulan)]} {int(tahun)}", ringkas["kegiatan"], ringkas["anggota"], ringkas["jadwal"])
                )
        finally:
            conn.close()

    def bersihkan_periode_terpilih(self):
        tree = getattr(self, "kelola_data_tree", None)
        if tree is None:
            return
        pilihan = tree.selection()
        if not pilihan:
            messagebox.showinfo("Bersihkan Data", "Pilih bulan yang ingin dibersihkan terlebih dahulu.", parent=self.kelola_data_window)
            return
        periode_id = int(pilihan[0])
        ringkas = get_ringkasan_periode(periode_id)
        if not ringkas["periode"]:
            return
        bulan, tahun = int(ringkas["bulan"]), int(ringkas["tahun"])
        jawab = messagebox.askyesno(
            "Konfirmasi Bersihkan Bulan",
            f"Bersihkan seluruh data penugasan {NAMA_BULAN[bulan]} {tahun}?\n\n"
            f"• {ringkas['kegiatan']} kegiatan\n• {ringkas['anggota']} keanggotaan\n• {ringkas['jadwal']} kali turun/jadwal\n\n"
            "Yang dihapus: kegiatan, keanggotaan, jadwal, dan hari libur bulan tersebut.\n"
            "Yang tetap: Data Pegawai.\n\n"
            "Pastikan laporan Excel sudah disimpan jika masih diperlukan. Lanjutkan?",
            parent=self.kelola_data_window
        )
        if not jawab:
            return
        try:
            hapus_data_penugasan_periode(periode_id)
        except Exception as exc:
            messagebox.showerror("Gagal Membersihkan", "Data penugasan tidak dapat dibersihkan. Silakan coba lagi.", parent=self.kelola_data_window)
            return

        if periode_id == self.periode_id:
            self.periode_id = get_or_create_periode(self.bulan, self.tahun)
            self.bagian_aktif = load_bagian_codes(self.periode_id)[0]
            self.refresh_bagian_buttons()
            self.load_bagian_data()
            self.load_kalender()
            self.refresh_selection_table()
            self.refresh_pj_options()
            self.refresh_data_kegiatan()
            self.refresh_data_pegawai()
            self.load_kalender_keseluruhan()
        self.refresh_kelola_data()
        messagebox.showinfo("Berhasil", f"Data penugasan {NAMA_BULAN[bulan]} {tahun} sudah dibersihkan.\nMaster Data Pegawai tetap aman.", parent=self.kelola_data_window)

    def bersihkan_semua_histori(self):
        conn = get_connection()
        try:
            count = conn.execute("SELECT COUNT(*) FROM periode").fetchone()[0]
        finally:
            conn.close()
        if not count:
            messagebox.showinfo("Bersihkan Data", "Belum ada histori penugasan yang tersimpan.", parent=self.kelola_data_window)
            return
        if not messagebox.askyesno(
            "Bersihkan Semua Histori",
            "Semua data penugasan dari SEMUA BULAN akan dihapus.\n\n"
            "Yang dihapus: kegiatan, keanggotaan, jadwal, dan hari libur.\n"
            "Master Data Pegawai TIDAK dihapus.\n\n"
            "Pastikan semua laporan yang diperlukan sudah diekspor. Lanjutkan?",
            parent=self.kelola_data_window
        ):
            return
        konfirmasi = simpledialog.askstring("Konfirmasi Terakhir", "Ketik HAPUS untuk melanjutkan:", parent=self.kelola_data_window)
        if konfirmasi != "HAPUS":
            return
        try:
            hapus_semua_data_penugasan()
        except Exception as exc:
            messagebox.showerror("Gagal Membersihkan", "Data penugasan tidak dapat dibersihkan. Silakan coba lagi.", parent=self.kelola_data_window)
            return
        self.periode_id = get_or_create_periode(self.bulan, self.tahun)
        self.bagian_aktif = load_bagian_codes(self.periode_id)[0]
        self.refresh_bagian_buttons()
        self.load_bagian_data()
        self.load_kalender()
        self.refresh_selection_table()
        self.refresh_pj_options()
        self.refresh_data_kegiatan()
        self.refresh_data_pegawai()
        self.load_kalender_keseluruhan()
        self.refresh_kelola_data()
        messagebox.showinfo("Berhasil", "Semua histori penugasan sudah dibersihkan.\nMaster Data Pegawai tetap aman.", parent=self.kelola_data_window)

    def clear_search_data_pegawai(self):
        self.search_data_pegawai_var.set("")
        self.refresh_data_pegawai()

    def change_period(self, event=None):

        try:

            bulan = NAMA_BULAN.index(
                self.bulan_var.get()
            )

            tahun = int(
                self.tahun_var.get()
            )

        except:

            messagebox.showerror(
                "Periode",
                "Bulan atau tahun tidak valid."
            )

            return

        self.bulan = bulan
        self.tahun = tahun

        self.periode_id = get_or_create_periode(
            bulan,
            tahun
        )

        kode_list = load_bagian_codes(self.periode_id)
        if self.bagian_aktif not in kode_list:
            self.bagian_aktif = kode_list[0]

        self.refresh_bagian_buttons()

        self.header_period.config(
            text=f"{NAMA_BULAN[bulan]} {tahun}"
        )

        self.load_bagian_data()
        self.load_kalender()
        self.refresh_data_pegawai()

    def change_bagian(self, kode):
        self.bagian_aktif = kode
        self.load_bagian_data()
        self.load_kalender()
        self.refresh_data_kegiatan()
        self.update_bagian_button_state()

    def get_deskripsi(self):
        """Mengambil deskripsi dari kotak teks multi-baris."""
        if hasattr(self, "kegiatan_text"):
            return self.kegiatan_text.get("1.0", "end-1c").strip()
        return self.kegiatan_var.get().strip() if hasattr(self, "kegiatan_var") else ""

    def set_deskripsi(self, value):
        """Mengisi deskripsi ke kotak teks dan menjaga state tetap sinkron."""
        value = value or ""
        if hasattr(self, "kegiatan_text"):
            self.kegiatan_text.delete("1.0", "end")
            self.kegiatan_text.insert("1.0", value)
        if hasattr(self, "kegiatan_var"):
            self.kegiatan_var.set(value)

    def load_bagian_data(self):

        data = load_bagian(
            self.periode_id,
            self.bagian_aktif
        )

        if not data:
            return

        (
            bagian_id,
            kode,
            kegiatan,
            lokasi,
            dana,
            penanggung_jawab_id
        ) = data

        self.bagian_id = bagian_id
        conn = get_connection()
        komposisi_row = conn.execute("SELECT COALESCE(komposisi_target, '') FROM bagian WHERE id = ?", (bagian_id,)).fetchone()
        try:
            stored = json.loads(komposisi_row[0]) if komposisi_row and komposisi_row[0] else {}
            if isinstance(stored, dict) and "target" in stored:
                self.komposisi_target = {
                    str(k): int(v) for k, v in (stored.get("target") or {}).items()
                }
                self.komposisi_target_tarif = {
                    str(k): int(v) for k, v in (stored.get("tarif") or {}).items()
                }
            elif isinstance(stored, dict):
                # Kompatibel dengan data lama yang hanya menyimpan angka target.
                self.komposisi_target = {str(k): int(v) for k, v in stored.items()}
                self.komposisi_target_tarif = {}
            else:
                self.komposisi_target = {}
                self.komposisi_target_tarif = {}
        except Exception:
            self.komposisi_target = {}
            self.komposisi_target_tarif = {}

        conn.close()

        conn = get_connection()
        nama_row = conn.execute(
            "SELECT COALESCE(nama_bagian, '') FROM bagian WHERE id = ?",
            (bagian_id,)
        ).fetchone()
        conn.close()

        nama_bagian = nama_row[0] if nama_row else ""
        self.nama_bagian_var.set(
            nama_bagian or f"Kegiatan {kode}"
        )

        deskripsi = kegiatan or ""
        # Sinkronkan state StringVar dan widget Text.
        self.kegiatan_var.set(deskripsi)
        self.set_deskripsi(deskripsi)

        self.dana_var.set(
            str(dana or "")
        )

        self.penanggung_jawab_id = penanggung_jawab_id
        self.refresh_pj_options()
        self.set_pj_value(penanggung_jawab_id)

        anggota = load_anggota(bagian_id)

        self.selected_ids = set(
            pegawai_id
            for pegawai_id, nama in anggota
        )
    

        self.refresh_selection_table()
        self.refresh_data_kegiatan()

        nama_header = self.nama_bagian_var.get().strip() if hasattr(self, "nama_bagian_var") else f"Kegiatan {self.bagian_aktif}"
        self.header_period.config(
            text=f"{NAMA_BULAN[self.bulan]} {self.tahun}"
            f"  •  {nama_header}"
        )

        self.update_bagian_button_state()
        self.update_notebook_title()

    def refresh_pj_options(self):
        if not hasattr(self, "combo_pj"):
            return

        self.pj_options = {}
        values = []
        for p in self.pegawai:
            pegawai_id = p[0]
            nama = p[1] or "Tanpa Nama"
            nip = p[2] or "-"
            display = f"{nama} — {nip}"
            self.pj_options[display] = pegawai_id
            values.append(display)

        self.combo_pj["values"] = values

    def set_pj_value(self, pegawai_id):
        if not hasattr(self, "combo_pj"):
            return
        if not pegawai_id:
            self.pj_var.set("")
            return
        for display, pid in self.pj_options.items():
            if pid == pegawai_id:
                self.pj_var.set(display)
                return
        self.pj_var.set("")

    def get_selected_pj_id(self):
        return self.pj_options.get(self.pj_var.get())



# =========================================================
# RUN
# =========================================================

if __name__ == "__main__":

    configure_windows_dpi()
    app = App()

    app.mainloop()
