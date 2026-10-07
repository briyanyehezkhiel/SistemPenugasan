import openpyxl
from database import get_connection, create_database

FILE_EXCEL = "data/data.xlsm"


def normalisasi_golongan(gol):
    """
    Mengubah:
    III a -> III
    III b -> III
    III c -> III
    III d -> III
    IV a  -> IV
    IV b  -> IV

    Golongan lain seperti VII tetap VII.
    """

    if not gol:
        return ""

    gol = str(gol).strip().upper()

    if gol.startswith("III"):
        return "III"

    if gol.startswith("IV"):
        return "IV"

    if gol.startswith("II"):
        return "II"

    return gol


def import_pegawai():

    create_database()

    print("Membaca Excel...")
    workbook = openpyxl.load_workbook(
        FILE_EXCEL,
        data_only=True,
        read_only=True
    )

    sheet = workbook["DaftarPegawai"]

    conn = get_connection()
    cursor = conn.cursor()

    # Bersihkan data lama agar import tidak menggandakan pegawai
    cursor.execute("DELETE FROM pegawai")

    jumlah = 0

    for row in sheet.iter_rows(min_row=2, values_only=True):

        no = row[0]
        nama = row[1]
        nip = row[2]
        pangkat = row[3]
        golongan = row[4]
        jabatan = row[5]
        tarif = row[6]

        if not nama:
            continue

        golongan_utama = normalisasi_golongan(golongan)

        # Tarif mengikuti data client di Excel.
        # Kalau kosong, coba ambil dari tabel tarif.
        if not tarif:
            result = cursor.execute(
                "SELECT nominal FROM tarif WHERE golongan = ?",
                (golongan_utama,)
            ).fetchone()

            tarif = result[0] if result else 0

        cursor.execute("""
            INSERT INTO pegawai (
                no,
                nama,
                nip,
                pangkat,
                golongan_asli,
                golongan_utama,
                jabatan,
                tarif
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            no,
            str(nama).strip(),
            str(nip or "").strip(),
            str(pangkat or "").strip(),
            str(golongan or "").strip(),
            golongan_utama,
            str(jabatan or "").strip(),
            int(tarif)
        ))

        jumlah += 1

    conn.commit()
    conn.close()
    workbook.close()

    print(f"Import selesai. {jumlah} pegawai berhasil dimasukkan.")


if __name__ == "__main__":
    import_pegawai()