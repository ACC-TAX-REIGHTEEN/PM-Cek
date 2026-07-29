import os
import shutil
import subprocess
import sys

def main():
    root_dir = os.path.dirname(os.path.abspath(__file__))
    pk_file = os.path.join(root_dir, "PK.xls")
    dapur_dir = os.path.join(root_dir, "Dapur")
    required_dapur_files = ["__init__.py", "1_Cleaner.py", "2_Processing.py"]

    print("--> Memulai pengecekan kelengkapan berkas...")

    if not os.path.exists(pk_file):
        print("--> Error: File 'PK.xls' tidak ditemukan di folder utama.")
        return

    if not os.path.exists(dapur_dir) or not os.path.isdir(dapur_dir):
        print("--> Error: Folder 'Dapur' tidak ditemukan.")
        return

    missing_files = [
        f
        for f in required_dapur_files
        if not os.path.exists(os.path.join(dapur_dir, f))
    ]

    if missing_files:
        print(
            f"--> Error: File wajib berikut tidak ditemukan di dalam folder 'Dapur': {missing_files}"
        )
        return

    print("--> Semua persyaratan lengkap! Memulai proses otomatisasi...")

    try:
        dest_pk = os.path.join(dapur_dir, "PK.xls")
        shutil.copy2(pk_file, dest_pk)
        print(
            "--> [Step 1/5] File 'PK.xls' berhasil disalin ke folder 'Dapur'."
        )

        print("--> [Step 2/5] Menjalankan 1_Cleaner.py di folder Dapur...")
        subprocess.run(
            [sys.executable, "1_Cleaner.py"], cwd=dapur_dir, check=True
        )
        print("--> 1_Cleaner.py selesai dijalankan.")

        print("--> [Step 3/5] Menjalankan 2_Processing.py di folder Dapur...")
        subprocess.run(
            [sys.executable, "2_Processing.py"], cwd=dapur_dir, check=True
        )
        print("--> 2_Processing.py selesai dijalankan.")

        dapur_output = os.path.join(dapur_dir, "PK_Processed.xlsx")
        root_output = os.path.join(root_dir, "PK_Processed.xlsx")

        if os.path.exists(dapur_output):
            if os.path.exists(root_output):
                os.remove(root_output)
            shutil.move(dapur_output, root_output)
            print(
                "--> [Step 4/5] File 'PK_Processed.xlsx' berhasil dipindahkan ke folder utama."
            )
        else:
            print(
                "--> Peringatan: File 'PK_Processed.xlsx' tidak ditemukan di folder Dapur."
            )

    except subprocess.CalledProcessError as e:
        print(
            f"--> Terjadi kesalahan saat mengeksekusi skrip Python di Dapur: {e}"
        )
        return
    except Exception as e:
        print(f"--> Terjadi kesalahan tidak terduga: {e}")
        return

    print(
        "--> [Step 5/5] Membersihkan sisa file Excel dari folder 'Dapur'..."
    )
    cleaned_count = 0
    for filename in os.listdir(dapur_dir):
        if filename.lower().endswith((".xls", ".xlsx")):
            file_to_delete = os.path.join(dapur_dir, filename)
            try:
                os.remove(file_to_delete)
                print(f"--> Menghapus file sementara: {filename}")
                cleaned_count += 1
            except Exception as e:
                print(f"--> Gagal menghapus {filename}: {e}")

    print(
        f"--> Selesai! Seluruh alur kerja berhasil dijalankan ({cleaned_count} file sementara dibersihkan)."
    )

if __name__ == "__main__":
    try:
        main()
    finally:
        input("\nTekan Enter untuk keluar...")