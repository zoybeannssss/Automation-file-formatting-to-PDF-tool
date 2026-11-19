import os
import time
import subprocess
import pyautogui
from PIL import Image

CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
pyautogui.FAILSAFE = True


def salin_ke_clipboard(teks):
    subprocess.run(["clip"], input=teks.encode("utf-16le"), check=True)


def bersihkan_proses_office(target_exes):
    for exe in target_exes:
        subprocess.run(
            ["taskkill", "/F", "/IM", exe],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
    time.sleep(1)


def buka_blokir_protected_view(file_path):
    zone_stream = f"{file_path}:Zone.Identifier"
    try:
        if os.path.exists(zone_stream):
            os.remove(zone_stream)
    except OSError:
        pass

    subprocess.run(
        [
            "powershell",
            "-NoProfile",
            "-Command",
            f"Unblock-File -LiteralPath '{file_path.replace(chr(39), chr(39)*2)}' -ErrorAction SilentlyContinue",
        ],
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
    )


def convert_ppt_via_gui(input_path, output_path):
    if os.path.exists(output_path):
        os.remove(output_path)

    print("   [1/6] Unblock file...")
    buka_blokir_protected_view(input_path)

    print("   [2/6] Membuka PowerPoint...")
    os.startfile(input_path)
    time.sleep(5)

    print("   [3/6] Menekan F12 (Save As)...")
    pyautogui.press("f12")
    time.sleep(1.5)

    print("   [4/6] Menempelkan nama file output...")
    nama_tanpa_ext = os.path.splitext(output_path)[0]
    salin_ke_clipboard(nama_tanpa_ext)
    pyautogui.hotkey("ctrl", "v")
    time.sleep(0.5)

    print("   [5/6] Memilih format PDF...")
    pyautogui.press("tab")
    time.sleep(0.3)
    pyautogui.hotkey("alt", "down")
    time.sleep(0.4)

    # Kembali ke posisi paling atas dulu agar urutan selalu konsisten
    pyautogui.press("home")
    time.sleep(0.2)

    # Turun 3 langkah menuju 'PDF' (1: Macro, 2: 97-2003, 3: PDF)
    for _ in range(3):
        pyautogui.press("down")
        time.sleep(0.15)

    pyautogui.press("enter")
    time.sleep(0.5)
    time.sleep(0.5)

    print("   [6/6] Menyimpan semua slide ke PDF & menutup PowerPoint...")
    pyautogui.press("enter")
    time.sleep(1.2)
    pyautogui.press("enter")

    timeout = 15
    elapsed = 0
    while not os.path.exists(output_path) and elapsed < timeout:
        time.sleep(0.5)
        elapsed += 0.5
    time.sleep(1.5)

    pyautogui.hotkey("alt", "f4")
    time.sleep(1)


def convert_image_to_pdf(input_path, output_path):
    image = Image.open(input_path)
    if image.mode in ("RGBA", "P"):
        image = image.convert("RGB")
    image.save(output_path, "PDF", resolution=100.0)


def proses_tahap_non_com(files):
    print("\n=== TAHAP 1: PPT/PPTX & Gambar (Tanpa pywin32) ===")
    for filename in files:
        name_without_ext, ext = os.path.splitext(filename)
        ext = ext.lower()

        if ext not in (".ppt", ".pptx", ".png", ".jpg", ".jpeg"):
            continue

        input_path = os.path.join(CURRENT_DIR, filename)
        output_path = os.path.join(CURRENT_DIR, f"{name_without_ext}.pdf")
        print(f"\nFile: {filename}")

        try:
            if ext in (".ppt", ".pptx"):
                convert_ppt_via_gui(input_path, output_path)
                if os.path.exists(output_path):
                    print("   -> SUKSES (PPT/PPTX)")
                else:
                    print("   -> GAGAL (File PDF belum terbentuk)")
            elif ext in (".png", ".jpg", ".jpeg"):
                convert_image_to_pdf(input_path, output_path)
                print("   -> SUKSES (Gambar)")
        except Exception as e:
            print(f"   -> GAGAL: {e}")

    bersihkan_proses_office(["POWERPNT.EXE"])


def proses_tahap_com(files):
    doc_files = [
        f for f in files
        if os.path.splitext(f)[1].lower() in (".doc", ".docx", ".xls", ".xlsx")
    ]
    if not doc_files:
        print("\nTidak ada file Word/Excel untuk Tahap 2.")
        return

    print("\n=== TAHAP 2: Word & Excel (pywin32) ===")
    import win32com.client

    word_app = None
    excel_app = None

    try:
        for filename in doc_files:
            name_without_ext, ext = os.path.splitext(filename)
            ext = ext.lower()

            input_path = os.path.join(CURRENT_DIR, filename)
            output_path = os.path.join(CURRENT_DIR, f"{name_without_ext}.pdf")
            print(f"Memproses: {filename} ...", end=" ", flush=True)

            try:
                if ext in (".doc", ".docx"):
                    if not word_app:
                        word_app = win32com.client.Dispatch("Word.Application")
                        word_app.Visible = False
                    doc = word_app.Documents.Open(input_path)
                    doc.ExportAsFixedFormat(
                        OutputFileName=output_path,
                        ExportFormat=17,
                        OptimizeFor=1,
                    )
                    doc.Close(False)
                    print("SUKSES (Word)")

                elif ext in (".xls", ".xlsx"):
                    if not excel_app:
                        excel_app = win32com.client.Dispatch("Excel.Application")
                        excel_app.Visible = False
                    wb = excel_app.Workbooks.Open(input_path)
                    wb.ExportAsFixedFormat(
                        Type=0,
                        Filename=output_path,
                        Quality=1,
                    )
                    wb.Close(False)
                    print("SUKSES (Excel)")
            except Exception as e:
                print(f"GAGAL: {e}")
    finally:
        if word_app:
            word_app.Quit()
        if excel_app:
            excel_app.Quit()


def main():
    print("Folder aktif:", CURRENT_DIR)
    bersihkan_proses_office(["POWERPNT.EXE", "WINWORD.EXE", "EXCEL.EXE"])

    valid_ext = (".doc", ".docx", ".ppt", ".pptx", ".xls", ".xlsx", ".png", ".jpg", ".jpeg")
    script_name = os.path.basename(__file__)
    files = [
        f for f in os.listdir(CURRENT_DIR)
        if not f.startswith("~$")
        and f != script_name
        and os.path.splitext(f)[1].lower() in valid_ext
    ]

    print("File terdeteksi:", files)
    if not files:
        print("Tidak ada file yang didukung di folder ini.")
        return

    proses_tahap_non_com(files)
    proses_tahap_com(files)
    print("\nSelesai semua.")


if __name__ == "__main__":
    main()