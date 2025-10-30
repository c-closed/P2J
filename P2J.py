import os
import shutil
import threading
import sys
from pathlib import Path
import requests
import zipfile
import customtkinter as ctk
from tkinterdnd2 import TkinterDnD, DND_FILES
from pdf2image import convert_from_path
from tkinter import messagebox


# ----------- poppler 설치/경로 준비 -------------

def find_poppler_folder(base_dir: Path):
    for folder in base_dir.iterdir():
        if folder.is_dir() and "poppler" in folder.name.lower():
            bin_path = folder / "Library" / "bin"
            if bin_path.exists():
                return folder.name, str(bin_path)
    return None, None

def download_and_extract_poppler(dest_folder: Path):
    api_url = "https://api.github.com/repos/oschwartz10612/poppler-windows/releases/latest"
    response = requests.get(api_url)
    response.raise_for_status()
    data = response.json()
    assets = data.get("assets", [])
    download_url = None
    filename = None
    for asset in assets:
        if asset["name"].endswith(".zip") and "x86_64" in asset["name"]:
            download_url = asset["browser_download_url"]
            filename = asset["name"]
            break
    if not download_url:
        raise RuntimeError("Poppler 윈도우용 최신버전을 찾을 수 없습니다.")

    zip_path = dest_folder / filename

    print(f"Poppler 다운로드 중: {download_url}")
    with requests.get(download_url, stream=True) as r:
        r.raise_for_status()
        with open(zip_path, 'wb') as f:
            shutil.copyfileobj(r.raw, f)

    extract_dir = dest_folder / filename.replace(".zip", "")
    if extract_dir.exists():
        shutil.rmtree(extract_dir)

    print(f"Poppler 압축 해제 중: {extract_dir}")
    with zipfile.ZipFile(zip_path, 'r') as zip_ref:
        zip_ref.extractall(extract_dir)

    os.remove(zip_path)
    print(f"Poppler 준비 완료: {extract_dir}")
    return extract_dir

def prepare_poppler_path():
    base_dir = Path(os.path.abspath(os.path.dirname(sys.executable if getattr(sys, 'frozen', False) else __file__)))

    folder_name, poppler_bin_path = find_poppler_folder(base_dir)
    if poppler_bin_path:
        print(f"기존 poppler 폴더 발견: {folder_name}")
        return poppler_bin_path

    try:
        download_and_extract_poppler(base_dir)
    except Exception as e:
        print(f"Poppler 다운로드/압축 해제 실패: {e}")
        return None

    folder_name, poppler_bin_path = find_poppler_folder(base_dir)
    if poppler_bin_path:
        print(f"다운로드 후 poppler 폴더 발견: {folder_name}")
        return poppler_bin_path

    print("Poppler 폴더를 찾을 수 없습니다.")
    return None


# ----------- 진행 팝업 -------------

class ProgressPopup(ctk.CTkToplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("진행 중")
        self.geometry("400x140")
        self.resizable(False, False)
        self.grab_set()

        self.label = ctk.CTkLabel(self, text="진행 중입니다...")
        self.label.pack(pady=10)

        self.progress = ctk.CTkProgressBar(self, width=350)
        self.progress.pack(pady=10)

        self.cancel_button = ctk.CTkButton(self, text="취소", command=self._on_cancel)
        self.cancel_button.pack(pady=5)

        self.cancelled = False
        self.cancel_callback = None

    def _on_cancel(self):
        if messagebox.askyesno("작업 취소", "변환 작업을 정말 취소하시겠습니까?"):
            self.cancelled = True
            if self.cancel_callback:
                self.cancel_callback()
            self.cancel_button.configure(state="disabled")
            self.set_status("작업 취소 중입니다...")

    def set_status(self, message):
        self.label.configure(text=message)
        self.update_idletasks()

    def set_progress(self, value, maxval):
        ratio = value / maxval if maxval > 0 else 0
        self.progress.set(ratio)
        self.update_idletasks()

    def enable_confirm(self):
        self.cancel_button.configure(text="확인", state="normal", command=self.destroy)
        self.set_status("변환이 완료되었습니다.")


# ----------- PDF→JPG 변환기 -------------

class PDFtoJPGApp(ctk.CTkFrame):
    def __init__(self, master):
        super().__init__(master)
        self.pack(fill="both", expand=True)
        self.master = master
        self.master.title("PDF → JPG 변환기 - 자동업데이트 없음 버전")
        self.master.geometry("600x350")

        self.label = ctk.CTkLabel(self, text="PDF 파일을 아래 상자에 드래그 앤 드롭 하세요")
        self.label.pack(pady=10)

        self.drop_area = ctk.CTkTextbox(self, height=100)
        self.drop_area.pack(padx=10, pady=10, fill="x")

        self.convert_button = ctk.CTkButton(self, text="변환하기", command=self.start_conversion)
        self.convert_button.pack(pady=10)

        self.status_label = ctk.CTkLabel(self, text="", text_color="blue")
        self.status_label.pack()

        self.pdf_files = []

        self.install_dir = Path(os.path.abspath(os.path.dirname(sys.executable if getattr(sys, 'frozen', False) else __file__)))
        self.poppler_path = prepare_poppler_path()
        if not self.poppler_path:
            messagebox.showerror("오류", "poppler를 준비하지 못했습니다. 프로그램을 종료합니다.")
            self.master.destroy()
            return

        self.progress_popup = None

        # drag and drop binding to root window
        master.drop_target_register(DND_FILES)
        master.dnd_bind("<<Drop>>", self.on_drop)

    def on_drop(self, event):
        files = self.master.tk.splitlist(event.data)
        for f in files:
            if f.lower().endswith(".pdf"):
                if f not in self.pdf_files:
                    self.pdf_files.append(f)
                    current_text = self.drop_area.get("0.0", "end").strip()
                    new_text = current_text + ("\n" if current_text else "") + Path(f).name
                    self.drop_area.delete("0.0", "end")
                    self.drop_area.insert("0.0", new_text)
            else:
                messagebox.showwarning("알림", f"PDF 파일만 등록할 수 있습니다 : {f}")

    def start_conversion(self):
        if not self.pdf_files:
            messagebox.showwarning("경고", "등록된 PDF 파일이 없습니다.")
            return
        self.progress_popup = ProgressPopup(self.master)
        self.progress_popup.cancel_callback = self.cancel_conversion
        threading.Thread(target=self.convert_files, daemon=True).start()
        self._cancel_requested = False

    def cancel_conversion(self):
        self._cancel_requested = True

    def convert_files(self):
        try:
            from pdf2image.pdf2image import pdfinfo_from_path
            total_pages_all_files = sum(pdfinfo_from_path(pdf, poppler_path=self.poppler_path)["Pages"] for pdf in self.pdf_files)
            self.progress_popup.set_progress(0, total_pages_all_files)
            pages_converted = 0

            for pdf_file in self.pdf_files:
                info = pdfinfo_from_path(pdf_file, poppler_path=self.poppler_path)
                total_pages = info["Pages"]
                pdf_path = Path(pdf_file)
                folder_name = f"JPG 변환({pdf_path.stem})"
                output_folder = pdf_path.parent / folder_name
                output_folder.mkdir(exist_ok=True)
                digits = len(str(total_pages))
                images = convert_from_path(pdf_file, dpi=200, first_page=1, last_page=total_pages, fmt="jpeg",
                                           output_folder=str(output_folder), paths_only=True, poppler_path=self.poppler_path)

                for i, img_path in enumerate(images, start=1):
                    if self._cancel_requested:
                        self.progress_popup.set_status("작업이 취소되었습니다.")
                        break
                    dest_path = output_folder / f"{str(i).zfill(digits)}.jpg"
                    shutil.move(img_path, dest_path)
                    pages_converted += 1
                    self.progress_popup.set_status(f"변환 중 : {pages_converted} / {total_pages_all_files} 페이지")
                    self.progress_popup.set_progress(pages_converted, total_pages_all_files)

                if self._cancel_requested:
                    break

            if not self._cancel_requested:
                self.progress_popup.enable_confirm()
            else:
                self.progress_popup.cancel_button.configure(state="disabled")
            self._cancel_requested = False
        except Exception as e:
            messagebox.showerror("오류", f"변환 중 오류 발생 : {e}")
            if self.progress_popup:
                self.progress_popup.destroy()


if __name__ == "__main__":
    root = TkinterDnD.Tk()
    app = PDFtoJPGApp(root)
    root.mainloop()
