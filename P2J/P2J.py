import sys
import os
import ctypes
import requests
import zipfile
import shutil
import subprocess
import win32api
import platform
import tkinter as tk
from tkinter import messagebox, ttk, Toplevel, Label, Button
from pathlib import Path
from tkinterdnd2 import TkinterDnD, DND_FILES
from pdf2image import convert_from_path



# ----------- 자동 업데이트 부분 -------------


GITHUB_REPO = "c-closed/pdf_to_jpg_converter"
API_SERVER_URL = f"https://api.github.com/repos/{GITHUB_REPO}"



def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False



def run_as_admin():
    ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, " ".join(sys.argv), None, 1)



class UpdateWindow(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("업데이트 진행 중")
        self.geometry("400x150")
        self.resizable(False, False)
        self.progress_label = Label(self, text="업데이트를 시작합니다...")
        self.progress_label.pack(pady=10)
        self.progressbar = ttk.Progressbar(self, length=350, mode="determinate")
        self.progressbar.pack(pady=10)
        self.cancel_button = Button(self, text="취소", command=self.on_cancel)
        self.cancel_button.pack(pady=10)
        self.cancelled = False
        self.auto_confirm_timer = None
        self.confirm_enabled = False


    def on_cancel(self):
        if messagebox.askyesno("취소 확인", "업데이트를 취소하시겠습니까?"):
            self.cancelled = True
            if self.auto_confirm_timer:
                self.after_cancel(self.auto_confirm_timer)


    def set_status(self, message):
        self.progress_label.config(text=message)
        self.update_idletasks()


    def set_progress(self, value, maxval):
        self.progressbar.config(maximum=maxval)
        self.progressbar['value'] = value
        self.update_idletasks()


    def enable_confirm_button_auto(self):
        self.cancel_button.config(text="확인", state="normal")
        self.confirm_enabled = True
        self.auto_confirm_timer = self.after(3000, self.auto_confirm)


    def auto_confirm(self):
        if self.confirm_enabled:
            self.cancel_button.invoke()



def get_file_version(exe_path):
    try:
        info = win32api.GetFileVersionInfo(exe_path, "\\")
        ms = info['FileVersionMS']
        ls = info['FileVersionLS']
        current_version = (ms >> 16, ms & 0xffff, ls >> 16, ls & 0xffff)
        return '.'.join(map(str, current_version))
    except Exception as e:
        print(f"버전 정보를 가져오는 중 오류 발생: {e}")
        return None



def download_with_progress(url, filepath, window: UpdateWindow):
    response = requests.get(url, stream=True)
    total_length = int(response.headers.get('content-length', 0))
    downloaded = 0
    chunk_size = 8192
    with open(filepath, "wb") as f:
        for chunk in response.iter_content(chunk_size=chunk_size):
            if window.cancelled:
                return False
            if chunk:
                f.write(chunk)
                downloaded += len(chunk)
                window.set_progress(downloaded, total_length)
                window.set_status(f"다운로드 중... {downloaded // 1024}KB / {total_length // 1024}KB")
    return True



def do_update(root):
    if not is_admin():
        print("관리자 권한으로 재실행 중...")
        run_as_admin()
        sys.exit(0)


    if platform.architecture()[0] != '64bit':
        install_dir = r"C:\Program Files\PDF_TO_JPG_변환기"
        exe_path = os.path.join(install_dir, "pdf_to_jpg_변환기.exe")
    else:
        install_dir = r"C:\Program Files (x86)\PDF_TO_JPG_변환기"
        exe_path = os.path.join(install_dir, "pdf_to_jpg_변환기.exe")


    os.makedirs(install_dir, exist_ok=True)


    exe_exists = os.path.exists(exe_path)
    if exe_exists:
        ver = get_file_version(exe_path)
        if ver is None:
            messagebox.showerror("오류", "실행파일 버전 정보를 읽을 수 없습니다.")
            sys.exit(1)
        current_version = 'v' + ver
        need_update = False
    else:
        current_version = None
        need_update = True
        print("실행파일이 존재하지 않아 새로 다운로드합니다.")


    print(f"현재 실행파일 버전 → {current_version}")


    try:
        response = requests.get(f"{API_SERVER_URL}/releases/latest")
        if response.status_code != 200:
            messagebox.showerror("오류", "릴리스 체크 실패")
            sys.exit(1)
    except Exception as e:
        messagebox.showerror("오류", f"릴리스 체크 중 오류: {e}")
        sys.exit(1)


    receive = response.json()
    latest_version = receive["tag_name"]


    if exe_exists and latest_version != current_version:
        need_update = True


    if need_update:
        update_type_msg = "다운로드 중..." if not exe_exists else f"새 버전 {latest_version}이 있습니다. 업데이트 하시겠습니까?"
        if exe_exists:
            res = messagebox.askyesno("업데이트", update_type_msg)
            if not res:
                return False
        else:
            print("실행파일이 없으므로 프로그램을 다운로드합니다.")


        update_win = UpdateWindow(root)
        update_win.set_status(update_type_msg)
        root.update()


        download_url = receive["assets"][0]["browser_download_url"]
        update_newfile = os.path.join(install_dir, "pdf_to_jpg_translater_latest.zip")


        success = download_with_progress(download_url, update_newfile, update_win)
        if not success or update_win.cancelled:
            update_win.set_status("업데이트가 취소 또는 실패했습니다.")
            update_win.cancel_button.config(text="확인")
            root.mainloop()
            return False


        update_win.set_status("압축 해제 중...")
        try:
            update_temp_DIR = os.path.join(install_dir, "update_temp_DIR")
            with zipfile.ZipFile(update_newfile, 'r') as zip_ref:
                zip_ref.extractall(update_temp_DIR)
        except Exception as e:
            update_win.set_status(f"압축 해제 실패: {e}")
            update_win.cancel_button.config(text="확인")
            root.mainloop()
            return False


        update_win.set_status("파일 복사 중...")
        shutil.copytree(update_temp_DIR, install_dir, dirs_exist_ok=True)
        os.remove(update_newfile)
        shutil.rmtree(update_temp_DIR)


        update_win.set_status(f"프로그램이 설치되었습니다.")
        update_win.enable_confirm_button_auto()


        def on_close():
            update_win.destroy()
            if os.path.exists(exe_path):
                proc = subprocess.Popen([exe_path])
                proc.wait()


        update_win.protocol("WM_DELETE_WINDOW", on_close)
        update_win.cancel_button.config(command=on_close)
        root.mainloop()
        return True


    else:
        return True



# ----------- PDF->JPG 변환기 부분 -------------


class ProgressPopup:
    # (종전 구현 유지)
    def __init__(self, parent, title="진행 중", cancel_callback=None):
        self.top = Toplevel(parent)
        self.top.title(title)
        self.top.resizable(False, False)
        self.top.grab_set()


        parent.update_idletasks()
        main_x = parent.winfo_x()
        main_y = parent.winfo_y()
        main_w = parent.winfo_width()
        main_h = parent.winfo_height()
        popup_w, popup_h = 400, 140
        popup_x = main_x + (main_w - popup_w) // 2
        popup_y = main_y + (main_h - popup_h) // 2
        self.top.geometry(f"{popup_w}x{popup_h}+{popup_x}+{popup_y}")


        self.label = Label(self.top, text="진행 중입니다... ", pady=10)
        self.label.pack()


        self.progress = ttk.Progressbar(self.top, length=350, mode="determinate")
        self.progress.pack(pady=10)


        self.cancel_button = Button(self.top, text="취소", command=self._on_cancel)
        self.cancel_button.pack(pady=5)


        self.cancel_callback = cancel_callback
        self.cancelled = False


    def _on_cancel(self):
        if messagebox.askyesno("작업 취소", "변환 작업을 정말 취소하시겠습니까?"):
            self.cancelled = True
            if self.cancel_callback:
                self.cancel_callback()
            self.cancel_button.config(state="disabled")
            self.set_status("작업 취소 중입니다...")


    def set_status(self, message):
        self.label.config(text=message)
        self.top.update_idletasks()


    def set_progress(self, value, maximum):
        self.progress.config(maximum=maximum)
        self.progress["value"] = value
        self.top.update_idletasks()


    def enable_confirm(self):
        self.cancel_button.config(text="확인", state="normal", command=self.top.destroy)
        self.set_status("변환이 완료되었습니다.")



class PDFtoJPGApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF → JPG 변환기")
        self.root.geometry("550x200")


        self.label = tk.Label(root, text="PDF 파일을 아래 상자에 드래그 앤 드롭 하세요", pady=10)
        self.label.pack()


        self.drop_area = tk.Listbox(root, width=70, height=6, selectmode=tk.SINGLE)
        self.drop_area.pack(pady=10)


        self.convert_button = tk.Button(root, text="변환하기", command=self.start_conversion)
        self.convert_button.pack(pady=10)


        self.status_label = tk.Label(root, text="", fg="blue")
        self.status_label.pack()


        self.drop_area.drop_target_register(DND_FILES)
        self.drop_area.dnd_bind("<<Drop>>", self.on_drop)


        self.pdf_files = []
        if platform.architecture()[0] != '64bit':
            exe_path = os.path.join(install_dir, "pdf_to_jpg_변환기.exe")
        else:
            exe_path = os.path.join(install_dir, "pdf_to_jpg_변환기.exe")
        self.poppler_path = exe_path + r"poppler-25.07.0\\Library\\bin"
        self.progress_popup = None


    def on_drop(self, event):
        files = self.root.tk.splitlist(event.data)
        for f in files:
            if f.lower().endswith(".pdf"):
                if f not in self.pdf_files:
                    self.pdf_files.append(f)
                    self.drop_area.insert(tk.END, Path(f).name)
            else:
                messagebox.showwarning("알림", f"PDF 파일만 등록할 수 있습니다 : {f}")


    def start_conversion(self):
        if not self.pdf_files:
            messagebox.showwarning("경고", "등록된 PDF 파일이 없습니다.")
            return
        self.progress_popup = ProgressPopup(self.root, cancel_callback=self.cancel_conversion)
        self.root.after(100, self.convert_files)
        self._cancel_requested = False


    def cancel_conversion(self):
        self._cancel_requested = True


    def convert_files(self):
        try:
            total_pages_all_files = 0
            for pdf_file in self.pdf_files:
                total_pages_all_files += self.get_pdf_page_count(pdf_file)


            self.progress_popup.set_progress(0, total_pages_all_files)
            pages_converted = 0
            for pdf_file in self.pdf_files:
                total_pages = self.get_pdf_page_count(pdf_file)
                pdf_path = Path(pdf_file)
                folder_name = f"JPG 변환({pdf_path.stem})"
                output_folder = pdf_path.parent / folder_name
                output_folder.mkdir(exist_ok=True)
                digits = len(str(total_pages))


                for i in range(1, total_pages + 1):
                    if self._cancel_requested:
                        self.progress_popup.set_status("작업이 취소되었습니다.")
                        break


                    pages_converted += 1
                    self.progress_popup.set_status(f"변환 중 : {pages_converted} / {total_pages_all_files} 페이지")
                    self.convert_page(pdf_file, output_folder, i, digits)
                    self.progress_popup.set_progress(pages_converted, total_pages_all_files)


                if self._cancel_requested:
                    break


            if not self._cancel_requested:
                self.progress_popup.enable_confirm()
            else:
                self.progress_popup.cancel_button.config(state="disabled")
            self._cancel_requested = False


        except Exception as e:
            messagebox.showerror("오류", f"변환 중 오류 발생 : {e}")
            if self.progress_popup:
                self.progress_popup.top.destroy()


    def get_pdf_page_count(self, pdf_path):
        from pdf2image.pdf2image import pdfinfo_from_path
        info = pdfinfo_from_path(pdf_path, poppler_path=self.poppler_path)
        return info["Pages"]


    def convert_page(self, pdf_path, output_folder, page_num, digits, dpi=200):
        images = convert_from_path(
            pdf_path,
            dpi=dpi,
            first_page=page_num,
            last_page=page_num,
            fmt="jpeg",
            output_folder=str(output_folder),
            paths_only=True,
            poppler_path=self.poppler_path
        )
        if not images:
            raise RuntimeError(f"페이지 {page_num} 변환 실패")
        src_path = Path(images[0])
        dest_path = output_folder / f"{str(page_num).zfill(digits)}.jpg"
        shutil.move(str(src_path), str(dest_path))


    def update_status(self, message):
        self.status_label.config(text=message)
        self.root.update_idletasks()



if __name__ == "__main__":
    root = TkinterDnD.Tk()
    root.withdraw()  # 숨긴 상태로 시작
    if do_update(root):
        root.deiconify()  # 업데이트 통과 후 메인 window 보이기
        app = PDFtoJPGApp(root)
        root.mainloop()