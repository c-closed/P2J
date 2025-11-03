import os
import shutil
import sys
import re
from pathlib import Path

import customtkinter as ctk
from tkinterdnd2 import TkinterDnD, DND_FILES
from pdf2image import convert_from_path, pdfinfo_from_path
from tkinter import messagebox, filedialog
import ctypes


# ---------- 상수 및 설정 ----------

THEME_COLOR = "blue"
APP_TITLE = "PDF → JPG 변환기 [made by. 류호준]"
APP_SIZE = "600x300"
ICON_FILENAME = "icon.ico"

# CustomTkinter 기본 테마 설정
ctk.set_default_color_theme(THEME_COLOR)
ctk.set_appearance_mode("system")


# ---------- 유틸리티 함수 ----------

def get_app_directory():
    """애플리케이션 실행 경로 반환"""
    if getattr(sys, 'frozen', False):
        return Path(sys.executable).parent
    else:
        return Path(__file__).parent


def get_icon_path():
    """아이콘 경로 반환"""
    app_dir = get_app_directory()
    icon_path = app_dir / ICON_FILENAME
    return str(icon_path) if icon_path.exists() else None


def set_window_icon(window, icon_path):
    """윈도우 핸들과 Win32 API로 아이콘 설정"""
    try:
        hwnd = window.winfo_id()
        hicon = ctypes.windll.user32.LoadImageW(0, icon_path, 1, 0, 0, 0x00000010)
        if hicon == 0:
            print(f"아이콘 로드 실패: {icon_path}")
            return False
        WM_SETICON = 0x80
        ctypes.windll.user32.SendMessageW(hwnd, WM_SETICON, 0, hicon)
        ctypes.windll.user32.SendMessageW(hwnd, WM_SETICON, 1, hicon)
        return True
    except Exception as e:
        print(f"아이콘 설정 실패: {e}")
        return False


def get_poppler_path():
    """Poppler bin 경로 반환 (폴더명에서 버전 찾기)"""
    app_dir = get_app_directory()
    poppler_dir = app_dir / "poppler"
    
    if not poppler_dir.exists():
        return None
    
    # poppler-버전 형태의 폴더 찾기
    for item in poppler_dir.iterdir():
        if item.is_dir() and "poppler" in item.name.lower():
            poppler_bin = item / "Library" / "bin"
            if poppler_bin.exists() and (poppler_bin / "pdftoppm.exe").exists():
                return str(poppler_bin)
    
    return None


# ---------- UI 컴포넌트 정의 ----------

class ProgressPopup(ctk.CTkToplevel):
    """PDF 변환 진행 상황 표시 팝업"""

    def __init__(self, parent, total_files, total_pages):
        super().__init__(parent)
        self.title("진행 중")
        self.geometry("450x170")
        self.resizable(False, False)
        self.grab_set()

        self.update_idletasks()
        parent_x = parent.winfo_x()
        parent_y = parent.winfo_y()
        parent_width = parent.winfo_width()
        parent_height = parent.winfo_height()

        popup_width = 450
        popup_height = 170

        x = parent_x + (parent_width - popup_width) // 2
        y = parent_y + (parent_height - popup_height) // 2
        self.geometry(f"{popup_width}x{popup_height}+{x}+{y}")

        icon_path = get_icon_path()
        if icon_path:
            self.iconbitmap(icon_path)

        self.total_files = total_files
        self.total_pages = total_pages
        self.completed_files = 0
        self.completed_pages = 0

        self._auto_close_id = None

        self.file_label = ctk.CTkLabel(self, text=f"파일: 0 / {total_files}")
        self.file_label.pack(pady=(10, 5))

        self.file_progress = ctk.CTkProgressBar(self, width=400)
        self.file_progress.pack(pady=5)
        self.file_progress.set(0)

        self.page_label = ctk.CTkLabel(self, text=f"페이지: 0 / {total_pages}")
        self.page_label.pack(pady=(10, 5))

        self.page_progress = ctk.CTkProgressBar(self, width=400)
        self.page_progress.pack(pady=5)
        self.page_progress.set(0)

        self.cancel_button = ctk.CTkButton(self, text="취소", command=self._on_cancel)
        self.cancel_button.pack(pady=10)

        self.cancelled = False
        self.cancel_callback = None

    def _on_cancel(self):
        if messagebox.askyesno("작업 취소", "변환 작업을 정말 취소하시겠습니까?"):
            self.cancelled = True
            if self.cancel_callback:
                self.cancel_callback()
            self.cancel_button.configure(state="disabled")

    def update_file_progress(self, completed_files):
        self.completed_files = completed_files
        self.file_label.configure(text=f"파일: {completed_files} / {self.total_files}")
        ratio = completed_files / self.total_files if self.total_files else 0
        self.file_progress.set(ratio)
        self.update_idletasks()

    def update_page_progress(self, completed_pages):
        self.completed_pages = completed_pages
        self.page_label.configure(text=f"페이지: {completed_pages} / {self.total_pages}")
        ratio = completed_pages / self.total_pages if self.total_pages else 0
        self.page_progress.set(ratio)
        self.update_idletasks()

    def _close_window(self):
        if self._auto_close_id:
            self.after_cancel(self._auto_close_id)
            self._auto_close_id = None
        self.destroy()

    def _update_button_countdown(self, seconds):
        if seconds > 0:
            self.cancel_button.configure(text=f"확인 ({seconds}초)")
            self._auto_close_id = self.after(1000, lambda: self._update_button_countdown(seconds-1))
        else:
            self._close_window()

    def show_completion(self):
        self.cancel_button.configure(text="확인 (3초)", state="normal", command=self._close_window)
        self._auto_close_id = self.after(1000, lambda: self._update_button_countdown(2))


# ---------- 메인 앱 클래스 ----------

class PDFtoJPGApp(ctk.CTkFrame):
    def __init__(self, master):
        super().__init__(master)
        self.pack(fill="both", expand=True)
        self.master = master

        self.master.title(APP_TITLE)
        self.master.geometry(APP_SIZE)

        icon_path = get_icon_path()
        if icon_path:
            self.master.iconbitmap(icon_path)

        # Poppler 경로 확인
        self.poppler_path = get_poppler_path()
        if not self.poppler_path:
            messagebox.showerror(
                "오류", 
                "Poppler를 찾을 수 없습니다.\n런처(launcher.exe)를 통해 프로그램을 실행해주세요."
            )
            self.master.destroy()
            sys.exit()

        self.pdf_files = []
        self._cancel_requested = False
        self.progress_popup = None

        # UI 생성
        self.create_widgets()
        self.center_window()

        # 드래그 앤 드롭 바인딩
        self.master.drop_target_register(DND_FILES)
        self.master.dnd_bind("<<Drop>>", self.on_drop)

    def create_widgets(self):
        self.drop_area = ctk.CTkTextbox(self, height=220)
        self.drop_area.pack(padx=10, pady=10, fill="x")
        self.drop_area.configure(state="disabled")

        control_container = ctk.CTkFrame(self, fg_color="transparent")
        control_container.pack(pady=10, fill="x", padx=10)

        self.select_button = ctk.CTkButton(control_container, text="불러오기", command=self.select_files, width=100)
        self.select_button.pack(side="left", padx=5)

        self.remove_button = ctk.CTkButton(control_container, text="지우기", command=self.remove_selected, width=100)
        self.remove_button.pack(side="left", padx=5)

        self.clear_button = ctk.CTkButton(control_container, text="비우기", command=self.clear_list, width=100)
        self.clear_button.pack(side="left", padx=5)

        self.convert_button = ctk.CTkButton(control_container, text="변환하기", command=self.start_conversion, width=100)
        self.convert_button.pack(side="left", padx=5)

        # 버전 표시
        version = self.get_app_version()
        version_label = ctk.CTkLabel(control_container, text=f"v{version}", text_color="gray")
        version_label.pack(side="right", padx=5)

    def get_app_version(self):
        """manifest.json에서 버전 읽기"""
        app_dir = get_app_directory()
        manifest_path = app_dir / "manifest.json"
        
        if manifest_path.exists():
            try:
                import json
                with manifest_path.open("r", encoding="utf-8") as f:
                    manifest = json.load(f)
                    version = manifest.get("version", "버전 확인 불가")
                    return version if version != "0.0.0" else "버전 확인 불가"
            except Exception:
                pass
        
        return "버전 확인 불가"

    def center_window(self):
        self.master.update_idletasks()
        x = (self.master.winfo_screenwidth() // 2) - (600 // 2)
        y = (self.master.winfo_screenheight() // 2) - (300 // 2)
        self.master.geometry(f"600x300+{x}+{y}")

    def select_files(self):
        files = filedialog.askopenfilenames(title="PDF 파일 선택", filetypes=[("PDF files", "*.pdf")])
        for f in files:
            if f not in self.pdf_files:
                self.pdf_files.append(f)
        self.update_file_list()

    def remove_selected(self):
        if not self.pdf_files:
            messagebox.showinfo("알림", "목록에 파일이 없습니다.")
            return
        try:
            self.drop_area.configure(state="normal")
            cursor_index = self.drop_area.index("insert")
            self.drop_area.configure(state="disabled")

            line_num = int(cursor_index.split('.')[0]) - 1

            if 0 <= line_num < len(self.pdf_files):
                removed_file = Path(self.pdf_files[line_num]).name
                if messagebox.askyesno("파일 제거", f"'{removed_file}'을(를) 목록에서 제거하시겠습니까?"):
                    self.pdf_files.pop(line_num)
                    self.update_file_list()
            else:
                messagebox.showwarning("경고", "유효한 파일을 선택해주세요.")
        except Exception as e:
            messagebox.showerror("오류", f"파일 제거 중 오류 발생: {e}")

    def clear_list(self):
        if self.pdf_files:
            if messagebox.askyesno("목록 비우기", "등록된 모든 파일을 목록에서 제거하시겠습니까?"):
                self.pdf_files.clear()
                self.update_file_list()
        else:
            messagebox.showinfo("알림", "목록이 이미 비어있습니다.")

    def update_file_list(self):
        self.drop_area.configure(state="normal")
        self.drop_area.delete("0.0", "end")
        file_names = "\n".join(Path(f).name for f in self.pdf_files)
        self.drop_area.insert("0.0", file_names)
        self.drop_area.configure(state="disabled")

    def on_drop(self, event):
        files = self.master.tk.splitlist(event.data)
        for f in files:
            if f.lower().endswith(".pdf") and f not in self.pdf_files:
                self.pdf_files.append(f)
        self.update_file_list()

    def start_conversion(self):
        if not self.pdf_files:
            messagebox.showwarning("경고", "등록된 PDF 파일이 없습니다.")
            return

        total_files = len(self.pdf_files)
        total_pages = 0
        try:
            total_pages = sum(pdfinfo_from_path(pdf, poppler_path=self.poppler_path)["Pages"] for pdf in self.pdf_files)
        except Exception as e:
            messagebox.showerror("오류", f"PDF 정보 읽기 실패: {e}")
            return

        self.progress_popup = ProgressPopup(self.master, total_files, total_pages)
        self.progress_popup.cancel_callback = self.cancel_conversion
        self._cancel_requested = False

        import threading
        threading.Thread(target=self.convert_files, daemon=True).start()

    def cancel_conversion(self):
        self._cancel_requested = True

    def convert_files(self):
        try:
            completed_files = 0
            completed_pages = 0

            for pdf_file in self.pdf_files:
                if self._cancel_requested:
                    break

                info = pdfinfo_from_path(pdf_file, poppler_path=self.poppler_path)
                total_pages = info["Pages"]
                pdf_path = Path(pdf_file)

                output_folder = pdf_path.parent / f"JPG 변환({pdf_path.stem})"
                output_folder.mkdir(exist_ok=True)
                digits = len(str(total_pages))

                images = convert_from_path(pdf_file, dpi=200, first_page=1, last_page=total_pages, fmt="jpeg",
                                           output_folder=str(output_folder), paths_only=True,
                                           poppler_path=self.poppler_path)

                for i, img_path in enumerate(images, start=1):
                    if self._cancel_requested:
                        break
                    dest_path = output_folder / f"{str(i).zfill(digits)}.jpg"
                    shutil.move(img_path, dest_path)
                    completed_pages += 1
                    self.progress_popup.update_page_progress(completed_pages)

                if not self._cancel_requested:
                    completed_files += 1
                    self.progress_popup.update_file_progress(completed_files)

            if not self._cancel_requested:
                self.progress_popup.show_completion()
            else:
                self.progress_popup.cancel_button.configure(state="disabled")

            self._cancel_requested = False

        except Exception as e:
            messagebox.showerror("오류", f"변환 중 오류 발생: {e}")
            if self.progress_popup:
                self.progress_popup.destroy()


# ---------- 메인 진입점 ----------

if __name__ == "__main__":
    root = TkinterDnD.Tk()
    app = PDFtoJPGApp(root)
    root.mainloop()
