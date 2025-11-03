import shutil
import sys
import json
import threading
from pathlib import Path
from typing import List, Optional, Callable

import customtkinter as ctk
from tkinterdnd2 import TkinterDnD, DND_FILES
from pdf2image import convert_from_path, pdfinfo_from_path
from tkinter import messagebox, filedialog
import ctypes


# ==================== 설정 ====================

class Config:
    """애플리케이션 설정"""
    THEME_COLOR = "blue"
    APP_TITLE = "PDF → JPG 변환기 [made by. 류호준]"
    APP_SIZE = "600x300"
    ICON_FILENAME = "icon.ico"
    POPPLER_FOLDER_NAME = "poppler"
    MANIFEST_FILENAME = "manifest.json"
    
    # 변환 설정
    CONVERSION_DPI = 200
    OUTPUT_FORMAT = "jpeg"
    
    @staticmethod
    def init_theme():
        """테마 초기화"""
        ctk.set_default_color_theme(Config.THEME_COLOR)
        ctk.set_appearance_mode("system")


Config.init_theme()


# ==================== 유틸리티 ====================

class PathUtils:
    """경로 관련 유틸리티"""
    
    @staticmethod
    def get_app_directory() -> Path:
        """애플리케이션 실행 경로 반환"""
        if getattr(sys, 'frozen', False):
            return Path(sys.executable).parent
        return Path(__file__).parent
    
    @staticmethod
    def get_icon_path() -> Optional[str]:
        """아이콘 경로 반환"""
        icon_path = PathUtils.get_app_directory() / Config.ICON_FILENAME
        return str(icon_path) if icon_path.exists() else None
    
    @staticmethod
    def get_poppler_path() -> Optional[str]:
        """Poppler bin 경로 반환"""
        poppler_dir = PathUtils.get_app_directory() / Config.POPPLER_FOLDER_NAME
        
        if not poppler_dir.exists():
            return None
        
        for item in poppler_dir.iterdir():
            if item.is_dir() and "poppler" in item.name.lower():
                poppler_bin = item / "Library" / "bin"
                if poppler_bin.exists() and (poppler_bin / "pdftoppm.exe").exists():
                    return str(poppler_bin)
        
        return None


class IconManager:
    """아이콘 관리"""
    
    @staticmethod
    def set_window_icon(window, icon_path: str) -> bool:
        """Win32 API로 윈도우 아이콘 설정"""
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


class VersionManager:
    """버전 관리"""
    
    @staticmethod
    def get_version() -> str:
        """manifest.json에서 버전 읽기"""
        manifest_path = PathUtils.get_app_directory() / Config.MANIFEST_FILENAME
        
        if manifest_path.exists():
            try:
                with manifest_path.open("r", encoding="utf-8") as f:
                    manifest = json.load(f)
                    version = manifest.get("version", "버전 확인 불가")
                    return version if version != "0.0.0" else "버전 확인 불가"
            except Exception:
                pass
        
        return "버전 확인 불가"


# ==================== PDF 처리 ====================

class PDFProcessor:
    """PDF 파일 처리"""
    
    def __init__(self, poppler_path: str):
        self.poppler_path = poppler_path
    
    def get_page_count(self, pdf_path: str) -> int:
        """PDF 페이지 수 가져오기"""
        try:
            info = pdfinfo_from_path(pdf_path, poppler_path=self.poppler_path)
            return info["Pages"]
        except Exception as e:
            raise RuntimeError(f"PDF 정보 읽기 실패: {e}")
    
    def get_total_pages(self, pdf_files: List[str]) -> int:
        """전체 페이지 수 계산"""
        try:
            return sum(self.get_page_count(pdf) for pdf in pdf_files)
        except Exception as e:
            raise RuntimeError(f"PDF 정보 읽기 실패: {e}")
    
    def convert_to_images(self, pdf_path: str, output_folder: Path,
                         progress_callback: Optional[Callable[[int], None]] = None) -> int:
        """PDF를 이미지로 변환"""
        info = pdfinfo_from_path(pdf_path, poppler_path=self.poppler_path)
        total_pages = info["Pages"]
        digits = len(str(total_pages))
        
        # 변환
        images = convert_from_path(
            pdf_path,
            dpi=Config.CONVERSION_DPI,
            first_page=1,
            last_page=total_pages,
            fmt=Config.OUTPUT_FORMAT,
            output_folder=str(output_folder),
            paths_only=True,
            poppler_path=self.poppler_path
        )
        
        # 파일명 정리
        for i, img_path in enumerate(images, start=1):
            dest_path = output_folder / f"{str(i).zfill(digits)}.jpg"
            shutil.move(img_path, dest_path)
            
            if progress_callback:
                progress_callback(i)
        
        return len(images)


# ==================== UI 컴포넌트 ====================

class ProgressPopup(ctk.CTkToplevel):
    """PDF 변환 진행 상황 표시 팝업"""
    
    def __init__(self, parent, total_files: int, total_pages: int):
        super().__init__(parent)
        self._setup_window(parent, total_files, total_pages)
        self._setup_icon()
        self._create_widgets()
        
        self.cancelled = False
        self.cancel_callback = None
        self._auto_close_id = None
    
    def _setup_window(self, parent, total_files: int, total_pages: int) -> None:
        """윈도우 설정"""
        self.title("진행 중")
        self.geometry("450x170")
        self.resizable(False, False)
        self.grab_set()
        
        self.total_files = total_files
        self.total_pages = total_pages
        self.completed_files = 0
        self.completed_pages = 0
        
        # 부모 윈도우 중앙에 배치
        self.update_idletasks()
        x = parent.winfo_x() + (parent.winfo_width() - 450) // 2
        y = parent.winfo_y() + (parent.winfo_height() - 170) // 2
        self.geometry(f"450x170+{x}+{y}")
    
    def _setup_icon(self) -> None:
        """아이콘 설정"""
        icon_path = PathUtils.get_icon_path()
        if icon_path:
            self.iconbitmap(icon_path)
    
    def _create_widgets(self) -> None:
        """위젯 생성"""
        # 파일 진행률
        self.file_label = ctk.CTkLabel(self, text=f"파일: 0 / {self.total_files}")
        self.file_label.pack(pady=(10, 5))
        
        self.file_progress = ctk.CTkProgressBar(self, width=400)
        self.file_progress.pack(pady=5)
        self.file_progress.set(0)
        
        # 페이지 진행률
        self.page_label = ctk.CTkLabel(self, text=f"페이지: 0 / {self.total_pages}")
        self.page_label.pack(pady=(10, 5))
        
        self.page_progress = ctk.CTkProgressBar(self, width=400)
        self.page_progress.pack(pady=5)
        self.page_progress.set(0)
        
        # 취소 버튼
        self.cancel_button = ctk.CTkButton(self, text="취소", command=self._on_cancel)
        self.cancel_button.pack(pady=10)
    
    def _on_cancel(self) -> None:
        """취소 버튼 클릭"""
        if messagebox.askyesno("작업 취소", "변환 작업을 정말 취소하시겠습니까?"):
            self.cancelled = True
            if self.cancel_callback:
                self.cancel_callback()
            self.cancel_button.configure(state="disabled")
    
    def update_file_progress(self, completed_files: int) -> None:
        """파일 진행률 업데이트"""
        self.completed_files = completed_files
        self.file_label.configure(text=f"파일: {completed_files} / {self.total_files}")
        ratio = completed_files / self.total_files if self.total_files else 0
        self.file_progress.set(ratio)
        self.update_idletasks()
    
    def update_page_progress(self, completed_pages: int) -> None:
        """페이지 진행률 업데이트"""
        self.completed_pages = completed_pages
        self.page_label.configure(text=f"페이지: {completed_pages} / {self.total_pages}")
        ratio = completed_pages / self.total_pages if self.total_pages else 0
        self.page_progress.set(ratio)
        self.update_idletasks()
    
    def show_completion(self) -> None:
        """완료 표시 및 자동 닫기"""
        self.cancel_button.configure(
            text="확인 (3초)",
            state="normal",
            command=self._close_window
        )
        self._auto_close_id = self.after(1000, lambda: self._update_button_countdown(2))
    
    def _update_button_countdown(self, seconds: int) -> None:
        """버튼 카운트다운"""
        if seconds > 0:
            self.cancel_button.configure(text=f"확인 ({seconds}초)")
            self._auto_close_id = self.after(1000, lambda: self._update_button_countdown(seconds - 1))
        else:
            self._close_window()
    
    def _close_window(self) -> None:
        """윈도우 닫기"""
        if self._auto_close_id:
            self.after_cancel(self._auto_close_id)
            self._auto_close_id = None
        self.destroy()


# ==================== 메인 앱 ====================

class PDFtoJPGApp(ctk.CTkFrame):
    """PDF to JPG 변환기 메인 애플리케이션"""
    
    def __init__(self, master):
        super().__init__(master)
        self.pack(fill="both", expand=True)
        self.master = master
        
        self._setup_window()
        self._setup_icon()
        self._check_poppler()
        self._init_variables()
        self._create_widgets()
        self._setup_drag_drop()
        self._center_window()
    
    def _setup_window(self) -> None:
        """윈도우 설정"""
        self.master.title(Config.APP_TITLE)
        self.master.geometry(Config.APP_SIZE)
    
    def _setup_icon(self) -> None:
        """아이콘 설정"""
        icon_path = PathUtils.get_icon_path()
        if icon_path:
            self.master.iconbitmap(icon_path)
    
    def _check_poppler(self) -> None:
        """Poppler 경로 확인"""
        self.poppler_path = PathUtils.get_poppler_path()
        if not self.poppler_path:
            messagebox.showerror(
                "오류",
                "Poppler를 찾을 수 없습니다.\n런처(launcher.exe)를 통해 프로그램을 실행해주세요."
            )
            self.master.destroy()
            sys.exit()
        
        self.pdf_processor = PDFProcessor(self.poppler_path)
    
    def _init_variables(self) -> None:
        """변수 초기화"""
        self.pdf_files: List[str] = []
        self._cancel_requested = False
        self.progress_popup: Optional[ProgressPopup] = None
    
    def _create_widgets(self) -> None:
        """위젯 생성"""
        # 파일 목록 표시 영역
        self.drop_area = ctk.CTkTextbox(self, height=220)
        self.drop_area.pack(padx=10, pady=10, fill="x")
        self.drop_area.configure(state="disabled")
        
        # 컨트롤 버튼 컨테이너
        control_container = ctk.CTkFrame(self, fg_color="transparent")
        control_container.pack(pady=10, fill="x", padx=10)
        
        # 버튼들
        self.select_button = ctk.CTkButton(
            control_container, text="불러오기", command=self.select_files, width=100
        )
        self.select_button.pack(side="left", padx=5)
        
        self.remove_button = ctk.CTkButton(
            control_container, text="지우기", command=self.remove_selected, width=100
        )
        self.remove_button.pack(side="left", padx=5)
        
        self.clear_button = ctk.CTkButton(
            control_container, text="비우기", command=self.clear_list, width=100
        )
        self.clear_button.pack(side="left", padx=5)
        
        self.convert_button = ctk.CTkButton(
            control_container, text="변환하기", command=self.start_conversion, width=100
        )
        self.convert_button.pack(side="left", padx=5)
        
        # 버전 표시
        version = VersionManager.get_version()
        version_label = ctk.CTkLabel(
            control_container, text=f"v{version}", text_color="gray"
        )
        version_label.pack(side="right", padx=10)
    
    def _setup_drag_drop(self) -> None:
        """드래그 앤 드롭 설정"""
        self.master.drop_target_register(DND_FILES)
        self.master.dnd_bind("<<Drop>>", self.on_drop)
    
    def _center_window(self) -> None:
        """윈도우를 화면 중앙에 배치"""
        self.master.update_idletasks()
        x = (self.master.winfo_screenwidth() - 600) // 2
        y = (self.master.winfo_screenheight() - 300) // 2
        self.master.geometry(f"600x300+{x}+{y}")
    
    # ========== 파일 관리 ==========
    
    def select_files(self) -> None:
        """파일 선택 다이얼로그"""
        files = filedialog.askopenfilenames(
            title="PDF 파일 선택",
            filetypes=[("PDF files", "*.pdf")]
        )
        self._add_files(files)
    
    def on_drop(self, event) -> None:
        """드래그 앤 드롭 이벤트 처리"""
        files = self.master.tk.splitlist(event.data)
        pdf_files = [f for f in files if f.lower().endswith(".pdf")]
        self._add_files(pdf_files)
    
    def _add_files(self, files: List[str]) -> None:
        """파일 목록에 추가"""
        for f in files:
            if f not in self.pdf_files:
                self.pdf_files.append(f)
        self._update_file_list()
    
    def remove_selected(self) -> None:
        """선택한 파일 제거"""
        if not self.pdf_files:
            messagebox.showinfo("알림", "목록에 파일이 없습니다.")
            return
        
        try:
            # 커서 위치 가져오기
            self.drop_area.configure(state="normal")
            cursor_index = self.drop_area.index("insert")
            self.drop_area.configure(state="disabled")
            
            line_num = int(cursor_index.split('.')[0]) - 1
            
            if 0 <= line_num < len(self.pdf_files):
                removed_file = Path(self.pdf_files[line_num]).name
                if messagebox.askyesno("파일 제거", f"'{removed_file}'을(를) 목록에서 제거하시겠습니까?"):
                    self.pdf_files.pop(line_num)
                    self._update_file_list()
            else:
                messagebox.showwarning("경고", "유효한 파일을 선택해주세요.")
        except Exception as e:
            messagebox.showerror("오류", f"파일 제거 중 오류 발생: {e}")
    
    def clear_list(self) -> None:
        """파일 목록 비우기"""
        if self.pdf_files:
            if messagebox.askyesno("목록 비우기", "등록된 모든 파일을 목록에서 제거하시겠습니까?"):
                self.pdf_files.clear()
                self._update_file_list()
        else:
            messagebox.showinfo("알림", "목록이 이미 비어있습니다.")
    
    def _update_file_list(self) -> None:
        """파일 목록 표시 업데이트"""
        self.drop_area.configure(state="normal")
        self.drop_area.delete("0.0", "end")
        file_names = "\n".join(Path(f).name for f in self.pdf_files)
        self.drop_area.insert("0.0", file_names)
        self.drop_area.configure(state="disabled")
    
    # ========== 변환 처리 ==========
    
    def start_conversion(self) -> None:
        """변환 시작"""
        if not self.pdf_files:
            messagebox.showwarning("경고", "등록된 PDF 파일이 없습니다.")
            return
        
        try:
            total_files = len(self.pdf_files)
            total_pages = self.pdf_processor.get_total_pages(self.pdf_files)
        except Exception as e:
            messagebox.showerror("오류", str(e))
            return
        
        # 진행률 팝업 생성
        self.progress_popup = ProgressPopup(self.master, total_files, total_pages)
        self.progress_popup.cancel_callback = self._cancel_conversion
        self._cancel_requested = False
        
        # 백그라운드 스레드에서 변환 수행
        threading.Thread(target=self._convert_files, daemon=True).start()
    
    def _cancel_conversion(self) -> None:
        """변환 취소"""
        self._cancel_requested = True
    
    def _convert_files(self) -> None:
        """파일 변환 (백그라운드 스레드)"""
        try:
            completed_files = 0
            completed_pages = 0
            
            for pdf_file in self.pdf_files:
                if self._cancel_requested:
                    break
                
                # 출력 폴더 생성
                pdf_path = Path(pdf_file)
                output_folder = pdf_path.parent / f"JPG 변환({pdf_path.stem})"
                output_folder.mkdir(exist_ok=True)
                
                # 변환 수행
                def page_callback(page_num):
                    nonlocal completed_pages
                    if not self._cancel_requested:
                        completed_pages += 1
                        self.progress_popup.update_page_progress(completed_pages)
                
                self.pdf_processor.convert_to_images(
                    pdf_file,
                    output_folder,
                    progress_callback=page_callback
                )
                
                if not self._cancel_requested:
                    completed_files += 1
                    self.progress_popup.update_file_progress(completed_files)
            
            # 완료 처리
            if not self._cancel_requested:
                self.progress_popup.show_completion()
            else:
                self.progress_popup.cancel_button.configure(state="disabled")
            
            self._cancel_requested = False
        
        except Exception as e:
            messagebox.showerror("오류", f"변환 중 오류 발생: {e}")
            if self.progress_popup:
                self.progress_popup.destroy()


# ==================== 메인 진입점 ====================

def main():
    """메인 함수"""
    root = TkinterDnD.Tk()
    app = PDFtoJPGApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
