import os
import shutil
import threading
import sys
from pathlib import Path
import re
import requests
import zipfile
import hashlib
import subprocess
import json
import customtkinter as ctk
from tkinterdnd2 import TkinterDnD, DND_FILES
from pdf2image import convert_from_path
from tkinter import messagebox, filedialog

# 테마 설정 (앱 시작 전에 설정)
ctk.set_default_color_theme("dark-blue")
ctk.set_appearance_mode("system")

# 현재 버전
CURRENT_VERSION = "1.0.0"
GITHUB_MANIFEST_URL = "https://raw.githubusercontent.com/username/repo/main/manifest.json"  # manifest.json 파일 URL


# ----------- 자동 업데이트 관련 함수 -------------

def calculate_file_hash(file_path):
    """파일의 SHA256 해시값 계산"""
    if not os.path.exists(file_path):
        return None
    
    sha256_hash = hashlib.sha256()
    try:
        with open(file_path, "rb") as f:
            for byte_block in iter(lambda: f.read(4096), b""):
                sha256_hash.update(byte_block)
        return sha256_hash.hexdigest()
    except Exception as e:
        print(f"해시 계산 실패: {file_path} - {e}")
        return None

def check_for_updates():
    """GitHub manifest.json에서 업데이트 확인"""
    try:
        response = requests.get(GITHUB_MANIFEST_URL, timeout=10)
        response.raise_for_status()
        
        manifest = response.json()
        remote_version = manifest.get("version", "0.0.0")
        
        if remote_version > CURRENT_VERSION:
            # 업데이트가 필요한 파일 목록
            files_to_update = []
            
            if getattr(sys, 'frozen', False):
                base_dir = Path(sys.executable).parent
            else:
                base_dir = Path(__file__).parent
            
            for filename, file_info in manifest.get("files", {}).items():
                local_file = base_dir / filename
                remote_hash = file_info.get("sha256")
                
                # 로컬 파일이 없거나 해시가 다르면 업데이트 필요
                if not local_file.exists():
                    files_to_update.append((filename, file_info))
                    print(f"파일 없음: {filename}")
                else:
                    local_hash = calculate_file_hash(local_file)
                    if local_hash != remote_hash:
                        files_to_update.append((filename, file_info))
                        print(f"해시 불일치: {filename}")
            
            return True, remote_version, files_to_update
        
        return False, None, []
        
    except Exception as e:
        print(f"업데이트 확인 실패: {e}")
        return False, None, []

def download_file(url, destination, progress_callback=None):
    """파일 다운로드"""
    try:
        print(f"다운로드 중: {url}")
        response = requests.get(url, stream=True, timeout=30)
        response.raise_for_status()
        
        total_size = int(response.headers.get('content-length', 0))
        downloaded = 0
        
        # 임시 파일로 다운로드
        temp_file = destination.with_suffix(destination.suffix + '.tmp')
        
        with open(temp_file, 'wb') as f:
            for chunk in response.iter_content(chunk_size=8192):
                f.write(chunk)
                downloaded += len(chunk)
                if progress_callback and total_size > 0:
                    progress = (downloaded / total_size) * 100
                    progress_callback(progress)
        
        # 다운로드 완료 후 원본 파일명으로 변경
        if temp_file.exists():
            if destination.exists():
                destination.unlink()
            temp_file.rename(destination)
        
        print(f"다운로드 완료: {destination.name}")
        return True
        
    except Exception as e:
        print(f"다운로드 실패: {e}")
        return False

def apply_updates(files_to_update, base_dir, progress_callback=None):
    """여러 파일 업데이트 적용"""
    try:
        total_files = len(files_to_update)
        
        for idx, (filename, file_info) in enumerate(files_to_update):
            if progress_callback:
                progress_callback(f"다운로드 중: {filename} ({idx+1}/{total_files})")
            
            destination = base_dir / filename
            url = file_info.get("url")
            
            # .exe 파일은 _new 접미사로 임시 저장
            if filename.endswith(".exe"):
                temp_destination = base_dir / (filename.replace(".exe", "_new.exe"))
            else:
                temp_destination = destination
            
            if not download_file(url, temp_destination):
                return False
        
        # 모든 파일 다운로드 완료 후 .exe 파일 교체 처리
        exe_files = [f for f, _ in files_to_update if f.endswith(".exe")]
        
        if exe_files and getattr(sys, 'frozen', False):
            # 배치 파일로 exe 교체
            create_updater_script(base_dir, exe_files)
            return True
        else:
            print("개발 모드이거나 exe 파일이 없습니다.")
            return True
            
    except Exception as e:
        print(f"업데이트 적용 실패: {e}")
        return False

def create_updater_script(base_dir, exe_files):
    """업데이터 배치 스크립트 생성 및 실행"""
    try:
        current_exe = Path(sys.executable)
        updater_script = base_dir / "updater.bat"
        
        # 배치 파일 내용 생성
        bat_lines = ["@echo off", "timeout /t 2 /nobreak > nul"]
        
        for exe_file in exe_files:
            old_exe = base_dir / exe_file
            new_exe = base_dir / exe_file.replace(".exe", "_new.exe")
            backup_exe = base_dir / exe_file.replace(".exe", "_backup.exe")
            
            bat_lines.append(f'if exist "{old_exe}" move "{old_exe}" "{backup_exe}"')
            bat_lines.append(f'move "{new_exe}" "{old_exe}"')
            bat_lines.append(f'if exist "{backup_exe}" del "{backup_exe}"')
        
        bat_lines.append(f'start "" "{current_exe}"')
        bat_lines.append('del "%~f0"')
        
        bat_content = "\n".join(bat_lines)
        
        with open(updater_script, 'w', encoding='cp949') as f:
            f.write(bat_content)
        
        # 배치 파일 실행 후 프로그램 종료
        subprocess.Popen([str(updater_script)], shell=True, creationflags=subprocess.CREATE_NO_WINDOW)
        sys.exit(0)
        
    except Exception as e:
        print(f"업데이터 스크립트 생성 실패: {e}")


# ----------- poppler 설치/경로 준비 -------------

def get_latest_poppler_download_url():
    api_url = "https://api.github.com/repos/oschwartz10612/poppler-windows/releases/latest"
    response = requests.get(api_url)
    response.raise_for_status()
    data = response.json()
    assets = data.get("assets", [])
    for asset in assets:
        if asset["name"].lower().endswith(".zip"):
            return asset["browser_download_url"], asset["name"]
    return None, None

def find_poppler_folder(base_dir: Path):
    for folder in base_dir.iterdir():
        if folder.is_dir() and "poppler" in folder.name.lower():
            bin_path = folder / "Library" / "bin"
            if bin_path.exists():
                return folder.name, str(bin_path)
    return None, None

def download_and_extract_poppler(dest_folder: Path):
    download_url, filename = get_latest_poppler_download_url()
    if not download_url:
        raise RuntimeError("Poppler 윈도우용 최신 zip 파일을 찾을 수 없습니다.")

    zip_path = dest_folder / filename

    print(f"Poppler 다운로드 중: {download_url}")
    with requests.get(download_url, stream=True) as r:
        r.raise_for_status()
        with open(zip_path, 'wb') as f:
            shutil.copyfileobj(r.raw, f)

    release_folder_name = filename.replace(".zip", "")
    release_folder = dest_folder / release_folder_name
    if release_folder.exists():
        shutil.rmtree(release_folder)

    print(f"Poppler 압축 해제 중: {dest_folder}")
    with zipfile.ZipFile(zip_path, 'r') as zip_ref:
        zip_ref.extractall(dest_folder)

    os.remove(zip_path)
    print(f"Poppler 준비 완료: {dest_folder}")

    match = re.search(r'Release-([\d\.]+)', release_folder_name)
    if not match:
        raise RuntimeError("Release 폴더명에서 버전 정보를 추출할 수 없습니다.")
    version_str = match.group(1)

    poppler_bin_path = release_folder / f"poppler-{version_str}" / "Library" / "bin"
    if not poppler_bin_path.exists():
        raise RuntimeError(f"Poppler bin 폴더가 없습니다: {poppler_bin_path}")

    return poppler_bin_path

def prepare_poppler_path():
    base_dir = Path(os.path.abspath(os.path.dirname(sys.executable if getattr(sys, 'frozen', False) else __file__)))

    folder_name, poppler_bin_path = find_poppler_folder(base_dir)
    if poppler_bin_path:
        print(f"기존 poppler 폴더 발견: {folder_name}, 경로: {poppler_bin_path}")
        return poppler_bin_path

    try:
        poppler_bin_path = download_and_extract_poppler(base_dir)
    except Exception as e:
        print(f"Poppler 다운로드/압축 해제 실패: {e}")
        return None

    folder_name, found_bin_path = find_poppler_folder(base_dir)
    if found_bin_path:
        print(f"다운로드 후 poppler 폴더 발견: {folder_name}, 경로: {found_bin_path}")
        return found_bin_path

    print("Poppler 폴더를 찾을 수 없습니다. None 반환")
    return None


# ----------- 진행 팝업 (2개 프로그레스바) -------------

class ProgressPopup(ctk.CTkToplevel):
    def __init__(self, parent, total_files, total_pages):
        super().__init__(parent)
        self.title("진행 중")
        self.geometry("450x240")
        self.resizable(False, False)
        self.grab_set()

        self.total_files = total_files
        self.total_pages = total_pages
        self.completed_files = 0
        self.completed_pages = 0

        # 파일 진행 레이블
        self.file_label = ctk.CTkLabel(self, text=f"파일: 0 / {total_files}")
        self.file_label.pack(pady=(10, 5))

        # 파일 프로그레스바
        self.file_progress = ctk.CTkProgressBar(self, width=400)
        self.file_progress.pack(pady=5)
        self.file_progress.set(0)

        # 페이지 진행 레이블
        self.page_label = ctk.CTkLabel(self, text=f"페이지: 0 / {total_pages}")
        self.page_label.pack(pady=(10, 5))

        # 페이지 프로그레스바
        self.page_progress = ctk.CTkProgressBar(self, width=400)
        self.page_progress.pack(pady=5)
        self.page_progress.set(0)

        # 취소 버튼
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
        ratio = completed_files / self.total_files if self.total_files > 0 else 0
        self.file_progress.set(ratio)
        self.update_idletasks()

    def update_page_progress(self, completed_pages):
        self.completed_pages = completed_pages
        self.page_label.configure(text=f"페이지: {completed_pages} / {self.total_pages}")
        ratio = completed_pages / self.total_pages if self.total_pages > 0 else 0
        self.page_progress.set(ratio)
        self.update_idletasks()

    def show_completion(self):
        self.cancel_button.configure(text="확인", state="normal", command=self.destroy)
        messagebox.showinfo("변환 완료", 
                          f"총 파일 {self.completed_files}개, 총 페이지 {self.completed_pages}페이지를 변환 완료하였습니다.")


# ----------- PDF→JPG 변환기 -------------

class PDFtoJPGApp(ctk.CTkFrame):
    def __init__(self, master):
        super().__init__(master)
        self.pack(fill="both", expand=True)
        self.master = master
        self.master.title(f"PDF → JPG 변환기 v{CURRENT_VERSION}")
        self.master.geometry("600x450")

        # 아이콘 설정 (exe로 빌드 시 경로 고려)
        if getattr(sys, 'frozen', False):
            # PyInstaller로 빌드된 실행 파일
            icon_path = Path(sys.executable).parent / "icon.ico"
        else:
            # 일반 스크립트 실행
            icon_path = Path(__file__).parent / "icon.ico"

        if icon_path.exists():
            self.master.iconbitmap(str(icon_path))

        # 업데이트 확인 버튼
        update_frame = ctk.CTkFrame(self)
        update_frame.pack(pady=5)
        
        self.update_button = ctk.CTkButton(update_frame, text="업데이트 확인", 
                                          command=self.check_updates, width=120)
        self.update_button.pack(side="left", padx=5)
        
        self.version_label = ctk.CTkLabel(update_frame, text=f"v{CURRENT_VERSION}")
        self.version_label.pack(side="left", padx=5)

        self.label = ctk.CTkLabel(self, text="PDF 파일을 아래 상자에 드래그 앤 드롭 하세요")
        self.label.pack(pady=10)

        self.drop_area = ctk.CTkTextbox(self, height=100)
        self.drop_area.pack(padx=10, pady=10, fill="x")

        # 버튼 프레임 생성
        button_frame = ctk.CTkFrame(self)
        button_frame.pack(pady=10)

        # 파일 선택 버튼
        self.select_button = ctk.CTkButton(button_frame, text="파일 선택", command=self.select_files, width=120)
        self.select_button.grid(row=0, column=0, padx=5)

        # 목록 비우기 버튼
        self.clear_button = ctk.CTkButton(button_frame, text="목록 비우기", command=self.clear_list, width=120)
        self.clear_button.grid(row=0, column=1, padx=5)

        # 변환하기 버튼
        self.convert_button = ctk.CTkButton(button_frame, text="변환하기", command=self.start_conversion, width=120)
        self.convert_button.grid(row=0, column=2, padx=5)

        self.status_label = ctk.CTkLabel(self, text="", text_color="blue")
        self.status_label.pack()

        self.pdf_files = []
        self._cancel_requested = False
        self.progress_popup = None

        self.install_dir = Path(os.path.abspath(os.path.dirname(sys.executable if getattr(sys, 'frozen', False) else __file__)))
        self.poppler_path = prepare_poppler_path()
        if not self.poppler_path:
            messagebox.showerror("오류", "poppler를 준비하지 못했습니다. 프로그램을 종료합니다.")
            self.master.destroy()
            return

        # drag and drop binding to root window
        master.drop_target_register(DND_FILES)
        master.dnd_bind("<<Drop>>", self.on_drop)

    def check_updates(self):
        """업데이트 확인"""
        self.status_label.configure(text="업데이트 확인 중...")
        self.update_button.configure(state="disabled")
        
        def check_thread():
            has_update, new_version, files_to_update = check_for_updates()
            
            if has_update and files_to_update:
                self.status_label.configure(text=f"새 버전 발견: v{new_version}")
                
                file_list = "\n".join([f"- {f}" for f, _ in files_to_update])
                message = f"새 버전 v{new_version}이 있습니다.\n\n업데이트할 파일:\n{file_list}\n\n지금 업데이트하시겠습니까?"
                
                if messagebox.askyesno("업데이트 가능", message):
                    # 업데이트 다운로드 및 적용
                    self.status_label.configure(text="업데이트 다운로드 중...")
                    
                    def progress_update(msg):
                        self.status_label.configure(text=msg)
                    
                    if apply_updates(files_to_update, self.install_dir, progress_update):
                        self.status_label.configure(text="업데이트 완료! 프로그램을 재시작합니다...")
                        # exe 파일이 있으면 배치 파일이 실행되어 자동 종료됨
                    else:
                        messagebox.showerror("오류", "업데이트 다운로드에 실패했습니다.")
                        self.status_label.configure(text="")
            else:
                self.status_label.configure(text="최신 버전입니다.")
                messagebox.showinfo("업데이트", "현재 최신 버전을 사용 중입니다.")
            
            self.update_button.configure(state="normal")
        
        threading.Thread(target=check_thread, daemon=True).start()

    def select_files(self):
        """파일 선택 다이얼로그"""
        files = filedialog.askopenfilenames(
            title="PDF 파일 선택",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
        )
        
        for f in files:
            if f not in self.pdf_files:
                self.pdf_files.append(f)
        
        self.update_file_list()

    def clear_list(self):
        """목록 비우기"""
        if self.pdf_files:
            if messagebox.askyesno("목록 비우기", "등록된 모든 파일을 목록에서 제거하시겠습니까?"):
                self.pdf_files.clear()
                self.drop_area.delete("0.0", "end")
                self.status_label.configure(text="목록이 비워졌습니다.")
        else:
            messagebox.showinfo("알림", "목록이 이미 비어있습니다.")

    def update_file_list(self):
        """파일 목록 텍스트박스 업데이트"""
        self.drop_area.delete("0.0", "end")
        file_names = "\n".join([Path(f).name for f in self.pdf_files])
        self.drop_area.insert("0.0", file_names)

    def on_drop(self, event):
        files = self.master.tk.splitlist(event.data)
        for f in files:
            if f.lower().endswith(".pdf"):
                if f not in self.pdf_files:
                    self.pdf_files.append(f)
            else:
                messagebox.showwarning("알림", f"PDF 파일만 등록할 수 있습니다 : {f}")
        
        self.update_file_list()

    def start_conversion(self):
        if not self.pdf_files:
            messagebox.showwarning("경고", "등록된 PDF 파일이 없습니다.")
            return

        # 파일별 페이지 수 계산
        from pdf2image.pdf2image import pdfinfo_from_path
        total_files = len(self.pdf_files)
        total_pages = sum(pdfinfo_from_path(pdf, poppler_path=self.poppler_path)["Pages"] for pdf in self.pdf_files)

        self.progress_popup = ProgressPopup(self.master, total_files, total_pages)
        self.progress_popup.cancel_callback = self.cancel_conversion
        self._cancel_requested = False
        threading.Thread(target=self.convert_files, daemon=True).start()

    def cancel_conversion(self):
        self._cancel_requested = True

    def convert_files(self):
        try:
            from pdf2image.pdf2image import pdfinfo_from_path
            
            completed_files = 0
            completed_pages = 0

            for pdf_file in self.pdf_files:
                if self._cancel_requested:
                    break

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
            messagebox.showerror("오류", f"변환 중 오류 발생 : {e}")
            if self.progress_popup:
                self.progress_popup.destroy()


if __name__ == "__main__":
    root = TkinterDnD.Tk()
    app = PDFtoJPGApp(root)
    root.mainloop()
