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
import time

# 테마 설정 (앱 시작 전에 설정)
ctk.set_default_color_theme("dark-blue")
ctk.set_appearance_mode("system")

# 현재 버전
CURRENT_VERSION = "1.0.0"
GITHUB_REPO_OWNER = "c-closed"  # GitHub 사용자명
GITHUB_REPO_NAME = "P2J"  # GitHub 저장소명


# ----------- 시작 업데이트 확인 창 -------------

class StartupUpdateWindow(ctk.CTkToplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("업데이트 확인중...")
        self.geometry("500x280")
        self.resizable(False, False)
        
        # 창을 화면 중앙에 배치
        self.update_idletasks()
        x = (self.winfo_screenwidth() // 2) - (500 // 2)
        y = (self.winfo_screenheight() // 2) - (280 // 2)
        self.geometry(f"500x280+{x}+{y}")
        
        self.grab_set()
        
        self.update_available = False
        self.files_to_update = []
        self.new_version = None
        self._check_running = False

        # 상태 레이블
        self.status_label = ctk.CTkLabel(self, text="업데이트 확인 중...", 
                                        font=("Arial", 14))
        self.status_label.pack(pady=10)

        # 세부 정보 레이블 (파일 개수)
        self.detail_label = ctk.CTkLabel(self, text="", text_color="gray")
        self.detail_label.pack(pady=5)

        # 프로그레스바
        self.progress = ctk.CTkProgressBar(self, width=400)
        self.progress.pack(pady=20)
        self.progress.set(0)

        # 버튼 프레임
        self.button_frame = ctk.CTkFrame(self)
        self.button_frame.pack(pady=20)

        # 시작 시에는 버튼 숨김
        self.skip_button = None
        self.update_button = None
        self.file_label = None

        # 메인 루프 시작 후 업데이트 확인 시작 (1000ms 후)
        self.after(1000, self.start_check_updates)

    def start_check_updates(self):
        """업데이트 확인 시작 (메인 루프 시작 후)"""
        if not self._check_running:
            self._check_running = True
            threading.Thread(target=self.check_updates, daemon=True).start()

    def update_check_progress(self, current, total, filename):
        """파일 체크 진행 상황 업데이트"""
        try:
            self.detail_label.configure(text=f"확인 중: {filename} ({current}/{total})")
            if total > 0:
                self.progress.set(current / total)
        except Exception as e:
            print(f"진행 상황 업데이트 실패: {e}")

    def check_updates(self):
        """업데이트 확인 (백그라운드 스레드)"""
        try:
            has_update, new_version, files_to_update = check_for_updates(
                progress_callback=lambda c, t, f: self.after(0, self.update_check_progress, c, t, f)
            )
            
            # UI 업데이트를 안전하게 예약
            try:
                if has_update and files_to_update:
                    self.after(0, self.show_update_available, new_version, files_to_update)
                else:
                    self.after(0, self.show_no_update)
            except:
                # after 호출 실패 시 직접 호출 시도
                if has_update and files_to_update:
                    self.show_update_available(new_version, files_to_update)
                else:
                    self.show_no_update()
                
        except Exception as e:
            print(f"업데이트 확인 실패: {e}")
            try:
                self.after(0, self.show_check_failed)
            except:
                self.show_check_failed()

    def show_update_available(self, new_version, files_to_update):
        """업데이트가 있을 때 UI 업데이트 (메인 스레드)"""
        try:
            self.update_available = True
            self.files_to_update = files_to_update
            self.new_version = new_version
            
            file_list = ", ".join([f for f, _ in files_to_update[:3]])
            if len(files_to_update) > 3:
                file_list += f" 외 {len(files_to_update) - 3}개"
            
            self.status_label.configure(text=f"새 버전 v{new_version} 발견!")
            self.detail_label.configure(text=f"업데이트 파일: {file_list}")
            self.progress.set(1)
            
            # 버튼 표시
            self.update_button = ctk.CTkButton(self.button_frame, text="업데이트", 
                                              command=self.start_update, width=150)
            self.update_button.pack(side="left", padx=10)
            
            self.skip_button = ctk.CTkButton(self.button_frame, text="건너뛰기", 
                                            command=self.skip_update, width=150,
                                            fg_color="gray")
            self.skip_button.pack(side="left", padx=10)
        except Exception as e:
            print(f"UI 업데이트 실패: {e}")

    def show_no_update(self):
        """업데이트가 없을 때 UI 업데이트 (메인 스레드)"""
        try:
            self.progress.set(1)
            
            self.status_label.configure(text="최신 버전입니다")
            self.detail_label.configure(text="프로그램을 시작합니다...")
            
            # 직접 호출 대신 타이머 사용
            try:
                self.after(1000, self.skip_update)
            except:
                time.sleep(1)
                self.skip_update()
        except Exception as e:
            print(f"UI 업데이트 실패: {e}")

    def show_check_failed(self):
        """업데이트 확인 실패 시 UI 업데이트 (메인 스레드)"""
        try:
            self.progress.set(0)
            
            self.status_label.configure(text="업데이트 확인 실패")
            self.detail_label.configure(text="프로그램을 시작합니다...")
            
            try:
                self.after(1000, self.skip_update)
            except:
                time.sleep(1)
                self.skip_update()
        except Exception as e:
            print(f"UI 업데이트 실패: {e}")

    def start_update(self):
        """업데이트 시작"""
        if self.update_button:
            self.update_button.configure(state="disabled")
        if self.skip_button:
            self.skip_button.configure(state="disabled")
        
        # 업데이트 진행 화면으로 전환
        self.status_label.configure(text="업데이트 다운로드 중...")
        self.detail_label.configure(text="")
        
        # 프로그레스바 다시 표시
        self.progress.set(0)
        
        # 파일별 진행 레이블
        if not self.file_label:
            self.file_label = ctk.CTkLabel(self, text="준비 중...", text_color="gray")
            self.file_label.pack(pady=5)
        else:
            self.file_label.configure(text="준비 중...")
        
        threading.Thread(target=self.download_updates, daemon=True).start()

    def download_updates(self):
        """업데이트 다운로드 (백그라운드 스레드)"""
        try:
            total_files = len(self.files_to_update)
            
            if getattr(sys, 'frozen', False):
                base_dir = Path(sys.executable).parent
            else:
                base_dir = Path(__file__).parent
            
            for idx, (filename, file_info) in enumerate(self.files_to_update):
                # UI 업데이트 (메인 스레드에서)
                try:
                    self.after(0, self.update_download_progress, idx, total_files, filename)
                except:
                    pass
                
                destination = base_dir / filename
                url = file_info.get("url")
                
                # .exe 파일은 _new 접미사로 임시 저장
                if filename.endswith(".exe"):
                    temp_destination = base_dir / (filename.replace(".exe", "_new.exe"))
                else:
                    temp_destination = destination
                
                # 파일 다운로드
                if not download_file(url, temp_destination):
                    try:
                        self.after(0, self.show_download_error, filename)
                    except:
                        self.show_download_error(filename)
                    return
            
            # 모든 파일 다운로드 완료
            try:
                self.after(0, self.complete_download, base_dir)
            except:
                self.complete_download(base_dir)
                
        except Exception as e:
            print(f"업데이트 다운로드 실패: {e}")
            try:
                self.after(0, self.show_download_error, "알 수 없는 파일")
            except:
                self.show_download_error("알 수 없는 파일")

    def update_download_progress(self, idx, total_files, filename):
        """다운로드 진행 상황 업데이트 (메인 스레드)"""
        try:
            self.progress.set((idx) / total_files)
            self.detail_label.configure(text=f"다운로드 중: {filename} ({idx+1}/{total_files})")
        except Exception as e:
            print(f"진행 상황 업데이트 실패: {e}")

    def show_download_error(self, filename):
        """다운로드 오류 표시 (메인 스레드)"""
        try:
            messagebox.showerror("오류", f"{filename} 다운로드에 실패했습니다.")
            self.skip_update()
        except Exception as e:
            print(f"오류 표시 실패: {e}")
            self.skip_update()

    def complete_download(self, base_dir):
        """다운로드 완료 처리 (메인 스레드)"""
        try:
            self.progress.set(1)
            self.status_label.configure(text="업데이트 완료!")
            self.detail_label.configure(text="프로그램을 재시작합니다...")
            
            # .exe 파일 교체 처리
            exe_files = [f for f, _ in self.files_to_update if f.endswith(".exe")]
            
            if exe_files and getattr(sys, 'frozen', False):
                try:
                    self.after(1000, lambda: create_updater_script(base_dir, exe_files))
                except:
                    time.sleep(1)
                    create_updater_script(base_dir, exe_files)
            else:
                try:
                    self.after(1000, self.skip_update)
                except:
                    time.sleep(1)
                    self.skip_update()
        except Exception as e:
            print(f"완료 처리 실패: {e}")

    def skip_update(self):
        """업데이트 건너뛰고 메인 앱 시작"""
        try:
            self.destroy()
        except Exception as e:
            print(f"창 닫기 실패: {e}")


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

def get_latest_release_info():
    """GitHub Releases에서 최신 릴리스 정보 가져오기"""
    try:
        api_url = f"https://api.github.com/repos/{GITHUB_REPO_OWNER}/{GITHUB_REPO_NAME}/releases/latest"
        response = requests.get(api_url, timeout=10)
        response.raise_for_status()
        
        release_data = response.json()
        version = release_data["tag_name"].replace("v", "")
        assets = release_data["assets"]
        
        return version, assets
        
    except Exception as e:
        print(f"릴리스 정보 가져오기 실패: {e}")
        return None, None

def download_manifest_from_release(assets):
    """Releases의 manifest.json 다운로드"""
    try:
        manifest_asset = None
        
        # manifest.json 찾기
        for asset in assets:
            if asset["name"] == "manifest.json":
                manifest_asset = asset
                break
        
        if not manifest_asset:
            print("manifest.json을 찾을 수 없습니다.")
            return None
        
        # manifest.json 다운로드
        response = requests.get(manifest_asset["browser_download_url"], timeout=10)
        response.raise_for_status()
        
        manifest = response.json()
        return manifest
        
    except Exception as e:
        print(f"manifest.json 다운로드 실패: {e}")
        return None

def check_for_updates(progress_callback=None):
    """GitHub Releases에서 업데이트 확인"""
    try:
        # 최신 릴리스 정보 가져오기
        remote_version, assets = get_latest_release_info()
        
        if not remote_version or not assets:
            return False, None, []
        
        print(f"원격 버전: {remote_version}, 현재 버전: {CURRENT_VERSION}")
        
        # 버전 비교
        if remote_version <= CURRENT_VERSION:
            if progress_callback:
                progress_callback(1, 1, "최신 버전")
            return False, None, []
        
        # manifest.json 다운로드
        manifest = download_manifest_from_release(assets)
        
        if not manifest:
            return False, None, []
        
        # 업데이트가 필요한 파일 목록
        files_to_update = []
        
        if getattr(sys, 'frozen', False):
            base_dir = Path(sys.executable).parent
        else:
            base_dir = Path(__file__).parent
        
        manifest_files = manifest.get("files", {})
        total_files = len(manifest_files)
        current_file = 0
        
        for filename, file_info in manifest_files.items():
            current_file += 1
            
            # 진행 상황 콜백
            if progress_callback:
                progress_callback(current_file, total_files, filename)
            
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
                else:
                    print(f"최신 상태: {filename}")
        
        return True, remote_version, files_to_update
        
    except Exception as e:
        print(f"업데이트 확인 실패: {e}")
        return False, None, []

def download_file(url, destination):
    """파일 다운로드"""
    try:
        print(f"다운로드 중: {url}")
        response = requests.get(url, stream=True, timeout=30)
        response.raise_for_status()
        
        # 임시 파일로 다운로드
        temp_file = destination.with_suffix(destination.suffix + '.tmp')
        
        with open(temp_file, 'wb') as f:
            for chunk in response.iter_content(chunk_size=8192):
                f.write(chunk)
        
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


# ----------- 변환 진행 팝업 -------------

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
        self.master.title(f"PDF → JPG 변환기 [made by. 류호준]")
        self.master.geometry("600x250")

        # 아이콘 설정 (exe로 빌드 시 경로 고려)
        if getattr(sys, 'frozen', False):
            icon_path = Path(sys.executable).parent / "icon.ico"
        else:
            icon_path = Path(__file__).parent / "icon.ico"

        if icon_path.exists():
            self.master.iconbitmap(str(icon_path))

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

        # 버전 정보
        version_label = ctk.CTkLabel(self, text=f"v{CURRENT_VERSION}", text_color="gray")
        version_label.pack(pady=5)

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


# ----------- 메인 실행 -------------

if __name__ == "__main__":
    # 임시 루트 창 생성 (숨김)
    root = TkinterDnD.Tk()
    root.withdraw()  # 메인 창 숨기기
    
    # 업데이트 확인 창 표시
    update_window = StartupUpdateWindow(root)
    
    # 업데이트 창이 닫힐 때까지 대기
    root.wait_window(update_window)
    
    # 메인 창 표시
    root.deiconify()
    app = PDFtoJPGApp(root)
    root.mainloop()
