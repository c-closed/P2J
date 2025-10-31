import os
import shutil
import threading
import sys
from pathlib import Path
import re
import requests
import zipfile
import queue
import ctypes
import ctypes.wintypes
import hashlib
import json
import customtkinter as ctk
from tkinterdnd2 import TkinterDnD, DND_FILES
from pdf2image import convert_from_path
from tkinter import messagebox, filedialog
import time


# 테마 설정 (앱 시작 전에 설정)
ctk.set_default_color_theme("blue")
ctk.set_appearance_mode("system")



# ----------- 아이콘 경로 가져오기 함수 -------------


def get_icon_path():
    """아이콘 경로 반환"""
    if getattr(sys, 'frozen', False):
        icon_path = Path(sys.executable).parent / "icon.ico"
    else:
        icon_path = Path(__file__).parent / "icon.ico"
    
    if icon_path.exists():
        return str(icon_path)
    return None


def set_window_icon(window, icon_path):
    """윈도우 핸들을 통해 아이콘 설정 (Win32 API 사용)"""
    try:
        hwnd = window.winfo_id()
        print(f"윈도우 핸들: {hwnd}")
        
        hicon = ctypes.windll.user32.LoadImageW(
            0, icon_path, 1, 0, 0, 0x00000010
        )
        if hicon == 0:
            print(f"아이콘 로드 실패: {icon_path}")
            return False
        
        print(f"아이콘 로드 성공: {hicon}")
        
        WM_SETICON = 0x80
        ICON_SMALL = 0
        ICON_BIG = 1
        ctypes.windll.user32.SendMessageW(hwnd, WM_SETICON, ICON_SMALL, hicon)
        ctypes.windll.user32.SendMessageW(hwnd, WM_SETICON, ICON_BIG, hicon)
        print(f"윈도우 핸들 아이콘 설정 완료: {icon_path}")
        return True
    except Exception as e:
        print(f"윈도우 핸들 아이콘 설정 실패: {e}")
        return False



# ----------- 통합 설치/업데이트 팝업 -------------


class UnifiedInstallPopup(ctk.CTkToplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("초기화 중...")
        self.geometry("600x300")
        self.resizable(False, False)
        
        self.update_idletasks()
        x = (self.winfo_screenwidth() // 2) - (600 // 2)
        y = (self.winfo_screenheight() // 2) - (300 // 2)
        self.geometry(f"600x300+{x}+{y}")
        
        self.grab_set()
        
        icon_path = get_icon_path()
        if icon_path:
            self.after(50, lambda: set_window_icon(self, icon_path))
        
        self._last_update_time = 0
        self._update_interval = 0.05
        
        # 로그 큐
        self.log_queue = queue.Queue()
        self._closing = False
        
        # 통합 로그 텍스트박스
        self.log_box = ctk.CTkTextbox(self, height=270, width=550, font=("Consolas", 11))
        self.log_box.pack(pady=10, padx=25)
        self.log_box.configure(state="disabled")
        
        # 로그 저장
        self.logs = []
        
        self.process_log_queue()
    
    def add_log(self, message, is_progress=False):
        """로그 추가 (메인 스레드에서만 호출)"""
        try:
            self.log_box.configure(state="normal")
            
            if is_progress:
                if self.logs:
                    last_line_index = len(self.logs)
                    self.log_box.delete(f"{last_line_index}.0", f"{last_line_index + 1}.0")
                    self.logs.pop()
                
                self.logs.append(message)
                self.log_box.insert("end", message + "\n")
            else:
                self.logs.append(message)
                self.log_box.insert("end", message + "\n")
            
            self.log_box.see("end")
            self.log_box.configure(state="disabled")
            self.update()
        except Exception as e:
            print(f"로그 추가 실패: {e}")
    
    def safe_update_status(self, status, progress=None, detail="", step_name=""):
        """상태 업데이트 (빈도 제한 포함)"""
        current_time = time.time()
        if current_time - self._last_update_time < self._update_interval and progress not in [0, 1.0]:
            return
        self._last_update_time = current_time
        
        try:
            if progress is not None:
                percent = int(progress * 100)
                
                if progress == 0:
                    self.safe_add_log(f"→ {status} {percent}%", is_progress=True)
                elif progress < 1.0:
                    self.safe_add_log(f"→ {status} {percent}%", is_progress=True)
                elif progress == 1.0:
                    self.safe_add_log(f"→ {status} 100%", is_progress=True)
                    if detail:
                        self.safe_add_log(f"  ✓ {detail}", is_progress=False)
                    else:
                        self.safe_add_log(f"  ✓ 완료", is_progress=False)
        except Exception as e:
            pass
    
    def process_log_queue(self):
        """큐에서 로그를 꺼내 처리"""
        if self._closing:
            return
        try:
            while not self.log_queue.empty():
                item = self.log_queue.get_nowait()
                if isinstance(item, tuple) and len(item) == 2:
                    message, is_progress = item
                    self.add_log(message, is_progress)
        except queue.Empty:
            pass
        except Exception as e:
            print(f"큐 처리 실패: {e}")
        try:
            self.after(100, self.process_log_queue)
        except:
            pass
    
    def safe_add_log(self, message, is_progress=False):
        """스레드 안전 로그 추가"""
        try:
            self.log_queue.put((message, is_progress))
        except Exception as e:
            print(f"safe_add_log 실패: {e}")

    def close_window(self):
        """창 닫기"""
        try:
            self._closing = True
            self.grab_release()
            self.destroy()
        except Exception as e:
            print(f"창 닫기 실패: {e}")



# ----------- 자동 업데이트 시스템 -------------


def calculate_file_hash(file_path: Path) -> str:
    """파일의 SHA256 해시 계산"""
    sha256_hash = hashlib.sha256()
    try:
        with open(file_path, "rb") as f:
            for byte_block in iter(lambda: f.read(65536), b""):
                sha256_hash.update(byte_block)
        return sha256_hash.hexdigest()
    except Exception as e:
        print(f"해시 계산 실패 ({file_path}): {e}")
        return None


def get_remote_manifest(repo_owner: str, repo_name: str, branch: str = "main") -> dict:
    """GitHub 리포지토리에서 manifest.json 가져오기"""
    manifest_url = f"https://raw.githubusercontent.com/{repo_owner}/{repo_name}/{branch}/release/manifest.json"
    
    try:
        response = requests.get(manifest_url, timeout=10)
        response.raise_for_status()
        manifest = response.json()
        print(f"원격 매니페스트 로드 완료: {len(manifest.get('files', []))}개 파일")
        return manifest
    except Exception as e:
        print(f"원격 매니페스트 로드 실패: {e}")
        return None


def get_local_manifest(app_dir: Path) -> dict:
    """로컬 manifest.json 가져오기"""
    manifest_file = app_dir / "manifest.json"
    
    if not manifest_file.exists():
        print("로컬 매니페스트 없음")
        return {"version": "0.0.0", "files": []}
    
    try:
        with open(manifest_file, 'r', encoding='utf-8') as f:
            manifest = json.load(f)
        print(f"로컬 매니페스트 로드 완료: v{manifest.get('version', '0.0.0')}")
        return manifest
    except Exception as e:
        print(f"로컬 매니페스트 로드 실패: {e}")
        return {"version": "0.0.0", "files": []}


def save_local_manifest(app_dir: Path, manifest: dict):
    """로컬 manifest.json 저장"""
    manifest_file = app_dir / "manifest.json"
    
    try:
        with open(manifest_file, 'w', encoding='utf-8') as f:
            json.dump(manifest, f, indent=2, ensure_ascii=False)
        print(f"로컬 매니페스트 저장 완료: v{manifest.get('version', '0.0.0')}")
    except Exception as e:
        print(f"로컬 매니페스트 저장 실패: {e}")


def compare_manifests(local_manifest: dict, remote_manifest: dict, app_dir: Path, 
                      progress_callback=None) -> list:
    """매니페스트 비교 및 업데이트 필요 파일 목록 반환"""
    local_files = {f['path']: f['hash'] for f in local_manifest.get('files', [])}
    remote_files = {f['path']: f['hash'] for f in remote_manifest.get('files', [])}
    
    files_to_update = []
    total_files = len(remote_files)
    
    if progress_callback:
        progress_callback("파일 유효성 검사 중", 0, "", "check")
    
    for i, (path, remote_hash) in enumerate(remote_files.items(), start=1):
        local_hash = local_files.get(path)
        file_path = app_dir / path
        
        if not file_path.exists():
            files_to_update.append({
                'path': path,
                'hash': remote_hash,
                'reason': '파일 없음'
            })
            print(f"업데이트 필요: {path} (파일 없음)")
        elif local_hash != remote_hash:
            actual_hash = calculate_file_hash(file_path)
            if actual_hash != remote_hash:
                files_to_update.append({
                    'path': path,
                    'hash': remote_hash,
                    'reason': '해시 불일치'
                })
                print(f"업데이트 필요: {path} (해시 불일치)")
        
        if progress_callback:
            progress = i / total_files
            progress_callback("파일 유효성 검사 중", progress, "", "check")
    
    if progress_callback:
        progress_callback("파일 유효성 검사 중", 1.0, f"{total_files}개 파일 검사 완료", "check")
    
    return files_to_update


def download_file_from_github(repo_owner: str, repo_name: str, file_path: str, 
                               dest_path: Path, branch: str = "main") -> bool:
    """GitHub에서 개별 파일 다운로드"""
    raw_url = f"https://raw.githubusercontent.com/{repo_owner}/{repo_name}/{branch}/release/{file_path}"
    
    try:
        response = requests.get(raw_url, timeout=30)
        response.raise_for_status()
        
        dest_path.parent.mkdir(parents=True, exist_ok=True)
        
        with open(dest_path, 'wb') as f:
            f.write(response.content)
        
        print(f"다운로드 완료: {file_path}")
        return True
    except Exception as e:
        print(f"다운로드 실패 ({file_path}): {e}")
        return False


def update_application(repo_owner: str, repo_name: str, app_dir: Path, 
                       progress_callback=None, branch: str = "main") -> bool:
    """애플리케이션 자동 업데이트"""
    try:
        if progress_callback:
            progress_callback("업데이트 확인 중", 0, "", "check")
        
        remote_manifest = get_remote_manifest(repo_owner, repo_name, branch)
        if not remote_manifest:
            print("원격 매니페스트를 가져올 수 없습니다")
            if progress_callback:
                progress_callback("업데이트 확인 중", 1.0, "원격 매니페스트 로드 실패", "check")
            return False
        
        local_manifest = get_local_manifest(app_dir)
        
        local_version = local_manifest.get('version', '0.0.0')
        remote_version = remote_manifest.get('version', '0.0.0')
        
        print(f"로컬 버전: v{local_version}")
        print(f"원격 버전: v{remote_version}")
        
        if progress_callback:
            progress_callback("업데이트 확인 중", 1.0, f"원격 버전: v{remote_version}", "check")
        
        files_to_update = compare_manifests(local_manifest, remote_manifest, app_dir, progress_callback)
        
        if not files_to_update:
            print("업데이트 필요한 파일 없음")
            return False
        
        print(f"업데이트 필요 파일: {len(files_to_update)}개")
        
        if progress_callback:
            progress_callback("새로운 파일 발견", 0, "", "found")
            time.sleep(0.1)
            progress_callback("새로운 파일 발견", 1.0, f"{len(files_to_update)}개 파일 업데이트 필요", "found")
            time.sleep(0.3)
        
        if progress_callback:
            progress_callback("새로운 파일 다운로드 중", 0, "", "download")
        
        success_count = 0
        total_count = len(files_to_update)
        
        for i, file_info in enumerate(files_to_update):
            file_path = file_info['path']
            dest_path = app_dir / file_path
            
            print(f"업데이트 중 ({i+1}/{total_count}): {file_path}")
            
            if download_file_from_github(repo_owner, repo_name, file_path, dest_path, branch):
                downloaded_hash = calculate_file_hash(dest_path)
                if downloaded_hash == file_info['hash']:
                    success_count += 1
                    print(f"  ✓ 검증 완료")
                else:
                    print(f"  ✗ 해시 불일치")
            
            if progress_callback:
                progress = (i + 1) / total_count
                progress_callback("새로운 파일 다운로드 중", progress, "", "download")
        
        if progress_callback:
            progress_callback("새로운 파일 다운로드 중", 1.0, 
                            f"{success_count}개 파일 다운로드 완료", "download")
        
        if success_count > 0:
            save_local_manifest(app_dir, remote_manifest)
            print(f"업데이트 완료: {success_count}/{total_count} 파일")
            return True
        
        return False
        
    except Exception as e:
        print(f"업데이트 실패: {e}")
        return False


def check_and_update_application(unified_popup, repo_owner: str, repo_name: str, branch: str = "main"):
    """애플리케이션 업데이트 확인 및 실행"""
    if getattr(sys, 'frozen', False):
        app_dir = Path(sys.executable).parent
    else:
        app_dir = Path(__file__).parent
    
    result = {"updated": False, "error": None}
    
    try:
        unified_popup.safe_add_log("→ 업데이트 확인 중...", is_progress=False)
        time.sleep(0.3)
        
        def progress_callback(status, progress, detail, step_name):
            try:
                unified_popup.safe_update_status(status, progress, detail, step_name)
            except:
                pass
        
        updated = update_application(repo_owner, repo_name, app_dir, progress_callback, branch)
        
        if updated:
            unified_popup.safe_add_log("", is_progress=False)
            unified_popup.safe_add_log("  ✓ 업데이트 완료", is_progress=False)
            time.sleep(0.5)
            unified_popup.safe_add_log("  ✓ 프로그램을 다시 시작해주세요", is_progress=False)
            result["updated"] = True
        else:
            unified_popup.safe_add_log("", is_progress=False)
            unified_popup.safe_add_log("  ✓ 최신 버전입니다", is_progress=False)
        
    except Exception as e:
        result["error"] = str(e)
        print(f"업데이트 오류: {e}")
        try:
            unified_popup.safe_add_log(f"  ✗ 오류 발생: {str(e)}", is_progress=False)
        except:
            print(f"로그 표시 실패: {e}")
    
    return result



# ----------- Poppler 설치 및 경로 준비 -------------


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
                match = re.search(r'(\d+\.\d+(?:\.\d+)?)', folder.name)
                if match:
                    version = "v" + match.group(1)
                else:
                    version = folder.name
                return version, str(bin_path)
    return None, None


def download_and_extract_poppler(dest_folder: Path, progress_callback=None):
    """Poppler 다운로드 및 압축 해제"""
    if progress_callback:
        progress_callback("Poppler 다운로드 정보 확인 중", 0, "", "check")
    
    download_url, filename = get_latest_poppler_download_url()
    if not download_url:
        raise RuntimeError("Poppler 윈도우용 최신 zip 파일을 찾을 수 없습니다.")

    if progress_callback:
        progress_callback("Poppler 다운로드 정보 확인 중", 1.0, "다운로드 URL 확인 완료", "check")

    zip_path = dest_folder / filename

    if progress_callback:
        progress_callback("Poppler 다운로드 중", 0, "", "download")

    with requests.get(download_url, stream=True) as r:
        r.raise_for_status()
        total_size = int(r.headers.get('content-length', 0))
        downloaded = 0
        last_reported = 0
        
        with open(zip_path, 'wb') as f:
            for chunk in r.iter_content(chunk_size=65536):
                f.write(chunk)
                downloaded += len(chunk)
                
                if progress_callback and total_size > 0:
                    percent = (downloaded / total_size) * 100
                    if percent - last_reported >= 1.0 or downloaded == total_size:
                        progress = downloaded / total_size
                        progress_callback("Poppler 다운로드 중", progress, "", "download")
                        last_reported = percent

    if progress_callback:
        progress_callback("Poppler 다운로드 중", 1.0, f"{filename} 다운로드 완료", "download")

    if progress_callback:
        progress_callback("Poppler 압축 해제 중", 0, "", "extract")

    release_folder_name = filename.replace(".zip", "")
    release_folder = dest_folder / release_folder_name
    if release_folder.exists():
        shutil.rmtree(release_folder)

    with zipfile.ZipFile(zip_path, 'r') as zip_ref:
        members = zip_ref.namelist()
        total_files = len(members)
        last_reported = 0
        
        for i, member in enumerate(members, start=1):
            zip_ref.extract(member, dest_folder)
            
            percent = (i / total_files) * 100
            if progress_callback and (percent - last_reported >= 1.0 or i == total_files):
                progress = i / total_files
                progress_callback("Poppler 압축 해제 중", progress, "", "extract")
                last_reported = percent

    os.remove(zip_path)

    if progress_callback:
        progress_callback("Poppler 압축 해제 중", 1.0, f"{total_files}개 파일 압축 해제 완료", "extract")

    def find_bin_folder(base_path: Path) -> Path:
        for root, dirs, files in os.walk(base_path):
            root_path = Path(root)
            if root_path.name == "bin" and root_path.parent.name == "Library":
                if (root_path / "pdftoppm.exe").exists():
                    return root_path
        return None
    
    poppler_bin_path = find_bin_folder(dest_folder)
    
    if poppler_bin_path and poppler_bin_path.exists():
        if progress_callback:
            progress_callback("Poppler 설치 완료", 1.0, "PDF → JPG 변환기를 실행합니다", "complete")
        return poppler_bin_path
    
    raise RuntimeError(f"Poppler bin 폴더를 찾을 수 없습니다.")


def prepare_poppler_path_with_ui(unified_popup):
    """UI와 함께 Poppler 경로 준비"""
    base_dir = Path("C:/PDF TO JPG 변환기")
    
    try:
        base_dir.mkdir(parents=True, exist_ok=True)
    except Exception as e:
        return None, f"설치 폴더를 생성할 수 없습니다: {e}"
    
    result = {"path": None, "error": None}
    
    try:
        unified_popup.safe_add_log("", is_progress=False)
        unified_popup.safe_add_log("→ Poppler 유효성 검사 중...", is_progress=False)
        time.sleep(0.3)
        
        latest_version = None
        try:
            download_url, latest_filename = get_latest_poppler_download_url()
            if latest_filename:
                latest_match = re.search(r'(\d+\.\d+\.\d+)', latest_filename)
                if latest_match:
                    latest_version = latest_match.group(1)
                    print(f"최신 버전: {latest_version}")
        except Exception as e:
            print(f"최신 버전 확인 실패: {e}")
            latest_version = None
        
        version, poppler_bin_path = find_poppler_folder(base_dir)
        
        if poppler_bin_path and version:
            unified_popup.safe_add_log(f"  ✓ 검사 완료", is_progress=False)
            time.sleep(0.1)
            unified_popup.safe_add_log(f"  ✓ 설치된 버전: {version}", is_progress=False)
            time.sleep(0.1)
            
            current_version_num = version.replace("v", "")
            
            needs_update = False
            if latest_version and current_version_num != latest_version:
                needs_update = True
            
            if needs_update:
                unified_popup.safe_add_log(f"  ! 최신 버전: v{latest_version}", is_progress=False)
                time.sleep(0.1)
                unified_popup.safe_add_log(f"  ! 이전 버전이 설치되어 있습니다", is_progress=False)
                time.sleep(0.1)
                unified_popup.safe_add_log(f"  ! 이전 버전을 삭제합니다", is_progress=False)
                time.sleep(0.2)
                
                for folder in base_dir.iterdir():
                    if folder.is_dir() and "poppler" in folder.name.lower():
                        try:
                            shutil.rmtree(folder)
                            unified_popup.safe_add_log(f"  ✓ {folder.name} 삭제 완료", is_progress=False)
                            time.sleep(0.1)
                        except Exception as e:
                            unified_popup.safe_add_log(f"  ✗ 삭제 실패: {e}", is_progress=False)
                
                unified_popup.safe_add_log(f"  ! 최신 버전을 다운로드합니다", is_progress=False)
                time.sleep(0.2)
                
                def progress_callback(status, progress, detail, step_name):
                    try:
                        unified_popup.safe_update_status(status, progress, detail, step_name)
                    except:
                        pass
                
                poppler_path = download_and_extract_poppler(base_dir, progress_callback)
                result["path"] = poppler_path
                time.sleep(0.5)
                
            else:
                if latest_version:
                    unified_popup.safe_add_log(f"  ✓ 최신 버전이 설치되어 있습니다", is_progress=False)
                else:
                    unified_popup.safe_add_log(f"  ✓ Poppler가 설치되어 있습니다", is_progress=False)
                time.sleep(0.1)
                
                result["path"] = poppler_bin_path
                
        else:
            unified_popup.safe_add_log(f"  ✓ 검사 완료", is_progress=False)
            time.sleep(0.1)
            unified_popup.safe_add_log(f"  ! Poppler를 찾을 수 없습니다", is_progress=False)
            time.sleep(0.1)
            unified_popup.safe_add_log(f"  ! Poppler를 다운로드합니다", is_progress=False)
            time.sleep(0.2)
            
            def progress_callback(status, progress, detail, step_name):
                try:
                    unified_popup.safe_update_status(status, progress, detail, step_name)
                except:
                    pass
            
            poppler_path = download_and_extract_poppler(base_dir, progress_callback)
            result["path"] = poppler_path
            time.sleep(0.5)
        
    except Exception as e:
        result["error"] = str(e)
        print(f"설치 오류: {e}")
        try:
            unified_popup.safe_add_log(f"  ✗ 오류 발생: {str(e)}", is_progress=False)
        except:
            print(f"로그 표시 실패: {e}")
    
    return result["path"], result["error"]



# ----------- 진행 팝업 -------------


class ProgressPopup(ctk.CTkToplevel):
    def __init__(self, parent, total_files, total_pages):
        super().__init__(parent)
        self.title("진행 중")
        self.geometry("450x170")
        self.resizable(False, False)
        
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
        
        self.grab_set()
        
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
        ratio = completed_files / self.total_files if self.total_files > 0 else 0
        self.file_progress.set(ratio)
        self.update_idletasks()

    def update_page_progress(self, completed_pages):
        self.completed_pages = completed_pages
        self.page_label.configure(text=f"페이지: {completed_pages} / {self.total_pages}")
        ratio = completed_pages / self.total_pages if self.total_pages > 0 else 0
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
            self._auto_close_id = self.after(1000, lambda: self._update_button_countdown(seconds - 1))
        else:
            self._close_window()

    def show_completion(self):
        self.cancel_button.configure(text="확인 (3초)", state="normal", command=self._close_window)
        self._auto_close_id = self.after(1000, lambda: self._update_button_countdown(2))



# ----------- PDF→JPG 변환기 -------------


class PDFtoJPGApp(ctk.CTkFrame):
    def __init__(self, master):
        super().__init__(master)
        self.pack(fill="both", expand=True)
        self.master = master
        self.master.title(f"PDF → JPG 변환기 [made by. 류호준]")
        self.master.geometry("600x300")

        # 메인 창 숨기기
        self.master.withdraw()

        icon_path = get_icon_path()
        if icon_path:
            self.master.iconbitmap(icon_path)
        
        # ============ 통합 초기화 (업데이트 + Poppler) ============
        unified_popup = UnifiedInstallPopup(self.master)
        
        result = {"update_done": False, "poppler_path": None, "should_close": False, "restart_required": False}
        
        def init_thread():
            try:
                # GitHub 리포지토리 정보
                REPO_OWNER = "c-closed"  # 본인의 GitHub 사용자명으로 변경
                REPO_NAME = "P2J"
                BRANCH = "main"
                
                # 1. 애플리케이션 업데이트 확인
                update_result = check_and_update_application(unified_popup, REPO_OWNER, REPO_NAME, BRANCH)
                result["update_done"] = update_result.get("updated", False)
                
                time.sleep(0.5)
                
                # 2. Poppler 준비
                poppler_path, poppler_error = prepare_poppler_path_with_ui(unified_popup)
                result["poppler_path"] = poppler_path
                
                if result["update_done"]:
                    unified_popup.safe_add_log("", is_progress=False)
                    unified_popup.safe_add_log("  ✓ 업데이트 완료 - 재시작 필요", is_progress=False)
                    result["restart_required"] = True
                else:
                    unified_popup.safe_add_log("", is_progress=False)
                    unified_popup.safe_add_log("  ✓ 3초 후 프로그램이 시작됩니다", is_progress=False)
                    time.sleep(1.0)
                    unified_popup.safe_add_log("  ✓ 2초 후 프로그램이 시작됩니다", is_progress=True)
                    time.sleep(1.0)
                    unified_popup.safe_add_log("  ✓ 1초 후 프로그램이 시작됩니다", is_progress=True)
                    time.sleep(1.0)
                
                result["should_close"] = True
                
            except Exception as e:
                print(f"초기화 오류: {e}")
                result["should_close"] = True
        
        def start_init():
            thread = threading.Thread(target=init_thread, daemon=True)
            thread.start()
        
        unified_popup.after(100, start_init)
        
        def check_close():
            if result["should_close"]:
                try:
                    unified_popup.close_window()
                except:
                    pass
            else:
                try:
                    unified_popup.after(100, check_close)
                except:
                    pass
        
        unified_popup.after(500, check_close)
        
        try:
            self.master.wait_window(unified_popup)
        except:
            pass
        
        # 업데이트 완료 시 재시작 요청
        if result["restart_required"]:
            messagebox.showinfo("업데이트 완료", "애플리케이션이 업데이트되었습니다.\n프로그램을 다시 시작해주세요.")
            self.master.destroy()
            return
        
        # Poppler 경로 확인
        if not result["poppler_path"]:
            messagebox.showerror("오류", "Poppler를 준비하지 못했습니다. 프로그램을 종료합니다.")
            self.master.destroy()
            return
        
        self.poppler_path = result["poppler_path"]
        # =========================================

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

        version_frame = ctk.CTkFrame(control_container)
        version_frame.pack(side="right", padx=5)
        
        version_label = ctk.CTkLabel(version_frame, text="v1.0.0", text_color="black")
        version_label.pack(padx=10, pady=5)

        self.pdf_files = []
        self._cancel_requested = False
        self.progress_popup = None
        
        # 메인 창 다시 표시
        self.master.deiconify()
        
        self.master.update_idletasks()
        x = (self.master.winfo_screenwidth() // 2) - (600 // 2)
        y = (self.master.winfo_screenheight() // 2) - (300 // 2)
        self.master.geometry(f"600x300+{x}+{y}")

        master.drop_target_register(DND_FILES)
        master.dnd_bind("<<Drop>>", self.on_drop)

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
            cursor_position = self.drop_area.index("insert")
            self.drop_area.configure(state="disabled")
            
            line_number = int(cursor_position.split('.')[0]) - 1
            
            if 0 <= line_number < len(self.pdf_files):
                removed_file = Path(self.pdf_files[line_number]).name
                if messagebox.askyesno("파일 제거", f"'{removed_file}'을(를) 목록에서 제거하시겠습니까?"):
                    self.pdf_files.pop(line_number)
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
        file_names = "\n".join([Path(f).name for f in self.pdf_files])
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
