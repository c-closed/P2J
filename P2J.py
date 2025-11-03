import os
import shutil
import threading
import sys
import queue
import ctypes
import hashlib
import json
import re
import time
import requests
import zipfile
from pathlib import Path

import customtkinter as ctk
from tkinterdnd2 import TkinterDnD, DND_FILES
from pdf2image import convert_from_path, pdfinfo_from_path
from tkinter import messagebox, filedialog


# ---------- 상수 및 설정 ----------

THEME_COLOR = "blue"
APP_TITLE = "PDF → JPG 변환기 [made by. 류호준]"
APP_SIZE = "600x300"
POPPLER_REPO_OWNER = "oschwartz10612"
POPPLER_REPO_NAME = "poppler-windows"
APP_REPO_OWNER = "c-closed"
APP_REPO_NAME = "P2J"
APP_BRANCH = "main"
ICON_FILENAME = "icon.ico"
POPPLER_FOLDER_NAME = "poppler"  # 애플리케이션 하위에 생성할 폴더명

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
        ctypes.windll.user32.SendMessageW(hwnd, WM_SETICON, 0, hicon)  # ICON_SMALL
        ctypes.windll.user32.SendMessageW(hwnd, WM_SETICON, 1, hicon)  # ICON_BIG
        return True
    except Exception as e:
        print(f"아이콘 설정 실패: {e}")
        return False


def calculate_file_hash(file_path: Path) -> str:
    """파일 SHA256 해시 계산"""
    sha256_hash = hashlib.sha256()
    try:
        with open(file_path, "rb") as f:
            for byte_block in iter(lambda: f.read(65536), b""):
                sha256_hash.update(byte_block)
        return sha256_hash.hexdigest()
    except Exception as e:
        print(f"해시 계산 실패 ({file_path}): {e}")
        return None


# ---------- 업데이트 관련 함수 ----------

def get_remote_manifest(repo_owner, repo_name, branch=APP_BRANCH, log_callback=None) -> dict:
    """원격 manifest.json 읽기"""
    url = f"https://raw.githubusercontent.com/{repo_owner}/{repo_name}/{branch}/release/manifest.json"
    try:
        if log_callback:
            log_callback("  → 원격 manifest 요청 중...", False)
        resp = requests.get(url, timeout=10)
        resp.raise_for_status()
        if log_callback:
            log_callback("  ✓ 원격 manifest 로드 완료", False)
        return resp.json()
    except Exception as e:
        if log_callback:
            log_callback(f"  ✗ 원격 manifest 로드 실패: {e}", False)
        print(f"원격 manifest 불러오기 실패: {e}")
        return None


def get_local_manifest(app_dir: Path) -> dict:
    """로컬 manifest.json 읽기"""
    manifest_path = app_dir / "manifest.json"
    if not manifest_path.exists():
        return {"version": "0.0.0", "files": []}
    try:
        with manifest_path.open("r", encoding="utf-8") as f:
            return json.load(f)
    except Exception as e:
        print(f"로컬 manifest 불러오기 실패: {e}")
        return {"version": "0.0.0", "files": []}


def save_local_manifest(app_dir: Path, manifest: dict):
    """로컬 manifest.json 저장"""
    manifest_path = app_dir / "manifest.json"
    try:
        with manifest_path.open("w", encoding="utf-8") as f:
            json.dump(manifest, f, indent=2, ensure_ascii=False)
    except Exception as e:
        print(f"로컬 manifest 저장 실패: {e}")


def compare_manifests(local_manifest, remote_manifest, app_dir, log_callback=None) -> dict:
    """매니페스트 비교 후 업데이트 및 삭제 대상 파일 목록 만들어 반환"""
    local_files = {f['path']: f['hash'] for f in local_manifest.get('files', [])}
    remote_files = {f['path']: f['hash'] for f in remote_manifest.get('files', [])}
    files_to_update = []
    files_to_delete = []

    total_files = len(remote_files) + len(local_files)
    current_file = 0

    if log_callback:
        log_callback("  → 파일 검사 시작...", False)

    # 업데이트 대상 검사
    for path, remote_hash in remote_files.items():
        current_file += 1
        file_path = app_dir / path
        needs_update = False
        reason = ""

        if not file_path.exists():
            needs_update = True
            reason = '파일 없음'
        else:
            local_hash = local_files.get(path)
            if local_hash is None:
                actual_hash = calculate_file_hash(file_path)
                if actual_hash != remote_hash:
                    needs_update = True
                    reason = '매니페스트 누락'
            elif local_hash != remote_hash:
                needs_update = True
                reason = '해시 불일치'

        if needs_update:
            files_to_update.append({'path': path, 'hash': remote_hash, 'reason': reason})
            if log_callback:
                log_callback(f"    • {path}: {reason}", False)

    # 삭제 대상 검사
    for path in local_files.keys():
        current_file += 1
        if path not in remote_files:
            file_path = app_dir / path
            if file_path.exists():
                files_to_delete.append({'path': path, 'reason': '원격에 없음'})
                if log_callback:
                    log_callback(f"    • {path}: 삭제 대상", False)

    if log_callback:
        log_callback(f"  ✓ 파일 검사 완료: 업데이트 {len(files_to_update)}개, 삭제 {len(files_to_delete)}개", False)

    return {'to_update': files_to_update, 'to_delete': files_to_delete}


def download_file_from_github(repo_owner, repo_name, file_path, dest_path: Path, branch=APP_BRANCH) -> bool:
    """GitHub에서 단일 파일 다운로드"""
    url = f"https://raw.githubusercontent.com/{repo_owner}/{repo_name}/{branch}/release/{file_path}"
    try:
        resp = requests.get(url, timeout=30)
        resp.raise_for_status()
        dest_path.parent.mkdir(parents=True, exist_ok=True)
        with open(dest_path, "wb") as f:
            f.write(resp.content)
        return True
    except Exception as e:
        print(f"다운로드 실패({file_path}): {e}")
        return False


def update_application(repo_owner, repo_name, app_dir: Path, log_callback=None, branch=APP_BRANCH) -> bool:
    """앱 자동 업데이트 수행"""
    try:
        if log_callback:
            log_callback("→ 업데이트 확인 시작", False)

        remote_manifest = get_remote_manifest(repo_owner, repo_name, branch, log_callback)
        if not remote_manifest:
            if log_callback:
                log_callback("  ✗ 업데이트 확인 실패", False)
            return False

        local_manifest = get_local_manifest(app_dir)
        local_version = local_manifest.get('version', '0.0.0')
        remote_version = remote_manifest.get('version', '0.0.0')

        if log_callback:
            log_callback(f"  • 현재 버전: v{local_version}", False)
            log_callback(f"  • 최신 버전: v{remote_version}", False)

        cmp_result = compare_manifests(local_manifest, remote_manifest, app_dir, log_callback)
        files_to_update = cmp_result['to_update']
        files_to_delete = cmp_result['to_delete']

        if not files_to_update and not files_to_delete:
            if log_callback:
                log_callback("  ✓ 변경 사항 없음", False)
            return False

        # 삭제
        if files_to_delete:
            if log_callback:
                log_callback(f"→ 파일 삭제 시작 ({len(files_to_delete)}개)", False)
            deleted_count = 0
            for f in files_to_delete:
                p = app_dir / f['path']
                try:
                    if p.exists():
                        p.unlink()
                        deleted_count += 1
                        if log_callback:
                            log_callback(f"  ✓ 삭제: {f['path']}", False)
                except Exception as e:
                    if log_callback:
                        log_callback(f"  ✗ 삭제 실패({f['path']}): {e}", False)
            if log_callback:
                log_callback(f"  ✓ 파일 삭제 완료: {deleted_count}개", False)

        # 업데이트
        has_changes = False
        if files_to_update:
            if log_callback:
                log_callback(f"→ 파일 다운로드 시작 ({len(files_to_update)}개)", False)
            success_count = 0
            for i, f in enumerate(files_to_update, start=1):
                dest_path = app_dir / f['path']
                if log_callback:
                    log_callback(f"  → [{i}/{len(files_to_update)}] {f['path']} 다운로드 중...", False)
                if download_file_from_github(repo_owner, repo_name, f['path'], dest_path, branch):
                    downloaded_hash = calculate_file_hash(dest_path)
                    if downloaded_hash == f['hash']:
                        success_count += 1
                        if log_callback:
                            log_callback(f"  ✓ [{i}/{len(files_to_update)}] {f['path']} 완료", False)
                    else:
                        if log_callback:
                            log_callback(f"  ✗ [{i}/{len(files_to_update)}] {f['path']} 해시 불일치", False)
                else:
                    if log_callback:
                        log_callback(f"  ✗ [{i}/{len(files_to_update)}] {f['path']} 다운로드 실패", False)

            if log_callback:
                log_callback(f"  ✓ 파일 다운로드 완료: {success_count}개", False)

            if success_count > 0:
                has_changes = True

        # 매니페스트 저장
        if has_changes or files_to_delete:
            save_local_manifest(app_dir, remote_manifest)
            if log_callback:
                log_callback(f"✓ 업데이트 완료 (삭제: {len(files_to_delete)}개, 업데이트: {len(files_to_update)}개)", False)
            return True

        return False

    except Exception as e:
        if log_callback:
            log_callback(f"✗ 업데이트 실패: {e}", False)
        print(f"업데이트 실패: {e}")
        return False


def check_and_update_application(popup_window, repo_owner, repo_name, branch=APP_BRANCH):
    """UI와 연동된 업데이트 확인 및 실행"""
    app_dir = get_app_directory()
    result = {"updated": False, "error": None}

    try:
        popup_window.safe_add_log("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━", False)
        popup_window.safe_add_log("[ 애플리케이션 업데이트 확인 ]", False)
        popup_window.safe_add_log("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━", False)
        time.sleep(0.2)

        def log_callback(msg, is_progress):
            popup_window.safe_add_log(msg, is_progress)
            time.sleep(0.05)

        updated = update_application(repo_owner, repo_name, app_dir, log_callback, branch)
        
        if updated:
            popup_window.safe_add_log("", False)
            popup_window.safe_add_log("✓ 업데이트가 완료되었습니다", False)
            popup_window.safe_add_log("✓ 프로그램을 다시 시작해주세요", False)
            result["updated"] = True
        else:
            popup_window.safe_add_log("", False)
            popup_window.safe_add_log("✓ 최신 버전입니다", False)

    except Exception as e:
        result["error"] = str(e)
        popup_window.safe_add_log(f"✗ 오류 발생: {e}", False)

    return result


# ---------- Poppler 관련 ----------

def get_poppler_directory():
    """Poppler 설치 디렉토리 반환 (애플리케이션 하위)"""
    app_dir = get_app_directory()
    poppler_dir = app_dir / POPPLER_FOLDER_NAME
    poppler_dir.mkdir(parents=True, exist_ok=True)
    return poppler_dir


def get_latest_poppler_download_url(log_callback=None):
    """Poppler 최신 윈도우 zip 다운로드 URL 및 파일명 가져오기"""
    api_url = f"https://api.github.com/repos/{POPPLER_REPO_OWNER}/{POPPLER_REPO_NAME}/releases/latest"
    try:
        if log_callback:
            log_callback("  → GitHub API 요청 중...", False)
        resp = requests.get(api_url, timeout=10)
        resp.raise_for_status()
        data = resp.json()
        if log_callback:
            log_callback("  ✓ 최신 릴리즈 정보 확인 완료", False)
        
        for asset in data.get("assets", []):
            if asset["name"].lower().endswith(".zip"):
                if log_callback:
                    log_callback(f"  ✓ 다운로드 파일: {asset['name']}", False)
                return asset["browser_download_url"], asset["name"]
        return None, None
    except Exception as e:
        if log_callback:
            log_callback(f"  ✗ API 요청 실패: {e}", False)
        return None, None


def find_poppler_folder(base_dir: Path, log_callback=None):
    """Poppler 설치 폴더 찾기"""
    if log_callback:
        log_callback(f"  → Poppler 폴더 검색 중: {base_dir}", False)
    
    for folder in base_dir.iterdir():
        if folder.is_dir() and "poppler" in folder.name.lower():
            bin_path = folder / "Library" / "bin"
            if bin_path.exists() and (bin_path / "pdftoppm.exe").exists():
                version_match = re.search(r'(\d+\.\d+(?:\.\d+)?)', folder.name)
                version = version_match.group(1) if version_match else "unknown"
                if log_callback:
                    log_callback(f"  ✓ Poppler 발견: {folder.name}", False)
                    log_callback(f"  ✓ 버전: v{version}", False)
                    log_callback(f"  ✓ 경로: {bin_path}", False)
                return version, str(bin_path)
    
    if log_callback:
        log_callback("  ! Poppler를 찾을 수 없습니다", False)
    return None, None


def download_and_extract_poppler(dest_folder: Path, log_callback=None):
    """Poppler 다운로드 및 압축 해제"""
    if log_callback:
        log_callback("→ Poppler 다운로드 준비", False)

    download_url, filename = get_latest_poppler_download_url(log_callback)
    if not download_url:
        raise RuntimeError("Poppler 윈도우용 최신 zip 파일을 찾을 수 없습니다.")

    zip_path = dest_folder / filename

    if log_callback:
        log_callback(f"→ Poppler 다운로드 시작: {filename}", False)

    # 다운로드 진행
    try:
        with requests.get(download_url, stream=True, timeout=60) as r:
            r.raise_for_status()
            total_size = int(r.headers.get('content-length', 0))
            downloaded = 0
            last_percent = 0
            
            with open(zip_path, 'wb') as f:
                for chunk in r.iter_content(chunk_size=65536):
                    f.write(chunk)
                    downloaded += len(chunk)
                    if total_size > 0:
                        percent = int((downloaded / total_size) * 100)
                        if percent - last_percent >= 10:
                            if log_callback:
                                log_callback(f"  → 다운로드 진행: {percent}% ({downloaded:,} / {total_size:,} bytes)", True)
                            last_percent = percent
        
        if log_callback:
            log_callback(f"  ✓ 다운로드 완료: {filename}", False)
    except Exception as e:
        if log_callback:
            log_callback(f"  ✗ 다운로드 실패: {e}", False)
        raise

    # 압축 해제
    if log_callback:
        log_callback("→ Poppler 압축 해제 시작", False)

    release_folder_name = filename.replace(".zip", "")
    release_folder = dest_folder / release_folder_name
    if release_folder.exists():
        shutil.rmtree(release_folder)

    try:
        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
            members = zip_ref.namelist()
            total_files = len(members)
            last_percent = 0
            
            for i, member in enumerate(members, start=1):
                zip_ref.extract(member, dest_folder)
                percent = int((i / total_files) * 100)
                if percent - last_percent >= 10:
                    if log_callback:
                        log_callback(f"  → 압축 해제 진행: {percent}% ({i} / {total_files} 파일)", True)
                    last_percent = percent
        
        if log_callback:
            log_callback(f"  ✓ 압축 해제 완료: {total_files}개 파일", False)
    except Exception as e:
        if log_callback:
            log_callback(f"  ✗ 압축 해제 실패: {e}", False)
        raise
    finally:
        if zip_path.exists():
            os.remove(zip_path)
            if log_callback:
                log_callback("  ✓ 임시 zip 파일 삭제 완료", False)

    # bin 폴더 탐색
    if log_callback:
        log_callback("→ Poppler bin 폴더 검색 중...", False)

    def find_bin_folder(base_path):
        for root, dirs, files in os.walk(base_path):
            root_path = Path(root)
            if root_path.name == "bin" and root_path.parent.name == "Library":
                if (root_path / "pdftoppm.exe").exists():
                    return root_path
        return None

    bin_path = find_bin_folder(dest_folder)

    if bin_path and bin_path.exists():
        if log_callback:
            log_callback(f"  ✓ bin 폴더 발견: {bin_path}", False)
            log_callback("✓ Poppler 설치 완료", False)
        return bin_path

    raise RuntimeError("Poppler bin 폴더를 찾을 수 없습니다.")


def prepare_poppler_path_with_ui(popup):
    """Poppler 경로 준비 기능 (UI 연동)"""
    poppler_dir = get_poppler_directory()
    result = {"path": None, "error": None}

    try:
        popup.safe_add_log("", False)
        popup.safe_add_log("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━", False)
        popup.safe_add_log("[ Poppler 유효성 검사 ]", False)
        popup.safe_add_log("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━", False)
        time.sleep(0.2)

        def log_callback(msg, is_progress):
            popup.safe_add_log(msg, is_progress)
            time.sleep(0.05)

        log_callback(f"→ Poppler 설치 경로: {poppler_dir}", False)

        latest_version = None
        try:
            log_callback("→ 최신 버전 확인 중...", False)
            download_url, latest_filename = get_latest_poppler_download_url(log_callback)
            if latest_filename:
                match = re.search(r'(\d+\.\d+\.\d+)', latest_filename)
                if match:
                    latest_version = match.group(1)
                    log_callback(f"  ✓ 최신 버전: v{latest_version}", False)
        except Exception as e:
            log_callback(f"  ! 최신 버전 확인 실패: {e}", False)
            latest_version = None

        log_callback("→ 로컬 Poppler 검색 중...", False)
        version, poppler_path = find_poppler_folder(poppler_dir, log_callback)

        if poppler_path and version:
            needs_update = latest_version and version != latest_version

            if needs_update:
                log_callback("", False)
                log_callback("! 이전 버전이 설치되어 있습니다", False)
                log_callback(f"  • 설치된 버전: v{version}", False)
                log_callback(f"  • 최신 버전: v{latest_version}", False)
                log_callback("→ 이전 버전 삭제 중...", False)
                time.sleep(0.2)
                
                for folder in poppler_dir.iterdir():
                    if folder.is_dir() and "poppler" in folder.name.lower():
                        try:
                            shutil.rmtree(folder)
                            log_callback(f"  ✓ 삭제 완료: {folder.name}", False)
                        except Exception as e:
                            log_callback(f"  ✗ 삭제 실패: {e}", False)

                log_callback("", False)
                poppler_bin_path = download_and_extract_poppler(poppler_dir, log_callback)
                result["path"] = poppler_bin_path
            else:
                log_callback("", False)
                log_callback("✓ 최신 버전이 설치되어 있습니다", False)
                result["path"] = poppler_path

        else:
            log_callback("", False)
            log_callback("! Poppler를 찾을 수 없습니다", False)
            log_callback("→ Poppler 다운로드 시작", False)
            time.sleep(0.2)

            poppler_bin_path = download_and_extract_poppler(poppler_dir, log_callback)
            result["path"] = poppler_bin_path

    except Exception as e:
        result["error"] = str(e)
        try:
            popup.safe_add_log(f"✗ 오류 발생: {str(e)}", False)
        except Exception:
            pass

    return result["path"], result["error"]


# ---------- UI 컴포넌트 정의 ----------

class UnifiedInstallPopup(ctk.CTkToplevel):
    """초기화 및 업데이트 진행 로그 표시 팝업"""

    def __init__(self, parent):
        super().__init__(parent)
        self.title("초기화 중...")
        self.geometry("700x400")
        self.resizable(False, False)

        # 화면 중앙 위치 조정
        self.update_idletasks()
        x = (self.winfo_screenwidth() // 2) - (700 // 2)
        y = (self.winfo_screenheight() // 2) - (400 // 2)
        self.geometry(f"700x400+{x}+{y}")

        self.grab_set()

        icon_path = get_icon_path()
        if icon_path:
            self.after(50, lambda: set_window_icon(self, icon_path))

        self._closing = False
        self.log_queue = queue.Queue()
        self.logs = []

        self.log_box = ctk.CTkTextbox(self, height=370, width=650, font=("Consolas", 10))
        self.log_box.pack(pady=10, padx=25)
        self.log_box.configure(state="disabled")

        self.process_log_queue()

    def add_log(self, message, is_progress=False):
        """메인 스레드에서 호출하는 로그 추가"""
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

    def process_log_queue(self):
        """큐에서 로그 읽어 UI에 반영"""
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
        except Exception:
            pass

    def safe_add_log(self, message, is_progress=False):
        """스레드 안전 로그 추가 (큐에 저장)"""
        try:
            self.log_queue.put((message, is_progress))
        except Exception as e:
            print(f"safe_add_log 실패: {e}")

    def close_window(self):
        """윈도우 종료"""
        try:
            self._closing = True
            self.grab_release()
            self.destroy()
        except Exception as e:
            print(f"창 닫기 실패: {e}")


class ProgressPopup(ctk.CTkToplevel):
    """PDF 변환 진행 상황 표시 팝업"""

    def __init__(self, parent, total_files, total_pages):
        super().__init__(parent)
        self.title("진행 중")
        self.geometry("450x170")
        self.resizable(False, False)
        self.grab_set()

        # 중앙 위치 설정
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

        self.master.withdraw()

        icon_path = get_icon_path()
        if icon_path:
            self.master.iconbitmap(icon_path)

        # 초기화 팝업(업데이트 + Poppler 설치)
        self.initialize_resources()

        self.pdf_files = []
        self._cancel_requested = False
        self.progress_popup = None

        # UI 생성
        self.create_widgets()

        self.master.deiconify()
        self.center_window()

        # 드래그 앤 드롭 바인딩
        self.master.drop_target_register(DND_FILES)
        self.master.dnd_bind("<<Drop>>", self.on_drop)

    def initialize_resources(self):
        """초기 업데이트 및 poppler 경로 준비"""
        self.unified_popup = UnifiedInstallPopup(self.master)
        self.result = {"update_done": False, "poppler_path": None, "should_close": False, "restart_required": False}

        def init_thread():
            try:
                # 업데이트 확인
                update_res = check_and_update_application(
                    self.unified_popup, APP_REPO_OWNER, APP_REPO_NAME, APP_BRANCH)
                self.result["update_done"] = update_res.get("updated", False)
                time.sleep(0.3)

                # Poppler 확인
                poppler_path, poppler_err = prepare_poppler_path_with_ui(self.unified_popup)
                self.result["poppler_path"] = poppler_path

                if self.result["update_done"]:
                    self.unified_popup.safe_add_log("", False)
                    self.unified_popup.safe_add_log("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━", False)
                    self.unified_popup.safe_add_log("✓ 업데이트 완료 - 재시작 필요", False)
                    self.unified_popup.safe_add_log("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━", False)
                    self.result["restart_required"] = True
                else:
                    self.unified_popup.safe_add_log("", False)
                    self.unified_popup.safe_add_log("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━", False)
                    for i in range(3, 0, -1):
                        self.unified_popup.safe_add_log(f"✓ {i}초 후 프로그램이 시작됩니다...", False)
                        time.sleep(1.0)
                    self.unified_popup.safe_add_log("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━", False)

                self.result["should_close"] = True

            except Exception as e:
                print(f"초기화 오류: {e}")
                self.unified_popup.safe_add_log(f"✗ 초기화 오류: {e}", False)
                self.result["should_close"] = True

        def start_init():
            threading.Thread(target=init_thread, daemon=True).start()

        self.unified_popup.after(100, start_init)

        def check_close():
            if self.result["should_close"]:
                try:
                    self.unified_popup.close_window()
                except Exception:
                    pass
            else:
                try:
                    self.unified_popup.after(100, check_close)
                except Exception:
                    pass

        self.unified_popup.after(500, check_close)
        try:
            self.master.wait_window(self.unified_popup)
        except Exception:
            pass

        if self.result["restart_required"]:
            messagebox.showinfo("업데이트 완료", "애플리케이션이 업데이트되었습니다.\n프로그램을 다시 시작해주세요.")
            self.master.destroy()
            sys.exit()

        if not self.result["poppler_path"]:
            messagebox.showerror("오류", "Poppler를 준비하지 못했습니다. 프로그램을 종료합니다.")
            self.master.destroy()
            sys.exit()

        self.poppler_path = self.result["poppler_path"]

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

        # 버전 표시 (로컬 manifest에서 가져오기)
        app_dir = get_app_directory()
        local_manifest = get_local_manifest(app_dir)
        version = local_manifest.get("version", "버전 확인 불가")
        if version == "0.0.0":
            version = "버전 확인 불가"
        version_label = ctk.CTkLabel(control_container, text=f"v{version}", text_color="gray")
        version_label.pack(side="right", padx=5)

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
