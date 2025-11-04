"""
PDF to JPG Converter with Auto-Update
Author: 류호준
"""

import os
import sys
import shutil
import threading
import subprocess
import re
import queue
import time
import zipfile
import tkinter as tk
from pathlib import Path
from typing import List, Optional, Callable, Tuple
from dataclasses import dataclass
from contextlib import suppress

import customtkinter as ctk
from tkinterdnd2 import TkinterDnD, DND_FILES
from pdf2image import convert_from_path, pdfinfo_from_path
from tkinter import messagebox, filedialog
import ctypes


# ==================== SSL 및 초기화 ====================


def init_ssl() -> bool:
    """SSL 인증서 설정"""
    try:
        import certifi
        os.environ['REQUESTS_CA_BUNDLE'] = certifi.where()
        os.environ['SSL_CERT_FILE'] = certifi.where()
        return True
    except (ImportError, AttributeError) as e:
        print(f"Warning: SSL certificate setup failed: {e}")
        with suppress(ImportError):
            import urllib3
            urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
        return False


VERIFY_SSL = init_ssl()
import requests


def init_theme() -> None:
    """CustomTkinter 테마 초기화"""
    ctk.set_default_color_theme("blue")
    ctk.set_appearance_mode("system")


init_theme()


# ==================== 설정 ====================


@dataclass(frozen=True)
class AppConfig:
    """애플리케이션 설정"""
    APP_TITLE: str = "PDF → JPG 변환기 [made by. 류호준]"
    APP_SIZE: str = "600x300"
    INIT_WINDOW_SIZE: str = "700x400"
    ICON_FILENAME: str = "icon.ico"
    POPPLER_FOLDER_NAME: str = "poppler"
    
    APP_REPO_OWNER: str = "c-closed"
    APP_REPO_NAME: str = "P2J"
    CURRENT_VERSION: str = "2.1.5"
    POPPLER_REPO_OWNER: str = "oschwartz10612"
    POPPLER_REPO_NAME: str = "poppler-windows"
    
    CONVERSION_DPI: int = 200
    OUTPUT_FORMAT: str = "jpeg"
    
    REQUEST_TIMEOUT: int = 10
    DOWNLOAD_TIMEOUT: int = 90
    DOWNLOAD_CHUNK_SIZE: int = 65536
    PROGRESS_UPDATE_INTERVAL: int = 1
    LOG_PROCESS_INTERVAL: int = 100


CONFIG = AppConfig()


# ==================== 유틸리티 ====================


class PathUtils:
    """경로 관리"""
    
    @staticmethod
    def get_app_directory() -> Path:
        if getattr(sys, 'frozen', False):
            return Path(sys.executable).parent
        return Path(__file__).parent
    
    @staticmethod
    def get_icon_path() -> Optional[str]:
        icon_path = PathUtils.get_app_directory() / CONFIG.ICON_FILENAME
        return str(icon_path) if icon_path.exists() else None
    
    @staticmethod
    def get_poppler_path() -> Optional[str]:
        poppler_dir = PathUtils.get_app_directory() / CONFIG.POPPLER_FOLDER_NAME
        
        if not poppler_dir.exists():
            return None
        
        for item in poppler_dir.iterdir():
            if item.is_dir() and "poppler" in item.name.lower():
                bin_path = item / "Library" / "bin"
                if (bin_path / "pdftoppm.exe").exists():
                    return str(bin_path)
        
        return None


class IconManager:
    """아이콘 관리"""
    
    WM_SETICON = 0x80
    ICON_SMALL = 0
    ICON_BIG = 1
    LR_LOADFROMFILE = 0x00000010
    
    @staticmethod
    def set_window_icon(window, icon_path: str) -> bool:
        if not icon_path or not os.path.exists(icon_path):
            return False
        
        try:
            window.update_idletasks()
            hwnd = window.winfo_id()
            if not hwnd:
                return False
            
            load_image = ctypes.windll.user32.LoadImageW
            hicon_small = load_image(0, icon_path, 1, 16, 16, IconManager.LR_LOADFROMFILE)
            hicon_big = load_image(0, icon_path, 1, 32, 32, IconManager.LR_LOADFROMFILE)
            
            if not (hicon_small or hicon_big):
                return False
            
            send_message = ctypes.windll.user32.SendMessageW
            if hicon_small:
                send_message(hwnd, IconManager.WM_SETICON, IconManager.ICON_SMALL, hicon_small)
            if hicon_big:
                send_message(hwnd, IconManager.WM_SETICON, IconManager.ICON_BIG, hicon_big)
            
            return True
        except Exception:
            return False


class VersionManager:
    """버전 관리"""
    
    @staticmethod
    def parse_version(version_str: str) -> Tuple[int, int, int]:
        try:
            clean_ver = version_str.lstrip('v')
            parts = clean_ver.split('.')
            return tuple(int(p) for p in parts[:3])
        except (ValueError, AttributeError):
            return (0, 0, 0)
    
    @staticmethod
    def is_newer(current: str, latest: str) -> bool:
        return VersionManager.parse_version(latest) > VersionManager.parse_version(current)


# ==================== GitHub API ====================


class GitHubAPIClient:
    """GitHub API 클라이언트"""
    
    def get_latest_release(self, owner: str, repo: str) -> Optional[dict]:
        api_url = f"https://api.github.com/repos/{owner}/{repo}/releases/latest"
        
        try:
            resp = requests.get(api_url, timeout=CONFIG.REQUEST_TIMEOUT, verify=VERIFY_SSL)
            resp.raise_for_status()
            return resp.json()
        except requests.RequestException:
            return None
    
    def download_file(self, url: str, dest_path: Path,
                     progress_callback: Optional[Callable[[int], None]] = None) -> bool:
        try:
            with requests.get(url, stream=True, timeout=CONFIG.DOWNLOAD_TIMEOUT, 
                            verify=VERIFY_SSL) as resp:
                resp.raise_for_status()
                
                total_size = int(resp.headers.get('content-length', 0))
                downloaded = 0
                last_percent = -1
                
                with open(dest_path, 'wb') as f:
                    for chunk in resp.iter_content(chunk_size=CONFIG.DOWNLOAD_CHUNK_SIZE):
                        if not chunk:
                            continue
                        
                        f.write(chunk)
                        downloaded += len(chunk)
                        
                        if total_size and progress_callback:
                            percent = int((downloaded / total_size) * 100)
                            if percent != last_percent and percent % CONFIG.PROGRESS_UPDATE_INTERVAL == 0:
                                progress_callback(percent)
                                last_percent = percent
            
            if progress_callback:
                progress_callback(100)
            
            return True
        except Exception:
            dest_path.unlink(missing_ok=True)
            return False


# ==================== Release 관리 ====================


@dataclass
class ReleaseInfo:
    """Release 정보"""
    version: str
    name: str
    msi_url: str
    msi_filename: str
    msi_size: int


class ReleaseManager:
    """Release 관리"""
    
    def __init__(self, api_client: GitHubAPIClient):
        self.api_client = api_client
    
    def get_latest_release_info(self, log_callback: Optional[Callable] = None) -> Optional[ReleaseInfo]:
        self._log(log_callback, "  → GitHub Release 확인 중...")
        
        data = self.api_client.get_latest_release(CONFIG.APP_REPO_OWNER, CONFIG.APP_REPO_NAME)
        
        if not data:
            self._log(log_callback, "  ! Release 확인 실패")
            return None
        
        tag_name = data.get('tag_name', '')
        self._log(log_callback, f"  ✓ 최신 Release 버전: {tag_name}")
        
        msi_asset = next((a for a in data.get('assets', []) 
                         if a['name'].lower().endswith('.msi')), None)
        
        if not msi_asset:
            self._log(log_callback, "  ! Release에 .msi 파일이 없습니다")
            return None
        
        return ReleaseInfo(
            version=tag_name,
            name=data.get('name', ''),
            msi_url=msi_asset['browser_download_url'],
            msi_filename=msi_asset['name'],
            msi_size=msi_asset['size']
        )
    
    def download_msi(self, release_info: ReleaseInfo, dest_folder: Path,
                    log_callback: Optional[Callable] = None) -> Optional[Path]:
        msi_path = dest_folder / release_info.msi_filename
        size_mb = release_info.msi_size / (1024 * 1024)
        
        self._log(log_callback, f"→ MSI 다운로드 시작: {release_info.msi_filename}")
        self._log(log_callback, f"  • 파일 크기: {size_mb:.1f} MB")
        
        def progress_callback(percent: int):
            if percent >= 100:
                self._log(log_callback, "  ✓ 다운로드 완료", True)
            else:
                self._log(log_callback, f"  → 다운로드 진행: {percent}%", True)
        
        if self.api_client.download_file(release_info.msi_url, msi_path, progress_callback):
            return msi_path
        
        self._log(log_callback, "  ✗ 다운로드 실패")
        return None
    
    @staticmethod
    def run_msi_installer(msi_path: Path, log_callback: Optional[Callable] = None) -> bool:
        """MSI 설치 프로그램을 독립 프로세스로 실행"""
        if not msi_path.exists():
            ReleaseManager._log(log_callback, "  ✗ MSI 파일을 찾을 수 없습니다")
            return False

        try:
            ReleaseManager._log(log_callback, "→ MSI 설치 프로그램 실행 중...")
            ReleaseManager._log(log_callback, "  • 설치 창이 백그라운드에서 열립니다")

            # Windows 전용: 완전히 독립된 프로세스로 실행
            if sys.platform == 'win32':
                subprocess.Popen(
                    ['msiexec', '/i', str(msi_path)],
                    shell=False,
                    stdin=subprocess.DEVNULL,
                    stdout=subprocess.DEVNULL,
                    stderr=subprocess.DEVNULL,
                    close_fds=True,
                    creationflags=subprocess.CREATE_NEW_PROCESS_GROUP | subprocess.DETACHED_PROCESS
                )
            else:
                # 기타 OS (사용 안함)
                subprocess.Popen(['msiexec', '/i', str(msi_path)], shell=False)

            ReleaseManager._log(log_callback, "  ✓ 설치 프로그램 실행 완료")
            time.sleep(1)
            return True
        except Exception as e:
            ReleaseManager._log(log_callback, f"  ✗ 설치 프로그램 실행 실패: {e}")
            return False

    @staticmethod
    def _log(callback: Optional[Callable], message: str, is_progress: bool = False):
        if callback:
            callback(message, is_progress)


# ==================== Poppler 관리 ====================


class PopplerManager:
    """Poppler 관리"""
    
    VERSION_PATTERN = re.compile(r'(\d+\.\d+\.\d+)')
    
    def __init__(self, api_client: GitHubAPIClient):
        self.api_client = api_client
        self._install_completed = False
    
    def get_installed_version(self, poppler_dir: Path) -> Optional[str]:
        if not poppler_dir.exists():
            return None
        
        for item in poppler_dir.iterdir():
            if item.is_dir() and "poppler" in item.name.lower():
                match = self.VERSION_PATTERN.search(item.name)
                if match:
                    return match.group(1)
        
        return None
    
    def get_latest_version_info(self) -> Optional[Tuple[str, str, str]]:
        data = self.api_client.get_latest_release(
            CONFIG.POPPLER_REPO_OWNER, CONFIG.POPPLER_REPO_NAME
        )
        
        if not data:
            return None
        
        zip_asset = next((a for a in data.get("assets", [])
                         if a["name"].lower().endswith(".zip")), None)
        
        if not zip_asset:
            return None
        
        filename = zip_asset["name"]
        match = self.VERSION_PATTERN.search(filename)
        version = match.group(1) if match else "unknown"
        
        return (zip_asset["browser_download_url"], filename, version)
    
    def download_and_extract(self, dest_folder: Path,
                            log_callback: Optional[Callable] = None) -> Path:
        version_info = self.get_latest_version_info()
        if not version_info:
            raise RuntimeError("Poppler 최신 버전을 찾을 수 없습니다.")
        
        download_url, filename, version = version_info
        zip_path = dest_folder / filename
        
        self._log(log_callback, f"→ 다운로드 파일: {filename}")
        self._log(log_callback, "→ 다운로드 시작...")
        
        def progress_callback(percent: int):
            if percent >= 100:
                self._log(log_callback, "  ✓ 다운로드 완료", True)
            else:
                self._log(log_callback, f"  → 다운로드 진행: {percent}%", True)
        
        if not self.api_client.download_file(download_url, zip_path, progress_callback):
            raise RuntimeError("Poppler 다운로드 실패")
        
        self._log(log_callback, "→ 압축 해제 시작")
        
        try:
            with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                members = zip_ref.namelist()
                total_files = len(members)
                last_percent = -1
                
                for i, member in enumerate(members, start=1):
                    zip_ref.extract(member, dest_folder)
                    percent = int((i / total_files) * 100)
                    
                    if percent != last_percent and percent % CONFIG.PROGRESS_UPDATE_INTERVAL == 0:
                        if percent >= 100:
                            self._log(log_callback, f"  ✓ 압축 해제 완료: {total_files}개 파일", True)
                        else:
                            self._log(log_callback, f"  → 압축 해제 진행: {percent}%", True)
                        last_percent = percent
        finally:
            zip_path.unlink(missing_ok=True)
            self._log(log_callback, "  ✓ 임시 파일 삭제 완료")
        
        self._log(log_callback, "→ Poppler 폴더 확인 중...")
        
        for item in dest_folder.iterdir():
            if item.is_dir() and "poppler" in item.name.lower():
                bin_path = item / "Library" / "bin"
                if (bin_path / "pdftoppm.exe").exists():
                    if not self._install_completed:
                        self._log(log_callback, f"  ✓ Poppler 설치 완료 (v{version})")
                        self._install_completed = True
                    return bin_path
        
        raise RuntimeError("Poppler 폴더를 찾을 수 없습니다.")
    
    def check_and_update(self, app_dir: Path, log_callback: Optional[Callable] = None) -> None:
        self._log(log_callback, "")
        self._log(log_callback, "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━")
        self._log(log_callback, "[ Poppler 확인 ]")
        self._log(log_callback, "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━")
        time.sleep(0.2)
        
        poppler_dir = app_dir / CONFIG.POPPLER_FOLDER_NAME
        poppler_dir.mkdir(parents=True, exist_ok=True)
        
        installed_version = self.get_installed_version(poppler_dir)
        
        self._log(log_callback, "→ 최신 버전 확인 중...")
        version_info = self.get_latest_version_info()
        
        if not version_info:
            self._log(log_callback, "  ✗ 최신 버전 확인 실패")
            if installed_version:
                self._log(log_callback, "")
                self._log(log_callback, f"✓ 기존 Poppler 사용 (v{installed_version})")
            else:
                self._log(log_callback, "")
                self._log(log_callback, "✗ Poppler 설치 불가")
                raise RuntimeError("Poppler 최신 버전 확인 실패")
            return
        
        _, _, latest_version = version_info
        self._log(log_callback, f"  • 최신 버전: v{latest_version}")
        
        if installed_version:
            self._log(log_callback, f"  • 설치된 버전: v{installed_version}")
            
            if installed_version == latest_version:
                if not self._install_completed:
                    self._log(log_callback, "")
                    self._log(log_callback, "✓ Poppler가 최신 버전입니다")
                self._install_completed = True
                return
            
            self._log(log_callback, "")
            self._log(log_callback, "! 새 버전 발견 - Poppler 업데이트 시작")
            self._log(log_callback, "→ 기존 Poppler 삭제 중...")
            
            self._remove_old_poppler(poppler_dir, log_callback)
            self._log(log_callback, "")
            self._install_completed = False
        else:
            self._log(log_callback, "")
            self._log(log_callback, "! Poppler가 설치되지 않았습니다")
            self._log(log_callback, "")
            time.sleep(0.2)
            self._install_completed = False
        
        try:
            self.download_and_extract(poppler_dir, log_callback)
        except Exception as e:
            self._log(log_callback, f"✗ Poppler 설치 실패: {e}")
            raise
    
    def _remove_old_poppler(self, poppler_dir: Path, log_callback: Optional[Callable]):
        for item in poppler_dir.iterdir():
            if item.is_dir() and "poppler" in item.name.lower():
                try:
                    shutil.rmtree(item)
                    self._log(log_callback, f"  ✓ 삭제 완료: {item.name}")
                except Exception as e:
                    self._log(log_callback, f"  ✗ 삭제 실패: {e}")
    
    @staticmethod
    def _log(callback: Optional[Callable], message: str, is_progress: bool = False):
        if callback:
            callback(message, is_progress)


# ==================== PDF 처리 ====================


class PDFProcessor:
    """PDF 처리"""
    
    def __init__(self, poppler_path: str):
        self.poppler_path = poppler_path
    
    def get_page_count(self, pdf_path: str) -> int:
        try:
            info = pdfinfo_from_path(pdf_path, poppler_path=self.poppler_path)
            return info["Pages"]
        except Exception as e:
            raise RuntimeError(f"PDF 정보 읽기 실패: {e}")
    
    def get_total_pages(self, pdf_files: List[str]) -> int:
        return sum(self.get_page_count(pdf) for pdf in pdf_files)
    
    def convert_to_images(self, pdf_path: str, output_folder: Path,
                         progress_callback: Optional[Callable[[int], None]] = None) -> int:
        info = pdfinfo_from_path(pdf_path, poppler_path=self.poppler_path)
        total_pages = info["Pages"]
        digits = len(str(total_pages))
        
        images = convert_from_path(
            pdf_path,
            dpi=CONFIG.CONVERSION_DPI,
            first_page=1,
            last_page=total_pages,
            fmt=CONFIG.OUTPUT_FORMAT,
            output_folder=str(output_folder),
            paths_only=True,
            poppler_path=self.poppler_path
        )
        
        for i, img_path in enumerate(images, start=1):
            dest_path = output_folder / f"{str(i).zfill(digits)}.jpg"
            shutil.move(img_path, dest_path)
            
            if progress_callback:
                progress_callback(i)
        
        return len(images)


# ==================== UI: 초기화 창 (순수 tkinter) ====================


class InitializationWindow(tk.Tk):
    """초기화 창 - 순수 tkinter"""
    
    def __init__(self):
        super().__init__()
        self.title("파일 무결성 검사 중...")
        self.geometry(CONFIG.INIT_WINDOW_SIZE)
        self.resizable(False, False)
        
        self.app_dir = PathUtils.get_app_directory()
        self.result = {
            "should_close": False,
            "launch_main": True,
            "update_started": False
        }
        
        self.logs = []
        self._closing = False
        
        self._setup_icon()
        self._center_window()
        self._setup_ui()
        self._start_initialization()
    
    def _setup_icon(self) -> None:
        icon_path = PathUtils.get_icon_path()
        if icon_path:
            with suppress(Exception):
                self.iconbitmap(icon_path)
    
    def _center_window(self) -> None:
        self.update_idletasks()
        width = self.winfo_width()
        height = self.winfo_height()
        x = (self.winfo_screenwidth() - width) // 2
        y = (self.winfo_screenheight() - height) // 2
        self.geometry(f"{width}x{height}+{x}+{y}")
    
    def _setup_ui(self) -> None:
        self.log_box = tk.Text(
            self, 
            height=22, 
            width=85,
            font=("Consolas", 8),
            bg="#ffffff",
            fg="#000000",
            state="disabled"
        )
        self.log_box.pack(pady=10, padx=25, fill="both", expand=True)
    
    def _start_initialization(self) -> None:
        threading.Thread(target=self._init_thread, daemon=True).start()
        self.after(500, self._check_close)
    
    def _init_thread(self) -> None:
        try:
            log_cb = lambda msg, is_progress: self._add_log(msg, is_progress)
            
            api_client = GitHubAPIClient()
            
            self._check_update(api_client, log_cb)
            if self.result["update_started"]:
                return
            
            time.sleep(0.3)
            
            poppler_manager = PopplerManager(api_client)
            poppler_manager.check_and_update(self.app_dir, log_cb)
            time.sleep(0.3)
            
            self._countdown()
            self.result["should_close"] = True
        
        except Exception as e:
            self._add_log(f"✗ 초기화 오류: {e}", False)
            self.result["should_close"] = True
            self.result["launch_main"] = False
    
    def _check_update(self, api_client: GitHubAPIClient, log_cb: Callable) -> None:
        self._add_log("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━", False)
        self._add_log("[ 업데이트 확인 ]", False)
        self._add_log("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━", False)
        self._add_log(f"  • 현재 버전: v{CONFIG.CURRENT_VERSION}", False)
        
        release_manager = ReleaseManager(api_client)
        release_info = release_manager.get_latest_release_info(log_cb)
        
        if not release_info:
            self._add_log("", False)
            self._add_log("✓ 업데이트 확인 완료", False)
            return
        
        if not VersionManager.is_newer(CONFIG.CURRENT_VERSION, release_info.version):
            self._add_log("", False)
            self._add_log("✓ 이미 최신 버전입니다", False)
            return
        
        self._add_log("", False)
        self._add_log("! 새 버전 발견!", False)
        self._add_log(f"  • 새 버전: {release_info.version}", False)
        self._add_log(f"  • 파일명: {release_info.msi_filename}", False)
        self._add_log("", False)
        
        msi_path = release_manager.download_msi(release_info, self.app_dir, log_cb)
        
        if msi_path and release_manager.run_msi_installer(msi_path, log_cb):
            self._add_log("", False)
            self._add_log("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━", False)
            self._add_log("✓ 업데이트 시작됨", False)
            self._add_log("설치 프로그램이 실행되었습니다.", False)
            self._add_log("프로그램을 종료합니다.", False)
            self._add_log("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━", False)
            time.sleep(3)
            
            self.result["update_started"] = True
            self.result["should_close"] = True
            self.result["launch_main"] = False
    
    def _countdown(self) -> None:
        self._add_log("", False)
        self._add_log("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━", False)
        for i in range(3, 0, -1):
            if self._closing:
                break
            self._add_log(f"✓ {i}초 후 프로그램이 시작됩니다...", False)
            time.sleep(1.0)
        self._add_log("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━", False)
    
    def _add_log(self, message: str, is_progress: bool = False) -> None:
        if self._closing:
            return
        
        try:
            self.log_box.config(state="normal")
            
            if is_progress and self.logs:
                self.log_box.delete(f"{len(self.logs)}.0", f"{len(self.logs) + 1}.0")
                self.logs.pop()
            
            self.logs.append(message)
            self.log_box.insert("end", message + "\n")
            self.log_box.see("end")
            self.log_box.config(state="disabled")
            self.update_idletasks()
        except Exception:
            pass
    
    def _check_close(self) -> None:
        if self._closing:
            return
        
        if self.result["should_close"]:
            self._closing = True
            self.quit()
        else:
            self.after(100, self._check_close)


# ==================== UI: 진행 팝업 (customtkinter) ====================


class ProgressPopup(ctk.CTkToplevel):
    """진행 팝업"""
    
    def __init__(self, parent, total_files: int, total_pages: int):
        super().__init__(parent)
        self.total_files = total_files
        self.total_pages = total_pages
        self.cancelled = False
        self.cancel_callback = None
        self._auto_close_id = None
        
        self._setup_window(parent)
        self._setup_icon()
        self._create_widgets()
    
    def _setup_window(self, parent) -> None:
        self.title("진행 중")
        self.geometry("450x170")
        self.resizable(False, False)
        self.grab_set()
        
        self.update_idletasks()
        x = parent.winfo_x() + (parent.winfo_width() - 450) // 2
        y = parent.winfo_y() + (parent.winfo_height() - 170) // 2
        self.geometry(f"450x170+{x}+{y}")
    
    def _setup_icon(self) -> None:
        icon_path = PathUtils.get_icon_path()
        if icon_path:
            with suppress(Exception):
                self.iconbitmap(icon_path)
    
    def _create_widgets(self) -> None:
        self.file_label = ctk.CTkLabel(self, text=f"파일: 0 / {self.total_files}")
        self.file_label.pack(pady=(10, 5))
        
        self.file_progress = ctk.CTkProgressBar(self, width=400)
        self.file_progress.pack(pady=5)
        self.file_progress.set(0)
        
        self.page_label = ctk.CTkLabel(self, text=f"페이지: 0 / {self.total_pages}")
        self.page_label.pack(pady=(10, 5))
        
        self.page_progress = ctk.CTkProgressBar(self, width=400)
        self.page_progress.pack(pady=5)
        self.page_progress.set(0)
        
        self.cancel_button = ctk.CTkButton(self, text="취소", command=self._on_cancel)
        self.cancel_button.pack(pady=10)
    
    def _on_cancel(self) -> None:
        if messagebox.askyesno("작업 취소", "변환 작업을 정말 취소하시겠습니까?"):
            self.cancelled = True
            if self.cancel_callback:
                self.cancel_callback()
            self.cancel_button.configure(state="disabled")
    
    def update_file_progress(self, completed: int) -> None:
        self.file_label.configure(text=f"파일: {completed} / {self.total_files}")
        self.file_progress.set(completed / self.total_files if self.total_files else 0)
        self.update_idletasks()
    
    def update_page_progress(self, completed: int) -> None:
        self.page_label.configure(text=f"페이지: {completed} / {self.total_pages}")
        self.page_progress.set(completed / self.total_pages if self.total_pages else 0)
        self.update_idletasks()
    
    def show_completion(self) -> None:
        self.cancel_button.configure(text="확인 (3초)", state="normal", command=self._close)
        self._auto_close_id = self.after(1000, lambda: self._countdown(2))
    
    def _countdown(self, seconds: int) -> None:
        if seconds > 0:
            self.cancel_button.configure(text=f"확인 ({seconds}초)")
            self._auto_close_id = self.after(1000, lambda: self._countdown(seconds - 1))
        else:
            self._close()
    
    def _close(self) -> None:
        if self._auto_close_id:
            self.after_cancel(self._auto_close_id)
        self.destroy()


# ==================== UI: 메인 앱 (customtkinter) ====================


class PDFtoJPGApp(ctk.CTkFrame):
    """메인 애플리케이션"""
    
    def __init__(self, master):
        super().__init__(master)
        self.pack(fill="both", expand=True)
        self.master = master
        
        self.pdf_files: List[str] = []
        self._cancel_requested = False
        self.progress_popup: Optional[ProgressPopup] = None
        
        self._setup_window()
        self._check_poppler()
        self._create_widgets()
        self._setup_drag_drop()
    
    def _setup_window(self) -> None:
        self.master.title(CONFIG.APP_TITLE)
        self.master.geometry(CONFIG.APP_SIZE)
        
        icon_path = PathUtils.get_icon_path()
        if icon_path:
            with suppress(Exception):
                self.master.iconbitmap(icon_path)
        
        self.master.update_idletasks()
        x = (self.master.winfo_screenwidth() - 600) // 2
        y = (self.master.winfo_screenheight() - 300) // 2
        self.master.geometry(f"600x300+{x}+{y}")
    
    def _check_poppler(self) -> None:
        self.poppler_path = PathUtils.get_poppler_path()
        
        if not self.poppler_path:
            messagebox.showerror("오류", "Poppler를 찾을 수 없습니다.\n프로그램을 다시 시작해주세요.")
            self.master.destroy()
            sys.exit()
        
        self.pdf_processor = PDFProcessor(self.poppler_path)
    
    def _create_widgets(self) -> None:
        self.drop_area = ctk.CTkTextbox(self, height=220)
        self.drop_area.pack(padx=10, pady=10, fill="x")
        self.drop_area.configure(state="disabled")
        
        control_container = ctk.CTkFrame(self, fg_color="transparent")
        control_container.pack(pady=10, fill="x", padx=10)
        
        buttons = [
            ("불러오기", self.select_files),
            ("지우기", self.remove_selected),
            ("비우기", self.clear_list),
            ("변환하기", self.start_conversion)
        ]
        
        for text, command in buttons:
            btn = ctk.CTkButton(control_container, text=text, command=command, width=100)
            btn.pack(side="left", padx=5)
        
        version_label = ctk.CTkLabel(
            control_container, text=f"v{CONFIG.CURRENT_VERSION}", text_color="gray"
        )
        version_label.pack(side="right", padx=10)
    
    def _setup_drag_drop(self) -> None:
        self.master.drop_target_register(DND_FILES)
        self.master.dnd_bind("<<Drop>>", self._on_drop)
    
    def select_files(self) -> None:
        files = filedialog.askopenfilenames(
            title="PDF 파일 선택",
            filetypes=[("PDF files", "*.pdf")]
        )
        self._add_files(files)
    
    def _on_drop(self, event) -> None:
        files = self.master.tk.splitlist(event.data)
        pdf_files = [f for f in files if f.lower().endswith(".pdf")]
        self._add_files(pdf_files)
    
    def _add_files(self, files: List[str]) -> None:
        for f in files:
            if f not in self.pdf_files:
                self.pdf_files.append(f)
        self._update_file_list()
    
    def remove_selected(self) -> None:
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
                    self._update_file_list()
            else:
                messagebox.showwarning("경고", "유효한 파일을 선택해주세요.")
        except Exception as e:
            messagebox.showerror("오류", f"파일 제거 중 오류 발생: {e}")
    
    def clear_list(self) -> None:
        if self.pdf_files:
            if messagebox.askyesno("목록 비우기", "등록된 모든 파일을 목록에서 제거하시겠습니까?"):
                self.pdf_files.clear()
                self._update_file_list()
        else:
            messagebox.showinfo("알림", "목록이 이미 비어있습니다.")
    
    def _update_file_list(self) -> None:
        self.drop_area.configure(state="normal")
        self.drop_area.delete("0.0", "end")
        file_names = "\n".join(Path(f).name for f in self.pdf_files)
        self.drop_area.insert("0.0", file_names)
        self.drop_area.configure(state="disabled")
    
    def start_conversion(self) -> None:
        if not self.pdf_files:
            messagebox.showwarning("경고", "등록된 PDF 파일이 없습니다.")
            return
        
        try:
            total_files = len(self.pdf_files)
            total_pages = self.pdf_processor.get_total_pages(self.pdf_files)
        except Exception as e:
            messagebox.showerror("오류", str(e))
            return
        
        self.progress_popup = ProgressPopup(self.master, total_files, total_pages)
        self.progress_popup.cancel_callback = self._cancel_conversion
        self._cancel_requested = False
        
        threading.Thread(target=self._convert_files, daemon=True).start()
    
    def _cancel_conversion(self) -> None:
        self._cancel_requested = True
    
    def _convert_files(self) -> None:
        try:
            completed_files = 0
            completed_pages = 0
            
            for pdf_file in self.pdf_files:
                if self._cancel_requested:
                    break
                
                pdf_path = Path(pdf_file)
                output_folder = pdf_path.parent / f"JPG 변환({pdf_path.stem})"
                output_folder.mkdir(exist_ok=True)
                
                def page_callback(page_num):
                    nonlocal completed_pages
                    if not self._cancel_requested:
                        completed_pages += 1
                        self.progress_popup.update_page_progress(completed_pages)
                
                self.pdf_processor.convert_to_images(pdf_file, output_folder, page_callback)
                
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


# ==================== 메인 ====================


def main():
    """메인 함수"""
    init_window = None
    
    try:
        init_window = InitializationWindow()
        init_window.mainloop()
    except KeyboardInterrupt:
        sys.exit(0)
    except Exception as e:
        print(f"초기화 오류: {e}")
        sys.exit(1)
    
    result = {"update_started": False, "launch_main": False}
    
    if init_window:
        try:
            result["update_started"] = init_window.result["update_started"]
            result["launch_main"] = init_window.result["launch_main"]
        except:
            pass
        
        with suppress(Exception):
            init_window.destroy()
        
        del init_window
    
    time.sleep(0.5)
    
    if result["update_started"]:
        sys.exit(0)
    
    if not result["launch_main"]:
        sys.exit(1)
    
    try:
        root = TkinterDnD.Tk()
        app = PDFtoJPGApp(root)
        root.mainloop()
    except Exception as e:
        messagebox.showerror("오류", f"프로그램 실행 중 오류 발생:\n{e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
