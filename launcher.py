import os
import sys
import subprocess
import shutil
import threading
import queue
import ctypes
import hashlib
import json
import re
import time
import requests
import zipfile
import concurrent.futures
from pathlib import Path
from typing import Optional, Dict, List, Callable, Tuple

import customtkinter as ctk
from tkinter import messagebox


# ==================== 설정 ====================

class Config:
    """애플리케이션 설정"""
    THEME_COLOR = "blue"
    APP_REPO_OWNER = "c-closed"
    APP_REPO_NAME = "P2J"
    APP_BRANCH = "main"
    MAIN_EXE_NAME = "P2J.exe"
    ICON_FILENAME = "icon.ico"
    POPPLER_FOLDER_NAME = "poppler"
    POPPLER_REPO_OWNER = "oschwartz10612"
    POPPLER_REPO_NAME = "poppler-windows"
    
    # 다운로드 설정
    MAX_RETRIES = 5
    RETRY_DELAY = 5
    DOWNLOAD_TIMEOUT = 90
    MAX_CONCURRENT_DOWNLOADS = 3
    
    # GitHub Token (옵션 - Rate Limit 해결용)
    GITHUB_TOKEN = os.getenv("GITHUB_TOKEN","")
    
    @staticmethod
    def init_ssl():
        """SSL 인증서 설정"""
        try:
            import certifi
            os.environ['REQUESTS_CA_BUNDLE'] = certifi.where()
            os.environ['SSL_CERT_FILE'] = certifi.where()
            return True
        except ImportError:
            import urllib3
            urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
            return False


VERIFY_SSL = Config.init_ssl()
ctk.set_default_color_theme(Config.THEME_COLOR)
ctk.set_appearance_mode("system")


# ==================== 유틸리티 ====================

class PathUtils:
    """파일 경로 관련 유틸리티"""
    
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


class FileUtils:
    """파일 처리 관련 유틸리티"""
    
    @staticmethod
    def calculate_hash(file_path: Path) -> Optional[str]:
        """파일 SHA256 해시 계산"""
        sha256_hash = hashlib.sha256()
        try:
            with open(file_path, "rb") as f:
                for byte_block in iter(lambda: f.read(65536), b""):
                    sha256_hash.update(byte_block)
            return sha256_hash.hexdigest()
        except Exception as e:
            print(f"해시 계산 실패: {e}")
            return None


class IconManager:
    """윈도우 아이콘 관리"""
    
    @staticmethod
    def set_window_icon(window, icon_path: str) -> bool:
        """Win32 API로 윈도우 아이콘 설정"""
        if not icon_path or not os.path.exists(icon_path):
            return False
        
        try:
            window.update_idletasks()
            hwnd = window.winfo_id()
            if not hwnd:
                return False
            
            # 아이콘 로드
            hicon_small = ctypes.windll.user32.LoadImageW(
                0, icon_path, 1, 16, 16, 0x00000010
            )
            hicon_big = ctypes.windll.user32.LoadImageW(
                0, icon_path, 1, 32, 32, 0x00000010
            )
            
            if hicon_small == 0 and hicon_big == 0:
                return False
            
            # 아이콘 설정
            WM_SETICON = 0x80
            if hicon_small:
                ctypes.windll.user32.SendMessageW(hwnd, WM_SETICON, 0, hicon_small)
            if hicon_big:
                ctypes.windll.user32.SendMessageW(hwnd, WM_SETICON, 1, hicon_big)
            
            return True
        except Exception as e:
            print(f"아이콘 설정 실패: {e}")
            return False


# ==================== 매니페스트 관리 ====================

class ManifestManager:
    """매니페스트 파일 관리"""
    
    @staticmethod
    def get_remote(repo_owner: str, repo_name: str, branch: str, 
                   log_callback: Optional[Callable] = None) -> Optional[Dict]:
        """원격 manifest.json 가져오기"""
        url = f"https://raw.githubusercontent.com/{repo_owner}/{repo_name}/{branch}/release/manifest.json"
        try:
            if log_callback:
                log_callback("  → 원격 manifest 요청 중...", False)
            
            resp = requests.get(url, timeout=10, verify=VERIFY_SSL)
            resp.raise_for_status()
            
            if log_callback:
                log_callback("  ✓ 원격 manifest 로드 완료", False)
            
            return resp.json()
        except Exception as e:
            if log_callback:
                log_callback(f"  ✗ 원격 manifest 로드 실패: {e}", False)
            print(f"원격 manifest 불러오기 실패: {e}")
            return None
    
    @staticmethod
    def get_local(app_dir: Path) -> Dict:
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
    
    @staticmethod
    def save_local(app_dir: Path, manifest: Dict) -> None:
        """로컬 manifest.json 저장"""
        manifest_path = app_dir / "manifest.json"
        try:
            with manifest_path.open("w", encoding="utf-8") as f:
                json.dump(manifest, f, indent=2, ensure_ascii=False)
        except Exception as e:
            print(f"manifest 저장 실패: {e}")
    
    @staticmethod
    def parse_files(manifest: Dict) -> Dict[str, str]:
        """manifest에서 파일 정보 파싱 (경로: 해시)"""
        files = {}
        for item in manifest.get('files', []):
            if isinstance(item, dict) and 'relative_path' in item and 'hash' in item:
                files[item['relative_path']] = item['hash']
        return files


# ==================== 파일 다운로드 ====================

class GitHubDownloader:
    """GitHub 파일 다운로드"""
    
    @staticmethod
    def download_file(repo_owner: str, repo_name: str, file_path: str, 
                     dest_path: Path, branch: str, verbose: bool = False) -> bool:
        """단일 파일 다운로드"""
        timestamp = int(time.time() * 1000)
        url = f"https://raw.githubusercontent.com/{repo_owner}/{repo_name}/{branch}/release/{file_path}?nocache={timestamp}"
        
        if verbose:
            print(f"\n{'='*60}")
            print(f"다운로드 시도: {file_path}")
            print(f"URL: {url}")
        
        for attempt in range(Config.MAX_RETRIES):
            try:
                headers = {
                    'Cache-Control': 'no-cache, no-store, must-revalidate',
                    'Pragma': 'no-cache',
                    'Expires': '0'
                }
                
                # GitHub Token 추가
                if Config.GITHUB_TOKEN:
                    headers['Authorization'] = f'token {Config.GITHUB_TOKEN}'
                    if verbose:
                        print(f"  GitHub Token 사용 중")
                
                if verbose:
                    print(f"  시도 {attempt + 1}/{Config.MAX_RETRIES}")
                
                with requests.get(url, timeout=Config.DOWNLOAD_TIMEOUT, 
                                verify=VERIFY_SSL, stream=True, headers=headers) as resp:
                    
                    if verbose:
                        print(f"  응답 코드: {resp.status_code}")
                        
                        # Rate Limit 정보
                        if 'X-RateLimit-Limit' in resp.headers:
                            print(f"  Rate Limit:")
                            print(f"    - 총 제한: {resp.headers.get('X-RateLimit-Limit')}")
                            print(f"    - 남은 횟수: {resp.headers.get('X-RateLimit-Remaining')}")
                            reset_time = int(resp.headers.get('X-RateLimit-Reset', 0))
                            if reset_time:
                                reset_str = time.strftime('%H:%M:%S', time.localtime(reset_time))
                                print(f"    - 리셋 시간: {reset_str}")
                    
                    # Rate Limit 처리
                    if resp.status_code == 403:
                        remaining = resp.headers.get('X-RateLimit-Remaining', '?')
                        if remaining == '0':
                            reset_time = int(resp.headers.get('X-RateLimit-Reset', 0))
                            wait_seconds = max(reset_time - time.time(), 0) + 10
                            
                            if wait_seconds < 300:  # 5분 이내면 대기
                                print(f"  Rate Limit 도달. {wait_seconds:.0f}초 대기 중...")
                                time.sleep(wait_seconds)
                                continue
                            else:
                                print(f"  Rate Limit 대기 시간이 너무 김 ({wait_seconds:.0f}초)")
                                return False
                    
                    # 404 오류
                    if resp.status_code == 404:
                        print(f"  ✗ 파일을 찾을 수 없음 (404): {file_path}")
                        print(f"  GitHub에서 파일 존재 여부 확인 필요")
                        return False
                    
                    resp.raise_for_status()
                    
                    # 파일 정보
                    content_length = resp.headers.get('Content-Length')
                    expected_size = int(content_length) if content_length else None
                    
                    if verbose and expected_size:
                        print(f"  파일 크기: {expected_size:,} bytes ({expected_size/1024:.1f} KB)")
                    
                    dest_path.parent.mkdir(parents=True, exist_ok=True)
                    
                    if dest_path.exists() and not os.access(dest_path, os.W_OK):
                        print(f"  ✗ 쓰기 권한 없음: {dest_path}")
                        return False
                    
                    # 임시 파일로 다운로드
                    temp_path = dest_path.with_suffix(dest_path.suffix + '.tmp')
                    actual_size = 0
                    start_time = time.time()
                    
                    with open(temp_path, "wb") as f:
                        for chunk in resp.iter_content(chunk_size=65536):
                            if chunk:
                                f.write(chunk)
                                actual_size += len(chunk)
                    
                    elapsed_time = time.time() - start_time
                    
                    if verbose:
                        speed = actual_size / elapsed_time / 1024 if elapsed_time > 0 else 0
                        print(f"  다운로드 완료: {actual_size:,} bytes ({elapsed_time:.2f}초, {speed:.1f} KB/s)")
                    
                    # 크기 검증
                    if expected_size and actual_size != expected_size:
                        print(f"  ✗ 크기 불일치: {file_path}")
                        print(f"    예상: {expected_size:,} bytes")
                        print(f"    실제: {actual_size:,} bytes")
                        temp_path.unlink()
                        if attempt < Config.MAX_RETRIES - 1:
                            time.sleep(Config.RETRY_DELAY)
                            continue
                        return False
                    
                    # 최종 파일로 이동
                    if dest_path.exists():
                        dest_path.unlink()
                    temp_path.rename(dest_path)
                    
                    if not dest_path.exists():
                        print(f"  ✗ 파일 이동 실패: {file_path}")
                        return False
                    
                    if verbose:
                        print(f"  ✓ 성공")
                        print(f"{'='*60}\n")
                    
                    return True
                    
            except requests.exceptions.Timeout as e:
                print(f"  ✗ 타임아웃 ({attempt + 1}/{Config.MAX_RETRIES}): {file_path}")
                if verbose:
                    print(f"     {e}")
                if attempt < Config.MAX_RETRIES - 1:
                    time.sleep(Config.RETRY_DELAY)
                continue
                
            except requests.exceptions.ConnectionError as e:
                print(f"  ✗ 연결 오류 ({attempt + 1}/{Config.MAX_RETRIES}): {file_path}")
                if verbose:
                    print(f"     {e}")
                if attempt < Config.MAX_RETRIES - 1:
                    time.sleep(Config.RETRY_DELAY)
                continue
                
            except requests.exceptions.HTTPError as e:
                status_code = e.response.status_code if e.response else 'N/A'
                print(f"  ✗ HTTP 오류 ({status_code}): {file_path}")
                if verbose:
                    print(f"     {e}")
                if attempt < Config.MAX_RETRIES - 1:
                    time.sleep(Config.RETRY_DELAY)
                continue
                
            except PermissionError as e:
                print(f"  ✗ 권한 오류: {file_path}")
                print(f"     관리자 권한으로 실행 필요")
                return False
                
            except Exception as e:
                print(f"  ✗ 예상치 못한 오류 ({attempt + 1}/{Config.MAX_RETRIES}): {file_path}")
                print(f"     {type(e).__name__}: {e}")
                if verbose:
                    import traceback
                    traceback.print_exc()
                if attempt < Config.MAX_RETRIES - 1:
                    time.sleep(Config.RETRY_DELAY)
                continue
        
        print(f"  ✗ 모든 재시도 실패: {file_path}")
        if verbose:
            print(f"{'='*60}\n")
        return False
    
    @staticmethod
    def download_file_with_verification(repo_owner: str, repo_name: str, file_path: str, 
                                       dest_path: Path, branch: str, expected_hash: str) -> bool:
        """해시 검증을 포함한 다운로드"""
        for attempt in range(Config.MAX_RETRIES):
            # 다운로드 시도 (상세 로그는 마지막 시도에만)
            verbose = (attempt == Config.MAX_RETRIES - 1)
            
            if GitHubDownloader.download_file(repo_owner, repo_name, file_path, dest_path, branch, verbose):
                # 해시 검증
                downloaded_hash = FileUtils.calculate_hash(dest_path)
                
                if downloaded_hash == expected_hash:
                    return True  # 성공
                else:
                    print(f"  ✗ 해시 불일치 ({attempt + 1}/{Config.MAX_RETRIES}): {file_path}")
                    print(f"     예상: {expected_hash}")
                    print(f"     실제: {downloaded_hash}")
                    
                    # 파일 삭제 후 재시도
                    if dest_path.exists():
                        dest_path.unlink()
                    
                    if attempt < Config.MAX_RETRIES - 1:
                        time.sleep(Config.RETRY_DELAY * 2)
                        continue
            else:
                if attempt < Config.MAX_RETRIES - 1:
                    time.sleep(Config.RETRY_DELAY)
                    continue
        
        return False
    
    @staticmethod
    def _download_single_file(args: Tuple) -> Tuple[int, str, str, bool, Optional[str]]:
        """단일 파일 다운로드 (병렬 처리용)"""
        index, total, file_path, expected_hash, repo_owner, repo_name, branch, app_dir = args
        
        dest_path = app_dir / file_path
        
        # 해시 검증을 포함한 다운로드
        success = GitHubDownloader.download_file_with_verification(
            repo_owner, repo_name, file_path, dest_path, branch, expected_hash
        )
        
        if success:
            downloaded_hash = FileUtils.calculate_hash(dest_path)
            hash_match = (downloaded_hash == expected_hash)
            return (index, file_path, expected_hash, hash_match, downloaded_hash)
        
        return (index, file_path, expected_hash, False, None)
    
    @staticmethod
    def download_all(repo_owner: str, repo_name: str, manifest: Dict, 
                    app_dir: Path, branch: str, 
                    log_callback: Optional[Callable] = None,
                    max_workers: int = None) -> int:
        """모든 파일 다운로드 (병렬 처리)"""
        if max_workers is None:
            max_workers = Config.MAX_CONCURRENT_DOWNLOADS
        
        files = ManifestManager.parse_files(manifest)
        
        if not files:
            if log_callback:
                log_callback("  ! 다운로드할 파일이 없습니다", False)
            return 0
        
        if log_callback:
            log_callback(f"→ 파일 다운로드 시작 ({len(files)}개)", False)
        
        success_count = 0
        fail_count = 0
        
        # 병렬 다운로드 준비
        file_items = list(files.items())
        total_files = len(file_items)
        
        download_args = [
            (i + 1, total_files, file_path, expected_hash, repo_owner, repo_name, branch, app_dir)
            for i, (file_path, expected_hash) in enumerate(file_items)
        ]
        
        # 병렬 다운로드 실행
        with concurrent.futures.ThreadPoolExecutor(max_workers=max_workers) as executor:
            futures = {executor.submit(GitHubDownloader._download_single_file, args): args 
                      for args in download_args}
            
            for future in concurrent.futures.as_completed(futures):
                try:
                    index, file_path, expected_hash, hash_match, downloaded_hash = future.result()
                    
                    if hash_match:
                        success_count += 1
                        if log_callback:
                            log_callback(f"  ✓ [{index}/{total_files}] {file_path}", False)
                    elif downloaded_hash is not None:
                        fail_count += 1
                        if log_callback:
                            log_callback(f"  ✗ [{index}/{total_files}] {file_path} (해시 불일치)", False)
                    else:
                        fail_count += 1
                        if log_callback:
                            log_callback(f"  ✗ [{index}/{total_files}] {file_path} (다운로드 실패)", False)
                
                except Exception as e:
                    fail_count += 1
                    if log_callback:
                        log_callback(f"  ✗ 다운로드 오류: {e}", False)
        
        if log_callback:
            if fail_count > 0:
                log_callback(f"  ✓ 다운로드 완료: 성공 {success_count}개, 실패 {fail_count}개", False)
            else:
                log_callback(f"  ✓ 다운로드 완료: {success_count}개", False)
        
        return success_count


# ==================== 파일 정리 ====================

class FileCleanup:
    """파일 삭제 및 정리"""
    
    @staticmethod
    def delete_all_except_protected(app_dir: Path, 
                                    log_callback: Optional[Callable] = None) -> int:
        """launcher와 poppler를 제외한 모든 파일 삭제"""
        if log_callback:
            log_callback("→ 기존 파일 삭제 중...", False)
        
        deleted_count = 0
        poppler_dir = app_dir / Config.POPPLER_FOLDER_NAME
        launcher_exe = Path(sys.executable) if getattr(sys, 'frozen', False) else None
        
        for item in app_dir.iterdir():
            # 보호 대상 확인
            if item == poppler_dir:
                if log_callback:
                    log_callback(f"  • 보호됨: {item.name}/", False)
                continue
            
            if launcher_exe and item == launcher_exe:
                continue
            
            if "launcher" in item.name.lower():
                if log_callback:
                    log_callback(f"  • 보호됨: {item.name}", False)
                continue
            
            # 삭제 시도
            try:
                if item.is_file():
                    item.unlink()
                    deleted_count += 1
                    if log_callback:
                        log_callback(f"  ✓ 삭제: {item.name}", False)
                elif item.is_dir():
                    shutil.rmtree(item)
                    deleted_count += 1
                    if log_callback:
                        log_callback(f"  ✓ 삭제: {item.name}/", False)
            except Exception as e:
                if log_callback:
                    log_callback(f"  ✗ 삭제 실패({item.name}): {e}", False)
        
        if log_callback:
            log_callback(f"  ✓ 총 {deleted_count}개 항목 삭제 완료", False)
        
        return deleted_count


# ==================== 무결성 검사 ====================

class IntegrityChecker:
    """파일 무결성 검사"""
    
    @staticmethod
    def check(local_manifest: Dict, remote_manifest: Dict, app_dir: Path,
              log_callback: Optional[Callable] = None) -> Dict[str, List]:
        """무결성 검사 수행"""
        if log_callback:
            log_callback("→ 파일 무결성 검사 시작...", False)
        
        local_files = ManifestManager.parse_files(local_manifest)
        remote_files = ManifestManager.parse_files(remote_manifest)
        
        if not remote_files:
            if log_callback:
                log_callback("  ✗ 원격 manifest 구조 오류", False)
            return {'to_download': [], 'to_delete': []}
        
        files_to_download = []
        files_to_delete = []
        
        poppler_dir = app_dir / Config.POPPLER_FOLDER_NAME
        launcher_exe = Path(sys.executable) if getattr(sys, 'frozen', False) else None
        
        # 원격 파일 검사
        for path, remote_hash in remote_files.items():
            file_path = app_dir / path
            
            # 보호된 파일 제외
            if IntegrityChecker._is_protected_file(file_path, launcher_exe, poppler_dir):
                if log_callback:
                    log_callback(f"    • {path}: 보호됨 (스킵)", False)
                continue
            
            needs_download = False
            reason = ""
            
            if not file_path.exists():
                needs_download = True
                reason = '파일 없음'
            else:
                local_hash = local_files.get(path)
                if local_hash is None:
                    actual_hash = FileUtils.calculate_hash(file_path)
                    if actual_hash != remote_hash:
                        needs_download = True
                        reason = '매니페스트 누락 및 해시 불일치'
                elif local_hash != remote_hash:
                    needs_download = True
                    reason = '해시 불일치'
            
            if needs_download:
                files_to_download.append({'path': path, 'hash': remote_hash, 'reason': reason})
                if log_callback:
                    log_callback(f"    • {path}: {reason}", False)
        
        # 로컬 파일 검사 (보호 폴더 제외)
        for path in local_files.keys():
            if path not in remote_files:
                file_path = app_dir / path
                
                # 보호 대상 확인
                if IntegrityChecker._is_protected_file(file_path, launcher_exe, poppler_dir):
                    continue
                
                if file_path.exists():
                    files_to_delete.append({'path': path, 'reason': '원격에 없음'})
                    if log_callback:
                        log_callback(f"    • {path}: 삭제 대상 (원격에 없음)", False)
        
        if log_callback:
            if files_to_download or files_to_delete:
                log_callback(f"  ✓ 검사 완료: 다운로드 {len(files_to_download)}개, 삭제 {len(files_to_delete)}개", False)
            else:
                log_callback("  ✓ 검사 완료: 모든 파일 정상", False)
        
        return {'to_download': files_to_download, 'to_delete': files_to_delete}
    
    @staticmethod
    def _is_protected_file(file_path: Path, launcher_exe: Optional[Path], poppler_dir: Path) -> bool:
        """보호 대상 파일인지 확인"""
        # 1. Poppler 폴더 내부
        try:
            if file_path.is_relative_to(poppler_dir):
                return True
        except:
            pass
        
        # 2. 런처 실행 파일 자신
        if launcher_exe and file_path == launcher_exe:
            return True
        
        # 3. 런처 관련 파일 (이름에 "launcher" 포함)
        if "launcher" in file_path.name.lower():
            return True
        
        # 4. 현재 실행 중인 파일
        if getattr(sys, 'frozen', False):
            current_exe = Path(sys.executable)
            if file_path == current_exe:
                return True
        
        return False


# ==================== 업데이트 관리자 ====================

class UpdateManager:
    """애플리케이션 업데이트 관리"""
    
    @staticmethod
    def update(app_dir: Path, log_callback: Optional[Callable] = None) -> bool:
        """메인 앱 업데이트 수행"""
        try:
            if log_callback:
                log_callback("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━", False)
                log_callback("[ 애플리케이션 업데이트 확인 ]", False)
                log_callback("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━", False)
                log_callback("→ 업데이트 확인 시작", False)
            
            # 원격 매니페스트 가져오기
            remote_manifest = ManifestManager.get_remote(
                Config.APP_REPO_OWNER, Config.APP_REPO_NAME, Config.APP_BRANCH, log_callback
            )
            
            if not remote_manifest or not isinstance(remote_manifest, dict) or 'version' not in remote_manifest:
                if log_callback:
                    log_callback("  ✗ 업데이트 확인 실패", False)
                return False
            
            # 로컬 매니페스트와 비교
            local_manifest = ManifestManager.get_local(app_dir)
            local_version = local_manifest.get('version', '0.0.0')
            remote_version = remote_manifest.get('version', '0.0.0')
            
            if log_callback:
                log_callback(f"  • 현재 버전: v{local_version}", False)
                log_callback(f"  • 최신 버전: v{remote_version}", False)
            
            # 버전 비교
            if local_version != remote_version:
                return UpdateManager._full_update(app_dir, remote_manifest, local_version, remote_version, log_callback)
            else:
                return UpdateManager._integrity_update(app_dir, local_manifest, remote_manifest, log_callback)
        
        except Exception as e:
            if log_callback:
                log_callback(f"✗ 업데이트 중 오류 발생: {e}", False)
            print(f"업데이트 실패: {e}")
            return False
    
    @staticmethod
    def _full_update(app_dir: Path, remote_manifest: Dict, local_version: str, 
                    remote_version: str, log_callback: Optional[Callable]) -> bool:
        """전체 업데이트"""
        if log_callback:
            log_callback("", False)
            log_callback("! 새 버전 발견 - 전체 업데이트 시작", False)
        
        FileCleanup.delete_all_except_protected(app_dir, log_callback)
        
        if log_callback:
            log_callback("", False)
        
        success_count = GitHubDownloader.download_all(
            Config.APP_REPO_OWNER, Config.APP_REPO_NAME, remote_manifest, 
            app_dir, Config.APP_BRANCH, log_callback
        )
        
        if success_count > 0:
            ManifestManager.save_local(app_dir, remote_manifest)
            if log_callback:
                log_callback("", False)
                log_callback(f"✓ 업데이트 완료 (v{local_version} → v{remote_version})", False)
            return True
        else:
            if log_callback:
                log_callback("", False)
                log_callback("✗ 업데이트 실패 (다운로드 실패)", False)
            return False
    
    @staticmethod
    def _integrity_update(app_dir: Path, local_manifest: Dict, remote_manifest: Dict,
                         log_callback: Optional[Callable]) -> bool:
        """무결성 검사 및 복구"""
        if log_callback:
            log_callback("", False)
            log_callback("✓ 버전 일치 - 무결성 검사 시작", False)
        
        check_result = IntegrityChecker.check(local_manifest, remote_manifest, app_dir, log_callback)
        files_to_download = check_result['to_download']
        files_to_delete = check_result['to_delete']
        
        if not files_to_download and not files_to_delete:
            if log_callback:
                log_callback("", False)
                log_callback("✓ 모든 파일이 정상입니다", False)
            return False
        
        # 삭제 처리
        if files_to_delete:
            UpdateManager._delete_files(app_dir, files_to_delete, log_callback)
        
        # 다운로드 처리
        if files_to_download:
            UpdateManager._repair_files(app_dir, files_to_download, log_callback)
        
        # manifest 업데이트
        if files_to_download or files_to_delete:
            ManifestManager.save_local(app_dir, remote_manifest)
            if log_callback:
                log_callback("", False)
                log_callback("✓ 무결성 검사 및 복구 완료", False)
            return True
        
        return False
    
    @staticmethod
    def _delete_files(app_dir: Path, files_to_delete: List[Dict], 
                     log_callback: Optional[Callable]) -> None:
        """파일 삭제"""
        if log_callback:
            log_callback("", False)
            log_callback(f"→ 불필요한 파일 삭제 시작 ({len(files_to_delete)}개)", False)
        
        deleted_count = 0
        for file_info in files_to_delete:
            file_path = app_dir / file_info['path']
            try:
                if file_path.exists():
                    if file_path.is_file():
                        file_path.unlink()
                    elif file_path.is_dir():
                        shutil.rmtree(file_path)
                    deleted_count += 1
                    if log_callback:
                        log_callback(f"  ✓ 삭제: {file_info['path']}", False)
            except Exception as e:
                if log_callback:
                    log_callback(f"  ✗ 삭제 실패({file_info['path']}): {e}", False)
        
        if log_callback:
            log_callback(f"  ✓ 삭제 완료: {deleted_count}개", False)
    
    @staticmethod
    def _repair_files(app_dir: Path, files_to_download: List[Dict],
                     log_callback: Optional[Callable]) -> None:
        """파일 복구 (병렬 처리 + 실패 파일 재시도)"""
        if log_callback:
            log_callback("", False)
            log_callback(f"→ 손상된 파일 복구 시작 ({len(files_to_download)}개)", False)
        
        success_count = 0
        fail_count = 0
        total_files = len(files_to_download)
        failed_files = []
        
        # 병렬 다운로드 준비
        download_args = [
            (i + 1, total_files, file_info['path'], file_info['hash'], 
             Config.APP_REPO_OWNER, Config.APP_REPO_NAME, Config.APP_BRANCH, app_dir)
            for i, file_info in enumerate(files_to_download)
        ]
        
        # 1차 시도: 병렬 다운로드
        with concurrent.futures.ThreadPoolExecutor(max_workers=Config.MAX_CONCURRENT_DOWNLOADS) as executor:
            futures = {executor.submit(GitHubDownloader._download_single_file, args): args 
                      for args in download_args}
            
            for future in concurrent.futures.as_completed(futures):
                try:
                    index, file_path, expected_hash, hash_match, downloaded_hash = future.result()
                    
                    if hash_match:
                        success_count += 1
                        if log_callback:
                            log_callback(f"  ✓ [{index}/{total_files}] {file_path}", False)
                    else:
                        fail_count += 1
                        failed_files.append((index, file_path, expected_hash))
                        if log_callback:
                            if downloaded_hash is not None:
                                log_callback(f"  ✗ [{index}/{total_files}] {file_path} (해시 불일치)", False)
                            else:
                                log_callback(f"  ✗ [{index}/{total_files}] {file_path} (다운로드 실패)", False)
                
                except Exception as e:
                    fail_count += 1
                    if log_callback:
                        log_callback(f"  ✗ 다운로드 오류: {e}", False)
        
        # 2차 시도: 실패한 파일 순차적으로 재시도
        if failed_files:
            if log_callback:
                log_callback("", False)
                log_callback(f"→ 실패한 파일 재시도 ({len(failed_files)}개)", False)
            
            for index, file_path, expected_hash in failed_files:
                if log_callback:
                    log_callback(f"  → 재시도 중: {file_path}", False)
                
                dest_path = app_dir / file_path
                
                # 순차적 재시도
                retry_success = GitHubDownloader.download_file_with_verification(
                    Config.APP_REPO_OWNER, Config.APP_REPO_NAME, 
                    file_path, dest_path, Config.APP_BRANCH, expected_hash
                )
                
                if retry_success:
                    success_count += 1
                    fail_count -= 1
                    if log_callback:
                        log_callback(f"  ✓ 재시도 성공: {file_path}", False)
                else:
                    if log_callback:
                        log_callback(f"  ✗ 재시도 실패: {file_path}", False)
        
        if log_callback:
            if fail_count > 0:
                log_callback(f"  ✓ 복구 완료: 성공 {success_count}개, 실패 {fail_count}개", False)
            else:
                log_callback(f"  ✓ 복구 완료: {success_count}개", False)


# ==================== Poppler 관리자 ====================

class PopplerManager:
    """Poppler 설치 및 관리"""
    
    @staticmethod
    def get_installed_version(poppler_dir: Path) -> Optional[str]:
        """설치된 Poppler 버전 확인"""
        if not poppler_dir.exists():
            return None
        
        for item in poppler_dir.iterdir():
            if item.is_dir() and "poppler" in item.name.lower():
                match = re.search(r'(\d+\.\d+\.\d+)', item.name)
                if match:
                    return match.group(1)
        
        return None
    
    @staticmethod
    def get_latest_version() -> tuple:
        """최신 Poppler 버전 정보 가져오기"""
        api_url = f"https://api.github.com/repos/{Config.POPPLER_REPO_OWNER}/{Config.POPPLER_REPO_NAME}/releases/latest"
        try:
            resp = requests.get(api_url, timeout=10, verify=VERIFY_SSL)
            resp.raise_for_status()
            data = resp.json()
            
            for asset in data.get("assets", []):
                if asset["name"].lower().endswith(".zip"):
                    filename = asset["name"]
                    match = re.search(r'(\d+\.\d+\.\d+)', filename)
                    version = match.group(1) if match else "unknown"
                    return asset["browser_download_url"], filename, version
            
            return None, None, None
        except Exception:
            return None, None, None
    
    @staticmethod
    def download_and_extract(dest_folder: Path, 
                           log_callback: Optional[Callable] = None) -> Path:
        """Poppler 다운로드 및 압축 해제"""
        download_url, filename, version = PopplerManager.get_latest_version()
        if not download_url:
            raise RuntimeError("Poppler 윈도우용 최신 zip 파일을 찾을 수 없습니다.")
        
        zip_path = dest_folder / filename
        
        if log_callback:
            log_callback(f"→ 다운로드 파일: {filename}", False)
            log_callback(f"→ 다운로드 시작...", False)
        
        # 다운로드
        try:
            with requests.get(download_url, stream=True, timeout=60, verify=VERIFY_SSL) as r:
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
                                    log_callback(f"  → 다운로드 진행: {percent}%", True)
                                last_percent = percent
            
            if log_callback:
                log_callback(f"  ✓ 다운로드 완료", False)
        except Exception as e:
            if log_callback:
                log_callback(f"  ✗ 다운로드 실패: {e}", False)
            raise
        
        # 압축 해제
        if log_callback:
            log_callback("→ 압축 해제 시작", False)
        
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
                            log_callback(f"  → 압축 해제 진행: {percent}%", True)
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
                    log_callback("  ✓ 임시 파일 삭제 완료", False)
        
        # 폴더 찾기
        if log_callback:
            log_callback("→ Poppler 폴더 확인 중...", False)
        
        for item in dest_folder.iterdir():
            if item.is_dir() and "poppler" in item.name.lower():
                bin_path = item / "Library" / "bin"
                if bin_path.exists() and (bin_path / "pdftoppm.exe").exists():
                    if log_callback:
                        log_callback(f"  ✓ Poppler 설치 완료 (v{version})", False)
                    return bin_path
        
        raise RuntimeError("Poppler 폴더를 찾을 수 없습니다.")
    
    @staticmethod
    def check_and_update(app_dir: Path, log_callback: Optional[Callable] = None) -> None:
        """Poppler 확인 및 업데이트"""
        if log_callback:
            log_callback("", False)
            log_callback("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━", False)
            log_callback("[ Poppler 유효성 검사 ]", False)
            log_callback("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━", False)
            time.sleep(0.2)
        
        poppler_dir = app_dir / Config.POPPLER_FOLDER_NAME
        poppler_dir.mkdir(parents=True, exist_ok=True)
        
        # 버전 확인
        installed_version = PopplerManager.get_installed_version(poppler_dir)
        
        if log_callback:
            log_callback("→ 최신 버전 확인 중...", False)
        
        _, _, latest_version = PopplerManager.get_latest_version()
        
        if not latest_version:
            if log_callback:
                log_callback("  ✗ 최신 버전 확인 실패", False)
            if installed_version:
                if log_callback:
                    log_callback("", False)
                    log_callback(f"✓ 기존 Poppler 사용 (v{installed_version})", False)
            else:
                if log_callback:
                    log_callback("", False)
                    log_callback("✗ Poppler 설치 불가 (네트워크 오류)", False)
                raise RuntimeError("Poppler 최신 버전 확인 실패")
            return
        
        if log_callback:
            log_callback(f"  • 최신 버전: v{latest_version}", False)
        
        # 업데이트 필요 여부 확인
        if installed_version:
            if log_callback:
                log_callback(f"  • 설치된 버전: v{installed_version}", False)
            
            if installed_version == latest_version:
                if log_callback:
                    log_callback("", False)
                    log_callback("✓ Poppler가 최신 버전입니다", False)
                return
            else:
                if log_callback:
                    log_callback("", False)
                    log_callback("! 새 버전 발견 - Poppler 업데이트 시작", False)
                    log_callback("→ 기존 Poppler 삭제 중...", False)
                
                # 기존 삭제
                for item in poppler_dir.iterdir():
                    if item.is_dir() and "poppler" in item.name.lower():
                        try:
                            shutil.rmtree(item)
                            if log_callback:
                                log_callback(f"  ✓ 삭제 완료: {item.name}", False)
                        except Exception as e:
                            if log_callback:
                                log_callback(f"  ✗ 삭제 실패: {e}", False)
                
                if log_callback:
                    log_callback("", False)
        else:
            if log_callback:
                log_callback("", False)
                log_callback("! Poppler가 설치되지 않았습니다", False)
                log_callback("", False)
                time.sleep(0.2)
        
        # 설치
        try:
            PopplerManager.download_and_extract(poppler_dir, log_callback)
            if log_callback:
                log_callback("", False)
                log_callback("✓ Poppler 설치 완료", False)
        except Exception as e:
            if log_callback:
                log_callback(f"✗ Poppler 설치 실패: {e}", False)
            raise


# ==================== 런처 UI ====================

class LauncherApp(ctk.CTk):
    """런처 메인 애플리케이션"""

    def __init__(self):
        super().__init__()
        self._setup_window()
        self._setup_icon()
        self._setup_variables()
        self._setup_ui()
        self._start_initialization()

    def _setup_window(self) -> None:
        """윈도우 설정"""
        self.title("파일 무결성 검사 중...")
        self.geometry("700x400")
        self.resizable(False, False)
        
        # 화면 중앙 배치
        self.update_idletasks()
        x = (self.winfo_screenwidth() - 700) // 2
        y = (self.winfo_screenheight() - 400) // 2
        self.geometry(f"700x400+{x}+{y}")

    def _setup_icon(self) -> None:
        """아이콘 설정"""
        icon_path = PathUtils.get_icon_path()
        if icon_path:
            try:
                self.iconbitmap(icon_path)
            except:
                pass
            self.after(50, lambda: IconManager.set_window_icon(self, icon_path))

    def _setup_variables(self) -> None:
        """변수 초기화"""
        self.app_dir = PathUtils.get_app_directory()
        self.result = {"should_close": False, "launch_main": True}
        self._closing = False
        self.log_queue = queue.Queue()
        self.logs = []

    def _setup_ui(self) -> None:
        """UI 구성"""
        self.log_box = ctk.CTkTextbox(
            self, height=370, width=650, font=("Consolas", 10)
        )
        self.log_box.pack(pady=10, padx=25)
        self.log_box.configure(state="disabled")
        
        self.process_log_queue()

    def _start_initialization(self) -> None:
        """초기화 작업 시작"""
        self.after(100, lambda: threading.Thread(
            target=self.init_thread, daemon=True
        ).start())
        self.after(500, self.check_close)

    def add_log(self, message: str, is_progress: bool = False) -> None:
        """로그 추가"""
        try:
            self.log_box.configure(state="normal")
            if is_progress and self.logs:
                last_line_index = len(self.logs)
                self.log_box.delete(f"{last_line_index}.0", f"{last_line_index + 1}.0")
                self.logs.pop()
            
            self.logs.append(message)
            self.log_box.insert("end", message + "\n")
            self.log_box.see("end")
            self.log_box.configure(state="disabled")
            self.update()
        except Exception as e:
            print(f"로그 추가 실패: {e}")

    def process_log_queue(self) -> None:
        """로그 큐 처리"""
        if self._closing:
            return
        
        try:
            while not self.log_queue.empty():
                message, is_progress = self.log_queue.get_nowait()
                self.add_log(message, is_progress)
        except queue.Empty:
            pass
        except Exception as e:
            print(f"큐 처리 실패: {e}")
        
        try:
            self.after(100, self.process_log_queue)
        except Exception:
            pass

    def safe_add_log(self, message: str, is_progress: bool = False) -> None:
        """스레드 안전 로그 추가"""
        try:
            self.log_queue.put((message, is_progress))
        except Exception as e:
            print(f"safe_add_log 실패: {e}")

    def init_thread(self) -> None:
        """초기화 작업 스레드"""
        try:
            log_cb = lambda msg, is_progress: (
                self.safe_add_log(msg, is_progress),
                time.sleep(0.05)
            )
            
            # 업데이트 확인
            UpdateManager.update(self.app_dir, log_cb)
            time.sleep(0.3)
            
            # Poppler 확인
            PopplerManager.check_and_update(self.app_dir, log_cb)
            time.sleep(0.3)
            
            # 카운트다운
            self.safe_add_log("", False)
            self.safe_add_log("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━", False)
            for i in range(3, 0, -1):
                self.safe_add_log(f"✓ {i}초 후 프로그램이 시작됩니다...", False)
                time.sleep(1.0)
            self.safe_add_log("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━", False)
            
            self.result["should_close"] = True
        
        except Exception as e:
            print(f"초기화 오류: {e}")
            import traceback
            traceback.print_exc()
            try:
                self.safe_add_log(f"✗ 초기화 오류: {e}", False)
            except:
                pass
            self.result["should_close"] = True
            self.result["launch_main"] = False

    def check_close(self) -> None:
        """종료 확인"""
        if self._closing:
            return
        
        if self.result["should_close"]:
            self._closing = True
            self.after(500, self.launch_main_app)
        else:
            try:
                self.after(100, self.check_close)
            except Exception:
                pass

    def launch_main_app(self) -> None:
        """메인 앱 실행"""
        if not self.result["launch_main"]:
            self.quit_app()
            return
        
        main_exe = self.app_dir / Config.MAIN_EXE_NAME
        
        if not main_exe.exists():
            print(f"메인 앱을 찾을 수 없습니다: {main_exe}")
            try:
                messagebox.showerror("오류", f"메인 앱을 찾을 수 없습니다:\n{main_exe}")
            except:
                pass
            self.quit_app()
            return
        
        try:
            print(f"메인 앱 실행 시도: {main_exe}")
            subprocess.Popen(
                [str(main_exe)],
                cwd=str(self.app_dir),
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE
            )
            print(f"메인 앱 실행 성공")
            time.sleep(1)
        except Exception as e:
            print(f"메인 앱 실행 실패: {e}")
            import traceback
            traceback.print_exc()
            try:
                messagebox.showerror("오류", f"메인 앱 실행 실패:\n{e}")
            except:
                pass
        
        self.quit_app()

    def quit_app(self) -> None:
        """애플리케이션 종료"""
        try:
            self.quit()
        except:
            pass
        
        try:
            self.destroy()
        except:
            pass


# ==================== 메인 진입점 ====================

def main():
    """메인 함수"""
    try:
        app = LauncherApp()
        app.mainloop()
    except Exception as e:
        print(f"Launcher 오류: {e}")
        import traceback
        traceback.print_exc()
    finally:
        sys.exit(0)


if __name__ == "__main__":
    main()
