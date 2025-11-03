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
from pathlib import Path

import customtkinter as ctk
from tkinter import messagebox

# SSL 인증서 경로 설정
try:
    import certifi
    os.environ['REQUESTS_CA_BUNDLE'] = certifi.where()
    os.environ['SSL_CERT_FILE'] = certifi.where()
except ImportError:
    import urllib3
    urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)


# ---------- 상수 및 설정 ----------

THEME_COLOR = "blue"
APP_REPO_OWNER = "c-closed"
APP_REPO_NAME = "P2J"
APP_BRANCH = "main"
MAIN_EXE_NAME = "P2J.exe"
ICON_FILENAME = "icon.ico"
POPPLER_FOLDER_NAME = "poppler"
POPPLER_REPO_OWNER = "oschwartz10612"
POPPLER_REPO_NAME = "poppler-windows"

VERIFY_SSL = True
try:
    import certifi
except ImportError:
    VERIFY_SSL = False

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
            return False
        WM_SETICON = 0x80
        ctypes.windll.user32.SendMessageW(hwnd, WM_SETICON, 0, hicon)
        ctypes.windll.user32.SendMessageW(hwnd, WM_SETICON, 1, hicon)
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
        print(f"해시 계산 실패: {e}")
        return None


# ---------- 업데이트 관련 함수 ----------

def get_remote_manifest(repo_owner, repo_name, branch, log_callback=None):
    """원격 manifest.json 읽기"""
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


def get_local_manifest(app_dir: Path):
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
        print(f"manifest 저장 실패: {e}")


def delete_all_except_protected(app_dir: Path, log_callback=None):
    """launcher와 poppler 폴더를 제외한 모든 파일 삭제"""
    if log_callback:
        log_callback("→ 기존 파일 삭제 중...", False)
    
    deleted_count = 0
    poppler_dir = app_dir / POPPLER_FOLDER_NAME
    launcher_exe = Path(sys.executable) if getattr(sys, 'frozen', False) else None
    
    for item in app_dir.iterdir():
        # poppler 폴더 보호
        if item == poppler_dir:
            if log_callback:
                log_callback(f"  • 보호됨: {item.name}/", False)
            continue
        
        # launcher.exe 자신 보호
        if launcher_exe and item == launcher_exe:
            continue
        
        # launcher 관련 파일/폴더 보호
        if "launcher" in item.name.lower():
            if log_callback:
                log_callback(f"  • 보호됨: {item.name}", False)
            continue
        
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


def download_file_from_github(repo_owner, repo_name, file_path, dest_path: Path, branch):
    """GitHub에서 단일 파일 다운로드"""
    url = f"https://raw.githubusercontent.com/{repo_owner}/{repo_name}/{branch}/release/{file_path}"
    try:
        resp = requests.get(url, timeout=30, verify=VERIFY_SSL)
        resp.raise_for_status()
        dest_path.parent.mkdir(parents=True, exist_ok=True)
        with open(dest_path, "wb") as f:
            f.write(resp.content)
        return True
    except Exception as e:
        print(f"다운로드 실패({file_path}): {e}")
        return False


def download_all_files(repo_owner, repo_name, remote_manifest, app_dir: Path, log_callback, branch):
    """모든 파일 다운로드"""
    files = remote_manifest.get('files', [])
    if not files:
        return 0
    
    if log_callback:
        log_callback(f"→ 파일 다운로드 시작 ({len(files)}개)", False)
    
    success_count = 0
    fail_count = 0
    
    for i, file_info in enumerate(files, start=1):
        file_path = file_info['path']
        expected_hash = file_info['hash']
        dest_path = app_dir / file_path
        
        if log_callback:
            log_callback(f"  → [{i}/{len(files)}] {file_path} 다운로드 중...", False)
        
        if download_file_from_github(repo_owner, repo_name, file_path, dest_path, branch):
            downloaded_hash = calculate_file_hash(dest_path)
            if downloaded_hash == expected_hash:
                success_count += 1
                if log_callback:
                    log_callback(f"  ✓ [{i}/{len(files)}] {file_path} 완료", False)
            else:
                fail_count += 1
                if log_callback:
                    log_callback(f"  ✗ [{i}/{len(files)}] {file_path} 해시 불일치", False)
        else:
            fail_count += 1
            if log_callback:
                log_callback(f"  ✗ [{i}/{len(files)}] {file_path} 다운로드 실패", False)
    
    if log_callback:
        if fail_count > 0:
            log_callback(f"  ✓ 다운로드 완료: 성공 {success_count}개, 실패 {fail_count}개", False)
        else:
            log_callback(f"  ✓ 다운로드 완료: {success_count}개", False)
    
    return success_count


def integrity_check_and_fix(repo_owner, repo_name, local_manifest, remote_manifest, app_dir: Path, log_callback, branch):
    """무결성 검사 및 수정"""
    if log_callback:
        log_callback("→ 파일 무결성 검사 시작...", False)
    
    local_files = {f['path']: f['hash'] for f in local_manifest.get('files', [])}
    remote_files = {f['path']: f['hash'] for f in remote_manifest.get('files', [])}
    
    files_to_download = []
    files_to_delete = []
    
    poppler_dir = app_dir / POPPLER_FOLDER_NAME
    launcher_exe = Path(sys.executable) if getattr(sys, 'frozen', False) else None
    
    # 1. 원격 파일 검사
    for path, remote_hash in remote_files.items():
        file_path = app_dir / path
        needs_download = False
        reason = ""
        
        if not file_path.exists():
            needs_download = True
            reason = '파일 없음'
        else:
            local_hash = local_files.get(path)
            if local_hash is None:
                actual_hash = calculate_file_hash(file_path)
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
    
    # 2. 로컬 파일 검사 (보호 폴더 제외)
    for path in local_files.keys():
        if path not in remote_files:
            file_path = app_dir / path
            
            # 보호 대상 확인
            is_protected = False
            try:
                if file_path.is_relative_to(poppler_dir):
                    is_protected = True
            except:
                pass
            
            if launcher_exe and "launcher" in file_path.name.lower():
                is_protected = True
            
            if is_protected:
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


def update_main_app(app_dir: Path, log_callback):
    """메인 앱 업데이트 수행"""
    try:
        if log_callback:
            log_callback("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━", False)
            log_callback("[ 애플리케이션 업데이트 확인 ]", False)
            log_callback("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━", False)
            log_callback("→ 업데이트 확인 시작", False)

        remote_manifest = get_remote_manifest(APP_REPO_OWNER, APP_REPO_NAME, APP_BRANCH, log_callback)
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

        # 버전 비교
        if local_version != remote_version:
            # 케이스 1: 버전이 다름 - 전체 업데이트
            if log_callback:
                log_callback("", False)
                log_callback("! 새 버전 발견 - 전체 업데이트 시작", False)
            
            delete_all_except_protected(app_dir, log_callback)
            
            if log_callback:
                log_callback("", False)
            
            success_count = download_all_files(APP_REPO_OWNER, APP_REPO_NAME, remote_manifest, app_dir, log_callback, APP_BRANCH)
            
            if success_count > 0:
                save_local_manifest(app_dir, remote_manifest)
                if log_callback:
                    log_callback("", False)
                    log_callback(f"✓ 업데이트 완료 (v{local_version} → v{remote_version})", False)
                return True
            else:
                if log_callback:
                    log_callback("", False)
                    log_callback("✗ 업데이트 실패 (다운로드 실패)", False)
                return False
        
        else:
            # 케이스 2: 버전이 같음 - 무결성 검사
            if log_callback:
                log_callback("", False)
                log_callback("✓ 버전 일치 - 무결성 검사 시작", False)
            
            check_result = integrity_check_and_fix(APP_REPO_OWNER, APP_REPO_NAME, local_manifest, remote_manifest, app_dir, log_callback, APP_BRANCH)
            files_to_download = check_result['to_download']
            files_to_delete = check_result['to_delete']
            
            if not files_to_download and not files_to_delete:
                if log_callback:
                    log_callback("", False)
                    log_callback("✓ 모든 파일이 정상입니다", False)
                return False
            
            # 삭제 처리
            if files_to_delete:
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
            
            # 다운로드 처리
            if files_to_download:
                if log_callback:
                    log_callback("", False)
                    log_callback(f"→ 손상된 파일 복구 시작 ({len(files_to_download)}개)", False)
                
                success_count = 0
                fail_count = 0
                
                for i, file_info in enumerate(files_to_download, start=1):
                    file_path = file_info['path']
                    expected_hash = file_info['hash']
                    dest_path = app_dir / file_path
                    
                    if log_callback:
                        log_callback(f"  → [{i}/{len(files_to_download)}] {file_path} 다운로드 중...", False)
                    
                    if download_file_from_github(APP_REPO_OWNER, APP_REPO_NAME, file_path, dest_path, APP_BRANCH):
                        downloaded_hash = calculate_file_hash(dest_path)
                        if downloaded_hash == expected_hash:
                            success_count += 1
                            if log_callback:
                                log_callback(f"  ✓ [{i}/{len(files_to_download)}] {file_path} 완료", False)
                        else:
                            fail_count += 1
                            if log_callback:
                                log_callback(f"  ✗ [{i}/{len(files_to_download)}] {file_path} 해시 불일치", False)
                    else:
                        fail_count += 1
                        if log_callback:
                            log_callback(f"  ✗ [{i}/{len(files_to_download)}] {file_path} 다운로드 실패", False)
                
                if log_callback:
                    if fail_count > 0:
                        log_callback(f"  ✓ 복구 완료: 성공 {success_count}개, 실패 {fail_count}개", False)
                    else:
                        log_callback(f"  ✓ 복구 완료: {success_count}개", False)
            
            # manifest 업데이트
            if files_to_download or files_to_delete:
                save_local_manifest(app_dir, remote_manifest)
                if log_callback:
                    log_callback("", False)
                    log_callback("✓ 무결성 검사 및 복구 완료", False)
                return True
            
            return False

    except Exception as e:
        if log_callback:
            log_callback(f"✗ 업데이트 중 오류 발생: {e}", False)
        print(f"업데이트 실패: {e}")
        return False


# ---------- Poppler 관련 함수 ----------

def get_poppler_directory(app_dir: Path):
    """Poppler 디렉토리 반환"""
    poppler_dir = app_dir / POPPLER_FOLDER_NAME
    poppler_dir.mkdir(parents=True, exist_ok=True)
    return poppler_dir


def get_installed_poppler_version(poppler_dir: Path):
    """설치된 Poppler 버전 확인 (폴더명에서 추출)"""
    if not poppler_dir.exists():
        return None
    
    # poppler 디렉토리 내에서 버전 정보가 포함된 폴더 찾기
    for item in poppler_dir.iterdir():
        if item.is_dir() and "poppler" in item.name.lower():
            # 폴더명에서 버전 추출 (예: poppler-24.08.0 -> 24.08.0)
            match = re.search(r'poppler[_-]?(\d+\.\d+\.\d+)', item.name.lower())
            if match:
                return match.group(1)
    
    return None


def get_latest_poppler_version(log_callback=None):
    """Poppler 최신 버전 정보 가져오기"""
    api_url = f"https://api.github.com/repos/{POPPLER_REPO_OWNER}/{POPPLER_REPO_NAME}/releases/latest"
    try:
        if log_callback:
            log_callback("  → GitHub API 요청 중...", False)
        resp = requests.get(api_url, timeout=10, verify=VERIFY_SSL)
        resp.raise_for_status()
        data = resp.json()
        
        if log_callback:
            log_callback("  ✓ 최신 릴리즈 정보 확인 완료", False)
        
        for asset in data.get("assets", []):
            if asset["name"].lower().endswith(".zip"):
                # 파일명에서 버전 추출 (예: poppler-24.08.0.zip -> 24.08.0)
                match = re.search(r'poppler[_-]?(\d+\.\d+\.\d+)', asset["name"].lower())
                version = match.group(1) if match else "unknown"
                
                if log_callback:
                    log_callback(f"  ✓ 다운로드 파일: {asset['name']}", False)
                return asset["browser_download_url"], asset["name"], version
        return None, None, None
    except Exception as e:
        if log_callback:
            log_callback(f"  ✗ API 요청 실패: {e}", False)
        return None, None, None


def download_and_extract_poppler(dest_folder: Path, log_callback=None):
    """Poppler 다운로드 및 압축 해제"""
    if log_callback:
        log_callback("→ Poppler 다운로드 준비", False)

    download_url, filename, version = get_latest_poppler_version(log_callback)
    if not download_url:
        raise RuntimeError("Poppler 윈도우용 최신 zip 파일을 찾을 수 없습니다.")

    zip_path = dest_folder / filename

    if log_callback:
        log_callback(f"→ Poppler 다운로드 시작: {filename}", False)

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
            log_callback(f"  ✓ 다운로드 완료: {filename}", False)
    except Exception as e:
        if log_callback:
            log_callback(f"  ✗ 다운로드 실패: {e}", False)
        raise

    # 압축 해제
    if log_callback:
        log_callback("→ Poppler 압축 해제 시작", False)

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
                log_callback("  ✓ 임시 zip 파일 삭제 완료", False)

    # 압축 해제된 폴더 찾기 (poppler-버전 형태)
    if log_callback:
        log_callback("→ Poppler 폴더 확인 중...", False)

    poppler_folder = None
    for item in dest_folder.iterdir():
        if item.is_dir() and "poppler" in item.name.lower():
            # Library/bin 폴더가 있는지 확인
            bin_path = item / "Library" / "bin"
            if bin_path.exists() and (bin_path / "pdftoppm.exe").exists():
                poppler_folder = item
                break
    
    if poppler_folder:
        if log_callback:
            log_callback(f"  ✓ Poppler 폴더 발견: {poppler_folder.name}", False)
            log_callback(f"  ✓ 버전: v{version}", False)
            log_callback("✓ Poppler 설치 완료", False)
        
        return poppler_folder / "Library" / "bin"
    else:
        raise RuntimeError("Poppler 폴더를 찾을 수 없습니다.")


def check_poppler(app_dir: Path, log_callback):
    """Poppler 확인, 버전 체크 및 설치/업데이트"""
    if log_callback:
        log_callback("", False)
        log_callback("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━", False)
        log_callback("[ Poppler 유효성 검사 ]", False)
        log_callback("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━", False)
        time.sleep(0.2)
    
    poppler_dir = get_poppler_directory(app_dir)
    
    if log_callback:
        log_callback(f"→ Poppler 설치 경로: {poppler_dir}", False)

    # 설치된 버전 확인 (폴더명에서 추출)
    installed_version = get_installed_poppler_version(poppler_dir)
    
    if installed_version:
        if log_callback:
            log_callback(f"  • 설치된 버전: v{installed_version}", False)
        
        # 최신 버전 확인
        if log_callback:
            log_callback("→ 최신 버전 확인 중...", False)
        
        _, _, latest_version = get_latest_poppler_version(log_callback)
        
        if latest_version:
            if log_callback:
                log_callback(f"  • 최신 버전: v{latest_version}", False)
            
            if installed_version == latest_version:
                if log_callback:
                    log_callback("", False)
                    log_callback("✓ Poppler가 최신 버전입니다", False)
            else:
                if log_callback:
                    log_callback("", False)
                    log_callback("! 새 버전 발견 - Poppler 업데이트 시작", False)
                    log_callback("→ 기존 Poppler 삭제 중...", False)
                
                # 기존 Poppler 폴더 삭제
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
                
                # 새 버전 설치
                try:
                    download_and_extract_poppler(poppler_dir, log_callback)
                except Exception as e:
                    if log_callback:
                        log_callback(f"✗ Poppler 업데이트 실패: {e}", False)
                    raise
        else:
            if log_callback:
                log_callback("", False)
                log_callback("✓ Poppler가 설치되어 있습니다 (버전 확인 실패)", False)
    
    else:
        # Poppler 없음 - 새로 설치
        if log_callback:
            log_callback("", False)
            log_callback("! Poppler를 찾을 수 없습니다", False)
            log_callback("→ Poppler 다운로드 시작", False)
            time.sleep(0.2)

        try:
            download_and_extract_poppler(poppler_dir, log_callback)
        except Exception as e:
            if log_callback:
                log_callback(f"✗ Poppler 설치 실패: {e}", False)
            raise


# ---------- UI 컴포넌트 ----------

class LauncherPopup(ctk.CTkToplevel):
    """런처 로그 표시 팝업"""

    def __init__(self, parent):
        super().__init__(parent)
        self.title("초기화 중...")
        self.geometry("700x400")
        self.resizable(False, False)

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
        try:
            self.log_queue.put((message, is_progress))
        except Exception as e:
            print(f"safe_add_log 실패: {e}")

    def close_window(self):
        try:
            self._closing = True
            self.grab_release()
            self.destroy()
        except Exception as e:
            print(f"창 닫기 실패: {e}")


class LauncherApp(ctk.CTk):
    """런처 메인 앱"""

    def __init__(self):
        super().__init__()
        self.title("PDF → JPG 변환기 런처")
        self.geometry("1x1")
        self.withdraw()

        self.app_dir = get_app_directory()
        self.result = {"should_close": False, "launch_main": True}

        self.popup = LauncherPopup(self)

        def init_thread():
            try:
                def log_cb(msg, is_progress):
                    self.popup.safe_add_log(msg, is_progress)
                    time.sleep(0.05)

                # 업데이트 확인
                updated = update_main_app(self.app_dir, log_cb)
                time.sleep(0.3)

                # Poppler 확인
                check_poppler(self.app_dir, log_cb)
                time.sleep(0.3)

                self.popup.safe_add_log("", False)
                self.popup.safe_add_log("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━", False)
                for i in range(3, 0, -1):
                    self.popup.safe_add_log(f"✓ {i}초 후 프로그램이 시작됩니다...", False)
                    time.sleep(1.0)
                self.popup.safe_add_log("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━", False)

                self.result["should_close"] = True

            except Exception as e:
                print(f"초기화 오류: {e}")
                self.popup.safe_add_log(f"✗ 초기화 오류: {e}", False)
                self.result["should_close"] = True
                self.result["launch_main"] = False

        def start_init():
            threading.Thread(target=init_thread, daemon=True).start()

        self.popup.after(100, start_init)

        def check_close():
            if self.result["should_close"]:
                try:
                    self.popup.close_window()
                except Exception:
                    pass
            else:
                try:
                    self.popup.after(100, check_close)
                except Exception:
                    pass

        self.popup.after(500, check_close)

        try:
            self.wait_window(self.popup)
        except Exception:
            pass

        # 메인 앱 실행
        if self.result["launch_main"]:
            main_exe = self.app_dir / MAIN_EXE_NAME
            if main_exe.exists():
                subprocess.Popen([str(main_exe)])
            else:
                messagebox.showerror("오류", f"메인 앱을 찾을 수 없습니다:\n{main_exe}")

        self.destroy()


# ---------- 메인 진입점 ----------

if __name__ == "__main__":
    app = LauncherApp()
    app.mainloop()
