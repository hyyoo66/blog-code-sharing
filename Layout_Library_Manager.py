#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Layout_Library_Manager.py
Layout_Library_Manager.html 을 크롬에 띄워주고, 그 페이지가 요청하는
로컬 작업(외부 exe 실행)을 대신 처리해주는 서버.

  - Layout_Library_Manager.html 자체를 서빙 (실행하면 크롬이 자동으로 열림)
  - .LLB/.MAX 파일을 추가하면 maxminw.exe로 .MIN 텍스트로 변환
  - "Run Layout" 버튼을 누르면 lsession.exe(Layout Plus)를 실행
  - (참고용) maxminb.exe로 .MIN -> .LLB 변환 시도
  - 페이지와 서버가 하트비트로 서로의 생사를 지켜본다: 창을 닫으면 서버가 스스로
    종료되고, 서버를 끄면(이 창에서 Ctrl+C 등) 페이지가 안내 후 스스로 탭을 닫는다.

실행
----
    python Layout_Library_Manager.py
실행하면 http://127.0.0.1:8765 가 브라우저에 자동으로 열린다. 끄려면 Ctrl+C.

파일을 직접 더블클릭해서(file://) 써도 되지만, 이 서버로 열면 localhost라
보안 컨텍스트가 되어 클립보드 복사·저장 대화상자가 확실하게 동작하고,
브릿지와 페이지의 출처가 같아져 CORS 문제도 생기지 않는다.

주의: maxminb.exe(MIN->MAX/LLB 방향)는 과거 테스트에서 인자 조합을 여러 개
바꿔봐도 항상 exit code 8("Unexpected DB Version found")로 실패했던 이력이
있다. 확실하게 되는 방법은 여전히 Layout Plus의 File -> Import -> MIN
Interchange 메뉴를 손으로 3번 클릭하는 것뿐이다. 이 엔드포인트는 혹시나 하는
용도로 만들어둔 것이라 실패해도 정상이다.
"""

import http.server
import subprocess
import tempfile
import shutil
import os
import json
import mimetypes
import threading
import socket
import time
import urllib.request
import webbrowser
from pathlib import Path
from urllib.parse import unquote

# 경로가 바뀌면 여기만 고치면 됨
MAXMINW = r"C:\OrCAD\OrCAD_10.5\tools\layout_plus\maxminw.exe"
MAXMINB = r"C:\OrCAD\OrCAD_10.5\tools\layout_plus\maxminb.exe"
# LSESSION = r"C:\OrCAD\OrCAD_10.5\tools\layout_plus\lsession.exe"  # TODO: 경로설정 기능 테스트 끝나면 이 줄로 되돌리기
LSESSION = r"Z:\OrCAD\OrCAD_10.5\tools\layout_plus\lsession.exe"  # C: -> Z: 로만 바꿔 일부러 없는 경로로 — ensure_lsession_path() 테스트용
PORT = 8765

# lsession.exe(Layout Plus) 경로를 사용자가 직접 골랐을 때 기억해두는 설정 파일.
# 다운로드 폴더에 저장 — 이 스크립트를 다른 곳으로 옮겨도(또는 재실행해도) 계속 남아있게.
LSESSION_CONFIG_FILE = Path.home() / "Downloads" / "Layout_Library_Manager_config.json"

# 크롬 실행 파일 위치. 여기 목록에서 먼저 찾아지는 걸 쓰고, 하나도 없으면 기본 브라우저로 연다.
CHROME_CANDIDATES = [
    r"C:\Program Files\Google\Chrome\Application\chrome.exe",
    r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
    os.path.expandvars(r"%LOCALAPPDATA%\Google\Chrome\Application\chrome.exe"),
]

# 이 스크립트가 있는 폴더를 웹 루트로 쓴다. 주소창에 "/"만 치면 아래 HTML을 준다.
WEB_ROOT = Path(__file__).resolve().parent
INDEX_FILE = "Layout_Library_Manager.html"

# 페이지가 이 시간(초) 넘게 하트비트를 안 보내면 창이 닫힌 것으로 보고 서버를 스스로 끈다.
# 페이지는 1.5초마다 보내므로 몇 번 놓쳐도 여유 있게 잡되, 너무 길면 창을 닫아도 한참 안 꺼진다.
HEARTBEAT_TIMEOUT = 6.0

_last_heartbeat = None      # None = 아직 페이지가 한 번도 연결한 적 없음(감시 시작 전)
_heartbeat_lock = threading.Lock()


def _read_lsession_config():
    if LSESSION_CONFIG_FILE.exists():
        try:
            return json.loads(LSESSION_CONFIG_FILE.read_text(encoding="utf-8"))
        except Exception:
            return {}
    return {}


def _write_lsession_config(path):
    try:
        LSESSION_CONFIG_FILE.parent.mkdir(parents=True, exist_ok=True)
        config = _read_lsession_config()
        config["lsession_path"] = path
        LSESSION_CONFIG_FILE.write_text(json.dumps(config, ensure_ascii=False, indent=2), encoding="utf-8")
        print(f"[OK] lsession.exe 경로를 저장했습니다: {LSESSION_CONFIG_FILE}")
    except Exception as e:
        print(f"[경고] 설정 저장 실패: {e}")


def _pick_lsession_path():
    """탐색기 창으로 사용자에게 lsession.exe 위치를 직접 고르게 한다. 취소하면 None."""
    try:
        import tkinter as tk
        from tkinter import filedialog, messagebox
        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        messagebox.showinfo(
            "Layout Plus 위치 지정",
            f"lsession.exe(Layout Plus)를 다음 경로에서 찾지 못했습니다:\n{LSESSION}\n\n"
            "이어지는 탐색기 창에서 lsession.exe 파일을 직접 선택해주세요.",
            parent=root,
        )
        path = filedialog.askopenfilename(
            title="lsession.exe 선택 (Layout Plus)",
            filetypes=[("lsession.exe", "lsession.exe"), ("실행 파일", "*.exe"), ("모든 파일", "*.*")],
            parent=root,
        )
        root.destroy()
        return path or None
    except Exception as e:
        print(f"[경고] 파일 선택 창을 열지 못했습니다: {e}")
        return None


def ensure_lsession_path():
    """LSESSION 경로가 실제로 있는지 확인하고, 없으면 저장된 설정 -> 탐색기로 직접 지정 순으로 찾는다."""
    global LSESSION
    if os.path.exists(LSESSION):
        return

    saved = _read_lsession_config().get("lsession_path")
    if saved and os.path.exists(saved):
        print(f"[OK] 저장된 lsession.exe 경로를 사용합니다: {saved}")
        LSESSION = saved
        return

    print(f"[알림] lsession.exe 경로가 없습니다: {LSESSION}")
    print("       탐색기 창을 띄워 직접 선택하도록 안내합니다...")
    picked = _pick_lsession_path()
    if picked:
        LSESSION = picked
        _write_lsession_config(picked)
    else:
        print("[경고] lsession.exe 경로를 지정하지 않았습니다 — 'Run Layout' 버튼이 동작하지 않습니다.")
        print(f"       나중에라도 이 파일 상단의 LSESSION 값을 고치거나, {LSESSION_CONFIG_FILE} 을 직접 편집하세요.")


def _is_layout_plus_running():
    """lsession.exe(Layout Plus)가 이미 떠 있는지 확인 — 이 도구로 띄웠든 손으로 띄웠든 상관없이
    Windows tasklist로 실제 프로세스 목록을 봐서 판단한다(버튼 두 번 눌러서 창이 두 개 뜨는 것 방지)."""
    exe_name = Path(LSESSION).name
    try:
        out = subprocess.run(
            ["tasklist", "/FI", f"IMAGENAME eq {exe_name}", "/NH"],
            capture_output=True, text=True, timeout=5
        )
        return exe_name.lower() in out.stdout.lower()
    except Exception:
        return False  # 확인 실패 시엔 막지 말고 그냥 실행 시도로 넘어감


class Handler(http.server.BaseHTTPRequestHandler):
    def _cors(self):
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type, X-File-Ext")

    def do_OPTIONS(self):
        self.send_response(204)
        self._cors()
        self.end_headers()

    def do_GET(self):
        """WEB_ROOT 폴더 안의 파일만 서빙한다. "/"는 Layout_Library_Manager.html."""
        # 브라우저는 한글 등 비ASCII 경로를 %xx로 인코딩해서 보내므로 먼저 디코드한다
        rel = unquote(self.path.split("?", 1)[0].split("#", 1)[0]).lstrip("/")
        if rel in ("", "/"):
            rel = INDEX_FILE
        target = (WEB_ROOT / rel).resolve()

        # 상위 폴더로 빠져나가는 요청(../ 등)은 거부
        if WEB_ROOT not in target.parents and target != WEB_ROOT:
            self._error(403, "허용되지 않은 경로입니다.")
            return
        if not target.is_file():
            self._error(404, f"파일을 찾을 수 없습니다: {rel}\n"
                             f"(웹 루트: {WEB_ROOT})")
            return

        ctype = mimetypes.guess_type(str(target))[0] or "application/octet-stream"
        if ctype.startswith("text/") or ctype in ("application/javascript", "application/json"):
            ctype += "; charset=utf-8"
        body = target.read_bytes()
        self.send_response(200)
        self._cors()
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")  # 코드를 고치면 새로고침만으로 바로 반영
        self.end_headers()
        self.wfile.write(body)

    def do_POST(self):
        if self.path == "/convert":
            self._handle_convert_to_min()
        elif self.path == "/convert-to-llb":
            self._handle_convert_to_llb()
        elif self.path == "/launch-layout-plus":
            self._handle_launch_layout_plus()
        elif self.path == "/heartbeat":
            self._handle_heartbeat()
        elif self.path == "/shutdown":
            self._handle_shutdown()
        else:
            self.send_response(404)
            self._cors()
            self.end_headers()

    def _handle_heartbeat(self):
        global _last_heartbeat
        with _heartbeat_lock:
            _last_heartbeat = time.time()
        body = b"OK"
        self.send_response(200)
        self._cors()
        self.send_header("Content-Type", "text/plain; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def _handle_shutdown(self):
        # 페이지가 정상적으로 닫힐 때(pagehide -> sendBeacon) 곧바로 서버를 끈다.
        # 하트비트 시간초과(HEARTBEAT_TIMEOUT)까지 기다리지 않아도 되는 빠른 경로.
        body = b"OK"
        self.send_response(200)
        self._cors()
        self.send_header("Content-Type", "text/plain; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)
        print("[알림] 페이지가 닫혀 서버를 종료합니다.")
        # 응답을 다 보낸 뒤에 끄기 위해 별도 스레드에서 shutdown() 호출 (여기서 바로 부르면
        # 응답이 나가기 전에 연결이 끊길 수 있음)
        threading.Thread(target=self.server.shutdown, daemon=True).start()

    def _handle_convert_to_min(self):
        length = int(self.headers.get("Content-Length", 0))
        data = self.rfile.read(length)
        ext = self.headers.get("X-File-Ext", ".LLB")
        if not ext.startswith("."):
            ext = "." + ext

        if not os.path.exists(MAXMINW):
            self._error(500, f"maxminw.exe 경로를 찾을 수 없습니다: {MAXMINW}\n"
                              f"Layout_Library_Manager.py 상단의 MAXMINW 값을 실제 경로로 고쳐주세요.")
            return

        tmpdir = tempfile.mkdtemp(prefix="minbridge_")
        try:
            # maxminw.exe가 한글/괄호 섞인 경로에서 조용히 실패하는 걸 이미 겪어봤으므로
            # 항상 순수 ASCII 임시 폴더(tempfile 기본 위치)에서 변환한다.
            in_path = os.path.join(tmpdir, "in" + ext)
            out_path = os.path.join(tmpdir, "out.MIN")
            with open(in_path, "wb") as f:
                f.write(data)

            result = subprocess.run([MAXMINW, in_path, out_path],
                                     capture_output=True, timeout=60)
            if not os.path.exists(out_path):
                self._error(500, f"변환 실패 (maxminw.exe exit={result.returncode})\n"
                                  f"{result.stdout.decode('cp949', 'replace')}\n"
                                  f"{result.stderr.decode('cp949', 'replace')}")
                return

            with open(out_path, "r", encoding="utf-8", errors="replace") as f:
                min_text = f.read()

            body = min_text.encode("utf-8")
            self.send_response(200)
            self._cors()
            self.send_header("Content-Type", "text/plain; charset=utf-8")
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)
            print(f"[OK] 변환: {in_path} ({len(data)} bytes) -> MIN {len(body)} bytes")
        except subprocess.TimeoutExpired:
            self._error(500, "maxminw.exe 실행이 60초를 넘겨 중단했습니다.")
        except Exception as e:
            self._error(500, f"서버 에러: {e}")
        finally:
            shutil.rmtree(tmpdir, ignore_errors=True)

    def _handle_convert_to_llb(self):
        # 주의: maxminb.exe는 과거 테스트에서 항상 exit 8로 실패했다(파일 상단 설명 참고).
        # 실패해도 놀라지 말 것 — 확실한 경로는 Layout Plus의 수동 Import뿐이다.
        length = int(self.headers.get("Content-Length", 0))
        min_text = self.rfile.read(length).decode("utf-8", "replace")

        if not os.path.exists(MAXMINB):
            self._error(500, f"maxminb.exe 경로를 찾을 수 없습니다: {MAXMINB}\n"
                              f"Layout_Library_Manager.py 상단의 MAXMINB 값을 실제 경로로 고쳐주세요.")
            return

        tmpdir = tempfile.mkdtemp(prefix="minbridge_")
        try:
            in_path = os.path.join(tmpdir, "in.MIN")
            out_path = os.path.join(tmpdir, "out.LLB")
            with open(in_path, "w", encoding="utf-8") as f:
                f.write(min_text)

            result = subprocess.run([MAXMINB, in_path, out_path],
                                     capture_output=True, timeout=60)
            if not os.path.exists(out_path) or os.path.getsize(out_path) == 0:
                self._error(500, f"변환 실패 (maxminb.exe exit={result.returncode}) — 알려진 문제입니다.\n"
                                  f"{result.stdout.decode('cp949', 'replace')}\n"
                                  f"{result.stderr.decode('cp949', 'replace')}\n"
                                  f"대신 .MIN을 저장한 뒤 Layout Plus에서 File -> Import -> MIN Interchange로 넣어주세요.")
                return

            with open(out_path, "rb") as f:
                body = f.read()

            self.send_response(200)
            self._cors()
            self.send_header("Content-Type", "application/octet-stream")
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)
            print(f"[OK] 변환: in.MIN ({len(min_text)} chars) -> LLB {len(body)} bytes")
        except subprocess.TimeoutExpired:
            self._error(500, "maxminb.exe 실행이 60초를 넘겨 중단했습니다.")
        except Exception as e:
            self._error(500, f"서버 에러: {e}")
        finally:
            shutil.rmtree(tmpdir, ignore_errors=True)

    def _handle_launch_layout_plus(self):
        # kicad2min.py의 실행 방식과 동일: lsession.exe를 tools 폴더(부모의 부모)에서 실행
        if not os.path.exists(LSESSION):
            self._error(500, f"lsession.exe 경로를 찾을 수 없습니다: {LSESSION}\n"
                              f"Layout_Library_Manager.py 상단의 LSESSION 값을 실제 경로로 고쳐주세요.")
            return
        if _is_layout_plus_running():
            self._error(409, "Layout Plus(lsession.exe)가 이미 실행 중입니다. 기존 창을 사용하세요.")
            return
        try:
            cwd = str(Path(LSESSION).parent.parent)  # ...\tools\layout_plus\.. -> ...\tools
            subprocess.Popen([LSESSION], cwd=cwd)
            body = b"OK"
            self.send_response(200)
            self._cors()
            self.send_header("Content-Type", "text/plain; charset=utf-8")
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)
            print("[OK] Layout Plus 실행")
        except Exception as e:
            self._error(500, f"Layout Plus 실행 실패: {e}")

    def _error(self, code, message):
        body = message.encode("utf-8")
        self.send_response(code)
        self._cors()
        self.send_header("Content-Type", "text/plain; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)
        print(f"[ERROR] {message}")

    def log_message(self, format, *args):
        pass  # 콘솔 조용히


class Server(http.server.ThreadingHTTPServer):
    # 윈도우는 SO_REUSEADDR가 켜져 있으면 "이미 듣고 있는 포트"에도 바인딩이 되어버려서
    # 두 서버가 같은 포트를 물고 요청이 뒤섞인다. 중복 실행을 확실히 막으려고 꺼둔다.
    allow_reuse_address = False


def probe_running_server(url):
    """이미 이 포트를 쓰는 서버가 있는지 확인.
    None = 없음 / "ok" = 지금 버전이 이미 실행 중 / "old" = 응답은 하는데 GET을 모르는 옛 버전."""
    with socket.socket() as s:
        s.settimeout(0.3)
        if s.connect_ex(("127.0.0.1", PORT)) != 0:
            return None
    try:
        with urllib.request.urlopen(url, timeout=2) as r:
            return "ok" if r.status == 200 else "old"
    except Exception:
        # 501(GET 미지원) 등 -> do_GET 없는 예전 min_bridge_server.py가 떠 있는 경우
        return "old"


def heartbeat_watchdog(server):
    """페이지가 창을 닫아 하트비트가 끊기면 서버를 스스로 종료한다.
    /shutdown(정상 종료 시 즉시 호출)이 이미 처리했다면 여기서는 아무 일도 안 하고 조용히 끝난다."""
    while True:
        time.sleep(1)
        with _heartbeat_lock:
            last = _last_heartbeat
        if last is not None and (time.time() - last) > HEARTBEAT_TIMEOUT:
            print("[알림] 페이지 연결(하트비트)이 끊겨 서버를 종료합니다.")
            server.shutdown()
            return


def open_in_chrome(url):
    """크롬으로 연다. 크롬을 못 찾으면 기본 브라우저로 대체."""
    for exe in CHROME_CANDIDATES:
        if os.path.exists(exe):
            try:
                subprocess.Popen([exe, "--new-window", url])
                print(f"[OK] 크롬으로 열었습니다: {exe}")
                return
            except Exception as e:
                print(f"[경고] 크롬 실행 실패({e}) — 기본 브라우저로 엽니다.")
                break
    else:
        print("[경고] 크롬을 찾지 못했습니다 — 기본 브라우저로 엽니다.")
        print("       크롬 경로가 다르면 이 파일 상단의 CHROME_CANDIDATES에 추가하세요.")
    webbrowser.open(url)


def main():
    ensure_lsession_path()
    if not os.path.exists(MAXMINW):
        print(f"[경고] MAXMINW 경로가 없습니다: {MAXMINW}")
        print(f"       이 파일 상단의 MAXMINW 값을 고친 뒤 다시 실행하세요.")
    if not (WEB_ROOT / INDEX_FILE).is_file():
        print(f"[경고] {INDEX_FILE} 을 찾을 수 없습니다: {WEB_ROOT}")
        print(f"       이 스크립트를 HTML과 같은 폴더에 두세요.")

    url = f"http://127.0.0.1:{PORT}/"

    # 이미 떠 있으면 또 띄우지 않는다 (두 번 실행해도 창만 하나 더 열릴 뿐)
    state = probe_running_server(url)
    if state == "ok":
        print(f"[알림] 이미 실행 중입니다 ({url})")
        print("       서버를 새로 띄우지 않고 창만 엽니다.")
        open_in_chrome(url)
        return
    if state == "old":
        print(f"[경고] 포트 {PORT}에 예전 버전 서버가 떠 있습니다.")
        print("       그 서버를 실행한 창에서 Ctrl+C로 끈 뒤 다시 실행하세요.")
        print("       (예전 min_bridge_server.py는 페이지 서빙을 못 해서 501 오류가 납니다)")
        return

    try:
        server = Server(("127.0.0.1", PORT), Handler)
    except OSError as e:
        print(f"[에러] 포트 {PORT}을 열 수 없습니다: {e}")
        print("       다른 프로그램이 쓰고 있으면 이 파일 상단의 PORT 값을 바꿔주세요.")
        return

    print(f"[Layout Library Manager] {url}")
    print("브라우저를 띄웁니다. 끄려면 이 창에서 Ctrl+C.")

    # 서버가 요청을 받을 준비가 된 뒤에 열리도록 살짝 늦춰서 브라우저 실행
    threading.Timer(0.5, lambda: open_in_chrome(url)).start()
    # 페이지가 창을 닫으면(=하트비트가 끊기면) 서버도 같이 종료되도록 감시 스레드를 띄운다
    threading.Thread(target=heartbeat_watchdog, args=(server,), daemon=True).start()
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\n종료합니다.")
    finally:
        # shutdown()은 서비스 루프만 멈출 뿐 소켓은 그대로 붙잡고 있어서, 곧바로 다시 실행하면
        # "주소 사용 중" 에러가 난다. 어떤 경로로 빠져나오든 여기서 확실히 반납한다.
        server.server_close()


if __name__ == "__main__":
    main()
