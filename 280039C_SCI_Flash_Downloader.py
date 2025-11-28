import sys
import os
import subprocess
from datetime import datetime

# ===== 사용자 설정 =====
COM_PORT   = "COM10"      # ★ 윈도우 장치관리자에서 CP2102가 할당된 포트 확인
BAUD_RATE  = "38400"      # ★ 보레이트, 9600 / 19200 / 38400
# ======================

HEX2000     = r"C:\ti\ccs2020\ccs\tools\compiler\ti-cgt-c2000_22.6.2.LTS\bin\hex2000.exe"
SERIAL_EXE  = r"C:\ti\c2000\C2000Ware_5_05_00_00\utilities\flash_programmers\serial_flash_programmer\serial_flash_programmer.exe"
KERNEL_TXT  = r"C:\ti\F28003x_sci_flash_kernel.txt"

SAVE_DIR = r"C:\ti\tmp"  # 임시 txt 저장 폴더
os.makedirs(SAVE_DIR, exist_ok=True)

print("\n=======================================")
print("     F28003x 시리얼 플래시롬 다운로더")
print("=======================================\n")

# ===============================
#   .out 파일 자동 탐색 기능
# ===============================

def find_out_files(project_path):
    out_list = []
    for root, dirs, files in os.walk(project_path):
        for f in files:
            if f.lower().endswith(".out"):
                fullpath = os.path.join(root, f)
                mtime = os.path.getmtime(fullpath)
                out_list.append((fullpath, mtime))
    return out_list


# -------------------------------
# 입력 처리
# -------------------------------
if len(sys.argv) >= 2:
    input_path = sys.argv[1]
else:
    input_path = input(">> OUT 파일 또는 프로젝트 폴더 경로 : ").strip()


# 경로 분류
if os.path.isdir(input_path):

    print("\n>> 프로젝트 폴더 탐색 중...\n")
    out_files = find_out_files(input_path)

    if not out_files:
        print("⚠️  오류: 해당 폴더에서 .out 파일을 찾지 못했습니다.")
        sys.exit(1)

    # 최신순 정렬
    out_files.sort(key=lambda x: x[1], reverse=True)

    print(">> 찾은 .out 파일 목록입니다.\n")
    for i, (fpath, ts) in enumerate(out_files):
        tstr = datetime.fromtimestamp(ts).strftime("%Y-%m-%d %H:%M:%S")
        mark = "  (최신)" if i == 0 else ""
        print(f" {i}) {fpath}\n      → {tstr}{mark}\n")

    # 선택: 엔터 치면 자동 0번
    raw = input(">> 다운로드할 OUT 파일을 선택하세요(엔터=0번) : ").strip()
    idx = 0 if raw == "" else int(raw)

    out_path = out_files[idx][0]

else:
    out_path = input_path

print(f"\n>> 선택된 OUT 파일 → {out_path}")

# 변환될 TXT 파일 경로
txt_path = os.path.join(
    SAVE_DIR,
    os.path.basename(out_path).replace(".out", ".txt")
)

# ===============================
#   1) HEX 변환
# ===============================
print("\n>> hex2000 변환 중...\n")

cmd_hex = [
    HEX2000,
    out_path,
    "-boot",
    "-sci8",
    "-a",
    "-o", txt_path
]

try:
    subprocess.run(cmd_hex, check=True)
    print(f"✅ 변환 성공 → {txt_path}\n")
except Exception as e:
    print(f"⚠️  HEX 변환 실패: {e}")
    sys.exit(1)

print(">> DSP를 SCI BOOT 모드로 설정하십시오.")
input("\n>> 엔터 키를 누르면 '커널' 다운로드를 시작합니다 : ")


# ===============================
#   2) serial_flash_programmer 실행
# ===============================
print("\n>> 'serial_flash_programmer.exe' 실행 중...")

try:
    subprocess.run([
        SERIAL_EXE,
        "-d", "f28003x",
        "-k", KERNEL_TXT,
        "-a", txt_path,
        "-p", COM_PORT,
        "-b", BAUD_RATE,
        "-q"
    ])
    input("\n✅ serial_flash_programmer.exe 실행 완료\n")

except Exception as e:
    print(f"⚠️  DSP 프로그래밍 실패: {e}")
