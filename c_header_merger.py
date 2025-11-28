import sys
import os

# 출력 디렉토리: 바탕화면
OUTPUT_DIR = os.path.join(os.path.expanduser("~"), "OneDrive", "Desktop")
os.makedirs(OUTPUT_DIR, exist_ok=True)

# 한글 설명 + 사용법
HELP_TEXT = """
============================================================
여러 개의 .c .h 파일을 하나로 묶어 하나의 txt 파일로 만드는 도구
 - 디렉토리를 입력하면 디렉토리 안의 .c / .h 파일을 모두 병합
 - 파일 여러 개를 직접 입력하면 입력한 파일만 병합
 - main.c 가 있으면 결과 파일 맨 앞에 오도록 자동 정렬
 - 출력 파일은 바탕화면에 저장됨
============================================================

사용법:
    DTMC3.py <디렉토리>
    DTMC3.py <파일1> <파일2> <파일3> ...

"""

def pause():
    input("\n아무 키나 누르면 종료합니다...")

def collect_files_from_dir(dir_path):
    file_list = []
    for root, _, files in os.walk(dir_path):
        for f in files:
            if f.endswith(".c") or f.endswith(".h"):
                file_list.append(os.path.join(root, f))
    return file_list


def sort_main_first(file_list):
    """main.c 있으면 맨 앞으로 정렬"""
    main_files = [f for f in file_list if os.path.basename(f).lower() == "main.c"]
    others = [f for f in file_list if os.path.basename(f).lower() != "main.c"]

    if main_files:
        return main_files + others, "main"
    else:
        return file_list, os.path.splitext(os.path.basename(file_list[0]))[0]


def merge_files(file_list, output_name):
    """파일을 UTF-8로 병합"""
    output_path = os.path.join(OUTPUT_DIR, output_name)

    with open(output_path, "w", encoding="utf-8") as fout:
        for f in file_list:
            fout.write(f"\n\n================== {f} ==================\n\n")
            try:
                with open(f, "r", encoding="utf-8", errors="ignore") as fin:
                    fout.write(fin.read())
            except Exception as e:
                fout.write(f"\n[파일 읽기 오류] {e}\n")

    return output_path


def main():
    args = sys.argv[1:]

    # 인수 없는 경우 → 설명 출력 후 종료
    if len(args) == 0:
        print(HELP_TEXT)
        pause()
        return

    # 디렉토리 하나만 입력한 경우
    if len(args) == 1 and os.path.isdir(args[0]):
        dir_path = args[0]
        file_list = collect_files_from_dir(dir_path)

        if len(file_list) == 0:
            print("디렉토리에 .c / .h 파일이 없습니다.")
            pause()
            return

        sorted_files, rep_name = sort_main_first(file_list)
        output_name = f"{os.path.basename(dir_path)}_combined.txt"

        out_path = merge_files(sorted_files, output_name)
        print(f"\n완료: {out_path}")
        pause()
        return

    # 파일 여러 개 입력한 경우
    elif all(os.path.isfile(a) for a in args):
        file_list = args
        sorted_files, rep_name = sort_main_first(file_list)
        output_name = f"{rep_name}_etc_combined.txt"

        out_path = merge_files(sorted_files, output_name)
        print(f"\n완료: {out_path}")
        pause()
        return

    else:
        print("\n오류: 디렉토리 또는 파일 여러 개 중 하나만 입력해야 합니다.")
        print(HELP_TEXT)
        pause()
        return


if __name__ == "__main__":
    main()
