import re
import pyperclip
from bs4 import BeautifulSoup
import sys
import hashlib

# 디버그 파일 생성 여부 설정
DEBUG_MODE = False  # True: 디버그 파일 생성 / False: 생성 안 함

# 마크다운 헤더 크기 설정
# 1개 이상의 #으로 시작하는 제목은 모두 이 개수의 #으로 통일
# # 의 수가 작을 수록 글자가 커짐
UNIFIED_HEADER_COUNT = 3

def html_to_markdown(html):
    """Gemini HTML → Markdown 변환 (스타일 초기화 + 2단계 크기 통일)"""
    soup = BeautifulSoup(html, 'html.parser')

    # 0. 스타일 제거
    for tag in soup.find_all(True):
        if tag.has_attr('style'): del tag['style']
        if tag.has_attr('class'): del tag['class']

    # 1. 제목 변환 (h1~h6 모두 지정된 개수의 #으로 통일)
    header_symbol = '#' * UNIFIED_HEADER_COUNT
    for i in range(1, 7):
        for h in soup.find_all(f"h{i}"):
            header_text = h.get_text().strip()
            if header_text:
                h.replace_with(f"\n\n{header_symbol} {header_text}\n\n")

    # 2. 서식 변환
    for pre in soup.find_all("pre"):
        code = pre.get_text("\n")
        pre.replace_with(f"\n```\n{code}\n```\n")
    for b in soup.find_all(["b", "strong"]):
        b.replace_with(f"**{b.get_text()}**")
    for i in soup.find_all(["i", "em"]):
        i.replace_with(f"*{i.get_text()}*")

    # 3. 수식 보호
    for mjx in soup.find_all("mjx-container"):
        tex = mjx.get_text().strip()
        if mjx.get("display") == "block":
            mjx.replace_with(f"\n$$\n{tex}\n$$\n")
        else:
            mjx.replace_with(f"${tex}$")

    # 4. 리스트
    for li in soup.find_all("li"):
        li.insert_before("* ")
        li.append("\n")
        li.unwrap()
    for ul in soup.find_all(["ul", "ol"]):
        ul.insert_before("\n")
        ul.append("\n")
        ul.unwrap()

    # 5. 문단 및 기타
    for p in soup.find_all(["p", "div"]):
        p.append("\n\n")
        p.unwrap()
    for br in soup.find_all("br"):
        br.replace_with("\n")
    for span in soup.find_all("span"):
        span.unwrap()

    return soup.get_text()

def is_html(text):
    """HTML 태그가 있는지 검사"""
    return bool(re.search(r'<[a-zA-Z][^>]*>', text))

def normalize_markdown_headers(text):
    """마크다운 헤더를 지정된 개수의 #으로 통일"""
    header_symbol = '#' * UNIFIED_HEADER_COUNT
    text = re.sub(r'^#{1,6}\s+', f'{header_symbol} ', text, flags=re.MULTILINE)
    return text

def process_gemini_html(raw_input):
    # HTML인지 마크다운인지 자동 감지
    if is_html(raw_input):
        # HTML → 마크다운 변환
        md = html_to_markdown(raw_input)
        md = re.sub(r'background[^;"]*;?', '', md)
    else:
        # 이미 마크다운인 경우 그대로 사용
        md = raw_input
    
    # 모든 마크다운 헤더를 ##로 통일
    md = normalize_markdown_headers(md)
    
    md = re.sub(r'\n{3,}', '\n\n', md)
    md = re.sub(r'\$\$\s*\n*', '$$\n', md)
    md = re.sub(r'\n*\s*\$\$', '\n$$', md)
    return md

def is_forbidden_code(text):
    """
    변환 금지 키워드로 시작하는지 검사
    (import, #include, #define)
    """
    # 공백 제거 후 시작 단어 확인
    clean_text = text.strip()
    
    # 1. Python import
    if re.match(r'^(import|from)\s+', clean_text):
        return True
    
    # 2. C/C++ Header
    if re.match(r'^#include', clean_text):
        return True
        
    # 3. C/C++ Define
    if re.match(r'^#define', clean_text):
        return True
        
    return False

def beep_sound():
    """시스템 종소리"""
    print('\a')
    sys.stdout.flush()

def main():
    print("=" * 60)
    print("  Gemini → Markdown 변환기 (수동 실행 모드)")
    print("=" * 60)
    print("  y: 현재 클립보드 내용 변환")
    print("  q: 종료")
    print("-" * 60)

    try:
        while True:
            # [처음] 상태: 키 입력 대기
            user_input = input("\n변환 할까요 ?('y') : ").strip().lower()

            # 1. 종료 조건
            if user_input == 'q':
                print("👋 프로그램을 종료합니다.")
                break

            # 2. 변환 시도 조건
            elif user_input == 'y':
                raw = pyperclip.paste()

                # 내용이 없는 경우
                if not raw or len(raw.strip()) == 0:
                    print("⚠️ 클립보드가 비어있습니다.")
                    continue

                # 금지된 코드(import, #include, #define)인지 확인
                if is_forbidden_code(raw):
                    print("🚫 [변환 거부] 코드(import/#include/#define)가 감지되었습니다.")
                    continue
                
                # 모든 조건을 통과했을 때: 종소리 -> 변환
                beep_sound() # 🔔 띵!
                print("🔄 변환 중...")
                
                try:
                    # 디버그 파일 저장 (변환 전 HTML)
                    if DEBUG_MODE:
                        with open('debug_before.html', 'w', encoding='utf-8') as f:
                            f.write(raw)
                    
                    md = process_gemini_html(raw)
                    
                    # 디버그 파일 저장 (변환 후 Markdown)
                    if DEBUG_MODE:
                        with open('debug_after.md', 'w', encoding='utf-8') as f:
                            f.write(md)
                    
                    pyperclip.copy(md)
                    print("✅ 변환 완료! (클립보드 업데이트됨)")
                    if DEBUG_MODE:
                        print("📁 디버그 파일 생성: debug_before.html, debug_after.md")
                except Exception as e:
                    print(f"⚠️ 오류 발생: {e}")

            # y, q 이외의 키는 무시하고 다시 [처음]으로 (while loop)

    except KeyboardInterrupt:
        print("\n\n👋 강제 종료됨")

if __name__ == "__main__":
    main()
