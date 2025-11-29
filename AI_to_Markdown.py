import re
import pyperclip
from bs4 import BeautifulSoup, NavigableString
import sys
import os
from datetime import datetime

# 디버그 파일 생성 여부 설정
DEBUG_MODE = False

# 마크다운 제목(헤더) 고정 개수 설정
# 예: 3으로 설정 시 -> #, ##, ###### 모두 ### 로 변경됨 (개수 통일)
UNIFIED_HEADER_COUNT = 3

# 화면 구분선 길이 설정
SEPARATOR_LENGTH = 50

def html_to_markdown(html):
    """Gemini HTML → Markdown 변환"""
    
    # 백틱 안의 HTML 태그 보호
    def protect_backtick_tags(text):
        protected = []
        def replacer(match):
            protected.append(match.group(0))
            return f"___BACKTICK_{len(protected)-1}___"
        text = re.sub(r'`[^`]+`', replacer, text)
        return text, protected
    
    html, protected_backticks = protect_backtick_tags(html)
    
    soup = BeautifulSoup(html, 'html.parser')

    # 0. 스타일 제거
    for tag in soup.find_all(True):
        if tag.has_attr('style'): del tag['style']
        if tag.has_attr('class'): del tag['class']

    # 1. 코드 블록(pre, code) 처리
    for pre in soup.find_all("pre"):
        code_text = pre.get_text("\n")
        if '\n' in code_text.strip() or len(code_text) > 50:
            pre.replace_with(f"\n```\n{code_text}\n```\n")
        else:
            pre.replace_with(f"`{code_text.strip()}`")

    for code in soup.find_all("code"):
        if code.parent.name == 'pre': continue
        text = code.get_text()
        code.replace_with(f"`{text}`")

    # 2. 제목 변환 (모든 제목을 고정된 개수로 통일)
    header_symbol = '#' * UNIFIED_HEADER_COUNT  # 미리 계산

    for i in range(1, 7):
        for h in soup.find_all(f"h{i}"):
            is_inline_mention = False
            if h.parent.name in ['p', 'span', 'li', 'a']:
                is_inline_mention = True
            prev = h.previous_sibling
            if prev and isinstance(prev, NavigableString) and len(prev.strip()) > 0:
                is_inline_mention = True

            if is_inline_mention:
                h.replace_with(f"`<{h.name}>{h.get_text()}</{h.name}>`")
            else:
                header_text = h.get_text().strip()
                if header_text:
                    # 무조건 설정된 개수(header_symbol)로 변경
                    h.replace_with(f"\n\n{header_symbol} {header_text}\n\n")

    # 3. 서식 변환
    for b in soup.find_all(["b", "strong"]):
        b.replace_with(f"**{b.get_text()}**")
    for i in soup.find_all(["i", "em"]):
        i.replace_with(f"*{i.get_text()}*")

    # 4. 수식 보호
    for mjx in soup.find_all("mjx-container"):
        tex = mjx.get_text().strip()
        if mjx.get("display") == "block":
            mjx.replace_with(f"\n$$\n{tex}\n$$\n")
        else:
            mjx.replace_with(f"${tex}$")

    # 5. 리스트
    for li in soup.find_all("li"):
        li.insert_before("* ")
        li.append("\n")
        li.unwrap()
    for ul in soup.find_all(["ul", "ol"]):
        ul.insert_before("\n")
        ul.append("\n")
        ul.unwrap()

    # 6. 문단 및 기타
    for p in soup.find_all(["p", "div"]):
        p.append("\n\n")
        p.unwrap()
    for br in soup.find_all("br"):
        br.replace_with("\n")
    for span in soup.find_all("span"):
        span.unwrap()

    result = soup.get_text()
    
    for i, backtick in enumerate(protected_backticks):
        result = result.replace(f"___BACKTICK_{i}___", backtick)
    
    return result

def is_html(text):
    return bool(re.search(r'<[a-zA-Z][^>]*>', text))

def insert_tilde_in_hashes(text):
    """[안전 장치] # -> #~ 변환 (코드 블록 제외)"""
    lines = text.split('\n')
    processed_lines = []
    
    for line in lines:
        header_match = re.match(r'^(#{1,6}\s+)', line)
        if header_match:
            header_part = header_match.group(1)
            content_part = line[len(header_part):]
            content_part = content_part.replace('#', '#~')
            processed_lines.append(header_part + content_part)
        else:
            processed_lines.append(line.replace('#', '#'))
    
    return '\n'.join(processed_lines)

def remove_hr_lines(text):
    lines = text.split('\n')
    filtered_lines = []
    for line in lines:
        stripped = line.strip()
        if stripped and re.match(r'^-{3,}$', stripped):
            continue
        filtered_lines.append(line)
    return '\n'.join(filtered_lines)

def process_gemini_html(raw_input):
    if is_html(raw_input):
        md = html_to_markdown(raw_input)
        md = re.sub(r'background[^;"]*;?', '', md)
    else:
        md = raw_input
    
    md = remove_hr_lines(md)
    
    # 코드 블록 보호용 정규식 (줄 시작 부분의 ```만 인식)
    pattern = r'(?m)(^\s*```[\s\S]*?^\s*```)'
    parts = re.split(pattern, md)
    
    final_parts = []
    
    def resize_header_in_text(match):
        # 원본 개수 무시하고 설정값(UNIFIED_HEADER_COUNT)으로 고정
        return ('#' * UNIFIED_HEADER_COUNT) + ' '

    for part in parts:
        if re.match(r'^\s*```', part):
            final_parts.append(part)
        else:
            # 제목 개수 강제 통일
            part = re.sub(r'^\s*(#{1,6})\s+', resize_header_in_text, part, flags=re.MULTILINE)
            part = insert_tilde_in_hashes(part)
            final_parts.append(part)
    
    md = "".join(final_parts)
    
    # 줄바꿈 정리
    md = re.sub(r'\n{3,}', '\n\n', md)
    md = re.sub(r'\$\$\s*\n*', '$$\n', md)
    md = re.sub(r'\n*\s*\$\$', '\n$$', md)
    
    return md

def is_forbidden_code(text):
    clean_text = text.strip()
    if re.match(r'^(import|from)\s+', clean_text): return True
    if re.match(r'^#include', clean_text): return True
    if re.match(r'^#define', clean_text): return True
    return False

def beep_sound():
    print('\a')
    sys.stdout.flush()

def save_backup(content):
    """클립보드 내용 백업 파일 생성"""
    try:
        now_str = datetime.now().strftime('%y%m%d_%H%M%S')
        filename = f"clipboard backup_{now_str}.txt"
        
        with open(filename, 'w', encoding='utf-8') as f:
            f.write(content)
        print(f"📁 [백업 완료] {filename}")
    except Exception as e:
        print(f"⚠️ [백업 실패] {e}")

def main():
    script_name = os.path.basename(__file__)
    file_path = __file__
    if os.path.exists(file_path):
        timestamp = os.path.getmtime(file_path)
        mod_time = datetime.fromtimestamp(timestamp).strftime('%Y-%m-%d %H:%M:%S')
    else:
        mod_time = "Unknown"

    file_line = f"  파일: {script_name}"
    date_line = f"  수정일: {mod_time}"
    y_line = "  y: 현재 클립보드 내용 변환"
    q_line = "  q: 종료"
    
    separator = "=" * SEPARATOR_LENGTH
    dash_separator = "-" * SEPARATOR_LENGTH

    print(separator)
    print(file_line)
    print(date_line)
    print(separator)
    print(y_line)
    print(q_line)
    print(dash_separator)

    try:
        while True:
            user_input = input("\n변환 할까요 ?('y') : ").strip().lower()

            if user_input == 'q':
                print("👋 프로그램을 종료합니다.")
                break

            elif user_input == 'y':
                raw = pyperclip.paste()

                if not raw or len(raw.strip()) == 0:
                    print("⚠️ 클립보드가 비어있습니다.")
                    continue

                if is_forbidden_code(raw):
                    print("🚫 [변환 거부] 코드(import/#include/#define)가 감지되었습니다.")
                    continue
                
                # 백업 실행
                save_backup(raw)
                
                beep_sound()
                print("🔄 변환 중...")
                
                try:
                    if DEBUG_MODE:
                        with open('debug_before.html', 'w', encoding='utf-8') as f:
                            f.write(raw)
                    
                    md = process_gemini_html(raw)
                    
                    if DEBUG_MODE:
                        with open('debug_after.md', 'w', encoding='utf-8') as f:
                            f.write(md)
                    
                    pyperclip.copy(md)
                    print("✅ 변환 완료! (클립보드 업데이트됨)")
                except Exception as e:
                    print(f"⚠️ 오류 발생: {e}")

    except KeyboardInterrupt:
        print("\n\n👋 강제 종료됨")

if __name__ == "__main__":
    main()
