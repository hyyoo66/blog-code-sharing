'''
Gemini에서 글 복사
2) 파이썬 스크립트 실행

→ 클립보드에 완성된 Markdown 생성
→ 티스토리 마크다운에 그대로 붙여넣기
'''
import re
import pyperclip
from bs4 import BeautifulSoup
import time

def html_to_markdown(html):
    """Gemini가 복사한 HTML을 실제 Markdown 구조로 완전히 변환"""

    soup = BeautifulSoup(html, 'html.parser')

    # 1) 코드블록 변환 <pre><code> → ``` 
    for pre in soup.find_all("pre"):
        code = pre.get_text("\n")
        pre.replace_with(f"\n```\n{code}\n```\n")

    # 2) 굵은 글씨 <b>, <strong> → **텍스트**
    for b in soup.find_all(["b", "strong"]):
        b.replace_with(f"**{b.get_text()}**")

    # 3) 이탤릭 <i> → *
    for i in soup.find_all("i"):
        i.replace_with(f"*{i.get_text()}*")

    # 4) 줄바꿈 <br> → \n
    for br in soup.find_all("br"):
        br.replace_with("\n")

    # 5) p 태그는 Markdown 문단으로 변환
    for p in soup.find_all("p"):
        p.replace_with(p.get_text() + "\n\n")

    # 6) 인라인 수식 <mjx-container> → $...$
    for mjx in soup.find_all("mjx-container"):
        tex = mjx.get_text()
        mjx.replace_with(f"${tex}$")

    # 7) 블록 수식 <mjx-container display="block"> → $$...$$
    for mjx in soup.find_all("mjx-container", {"display": "block"}):
        tex = mjx.get_text()
        mjx.replace_with(f"\n$$\n{tex}\n$$\n")

    # 8) 모든 span의 스타일 제거(배경 포함)
    for span in soup.find_all("span"):
        span.replace_with(span.get_text())

    # 최종 텍스트
    text = soup.get_text()

    return text


def fix_mathjax(text):
    """Markdown 수식을 티스토리 MathJax에 맞게 정리"""
    # 백슬래시 깨짐 방지
    text = text.replace("\\(", "(").replace("\\)", ")")

    # $$ ... $$ 사이 공백 정리
    text = re.sub(r'\$\$\s*\n*', '$$\n', text)
    text = re.sub(r'\n*\s*\$\$', '\n$$', text)

    return text


def clean_background(text):
    # background 제거 정리
    return re.sub(r'background[^;"]*;?', '', text)


def process_gemini_html(raw_html):
    md = html_to_markdown(raw_html)
    md = clean_background(md)
    md = fix_mathjax(md)

    # 불필요한 빈줄 정리
    md = re.sub(r'\n{3,}', '\n\n', md)

    return md


if __name__ == "__main__":
    print("📌 Gemini → Markdown 자동 변환 중...\n")

    raw = pyperclip.paste()
    processed = process_gemini_html(raw)

    pyperclip.copy(processed)

    print("✨ 완료! 변환된 Markdown이 클립보드에 저장되었습니다.")
    print(f"\n3초 후 종료합니다.", end="")
    time.sleep(3)
