import re
import pyperclip
from bs4 import BeautifulSoup
import time
import hashlib

def html_to_markdown(html):
    """Gemini HTML → Markdown 구조 변환"""

    soup = BeautifulSoup(html, 'html.parser')

    # 
```
 → ```
    for pre in soup.find_all("pre"):
        code = pre.get_text("\n")
        pre.replace_with(f"\n```\n{code}\n```\n")

    # 
, 
 → **
    for b in soup.find_all(["b", "strong"]):
        b.replace_with(f"**{b.get_text()}**")

    # 
 → *
    for i in soup.find_all("i"):
        i.replace_with(f"*{i.get_text()}*")

    # 
 → \n
    for br in soup.find_all("br"):
        br.replace_with("\n")

    # 
 → 문단
    for p in soup.find_all("p"):
        p.replace_with(p.get_text() + "\n\n")

    # 인라인 수식
    for mjx in soup.find_all("mjx-container"):
        tex = mjx.get_text()
        mjx.replace_with(f"${tex}$")

    # 블록 수식
    for mjx in soup.find_all("mjx-container", {"display": "block"}):
        tex = mjx.get_text()
        mjx.replace_with(f"\n
$$
\n{tex}\n
$$
\n")

    # span 스타일 제거
    for span in soup.find_all("span"):
        span.replace_with(span.get_text())

    text = soup.get_text()
    return text


def fix_mathjax(text):
    """수식 영역 정리"""

    text = text.replace("(", "(").replace(")", ")")

    #
$$
...
$$
포맷 정리
    text = re.sub(r'\$\$\s*\n*', '
$$
\n', text)
    text = re.sub(r'\n*\s*\$\$', '\n
$$
', text)

    return text


def clean_""""""
    return re.sub(r'"]*;?', '', text)


def process_gemini_html(raw_html):
    """Gemini HTML 전체 처리"""

    md = html_to_markdown(raw_html)
    md = clean_"""변경 감지용 해시"""
    if text is None:
        return None
    return hashlib.md5(text.encode('utf-8')).hexdigest()


def main():
    print("=" * 60)
    print("  Gemini → Markdown 자동 변환기 (상주 모드)")
    print("=" * 60)
    print()
    print("📋 클립보드를 감시하고 있습니다...")
    print("💡 Gemini에서 HTML 복사 → 자동으로 Markdown 변환")
    print("⏹️  종료하려면 Ctrl+C")
    print()
    print("-" * 60)

    last_hash = None

    try:
        while True:
            raw = pyperclip.paste()

            if raw:
                current_hash = get_text_hash(raw)

                if current_hash != last_hash and len(raw.strip()) > 5:
                    print(f"\n🔄 [{time.strftime('%H:%M:%S')}] 클립보드 변경 감지!")

                    try:
                        md = process_gemini_html(raw)
                        pyperclip.copy(md)
                        print("✅ 변환 완료! Markdown이 클립보드에 저장되었습니다.")
                    except Exception as e:
                        print(f"⚠️ 변환 중 오류: {e}")

                    last_hash = current_hash

            time.sleep(0.5)

    except KeyboardInterrupt:
        print("\n\n" + "=" * 60)
        print("⏹️  프로그램 종료")
        print("=" * 60)


if __name__ == "__main__":
    main()

```
