import re
import time
import win32clipboard
import latex2mathml.converter
import sys

def latex_to_mathml(latex_str):
    """
    LaTeX 문자열을 MathML로 변환합니다.
    변환 실패 시 원본 문자열을 반환합니다.
    """
    try:
        return latex2mathml.converter.convert(latex_str)
    except Exception:
        return latex_str

def process_tables(text):
    """
    텍스트 내의 마크다운 표 문자열을 찾아 HTML Table 태그로 변환합니다.
    """
    lines = text.split('\n')
    new_lines = []
    table_buffer = []
    in_table = False

    for line in lines:
        stripped = line.strip()
        if stripped.startswith('|') and stripped.endswith('|'):
            in_table = True
            table_buffer.append(stripped)
        else:
            if in_table:
                new_lines.append(convert_table_block(table_buffer))
                table_buffer = []
                in_table = False
            new_lines.append(line)
    
    if in_table:
        new_lines.append(convert_table_block(table_buffer))
        
    return '\n'.join(new_lines)

def convert_table_block(lines):
    """
    마크다운 표 라인 리스트를 HTML Table 문자열로 변환합니다.
    """
    if len(lines) < 2:
        return '\n'.join(lines)
    
    if not set(lines[1]).issubset(set('|:- ')):
        return '\n'.join(lines)

    # 표 스타일: 맑은 고딕, 10pt (13pt 이하이므로 유지)
    table_style = "border-collapse: collapse; width: 100%; border: 1px solid black; font-family: 'Malgun Gothic', sans-serif; font-size: 10pt; line-height: 1.1; margin: 0px; mso-para-margin: 0px; font-weight: normal;"
    th_style = "border: 1px solid black; padding: 5px; background-color: #f2f2f2;"
    td_style = "border: 1px solid black; padding: 5px;"

    html = f'<table border="1" cellspacing="0" cellpadding="5" style="{table_style}">'
    
    headers = [h.strip() for h in lines[0].strip('|').split('|')]
    html += '<thead><tr>'
    for h in headers:
        html += f'<th style="{th_style}">{h}</th>'
    html += '</tr></thead>'
    
    html += '<tbody>'
    for line in lines[2:]:
        cells = [c.strip() for c in line.strip('|').split('|')]
        html += '<tr>'
        for i, c in enumerate(cells):
            html += f'<td style="{td_style}">{c}</td>'
        html += '</tr>'
    html += '</tbody></table>' 
    
    return html

def process_inline_markdown(content):
    """
    텍스트 내의 굵은 글씨와 기울임체를 HTML 태그로 변환합니다.
    """
    # 1. 굵은 글씨 (**내용**을 <strong>으로 변환)
    content = re.sub(r'\*\*(.*?)\*\*', r'<strong>\1</strong>', content)

    # 2. 기울임체 (*내용* 또는 _내용_ 을 <em>으로 변환)
    content = re.sub(r'(?<!\*)\*(?!\*)([^\*]+)\*(?!\*)', r'<em>\1</em>', content)
    content = re.sub(r'_([^_]+)_', r'<em>\1</em>', content)
    
    return content

def process_lists_to_text(text):
    """
    리스트 처리 함수 - 리스트 내용에서 볼드/이탤릭도 함께 처리
    """
    lines = text.split('\n')
    new_lines = []
    list_buffer = []
    in_list = False

    # 리스트 폰트: 11pt (13pt 이하이므로 유지)
    li_style = "line-height: 1.1; font-size: 11pt; font-family: 'Malgun Gothic', sans-serif;"
    ul_style = "margin: 0; padding-left: 20px;"

    for line in lines:
        stripped = line.strip().lstrip('\u200b')
        
        if not stripped:
            if in_list:
                # 빈 줄이 나오면 리스트 종료
                new_lines.append(f'<ul style="{ul_style}">{"".join(list_buffer)}</ul>')
                list_buffer = []
                in_list = False
            continue

        # 리스트 항목인지 체크
        match = re.match(r'^([*+\-•·●○▪■◆])\s+(.*)', stripped)
        
        # 가로줄이 아닌 경우만 리스트로 처리
        is_horizontal_rule = re.match(r'^[-*_]{3,}$', stripped)
        
        if match and not is_horizontal_rule:
            in_list = True
            content = match.group(2)
            
            # 숨겨진 공백 제거
            content = content.replace('\u200b', '').replace('\u00a0', ' ')
            
            # 볼드와 이탤릭 처리
            content = process_inline_markdown(content)
            
            list_buffer.append(f'<li style="{li_style}">{content}</li>')
        else:
            if in_list:
                new_lines.append(f'<ul style="{ul_style}">{"".join(list_buffer)}</ul>')
                list_buffer = []
                in_list = False
            new_lines.append(line)
            
    if in_list:
        new_lines.append(f'<ul style="{ul_style}">{"".join(list_buffer)}</ul>')
        
    return '\n'.join(new_lines)

def convert_text_to_html(text):
    """
    텍스트 내의 수식, 마크다운 요소들을 변환하고 HTML을 생성합니다.
    """
    # 1. 블록 수식 처리
    text = re.sub(r'\$\$(.*?)\$\$', lambda m: f'{latex_to_mathml(m.group(1))}', text, flags=re.DOTALL)
    text = re.sub(r'\\\[(.*?)\\\]', lambda m: f'{latex_to_mathml(m.group(1))}', text, flags=re.DOTALL)
    
    # 2. 인라인 수식 처리
    text = re.sub(r'\$(.*?)\$', lambda m: f'{latex_to_mathml(m.group(1))}', text)
    text = re.sub(r'\\\((.*?)\\\)', lambda m: f'{latex_to_mathml(m.group(1))}', text)
    
    # 3. 마크다운 표 처리
    text = process_tables(text)

    # 4. 마크다운 리스트 처리
    text = process_lists_to_text(text)
    
    # 5. 마크다운 인라인 요소 처리 (리스트 밖의 텍스트만 처리)
    lines = text.split('\n')
    processed_lines = []
    for line in lines:
        if re.match(r'^\s*<(ul|table|div|hr)', line, re.IGNORECASE):
            processed_lines.append(line)
        else:
            processed_lines.append(process_inline_markdown(line))
    text = '\n'.join(processed_lines)

    # 6. 가로줄 처리
    text = re.sub(r'^\s*([-*_]){3,}\s*$', r'<hr style="border:none; border-top:1px solid #000000;">', text, flags=re.MULTILINE)

    # 7. 마크다운 헤더 처리
    def header_replace(m):
        level = len(m.group(1))
        content = m.group(2).strip()
        
        # 폰트 크기 계산
        font_size = 18 - (level * 2) 
        if font_size < 12: font_size = 12
        
        # [수정] 폰트 크기 13 이상은 13으로 제한
        if font_size >= 13:
            font_size = 13
            
        return f'<div style="font-size: {font_size}pt; line-height: 1.1; font-weight: bold; color: #000000; font-family: \'Malgun Gothic\', sans-serif;">{content}</div>'
    
    text = re.sub(r'^(#{1,6})\s+(.*)$', header_replace, text, flags=re.MULTILINE)

    # 8. 연속된 텍스트 라인 처리
    lines = text.split('\n')
    final_html_parts = []
    text_buffer = []
    
    # 일반 본문 폰트: 11pt (13pt 이하이므로 유지)
    common_style = "line-height: 1.1; font-size: 11pt; font-family: 'Malgun Gothic', sans-serif; color: #000000; font-weight: normal;"

    def flush_buffer():
        if text_buffer:
            joined = '<br>'.join(text_buffer)
            final_html_parts.append(f'<p style="{common_style}">{joined}</p>')
            text_buffer.clear()

    for line in lines:
        stripped = line.strip()
        
        if not stripped:
            continue
        
        if re.match(r'^\s*<(table|hr|ul|ol|div)', line, re.IGNORECASE):
            flush_buffer()
            final_html_parts.append(line)
        else:
            text_buffer.append(line)
            
    flush_buffer()

    final_body_content = ''.join(final_html_parts)
    
    # body 기본 폰트도 11pt 설정
    html_body = f'<html><body style="font-weight: normal; font-family: \'Malgun Gothic\', sans-serif; font-size: 11pt;">{final_body_content}</body></html>'
    return html_body

def copy_html_to_clipboard(html):
    """
    생성된 HTML을 윈도우 클립보드 포맷에 맞춰 복사합니다.
    """
    header = (
        "Version:0.9\r\n"
        "StartHTML:{0:08d}\r\n"
        "EndHTML:{1:08d}\r\n"
        "StartFragment:{2:08d}\r\n"
        "EndFragment:{3:08d}\r\n"
    )
    
    html_bytes = html.encode('utf-8')
    fragment_start_marker = "<html><body><!--StartFragment-->"
    fragment_end_marker = "<!--EndFragment--></body></html>"
    
    start_html = len(header.format(0, 0, 0, 0))
    start_fragment = start_html + len(fragment_start_marker)
    end_fragment = start_fragment + len(html_bytes)
    end_html = end_fragment + len(fragment_end_marker)
    
    formatted_html = (
        header.format(start_html, end_html, start_fragment, end_fragment)
        + fragment_start_marker
    )
    final_payload = formatted_html.encode('utf-8') + html_bytes + fragment_end_marker.encode('utf-8')

    win32clipboard.OpenClipboard()
    win32clipboard.EmptyClipboard()
    cf_html = win32clipboard.RegisterClipboardFormat("HTML Format")
    win32clipboard.SetClipboardData(cf_html, final_payload)
    win32clipboard.CloseClipboard()

def get_clipboard_text():
    """
    클립보드에서 텍스트를 가져옵니다.
    """
    try:
        win32clipboard.OpenClipboard()
        if win32clipboard.IsClipboardFormatAvailable(win32clipboard.CF_UNICODETEXT):
            text = win32clipboard.GetClipboardData(win32clipboard.CF_UNICODETEXT)
        else:
            text = None
        win32clipboard.CloseClipboard()
        return text
    except:
        return None

def is_forbidden_code(text):
    """
    변환 금지 키워드로 시작하는지 검사
    (import, #include, #define)
    """
    clean_text = text.strip()
    
    if re.match(r'^(import|from)\s+', clean_text):
        return True
    
    if re.match(r'^#include', clean_text):
        return True
        
    if re.match(r'^#define', clean_text):
        return True
        
    return False

def beep_sound():
    """시스템 종소리"""
    print('\a')
    sys.stdout.flush()

def main():
    print("=" * 60)
    print("  Gemini → Word HTML 변환기 (수동 실행 모드)")
    print("=" * 60)
    print("  y: 현재 클립보드 내용 변환")
    print("  q: 종료")
    print("-" * 60)

    try:
        while True:
            # 키 입력 대기
            user_input = input("\n변환 할까요 ?('y') : ").strip().lower()

            # 1. 종료 조건
            if user_input == 'q':
                print("👋 프로그램을 종료합니다.")
                break

            # 2. 변환 시도 조건
            elif user_input == 'y':
                current_text = get_clipboard_text()

                # 내용이 없는 경우
                if not current_text or len(current_text.strip()) == 0:
                    print("⚠️ 클립보드가 비어있습니다.")
                    continue

                # 금지된 코드(import, #include, #define)인지 확인
                if is_forbidden_code(current_text):
                    print("🚫 [변환 거부] 코드(import/#include/#define)가 감지되었습니다.")
                    continue
                
                # 변환 진행
                beep_sound() # 🔔 띵!
                print("🔄 변환 중...")
                
                try:
                    # HTML 변환
                    html_result = convert_text_to_html(current_text)
                    
                    # 클립보드에 다시 복사
                    copy_html_to_clipboard(html_result)
                    
                    print("✅ 변환 완료! 워드에 바로 붙여넣을 수 있습니다.")
                except Exception as e:
                    print(f"⚠️ 오류 발생: {e}")

    except KeyboardInterrupt:
        print("\n\n👋 강제 종료됨")

if __name__ == "__main__":
    main()
