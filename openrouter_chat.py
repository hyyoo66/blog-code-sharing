import os
import sys
from datetime import datetime
from openai import OpenAI

# -----------------------------------------------
# 🔧 설정
# -----------------------------------------------
MODEL = "deepseek/deepseek-chat"
API_KEY_ENV = "OPENROUTER_API_KEY"
SEPARATOR_LENGTH = 50
# -----------------------------------------------


def create_client():
    api_key = os.environ.get(API_KEY_ENV)
    if not api_key:
        print(f"⚠️ 환경변수 {API_KEY_ENV} 가 설정되지 않았습니다.")
        print(f"   예) export {API_KEY_ENV}=sk-or-...")
        sys.exit(1)
    return OpenAI(
        api_key=api_key,
        base_url="https://openrouter.ai/api/v1",
    )


def chat(client, messages):
    response = client.chat.completions.create(
        model=MODEL,
        messages=messages,
        max_tokens=1000,
    )
    return response.choices[0].message.content


def main():
    script_name = os.path.basename(__file__)
    file_path = __file__
    if os.path.exists(file_path):
        timestamp = os.path.getmtime(file_path)
        mod_time = datetime.fromtimestamp(timestamp).strftime('%Y-%m-%d %H:%M:%S')
    else:
        mod_time = "Unknown"

    separator = "=" * SEPARATOR_LENGTH
    dash = "-" * SEPARATOR_LENGTH

    print(separator)
    print(f"  파일: {script_name}")
    print(f"  수정일: {mod_time}")
    print(f"  모델: {MODEL}")
    print(separator)
    print("  q / exit : 종료")
    print("  clear    : 대화 초기화")
    print(dash)

    client = create_client()
    messages = []

    while True:
        try:
            user_input = input("\n나: ").strip()
        except (KeyboardInterrupt, EOFError):
            print("\n\n👋 종료합니다.")
            break

        if not user_input:
            continue

        if user_input.lower() in ("q", "exit", "종료"):
            print("👋 종료합니다.")
            break

        if user_input.lower() == "clear":
            messages = []
            print("🗑️ 대화 내용이 초기화되었습니다.")
            continue

        messages.append({"role": "user", "content": user_input})

        print("🔄 생각 중...")
        try:
            reply = chat(client, messages)
            messages.append({"role": "assistant", "content": reply})
            print(f"\nAI: {reply}")
        except Exception as e:
            messages.pop()
            print(f"⚠️ 오류 발생: {e}")


if __name__ == "__main__":
    main()
