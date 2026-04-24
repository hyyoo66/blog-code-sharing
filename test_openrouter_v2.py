#!/usr/bin/env python3
import os
import requests
import json

api_key = "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"

print("🔄 OpenRouter API 상태 확인 (v2)...\n")

# 1. 간단한 텍스트 모델 테스트
url = "https://openrouter.ai/api/v1/chat/completions"

headers = {
    "Authorization": f"Bearer {api_key}",
    "Content-Type": "application/json",
}

payload = {
    "model": "openai/gpt-3.5-turbo",
    "messages": [{"role": "user", "content": "Say 'OpenRouter API works!'"}],
    "temperature": 0.7,
}

try:
    print(f"📤 텍스트 모델 테스트: openai/gpt-3.5-turbo\n")

    response = requests.post(url, headers=headers, json=payload, timeout=30)

    print(f"📊 상태 코드: {response.status_code}")

    if response.status_code == 200:
        data = response.json()
        print(f"✅ OpenRouter 연결 성공!")
        print(f"💬 응답: {data['choices'][0]['message']['content']}")
        print(f"\n📋 사용 토큰:")
        print(f"   - 입력: {data.get('usage', {}).get('prompt_tokens', 0)}")
        print(f"   - 출력: {data.get('usage', {}).get('completion_tokens', 0)}")

    else:
        print(f"❌ 오류: {response.status_code}")
        print(f"📋 응답: {response.text[:500]}")

except Exception as e:
    print(f"❌ 실패: {e}")

# 2. 이미지 모델 목록 확인
print("\n" + "="*50)
print("\n🖼️  이미지 모델 확인...\n")

image_models = [
    "openai/dall-e-3",
    "openai/dall-e-2",
    "google/imagen-3",
    "google/imagen-3-fast"
]

for model in image_models:
    test_payload = {
        "model": model,
        "prompt": "test",
    }
    try:
        response = requests.post(
            "https://openrouter.ai/api/v1/images/generations",
            headers=headers,
            json=test_payload,
            timeout=10
        )
        status = "✅" if response.status_code in [200, 400] else "❌"
        print(f"{status} {model}: {response.status_code}")
    except Exception as e:
        print(f"❌ {model}: {type(e).__name__}")
