#!/usr/bin/env python3
import os
import requests
import json

api_key = "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"

print("🔄 OpenRouter 이미지 생성 테스트 시작...")
print(f"✅ API 키 감지됨 (길이: {len(api_key)}자)\n")

# OpenRouter 이미지 생성 API
url = "https://openrouter.ai/api/v1/images/generations"

headers = {
    "Authorization": f"Bearer {api_key}",
    "Content-Type": "application/json",
}

payload = {
    "model": "openai/dall-e-3",  # 기본 모델
    "prompt": "A beautiful sunset over mountains",
    "size": "1024x1024",
    "n": 1,
}

try:
    print(f"📤 요청: {payload['model']}")
    print(f"📝 프롬프트: {payload['prompt']}\n")

    response = requests.post(url, headers=headers, json=payload, timeout=30)

    print(f"📊 상태 코드: {response.status_code}")

    if response.status_code == 200:
        data = response.json()
        print(f"✅ 이미지 생성 성공!")
        print(f"📋 응답: {json.dumps(data, indent=2)}")
    else:
        print(f"❌ 오류: {response.status_code}")
        print(f"📋 응답: {response.text}")

except requests.exceptions.RequestException as e:
    print(f"❌ 요청 실패: {e}")
except json.JSONDecodeError:
    print(f"❌ 응답 파싱 실패")
