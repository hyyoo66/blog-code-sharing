#!/usr/bin/env python3
import os
import requests
import json

api_key = os.getenv("OPENROUTER_API_KEY")

if not api_key:
    print("❌ OPENROUTER_API_KEY 환경변수가 설정되지 않았습니다")
    exit(1)

print(f"✅ API 키 감지됨 (길이: {len(api_key)}자)")

# 사용 가능한 모델 목록 조회
url = "https://openrouter.ai/api/v1/models"
headers = {
    "Authorization": f"Bearer {api_key}",
}

try:
    response = requests.get(url, headers=headers, timeout=10)
    response.raise_for_status()
    data = response.json()

    print(f"\n✅ OpenRouter 연결 성공!")
    print(f"📊 사용 가능한 모델: {len(data.get('data', []))}개")

    # 이미지 생성 모델만 필터링
    image_models = [m for m in data.get('data', []) if 'image' in m.get('id', '').lower()]
    print(f"🖼️  이미지 생성 모델: {len(image_models)}개")

    if image_models:
        print("\n📋 상위 5개 이미지 모델:")
        for i, model in enumerate(image_models[:5], 1):
            print(f"  {i}. {model['id']}")
            print(f"     - 입력: ${model.get('pricing', {}).get('prompt', 'N/A')}/M 토큰")
            print(f"     - 출력: ${model.get('pricing', {}).get('completion', 'N/A')}/M 토큰")

except requests.exceptions.RequestException as e:
    print(f"❌ API 요청 실패: {e}")
    exit(1)
except json.JSONDecodeError:
    print(f"❌ 응답 파싱 실패")
    exit(1)
