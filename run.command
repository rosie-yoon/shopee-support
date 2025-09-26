#!/bin/zsh
cd "$(dirname "$0")"

# 1) 가상환경 있으면 사용
if [ -d ".venv" ]; then
  source .venv/bin/activate
fi

# 2) 없으면 생성 + 의존성 설치
if ! command -v streamlit >/dev/null 2>&1; then
  python3 -m venv .venv
  source .venv/bin/activate
  pip install --upgrade pip
  if [ -f "requirements.txt" ]; then
    pip install -r requirements.txt
  else
    pip install streamlit pillow
  fi
fi

# 3) 실행
exec streamlit run Home.py --server.port=8501 --server.headless true

