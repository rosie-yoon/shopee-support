# pages/2_item_uploader.py  ← 교체용(깔끔 래퍼)
from pathlib import Path
import sys
import streamlit as st

st.set_page_config(page_title="Copy Template", layout="wide")

# 프로젝트 루트(shopee)를 임포트 경로에 추가
ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

# 표준 임포트 (폴더명: item_uploader, 파일명: app.py, 함수: run)
from item_uploader.app import run as item_uploader_run


# 실행
item_uploader_run()
