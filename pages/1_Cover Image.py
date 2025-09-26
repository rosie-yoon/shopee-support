import streamlit as st
from pathlib import Path
import sys

# pages/ 아래에 있으므로 프로젝트 루트(shopee)를 sys.path에 추가 (견고성 ↑)
ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.append(str(ROOT))

from image_compose.app import run as image_compose_run  # 폴더명이 image_compose 여야 함

st.set_page_config(page_title="Cover Image", layout="wide")

image_compose_run()
