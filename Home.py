# Home.py
import streamlit as st

# [추가] Firebase 로그인 위젯
from streamlit_firebase_auth import FirebaseAuth

# 페이지 설정 (기존)
st.set_page_config(
    page_title="Shopee Support",
    page_icon="🌐",
    layout="wide",
    initial_sidebar_state="expanded",
)

# [추가] Firebase 웹 앱 구성 (콘솔에서 복붙)
firebase_config = {
    "apiKey": "YOUR_API_KEY",
    "authDomain": "YOUR_AUTH_DOMAIN",
    "projectId": "YOUR_PROJECT_ID",
    "storageBucket": "YOUR_STORAGE_BUCKET",
    "messagingSenderId": "YOUR_SENDER_ID",
    "appId": "YOUR_APP_ID",
    # "measurementId": "YOUR_MEASUREMENT_ID",  # 있으면 추가
}

# [추가] 로그인 세션 객체 생성 & 체크
auth = FirebaseAuth(firebase_config)
user = auth.check_session()   # 로그인 상태면 dict, 아니면 None

# [추가] 로그인 요구 화면 + 도메인 제한(원하면 사용)
if not user:
    st.title("로그인이 필요합니다")
    auth.login_form()         # Google 로그인 버튼 렌더링
    st.stop()

email = user.get("email", "")
# 회사 도메인만 허용하려면 아래 조건을 유지, 모든 구글 계정 허용은 if 블록 제거
if not email.endswith("@brand2025.com"):
    st.error("허용되지 않은 계정입니다. @brand2025.com 구글 계정으로 로그인해주세요.")
    auth.logout_form()
    st.stop()

# 사이드바에 로그인 정보/로그아웃 버튼
with st.sidebar:
    st.success(f"로그인: {email}")
    auth.logout_form()

# ==================== 여기부터 '로그인 성공 시' 노출되는 기존 화면 ====================

# (기존) 버튼 스타일 커스텀
st.markdown("""
<style>
:root{
  --btn-bg: #2563EB;
  --btn-bg-hover: #1D4ED8;
  --btn-fg: #FFFFFF;
  --btn-radius: 14px;
  --btn-pad: 18px 24px;
  --btn-shadow: 0 8px 20px rgba(37,99,235,.20);
  --btn-font-size: 30px;
}
.stButton > button{
  background: var(--btn-bg) !important;
  color: var(--btn-fg) !important;
  border: 0 !important;
  border-radius: var(--btn-radius) !important;
  padding: var(--btn-pad) !important;
  font-weight: 700 !important;
  letter-spacing: .2px;
  height: auto !important;
  box-shadow: var(--btn-shadow);
  transition: transform .12s ease, filter .12s ease, background-color .12s ease;
  font-size: var(--btn-font-size) !important;
  line-height: 1.2 !important;
}
.stButton > button:hover{
  background: var(--btn-bg-hover) !important;
  transform: translateY(-1px);
  filter: brightness(1.02);
}
.stButton > button:active{
  transform: translateY(0) scale(.99);
}
@media (min-width: 1200px){
  .stButton > button{ font-size: 22px !important; }
}
</style>
""", unsafe_allow_html=True)

# ===== 헤더 (기존) =====
st.title("🌐 Shopee Support Tools")
st.info(
    "Cover Image : 썸네일로 사용할 커버 이미지를 생성하는 메뉴입니다.\n\n"
    "Copy Template : 샵 복제 시 사용할 Mass Upload 파일을 생성하는 메뉴입니다."
)

st.divider()

# ===== 네비게이션 버튼 (기존) =====
col1, col2 = st.columns(2)

with col1:
    if hasattr(st, "switch_page"):
        if st.button("Cover Image", use_container_width=True, key="btn_cover"):
            st.switch_page("pages/1_Cover Image.py")
    else:
        st.page_link("pages/1_Cover Image.py", label="Cover Image", use_container_width=True)

with col2:
    if hasattr(st, "switch_page"):
        if st.button("Copy Template", use_container_width=True, key="btn_copy"):
            st.switch_page("pages/2_Copy Template.py")
    else:
        st.page_link("pages/2_Copy Template.py", label="Copy Template", use_container_width=True)

# ==================== /기존 화면 끝 ====================
