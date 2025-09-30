# -*- coding: utf-8 -*-
from __future__ import annotations
import streamlit as st
from gspread.exceptions import WorksheetNotFound
from .utils_common import load_env, open_sheet_by_env, open_ref_by_env, safe_worksheet, with_retry

# 통합된 automation_steps 하나만 import 합니다.
from . import automation_steps

class ShopeeAutomation:
    """
    Shopee 자동화의 모든 단계를 제어하는 컨트롤러 클래스.
    Streamlit UI와 실제 로직 사이의 다리 역할을 합니다.
    """
    def __init__(self):
        # --- 디버깅 로그 추가 ---
        print("[ ShopeeAutomation ] 클래스 초기화를 시작합니다.")
        try:
            load_env()
            print("[ ShopeeAutomation ] Google Sheets 연결을 시도합니다...")
            self.sh = open_sheet_by_env()
            self.ref = open_ref_by_env()
            print("[ ShopeeAutomation ] Google Sheets 연결 성공 및 클래스 초기화 완료.")
        except Exception as e:
            print(f"[ ShopeeAutomation ] 초기화 중 심각한 오류 발생: {e}") # 터미널 로그용
            st.error(f"Google Sheets 연결에 실패했습니다: {e}")
            st.stop()

    def _initialize_failures_sheet(self):
        """(신규) Failures 시트를 찾아 초기화하거나 새로 생성합니다."""
        try:
            failures_ws = safe_worksheet(self.sh, "Failures")
            with_retry(lambda: failures_ws.clear())
        except WorksheetNotFound:
            failures_ws = with_retry(lambda: self.sh.add_worksheet(title="Failures", rows=1000, cols=10))
        
        # 헤더 다시 작성
        header = [["PID","Category","Name","Reason","Detail"]]
        with_retry(lambda: failures_ws.update(values=header, range_name="A1:E1"))
        print("[ INFO ] Failures sheet has been initialized.")


    def run_all_steps_with_progress(self, progress_bar, log_container, shop_code: str):
        """
        (수정) Streamlit UI에 진행 상황을 표시하며 모든 자동화 단계를 순차적으로 실행합니다.
        - `Failures` 시트 초기화 로직 추가
        - `shop_code`를 Step 6에 전달
        """
        
        # 1. (추가) 자동화 시작 직전에 Failures 시트 초기화
        try:
            self._initialize_failures_sheet()
        except Exception as e:
            st.error(f"Failures 시트 초기화 실패: {e}")
            return False, ["❌ Failures 시트 초기화 : **실패**"]

        # 2. 실행할 단계 목록 정의
        steps = [
            ("Step 1: TEM_OUTPUT 시트 생성", self.run_step1_build_template),
            ("Step 2: Mandatory 기본값 채우기", self.run_step2_fill_defaults),
            ("Step 3: FDA 코드 채우기", self.run_step3_fill_fda),
            ("Step 4: 기타 필드(재고, 브랜드 등) 채우기", self.run_step4_fill_etc),
            ("Step 5: 필수 정보(설명, 가격 등) 채우기", self.run_step5_fill_info),
            ("Step 6: 커버 이미지 URL 생성", lambda: self.run_step6_fill_images(shop_code)),
        ]

        total_steps = len(steps)
        results = []
        all_success = True

        for i, (title, func) in enumerate(steps):
            progress_text = f"({i+1}/{total_steps}) {title} 진행 중..."
            progress_bar.progress((i + 1) / total_steps, text=progress_text)
            
            try:
                func()
                log_container.markdown(f"✅ **{title}** : 완료")
                results.append(f"✅ {title} : **성공**")
            except Exception as e:
                error_message = f"❌ **{title}** : 실패\n   - 오류: `{e}`"
                log_container.error(error_message)
                results.append(error_message)
                all_success = False
                break 

        progress_bar.empty()
        return all_success, results

    def run_step1_build_template(self):
        # --- 디버깅 로그 추가 ---
        print("[ Step 1 ] 'run_step1_build_template' 함수가 호출되었습니다.")
        try:
            print("[ Step 1 ] 'automation_steps.run_step_1' 실행 직전입니다.")
            automation_steps.run_step_1(self.sh, self.ref)
            print("[ Step 1 ] 'automation_steps.run_step_1' 실행이 성공적으로 완료되었습니다.")
        except Exception as e:
            print(f"[ Step 1 ] 'automation_steps.run_step_1' 실행 중 심각한 오류가 발생했습니다: {e}")
            # 발생한 오류를 다시 상위로 보내 UI에 표시되도록 합니다.
            raise e

    def run_step2_fill_defaults(self):
        automation_steps.run_step_2(self.sh, self.ref)

    def run_step3_fill_fda(self):
        # (수정) 인자 변경
        automation_steps.run_step_3(self.sh, self.ref, overwrite=True)

    def run_step4_fill_etc(self):
        automation_steps.run_step_4(self.sh, self.ref)

    def run_step5_fill_info(self):
        automation_steps.run_step_5(self.sh)
    
    def run_step6_fill_images(self, shop_code: str):
        # (수정) shop_code 인자 전달
        automation_steps.run_step_6(self.sh, shop_code)

    def run_step7_generate_download(self):
        return automation_steps.run_step_7(self.sh)