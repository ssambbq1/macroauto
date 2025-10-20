try:
    import streamlit as st
except Exception:
    # This file is intended to be executed via `streamlit run streamlit_frontend.py`.
    if __name__ == "__main__":
        print("Streamlit is required to run this UI. Install with: pip install streamlit")
    raise SystemExit()

import os
import sys
import platform

st.set_page_config(page_title="T&L 매크로 - 상태", layout="centered")
st.title("T&L 자동 복사 매크로 — 상태 페이지")

st.caption("이 페이지는 호스트에서 데스크톱 자동화가 가능한지(예: pyautogui) 여부를 알려줍니다.")

# Try to import the backend module safely. streamlit_app was made import-safe earlier.
backend = None
import_error = None
try:
    import streamlit_app as backend
except Exception as e:
    import_error = e

if backend is None:
    st.error("백엔드 모듈을 가져오지 못했습니다: `streamlit_app`\n앱 로그를 확인하세요.")
    with st.expander("오류 세부정보"):
        st.write(repr(import_error))
    st.info("로컬에서 전체 기능(좌표 설정 등)을 사용하려면 이 저장소를 클론한 뒤 데스크톱 환경에서 실행하세요.")
else:
    # Determine automation availability
    automation_available = getattr(backend, 'AUTOMATION_AVAILABLE', False)
    pyautogui_present = getattr(backend, 'pyautogui', None) is not None

    if automation_available and pyautogui_present:
        st.success("자동화 라이브러리(예: pyautogui)가 사용 가능합니다.")
        st.write("이 호스트에서 마우스/키보드 자동화를 실행할 수 있습니다.")
    else:
        st.error("자동화 라이브러리가 현재 사용 불가합니다 (헤드리스 환경).")
        st.markdown("**원인 및 해결방법**")
        st.markdown("- Streamlit Cloud / 대부분의 클라우드 환경은 GUI 디스플레이(X11/Wayland)가 없어 `pyautogui`를 사용할 수 없습니다.")
        st.markdown("- 해결책: 로컬 Windows/GUI 환경에서 앱을 실행하거나, 자동화를 수행할 별도의 데스크톱 호스트(Windows VM 등)를 마련하세요.")

    with st.expander("호스트 정보"):
        st.write({
            "platform": platform.platform(),
            "python": sys.version.splitlines()[0],
            "os_name": os.name,
            "DISPLAY": os.environ.get('DISPLAY', None),
        })

    st.markdown("---")

    st.header("테스트 및 다음 단계")
    st.write("원하시면 다음 중 하나를 진행하세요:")
    st.write("1. 로컬 데스크톱에서 전체 앱(`python streamlit_app.py`) 실행 — 좌표 설정과 자동화가 동작합니다.")
    st.write("2. 자동화를 수행할 별도의 데스크톱 서버(Windows)에서 간단한 HTTP API를 실행하고, 이 Streamlit UI는 그 API를 호출하도록 구성합니다.")

    if not automation_available:
        with st.expander("원격 자동화 서버 예시 코드 (간단한 안전 경고 포함)"):
            st.code(
"""
# This is a minimal example (NOT production-ready).
from flask import Flask, request
import pyautogui

app = Flask(__name__)

@app.route('/run', methods=['POST'])
def run_action():
    # Add auth/checks here
    action = request.json.get('action')
    if action == 'click_sample':
        pyautogui.click(100, 100)
        return {'status': 'ok'}
    return {'status': 'unknown action'}, 400

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)
""", language='python')

    with st.expander("간단한 로컬 테스트/명령"):
        st.write("아래 명령은 로컬(데스크톱)에서 앱을 직접 실행하는 방법입니다:")
        st.code("python streamlit_app.py")

st.markdown("---")
st.caption("이 페이지는 자동화 기능의 가용성을 확인하기 위한 간단한 상태 페이지입니다.")
