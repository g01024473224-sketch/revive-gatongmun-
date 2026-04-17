"""
리바이브 수학 가통문 자동 생성기 (Streamlit 웹 앱)

워크플로우:
1. 연도/월 선택
2. 매쓰플랫 Excel 파일 업로드 (학년별 다중)
3. 진도계획표 Google Sheet URL (또는 로컬 xlsx)
4. 공지사항 입력
5. [가통문 생성] 클릭
6. AI 코멘트 생성 진행률 확인
7. 완성된 xlsx 다운로드
"""

from __future__ import annotations

import json
import os
import shutil
import sys
from pathlib import Path

import streamlit as st
import streamlit_authenticator as stauth
import yaml
from dotenv import load_dotenv

# --- 스크립트 경로 등록 ---
BASE = Path(__file__).parent
SCRIPTS = BASE / "scripts"
OUTPUT = BASE / "output"
UPLOAD = BASE / "_upload"
OUTPUT.mkdir(exist_ok=True)
UPLOAD.mkdir(exist_ok=True)
sys.path.insert(0, str(SCRIPTS))

load_dotenv(BASE / ".env")

# 모듈 import (스크립트 재사용)
from parse_mathflat import parse_mathflat_excel
from parse_progress import parse_progress_workbook
from merge_student_data import load_mathflat_all, load_progress, build_student_records

# --- 페이지 설정 ---
st.set_page_config(
    page_title="리바이브 가통문 자동 생성기",
    page_icon="📝",
    layout="wide",
)

# --- 로그인 인증 ---
config_path = BASE / "config.yaml"
with config_path.open(encoding="utf-8") as f:
    config = yaml.safe_load(f)

authenticator = stauth.Authenticate(
    config["credentials"],
    config["cookie"]["name"],
    config["cookie"]["key"],
    config["cookie"]["expiry_days"],
)

authenticator.login()

if st.session_state.get("authentication_status") is None:
    st.info("🔐 로그인이 필요합니다. 아이디와 비밀번호를 입력하세요.")
    st.stop()
elif st.session_state.get("authentication_status") is False:
    st.error("❌ 아이디 또는 비밀번호가 올바르지 않습니다.")
    st.stop()

# --- 로그인 성공 후 ---
# --- 사이드바: 기본 정보 ---
st.sidebar.title("📝 리바이브 가통문")
st.sidebar.markdown("매쓰플랫 + 진도계획표 → AI 코멘트 → 가통문 xlsx")
st.sidebar.divider()
st.sidebar.write(f"👤 {st.session_state.get('name', '')} 님")
authenticator.logout("로그아웃", "sidebar")
st.sidebar.divider()

st.sidebar.subheader("📅 대상 월")
col_y, col_m = st.sidebar.columns(2)
with col_y:
    year = st.number_input("연도", min_value=2024, max_value=2030, value=2026, step=1)
with col_m:
    month = st.number_input("월", min_value=1, max_value=12, value=4, step=1)

# API 키 상태 확인
api_status = "✅" if os.getenv("ANTHROPIC_API_KEY") else "❌"
st.sidebar.caption(f"Anthropic API 키: {api_status}")

# --- 메인 영역 ---
st.title("🧮 가통문 자동 생성기")

st.warning("⚠️ **꼭 인쇄는 '한 페이지에 모든 열 맞추기' 선택하기!!!**", icon="🖨️")

# ---------- STEP 1: 매쓰플랫 Excel 업로드 ----------
st.header("1️⃣ 매쓰플랫 Excel 업로드")
st.caption(
    "매쓰플랫에서 다운로드한 학년별 학습지 학습내역 Excel 파일을 모두 업로드하세요. "
    "파일명 예시: `이상준선생님_중1_2026년 03월_학습지 학습내역.xlsx`"
)
uploaded_files = st.file_uploader(
    "Excel 파일 (여러 개 선택 가능)",
    type=["xlsx"],
    accept_multiple_files=True,
)

# ---------- STEP 2: 진도계획표 ----------
st.header("2️⃣ 진도계획표 (Google Sheet)")
progress_source = st.radio(
    "진도계획표 가져오기 방식",
    ["공개 URL에서 자동 다운로드", "로컬 xlsx 업로드"],
    horizontal=True,
)

progress_url = ""
progress_file = None
if progress_source == "공개 URL에서 자동 다운로드":
    progress_url = st.text_input(
        "Google Sheet URL",
        value="https://docs.google.com/spreadsheets/d/13byzXWxaPeqqoVfrINvurKxq1GehXhXp_Q643M-muWA/edit",
        help="시트가 '링크가 있는 모든 사용자' 또는 공개 권한이어야 함",
    )
else:
    progress_file = st.file_uploader("진도계획표 xlsx", type=["xlsx"], key="progress_file")

# ---------- STEP 3: 공지사항 ----------
st.header("3️⃣ 공지사항")
st.caption("학원 휴일, 시험 일정, 특강 공지 등. 모든 학생의 가통문에 동일하게 표시됩니다.")
notice_path = OUTPUT / f"notice_{int(year)}_{int(month):02d}.txt"
default_notice = notice_path.read_text(encoding="utf-8") if notice_path.exists() else (
    "● \n"
    "● \n"
    "● \n"
)
notice_text = st.text_area(
    "공지사항 내용",
    value=default_notice,
    height=140,
    help="각 줄 시작에 '●' 기호를 쓰면 깔끔하게 표시됩니다.",
)

# ---------- STEP 4: AI 코멘트 설정 ----------
st.header("4️⃣ AI 코멘트 설정")
col1, col2 = st.columns(2)
with col1:
    skip_ai = st.checkbox("AI 코멘트 생략 (빠른 테스트용)", value=False)
with col2:
    force_regen = st.checkbox("기존 코멘트 무시하고 재생성", value=False)

# ---------- STEP 5: 실행 버튼 ----------
st.header("5️⃣ 가통문 생성")

if st.button("🚀 가통문 생성 시작", type="primary", use_container_width=True):
    if not uploaded_files:
        st.error("매쓰플랫 Excel 파일을 하나 이상 업로드하세요.")
        st.stop()

    progress_bar = st.progress(0.0)
    status = st.empty()

    try:
        # === 업로드 파일 저장 ===
        status.info("📁 업로드 파일 저장 중...")
        # 기존 업로드 정리
        for f in UPLOAD.glob("*.xlsx"):
            f.unlink()
        saved_files = []
        for uf in uploaded_files:
            p = UPLOAD / uf.name
            with p.open("wb") as f:
                f.write(uf.getbuffer())
            saved_files.append(p)
        progress_bar.progress(0.1)

        # === 매쓰플랫 파싱 ===
        status.info(f"📊 매쓰플랫 Excel 파싱 중... ({len(saved_files)}개)")
        # output/*_parsed.json 정리 후 다시 생성
        for jf in OUTPUT.glob("*_parsed.json"):
            jf.unlink()
        import re
        for xlsx in saved_files:
            data = parse_mathflat_excel(xlsx)
            m = re.search(r"(초\d|중\d)", xlsx.name)
            key = m.group(1) if m else xlsx.stem
            out_path = OUTPUT / f"{key}_parsed.json"
            with out_path.open("w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2, default=str)
        progress_bar.progress(0.25)

        # === 진도계획표 로드 ===
        status.info("📅 진도계획표 가져오는 중...")
        if progress_source == "공개 URL에서 자동 다운로드":
            import urllib.request
            import re as re_mod
            m = re_mod.search(r"/d/([^/]+)", progress_url)
            if not m:
                st.error("URL에서 Sheet ID를 추출할 수 없습니다.")
                st.stop()
            sheet_id = m.group(1)
            export_url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=xlsx"
            progress_xlsx_path = OUTPUT / "진도계획표_전체.xlsx"
            urllib.request.urlretrieve(export_url, progress_xlsx_path)
        else:
            if not progress_file:
                st.error("진도계획표 파일을 업로드하세요.")
                st.stop()
            progress_xlsx_path = OUTPUT / "진도계획표_전체.xlsx"
            with progress_xlsx_path.open("wb") as f:
                f.write(progress_file.getbuffer())

        progress_data = parse_progress_workbook(progress_xlsx_path, int(year), int(month))
        progress_json = OUTPUT / f"progress_{int(year)}_{int(month):02d}.json"
        with progress_json.open("w", encoding="utf-8") as f:
            json.dump(progress_data, f, ensure_ascii=False, indent=2)
        progress_bar.progress(0.4)

        # === 병합 ===
        status.info("🔗 학생별 데이터 병합 중...")
        mf = load_mathflat_all(OUTPUT)
        pg = load_progress(progress_json)
        records = build_student_records(mf, pg)
        merged_json = OUTPUT / f"merged_{int(year)}_{int(month):02d}.json"
        with merged_json.open("w", encoding="utf-8") as f:
            json.dump({"year": int(year), "month": int(month), "students": records},
                      f, ensure_ascii=False, indent=2)
        progress_bar.progress(0.55)
        status.success(f"✅ 병합 완료 — 전체 {len(records)}명")

        # === 공지사항 저장 ===
        notice_path.write_text(notice_text, encoding="utf-8")

        # === AI 코멘트 생성 ===
        if not skip_ai:
            if not os.getenv("ANTHROPIC_API_KEY"):
                st.error("❌ ANTHROPIC_API_KEY가 없어 AI 코멘트를 생성할 수 없습니다. .env 파일을 확인하세요.")
            else:
                status.info("🤖 AI 코멘트 생성 중... (1~2분 소요)")
                from anthropic import Anthropic
                from generate_comments import generate_comment
                client = Anthropic()

                # 기존 코멘트 유지/재생성 옵션
                with merged_json.open(encoding="utf-8") as f:
                    data = json.load(f)

                generated, skipped = 0, 0
                n_total = len(data["students"])

                for i, student in enumerate(data["students"]):
                    if not (student.get("weekly_evals") or student.get("unit_evals") or student.get("progress")):
                        continue
                    if student.get("ai_comment") and not force_regen:
                        skipped += 1
                        continue
                    try:
                        comment, _ = generate_comment(client, student)
                        student["ai_comment"] = comment
                        generated += 1
                    except Exception as e:
                        st.warning(f"코멘트 실패: {student.get('name')} - {e}")
                    progress_bar.progress(0.55 + 0.35 * (i + 1) / n_total)
                    status.info(f"🤖 AI 코멘트 생성 중... [{i+1}/{n_total}] {student.get('name','')}")

                with merged_json.open("w", encoding="utf-8") as f:
                    json.dump(data, f, ensure_ascii=False, indent=2)
                status.success(f"✅ AI 코멘트 완료 — 생성 {generated}명 / 재사용 {skipped}명")
        progress_bar.progress(0.9)

        # === 가통문 렌더링 ===
        status.info("📄 가통문 xlsx 생성 중...")
        from render_gatongmun import main as render_main
        # render_main은 sys.argv를 읽으므로 조작
        sys.argv = ["render_gatongmun.py", str(int(year)), str(int(month))]
        render_main()
        progress_bar.progress(1.0)

        out_xlsx = OUTPUT / f"가통문_{int(year)}_{int(month):02d}.xlsx"
        st.success(f"🎉 완료! 총 {len(records)}명의 가통문이 생성됐습니다.")
        st.balloons()

        # 다운로드 버튼
        if out_xlsx.exists():
            st.warning(
                "⚠️ **꼭 인쇄는 '한 페이지에 모든 열 맞추기' 선택하기!!!**",
                icon="🖨️",
            )
            with out_xlsx.open("rb") as f:
                st.download_button(
                    "📥 가통문 xlsx 다운로드",
                    f.read(),
                    file_name=out_xlsx.name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                    use_container_width=True,
                )

        # 통계 표시
        n_a = sum(1 for r in records if r["has_evaluations"])
        n_b = len(records) - n_a
        n_warn = sum(1 for r in records if r.get("warnings"))

        st.subheader("📊 생성 요약")
        col_a, col_b, col_c = st.columns(3)
        col_a.metric("총 학생", f"{len(records)}명")
        col_b.metric("레이아웃 A (평가+진도)", f"{n_a}명")
        col_c.metric("레이아웃 B (진도만)", f"{n_b}명")

        if n_warn > 0:
            st.warning(f"⚠️ 경고 있는 학생 {n_warn}명 — 가통문에 경고 메시지 표시됨")
            with st.expander("경고 상세 보기"):
                for r in records:
                    if r.get("warnings"):
                        st.write(f"- **{r['code']} {r['name']}**: {' / '.join(r['warnings'])}")

    except Exception as e:
        status.error(f"❌ 오류 발생: {e}")
        import traceback
        with st.expander("상세 오류"):
            st.code(traceback.format_exc())


# ---------- PDF 일괄 변환 ----------
import platform
_is_windows = platform.system() == "Windows"

st.divider()
st.header("6️⃣ PDF 일괄 변환 (선택)")

if not _is_windows:
    st.info(
        "📌 PDF 변환은 Excel이 설치된 **Windows PC**에서만 가능합니다. "
        "xlsx 파일을 다운로드한 뒤 로컬에서 변환하세요."
    )
else:
    st.caption(
        "생성된 가통문 xlsx를 학생별 PDF로 일괄 변환합니다. "
        "Excel이 설치된 Windows에서만 동작하며, 40~50명 기준 1~2분 소요됩니다."
    )

current_xlsx = OUTPUT / f"가통문_{int(year)}_{int(month):02d}.xlsx"
if not current_xlsx.exists():
    st.info(f"먼저 가통문 xlsx를 생성하세요. (찾는 파일: `{current_xlsx.name}`)")
elif not _is_windows:
    st.success(f"대상 파일: `{current_xlsx.name}` (xlsx 다운로드 후 로컬에서 PDF 변환)")
elif _is_windows:
    st.success(f"대상 파일: `{current_xlsx.name}`")
    if st.button("📄 PDF로 변환 시작", use_container_width=True):
        pdf_status = st.empty()
        pdf_status.info("📄 PDF 변환 중... (1~2분 소요, Excel이 백그라운드에서 실행됩니다)")
        try:
            import subprocess
            result = subprocess.run(
                [sys.executable, str(SCRIPTS / "export_pdf.py"),
                 str(int(year)), str(int(month))],
                capture_output=True, text=True, timeout=600,
                cwd=str(BASE),
            )
            if result.returncode == 0:
                zip_path = OUTPUT / f"가통문_{int(year)}_{int(month):02d}.zip"
                # 생성된 PDF 수 세기
                pdf_dir = OUTPUT / "pdf"
                n_pdfs = len(list(pdf_dir.glob("*.pdf"))) if pdf_dir.exists() else 0
                pdf_status.success(f"✅ {n_pdfs}개 PDF 변환 완료")

                if zip_path.exists():
                    with zip_path.open("rb") as f:
                        st.download_button(
                            f"📦 PDF 전체 다운로드 ({zip_path.name})",
                            f.read(),
                            file_name=zip_path.name,
                            mime="application/zip",
                            type="primary",
                            use_container_width=True,
                        )
                with st.expander("변환 로그"):
                    st.code(result.stdout)
            else:
                pdf_status.error("❌ PDF 변환 실패")
                st.code(result.stderr or result.stdout)
        except subprocess.TimeoutExpired:
            pdf_status.error("❌ PDF 변환 시간 초과 (10분). Excel 프로세스 확인 필요.")
        except Exception as e:
            pdf_status.error(f"❌ PDF 변환 실패: {e}")
            import traceback
            with st.expander("상세 오류"):
                st.code(traceback.format_exc())


# --- 푸터 ---
st.divider()
st.caption("리바이브수학전문학원 | Made with Claude Sonnet 4.5 + Streamlit")
