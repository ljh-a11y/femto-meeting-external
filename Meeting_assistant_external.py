import streamlit as st
from streamlit_mic_recorder import mic_recorder
from streamlit_lottie import st_lottie
import datetime
import time
import requests
import warnings
import urllib.parse
import os
import openpyxl
from openpyxl.styles import Alignment
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from google.cloud import storage, speech, firestore
import google.auth
import vertexai
from vertexai.generative_models import GenerativeModel

# --- 시스템 설정 및 경고 제어 ---
warnings.filterwarnings("ignore")

PROJECT_ID = "femto-ai-assistant-497500"
LOCATION = "asia-northeast3"
BUCKET_NAME = "femto-meeting-luke"
TEMPLATE_FILE = "ai 회의록.xlsx"  # 엑셀 템플릿 파일명

my_options = {"quota_project_id": PROJECT_ID}

# ========================================================
# [구글 인증] 파이참 터미널 로그인 정보(gcloud)와 연동
# ========================================================
try:
    creds, _ = google.auth.default(quota_project_id=PROJECT_ID)
    vertexai.init(project=PROJECT_ID, location=LOCATION, credentials=creds)
except Exception as e:
    st.error(f" 구글 인증서 로드 실패: {e}")
    st.info(" 해결 방법: 파이참 하단 [Terminal] 탭에서 'gcloud auth application-default login' 명령어를 입력해 로그인을 완료해 주세요.")

# --- 프리미엄 디자인 ---
st.set_page_config(page_title="Femto AI Dash (External)", layout="wide", initial_sidebar_state="expanded")
st.markdown("""
    <style>
    .stApp { background-color: #F8FAFC; color: #0F172A; }
    .stat-card {
        background: #FFFFFF; padding: 24px; border-radius: 20px; border: 1px solid #E2E8F0;
        text-align: center; box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05);
    }
    .stat-val { color: #2563EB; font-size: 1.8rem; font-weight: 800; }
    .stat-lbl { color: #64748B; font-size: 0.9rem; font-weight: 600; }
    .live-monitor {
        background-color: #FFFFFF; color: #334155; padding: 25px; border-radius: 16px;
        font-family: 'Consolas', monospace; font-size: 1.1rem; line-height: 1.8;
        height: 300px; overflow-y: auto; border: 1px solid #CBD5E1;
    }
    </style>
""", unsafe_allow_html=True)


# --- 이메일 발송 함수 ---
def send_email_with_excel(to_email, file_path, subject, body_text):
    SMTP_SERVER = "smtp.naver.com"
    SMTP_PORT = 465
    SENDER_EMAIL = "내이메일@naver.com"
    SENDER_PASSWORD = "내보안앱비밀번호"

    try:
        msg = MIMEMultipart()
        msg['From'] = SENDER_EMAIL
        msg['To'] = to_email
        msg['Subject'] = subject
        msg.attach(MIMEText(body_text, 'plain'))

        with open(file_path, "rb") as attachment:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(attachment.read())
            encoders.encode_base64(part)
            part.add_header("Content-Disposition", f"attachment; filename={os.path.basename(file_path)}")
            msg.attach(part)

        server = smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT)
        server.login(SENDER_EMAIL, SENDER_PASSWORD)
        server.sendmail(SENDER_EMAIL, to_email, msg.as_string())
        server.quit()
        return True
    except Exception as e:
        st.error(f"이메일 발송 실패: {e}")
        return False


# --- GCS 파일 가져오기 ---
def list_gcs_files():
    try:
        client_s = storage.Client(project=PROJECT_ID, client_options=my_options)
        blobs = list(client_s.list_blobs(BUCKET_NAME))
        return sorted([blob.name for blob in blobs if blob.name.endswith('.wav')], reverse=True)
    except Exception as e:
        st.error(f" 구글 스토리지 연결 실패: {e}")
        return []


# --- 1단계: 음성 인식 엔진 ---
def speech_to_text(audio_data, lang_code, gcs_uri=None, selected_filename=None):
    try:
        client_s = storage.Client(project=PROJECT_ID, client_options=my_options)
        bucket = client_s.bucket(BUCKET_NAME)

        if selected_filename:
            txt_filename = selected_filename.replace('.wav', '.txt')
            txt_blob = bucket.blob(txt_filename)
            if txt_blob.exists():
                st.toast(" 저장된 자막 파일(.txt)을 초고속으로 로드했습니다!")
                return txt_blob.download_as_text(encoding="utf-8")

        if not gcs_uri:
            fname = f"meeting_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.wav"
            blob = bucket.blob(fname)
            blob.upload_from_string(audio_data, content_type="audio/wav")
            gcs_uri = f"gs://{BUCKET_NAME}/{fname}"
            selected_filename = fname

        client_speech = speech.SpeechClient(client_options=my_options)
        audio = speech.RecognitionAudio(uri=gcs_uri)

        femto_phrases = [
            "펨토사이언스", "개발혁신팀", "통합사업팀", "제너레이터", "PECVD", "ICP",
            "김무환", "김무환 대표", "김대현", "김대현 고문", "류재홍", "류재홍 박사",
            "김준민", "김준민 팀장", "조인오", "조인오 프로", "윤진성", "윤진성 프로",
            "김다다", "김다다 팀장", "조이정", "조이정 대리", "박홍근", "박홍근 프로", "이주혁", "이주혁 프로"
        ]

        femto_context = speech.SpeechContext(phrases=femto_phrases)

        config = speech.RecognitionConfig(
            language_code=lang_code,
            enable_automatic_punctuation=True,
            use_enhanced=True,
            model="latest_long",
            speech_contexts=[femto_context]
        )

        operation = client_speech.long_running_recognize(config=config, audio=audio)

        p_bar = st.progress(0)
        p_status = st.empty()

        while not operation.done():
            metadata = operation.metadata
            if metadata:
                percent = getattr(metadata, 'progress_percent', 0)
                p_bar.progress(min(int(percent), 100))
                p_status.info(f"⏳ 구글 STT 인공지능 음성 분석 중... ({percent}%)")
            time.sleep(2)

        p_bar.empty()
        p_status.empty()

        response = operation.result(timeout=3600)
        full_text = " ".join([r.alternatives[0].transcript for r in response.results])
        final_result = full_text if full_text.strip() else "(인식된 내용 없음)"

        if final_result and final_result != "(인식된 내용 없음)":
            save_txt_name = selected_filename.replace('.wav', '.txt')
            new_txt_blob = bucket.blob(save_txt_name)
            new_txt_blob.upload_from_string(final_result, content_type="text/plain")

        return final_result
    except Exception as e:
        st.error(f"STT 오류: {e}")
        return ""


# --- 2단계: Gemini 2.5 압축 및 과거 이력 지속 학습 엔진 ---
def generate_report_text_only(text_content):
    try:
        model = GenerativeModel("gemini-2.5-flash")

        history_context = ""
        try:
            # 📌 [외부용 변경] 과거 이력 참고 시 외부 전용 DB(external-meetings)를 바라봅니다.
            db_fs = firestore.Client(project=PROJECT_ID, database="external-meetings", client_options=my_options)
            past_docs = list(
                db_fs.collection("meetings").order_by("date", direction=firestore.Query.DESCENDING).limit(3).stream())
            if past_docs:
                history_context = "\n[사전 학습 데이터 - 과거 회의 기록의 요약 서식 및 단어 패턴 참고]\n"
                for doc in past_docs:
                    d = doc.to_dict()
                    history_context += f"- 과거 안건: {d.get('title', '')} / 요약 스타일: {d.get('ai_summary', '')[:100]}...\n"
        except:
            pass

        prompt = f"""당신은 '펨토사이언스'의 전문 비서 AI입니다.
        제공된 회의 스크립트를 바탕으로 가시성이 뛰어난 고품질 핵심 회의록 리포트를 작성하세요.

        [ ⚠️ 핵심 제한 조건 - 절대 준수]
        1. 세부 사항을 일일이 나열하지 말고, 자잘한 지시사항은 과감히 통합하거나 생략하여 전체 분량을 콤팩트하게 압축하세요.
        2. 엑셀 출력 시 빈 칸을 포함하여 총 70줄 이내로 끊길 수 있도록 핵심 위주로 요약해야 4페이지를 넘지 않습니다.
        3. 다른 말 없이 무조건 본론 대괄호 헤더부터 즉시 시작하세요.

        [작성 형식]
        - 내용을 더 넓은 범위로 크게 묶어서 최대 4~5개의 [대괄호 주제]로만 분류하세요. (예: [영업 마케팅 및 전략 강화를 위한 지시사항], [핵심 개발 과제 및 표준화 추진 현황] 등)
        - 정중하고 깔끔한 비즈니스 개조식 어투(~할 것, ~ 요망, ~ 완료 바람)를 사용하세요.
        - 모든 세부 항목은 개별 줄마다 '• ' 기호로 시작하세요.

        {history_context}

        내용: {text_content}"""

        ai_res = model.generate_content(prompt).text
        return ai_res
    except Exception as e:
        st.error(f"리포트 생성 오류: {e}")
        return ""


# --- 3단계: 엑셀 정밀 틀 고정 데이터 기입 엔진 ---
def save_final_summary_to_excel(final_summary, meta_info):
    try:
        if os.path.exists(TEMPLATE_FILE):
            wb = openpyxl.load_workbook(TEMPLATE_FILE, data_only=False)
            ws = wb.active

            def safe_write(coord, value, alignment=None):
                target_cell = ws[coord]
                if type(target_cell).__name__ == 'MergedCell':
                    for r_range in ws.merged_cells.ranges:
                        if coord in r_range:
                            target_cell = ws.cell(row=r_range.min_row, column=r_range.min_col)
                            break
                target_cell.value = value
                if alignment:
                    target_cell.alignment = alignment
                return target_cell

            # 메타 정보 입력 구역 연동 (6행~9행)
            safe_write('B6', meta_info['일시'])
            safe_write('E6', meta_info['부서'])
            safe_write('B7', meta_info['주관'])
            safe_write('E7', meta_info['작성자'])
            safe_write('B8', meta_info['참석자'])
            safe_write('B9', meta_info['안건'])

            # 각 페이지별 실제 내용 입력 가능 행 범위 지정 (29, 59, 89, 119행은 빈 행으로 둡니다)
            p1_rows = list(range(10, 29))  # 1페이지 본문 (10~28행)
            p2_rows = list(range(32, 59))  # 2페이지 본문 (32~58행)
            p3_rows = list(range(62, 89))  # 3페이지 본문 (62~88행)
            p4_rows = list(range(92, 119)) # 4페이지 본문 (92~118행)

            paragraphs = []
            current_block = []
            category_counter = 1

            for line in final_summary.split('\n'):
                clean_line = line.replace('**', '').replace('###', '').replace('##', '').replace('* ', '• ').strip()
                if not clean_line:
                    continue

                if clean_line.startswith('[') and clean_line.endswith(']'):
                    if current_block:
                        paragraphs.append(current_block)
                    clean_line = f"{category_counter}. {clean_line}"
                    category_counter += 1
                    current_block = [clean_line]
                else:
                    MAX_CHAR_PER_ROW = 60
                    is_bullet = clean_line.startswith('•')
                    indent = "    " if is_bullet else ""

                    temp_text = clean_line
                    is_first_chunk = True

                    while len(temp_text) > MAX_CHAR_PER_ROW:
                        split_idx = temp_text.rfind(' ', 0, MAX_CHAR_PER_ROW)
                        if split_idx == -1 or split_idx == 0:
                            split_idx = MAX_CHAR_PER_ROW

                        chunk = temp_text[:split_idx].strip()
                        if not is_first_chunk:
                            chunk = indent + chunk
                        current_block.append(chunk)
                        temp_text = temp_text[split_idx:].strip()
                        is_first_chunk = False

                    if temp_text:
                        if not is_first_chunk:
                            temp_text = indent + temp_text
                        current_block.append(temp_text)

            if current_block:
                paragraphs.append(current_block)

            # 데이터 맵핑 및 단락간 1행 건너뛰기
            final_mapped_lines = []
            p1_idx = 0
            p2_idx = 0
            p3_idx = 0
            p4_idx = 0
            current_page = 1

            for idx_p, block in enumerate(paragraphs):
                block_len = len(block)
                actual_need_len = block_len + (1 if idx_p > 0 else 0)

                if current_page == 1:
                    if p1_idx + actual_need_len <= len(p1_rows):
                        if idx_p > 0:  # 항목 간 1행 비우기
                            final_mapped_lines.append((p1_rows[p1_idx], ""))
                            p1_idx += 1
                        for b_line in block:
                            final_mapped_lines.append((p1_rows[p1_idx], b_line))
                            p1_idx += 1
                    else:
                        current_page = 2

                if current_page == 2:
                    is_page_start = (p2_idx == 0)
                    need_len_p2 = block_len + (0 if is_page_start else 1)
                    if p2_idx + need_len_p2 <= len(p2_rows):
                        if not is_page_start:
                            final_mapped_lines.append((p2_rows[p2_idx], ""))
                            p2_idx += 1
                        for b_line in block:
                            final_mapped_lines.append((p2_rows[p2_idx], b_line))
                            p2_idx += 1
                    else:
                        current_page = 3

                if current_page == 3:
                    is_page_start = (p3_idx == 0)
                    need_len_p3 = block_len + (0 if is_page_start else 1)
                    if p3_idx + need_len_p3 <= len(p3_rows):
                        if not is_page_start:
                            final_mapped_lines.append((p3_rows[p3_idx], ""))
                            p3_idx += 1
                        for b_line in block:
                            final_mapped_lines.append((p3_rows[p3_idx], b_line))
                            p3_idx += 1
                    else:
                        current_page = 4

                if current_page == 4:
                    is_page_start = (p4_idx == 0)
                    need_len_p4 = block_len + (0 if is_page_start else 1)
                    if p4_idx + need_len_p4 <= len(p4_rows):
                        if not is_page_start:
                            final_mapped_lines.append((p4_rows[p4_idx], ""))
                            p4_idx += 1
                        for b_line in block:
                            final_mapped_lines.append((p4_rows[p4_idx], b_line))
                            p4_idx += 1
                    else:
                        if not is_page_start and p4_idx < len(p4_rows):
                            final_mapped_lines.append((p4_rows[p4_idx], ""))
                            p4_idx += 1
                        for b_line in block:
                            if p4_idx < len(p4_rows):
                                final_mapped_lines.append((p4_rows[p4_idx], b_line))
                                p4_idx += 1

            # 최종 문장 끝 바로 밑에 "-- 이 상 --" 표식 배치
            if current_page == 1:
                if p1_idx < len(p1_rows):
                    final_mapped_lines.append((p1_rows[p1_idx], "-- 이 상 --"))
                else:
                    final_mapped_lines.append((p2_rows[0], "-- 이 상 --"))
                    current_page = 2
            elif current_page == 2:
                if p2_idx < len(p2_rows):
                    final_mapped_lines.append((p2_rows[p2_idx], "-- 이 상 --"))
                else:
                    final_mapped_lines.append((p3_rows[0], "-- 이 상 --"))
                    current_page = 3
            elif current_page == 3:
                if p3_idx < len(p3_rows):
                    final_mapped_lines.append((p3_rows[p3_idx], "-- 이 상 --"))
                else:
                    final_mapped_lines.append((p4_rows[0], "-- 이 상 --"))
                    current_page = 4
            elif current_page == 4:
                if p4_idx < len(p4_rows):
                    final_mapped_lines.append((p4_rows[p4_idx], "-- 이 상 --"))
                else:
                    final_mapped_lines.append((p4_rows[-1], "-- 이 상 --"))

            # 실제 채워진 총 라인 구역 연산
            used_rows = [item[0] for item in final_mapped_lines]
            max_used_row = max(used_rows) if used_rows else 10

            if max_used_row <= 28:
                total_pages = 1
            elif max_used_row <= 58:
                total_pages = 2
            elif max_used_row <= 88:
                total_pages = 3
            else:
                total_pages = 4

            # 타이틀 서식 연동 구역 매핑 (30~31, 60~61, 90~91행이 타이틀 칸)
            safe_write('A4', f"회  의  록 ( 1 / {total_pages} )")
            if total_pages >= 2:
                safe_write('A30', f"회  의  록 ( 2 / {total_pages} )")
            if total_pages >= 3:
                safe_write('A60', f"회  의  록 ( 3 / {total_pages} )")
            if total_pages >= 4:
                safe_write('A90', f"회  의  록 ( 4 / {total_pages} )")

            # 사용하지 않는 뒷페이지 영역 자동 숨김 처리
            if total_pages == 1:
                for r_clear in range(29, 130):
                    ws.row_dimensions[r_clear].hidden = True
            elif total_pages == 2:
                for r_clear in range(59, 130):
                    ws.row_dimensions[r_clear].hidden = True
            elif total_pages == 3:
                for r_clear in range(89, 130):
                    ws.row_dimensions[r_clear].hidden = True
            elif total_pages == 4:
                for r_clear in range(119, 130):
                    ws.row_dimensions[r_clear].hidden = True

            # 최종 매핑 데이터 기입
            for r, line in final_mapped_lines:
                ws.row_dimensions[r].height = 22
                coord = f'B{r}'
                if line == "-- 이 상 --":
                    written_cell = safe_write(coord, line)
                    written_cell.alignment = Alignment(vertical='center', horizontal='center')
                else:
                    written_cell = safe_write(coord, line)
                    written_cell.alignment = Alignment(vertical='center', horizontal='left')

            output_filename = f"회의록_{datetime.date.today()}_{meta_info['주관']}.xlsx"
            wb.save(output_filename)
            return output_filename
        else:
            st.error("폴더 내에 'ai 회의록.xlsx' 원본 템플릿 파일이 없습니다.")
            return None
    except Exception as e:
        st.error(f"엑셀 빌드 오류: {e}")
        return None


# --- 메인 대시보드 UI ---
st.title("🏢 FEMTO SCIENCE AI COMMAND CENTER (EXTERNAL)")

tab1, tab2 = st.tabs(["🚀 LIVE CONSOLE", "📁 MEETING ARCHIVE"])

with tab1:
    col_meta, col_action = st.columns([1, 1.5], gap="large")

    with col_meta:
        st.subheader("📋 회의 기본 정보 입력")
        m_dept = st.text_input("부서", "통합사업팀 (IBD Team)")
        m_writer = st.text_input("작성자", "이주혁 사원")
        m_host = st.text_input("주관 (발표자/회의 리더)", "이주혁")
        m_attendees = st.text_input("참석자", "전직원 (대표님 포함)")
        m_title = st.text_input("안건 (회의 제목)", f"정기 업무 보고_{datetime.date.today()}")
        m_date = st.text_input("일시", datetime.datetime.now().strftime("%Y-%m-%d %H:%M"))

        meta_data = {"부서": m_dept, "작성자": m_writer, "주관": m_host, "참석자": m_attendees, "안건": m_title, "일시": m_date}

    with col_action:
        st.subheader("🎙️ 회의 오디오 입력 / 양식 테스트 선택")

        input_mode = st.radio("방식 선택", ["실시간 녹음", "기존 파일 선택", "📝 회의록 텍스트 직접 입력"])

        if input_mode == "실시간 녹음":
            audio_record = mic_recorder(start_prompt="▶️ 녹음 시작", stop_prompt="⏹️ 녹음 종료 및 자막 변환", key='rec')
            if audio_record:
                st.session_state.transcript = speech_to_text(audio_record['bytes'], "ko-KR")

        elif input_mode == "기존 파일 선택":
            gcs_files = list_gcs_files()
            if gcs_files:
                selected_file = st.selectbox("파일 선택", gcs_files)
                if st.button("🚀 선택 파일 자막 변환 시작", type="primary"):
                    gcs_uri = f"gs://{BUCKET_NAME}/{selected_file}"
                    st.session_state.transcript = speech_to_text(None, "ko-KR", gcs_uri=gcs_uri,
                                                                 selected_filename=selected_file)
            else:
                st.warning(" 구글 스토리지에서 오디오(.wav) 목록을 불러오지 못했거나 파일이 없습니다.")

        elif input_mode == "📝 회의록 텍스트 직접 입력":
            st.info("💡 녹음이나 AI 요약을 패스하고, 아래 텍스트 창에 입력한 대로 엑셀 양식(글자 쪼개기, 페이지 분할)이 어떻게 생성되는지 즉시 테스트합니다.")

            sample_init_value = (
                "[영업 및 마케팅 전략 강화]\n"
                "신규 장착 기능 및 개선 사항을 강조하되, 가격은 동결하여 경쟁력 확보 요망.\n"
                "제품의 강점과 가치를 명확히 전달하여 고객 구매를 유도하는 적극적인 영업 방식 필요.\n"
                "제품 사양서 및 견적서를 상세하고 체계적으로 작성하여 고객 문의에 즉각 대응할 준비 완료 요망.\n"
                "경쟁사 제품(예: 펌프) 사양 및 기능을 숙지하여 시장 대응 능력 강화.\n"
                "[3대 과제 표준화 추진 현황]\n"
                "ICP 표준화: 6월 말까지 완료 목표.\n"
                "PCVD 표준화: 일본 대상 발표자 지정 및 발표 자료 준비 철저.\n"
                "HLS 및 MRPG: 6월 첫째 주에 주역, 경력 담당으로 일정 확정 및 추진.\n"
                "표준화된 절차 및 내용을 사내 전체에 빠르게 적용할 것.\n"
                "[영업 발표 및 대외 커뮤니케이션 준비]\n"
                "발표 자료: PPT는 금일까지 최종 완료하고, 발표 시간 준수 연습 필수.\n"
                "온라인 회의 플랫폼: Zoom, MS Teams 등 통일된 플랫폼을 선정하고, 필요시 IBD팀에서 라이선스 구매 후 공지 요망.\n"
                "상세 사양서: 제품에 대한 깊이 있는 이해를 바탕으로 상세하게 작성하여 내부 지식 공유 및 활용 극대화.\n"
                "[AI 도입 및 활용 방안]\n"
                "AI 적용 분야:\n"
                "경비 정산: 현재 적용 방식 유지.\n"
                "도면 및 BOM 추출: PCVD 표준화 데이터를 활용하여 도면 및 BOM 추출 자동화 준비 (경민 담당).\n"
                "회의록 요약 및 브리핑: AI 활용하여 지난 회의 내용 요약 및 브리핑 기능 도입, 음성 인식 정확도 개선 지속 추진.\n"
                "레퍼런스 관리 (홍과장): AI를 활용한 홈페이지 레퍼런스 검색/추천 기능 구현 방안 구체화 및 AI의 역할 명확히 설명 요망.\n"
                "AI 도입을 지연 없이 적극적으로 추진하고 활용 범위를 확대할 것.\n"
                "[내부 인프라 및 자원 관리]\n"
                "PC 구매: AI 및 사내 활용 PC는 조립 PC로 진행, 인원 담당자가 최종 사양 검토 후 구매 준비. 기존 PC는 하드디스크 분리 후 폐기 등 정리 요망.\n"
                "NAS 접근: 김 팀장 및 조 대리는 6월 초까지 NAS 접근 권한 설정 및 활용 준비 완료.\n"
                "한수 활용 극대화: 조립 및 서비스 업무(특히 Class B 장비)를 한수로 과감히 이관하고, 이에 대한 구체적인 계획 및 일정 수립 요망. 핵심 인력은 고부가가치 업무에 집중.\n"
                "[개발 및 프로젝트 관리]\n"
                "개발 일정 준수: 성균관대 ICP 장비 등 모든 개발 프로젝트의 납기일을 철저히 준수할 것.\n"
                "PROEJCT 예산 및 BOM: 김 팀장은 각 프로젝트의 전체 예산 및 약식 BOM을 즉시 제출하여 집행 가능하도록 조치.\n"
                "툴로나 프로젝트: 기술 협의 후 상대방과의 상업적 조건(일정, 금액, 수수료) 협상 가속화. 가격 인하 및 기술 획득을 목표로 협상 전략 수립.\n"
                "청정실 및 연구소 관리:\n"
                "연구소 자산 실사 및 정리 정돈 철저: 종이 박스 사용 금지, 플라스틱 박스 활용, 재고 관리 시스템 구축.\n"
                "연구 공간을 명확히 구분하고 정리하여 연구 활동의 효율성 증대.\n"
                "[주요 프로젝트 현황 공유]\n"
                "AI 진단 장비 (유 박사):\n"
                "OS 장비에 AI 진단 기능(가스, 압력 점검) 탑재 완료, 경북대에서 시뮬레이션 테스트 진행 중.\n"
                "실제 가스 테스트는 청정실 설치 후 진행 예정이며, 이후 전 박사에게 이관하여 모델 비교 진행.\n"
                "유 박사는 현재 진행 중인 업무 내용을 팀장들에게 간략히 공유하여 협업 강화 요망.\n"
                "정부 과제 (고문님):\n"
                "총 5개 과제 진행 중이며, 진행 상황은 정기적으로 공유 예정 (1억 원 과제 협약 평가 완료 등).\n"
                "연구 노트 작성의 중요성을 강조하고, 9월 종료 과제는 마무리 준비에 만전.\n"
                "[경영 전략 및 방향]\n"
                "회사 경쟁력 강화를 위해 레벨 1에서 레벨 2로의 전환을 신속히 추진.\n"
                "상반기까지는 예산 투자를 유지하나, 하반기부터는 P.A.L.D. 등 전략적 신규 개발을 제외하고는 엄격한 예산 집행 원칙 적용.\n"
                "제품에 대한 정당한 가치를 인정받아 제값을 받고 판매하는 원칙을 준수, 불필요한 가격 인하 요구에 응하지 않을 것."
            )

            sandbox_text = st.text_area("엑셀 반영 테스트용 내용 입력 (마음대로 편집해 보세요)", value=sample_init_value, height=250)

            if st.button("📊 테스트 엑셀 문서 즉시 빌드", type="primary"):
                st.session_state.ai_summary = sandbox_text
                st.session_state.transcript = "(엑셀 양식 단독 레이아웃 테스트 모드로 생성됨)"
                with st.spinner("템플릿 양식에 맞춰 엑셀 빌드 중..."):
                    excel_path = save_final_summary_to_excel(sandbox_text, meta_data)
                    if excel_path:
                        st.session_state.excel_path = excel_path

    if input_mode != "📝 회의록 텍스트 직접 입력" and 'transcript' in st.session_state:
        st.divider()
        st.subheader("📝 1차 음성 인식 결과 (필요시 오타를 수정하세요)")
        edited_text = st.text_area("인식된 대화 내용 (STT)", value=st.session_state.transcript, height=220)

        if st.button("✨ 1차 AI 요약 리포트 초안 생성", type="secondary"):
            st.session_state.transcript = edited_text
            with st.spinner("Gemini가 회의 내용을 요약하고 있습니다..."):
                ai_summary = generate_report_text_only(edited_text)
                if ai_summary:
                    st.session_state.ai_summary = ai_summary
                    if 'excel_path' in st.session_state:
                        del st.session_state['excel_path']

    if input_mode != "📝 회의록 텍스트 직접 입력" and 'ai_summary' in st.session_state:
        st.divider()
        st.subheader("🤖 AI 최종 요약 리포트 검토 (엑셀 삽입 전 편집 가능)")
        st.caption(" AI가 요약한 내용 중 지우거나 수정할 문구가 있다면 아래 창에서 직접 수정한 후 최종 문서를 만드세요.")

        edited_summary = st.text_area("엑셀에 반영될 요약 내용 편집", value=st.session_state.ai_summary, height=300)

        if st.button("📊 이 내용으로 최종 엑셀 문서 빌드", type="primary"):
            st.session_state.ai_summary = edited_summary
            with st.spinner("편집된 내용을 토대로 'ai 회의록.xlsx' 템플릿에 쓰는 중..."):
                excel_path = save_final_summary_to_excel(edited_summary, meta_data)
                if excel_path:
                    st.session_state.excel_path = excel_path

                    try:
                        # 📌 [외부용 변경] 저장 시 외부 전용 DB(external-meetings) 창고로 들어가도록 지정
                        db = firestore.Client(project=PROJECT_ID, database="external-meetings", client_options=my_options)
                        db.collection("meetings").add({
                            "date": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                            "title": meta_data['안건'],
                            "host": meta_data['주관'],
                            "dept": meta_data['부서'],
                            "writer": meta_data['작성자'],
                            "attendees": meta_data['참석자'],
                            "transcript": st.session_state.transcript,
                            "ai_summary": st.session_state.ai_summary
                        })
                        st.toast("🔥 외부 전용 클라우드 DB에 회의록 기록이 누적 백업되었습니다!")
                    except Exception as db_err:
                        st.warning(f"DB 백업 중 권한 오류 발생 (로컬 엑셀 파일은 정상 빌드됨): {db_err}")

    if 'excel_path' in st.session_state:
        st.divider()
        st.success("🎉 최종 엑셀 회의록 문서 생성이 완료되었습니다!")

        with open(st.session_state.excel_path, "rb") as f:
            st.download_button(
                label="📥 완성된 엑셀 회의록 다운로드",
                data=f,
                file_name=st.session_state.excel_path,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        st.write("")
        st.markdown("### 📧 부서원에게 회의록 즉시 이메일 발송")
        target_email = st.text_input("수신자 이메일 주소 입력", placeholder="example@femtoscience.co.kr")

        if st.button("✉️ 회의록 이메일 전송"):
            if target_email:
                with st.spinner(" 메일에 엑셀 파일을 첨부하여 발송 중입니다..."):
                    mail_subject = f"[회의록 공유] {meta_data['안건']}"
                    mail_body = f"안녕하세요.\n\n{meta_data['일시']}에 진행된 회의의 AI 회의록을 공유해 드립니다.\n자세한 요약본 및 원본 내용은 첨부된 엑셀 파일을 확인해 주세요.\n\n감사합니다."

                    success = send_email_with_excel(target_email, st.session_state.excel_path, mail_subject, mail_body)
                    if success:
                        st.success(f" {target_email} 로 메일이 성공적으로 전송되었습니다!")
            else:
                st.warning("이메일 주소를 입력해 주세요.")

with tab2:
    st.subheader("📂 Meeting Archive (외부 전용 과거 이력 보관소)")
    st.caption("외부 전용 파이어스토어 DB에 백업된 기록 목록만 최신순으로 조회합니다.")

    try:
        # 📌 [외부용 변경] 과거 기록 조회 시 외부 전용 DB(external-meetings) 창고를 읽어옵니다.
        db = firestore.Client(project=PROJECT_ID, database="external-meetings", client_options=my_options)
        docs = list(db.collection("meetings").order_by("date", direction=firestore.Query.DESCENDING).stream())

        if not docs:
            st.info("아직 외부 데이터베이스에 누적된 회의록 기록이 없습니다. 본문 빌드를 완료해 보세요.")

        for doc in docs:
            data = doc.to_dict()
            doc_id = doc.id

            with st.expander(
                    f"📅 {data.get('date', '날짜 정보 없음')} | 🏢 {data.get('dept', '부서 미지정')} | 📋 {data.get('title', '제목 없음')}"):
                st.markdown(
                    f"**🗣️ 회의 주관:** {data.get('host', '-')}   |   **✍️ 작성자:** {data.get('writer', '-')}   |   **👥 참석자:** {data.get('attendees', '-')}")
                st.markdown("---")

                st.markdown("### 🤖 AI 최종 요약 리포트 본문")
                st.info(data.get('ai_summary', '(요약 내용 없음)'))

                c1, c2 = st.columns([1, 4])
                with c1:
                    if st.button(f"🗑️ 기록 삭제", key=f"del_{doc_id}", help="클라우드 DB에서 이 회의록을 영구히 지웁니다."):
                        db.collection("meetings").document(doc_id).delete()
                        st.toast("선택하신 회의록 이력이 정상적으로 삭제되었습니다.")
                        time.sleep(0.8)
                        st.rerun()
                with c2:
                    with st.expander("🔍 회의 원본 스크립트(STT) 전문 보기"):
                        st.write(data.get('transcript', '(인식된 원본 텍스트가 없습니다.)'))

    except Exception as e:
        st.write(f"데이터베이스 조회 권한을 확인 중이거나 대기 상태입니다. ({e})")