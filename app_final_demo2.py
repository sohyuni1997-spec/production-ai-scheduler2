import streamlit as st
import pandas as pd
from supabase import create_client, Client
import requests
from datetime import datetime, timedelta
import re
import os
import logging
from io import BytesIO
import json

# 로깅 설정
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('scheduler_app.log'),
        logging.StreamHandler()
    ]
)

# ==================== 설정 및 초기화 ====================

# 환경변수에서 API 키 로드 (없으면 기본값 - 개발용)
SUPABASE_URL = os.getenv("SUPABASE_URL", "https://suaajrdahixouinbfcfo.supabase.co")
SUPABASE_KEY = os.getenv("SUPABASE_KEY", "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InN1YWFqcmRhaGl4b3VpbmJmY2ZvIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NjYzMTk4NzAsImV4cCI6MjA4MTg5NTg3MH0.Ic4izQY-ihIw75jKh9iJicZvuZ4gCRs4OH3rCGyo0Zk")
POTENS_API_KEY = os.getenv("POTENS_API_KEY", "qD2gfuVAkMJexDAcFb5GnEb1SZksTs7o")

try:
    supabase: Client = create_client(SUPABASE_URL, SUPABASE_KEY)
    logging.info("Supabase 연결 성공")
except Exception as e:
    logging.error(f"Supabase 연결 실패: {e}")
    st.error("⚠️ 데이터베이스 연결에 실패했습니다.")

TABLE_NAME = "pattern_learning2"

# CAPA 정보
CAPA_INFO = {
    "조립1": 3000,
    "조립2": 2500,
    "조립3": 2000
}

# ==================== 유틸리티 함수 ====================

def parse_date(date_str):
    """날짜 문자열을 YYYY-MM-DD 형식으로 변환"""
    try:
        if '/' in date_str:
            parts = date_str.split('/')
            month = parts[0].zfill(2)
            day = parts[1].zfill(2)
            current_year = datetime.now().year
            
            # 입력된 날짜가 과거라면 다음 해로 추정
            test_date = datetime.strptime(f"{current_year}-{month}-{day}", '%Y-%m-%d')
            if test_date < datetime.now() - timedelta(days=180):
                current_year += 1
            
            return f"{current_year}-{month}-{day}"
        return date_str
    except Exception as e:
        logging.error(f"날짜 파싱 오류: {date_str} - {e}")
        return None

def get_date_range(target_date, days_before=14, days_after=7):
    """목표 날짜 기준으로 조회 범위 계산"""
    try:
        dt = datetime.strptime(target_date, '%Y-%m-%d')
        start = (dt - timedelta(days=days_before)).strftime('%Y-%m-%d')
        end = (dt + timedelta(days=days_after)).strftime('%Y-%m-%d')
        return start, end
    except Exception as e:
        logging.error(f"날짜 범위 계산 오류: {e}")
        return None, None

def convert_df_to_excel(df):
    """DataFrame을 Excel 파일로 변환"""
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Schedule')
        
        # 워크북 및 워크시트 가져오기
        workbook = writer.book
        worksheet = writer.sheets['Schedule']
        
        # 헤더 포맷
        header_format = workbook.add_format({
            'bold': True,
            'bg_color': '#4472C4',
            'font_color': 'white',
            'border': 1
        })
        
        # 헤더 적용
        for col_num, value in enumerate(df.columns.values):
            worksheet.write(0, col_num, value, header_format)
            worksheet.set_column(col_num, col_num, 15)
    
    return output.getvalue()

# ==================== 데이터 조회 함수 ====================

@st.cache_data(ttl=600, show_spinner=False)  # 10분 캐시
def fetch_production_data(target_date, version='2차'):
    """DB에서 생산 데이터 조회 (캐싱 적용)"""
    try:
        start_date, end_date = get_date_range(target_date)
        if not start_date:
            logging.warning(f"날짜 범위 계산 실패: {target_date}")
            return None
        
        logging.info(f"DB 조회 시작: {version}, {start_date} ~ {end_date}")
        
        response = supabase.table(TABLE_NAME)\
            .select("*")\
            .eq("version", version)\
            .gte("due_date", start_date)\
            .lte("due_date", end_date)\
            .order("due_date", desc=False)\
            .execute()
        
        if response.data:
            df = pd.DataFrame(response.data)
            logging.info(f"데이터 조회 성공: {len(df)}건")
            return df
        else:
            logging.warning(f"조회 결과 없음: {version}, {target_date}")
            return None
            
    except Exception as e:
        logging.error(f"DB 조회 오류: {e}")
        return None

def compare_versions(df_0, df_2):
    """0차와 2차 버전 비교"""
    if df_0 is None or df_2 is None:
        return None
    
    try:
        merged = pd.merge(
            df_0[['due_date', 'line', 'product_name', 'product_type', 'quantity']],
            df_2[['due_date', 'line', 'product_name', 'product_type', 'quantity', 'status', 'remark', 'worker_memo']],
            on=['due_date', 'line', 'product_name', 'product_type'],
            how='outer',
            suffixes=('_0차', '_2차')
        )
        
        merged['quantity_0차'] = merged['quantity_0차'].fillna(0)
        merged['quantity_2차'] = merged['quantity_2차'].fillna(0)
        merged['qty_diff'] = merged['quantity_2차'] - merged['quantity_0차']
        merged['changed'] = merged['qty_diff'] != 0
        
        logging.info(f"버전 비교 완료: 총 {len(merged)}건, 변경 {merged['changed'].sum()}건")
        return merged
        
    except Exception as e:
        logging.error(f"버전 비교 오류: {e}")
        return None

# ==================== 데이터 분석 함수 ====================

def analyze_data(df, version='2차'):
    """생산 데이터 종합 분석"""
    try:
        analysis = {'version': version}
        
        # 날짜 처리
        df['plan_date_dt'] = pd.to_datetime(df['due_date'])
        df['weekday_kr'] = df['plan_date_dt'].dt.strftime('%A').map({
            'Monday': '월', 'Tuesday': '화', 'Wednesday': '수',
            'Thursday': '목', 'Friday': '금', 'Saturday': '토', 'Sunday': '일'
        })
        
        qty_col = 'quantity'
        
        # 라인별 CAPA 분석
        for line in ["조립1", "조립2", "조립3"]:
            line_data = df[df['line'] == line]
            daily_sum = line_data.groupby('due_date')[qty_col].sum()
            max_capa = CAPA_INFO[line]
            target_capa = max_capa * 0.9
            
            analysis[line] = {
                'max_capa': max_capa,
                'target_90': int(target_capa),
                'daily_production': daily_sum.to_dict(),
                'over_capacity_days': daily_sum[daily_sum > target_capa].to_dict(),
                'avg_utilization': (daily_sum.mean() / max_capa * 100) if len(daily_sum) > 0 else 0
            }
        
        # BERGSTROM 분석
        bergstrom_data = df[df['product_name'].str.contains('BERGSTROM', case=False, na=False)]
        bergstrom_days = bergstrom_data.groupby('due_date')[qty_col].sum().to_dict()
        
        # 조립2 요일 규칙 위반 체크
        line2_data = df[df['line'] == '조립2'].copy()
        
        # FAN: 월/수/금 생산
        fan_data = line2_data[line2_data['product_type'] == 'FAN']
        fan_wrong = fan_data[~fan_data['weekday_kr'].isin(['월', '수', '금'])]
        
        # FLANGE: 화/목 생산
        flange_data = line2_data[line2_data['product_type'] == 'FLANGE']
        flange_wrong = flange_data[~flange_data['weekday_kr'].isin(['화', '목'])]
        
        # 조립2 일일 제품 종류 수
        line2_daily_products = line2_data.groupby('due_date')['product_name'].nunique()
        
        analysis['bergstrom_days'] = bergstrom_days
        analysis['fan_violations'] = fan_wrong[['due_date', 'product_name', qty_col, 'weekday_kr', 'remark']].to_dict('records')
        analysis['flange_violations'] = flange_wrong[['due_date', 'product_name', qty_col, 'weekday_kr', 'remark']].to_dict('records')
        analysis['line2_over_5products'] = line2_daily_products[line2_daily_products > 5].to_dict()
        
        # 상태별 집계
        status_summary = df['status'].value_counts().to_dict()
        analysis['status_summary'] = status_summary
        
        logging.info(f"데이터 분석 완료: {version}")
        return analysis
        
    except Exception as e:
        logging.error(f"데이터 분석 오류: {e}")
        return {'version': version, 'error': str(e)}

# ==================== AI 분석 함수 ====================

@st.cache_data(ttl=300, show_spinner=False)  # 5분 캐시 (동일 질문 재요청 방지)
def ask_professional_scheduler(question, df_json, analysis_json, comparison_json=None):
    """Potens AI API 호출 (캐싱 적용)"""
    api_url = "https://ai.potens.ai/api/chat"
    headers = {
        "Content-Type": "application/json",
        "Authorization": f"Bearer {POTENS_API_KEY}"
    }
    
    try:
        # JSON 문자열을 다시 딕셔너리로 변환
        df_dict = json.loads(df_json)
        analysis = json.loads(analysis_json)
        
        version = analysis.get('version', '2차')
        
        # 데이터 요약 (구조화된 형식)
        data_summary = {
            "총_데이터_건수": len(df_dict.get('due_date', {})),
            "분석_기간": f"{min(df_dict.get('due_date', {}).values())} ~ {max(df_dict.get('due_date', {}).values())}" if df_dict.get('due_date') else "N/A",
            "라인별_평균_가동률": {
                line: f"{info.get('avg_utilization', 0):.1f}%"
                for line, info in analysis.items()
                if line.startswith("조립")
            }
        }
        
        # 변경사항 요약
        change_summary = ""
        if comparison_json:
            comp_dict = json.loads(comparison_json)
            changed_count = sum(1 for v in comp_dict.get('changed', {}).values() if v)
            if changed_count > 0:
                change_summary = f"\n\n[📊 0차 대비 2차 변경사항]\n총 {changed_count}건 변경됨"
        
        # 위반사항 요약
        violations = []
        if analysis.get('fan_violations'):
            violations.append(f"⚠️ FAN 요일규칙 위반: {len(analysis['fan_violations'])}건")
        if analysis.get('flange_violations'):
            violations.append(f"⚠️ FLANGE 요일규칙 위반: {len(analysis['flange_violations'])}건")
        
        violations_summary = "\n".join(violations) if violations else "✅ 요일 규칙 위반 없음"
        
        # 간결한 시스템 프롬프트
        system_prompt = f"""당신은 자동차 부품 조립라인의 '수석 생산 스케줄러'입니다.

[핵심 규칙]
1. **[PLT] 태그**: remark에 '[PLT]' 포함 시 → 배수 무시, 박스/로트 단위 우선
2. **요일 규칙**: 조립2 - FAN(월수금), FLANGE(화목)
3. **CAPA 제약**: 조립1(3000), 조립2(2500), 조립3(2000) - 90% 목표

[현재 데이터 요약 - {version}]
{json.dumps(data_summary, ensure_ascii=False, indent=2)}
{change_summary}

[위반사항]
{violations_summary}

[출력 형식]
1. **상황 진단** (2-3줄 요약)
2. **대안 1**: [구체적 조치]
3. **대안 2**: [예비 방안]
4. **대안 3**: [장기 개선안]
5. **즉시 조치 사항** (있을 경우)

**중요**: [PLT] 태그가 있는 항목은 정상으로 간주하고 수정 제안하지 마세요.
"""
        
        payload = {
            "prompt": f"{system_prompt}\n\n[긴급 요청]\n{question}"
        }
        
        logging.info(f"AI 요청 전송: {question[:50]}...")
        
        response = requests.post(
            api_url,
            headers=headers,
            json=payload,
            timeout=90
        )
        
        if response.status_code == 200:
            result = response.json().get('message', '응답 형식 오류')
            logging.info("AI 응답 수신 성공")
            return result
        else:
            error_msg = f"❌ API 오류 (코드: {response.status_code}): {response.text}"
            logging.error(error_msg)
            return error_msg
            
    except requests.Timeout:
        error_msg = "⏱️ 요청 시간 초과 (90초). 잠시 후 다시 시도해주세요."
        logging.error(error_msg)
        return error_msg
    except Exception as e:
        error_msg = f"❌ 요청 실패: {str(e)}"
        logging.error(error_msg)
        return error_msg

# ==================== 대시보드 컴포넌트 ====================

def render_dashboard(analysis, target_date):
    """CAPA 현황 대시보드"""
    st.subheader(f"📈 라인별 CAPA 현황 ({target_date})")
    
    cols = st.columns(3)
    
    for idx, line in enumerate(["조립1", "조립2", "조립3"]):
        if line in analysis:
            info = analysis[line]
            total_production = sum(info['daily_production'].values())
            max_possible = info['max_capa'] * len(info['daily_production'])
            utilization = (total_production / max_possible * 100) if max_possible > 0 else 0
            
            with cols[idx]:
                st.metric(
                    label=f"{line}",
                    value=f"{total_production:,}개",
                    delta=f"가동률 {utilization:.1f}%"
                )
                
                # CAPA 초과 날짜
                over_days = info.get('over_capacity_days', {})
                if over_days:
                    st.warning(f"⚠️ CAPA 초과: {len(over_days)}일")
                    for date, qty in list(over_days.items())[:3]:
                        st.caption(f"  • {date}: {qty:,}개")
                else:
                    st.success("✅ CAPA 정상")

def render_violations(analysis):
    """요일 규칙 위반 현황"""
    fan_violations = analysis.get('fan_violations', [])
    flange_violations = analysis.get('flange_violations', [])
    
    if fan_violations or flange_violations:
        st.subheader("⚠️ 요일 규칙 위반 현황")
        
        col1, col2 = st.columns(2)
        
        with col1:
            if fan_violations:
                st.error(f"FAN 위반: {len(fan_violations)}건")
                for v in fan_violations[:5]:
                    # [PLT] 태그 확인
                    is_plt = '[PLT]' in str(v.get('remark', ''))
                    icon = "📦" if is_plt else "⚠️"
                    st.caption(f"{icon} {v['due_date']} ({v['weekday_kr']}): {v['product_name']}")
        
        with col2:
            if flange_violations:
                st.error(f"FLANGE 위반: {len(flange_violations)}건")
                for v in flange_violations[:5]:
                    is_plt = '[PLT]' in str(v.get('remark', ''))
                    icon = "📦" if is_plt else "⚠️"
                    st.caption(f"{icon} {v['due_date']} ({v['weekday_kr']}): {v['product_name']}")

# ==================== UI 구성 ====================

def main():
    st.set_page_config(
        page_title="수석 스케줄러 AI 관제 센터",
        page_icon="👨‍✈️",
        layout="wide",
        initial_sidebar_state="expanded"
    )
    
    st.title("👨‍✈️ 수석 스케줄러 AI 통합 제어 센터")
    st.caption("Real-time Production Scheduling with AI Assistant")
    
    # 세션 상태 초기화
    if "messages" not in st.session_state:
        st.session_state.messages = []
    if "current_date" not in st.session_state:
        st.session_state.current_date = None
    if "current_df" not in st.session_state:
        st.session_state.current_df = None
    
    # ==================== 사이드바 ====================
    with st.sidebar:
        st.header("⚙️ 설정")
        
        # 날짜 선택
        selected_date = st.date_input(
            "📅 분석 기준일",
            value=datetime.now(),
            min_value=datetime(2025, 1, 1),
            max_value=datetime(2026, 12, 31)
        )
        formatted_date = selected_date.strftime('%Y-%m-%d')
        
        # 버전 선택
        version_option = st.radio(
            "📂 분석 대상",
            options=['2차 (실제 조정본)', '0차 vs 2차 비교'],
            help="2차: 최종 조정된 계획 / 비교: 초기 계획 대비 변경사항"
        )
        
        # 빠른 질문 템플릿
        st.divider()
        st.subheader("💡 빠른 질문")
        
        quick_questions = {
            "CAPA 초과 분석": f"{selected_date.month}/{selected_date.day} CAPA 초과 날짜 분석해줘",
            "요일 규칙 위반": f"{selected_date.month}/{selected_date.day} 요일 규칙 위반 확인해줘",
            "변경사항 요약": f"{selected_date.month}/{selected_date.day} 0차 대비 주요 변경사항 요약해줘",
            "긴급 조치 필요": f"{selected_date.month}/{selected_date.day} 긴급 조치가 필요한 항목 알려줘"
        }
        
        for label, question in quick_questions.items():
            if st.button(label, use_container_width=True):
                st.session_state.quick_question = question
        
        # 통계 정보
        st.divider()
        st.caption(f"🔄 마지막 업데이트: {datetime.now().strftime('%H:%M:%S')}")
        st.caption(f"💾 캐시 상태: {'활성' if st.session_state.current_df is not None else '비활성'}")
    
    # ==================== 메인 영역 ====================
    
    # 대시보드 영역
    dashboard_container = st.container()
    
    # 채팅 및 데이터 영역
    left_col, right_col = st.columns([1, 1.2])
    
    with left_col:
        st.subheader("💬 AI 상담 창구")
        
        # 기존 메시지 표시
        chat_container = st.container()
        with chat_container:
            for msg in st.session_state.messages:
                with st.chat_message(msg["role"]):
                    st.markdown(msg["content"])
        
        # 채팅 입력
        prompt = st.chat_input("질문을 입력하세요 (예: 8/14 배수 이슈 분석해줘)")
        
        # 빠른 질문 처리
        if "quick_question" in st.session_state:
            prompt = st.session_state.quick_question
            del st.session_state.quick_question
        
        if prompt:
            # 사용자 메시지 추가
            st.session_state.messages.append({"role": "user", "content": prompt})
            with st.chat_message("user"):
                st.markdown(prompt)
            
            # 날짜 추출
            date_match = re.search(r'(\d{1,2})/(\d{1,2})', prompt)
            if date_match:
                date_val = f"{date_match.group(1)}/{date_match.group(2)}"
                target_date = parse_date(date_val)
            else:
                target_date = formatted_date
            
            if not target_date:
                with st.chat_message("assistant"):
                    error_msg = "❌ 날짜 형식을 인식할 수 없습니다. (예: 8/14)"
                    st.error(error_msg)
                    st.session_state.messages.append({"role": "assistant", "content": error_msg})
            else:
                with st.spinner("🤖 AI가 분석 중입니다..."):
                    try:
                        if '비교' in version_option:
                            # 0차 vs 2차 비교 모드
                            df_0 = fetch_production_data(target_date, version='0차')
                            df_2 = fetch_production_data(target_date, version='2차')
                            
                            if df_0 is None or df_2 is None:
                                with st.chat_message("assistant"):
                                    error_msg = f"❌ {date_val} 데이터를 찾을 수 없습니다.\n0차 데이터: {'있음' if df_0 is not None else '없음'}\n2차 데이터: {'있음' if df_2 is not None else '없음'}"
                                    st.error(error_msg)
                                    st.session_state.messages.append({"role": "assistant", "content": error_msg})
                            else:
                                comp_df = compare_versions(df_0, df_2)
                                analysis = analyze_data(df_2)
                                
                                # DataFrame을 JSON 문자열로 변환 (캐싱을 위해)
                                df_json = df_2.to_json(orient='columns')
                                analysis_json = json.dumps(analysis, ensure_ascii=False)
                                comp_json = comp_df.to_json(orient='columns')
                                
                                answer = ask_professional_scheduler(prompt, df_json, analysis_json, comp_json)
                                
                                with st.chat_message("assistant"):
                                    st.markdown(answer)
                                    st.session_state.messages.append({"role": "assistant", "content": answer})
                                
                                # 데이터 저장
                                st.session_state.current_df = comp_df
                                st.session_state.current_analysis = analysis
                                st.session_state.current_date = target_date
                                
                                # 오른쪽 패널에 데이터 표시
                                with right_col:
                                    st.subheader(f"📊 비교 데이터 ({date_val})")
                                    
                                    # 변경사항 필터
                                    show_changed_only = st.checkbox("변경된 항목만 보기", value=True)
                                    display_df = comp_df[comp_df['changed']] if show_changed_only else comp_df
                                    
                                    st.dataframe(
                                        display_df[['due_date', 'line', 'product_name', 'quantity_0차', 'quantity_2차', 'qty_diff', 'status', 'remark', 'worker_memo']],
                                        use_container_width=True,
                                        height=400
                                    )
                                    
                                    # 다운로드 버튼
                                    excel_data = convert_df_to_excel(display_df)
                                    st.download_button(
                                        label="📥 Excel 다운로드",
                                        data=excel_data,
                                        file_name=f"schedule_comparison_{target_date}.xlsx",
                                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                                    )
                        
                        else:
                            # 2차 단독 분석 모드
                            df = fetch_production_data(target_date, version='2차')
                            
                            if df is None:
                                with st.chat_message("assistant"):
                                    error_msg = f"❌ {date_val} 데이터를 찾을 수 없습니다.\n• 날짜를 확인해주세요\n• 해당 기간에 계획이 없을 수 있습니다"
                                    st.error(error_msg)
                                    st.session_state.messages.append({"role": "assistant", "content": error_msg})
                            else:
                                analysis = analyze_data(df)
                                
                                # DataFrame을 JSON 문자열로 변환
                                df_json = df.to_json(orient='columns')
                                analysis_json = json.dumps(analysis, ensure_ascii=False)
                                
                                answer = ask_professional_scheduler(prompt, df_json, analysis_json)
                                
                                with st.chat_message("assistant"):
                                    st.markdown(answer)
                                    st.session_state.messages.append({"role": "assistant", "content": answer})
                                
                                # 데이터 저장
                                st.session_state.current_df = df
                                st.session_state.current_analysis = analysis
                                st.session_state.current_date = target_date
                                
                                # 오른쪽 패널에 데이터 표시
                                with right_col:
                                    st.subheader(f"📊 2차 데이터 상세 ({date_val})")
                                    
                                    # 필터 옵션
                                    filter_line = st.multiselect(
                                        "라인 필터",
                                        options=['조립1', '조립2', '조립3'],
                                        default=['조립1', '조립2', '조립3']
                                    )
                                    
                                    filtered_df = df[df['line'].isin(filter_line)]
                                    
                                    st.dataframe(
                                        filtered_df[['due_date', 'line', 'product_name', 'product_type', 'quantity', 'plt', 'status', 'remark']],
                                        use_container_width=True,
                                        height=400
                                    )
                                    
                                    # 다운로드 버튼
                                    excel_data = convert_df_to_excel(filtered_df)
                                    st.download_button(
                                        label="📥 Excel 다운로드",
                                        data=excel_data,
                                        file_name=f"schedule_detail_{target_date}.xlsx",
                                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                                    )
                                
                                # 대시보드 업데이트
                                with dashboard_container:
                                    render_dashboard(analysis, date_val)
                                    render_violations(analysis)
                    
                    except Exception as e:
                        logging.error(f"처리 중 오류 발생: {e}")
                        with st.chat_message("assistant"):
                            error_msg = f"❌ 처리 중 오류가 발생했습니다: {str(e)}"
                            st.error(error_msg)
                            st.session_state.messages.append({"role": "assistant", "content": error_msg})
    
    # 초기 대시보드 표시 (데이터가 있을 경우)
    if st.session_state.current_df is not None and st.session_state.current_analysis is not None:
        with dashboard_container:
            render_dashboard(st.session_state.current_analysis, st.session_state.current_date)
            render_violations(st.session_state.current_analysis)

if __name__ == "__main__":
    main()
