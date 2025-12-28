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

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('scheduler_app.log'),
        logging.StreamHandler()
    ]
)

SUPABASE_URL = os.getenv("SUPABASE_URL", "https://suaajrdahixouinbfcfo.supabase.co")
SUPABASE_KEY = os.getenv("SUPABASE_KEY", "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InN1YWFqcmRhaGl4b3VpbmJmY2ZvIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NjYzMTk4NzAsImV4cCI6MjA4MTg5NTg3MH0.Ic4izQY-ihIw75jKh9iJicZvuZ4gCRs4OH3rCGyo0Zk")
POTENS_API_KEY = os.getenv("POTENS_API_KEY", "qD2gfuVAkMJexDAcFb5GnEb1SZksTs7o")

try:
    supabase: Client = create_client(SUPABASE_URL, SUPABASE_KEY)
    logging.info("Supabase 연결 성공")
except Exception as e:
    logging.error(f"Supabase 연결 실패: {e}")

TABLE_NAME = "pattern_learning2"

CAPA_INFO = {
    "조립1": 3000,
    "조립2": 2500,
    "조립3": 2000
}

def parse_date(date_str):
    try:
        if '/' in date_str:
            parts = date_str.split('/')
            month = parts[0].zfill(2)
            day = parts[1].zfill(2)
            current_year = datetime.now().year
            
            test_date = datetime.strptime(f"{current_year}-{month}-{day}", '%Y-%m-%d')
            if test_date < datetime.now() - timedelta(days=180):
                current_year += 1
            
            return f"{current_year}-{month}-{day}"
        return date_str
    except Exception as e:
        logging.error(f"날짜 파싱 오류: {date_str} - {e}")
        return None

def get_date_range(target_date, days_before=0, days_after=0):
    try:
        dt = datetime.strptime(target_date, '%Y-%m-%d')
        start = (dt - timedelta(days=days_before)).strftime('%Y-%m-%d')
        end = (dt + timedelta(days=days_after)).strftime('%Y-%m-%d')
        return start, end
    except Exception as e:
        logging.error(f"날짜 범위 계산 오류: {e}")
        return None, None

def convert_df_to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Schedule')
        workbook = writer.book
        worksheet = writer.sheets['Schedule']
        header_format = workbook.add_format({
            'bold': True,
            'bg_color': '#4472C4',
            'font_color': 'white',
            'border': 1
        })
        for col_num, value in enumerate(df.columns.values):
            worksheet.write(0, col_num, value, header_format)
            worksheet.set_column(col_num, col_num, 15)
    return output.getvalue()

@st.cache_data(ttl=600, show_spinner=False)
def fetch_production_data(target_date, version='2차', date_type='due_date'):
    try:
        start_date, end_date = get_date_range(target_date)
        if not start_date:
            return None
        
        logging.info(f"DB 조회: {version}, {date_type}, {start_date}~{end_date}")
        
        response = supabase.table(TABLE_NAME)\
            .select("*")\
            .eq("version", version)\
            .eq(date_type, target_date)\
            .order("due_date", desc=False)\
            .execute()
        
        if response.data:
            df = pd.DataFrame(response.data)
            if 'id' in df.columns:
                df = df.drop_duplicates(subset=['id'])
            else:
                df = df.drop_duplicates()
            
            logging.info(f"조회 성공: {len(df)}건")
            return df
        return None
    except Exception as e:
        logging.error(f"DB 조회 오류: {e}")
        return None

@st.cache_data(ttl=3600, show_spinner=False)
def fetch_historical_data(days=90):
    try:
        end_date = datetime.now().strftime('%Y-%m-%d')
        start_date = (datetime.now() - timedelta(days=days)).strftime('%Y-%m-%d')
        
        response = supabase.table(TABLE_NAME)\
            .select("*")\
            .gte("due_date", start_date)\
            .lte("due_date", end_date)\
            .order("due_date", desc=False)\
            .execute()
        
        if response.data:
            return pd.DataFrame(response.data)
        return None
    except Exception as e:
        logging.error(f"과거 데이터 조회 오류: {e}")
        return None

def detect_issue_type(question):
    question_lower = question.lower()
    if any(keyword in question_lower for keyword in ['capa', '초과', '가동률', '부하']):
        return 'CAPA초과'
    elif any(keyword in question_lower for keyword in ['요일', 'fan', 'flange', '규칙']):
        return '요일위반'
    elif any(keyword in question_lower for keyword in ['배수', 'plt', '팔레트']):
        return '배수문제'
    elif any(keyword in question_lower for keyword in ['변경', '수정', '조정']):
        return '계획변경'
    else:
        return '일반문의'

def find_similar_cases(historical_df, issue_type, target_date):
    if historical_df is None or historical_df.empty:
        return None
    
    try:
        similar_cases = []
        
        if issue_type == 'CAPA초과':
            for line, info in CAPA_INFO.items():
                line_data = historical_df[historical_df['line'] == line]
                daily_sum = line_data.groupby('due_date')['quantity'].sum()
                over_cases = daily_sum[daily_sum > info * 0.9]
                
                for date, qty in over_cases.items():
                    case_data = historical_df[
                        (historical_df['due_date'] == date) & 
                        (historical_df['line'] == line)
                    ]
                    if not case_data.empty:
                        similar_cases.append({
                            'date': date,
                            'line': line,
                            'quantity': qty,
                            'remark': case_data['remark'].iloc[0] if 'remark' in case_data.columns else '',
                            'worker_memo': case_data['worker_memo'].iloc[0] if 'worker_memo' in case_data.columns else ''
                        })
        
        elif issue_type == '요일위반':
            historical_df['weekday'] = pd.to_datetime(historical_df['due_date']).dt.dayofweek
            fan_violations = historical_df[
                (historical_df['product_type'] == 'FAN') & 
                (~historical_df['weekday'].isin([0, 2, 4]))
            ]
            for _, row in fan_violations.head(5).iterrows():
                similar_cases.append({
                    'date': row['due_date'],
                    'line': row['line'],
                    'product': row['product_name'],
                    'type': 'FAN 요일위반',
                    'remark': row.get('remark', ''),
                    'worker_memo': row.get('worker_memo', '')
                })
        
        elif issue_type == '배수문제':
            plt_cases = historical_df[
                historical_df['remark'].str.contains('\[PLT\]', na=False, regex=True)
            ]
            for _, row in plt_cases.head(5).iterrows():
                similar_cases.append({
                    'date': row['due_date'],
                    'product': row['product_name'],
                    'quantity': row['quantity'],
                    'plt': row.get('plt', ''),
                    'remark': row.get('remark', ''),
                    'worker_memo': row.get('worker_memo', '')
                })
        
        return similar_cases[:5] if similar_cases else None
    except Exception as e:
        logging.error(f"유사 사례 검색 오류: {e}")
        return None

def compare_versions(df_base, df_final):
    if df_base is None or df_final is None or df_base.empty or df_final.empty:
        return None
    
    try:
        base_version = df_base['version'].iloc[0] if 'version' in df_base.columns and len(df_base) > 0 else '0차'
        
        merged = pd.merge(
            df_base[['due_date', 'line', 'product_name', 'product_type', 'quantity']],
            df_final[['due_date', 'line', 'product_name', 'product_type', 'quantity', 'status', 'remark', 'worker_memo']],
            on=['due_date', 'line', 'product_name', 'product_type'],
            how='outer',
            suffixes=(f'_{base_version}', '_2차')
        )
        
        merged[f'quantity_{base_version}'] = merged[f'quantity_{base_version}'].fillna(0)
        merged['quantity_2차'] = merged['quantity_2차'].fillna(0)
        merged['qty_diff'] = merged['quantity_2차'] - merged[f'quantity_{base_version}']
        merged['changed'] = merged['qty_diff'] != 0
        merged['base_version'] = base_version
        
        logging.info(f"{base_version} vs 2차 비교 완료: {len(merged)}건, 변경 {merged['changed'].sum()}건")
        return merged
    except Exception as e:
        logging.error(f"버전 비교 오류: {e}")
        return None

def analyze_data(df, version='2차'):
    try:
        analysis = {'version': version}
        
        df['plan_date_dt'] = pd.to_datetime(df['due_date'])
        df['weekday_kr'] = df['plan_date_dt'].dt.strftime('%A').map({
            'Monday': '월', 'Tuesday': '화', 'Wednesday': '수',
            'Thursday': '목', 'Friday': '금', 'Saturday': '토', 'Sunday': '일'
        })
        
        qty_col = 'quantity'
        
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
        
        bergstrom_data = df[df['product_name'].str.contains('BERGSTROM', case=False, na=False)]
        bergstrom_days = bergstrom_data.groupby('due_date')[qty_col].sum().to_dict()
        
        line2_data = df[df['line'] == '조립2'].copy()
        fan_data = line2_data[line2_data['product_type'] == 'FAN']
        fan_wrong = fan_data[~fan_data['weekday_kr'].isin(['월', '수', '금'])]
        flange_data = line2_data[line2_data['product_type'] == 'FLANGE']
        flange_wrong = flange_data[~flange_data['weekday_kr'].isin(['화', '목'])]
        line2_daily_products = line2_data.groupby('due_date')['product_name'].nunique()
        
        analysis['bergstrom_days'] = bergstrom_days
        analysis['fan_violations'] = fan_wrong[['due_date', 'product_name', qty_col, 'weekday_kr', 'remark']].to_dict('records')
        analysis['flange_violations'] = flange_wrong[['due_date', 'product_name', qty_col, 'weekday_kr', 'remark']].to_dict('records')
        analysis['line2_over_5products'] = line2_daily_products[line2_daily_products > 5].to_dict()
        analysis['status_summary'] = df['status'].value_counts().to_dict()
        
        return analysis
    except Exception as e:
        logging.error(f"데이터 분석 오류: {e}")
        return {'version': version, 'error': str(e)}

@st.cache_data(ttl=300, show_spinner=False)
def ask_professional_scheduler(question, df_json, analysis_json, comparison_json=None, historical_cases_json=None, original_plan_json=None):
    api_url = "https://ai.potens.ai/api/chat"
    headers = {
        "Content-Type": "application/json",
        "Authorization": f"Bearer {POTENS_API_KEY}"
    }
    
    try:
        df_dict = json.loads(df_json)
        analysis = json.loads(analysis_json)
        version = analysis.get('version', '2차')
        
        data_summary = {
            "총_데이터_건수": len(df_dict.get('due_date', {})),
            "분석_기간": f"{min(df_dict.get('due_date', {}).values())} ~ {max(df_dict.get('due_date', {}).values())}" if df_dict.get('due_date') else "N/A"
        }
        
        original_plan_summary = ""
        if original_plan_json:
            original_df = json.loads(original_plan_json)
            original_plan_summary = f"\n\n[📋 1차 원래 계획]\n총 {len(original_df.get('due_date', {}))}건"
        
        historical_summary = ""
        if historical_cases_json:
            historical_cases = json.loads(historical_cases_json)
            if historical_cases:
                historical_summary = f"\n\n[🔍 과거 유사 사례 {len(historical_cases)}건]\n"
                for idx, case in enumerate(historical_cases[:3], 1):
                    historical_summary += f"\n{idx}. {case.get('date', 'N/A')}"
                    if 'line' in case:
                        historical_summary += f" | {case['line']}"
                    if 'quantity' in case:
                        historical_summary += f" | {case['quantity']}개"
                    if case.get('remark'):
                        historical_summary += f"\n   조치: {case['remark'][:50]}"
        
        change_summary = ""
        if comparison_json:
            comp_dict = json.loads(comparison_json)
            changed_count = sum(1 for v in comp_dict.get('changed', {}).values() if v)
            if changed_count > 0:
                change_summary = f"\n\n[📊 변경사항]\n총 {changed_count}건 변경됨"
        
        violations = []
        if analysis.get('fan_violations'):
            violations.append(f"⚠️ FAN 요일규칙 위반: {len(analysis['fan_violations'])}건")
        if analysis.get('flange_violations'):
            violations.append(f"⚠️ FLANGE 요일규칙 위반: {len(analysis['flange_violations'])}건")
        violations_summary = "\n".join(violations) if violations else "✅ 요일 규칙 위반 없음"
        
        system_prompt = f"""당신은 자동차 부품 조립라인의 '수석 생산 스케줄러'입니다.
과거 데이터를 학습하여 실제 해결 사례 기반으로 조언하세요.

[핵심 규칙]
1. [PLT] 태그: remark에 '[PLT]' 포함 시 → 배수 무시
2. 요일 규칙: 조립2 - FAN(월수금), FLANGE(화목)
3. CAPA 제약: 조립1(3000), 조립2(2500), 조립3(2000) - 90% 목표

[현재 데이터 - {version}]
{json.dumps(data_summary, ensure_ascii=False)}
{original_plan_summary}
{change_summary}
{violations_summary}
{historical_summary}

[출력 형식]
1. 상황 진단
2. 과거 사례 참고
3. 대안 1, 2, 3
4. 즉시 조치 사항
"""
        
        payload = {"prompt": f"{system_prompt}\n\n[긴급 요청]\n{question}"}
        response = requests.post(api_url, headers=headers, json=payload, timeout=90)
        
        if response.status_code == 200:
            return response.json().get('message', '응답 형식 오류')
        return f"❌ API 오류: {response.status_code}"
    except Exception as e:
        return f"❌ 요청 실패: {str(e)}"

def render_dashboard(analysis, target_date, highlight_date=None):
    st.subheader(f"📈 일일 CAPA 초과 현황")
    cols = st.columns(3)
    
    for idx, line in enumerate(["조립1", "조립2", "조립3"]):
        if line in analysis:
            info = analysis[line]
            over_days = info.get('over_capacity_days', {})
            target_90 = info['target_90']
            
            with cols[idx]:
                if over_days:
                    st.error(f"**{line}**")
                    st.caption(f"⚠️ CAPA 초과: {len(over_days)}일")
                    st.caption(f"기준: {target_90:,}개 (90%)")
                    
                    for date, qty in sorted(over_days.items()):
                        over_percent = ((qty - target_90) / target_90 * 100)
                        if highlight_date and date == highlight_date:
                            st.warning(f"🎯 **{date}: {qty:,}개 (+{over_percent:.0f}%)**")
                        else:
                            st.caption(f"• {date}: {qty:,}개 (+{over_percent:.0f}%)")
                else:
                    st.success(f"**{line}**")
                    st.caption(f"✅ 모든 날짜 CAPA 정상")
                    st.caption(f"기준: {target_90:,}개 (90%)")

def render_violations(analysis):
    fan_violations = analysis.get('fan_violations', [])
    flange_violations = analysis.get('flange_violations', [])
    
    if fan_violations or flange_violations:
        st.subheader("⚠️ 요일 규칙 위반 현황")
        col1, col2 = st.columns(2)
        
        with col1:
            if fan_violations:
                st.error(f"FAN 위반: {len(fan_violations)}건")
                for v in fan_violations[:5]:
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

def render_historical_cases(cases):
    if not cases:
        return
    
    st.subheader("🔍 과거 유사 사례")
    for idx, case in enumerate(cases, 1):
        with st.expander(f"사례 {idx}: {case.get('date', 'N/A')}"):
            cols = st.columns(2)
            with cols[0]:
                if 'line' in case:
                    st.write(f"**라인**: {case['line']}")
                if 'quantity' in case:
                    st.write(f"**수량**: {case['quantity']:,}개")
                if 'product' in case:
                    st.write(f"**제품**: {case['product']}")
            with cols[1]:
                if case.get('remark'):
                    st.write(f"**조치사항**: {case['remark']}")
                if case.get('worker_memo'):
                    st.write(f"**담당자 메모**: {case['worker_memo']}")

def main():
    st.set_page_config(
        page_title="수석 스케줄러 AI 관제 센터",
        page_icon="👨‍✈️",
        layout="wide",
        initial_sidebar_state="expanded"
    )
    
    st.title("👨‍✈️ 수석 스케줄러 AI 통합 제어 센터")
    st.caption("Real-time Production Scheduling with AI & Historical Pattern Learning")
    
    if "messages" not in st.session_state:
        st.session_state.messages = []
    if "current_date" not in st.session_state:
        st.session_state.current_date = None
    if "current_df" not in st.session_state:
        st.session_state.current_df = None
    
    with st.spinner("과거 데이터 로딩 중..."):
        historical_data = fetch_historical_data(days=90)
    
    with st.sidebar:
        st.header("⚙️ 설정")
        
        selected_date = st.date_input(
            "📅 분석 기준일",
            value=datetime.now(),
            min_value=datetime(2025, 1, 1),
            max_value=datetime(2026, 12, 31)
        )
        formatted_date = selected_date.strftime('%Y-%m-%d')
        
        version_option = st.radio(
            "📂 분석 대상",
            options=['2차 (최종 조정본)', '변경사항 비교 (과거 패턴 포함)'],
            help="2차: 최종 조정 / 비교: 1차/0차 대비 변경사항 + 과거 사례"
        )
        
        st.divider()
        st.subheader("💡 빠른 질문")
        
        quick_questions = {
            "CAPA 초과 분석": f"{selected_date.month}/{selected_date.day} CAPA 초과 분석해줘",
            "요일 규칙 위반": f"{selected_date.month}/{selected_date.day} 요일 규칙 위반 확인해줘",
            "변경사항 분석": f"{selected_date.month}/{selected_date.day} 변경사항 분석해줘",
            "긴급 조치 필요": f"{selected_date.month}/{selected_date.day} 긴급 조치 알려줘"
        }
        
        for label, question in quick_questions.items():
            if st.button(label, use_container_width=True):
                st.session_state.quick_question = question
        
        st.divider()
        st.caption(f"🔄 업데이트: {datetime.now().strftime('%H:%M:%S')}")
        st.caption(f"📚 과거 데이터: {len(historical_data) if historical_data is not None else 0}건")
    
    dashboard_container = st.container()
    left_col, right_col = st.columns([1, 1.2])
    
    with left_col:
        st.subheader("💬 AI 상담 창구")
        
        for msg in st.session_state.messages:
            with st.chat_message(msg["role"]):
                st.markdown(msg["content"])
        
        prompt = st.chat_input("질문을 입력하세요 (예: 11/14 조립1 CAPA 초과 어떻게 해?)")
        
        if "quick_question" in st.session_state:
            prompt = st.session_state.quick_question
            del st.session_state.quick_question
        
        if prompt:
            is_production_query = any(k in prompt for k in ['생산', '가동'])
            search_col = 'production_date' if is_production_query else 'due_date'
            
            st.session_state.messages.append({"role": "user", "content": prompt})
            with st.chat_message("user"):
                st.markdown(prompt)
            
            date_match = re.search(r'(\d{1,2})[\./](\d{1,2})', prompt)
            if date_match:
                date_val = f"{date_match.group(1)}/{date_match.group(2)}"
                target_date = parse_date(date_val)
            else:
                target_date = formatted_date
                date_val = f"{selected_date.month}/{selected_date.day}"
            
            if not target_date:
                with st.chat_message("assistant"):
                    st.error("❌ 날짜 형식 오류")
            else:
                with st.spinner("🤖 AI 분석 중..."):
                    try:
                        df_2 = fetch_production_data(target_date, version='2차', date_type=search_col)
                        
                        if df_2 is None or df_2.empty:
                            alt_col = 'due_date' if search_col == 'production_date' else 'production_date'
                            df_2 = fetch_production_data(target_date, version='2차', date_type=alt_col)
                            if df_2 is not None and not df_2.empty:
                                st.info(f"💡 {alt_col} 기준으로 조회했습니다.")
                                search_col = alt_col
                        
                        if df_2 is None or df_2.empty:
                            with st.chat_message("assistant"):
                                st.error(f"❌ {date_val} 데이터 없음")
                        else:
                            issue_type = detect_issue_type(prompt)
                            similar_cases = find_similar_cases(historical_data, issue_type, target_date)
                            
                            if '비교' in version_option:
                                df_1 = fetch_production_data(target_date, version='1차', date_type=search_col)
                                if df_1 is None or df_1.empty:
                                    df_base = fetch_production_data(target_date, version='0차', date_type=search_col)
                                    comparison_type = "0차 vs 2차"
                                else:
                                    df_base = df_1
                                    comparison_type = "1차 vs 2차"
                                
                                if df_base is None or df_base.empty:
                                    analysis = analyze_data(df_2)
                                    answer = ask_professional_scheduler(
                                        prompt, 
                                        df_2.to_json(orient='columns'), 
                                        json.dumps(analysis, ensure_ascii=False),
                                        None,
                                        json.dumps(similar_cases, ensure_ascii=False) if similar_cases else None,
                                        None
                                    )
                                    with st.chat_message("assistant"):
                                        st.warning("⚠️ 비교 기준 없음, 2차만 분석")
                                        st.markdown(answer)
                                        st.session_state.messages.append({"role": "assistant", "content": answer})
                                else:
                                    comp_df = compare_versions(df_base, df_2)
                                    
                                    if comp_df is None:
                                        with st.chat_message("assistant"):
                                            st.error("❌ 비교 실패")
                                    elif comp_df.empty:
                                        with st.chat_message("assistant"):
                                            st.warning("⚠️ 비교 데이터 없음")
                                    else:
                                        analysis = analyze_data(df_2)
                                        base_version = comp_df['base_version'].iloc[0] if 'base_version' in comp_df.columns else '0차'
                                        
                                        answer = ask_professional_scheduler(
                                            prompt, 
                                            df_2.to_json(orient='columns'), 
                                            json.dumps(analysis, ensure_ascii=False), 
                                            comp_df.to_json(orient='columns'),
                                            json.dumps(similar_cases, ensure_ascii=False) if similar_cases else None,
                                            df_base.to_json(orient='columns')
                                        )
                                        
                                        with st.chat_message("assistant"):
                                            st.markdown(f"**📊 {comparison_type}**\n\n")
                                            st.markdown(answer)
                                            st.session_state.messages.append({"role": "assistant", "content": answer})
                                        
                                        st.session_state.current_df = comp_df
                                        st.session_state.current_analysis = analysis
                                        st.session_state.current_date = target_date
                                        
                                        with right_col:
                                            st.subheader(f"📊 {comparison_type} ({date_val})")
                                            show_changed = st.checkbox("변경된 항목만", value=True)
                                            display_df = comp_df[comp_df['changed']] if show_changed else comp_df
                                            
                                            display_cols = ['due_date', 'line', 'product_name', 
                                                          f'quantity_{base_version}', 'quantity_2차', 
                                                          'qty_diff', 'status', 'remark']
                                            
                                            st.dataframe(display_df[display_cols], use_container_width=True, height=300)
                                            
                                            if similar_cases:
                                                render_historical_cases(similar_cases)
                                            
                                            st.download_button(
                                                "📥 Excel 다운로드",
                                                convert_df_to_excel(display_df),
                                                f"comparison_{target_date}.xlsx",
                                                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                                            )
                            else:
                                df_1 = fetch_production_data(target_date, version='1차', date_type=search_col)
                                if df_1 is None or df_1.empty:
                                    df_1 = fetch_production_data(target_date, version='0차', date_type=search_col)
                                
                                analysis = analyze_data(df_2)
                                
                                answer = ask_professional_scheduler(
                                    prompt, 
                                    df_2.to_json(orient='columns'), 
                                    json.dumps(analysis, ensure_ascii=False),
                                    None,
                                    json.dumps(similar_cases, ensure_ascii=False) if similar_cases else None,
                                    df_1.to_json(orient='columns') if df_1 is not None else None
                                )
                                
                                with st.chat_message("assistant"):
                                    st.markdown(answer)
                                    st.session_state.messages.append({"role": "assistant", "content": answer})
                                
                                st.session_state.current_df = df_2
                                st.session_state.current_analysis = analysis
                                st.session_state.current_date = target_date
                                
                                with right_col:
                                    st.subheader(f"📊 2차 데이터 ({date_val})")
                                    
                                    filter_line = st.multiselect(
                                        "라인 필터",
                                        ['조립1', '조립2', '조립3'],
                                        ['조립1', '조립2', '조립3']
                                    )
                                    
                                    filtered_df = df_2[df_2['line'].isin(filter_line)]
                                    
                                    st.dataframe(
                                        filtered_df[['due_date', 'line', 'product_name', 'product_type', 'quantity', 'plt', 'status', 'remark']],
                                        use_container_width=True,
                                        height=300
                                    )
                                    
                                    if similar_cases:
                                        render_historical_cases(similar_cases)
                                    
                                    st.download_button(
                                        "📥 Excel 다운로드",
                                        convert_df_to_excel(filtered_df),
                                        f"schedule_{target_date}.xlsx",
                                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                                    )
                                
                                with dashboard_container:
                                    render_dashboard(analysis, date_val, highlight_date=target_date)
                                    render_violations(analysis)
                    
                    except Exception as e:
                        logging.error(f"오류: {e}")
                        with st.chat_message("assistant"):
                            st.error(f"❌ 오류: {str(e)}")
    
    if st.session_state.current_df is not None and st.session_state.current_analysis is not None:
        with dashboard_container:
            render_dashboard(st.session_state.current_analysis, st.session_state.current_date)
            render_violations(st.session_state.current_analysis)

if __name__ == "__main__":
    main()
