import streamlit as st
import pandas as pd
import requests
from datetime import datetime
import plotly.express as px

# 1. 페이지 기본 설정
st.set_page_config(page_title="식품안전 검사부적합 모니터링", layout="wide")

# 2. 스트림릿 시크릿에서 API 키 불러오기
try:
    API_KEY = st.secrets["FOOD_API_KEY"]
except KeyError:
    st.error("⚠️ Streamlit Secrets에 'FOOD_API_KEY'가 설정되지 않았습니다. .streamlit/secrets.toml 셋팅을 확인해주세요.")
    st.stop()

# 3. 공공 API 호출 함수 (최신 데이터 1000건 추출)
@st.cache_data(ttl=3600) # 1시간마다 데이터 갱신
def fetch_food_safety_data():
    # 식품안전나라 API 엔드포인트 구조 (I2710: 국내 검사부적합)
    api_service_code = "I2710"
    
    try:
        # 1차 호출: 전체 데이터 건수 파악 및 최신 1000건 가져오기
        # 식품안전나라 API는 /시작인덱스/종료인덱스 형태로 호출
        base_url = f"http://openapi.foodsafetykorea.go.kr/api/{API_KEY}/{api_service_code}/json/1/1000"
        
        res = requests.get(base_url)
        res.raise_for_status()
        data = res.json()
        
        if api_service_code in data and 'row' in data[api_service_code]:
            df = pd.DataFrame(data[api_service_code]['row'])
            
            # 컬럼명 직관적으로 변경 (API 실제 응답 키 기준 매핑)
            rename_dict = {
                'BSSH_NM': '업체명',
                'PRDUCT_NM': '제품명',
                'PRDCT_TYPE': '식품유형',
                'INSPCT_RSLT': '부적합항목',
                'MNFCTUR_DE': '제조일자',
                'DISTB_TMLMT': '유통기한',
                'RTRVL_DSUSE_SEQ': '회수폐기일련번호',
                'ADDR': '소재지'
            }
            df = df.rename(columns=lambda x: rename_dict.get(x, x))
            
            # 결측치 처리
            df = df.fillna("")
            
            # 최근 제조일자 기준으로 내림차순 정렬
            if '제조일자' in df.columns:
                df['제조일자'] = df['제조일자'].astype(str)
                df = df.sort_values(by="제조일자", ascending=False).reset_index(drop=True)
                
            return df
        else:
            return pd.DataFrame()

    except Exception as e:
        st.error(f"API 호출 중 오류 발생: {e}")
        return pd.DataFrame()

# 4. 기준 일시 설정
today = datetime.now()

# 5. 화면 UI 구성 (사이드바 및 메인)
st.sidebar.header("🔍 상세 검색 필터")

# 데이터 로딩 상태 표시
with st.spinner("식품의약품안전처 최신 부적합 데이터를 불러오는 중입니다..."):
    raw_df = fetch_food_safety_data()

if raw_df.empty:
    st.warning("조회된 검사부적합 내역이 없거나, API 연동에 문제가 발생했습니다.")
else:
    # 사이드바 필터링 UI
    search_company = st.sidebar.text_input("업체명 검색", placeholder="예: 농업회사법인")
    search_product = st.sidebar.text_input("제품명 검색", placeholder="예: 우유, 치즈")
    
    unique_types = ["전체"] + sorted(list(raw_df['식품유형'].dropna().unique()))
    selected_type = st.sidebar.selectbox("식품유형 선택", unique_types)
    
    # 필터 적용 로직
    filtered_df = raw_df.copy()
    if search_company:
        filtered_df = filtered_df[filtered_df['업체명'].str.contains(search_company, na=False, case=False)]
    if search_product:
        filtered_df = filtered_df[filtered_df['제품명'].str.contains(search_product, na=False, case=False)]
    if selected_type != "전체":
        filtered_df = filtered_df[filtered_df['식품유형'] == selected_type]

    # 아산 지역 및 유가공품 데이터 분류 (리스크 매니지먼트용)
    asan_df = raw_df[raw_df['소재지'].astype(str).str.contains('아산', na=False)]
    dairy_df = raw_df[raw_df['식품유형'].astype(str).str.contains('우유|유가공|치즈|발효유', na=False)]

    # 메인 화면 타이틀
    st.title("🛡️ 실시간 식품안전 검사부적합 모니터링 대시보드")
    st.caption(f"기준일시: {today.strftime('%Y-%m-%d %H:%M')} / 데이터 출처: 식품의약품안전처 OpenAPI")

    # 6. 핵심 지표 시각화 (Metrics)
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric(label="🚨 전체 부적합 적발 (최근 1000건 기준)", value=f"{len(raw_df)} 건")
    with col2:
        st.metric(label="🥛 유가공품 관련 부적합 건수", value=f"{len(dairy_df)} 건")
    with col3:
        st.metric(label="📍 관내(아산시) 업체 부적합 건수", value=f"{len(asan_df)} 건")

    st.markdown("---")

    # 7. AI 기반 실시간 품질 리스크 리포트
    st.subheader("🤖 AI 기반 실시간 품질 리스크 리포트")
    
    report_text = "현재 조회된 식약처 최신 데이터를 바탕으로 분석한 QC 리스크 결과입니다.\n\n"
    
    if len(asan_df) > 0:
        report_text += "**[긴급 확인] 아산시 관내 업체 부적합 데이터 발생:**\n"
        report_text += "- 관내(아산) 소재 업체의 부적합 이력이 확인되었습니다. 원부자재 납품 리스트와 해당 업체를 대조하여 입고 여부를 즉시 크로스체크하시기 바랍니다.\n"
    else:
        report_text += "**[안정] 아산시 관내 업체 특이사항 없음:**\n"
        report_text += "- 관내 소재 업체의 신규 부적합 적발 내역이 없습니다.\n"
        
    if len(dairy_df) > 0:
        report_text += "\n**[주의] 유가공품 카테고리 부적합 발생:**\n"
        report_text += "- 동종 업계(유가공품, 우유류 등)의 부적합 사례가 포착되었습니다. 하단 상세 데이터에서 위반 기준규격(대장균군, 세균수 등)을 확인하고 자사 공정 관리 기준을 점검하십시오.\n"

    st.warning(report_text) if len(asan_df) > 0 or len(dairy_df) > 0 else st.info(report_text)

    # 8. Plotly 데이터 시각화 (두 가지 차트 나란히 배치)
    st.subheader("📊 부적합 발생 통계 분석")
    chart_col1, chart_col2 = st.columns(2)
    
    with chart_col1:
        # 식품유형별 부적합 파이 차트 (상위 10개)
        type_counts = filtered_df['식품유형'].value_counts().reset_index()
        type_counts.columns = ['식품유형', '건수']
        type_counts_top10 = type_counts.head(10)
        
        fig_pie = px.pie(type_counts_top10, values='건수', names='식품유형', 
                         title='주요 위반 식품유형 비율 (Top 10)',
                         hole=0.4, color_discrete_sequence=px.colors.qualitative.Pastel)
        st.plotly_chart(fig_pie, use_container_width=True)
        
    with chart_col2:
        # 제조일자별 부적합 발생 추이 라인 차트
        if '제조일자' in filtered_df.columns:
            # 연-월 단위로 그룹화하기 위해 전처리
            filtered_df['제조월'] = filtered_df['제조일자'].str[:6] # YYYYMM 추출
            date_counts = filtered_df[filtered_df['제조월'] != ""].groupby('제조월').size().reset_index(name='건수')
            date_counts = date_counts.sort_values('제조월')
            
            fig_line = px.line(date_counts, x='제조월', y='건수', markers=True,
                               title='제조월별 부적합 발생 추이',
                               line_shape='spline')
            fig_line.update_layout(xaxis_title="제조 연월", yaxis_title="발생 건수")
            st.plotly_chart(fig_line, use_container_width=True)

    # 9. 리콜 정보 강조 및 데이터 테이블 출력
    st.subheader("📋 필터링된 상세 데이터 및 리콜 여부")
    
    # 회수폐기일련번호가 있는 경우 시각적 강조
    display_df = filtered_df.copy()
    if '회수폐기일련번호' in display_df.columns:
        display_df['리콜대상여부'] = display_df['회수폐기일련번호'].apply(
            lambda x: "⚠️ 리콜대상" if pd.notna(x) and str(x).strip() != "" else "해당없음"
        )
        # 컬럼 순서 재배치 (리콜대상여부를 앞쪽으로)
        cols = ['리콜대상여부', '업체명', '제품명', '식품유형', '부적합항목', '제조일자', '유통기한', '소재지']
        existing_cols = [c for c in cols if c in display_df.columns]
        display_df = display_df[existing_cols]

    st.dataframe(display_df, use_container_width=True)

    # 10. CSV 다운로드 버튼
    csv_data = display_df.to_csv(index=False, encoding='utf-8-sig')
    st.download_button(
        label="📥 현재 조회된 데이터 CSV 다운로드",
        data=csv_data,
        file_name=f"식품안전_검사부적합_모니터링_{today.strftime('%Y%m%d')}.csv",
        mime="text/csv",
    )
