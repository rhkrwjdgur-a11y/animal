import streamlit as st
import pandas as pd
import requests
from datetime import datetime

# 1. 페이지 기본 설정
st.set_page_config(page_title="실시간 가축전염병 방역 동향 (API 연동)", layout="wide")

# 2. 스트림릿 시크릿에서 API 키 불러오기
try:
    API_KEY = st.secrets["PUBLIC_API_KEY"]
except KeyError:
    st.error("⚠️ Streamlit Secrets에 'PUBLIC_API_KEY'가 설정되지 않았습니다. 셋팅을 확인해주세요.")
    st.stop()

# 3. 공공 API 호출 함수 (농림축산식품 공공데이터포털 Grid API 방식 적용)
@st.cache_data(ttl=3600) # 1시간마다 데이터 갱신
def fetch_disease_data():
    # image_b07ce4.png 명세서에 따른 API_URL 값
    api_url_code = "Grid_20151204000000000316_1"
    
    # image_b07ca6.png 및 image_b07ce4.png 명세서에 따른 요청 URL 구조 조립 (JSON 타입으로 요청)
    url = f"http://211.237.50.150:7080/openapi/{API_KEY}/json/{api_url_code}/1/1000"
    
    try:
        response = requests.get(url)
        response.raise_for_status()
        data = response.json()
        
        # 데이터가 정상적으로 수신되었는지 확인
        if api_url_code in data and 'row' in data[api_url_code]:
            df = pd.DataFrame(data[api_url_code]['row'])
            
            # image_b07ce4.png 명세서의 항목명에 맞추어 컬럼명 변경
            rename_dict = {
                'OCCRRNC_DE': '발생일자',
                'LKNTS_NM': '질병명'
            }
            df = df.rename(columns=lambda x: rename_dict.get(x, x))
            
            if '발생일자' in df.columns:
                df = df.sort_values(by="발생일자", ascending=False).reset_index(drop=True)
                
            return df
        else:
            return pd.DataFrame()

    except Exception as e:
        st.error(f"API 호출 중 오류 발생: {e}")
        return pd.DataFrame()

# 4. 기준 일시 설정
today = datetime.now()

# 5. 화면 UI 구성
st.title("🛡️ 실시간 가축 전염병 및 방역동향 모니터링 (Live)")
st.caption(f"기준일시: {today.strftime('%Y-%m-%d %H:%M')} / 데이터 출처: 농림축산식품 공공데이터포털 API")

# 데이터 로딩 상태 표시
with st.spinner("농림축산검역본부 실시간 데이터를 불러오는 중입니다..."):
    raw_df = fetch_disease_data()

if raw_df.empty:
    st.warning("조회된 가축전염병 발생 내역이 없거나, API 키 설정에 문제가 있습니다.")
else:
    # 6. 아산/충남 지역 데이터 필터링 (반환되는 컬럼명이 가변적일 수 있어 전체 텍스트 검색 활용)
    asan_df = raw_df[raw_df.apply(lambda row: row.astype(str).str.contains('아산').any(), axis=1)]
    chungnam_df = raw_df[raw_df.apply(lambda row: row.astype(str).str.contains('충청남도|충남').any(), axis=1)]
    
    # 7. 주요 지표 시각화
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric(label="🚨 전국 누적 발생 (최근 1000건)", value=f"{len(raw_df)} 건")
    with col2:
        st.metric(label="⚠️ 충청남도 관련 건수", value=f"{len(chungnam_df)} 건")
    with col3:
        st.metric(label="📍 아산시 관련 건수", value=f"{len(asan_df)} 건")

    st.markdown("---")

    # 8. 자동화된 방역 분석 리포트
    st.subheader("🤖 AI 기반 실시간 방역 리포트 및 지시사항")
    
    report_text = "현재 조회된 최신 데이터를 바탕으로 분석한 결과입니다.\n\n"
    
    if len(asan_df) > 0:
        report_text += "**[긴급] 관내(아산시) 질병 발생 데이터 확인:**\n"
        report_text += "- 아산 지역 관련 방역 이력이 확인되었습니다. 해당 노선 집유 차량은 진출입 시 2차 고압 소독을 의무화하고, 착유 농가에 즉각적인 예찰 강화 메시지를 발송해야 합니다.\n"
    elif len(chungnam_df) > 0:
        report_text += "**[주의] 충청남도 인접 지역 발생 확인:**\n"
        report_text += "- 관내 발생은 없으나 충남 권역에 질병 데이터가 있습니다. 도 경계 및 인접 시군에서 진입하는 차량에 대한 소독 증명서 확인을 강화하십시오.\n"
    else:
        report_text += "**[양호] 충남 권역 방역 상태 안정적:**\n"
        report_text += "- 현재 충청남도 내 최근 보고된 주요 특이 질병 동향은 없습니다. 기존 정기 방역 루틴을 유지하십시오.\n"
        
    st.info(report_text)

    # 9. 원본 데이터프레임 시각화 및 검색 기능
    st.subheader("📋 실시간 전국 발생 상세 데이터")
    
    search_keyword = st.text_input("🔍 질병명 또는 지역명으로 검색 (예: 럼피스킨, 아산)")
    if search_keyword:
        filtered_df = raw_df[raw_df.apply(lambda row: row.astype(str).str.contains(search_keyword).any(), axis=1)]
        st.dataframe(filtered_df, use_container_width=True)
    else:
        st.dataframe(raw_df, use_container_width=True)
