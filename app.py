import streamlit as st
import pandas as pd
import requests
import xml.etree.ElementTree as ET
from datetime import datetime, timedelta

# 1. 페이지 기본 설정
st.set_page_config(page_title="실시간 가축전염병 방역 동향 (API 연동)", layout="wide")

# 2. 스트림릿 시크릿에서 API 키 불러오기
try:
    API_KEY = st.secrets["PUBLIC_API_KEY"]
except KeyError:
    st.error("⚠️ Streamlit Secrets에 'PUBLIC_API_KEY'가 설정되지 않았습니다. 셋팅을 확인해주세요.")
    st.stop()

# 3. 공공 API 호출 함수 (데이터 캐싱 적용으로 속도 향상)
@st.cache_data(ttl=3600) # 1시간마다 데이터 갱신
def fetch_disease_data(start_date, end_date):
    url = "http://apis.data.go.kr/1543061/animalDiseaseSrvc/getAnimalDiseaseList"
    
    params = {
        "ServiceKey": API_KEY,
        "bgnde": start_date,
        "endde": end_date,
        "numOfRows": "1000",
        "pageNo": "1"
    }
    
    try:
        response = requests.get(url, params=params)
        response.raise_for_status()
        
        # XML 데이터 파싱
        root = ET.fromstring(response.text)
        
        # 에러 코드 확인 (정상 코드는 '00')
        result_code = root.find('.//resultCode')
        if result_code is not None and result_code.text != '00':
            result_msg = root.find('.//resultMsg').text
            st.error(f"API 호출 오류: {result_msg}")
            return pd.DataFrame()

        items = root.findall('.//item')
        dataList = []
        
        for item in items:
            date_str = item.find('occrDe').text if item.find('occrDe') is not None else ""
            formatted_date = f"{date_str[:4]}-{date_str[4:6]}-{date_str[6:8]}" if len(date_str) == 8 else date_str
            
            row = {
                "발생일자": formatted_date,
                "질병명": item.find('dssNm').text if item.find('dssNm') is not None else "",
                "축종": item.find('lclsDtlNm').text if item.find('lclsDtlNm') is not None else "",
                "시도": item.find('ctprvnNm').text if item.find('ctprvnNm') is not None else "",
                "시군구": item.find('signguNm').text if item.find('signguNm') is not None else "",
                "농장명": item.find('farmNm').text if item.find('farmNm') is not None else "비공개"
            }
            dataList.append(row)
            
        df = pd.DataFrame(dataList)
        if not df.empty:
            # 날짜순 내림차순 정렬
            df = df.sort_values(by="발생일자", ascending=False).reset_index(drop=True)
        return df

    except Exception as e:
        st.error(f"데이터를 불러오는 중 오류가 발생했습니다: {e}")
        return pd.DataFrame()

# 4. 데이터 조회 기간 설정 (최근 3개월)
today = datetime.now()
three_months_ago = today - timedelta(days=90)
start_date_str = three_months_ago.strftime("%Y%m%d")
end_date_str = today.strftime("%Y%m%d")

# 5. 화면 UI 구성
st.title("🛡️ 실시간 가축 전염병 및 방역동향 모니터링 (Live)")
st.caption(f"기준일시: {today.strftime('%Y-%m-%d %H:%M')} / 데이터 출처: 농림축산검역본부 공공데이터 API")

# 데이터 로딩 상태 표시
with st.spinner("농림축산검역본부 실시간 데이터를 불러오는 중입니다..."):
    raw_df = fetch_disease_data(start_date_str, end_date_str)

if raw_df.empty:
    st.warning("조회된 가축전염병 발생 내역이 없거나, API 키 설정이 유효하지 않습니다.")
else:
    # 충남 아산 지역 데이터 필터링
    chungnam_df = raw_df[raw_df["시도"].str.contains("충청남도", na=False)]
    asan_df = raw_df[raw_df["시군구"].str.contains("아산", na=False)]
    
    # 6. 주요 지표 시각화
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric(label="🚨 전국 누적 발생 (최근 3개월)", value=f"{len(raw_df)} 건")
    with col2:
        st.metric(label="⚠️ 충청남도 발생 건수", value=f"{len(chungnam_df)} 건")
    with col3:
        st.metric(label="📍 아산시 발생 건수", value=f"{len(asan_df)} 건")

    st.markdown("---")

    # 7. 자동화된 방역 분석 리포트 (AI 활용을 가장한 텍스트 생성)
    st.subheader("🤖 AI 기반 실시간 방역 리포트 및 지시사항")
    
    report_text = f"현재 조회된 최신 데이터({end_date_str} 기준)를 바탕으로 분석한 결과입니다.\n\n"
    
    if len(asan_df) > 0:
        latest_asan_disease = asan_df.iloc[0]['질병명']
        report_text += f"**[긴급] 관내(아산시) {latest_asan_disease} 발생 확인:**\n"
        report_text += "- 아산 지역 내 발생 건이 확인되었습니다. 해당 노선 집유 차량은 진출입 시 2차 고압 소독을 의무화하고, 착유 농가에 즉각적인 예찰 강화 메시지를 발송해야 합니다.\n"
    elif len(chungnam_df) > 0:
        report_text += "**[주의] 충청남도 인접 지역 발생 확인:**\n"
        report_text += "- 아산시 내부 발생은 없으나, 인접한 충남 권역에 가축 질병이 발생했습니다. 도 경계 및 인접 시군에서 진입하는 차량에 대한 소독 증명서 확인을 강화하십시오.\n"
    else:
        report_text += "**[양호] 충남 권역 방역 상태 안정적:**\n"
        report_text += "- 현재 충청남도 내 최근 보고된 주요 특이 질병 동향은 없습니다. 기존 정기 방역 루틴을 유지하십시오.\n"
        
    recent_diseases = raw_df['질병명'].value_counts().head(2)
    report_text += f"\n**[전국 트렌드]** 최근 전국적으로 가장 빈발하는 질병은 **{recent_diseases.index[0]}** 입니다. 관련된 농가 지도 가이드를 점검하십시오."
    
    st.info(report_text)

    # 8. 원본 데이터프레임 시각화
    st.subheader("📋 실시간 전국 발생 상세 데이터")
    
    # 검색 필터 기능
    search_keyword = st.text_input("🔍 질병명 또는 지역명으로 검색 (예: 럼피스킨, 아산)")
    if search_keyword:
        filtered_df = raw_df[
            raw_df["질병명"].str.contains(search_keyword) | 
            raw_df["시도"].str.contains(search_keyword) | 
            raw_df["시군구"].str.contains(search_keyword)
        ]
        st.dataframe(filtered_df, use_container_width=True)
    else:
        st.dataframe(raw_df, use_container_width=True)