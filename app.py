import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Alignment, Font, Border, Side, PatternFill
import re
from io import BytesIO
import datetime

def generate_excel(text_data, excel_df):
    # 1. 텍스트 데이터 파싱을 통한 기본 정보 추출
    info = {}
    
    report_no_match = re.search(r'품목보고번호\s*(\d+)', text_data)
    info['품목보고번호'] = report_no_match.group(1) if report_no_match else ""
    
    type_match = re.search(r'식품의 유형\s*([^\n]+)', text_data)
    info['식품의 유형'] = type_match.group(1).replace("제품명", "").strip() if type_match else ""
    
    name_match = re.search(r'제품명\s*([^\n]+)', text_data)
    info['제품명'] = name_match.group(1).replace("소비기한", "").strip() if name_match else ""
    
    expire_match = re.search(r'소비기한\s*([^\n]+)', text_data)
    info['소비기한'] = expire_match.group(1).replace("품질유지기한", "").strip() if expire_match else ""
    
    storage_match = re.search(r'보관방법 및 포장재질\s*([^\n]+)', text_data)
    info['보관방법'] = storage_match.group(1).split('/')[0].strip() if storage_match else ""
    
    package_match = re.search(r'포장방법 및 포장단위\s*([^\n]+)', text_data)
    info['포장단위'] = package_match.group(1).replace("성상", "").strip() if package_match else ""
    
    appearance_match = re.search(r'성상\s*([^\n]+)', text_data)
    info['성상'] = appearance_match.group(1).replace("살균·멸균", "").strip() if appearance_match else ""

    usage_match = re.search(r'용도용법\s*([^\n]+)', text_data)
    info['용도용법'] = usage_match.group(1).replace("보관방법", "").strip() if usage_match else ""
    
    sterilization_match = re.search(r'살균·멸균\s*([^\n]+)', text_data)
    info['살균방법'] = sterilization_match.group(1).replace("기타", "").strip() if sterilization_match else ""

    # 2. 엑셀 파일(PRDLST_REPORT_LIST)과 매칭하여 추가 정보 추출
    report_date = ""
    ingredients = ""
    
    if info['품목보고번호'] and not excel_df.empty:
        # 데이터프레임에서 품목보고번호 매칭 (자료형 통일)
        excel_df['품목보고번호'] = excel_df['품목보고번호'].astype(str)
        matched_row = excel_df[excel_df['품목보고번호'] == info['품목보고번호']]
        
        if not matched_row.empty:
            report_date_raw = str(matched_row['신(보)고일자'].values[0])
            if len(report_date_raw) >= 10:
                report_date = f"{report_date_raw[:4]}년 {report_date_raw[5:7]}월 {report_date_raw[8:10]}일"
            
            ingredients = str(matched_row['원재료(배합비)'].values[0])

    # 3. 제품설명서 엑셀 양식 작성
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "제품설명서"

    border_thin = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    align_center = Alignment(horizontal='center', vertical='center', wrap_text=True)
    align_left = Alignment(horizontal='left', vertical='center', wrap_text=True)

    # 상단 헤더 영역 작성
    ws.merge_cells('A1:B3')
    ws['A1'] = "연세대학교 연세유업\n제품설명서(99)\n품질안전부문"
    ws['A1'].alignment = align_center
    ws['A1'].font = Font(bold=True)

    ws.merge_cells('C1:D3')
    ws['C1'] = "HACCP\n관리기준서"
    ws['C1'].alignment = align_center
    ws['C1'].font = Font(size=14, bold=True)

    ws.merge_cells('E1:F1')
    ws['E1'] = "YS-HACCP(멸균유)"
    ws['E1'].alignment = align_center

    ws['E2'] = "제정일자"
    ws['E2'].alignment = align_center
    ws['F2'] = "1998년 3월 30일"
    ws['F2'].alignment = align_center

    ws['E3'] = "개정일자"
    ws['E3'].alignment = align_center
    ws['F3'] = "2026년 2월 5일"
    ws['F3'].alignment = align_center

    ws.merge_cells('A5:F5')
    ws['A5'] = f"99) {info.get('제품명', '')}"
    ws['A5'].font = Font(size=12, bold=True)

    ws.merge_cells('A6:B6')
    ws['A6'] = "구 분"
    ws['A6'].alignment = align_center
    ws['A6'].fill = PatternFill(start_color="EAEAEA", end_color="EAEAEA", fill_type="solid")

    ws.merge_cells('C6:F6')
    ws['C6'] = "내 용"
    ws['C6'].alignment = align_center
    ws['C6'].fill = PatternFill(start_color="EAEAEA", end_color="EAEAEA", fill_type="solid")

    def add_row(row_idx, col1, col2, col3=None, col4=None, merge_cf=True):
        ws.merge_cells(f'A{row_idx}:B{row_idx}')
        ws[f'A{row_idx}'] = col1
        ws[f'A{row_idx}'].alignment = align_center
        if merge_cf:
            ws.merge_cells(f'C{row_idx}:F{row_idx}')
            ws[f'C{row_idx}'] = col2
            ws[f'C{row_idx}'].alignment = align_left
        else:
            ws.merge_cells(f'C{row_idx}:D{row_idx}')
            ws[f'C{row_idx}'] = col2
            ws[f'C{row_idx}'].alignment = align_center
            ws[f'E{row_idx}'] = col3
            ws[f'E{row_idx}'].alignment = align_center
            ws[f'F{row_idx}'] = col4
            ws[f'F{row_idx}'].alignment = align_center

    add_row(7, "1. 제품명", info.get('제품명', ''), "미출시", "", merge_cf=False)
    ws.merge_cells('C7:E7')
    ws['C7'] = info.get('제품명', '')
    ws['C7'].alignment = align_left
    ws['F7'] = "미출시"
    ws['F7'].font = Font(color="FF0000")

    add_row(8, "2. 식품의 유형", info.get('식품의 유형', ''))
    
    add_row(9, "3. 품목제조보고 연월일", report_date, "품목보고번호", info.get('품목보고번호', ''), merge_cf=False)
    
    current_date_str = datetime.datetime.now().strftime("%Y년 %m월 %d일")
    add_row(10, "4. 작성자 및 작성 연월일", "식품안전팀 곽정혁", "작성일", current_date_str, merge_cf=False)
    
    add_row(11, "5. 성분배합비율", ingredients)
    ws.row_dimensions[11].height = 60
    
    add_row(12, "6. 포장단위", info.get('포장단위', ''))

    # 완제품 규격 항목
    ws.merge_cells('A13:A18')
    ws['A13'] = "7. 완제품의 규격"
    ws['A13'].alignment = align_center

    ws['B13'] = "성상"
    ws['B13'].alignment = align_center
    ws.merge_cells('C13:F13')
    ws['C13'] = info.get('성상', '고유의 색과 향미를 가진 균일한 액체')
    ws['C13'].alignment = align_center

    ws.merge_cells('B14:B16')
    ws['B14'] = "생물학적"
    ws['B14'].alignment = align_center

    ws['C14'] = "구 분"
    ws['C14'].alignment = align_center
    ws.merge_cells('D14:E14')
    ws['D14'] = "법적규격"
    ws['D14'].alignment = align_center
    ws['F14'] = "사내규격"
    ws['F14'].alignment = align_center

    ws['C15'] = "세균수(cfu/ml)"
    ws.merge_cells('D15:E15')
    ws['D15'] = "n=5, c=0, m=0"
    ws['F15'] = "좌 동"

    ws['C16'] = "대장균군(cfu/ml)"
    ws.merge_cells('D16:E16')
    ws['D16'] = "음 성"
    ws['F16'] = "좌 동"

    ws['B17'] = "이화학적"
    ws['B17'].alignment = align_center
    ws['C17'] = "pH"
    ws.merge_cells('D17:E17')
    ws['D17'] = "-"
    ws['F17'] = "좌 동"

    ws['B18'] = "물리적"
    ws['B18'].alignment = align_center
    ws['C18'] = "이물"
    ws.merge_cells('D18:E18')
    ws['D18'] = "불검출"
    ws['F18'] = "좌 동"

    add_row(19, "8. 보존기준 및 운송조건", info.get('보존방법', '실온보관'))
    add_row(20, "9. 제품용도", info.get('용도용법', '직접음용'))
    add_row(21, "10. 소비기한", info.get('소비기한', ''))
    
    # 텍스트에 포함된 멸균 방법 파싱
    sterilization_val = info.get('살균방법', '')
    if '멸균' in sterilization_val and '기타' in sterilization_val:
         sterilization_val = "135~150 ℃에서 30~45초간 멸균"
    add_row(22, "11. 살균방법", sterilization_val)
    
    add_row(23, "12. 포장방법 및 재질", info.get('포장단위', ''))
    add_row(24, "13. 알러겐 주의사항", "본 제품은 알레르기 유발물질을 사용한 제품과 같은 제조시설에서 제조하고 있습니다. (필요 시 수정)")

    # 전체 테두리 및 너비 조정
    for row in ws.iter_rows(min_row=1, max_row=24, min_col=1, max_col=6):
        for cell in row:
            cell.border = border_thin
            if cell.alignment.horizontal is None:
                cell.alignment = align_center

    ws.column_dimensions['A'].width = 15
    ws.column_dimensions['B'].width = 15
    ws.column_dimensions['C'].width = 25
    ws.column_dimensions['D'].width = 15
    ws.column_dimensions['E'].width = 15
    ws.column_dimensions['F'].width = 15

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# Streamlit UI 구성
st.title("제품설명서 자동 생성 시스템")

st.subheader("1. 품목제조보고 목록 엑셀 파일 업로드")
uploaded_file = st.file_uploader("엑셀 파일 (PRDLST_REPORT_LIST.xls) 업로드", type=['xls', 'xlsx'])

st.subheader("2. 품목제조 기본정보 텍스트 입력")
text_input = st.text_area("식품안전나라 등에서 복사한 텍스트를 붙여넣으세요.", height=200)

if st.button("제품설명서 생성"):
    if uploaded_file and text_input:
        try:
            # 관공서 다운로드 xls 파일이 내부적으로 HTML 구조를 가지는 경우를 처리
            try:
                df = pd.read_excel(uploaded_file)
            except ValueError:
                df = pd.read_html(uploaded_file)[0]
                
            excel_data = generate_excel(text_input, df)
            
            st.success("제품설명서 생성이 완료되었습니다.")
            st.download_button(
                label="엑셀 파일 다운로드",
                data=excel_data,
                file_name=f"제품설명서_자동생성.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except Exception as e:
            st.error(f"엑셀 파일 생성 중 오류가 발생했습니다: {e}")
    else:
        st.warning("엑셀 파일과 텍스트 정보를 모두 입력해주세요.")
