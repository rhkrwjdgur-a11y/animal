import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Alignment, Font, Border, Side, PatternFill
from openpyxl.drawing.image import Image as OpenpyxlImage
from PIL import Image as PILImage
import re
from io import BytesIO
import datetime

def generate_excel(text_data, excel_df, doc_number, is_china_export, package_img_file):
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
        excel_df['품목보고번호'] = excel_df['품목보고번호'].astype(str)
        matched_row = excel_df[excel_df['품목보고번호'] == info['품목보고번호']]
        
        if not matched_row.empty:
            report_date_raw = str(matched_row['신(보)고일자'].values[0])
            if len(report_date_raw) >= 10:
                report_date = f"{report_date_raw[:4]}년 {report_date_raw[5:7]}월 {report_date_raw[8:10]}일"
            
            ingredients = str(matched_row['원재료(배합비)'].values[0])

    # 3. 제품설명서 엑셀 양식 작성 (A~L열 12칸 구조 적용)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "제품설명서"

    border_thin = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    align_center = Alignment(horizontal='center', vertical='center', wrap_text=True)
    align_left = Alignment(horizontal='left', vertical='center', wrap_text=True)
    header_fill = PatternFill(start_color="EAEAEA", end_color="EAEAEA", fill_type="solid")

    # 상단 헤더 영역 작성 (1~3행)
    ws.merge_cells('A1:E1')
    ws['A1'] = "연세대학교 연세유업"
    ws['A1'].alignment = align_center
    ws['A1'].font = Font(bold=True)

    ws.merge_cells('A2:E2')
    ws['A2'] = f"제품설명서 ({doc_number})"
    ws['A2'].alignment = align_center
    ws['A2'].font = Font(bold=True)

    ws.merge_cells('A3:E3')
    ws['A3'] = "품질안전부문"
    ws['A3'].alignment = align_center
    ws['A3'].font = Font(bold=True)

    ws.merge_cells('F1:I3')
    ws['F1'] = "HACCP\n관리기준서"
    ws['F1'].alignment = align_center
    ws['F1'].font = Font(size=14, bold=True)

    sterilization_val = info.get('살균방법', '')
    is_pasteurized = '살균' in sterilization_val and '멸균' not in sterilization_val
    
    ws.merge_cells('J1:L1')
    ws['J1'] = "YS-HACCP(우)" if is_pasteurized else "YS-HACCP(멸균유)"
    ws['J1'].alignment = align_center

    ws['J2'] = "제정일자"
    ws['J2'].alignment = align_center
    ws.merge_cells('K2:L2')
    ws['K2'] = "1998년 3월 30일"
    ws['K2'].alignment = align_center

    ws['J3'] = "개정일자"
    ws['J3'].alignment = align_center
    ws.merge_cells('K3:L3')
    ws['K3'] = "2026년 2월 5일"
    ws['K3'].alignment = align_center

    # 타이틀 (5행)
    ws.merge_cells('A5:L5')
    ws['A5'] = f"{doc_number}) {info.get('제품명', '')}"
    ws['A5'].font = Font(size=12, bold=True)

    # 본문 표 헤더 (7행)
    ws.merge_cells('A7:D7')
    ws['A7'] = "구 분"
    ws['A7'].alignment = align_center
    ws['A7'].fill = header_fill

    ws.merge_cells('E7:L7')
    ws['E7'] = "내 용"
    ws['E7'].alignment = align_center
    ws['E7'].fill = header_fill

    def add_row(row_idx, col1, col2, col3=None, col4=None, merge_el=True):
        ws.merge_cells(f'A{row_idx}:D{row_idx}')
        ws[f'A{row_idx}'] = col1
        ws[f'A{row_idx}'].alignment = align_center
        if merge_el:
            ws.merge_cells(f'E{row_idx}:L{row_idx}')
            ws[f'E{row_idx}'] = col2
            ws[f'E{row_idx}'].alignment = align_left
        else:
            ws.merge_cells(f'E{row_idx}:H{row_idx}')
            ws[f'E{row_idx}'] = col2
            ws[f'E{row_idx}'].alignment = align_center
            ws.merge_cells(f'I{row_idx}:J{row_idx}')
            ws[f'I{row_idx}'] = col3
            ws[f'I{row_idx}'].alignment = align_center
            ws.merge_cells(f'K{row_idx}:L{row_idx}')
            ws[f'K{row_idx}'] = col4
            ws[f'K{row_idx}'].alignment = align_center

    # 1. 제품명 ~ 4. 작성자 (8~11행)
    add_row(8, "1. 제품명", info.get('제품명', ''), "미출시", "", merge_el=False)
    ws.merge_cells('E8:J8') 
    ws['E8'] = info.get('제품명', '')
    ws['E8'].alignment = align_left
    ws.merge_cells('K8:L8')
    if is_china_export:
        ws['K8'] = "중국수출용"
        ws['K8'].font = Font(color="000000")
    else:
        ws['K8'] = "미출시"
        ws['K8'].font = Font(color="FF0000")
    ws['K8'].alignment = align_center

    add_row(9, "2. 식품의 유형", info.get('식품의 유형', ''))
    add_row(10, "3. 품목제조보고 연월일", report_date, "품목보고번호", info.get('품목보고번호', ''), merge_el=False)
    
    current_date_str = datetime.datetime.now().strftime("%Y년 %m월 %d일")
    add_row(11, "4. 작성자 및 작성 연월일", "식품안전팀 곽정혁", "작성일", current_date_str, merge_el=False)
    
    # 5. 성분배합비율 (12~13행 병합)
    ws.merge_cells('A12:D13')
    ws['A12'] = "5. 성분배합비율"
    ws['A12'].alignment = align_center
    ws.merge_cells('E12:L13')
    ws['E12'] = ingredients
    ws['E12'].alignment = align_left
    
    # 6. 포장단위 (14행)
    add_row(14, "6. 포장단위", info.get('포장단위', ''))

    # === 7. 완제품의 규격 ===
    food_type = info.get('식품의 유형', '')
    legal_header = "법적규격" 
    
    if is_china_export: 
        legal_header = "중국 GB 표준"
        bio_specs = [
            ("세균수(cfu/ml)", "n=5,c=2,m=10,000,M=50,000", "좌 동"),
            ("대장균군(cfu/ml)", "n=5,c=2,m=0,M=10", "좌 동"),
            ("Salmonella spp.", "n=5, c=0, m=0/25g", "좌 동"),
            ("L.monocytogenes", "n=5, c=0, m=0/25g", "좌 동"),
            ("황색포도상구균", "n=5, c=0, m=0/25g", "좌 동")
        ]
        chem_specs = [
            ("비중(15 ℃)", "-", "1.028 ~ 1.034"),
            ("산도(%)", "-", "0.18 이하"),
            ("유지방(%)", "2.5 이상", "3.0 이상"),
            ("단백질(%)", "2.3 이상", "-"),
            ("납(mg/kg)", "0.05이하", "0.05이하"),
            ("총수은(mg/kg)", "0.01이하", "0.01이하"),
            ("총비소(mg/kg)", "0.1이하", "0.1이하"),
            ("크롬(mg/kg)", "0.3이하", "0.3이하")
        ]
        phys_specs = [("이물", "불검출", "좌 동")]
        
    elif is_pasteurized: 
        if '우유' in food_type and '가공' not in food_type and '강화' not in food_type and '유당' not in food_type:
            bio_specs = [
                ("세균수(cfu/ml)", "n=5,c=2,m=10,000,M=50,000", "좌 동"),
                ("대장균군(cfu/ml)", "n=5,c=2,m=0,M=10", "좌 동"),
                ("Salmonella spp.", "n=5, c=0, m=0/25g", "좌 동"),
                ("L.monocytogenes", "n=5, c=0, m=0/25g", "좌 동"),
                ("황색포도상구균", "n=5, c=0, m=0", "음 성")
            ]
            chem_specs = [
                ("비중(15 ℃)", "1.028 ~ 1.034", "좌 동"),
                ("산도(%)", "0.18 이하", "좌 동"),
                ("총고형분(%)", "-", "11.6 이상"),
                ("유지방(%)", "3.0 이상", "좌 동")
            ]
        elif '가공유' in food_type:
            bio_specs = [
                ("세균수(cfu/ml)", "n=5,c=2,m=10,000,M=50,000", "좌 동"),
                ("대장균군(cfu/ml)", "n=5,c=2,m=0,M=10", "좌 동"),
                ("Salmonella spp.", "n=5, c=0, m=0/25g", "좌 동"),
                ("L.monocytogenes", "n=5, c=0, m=0/25g", "좌 동"),
                ("황색포도상구균", "n=5, c=0, m=0/25g", "좌 동")
            ]
            chem_specs = [
                ("비중(15 ℃)", "-", "1.028 ~ 1.034"),
                ("산도(%)", "-", "0.14 이하"),
                ("무지유고형분(%)", "4.0 이상", "5.5 이상"),
                ("총고형분(%)", "-", "6.3 이상"),
                ("조지방(%)", "0.6 ~ 2.6", "0.8 ~ 1.2")
            ]
        else: 
            bio_specs = [
                ("세균수(cfu/ml)", "n=5,c=2,m=10,000,M=50,000", "좌 동"),
                ("대장균군(cfu/ml)", "n=5,c=2,m=0,M=10", "좌 동"),
                ("Salmonella spp.", "n=5, c=0, m=0/25g", "좌 동"),
                ("L.monocytogenes", "n=5, c=0, m=0/25g", "좌 동"),
                ("황색포도상구균", "n=5, c=0, m=0", "좌 동")
            ]
            chem_specs = [
                ("비중(15 ℃)", "-", "1.028 ~ 1.034"),
                ("산도(%)", "-", "0.18 이하"),
                ("무지유고형분(%)", "8.0 이상", "8.2 이상"),
                ("유지방(%)", "2.5 이상", "3.0 이상")
            ]
        phys_specs = [("이물", "불검출", "좌 동")]
        
    else: 
        if '강화우유' in food_type:
            bio_specs = [
                ("세균수(cfu/ml)", "n=5, c=0, m=0", "좌 동"),
                ("황색포도상구균", "n=5, c=0, m=0/25g", "좌 동"),
                ("Salmonella spp.", "n=5, c=0, m=0/25g", "좌 동"),
                ("L.monocytogenes", "n=5, c=0, m=0/25g", "좌 동")
            ]
            chem_specs = [
                ("산도(%)", "0.18 이하", "좌 동"),
                ("무지유고형분(%)", "8.0 이상", "8.2 이상"),
                ("유지방(%)", "3.0 이상", "좌 동")
            ]
        elif '우유' in food_type and '가공' not in food_type and '유당' not in food_type:
            bio_specs = [
                ("세균수(cfu/ml)", "n=5, c=0, m=0", "좌 동"),
                ("황색포도상구균", "n=5, c=0, m=0/25g", "좌 동"),
                ("Salmonella spp.", "n=5, c=0, m=0/25g", "좌 동"),
                ("L.monocytogenes", "n=5, c=0, m=0/25g", "좌 동")
            ]
            chem_specs = [
                ("산도(%)", "0.18 이하", "좌 동"),
                ("유지방(%)", "3.0 이상", "좌 동")
            ]
        else: 
            bio_specs = [
                ("세균수(cfu/ml)", "n=5, c=0, m=0", "좌 동"),
                ("대장균군(cfu/ml)", "-", "음 성"),
                ("황색포도상구균", "n=5, c=0, m=0/25g", "좌 동"),
                ("Salmonella spp.", "n=5, c=0, m=0/25g", "좌 동"),
                ("L.monocytogenes", "n=5, c=0, m=0/25g", "좌 동")
            ]
            chem_specs = [
                ("무지유고형분(%)", "4.0 이상", "좌 동"),
                ("조지방(%)", "2.7 이상", "좌 동")
            ]
        phys_specs = [("이물", "불검출", "좌 동")]

    curr_row = 15
    num_spec_rows = 1 + 1 + len(bio_specs) + len(chem_specs) + len(phys_specs)
    start_spec_row = curr_row

    ws.merge_cells(f'A{start_spec_row}:B{start_spec_row + num_spec_rows - 1}')
    ws[f'A{start_spec_row}'] = "7. 완제품의 규격"
    
    ws.merge_cells(f'C{curr_row}:D{curr_row}')
    ws[f'C{curr_row}'] = "성상"
    ws.merge_cells(f'E{curr_row}:L{curr_row}')
    ws[f'E{curr_row}'] = info.get('성상', '유백색의 균일한 액체로서 이미, 이취가 없음' if is_pasteurized else '고유의 색과 향미를 가진 균일한 액체')
    curr_row += 1

    bio_start = curr_row
    ws.merge_cells(f'C{bio_start}:C{bio_start + len(bio_specs)}')
    ws[f'C{bio_start}'] = "생물학적"
    
    ws.merge_cells(f'D{curr_row}:F{curr_row}')
    ws[f'D{curr_row}'] = "구 분"
    ws[f'D{curr_row}'].fill = header_fill

    ws.merge_cells(f'G{curr_row}:J{curr_row}')
    ws[f'G{curr_row}'] = legal_header
    ws[f'G{curr_row}'].fill = header_fill

    ws.merge_cells(f'K{curr_row}:L{curr_row}')
    ws[f'K{curr_row}'] = "사내규격"
    ws[f'K{curr_row}'].fill = header_fill
    curr_row += 1

    for item in bio_specs:
        ws.merge_cells(f'D{curr_row}:F{curr_row}')
        ws[f'D{curr_row}'] = item[0]
        ws.merge_cells(f'G{curr_row}:J{curr_row}')
        ws[f'G{curr_row}'] = item[1]
        ws.merge_cells(f'K{curr_row}:L{curr_row}')
        ws[f'K{curr_row}'] = item[2]
        curr_row += 1

    chem_start = curr_row
    ws.merge_cells(f'C{chem_start}:C{chem_start + len(chem_specs) - 1}')
    ws[f'C{chem_start}'] = "이화학적"
    for item in chem_specs:
        ws.merge_cells(f'D{curr_row}:F{curr_row}')
        ws[f'D{curr_row}'] = item[0]
        ws.merge_cells(f'G{curr_row}:J{curr_row}')
        ws[f'G{curr_row}'] = item[1]
        ws.merge_cells(f'K{curr_row}:L{curr_row}')
        ws[f'K{curr_row}'] = item[2]
        curr_row += 1

    phys_start = curr_row
    ws.merge_cells(f'C{phys_start}:C{phys_start + len(phys_specs) - 1}')
    ws[f'C{phys_start}'] = "물리적"
    for item in phys_specs:
        ws.merge_cells(f'D{curr_row}:F{curr_row}')
        ws[f'D{curr_row}'] = item[0]
        ws.merge_cells(f'G{curr_row}:J{curr_row}')
        ws[f'G{curr_row}'] = item[1]
        ws.merge_cells(f'K{curr_row}:L{curr_row}')
        ws[f'K{curr_row}'] = item[2]
        curr_row += 1

    # === 8. 보존기준 ~ 13. 알러겐 ===
    add_row(curr_row, "8. 보존기준 및 운송조건", info.get('보존방법', '2 ~ 6 ℃에서 냉장보관' if is_pasteurized else '실온보관'))
    curr_row += 1
    add_row(curr_row, "9. 제품용도", info.get('용도용법', '직접음용'))
    curr_row += 1
    add_row(curr_row, "10. 소비기한" if not is_pasteurized else "10. 유통기한", info.get('소비기한', ''))
    curr_row += 1
    
    if is_pasteurized:
        if '우유' in food_type or '가공유' in food_type or '강화우유' in food_type:
            sterilization_val = "130 ~ 135 ℃에서 2초간 살균"
        else:
            sterilization_val = "130 ~ 135 ℃에서 2초간 살균" 
    elif '멸균' in sterilization_val:
        if '가공유' in food_type:
            sterilization_val = "135~150 ℃에서 30~45초간 멸균"
        elif '강화우유' in food_type:
            sterilization_val = "135~150 ℃에서 10~14초간 멸균"
        elif '유당분해우유' in food_type:
            sterilization_val = "135~150 ℃에서 30~45초간 멸균"
        elif '우유' in food_type:
            sterilization_val = "135~150 ℃에서 10~14초간 멸균"
        else:
            sterilization_val = "135~150 ℃에서 30~45초간 멸균"
            
    add_row(curr_row, "11. 살균방법", sterilization_val)
    curr_row += 1
    
    add_row(curr_row, "12. 포장방법 및 재질", info.get('포장단위', ''))
    curr_row += 1
    add_row(curr_row, "13. 알러겐 주의사항", "본 제품은 알레르기 유발물질을 사용한 제품과 같은 제조시설에서 제조하고 있습니다. (필요 시 수정)")
    curr_row += 1

    # === [수정됨] 14. 표시사항 (단일 이미지 통채로 병합) ===
    ws.merge_cells(f'A{curr_row}:D{curr_row+14}')
    ws[f'A{curr_row}'] = "14. 표시사항"

    ws.merge_cells(f'E{curr_row}:L{curr_row+14}') # E부터 L까지 통으로 병합

    if package_img_file:
        pil_img = PILImage.open(package_img_file)
        ratio = pil_img.width / float(pil_img.height)
        new_width = 400 # E~L 전체 너비 근사치
        new_height = int(new_width / ratio)
        if new_height > 280: # 병합된 칸의 총 높이를 초과하지 않도록 280px로 제한
            new_height = 280
            new_width = int(new_height * ratio)
            
        img_byte_arr = BytesIO()
        pil_img.save(img_byte_arr, format=pil_img.format if pil_img.format else 'PNG')
        opxl_img = OpenpyxlImage(img_byte_arr)
        opxl_img.width = new_width
        opxl_img.height = new_height
        ws.add_image(opxl_img, f'E{curr_row}')

    last_row = curr_row + 14

    # 전체 테두리 지정 및 정렬
    for row in ws.iter_rows(min_row=1, max_row=3, min_col=1, max_col=12):
        for cell in row:
            cell.border = border_thin
            if cell.alignment.horizontal is None:
                cell.alignment = align_center
            
    for row in ws.iter_rows(min_row=7, max_row=last_row, min_col=1, max_col=12):
        for cell in row:
            cell.border = border_thin
            if cell.alignment.horizontal is None:
                cell.alignment = align_center

    col_widths = {
        'A': 3.13, 'B': 9, 'C': 9, 'D': 4.5, 
        'E': 3, 'F': 5, 'G': 5, 'H': 7, 
        'I': 8, 'J': 9, 'K': 8.25, 'L': 8.25
    }
    for col_letter, width_val in col_widths.items():
        ws.column_dimensions[col_letter].width = width_val

    ws.row_dimensions[1].height = 25
    ws.row_dimensions[2].height = 25
    ws.row_dimensions[3].height = 25
    ws.row_dimensions[4].height = 8
    ws.row_dimensions[5].height = 16
    ws.row_dimensions[6].height = 6
    ws.row_dimensions[7].height = 17.25
    for i in range(8, 12):
        ws.row_dimensions[i].height = 16
    ws.row_dimensions[12].height = 60
    ws.row_dimensions[13].height = 60
    ws.row_dimensions[14].height = 16
    
    dynamic_row = 15
    while dynamic_row < (last_row - 14):
        ws.row_dimensions[dynamic_row].height = 14.25
        dynamic_row += 1
        
    tail_heights = [16, 16, 16, 16, 30, 20, 15, 15, 15, 15, 15, 15, 15, 15, 15]
    for h in tail_heights:
        ws.row_dimensions[dynamic_row].height = h
        dynamic_row += 1

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# Streamlit UI 구성
st.title("제품설명서 자동 생성 시스템")

st.subheader("1. 기본 옵션 설정")
col1, col2 = st.columns(2)
with col1:
    doc_number = st.text_input("문서 번호 입력 (예: 99, 44 등)", value="99")
with col2:
    is_china_export = st.checkbox("중국 수출용 (GB 표준) 규격 적용")

# [수정됨] 단일 이미지 업로드로 변경
st.subheader("2. 표시사항 이미지 업로드 (선택)")
package_img = st.file_uploader("표시사항 이미지 파일", type=['png', 'jpg', 'jpeg'])

st.subheader("3. 품목제조보고 목록 엑셀 파일 업로드")
uploaded_file = st.file_uploader("엑셀 파일 (PRDLST_REPORT_LIST.xls) 업로드", type=['xls', 'xlsx'])

st.subheader("4. 품목제조 기본정보 텍스트 입력")
text_input = st.text_area("식품안전나라 등에서 복사한 텍스트를 붙여넣으세요.", height=200)

if st.button("제품설명서 생성"):
    if uploaded_file and text_input and doc_number:
        try:
            try:
                df = pd.read_excel(uploaded_file)
            except ValueError:
                df = pd.read_html(uploaded_file)[0]
                
            excel_data = generate_excel(text_input, df, doc_number, is_china_export, package_img)
            
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
        st.warning("엑셀 파일, 텍스트 정보, 그리고 문서 번호를 모두 확인해주세요.")
