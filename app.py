import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Alignment, Font, Border, Side, PatternFill
from openpyxl.drawing.image import Image as OpenpyxlImage
from PIL import Image as PILImage
import re
from io import BytesIO
import datetime
import os
import glob
import difflib
import google.generativeai as genai

# ==========================================
# 환경 설정 (고정 변수)
# ==========================================
# 1. 같은 저장소 내의 기준 데이터 폴더 이름 설정
BASE_FOLDER_NAME = "새 폴더 (4)"

# 2. Gemini API 키 설정 (Streamlit Secrets 활용)
try:
    genai.configure(api_key=st.secrets["FOOD_API_KEY"])
except Exception as e:
    st.error("Gemini API 키가 Streamlit Secrets에 올바르게 설정되지 않았습니다.")

# ==========================================
# 1. 제품설명서 자동 생성 로직
# ==========================================
def generate_excel(text_data, excel_df, doc_number, is_china_export, package_img_file):
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

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "제품설명서"

    border_thin = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    align_center = Alignment(horizontal='center', vertical='center', wrap_text=True)
    align_left = Alignment(horizontal='left', vertical='center', wrap_text=True)
    header_fill = PatternFill(start_color="EAEAEA", end_color="EAEAEA", fill_type="solid")

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
    food_type = info.get('식품의 유형', '')
    is_pasteurized = '살균' in sterilization_val and '멸균' not in sterilization_val
    
    ws.merge_cells('J1:L1')
    if any(x in food_type for x in ['과채주스', '혼합음료', '커피', '유산균음료', '과채음료', '액상차']):
        ws['J1'] = "YS-HACCP(음)"
    elif '환자용' in food_type or '영양조제식품' in food_type:
        ws['J1'] = "YS-HACCP(환)"
    elif '농후발효유' in food_type or '발효유' in food_type:
        ws['J1'] = "YS-HACCP(발)"
    elif '가공두유' in food_type:
        ws['J1'] = "YS-HACCP(두)"
    elif is_pasteurized:
        ws['J1'] = "YS-HACCP(우)"
    else:
        ws['J1'] = "YS-HACCP(멸균유)"
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

    ws.merge_cells('A5:L5')
    ws['A5'] = f"{doc_number}) {info.get('제품명', '')}"
    ws['A5'].font = Font(size=12, bold=True)

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
    
    ws.merge_cells('A12:D13')
    ws['A12'] = "5. 성분배합비율"
    ws['A12'].alignment = align_center
    ws.merge_cells('E12:L13')
    ws['E12'] = ingredients
    ws['E12'].alignment = align_left
    
    add_row(14, "6. 포장단위", info.get('포장단위', ''))

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

    elif '유산균음료' in food_type:
        bio_specs = [
            ("일반세균(1 ㎖)", "n=5, c=0, m=0", "음 성"),
            ("대장균군(1 ㎖)", "-", "음 성")
        ]
        chem_specs = [
            ("보존료", "불검출", "좌 동"),
            ("유산균수", "1ml 당 1,000,000 이상", "좌 동")
        ]
        phys_specs = [("이물", "불검출", "좌 동")]

    elif '액상차' in food_type:
        bio_specs = [
            ("세균수(1㎖)", "n=5, c=1, m=100, M=1,000", "음 성"),
            ("대장균군(1㎖)", "n=5, c=1, m=0, M=10", "음 성")
        ]
        chem_specs = [
            ("납", "0.05이하", "좌 동"),
            ("카드뮴", "0.1이하", "좌 동"),
            ("타르색소", "불검출", "좌 동")
        ]
        phys_specs = [("이물", "-", "불검출")]

    elif '과채주스' in food_type or '혼합음료' in food_type or '과채음료' in food_type:
        bio_specs = [
            ("세균수(1㎖)", "n=5, c=1, m=100, M=1,000", "음 성"),
            ("대장균군(1㎖)", "n=5, c=1, m=0, M=10", "음 성")
        ]
        chem_specs = [
            ("납", "0.05이하", "좌 동"),
            ("카드뮴", "0.1이하", "좌 동"),
            ("보존료", "불검출", "좌 동")
        ]
        phys_specs = [("이물", "-", "불검출")]

    elif '커피' in food_type:
        bio_specs = [
            ("세균수(1㎖)", "n=5, c=0, m=0", "음 성"),
            ("대장균군(1㎖)", "n=5, c=1, m=0, M=10", "음 성")
        ]
        chem_specs = [
            ("납", "2.0이하", "좌 동"),
            ("허용외 타르색소", "불검출", "좌 동")
        ]
        phys_specs = [("이물", "-", "불검출")]

    elif '환자용' in food_type or '영양조제식품' in food_type:
        bio_specs = [
            ("일반세균(1 ㎖)", "n=5, c=1, m=10, M=100", "음 성"),
            ("대장균군(1 ㎖)", "n=5, c=0, m=0", "음 성"),
            ("바실러스세레우스(1 ㎖)", "n=5, c=0, m=100", "음 성")
        ]
        chem_specs = [
            ("타르색소", "불검출", "좌 동")
        ]
        phys_specs = [("이물", "불검출", "좌 동")]

    elif '농후발효유' in food_type:
        bio_specs = [
            ("효모,곰팡이(4 ㎖)", "-", "음 성"),
            ("대장균군(1 ㎖)", "n=5, c=2, m=0, M=10", "음 성"),
            ("유산균수(1 ㎖)", "1억 이상", "1억 이상"),
            ("Salmonella spp.", "n=5, c=0, m=0/25g", "좌 동"),
            ("황색포도상구균", "n=5, c=0, m=0/25g", "좌 동"),
            ("L.monocytogenes", "n=5, c=0, m=0/25g", "좌 동")
        ]
        chem_specs = [
            ("무지유고형분(%)", "8.0 이상", "좌 동")
        ]
        phys_specs = [("이물", "불검출", "좌 동")]

    elif '발효유' in food_type:
        bio_specs = [
            ("효모,곰팡이(4 ㎖)", "-", "음 성"),
            ("대장균군(1 ㎖)", "n=5, c=2, m=0 M=10", "음 성"),
            ("유산균수(1 ㎖)", "1,000만 이상", "좌 동"),
            ("Salmonella spp.", "n=5, c=0, m=0/25g", "좌 동"),
            ("황색포도상구균", "n=5, c=0, m=0/25g", "좌 동"),
            ("L.monocytogenes", "n=5, c=0, m=0/25g", "좌 동")
        ]
        chem_specs = [
            ("무지유고형분(%)", "3.0이상", "좌 동")
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
        if '가공두유' in food_type:
            bio_specs = [
                ("세균수", "n=5, c=0, m=0", "음 성"),
                ("대장균군", "-", "음 성"),
                ("L.monocytogenes", "-", "n=5, c=0, m=0/25g"),
                ("장출혈성 대장균", "-", "n=5, c=0, m=0/25g")
            ]
            chem_specs = [
                ("대두고형분(%)", "1.4% 이상", "좌 동")
            ]
        elif '강화우유' in food_type:
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

    row_item_8 = curr_row
    
    add_row(curr_row, "8. 보존기준 및 운송조건", info.get('보존방법', '2 ~ 6 ℃에서 냉장보관' if is_pasteurized else '실온보관'))
    curr_row += 1
    add_row(curr_row, "9. 제품용도", info.get('용도용법', '직접음용'))
    curr_row += 1
    add_row(curr_row, "10. 소비기한" if not is_pasteurized else "10. 유통기한", info.get('소비기한', ''))
    curr_row += 1
    
    drink_types = ['과채주스', '혼합음료', '커피', '유산균음료', '과채음료', '액상차']
    
    if ('농후발효유' in food_type or '발효유' in food_type) and not sterilization_val.strip():
        sterilization_val = "비살균"
    elif is_pasteurized:
        if any(x in food_type for x in drink_types):
            sterilization_val = "90 ~ 130℃에서 30 ~ 37초간 살균"
        elif '우유' in food_type or '가공유' in food_type or '강화우유' in food_type:
            sterilization_val = "130 ~ 135 ℃에서 2초간 살균"
        else:
            sterilization_val = "130 ~ 135 ℃에서 2초간 살균" 
    elif '멸균' in sterilization_val:
        if any(x in food_type for x in drink_types):
            sterilization_val = "135~150℃에서 30~37초 동안 멸균"
        elif '환자용' in food_type or '영양조제식품' in food_type:
            sterilization_val = "143 ~ 153℃에서 3.6 ~ 4.4초간 멸균"
        elif '가공두유' in food_type:
            sterilization_val = "135~150 ℃에서 30~45초간 멸균"
        elif '가공유' in food_type:
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
    
    if '가공두유' in food_type:
        allergen_text = "본 제품은 우유, 알류, 땅콩, 밀, 복숭아, 토마토, 호두, 메밀, 아황산류, 잣을 사용한 제품과 같은 시설에서 제조하고 있습니다. / 대두 함유"
    elif '환자용' in food_type or '영양조제식품' in food_type:
        allergen_text = "본 제품은 알류, 땅콩, 밀, 복숭아, 토마토, 호두, 메밀, 아황산류, 잣을 사용한 제품과 같은 시설에서 제조하고 있습니다. / 대두, 우유 함유"
    else:
        allergen_text = "본 제품은 대두, 땅콩, 알류, 밀, 복숭아, 토마토, 호두, 메밀, 아황산류, 잣을 사용한 제품과 같은 시설에서 제조하고 있습니다. / 우유 함유"
        
    add_row(curr_row, "13. 알러겐 주의사항", allergen_text)
    curr_row += 1

    ws.merge_cells(f'A{curr_row}:D{curr_row+14}')
    ws[f'A{curr_row}'] = "14. 표시사항"

    ws.merge_cells(f'E{curr_row}:L{curr_row+14}') 

    if package_img_file:
        pil_img = PILImage.open(package_img_file)
        ratio = pil_img.width / float(pil_img.height)
        new_width = 400 
        new_height = int(new_width / ratio)
        if new_height > 280: 
            new_height = 280
            new_width = int(new_height * ratio)
            
        img_byte_arr = BytesIO()
        pil_img.save(img_byte_arr, format=pil_img.format if pil_img.format else 'PNG')
        opxl_img = OpenpyxlImage(img_byte_arr)
        opxl_img.width = new_width
        opxl_img.height = new_height
        ws.add_image(opxl_img, f'E{curr_row}')

    last_row = curr_row + 14

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
    
    for r in range(15, row_item_8):
        ws.row_dimensions[r].height = 14.25
        
    ws.row_dimensions[row_item_8].height = 14.25     
    ws.row_dimensions[row_item_8 + 1].height = 14.25 
    ws.row_dimensions[row_item_8 + 2].height = 14.25 
    ws.row_dimensions[row_item_8 + 3].height = 14.25 
    ws.row_dimensions[row_item_8 + 4].height = 14.25 
    ws.row_dimensions[row_item_8 + 5].height = 50    
    
    row_item_14_start = row_item_8 + 6
    tail_heights = [16, 16, 16, 16, 30, 20, 15, 15, 15, 15, 15, 15, 15, 15, 15]
    for i, h in enumerate(tail_heights):
        ws.row_dimensions[row_item_14_start + i].height = h

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output


# ==========================================
# 2 & 3. 갱신 여부 점검 공통 함수 (AI 적용)
# ==========================================
def get_local_base_files(folder_name):
    if os.path.exists(folder_name) and os.path.isdir(folder_name):
        files = glob.glob(os.path.join(folder_name, "*.xls*"))
        return files
    else:
        return []

def load_and_concat(files_or_paths):
    df_list = []
    for item in files_or_paths:
        if not item:
            continue
            
        try:
            if isinstance(item, str): 
                try:
                    df = pd.read_excel(item)
                except ValueError:
                    df = pd.read_html(item)[0]
            else:
                try:
                    df = pd.read_excel(item)
                except ValueError:
                    item.seek(0)
                    df = pd.read_html(item)[0]
                    
            if '원재료(배합비)' in df.columns:
                df.rename(columns={'원재료(배합비)': '원재료'}, inplace=True)
            elif '원재료명' in df.columns:
                df.rename(columns={'원재료명': '원재료'}, inplace=True)
                
            if '품목보고번호' in df.columns and '제품명' in df.columns and '원재료' in df.columns:
                df['품목보고번호'] = df['품목보고번호'].astype(str)
                df_list.append(df[['품목보고번호', '제품명', '원재료']])
        except Exception as e:
            st.error(f"데이터 로드 중 오류 발생 ({item}): {e}")
            
    if df_list:
        return pd.concat(df_list).drop_duplicates(subset=['품목보고번호'], keep='last')
    else:
        return pd.DataFrame()

# [수정됨] Few-shot Prompting이 적용된 스마트 검수 로직
def is_ingredient_changed_ai(old_ing, new_ing):
    old_str = str(old_ing).strip()
    new_str = str(new_ing).strip()
    
    if old_str == new_str:
        return False
        
    try:
        model = genai.GenerativeModel('gemini-pro')
        prompt = f'''
        당신은 식품 배합비 검수 전문가입니다. 아래 두 원재료명 텍스트를 비교하여 실질적인 배합비(%)나 핵심 원재료가 변경되었는지 판단하세요.
        
        [중요 검수 규칙]
        1. 단순한 띄어쓰기, 쉼표, 마침표 등 기호의 차이는 무시하세요.
        2. 동일한 성분들의 단순 기재 순서 변경은 무시하세요.
        3. 품목제조보고서와 제품설명서 간의 동의어, 한글 라벨링 명칭 차이, 축약어 차이는 '변경되지 않음'으로 간주하세요. 
           (예: '액상과당'과 '기타과당', '옥수수기름'과 '옥배유', '비타민C'와 'L-아스코르브산')
        4. [핵심] 원재료명 뒤에 붙는 괄호 안의 세부 명칭이나 수식어가 생략되거나 추가된 경우 해당 원료의 배합 비율(%) 수치가 같다면 절대 '변경됨'으로 판단하지 마세요.
        5. 오직 포함된 원재료의 성분 자체가 완전히 삭제/추가되었거나, 배합 비율(%) 수치가 달라진 경우에만 '변경됨'으로 대답하세요.
        
        [판단 예시]
        예시 1:
        - 기존: 원액두유(70%),두류가공품(분리대두단백분말)(3.3%),향료(천연땅콩향 H24346)(0.05%)
        - 신규: 원액두유(70%),두류가공품(3.3%),향료(0.05%)
        - 정답: 변경 안됨 (이유: 괄호 안 상세명칭이 생략되었으나 비율이 동일함)
        
        예시 2:
        - 기존: 원액두유(70%),탄산칼슘(0.3%),두류가공품(약콩분말)(0.5%),곡류가공품(블랙5곡분말)(0.15%),정제수(28.55%)
        - 신규: 원액두유(70%),탄산칼슘(0.3%),두류가공품(0.5%),곡류가공품(0.15%),정제수(28.55%)
        - 정답: 변경 안됨
        
        예시 3:
        - 기존: 원액두유(70%), 혼합제제(비타민A, 비타민D)(0.1%)
        - 신규: 원액두유(70%), 혼합제제(비타민A, 비타민E)(0.1%)
        - 정답: 변경됨 (이유: 괄호 안의 실제 성분 종류가 비타민D에서 E로 변경됨)
        
        ---
        기존(또는 문서상) 원재료: {old_str}
        신규(최신 DB) 원재료: {new_str}
        
        결과는 반드시 "변경됨" 또는 "변경 안됨" 둘 중 하나로만 대답하세요.
        '''
        response = model.generate_content(prompt)
        result_text = response.text.strip()
        
        if "변경됨" in result_text and "안됨" not in result_text:
            return True
        else:
            return False
    except Exception as e:
        return old_str != new_str

def get_colored_diff(old_text, new_text):
    matcher = difflib.SequenceMatcher(None, old_text, new_text)
    old_html = ""
    new_html = ""
    
    for opcode, a0, a1, b0, b1 in matcher.get_opcodes():
        if opcode == 'equal':
            old_html += old_text[a0:a1]
            new_html += new_text[b0:b1]
        elif opcode == 'insert':
            new_html += f"<span style='color: #0056b3; font-weight: bold; background-color: #e6f2ff;'>{new_text[b0:b1]}</span>"
        elif opcode == 'delete':
            old_html += f"<span style='color: #dc3545; font-weight: bold; text-decoration: line-through; background-color: #ffe6e6;'>{old_text[a0:a1]}</span>"
        elif opcode == 'replace':
            old_html += f"<span style='color: #dc3545; font-weight: bold; text-decoration: line-through; background-color: #ffe6e6;'>{old_text[a0:a1]}</span>"
            new_html += f"<span style='color: #0056b3; font-weight: bold; background-color: #e6f2ff;'>{new_text[b0:b1]}</span>"
            
    return old_html, new_html

def compare_ingredients(old_data, new_data):
    df_old = load_and_concat(old_data)
    df_new = load_and_concat(new_data)
    
    if df_old.empty or df_new.empty:
        return pd.DataFrame()
        
    merged = pd.merge(df_old, df_new, on='품목보고번호', how='right', suffixes=('_기존', '_신규'))
    
    results = []
    progress_bar = st.progress(0)
    status_text = st.empty()
    total_items = len(merged)
    
    for i, row in merged.iterrows():
        status_text.text(f"AI 분석 중... ({i+1}/{total_items}) - {row['제품명_신규']}")
        old_str = str(row['원재료_기존']) if not pd.isna(row['원재료_기존']) else ""
        new_str = str(row['원재료_신규']) if not pd.isna(row['원재료_신규']) else ""
        
        if pd.isna(row['원재료_기존']):
            status = "신규 등록"
            msg_html = ""
        else:
            is_changed = is_ingredient_changed_ai(old_str, new_str)
            if is_changed:
                status = "갱신 필요 (변경됨)"
                old_colored, new_colored = get_colored_diff(old_str, new_str)
                msg_html = f"현재 품목제조보고번호 <b>{row['품목보고번호']}</b> 제품명 <b>{row['제품명_신규']}</b> 인 것 원재료명이<br><br>[기준 DB] {old_colored}<br>에서<br>[최신 DB] {new_colored}<br><br>으로 변경 확인되어 제품설명서 최신화 필요합니다."
            else:
                status = "변경 없음"
                msg_html = ""
            
        results.append({
            '품목보고번호': row['품목보고번호'],
            '제품명': row['제품명_신규'],
            '상태': status,
            '알림 메시지': msg_html,
            '원재료_기존': old_str,
            '원재료_신규': new_str
        })
        progress_bar.progress((i + 1) / total_items)
        
    status_text.empty()
    progress_bar.empty()
    return pd.DataFrame(results)

# ==========================================
# 3. 3번 탭 전용: 다중 시트 제품설명서 파일 스캔 로직
# ==========================================
def parse_existing_specs_from_workbook(file_obj):
    results = []
    try:
        wb = openpyxl.load_workbook(file_obj, data_only=True)
        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            found_report_no = ""
            found_product_name = ""
            found_ingredients = ""
            
            for row in ws.iter_rows(min_row=1, max_row=20):
                for cell in row:
                    val = str(cell.value).strip().replace(" ", "") if cell.value else ""
                    if "품목보고번호" in val and not found_report_no:
                        found_report_no = str(ws.cell(row=cell.row, column=11).value).strip()
                        found_report_no = re.sub(r'[^0-9]', '', found_report_no)
                    if "제품명" in val and len(val) < 5 and not found_product_name:
                        found_product_name = str(ws.cell(row=cell.row, column=5).value).strip()
                    if "성분배합비율" in val and not found_ingredients:
                        found_ingredients = str(ws.cell(row=cell.row, column=5).value).strip()
                        
            if found_report_no:
                results.append({
                    '파일명': f"{file_obj.name} [{sheet_name} 시트]",
                    '품목보고번호': found_report_no,
                    '제품명_문서': found_product_name,
                    '원재료_문서': found_ingredients
                })
        return results
    except Exception as e:
        st.error(f"{file_obj.name} 파일 분석 중 오류가 발생했습니다: {e}")
        return []

def check_existing_specs(spec_files, new_data):
    df_new = load_and_concat(new_data)
    if df_new.empty:
        return pd.DataFrame()

    parsed_data = []
    for f in spec_files:
        parsed_list = parse_existing_specs_from_workbook(f)
        parsed_data.extend(parsed_list)
            
    if not parsed_data:
        st.warning("업로드된 제품설명서에서 유효한 데이터를 찾지 못했습니다.")
        return pd.DataFrame()
        
    df_specs = pd.DataFrame(parsed_data)
    merged = pd.merge(df_specs, df_new, on='품목보고번호', how='inner')
    
    results = []
    progress_bar = st.progress(0)
    status_text = st.empty()
    total_items = len(merged)
    
    for i, row in merged.iterrows():
        status_text.text(f"기존 문서 점검 중... ({i+1}/{total_items}) - {row['파일명']}")
        
        old_str = str(row['원재료_문서']).strip()
        new_str = str(row['원재료']).strip()
        
        is_changed = is_ingredient_changed_ai(old_str, new_str)
        if is_changed:
            status = "최신화 필요"
            old_colored, new_colored = get_colored_diff(old_str, new_str)
            msg_html = f"기존 문서 <b>{row['파일명']}</b> ({row['제품명_문서']})의 원재료명이<br><br>[기존 문서] {old_colored}<br>에서<br>[최신 기준] {new_colored}<br><br>으로 변경 확인되어 제품설명서 최신화 필요합니다."
        else:
            status = "적합"
            msg_html = f"기존 문서 <b>{row['파일명']}</b> ({row['제품명_문서']})은(는) 최신 기준에 맞게 잘 표기되어 있습니다. <b>(적합)</b>"
            
        results.append({
            '파일명': row['파일명'],
            '품목보고번호': row['품목보고번호'],
            '제품명': row['제품명_문서'],
            '상태': status,
            '알림 메시지': msg_html
        })
        progress_bar.progress((i + 1) / total_items)
        
    status_text.empty()
    progress_bar.empty()
    return pd.DataFrame(results)


# ==========================================
# Streamlit UI 구성
# ==========================================
st.title("품질안전부문 업무 자동화 시스템")

tab1, tab2, tab3 = st.tabs(["📄 제품설명서 자동 생성", "🔄 기준 데이터 배합비 변경 확인", "🔎 기존 제품설명서 최신화 점검"])

with tab1:
    st.subheader("1. 기본 옵션 설정")
    col1, col2 = st.columns(2)
    with col1:
        doc_number = st.text_input("문서 번호 입력 (예: 99, 44 등)", value="99")
    with col2:
        is_china_export = st.checkbox("중국 수출용 (GB 표준) 규격 적용")

    st.subheader("2. 표시사항 이미지 업로드 (선택)")
    package_img = st.file_uploader("표시사항 이미지 파일", type=['png', 'jpg', 'jpeg'])

    st.subheader("3. 품목제조보고 목록 엑셀 파일 업로드")
    uploaded_file = st.file_uploader("엑셀 파일 (PRDLST_REPORT_LIST.xls) 업로드", type=['xls', 'xlsx'], key="gen_upload")

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

with tab2:
    st.subheader("원재료명(배합비) 변경 확인 및 갱신 여부 조회")
    
    current_date = datetime.datetime.now().strftime("%Y년 %m월 %d일")
    st.markdown(f"**오늘 날짜({current_date}) 기준**으로 새로 확보한 데이터를 업로드하면, Gemini AI가 저장소의 `{BASE_FOLDER_NAME}` 안의 기준 파일들과 의미적으로 분석 및 대조합니다.")
    
    new_files_tab2 = st.file_uploader("최신 기준 데이터 업로드 (일반식품, 축산물 등)", type=['xls', 'xlsx'], accept_multiple_files=True, key="new_files_tab2")
        
    if st.button("AI 갱신 필요 여부 확인 (기준DB 대조)"):
        local_base_files = get_local_base_files(BASE_FOLDER_NAME)
        
        if not local_base_files:
            st.warning(f"기준 데이터를 가져올 수 없습니다. 깃허브 동일 경로에 '{BASE_FOLDER_NAME}' 폴더가 존재하는지 확인해 주세요.")
        elif not new_files_tab2:
            st.warning("비교할 최신 데이터를 업로드해주세요.")
        else:
            with st.spinner('배합비 대조 중...'):
                result_df = compare_ingredients(local_base_files, new_files_tab2)
                
            if not result_df.empty:
                st.success("AI 배합비 분석이 완료되었습니다.")
                
                changes = result_df[result_df['상태'] == '갱신 필요 (변경됨)']
                if not changes.empty:
                    st.error(f"총 {len(changes)}건의 실질적 배합비 변경이 감지되었습니다.")
                    
                    for idx, row in changes.iterrows():
                        st.markdown(f'''
                        <div style="background-color: #fff3cd; color: #856404; padding: 15px; border-radius: 5px; margin-bottom: 10px; border-left: 5px solid #ffeeba;">
                            {row['알림 메시지']}
                        </div>
                        ''', unsafe_allow_html=True)
                else:
                    st.info("실질적인 배합비가 변경된 품목이 없습니다.")
                    
                st.dataframe(result_df[['품목보고번호', '제품명', '원재료_기존', '원재료_신규', '상태']])
            else:
                st.warning("데이터를 찾을 수 없거나 형식이 일치하지 않습니다.")

with tab3:
    st.subheader("기존에 작성된 제품설명서 최신화 점검")
    st.markdown("기존에 만들어둔 **'다중 시트 제품설명서 엑셀 파일'**들을 올리고, 오늘 기준 **'최신 품목제조보고 목록'**과 대조하여 최신화(갱신)가 필요한 문서를 색출합니다.")
    
    col_spec, col_ref = st.columns(2)
    with col_spec:
        st.markdown("**기존 제품설명서 업로드**")
        spec_files = st.file_uploader("제품설명서 엑셀 파일 (다중 시트 포함, 복수 선택 가능)", type=['xlsx'], accept_multiple_files=True, key="spec_files")
    with col_ref:
        st.markdown("**최신 기준 데이터 업로드**")
        new_ref_files = st.file_uploader("최신 품목제조보고 데이터 (일반식품, 축산물 등)", type=['xls', 'xlsx'], accept_multiple_files=True, key="new_ref_files")

    if st.button("문서 최신화 점검 시작"):
        if spec_files and new_ref_files:
            with st.spinner('문서 내 데이터 전체 추출 및 AI 대조 중...'):
                check_df = check_existing_specs(spec_files, new_ref_files)
                
            if not check_df.empty:
                st.success("기존 제품설명서 점검이 완료되었습니다.")
                
                outdated = check_df[check_df['상태'] == '최신화 필요']
                valid = check_df[check_df['상태'] == '적합']
                
                if not outdated.empty:
                    st.error(f"업로드된 문서 중 총 {len(outdated)}건이 최신 기준과 달라 업데이트가 필요합니다. 알림을 확인하고 1번 탭에서 다시 생성해 주세요.")
                    for idx, row in outdated.iterrows():
                        st.markdown(f'''
                        <div style="background-color: #f8d7da; color: #721c24; padding: 15px; border-radius: 5px; margin-bottom: 10px; border-left: 5px solid #f5c6cb;">
                            {row['알림 메시지']}
                        </div>
                        ''', unsafe_allow_html=True)
                
                if not valid.empty:
                    st.info(f"총 {len(valid)}건의 문서가 최신 기준에 적합합니다.")
                    for idx, row in valid.iterrows():
                        st.markdown(f'''
                        <div style="background-color: #d4edda; color: #155724; padding: 15px; border-radius: 5px; margin-bottom: 10px; border-left: 5px solid #c3e6cb;">
                            {row['알림 메시지']}
                        </div>
                        ''', unsafe_allow_html=True)
                    
                st.dataframe(check_df[['파일명', '품목보고번호', '제품명', '상태']])
        else:
            st.warning("점검할 '기존 제품설명서'와 대조할 '최신 기준 데이터'를 모두 업로드해주세요.")
