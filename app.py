import streamlit as st
import openpyxl
from openpyxl.styles import Font
from openpyxl.drawing.image import Image
from PIL import Image as PILImage
import io
import re

st.set_page_config(page_title="자동창고 보고서 생성기", layout="centered")
st.title("📝 현장 보고서 자동 생성기")

# ==========================================
# 0. 작업 분류 (공사 / 점검)
# ==========================================
st.markdown("### 📋 작업 분류 선택")
task_type = st.radio("어떤 종류의 보고서를 작성하시나요?", ["점검", "공사"], horizontal=True)

st.divider()

# ==========================================
# 1. 대분류 / 중분류 현장 선택
# ==========================================
st.markdown("### 🏢 업체 및 현장 정보")

site_dict = {
    "MXROBOTICS": ["태준제약", "직접 입력..."],
    "SFA SERVICE": ["BGF 로지스 광주 현장", "BGF 로지스 진천 현장", "직접 입력..."],
    "BLUEONE": ["유한킴벌리 충주공장", "직접 입력..."]
}

col_v, col_s = st.columns(2)
with col_v:
    vendor = st.selectbox("업체 대분류", list(site_dict.keys()))
with col_s:
    site_select = st.selectbox("현장명 중분류", site_dict[vendor])

if site_select == "직접 입력...":
    site_name = st.text_input("새로운 현장명을 직접 입력해주세요")
else:
    site_name = site_select

address = st.text_input("현장 주소", placeholder="예: 경기도 용인시...")

st.divider()

# ==========================================
# 2. 작업자 및 기간 입력 
# ==========================================
col1, col2 = st.columns(2)
with col1:
    author = st.text_input("작성자", value="지창현")
with col2:
    manager = st.selectbox("담당자 선택", ["김주영 책임", "최진명 차장", "조상길 부장", "직접 입력..."])

date_range = st.date_input("작업 일자 (기간 선택)", [])
date_str = ""
if len(date_range) == 2:
    start, end = date_range
    if start.month == end.month:
        date_str = f"{start.strftime('%y. %m. %d')} ~ {end.strftime('%d')}"
    else:
        date_str = f"{start.strftime('%y. %m. %d')} ~ {end.strftime('%m. %d')}"

workers = st.text_input("작업자명 및 인원", placeholder="예: 최진명 차장 외 6명")

# ==========================================
# 3. 작업 내용 메모장
# ==========================================
st.markdown(f"**{task_type} 내용 (조치 완료 우선 ➔ 교체 필요 항목 후순위)**")
contents = st.text_area("숫자 번호는 안 적으셔도 됩니다. 기본 항목(1번) 다음인 2번부터 자동 시작됩니다.", height=150)

# ==========================================
# [신규] 4. 스마트 사진 업로드 및 개별 설명 입력창
# ==========================================
st.divider()
st.markdown("### 📷 현장 사진 업로드 (선택사항)")
uploaded_photos = st.file_uploader("사진을 업로드하면 '사진 대장'이 추가로 생성됩니다.", type=['png', 'jpg', 'jpeg'], accept_multiple_files=True)

photo_descriptions = []
if uploaded_photos:
    st.info(f"총 {len(uploaded_photos)}장의 사진이 업로드되었습니다. 각 사진의 설명을 입력해주세요.")
    for i, photo in enumerate(uploaded_photos):
        # 업로드한 사진의 개수만큼 설명 입력칸이 자동으로 생성됨
        desc = st.text_input(f"[{i+1}번 사진] '{photo.name}' 내용", placeholder=f"예: S/C #{i+1}호기 부품 교체 전/후")
        photo_descriptions.append(desc)


# --- 정렬 로직 ---
def sort_rules(text):
    is_need_replace = 1 if "교체 필요" in text else 0
    match = re.search(r'(#)?(\d+)호기', text)
    ho_number = int(match.group(2)) if match else 9999
    return (is_need_replace, ho_number)

# ==========================================
# 5. 엑셀 생성 실행
# ==========================================
st.divider()
if st.button(f"🚀 {vendor} {task_type}보고서 생성하기", use_container_width=True):
    if not site_name or not contents or not date_str or not author:
        st.warning("작성자, 현장명, 날짜, 내용은 필수입니다!")
    else:
        with st.spinner("엑셀 파일을 만들고 있습니다..."):
            try:
                prefix_map = {"MXROBOTICS": "mxr", "SFA SERVICE": "sfa", "BLUEONE": "blueone"}
                template_filename = f"template_{prefix_map[vendor]}_{task_type}.xlsx"

                # --------------------------------------------------
                # [A] 보고서 엑셀 처리
                # --------------------------------------------------
                wb_report = openpyxl.load_workbook(template_filename)
                ws_report = wb_report.active

                ws_report['C5'] = site_name
                ws_report['C6'] = address
                ws_report['C9'] = f"{site_name} 정기 {task_type}" 
                ws_report['C10'] = date_str
                ws_report['I10'] = workers
                ws_report['H6'] = author    
                ws_report['C8'] = manager   
                
                # 1. 기존 잔여 데이터 깨끗하게 지우기 (18줄 ~ 38줄)
                for r in range(18, 39):
                    ws_report.cell(row=r, column=2).value = None
                    ws_report.cell(row=r, column=2).font = Font(name='맑은 고딕', size=11, color="000000")
                
                # 2. 사용자 입력 내용 정렬 및 기입 (18번째 줄, 2번부터 시작)
                raw_lines = [line.strip() for line in contents.split('\n') if line.strip()]
                sorted_lines = sorted(raw_lines, key=sort_rules)
                
                start_row = 18
                for idx, text in enumerate(sorted_lines):
                    if start_row + idx > 38: break # 38줄 넘침 임시 방지
                    
                    cell = ws_report.cell(row=start_row + idx, column=2)
                    cell.value = f"{idx + 2}. {text}" # 2번부터 시작하도록 +2 적용
                    
                    base_font = ws_report.cell(row=11, column=2).font
                    if "교체 필요" in text:
                        cell.font = Font(name=base_font.name, size=base_font.size, bold=True, color="FF0000")
                    elif "조치" in text or "교체" in text:
                        cell.font = Font(name=base_font.name, size=base_font.size, bold=True, color="0000FF")
                    else:
                        cell.font = Font(name=base_font.name, size=base_font.size, bold=base_font.bold, color="000000")

                output_report = io.BytesIO()
                wb_report.save(output_report)
                output_report.seek(0)
                
                # --------------------------------------------------
                # [B] 사진 대장 엑셀 처리 (사진이 있을 때만 실행!)
                # --------------------------------------------------
                output_photo = None
                if uploaded_photos:
                    wb_photo = openpyxl.load_workbook('photo_template.xlsx')
                    ws_photo = wb_photo.active
                    
                    ws_photo['B5'] = f"{site_name} 자동화 창고 {task_type} 사진"
                    ws_photo['I10'] = f"1. 현장명 : {site_name}"
                    ws_photo['I14'] = f"3. 작업일자 : {date_str}"
                    ws_photo['I15'] = f"4. 작업인원 : {workers}"

                    # 사진이 들어갈 셀 좌표 (B10, D10, B29, D29 ...)
                    photo_coords = [('B', 10), ('D', 10), ('B', 29), ('D', 29), ('B', 48), ('D', 48), ('B', 67), ('D', 67)]
                    
                    for i, photo_file in enumerate(uploaded_photos):
                        if i >= len(photo_coords): break 
                        col, row = photo_coords[i]
                        
                        img_pil = PILImage.open(photo_file)
                        img_pil.thumbnail((350, 350)) 
                        img_byte_arr = io.BytesIO()
                        img_pil.save(img_byte_arr, format='PNG')
                        img_byte_arr.seek(0)
                        xl_img = Image(img_byte_arr)
                        ws_photo.add_image(xl_img, f"{col}{row}")
                        
                        # 사진 설명 텍스트 기입 (사진 칸에서 16칸 아래가 텍스트 칸이라고 가정)
                        desc_row = row + 16 
                        ws_photo[f"{col}{desc_row}"] = photo_descriptions[i]
                            
                    output_photo = io.BytesIO()
                    wb_photo.save(output_photo)
                    output_photo.seek(0)

                # --------------------------------------------------
                # [C] 결과 및 다운로드 버튼 출력
                # --------------------------------------------------
                st.success("🎉 생성이 완료되었습니다!")
                
                col1, col2 = st.columns(2)
                vendor_prefix = prefix_map[vendor].upper()
                
                with col1:
                    st.download_button(
                        label=f"📥 [{task_type} 보고서] 다운로드",
                        data=output_report,
                        file_name=f"({vendor_prefix}_{task_type}보고서){site_name}_{date_str[:8].replace('. ', '')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                # 사진 파일이 만들어졌을 때만 2번째 다운로드 버튼 띄우기
                if output_photo:
                    with col2:
                        st.download_button(
                            label=f"🖼️ [{task_type} 사진] 다운로드",
                            data=output_photo,
                            file_name=f"({vendor_prefix}_{task_type}사진){site_name}_{date_str[:8].replace('. ', '')}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True
                        )

            except Exception as e:
                st.error(f"에러 발생! 깃허브에 템플릿 파일이 모두 있는지 확인해주세요. 에러내용: {e}")
