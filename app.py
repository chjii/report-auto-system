import streamlit as st
import openpyxl
from openpyxl.styles import Font
from openpyxl.drawing.image import Image
from PIL import Image as PILImage
import io
import re

st.set_page_config(page_title="자동창고 점검보고서", layout="centered")
st.title("📝 현장 점검보고서 자동 생성기")

# 1. 정보 입력 (달력 포함)
site_name = st.text_input("현장명", placeholder="예: 태준제약")
address = st.text_input("현장 주소", placeholder="예: 경기도 용인시")

date_range = st.date_input("작업 일자 (기간 선택)", [])
date_str = ""
if len(date_range) == 2:
    start, end = date_range
    if start.month == end.month:
        date_str = f"{start.strftime('%y. %m. %d')} ~ {end.strftime('%d')}"
    else:
        date_str = f"{start.strftime('%y. %m. %d')} ~ {end.strftime('%m. %d')}"

workers = st.text_input("점검자명 및 인원", placeholder="예: 최진명 차장 외 6명")

# 2. 점검 내용 메모장
st.markdown("**점검 내용 (조치 완료 항목 우선 ➔ 교체 필요 항목 후순위 정렬)**")
contents = st.text_area("숫자 번호(1. 2. 3.)는 안 적으셔도 됩니다.", height=200)

uploaded_photos = st.file_uploader("현장 사진 업로드 (여러 장 선택 가능)", type=['png', 'jpg', 'jpeg'], accept_multiple_files=True)

# --- 2단계 정렬 로직 ---
def sort_rules(text):
    is_need_replace = 1 if "교체 필요" in text else 0
    match = re.search(r'(#)?(\d+)호기', text)
    ho_number = int(match.group(2)) if match else 9999
    return (is_need_replace, ho_number)

# 3. 엑셀 생성 실행
if st.button("🚀 전체 엑셀 생성하기 (보고서 + 사진대장)"):
    if not site_name or not contents or not date_str:
        st.warning("현장명, 날짜(시작/종료 모두 지정), 점검 내용은 필수입니다!")
    else:
        with st.spinner("엑셀 파일을 만들고 있습니다..."):
            try:
                # ==========================================
                # [1] 점검 보고서 (template.xlsx) 생성 로직
                # ==========================================
                wb_report = openpyxl.load_workbook('template.xlsx')
                ws_report = wb_report.active

                ws_report['C5'] = site_name
                ws_report['C6'] = address
                ws_report['C9'] = f"{site_name} 정기 점검" 
                ws_report['C10'] = date_str
                ws_report['I10'] = workers
                
                raw_lines = [line.strip() for line in contents.split('\n') if line.strip()]
                sorted_lines = sorted(raw_lines, key=sort_rules)
                
                start_row = 12
                for idx, text in enumerate(sorted_lines):
                    cell = ws_report.cell(row=start_row + idx, column=2)
                    cell.value = f"{idx + 1}. {text}"
                    
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
                
                # ==========================================
                # [2] 점검 사진 (photo_template.xlsx) 생성 로직
                # ==========================================
                wb_photo = openpyxl.load_workbook('photo_template.xlsx')
                ws_photo = wb_photo.active
                
                # 사진 대장 기본 정보 덮어쓰기 (아까 맞춰둔 좌표)
                ws_photo['B5'] = f"{site_name} 자동화 창고 점검 사진"
                ws_photo['I10'] = f"1. 현장명 : {site_name}"
                ws_photo['I14'] = f"3. 작업일자 : {date_str}"
                ws_photo['I15'] = f"4. 작업인원 : {workers}"

                # 업로드된 사진들을 엑셀 셀 위치에 맞게 넣기
                # (예시 좌푯값: 좌측 B10, 우측 D10, 그 다음줄 B29, D29 ...)
                photo_cells = ['B10', 'D10', 'B29', 'D29', 'B48', 'D48', 'B67', 'D67']
                
                if uploaded_photos:
                    for i, photo_file in enumerate(uploaded_photos):
                        if i >= len(photo_cells):
                            break # 준비된 칸보다 사진이 많으면 일단 스킵
                        
                        # 이미지 열기 및 엑셀 칸 크기에 맞춰 축소 (가로 약 350px 기준)
                        img_pil = PILImage.open(photo_file)
                        img_pil.thumbnail((350, 350)) 
                        
                        # 메모리에 임시 저장 후 openpyxl 이미지로 변환
                        img_byte_arr = io.BytesIO()
                        img_pil.save(img_byte_arr, format='PNG')
                        img_byte_arr.seek(0)
                        
                        xl_img = Image(img_byte_arr)
                        
                        # 지정된 셀 위치에 이미지 삽입
                        target_cell = photo_cells[i]
                        ws_photo.add_image(xl_img, target_cell)
                        
                output_photo = io.BytesIO()
                wb_photo.save(output_photo)
                output_photo.seek(0)

                # ==========================================
                # [3] 다운로드 버튼 표시
                # ==========================================
                st.success("🎉 보고서와 사진 대장이 모두 완성되었습니다!")
                
                col1, col2 = st.columns(2)
                with col1:
                    st.download_button(
                        label="📥 [점검 보고서] 다운로드",
                        data=output_report,
                        file_name=f"(점검보고서){site_name}_{date_str[:8].replace('. ', '')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                with col2:
                    st.download_button(
                        label="🖼️ [점검 사진] 다운로드",
                        data=output_photo,
                        file_name=f"(점검사진){site_name}_{date_str[:8].replace('. ', '')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

            except Exception as e:
                st.error(f"에러가 발생했습니다. 파일 이름을 'template.xlsx', 'photo_template.xlsx'로 정확히 올리셨는지 확인해주세요. 에러내용: {e}")
