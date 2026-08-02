import streamlit as st
import openpyxl
from openpyxl.styles import Font
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
contents = st.text_area("숫자 번호(1. 2. 3.)는 안 적으셔도 됩니다.", height=200, 
                        placeholder="S/C #3호기 가이드 롤러 교체 필요\nS/C #5호기 체인 장력 조치\nS/C #1호기 센서 교체")

uploaded_photos = st.file_uploader("현장 사진 업로드 (여러 장 선택 가능)", accept_multiple_files=True)

# --- 뒷단 2단계 정렬 로직 ---
def sort_rules(text):
    # 1. 상태값 확인: '교체 필요'가 들어가면 우선순위가 뒤(1)로 밀림, 완료건은 앞(0)
    is_need_replace = 1 if "교체 필요" in text else 0
    
    # 2. 호기 숫자 추출
    match = re.search(r'(#)?(\d+)호기', text)
    ho_number = int(match.group(2)) if match else 9999
    
    # 리턴: (상태 우선순위, 호기 숫자) 순으로 2단계 정렬
    return (is_need_replace, ho_number)

# 3. 엑셀 생성 실행
if st.button("🚀 보고서 엑셀 생성하기"):
    if not site_name or not contents or not date_str:
        st.warning("현장명, 날짜(시작/종료 모두 지정), 점검 내용은 필수입니다!")
    else:
        try:
            wb = openpyxl.load_workbook('template.xlsx')
            ws = wb.active

            ws['C5'] = site_name
            ws['C6'] = address
            ws['C9'] = f"{site_name} 정기 점검" 
            ws['C10'] = date_str
            ws['I10'] = workers
            
            # 입력 텍스트 줄바꿈으로 나누기
            raw_lines = [line.strip() for line in contents.split('\n') if line.strip()]
            
            # 🌟 마법의 2단계 정렬 실행 (파란글씨 호기순 -> 빨간글씨 호기순)
            sorted_lines = sorted(raw_lines, key=sort_rules)
            
            # 엑셀에 쓰기 및 색상 입히기
            start_row = 12
            for idx, text in enumerate(sorted_lines):
                cell = ws.cell(row=start_row + idx, column=2)
                # 앞쪽에 자동으로 1. 2. 3. 통합 넘버링 부여
                cell.value = f"{idx + 1}. {text}"
                
                base_font = ws.cell(row=11, column=2).font
                
                if "교체 필요" in text:
                    cell.font = Font(name=base_font.name, size=base_font.size, bold=True, color="FF0000")
                elif "조치" in text or "교체" in text:
                    cell.font = Font(name=base_font.name, size=base_font.size, bold=True, color="0000FF")
                else:
                    cell.font = Font(name=base_font.name, size=base_font.size, bold=base_font.bold, color="000000")

            output = io.BytesIO()
            wb.save(output)
            output.seek(0)

            st.success("🎉 보고서 생성이 완료되었습니다! 아래 버튼을 눌러 다운로드하세요.")
            st.download_button(
                label="📥 엑셀 파일 다운로드",
                data=output,
                file_name=f"(점검보고서){site_name}_{date_str[:8].replace('. ', '')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except Exception as e:
            st.error(f"에러가 발생했습니다. 원본 엑셀 파일(template.xlsx) 확인 필요. 에러내용: {e}")