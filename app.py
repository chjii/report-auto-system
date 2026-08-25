import streamlit as st
import openpyxl
from openpyxl.styles import Font, Alignment
from openpyxl.drawing.image import Image
from PIL import Image as PILImage
import io
import re
import json
import os

st.set_page_config(page_title="자동창고 보고서 생성기", layout="centered")
st.title("📝 현장 보고서 자동 생성기")

# ==========================================
# [신규] 현장 기억 데이터 불러오기/저장하기 함수
# ==========================================
DATA_FILE = "site_memory.json"

def load_memory():
    if os.path.exists(DATA_FILE):
        with open(DATA_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}

def save_memory(data):
    with open(DATA_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=4)

memory_db = load_memory()

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

# 💡 기본 목록 외에, 기억상자에 있는 현장들도 '직접 입력...' 위쪽에 불러옵니다.
site_dict = {
    "MXROBOTICS": ["태준제약"],
    "SFA SERVICE": ["BGF 로지스 광주 현장", "BGF 로지스 진천 현장"],
    "BLUEONE": ["유한킴벌리 충주공장"]
}

col_v, col_s = st.columns(2)
with col_v:
    vendor = st.selectbox("업체 대분류", list(site_dict.keys()))

# 기억상자(memory_db)에서 해당 업체의 현장들을 추가로 가져와 목록에 합침
saved_sites = [site for site, info in memory_db.items() if info.get("vendor") == vendor and site not in site_dict[vendor]]
current_site_list = site_dict[vendor] + saved_sites + ["직접 입력..."]

with col_s:
    site_select = st.selectbox("현장명 중분류", current_site_list)

# 자동 완성을 위한 기본값 세팅
default_address = ""
default_manager = "김주영 책임"

if site_select == "직접 입력...":
    site_name = st.text_input("새로운 현장명을 직접 입력해주세요 (필수)")
else:
    site_name = site_select
    # 기존 현장을 고르면 기억상자에서 주소와 담당자를 꺼내옵니다.
    if site_name in memory_db:
        default_address = memory_db[site_name].get("address", "")
        default_manager = memory_db[site_name].get("manager", "김주영 책임")

# 꺼내온 주소를 기본값(value)으로 넣음. 당연히 직접 수정도 가능!
address = st.text_input("현장 주소", value=default_address, placeholder="예: 경기도 용인시...")

st.divider()

# ==========================================
# 2. 작업자 및 기간 입력 
# ==========================================
col1, col2 = st.columns(2)
with col1:
    author = st.text_input("작성자", value="지창현")
with col2:
    manager_list = ["김주영 책임", "최진명 차장", "조상길 부장"]
    # 기억상자에서 꺼낸 담당자가 기본 목록에 없으면 맨 앞에 추가해줌
    if default_manager not in manager_list:
        manager_list.insert(0, default_manager)
    manager_list.append("직접 입력...")
    
    manager_select = st.selectbox("담당자 선택", manager_list, index=manager_list.index(default_manager))

if manager_select == "직접 입력...":
    manager = st.text_input("담당자명을 직접 입력해주세요 (필수)")
else:
    manager = manager_select

date_range = st.date_input("작업 일자 (기간 선택)", [])
date_str = ""
if len(date_range) == 2:
    start, end = date_range
    if start.month == end.month:
        date_str = f"{start.strftime('%y. %m. %d')} ~ {end.strftime('%d')}"
    else:
        date_str = f"{start.strftime('%y. %m. %d')} ~ {end.strftime('%m. %d')}"

workers = st.text_input("작업자명 및 인원", placeholder="예: 최진명 차장 외 6명")

st.divider()

# ==========================================
# 3. 설비 선택 및 기본 점검 항목 세팅
# ==========================================
st.markdown("### ⚙️ 점검 진행 설비 선택")
equipments = st.multiselect(
    "현장에서 점검한 설비를 모두 선택하세요 (선택 시 기본 항목 자동 출력)", 
    ["STACKER CRANE", "CONVEYOR", "RGV", "LIFT"], 
    default=["STACKER CRANE"]
)

DEFAULT_TEXTS = {
    "STACKER CRANE": [
        "1. S/C 점검 공통사항",
        "  1) 승강부,주행부,FORK부 구동 MOTOR 및 감속기 발열 상태 및 OIL 누유 상태 점검",
        "  2) CARRIAGE INNER ROLLER, GUIDE ROLLER 구름 상태 및 마모 상태 점검",
        "  3) FORK CHAIN TENSION 점검 및 C/F BEARING, MC GUIDE GREASE 도포",
        "  4) STC(기상반) 단자대 풀림 상태 CHECK 및 재조임",
        "  5) 주행부 구동 WHEEL,종동 WHEEL, GUIDE ROLLER 구름 상태 및 마모 상태 점검"
    ],
    "CONVEYOR": [
        "1. C/V 점검 공통사항",
        "  1) 구동 MOTOR 및 감속기 발열/소음 상태 및 OIL 누유 상태 점검",
        "  2) 체인/벨트 장력 상태 및 마모 상태 점검",
        "  3) 구동/종동 ROLLER 구름 상태 점검 및 베어링 소음 확인",
        "  4) 센서(광전, 근접 등) 취부 상태 및 동작 상태 점검"
    ],
    "RGV": [
        "1. RGV 점검 공통사항",
        "  1) 주행부 구동 MOTOR 발열 및 소음, 누유 상태 점검",
        "  2) 주행 WHEEL 및 GUIDE ROLLER 마모 상태 점검",
        "  3) 집전기(Collector) 마모 상태 및 단자대 조임 상태 점검",
        "  4) 충돌 방지 센서 및 통신 장치 상태 점검"
    ],
    "LIFT": [
        "1. LIFT 점검 공통사항",
        "  1) 승강 MOTOR 및 감속기 소음/발열, 누유 상태 점검",
        "  2) 승강 CHAIN 및 장력, 마모 상태 점검",
        "  3) GUIDE ROLLER 구름 상태 및 마모 상태 점검",
        "  4) 상/하한 리미트 센서 및 낙하 방지 장치 동작 상태 점검"
    ]
}

# ==========================================
# 4. 작업 내용 메모장
# ==========================================
st.markdown(f"**{task_type} 내용 (설비별 자동 분류 및 2번부터 자동 넘버링)**")
contents = st.text_area("S/C, CV, RGV, LIFT 등 키워드를 포함해서 적어주시면 코드가 알아서 각 설비 구역으로 나눠서 정리합니다.", height=150)

# ==========================================
# 5. 스마트 사진 업로드 및 개별 설명 입력창
# ==========================================
st.divider()
st.markdown("### 📷 현장 사진 업로드 (선택사항)")
uploaded_photos = st.file_uploader("사진을 업로드하면 '사진 대장'이 추가로 생성됩니다.", type=['png', 'jpg', 'jpeg'], accept_multiple_files=True)

photo_descriptions = []
if uploaded_photos:
    st.info(f"총 {len(uploaded_photos)}장의 사진이 업로드되었습니다. 각 사진의 설명을 입력해주세요.")
    for i, photo in enumerate(uploaded_photos):
        desc = st.text_input(f"[{i+1}번 사진] '{photo.name}' 내용", placeholder=f"예: S/C #{i+1}호기 부품 교체 전/후")
        photo_descriptions.append(desc)

def sort_rules(text):
    is_need_replace = 1 if "교체 필요" in text else 0
    match = re.search(r'(#)?(\d+)호기', text)
    ho_number = int(match.group(2)) if match else 9999
    return (is_need_replace, ho_number)

def write_equipment_block(ws, eq_title, defaults, user_lines, row_idx):
    ws.cell(row=row_idx, column=2).value = eq_title
    ws.cell(row=row_idx, column=2).font = Font(name='맑은 고딕', size=11, bold=True)
    ws.cell(row=row_idx, column=2).alignment = Alignment(horizontal='center', vertical='center')
    row_idx += 1
    
    for i, df_text in enumerate(defaults):
        ws.cell(row=row_idx, column=2).value = df_text
        ws.cell(row=row_idx, column=2).font = Font(name='맑은 고딕', size=11, bold=(i==0), color="000000")
        ws.cell(row=row_idx, column=2).alignment = Alignment(horizontal='left', vertical='center')
        row_idx += 1
        
    for idx, text in enumerate(user_lines):
        cell = ws.cell(row=row_idx, column=2)
        cell.value = f"{idx + 2}. {text}"
        
        if "교체 필요" in text:
            cell.font = Font(name='맑은 고딕', size=11, bold=True, color="FF0000")
        elif "조치" in text or "교체" in text:
            cell.font = Font(name='맑은 고딕', size=11, bold=True, color="0000FF")
        else:
            cell.font = Font(name='맑은 고딕', size=11, bold=False, color="000000")
            
        cell.alignment = Alignment(horizontal='left', vertical='center')
        row_idx += 1
    
    row_idx += 1 
    return row_idx

# ==========================================
# 6. 엑셀 생성 실행
# ==========================================
st.divider()
if st.button(f"🚀 {vendor} {task_type}보고서 생성하기", use_container_width=True):
    if not site_name or not date_str or not author or not manager:
        st.warning("작성자, 담당자, 현장명, 날짜는 필수입니다!")
    elif not equipments:
        st.warning("점검 진행 설비를 최소 1개 이상 선택해주세요!")
    else:
        with st.spinner("엑셀 파일을 만들고 있습니다..."):
            try:
                # 💡 [신규] 생성 버튼을 누를 때, 입력한 현장/주소/담당자 정보를 파일에 기억시킵니다!
                memory_db[site_name] = {
                    "vendor": vendor,
                    "address": address,
                    "manager": manager
                }
                save_memory(memory_db)

                prefix_map = {"MXROBOTICS": "mxr", "SFA SERVICE": "sfa", "BLUEONE": "blueone"}
                template_filename = f"template_{prefix_map[vendor]}_{task_type}.xlsx"

                # [A] 보고서 엑셀 처리
                wb_report = openpyxl.load_workbook(template_filename)
                ws_report = wb_report.active

                ws_report['C5'] = site_name
                ws_report['C6'] = address
                ws_report['C9'] = f"{site_name} 정기 {task_type}" 
                ws_report['C10'] = date_str
                ws_report['I10'] = workers
                ws_report['H6'] = author    
                ws_report['C8'] = manager   
                
                for r in range(12, 60):
                    ws_report.cell(row=r, column=2).value = None
                    ws_report.cell(row=r, column=2).font = Font(name='맑은 고딕', size=11, color="000000")
                
                raw_lines = [line.strip() for line in contents.split('\n') if line.strip()]
                sorted_lines = sorted(raw_lines, key=sort_rules)
                
                sc_lines, cv_lines, rgv_lines, lift_lines = [], [], [], []
                for line in sorted_lines:
                    u_line = line.upper()
                    if "RGV" in u_line: rgv_lines.append(line)
                    elif "LIFT" in u_line or "리프트" in u_line: lift_lines.append(line)
                    elif "CV" in u_line or "CONVEYOR" in u_line or "컨베이어" in u_line: cv_lines.append(line)
                    elif "S/C" in u_line or "STC" in u_line or "크레인" in u_line or "호기" in u_line: sc_lines.append(line)
                    else: sc_lines.append(line) 

                current_row = 12
                
                if "STACKER CRANE" in equipments:
                    current_row = write_equipment_block(ws_report, " - STACKER CRANE - ", DEFAULT_TEXTS["STACKER CRANE"], sc_lines, current_row)
                if "CONVEYOR" in equipments:
                    current_row = write_equipment_block(ws_report, " - CONVEYOR - ", DEFAULT_TEXTS["CONVEYOR"], cv_lines, current_row)
                if "RGV" in equipments:
                    current_row = write_equipment_block(ws_report, " - RGV - ", DEFAULT_TEXTS["RGV"], rgv_lines, current_row)
                if "LIFT" in equipments:
                    current_row = write_equipment_block(ws_report, " - LIFT - ", DEFAULT_TEXTS["LIFT"], lift_lines, current_row)

                output_report = io.BytesIO()
                wb_report.save(output_report)
                output_report.seek(0)
                
                # [B] 사진 대장 엑셀 처리
                output_photo = None
                if uploaded_photos:
                    wb_photo = openpyxl.load_workbook('photo_template.xlsx')
                    ws_photo = wb_photo.active
                    
                    ws_photo['B5'] = f"{site_name} 자동화 창고 {task_type} 사진"
                    ws_photo['I10'] = f"1. 현장명 : {site_name}"
                    ws_photo['I14'] = f"3. 작업일자 : {date_str}"
                    ws_photo['I15'] = f"4. 작업인원 : {workers}"

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
                        
                        desc_row = row + 16 
                        ws_photo[f"{col}{desc_row}"] = photo_descriptions[i]
                            
                    output_photo = io.BytesIO()
                    wb_photo.save(output_photo)
                    output_photo.seek(0)

                # [C] 다운로드 버튼 출력
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
