import streamlit as st
import openpyxl
from openpyxl.styles import Font, Alignment
from openpyxl.drawing.image import Image
from PIL import Image as PILImage
import io
import re
import json
import os
from copy import copy

st.set_page_config(page_title="자동창고 보고서 생성기", layout="centered")
st.title("📝 현장 보고서 자동 생성기")

# ==========================================
# 0. 55개 고객사 마스터 데이터베이스
# ==========================================
BASE_SITE_DB = {
    "아이티센엔텍": {"vendor": "블루원", "address": "강원도 인제군", "equipments": ["STACKER CRANE", "CONVEYOR"], "pm": "김은호 수석"},
    "LGL(롯데글로벌로지스)": {"vendor": "블루원", "address": "인천광역시 영종구", "equipments": ["CONVEYOR"], "pm": "김은기"},
    "일성IS": {"vendor": "블루원", "address": "경기도 안산시", "equipments": ["STACKER CRANE", "CONVEYOR", "RGV"], "pm": "권욱태 부장"},
    "우진플라임": {"vendor": "블루원", "address": "충청북도 보은군", "equipments": ["STACKER CRANE", "CONVEYOR", "RGV"], "pm": "백기철 차장"},
    "일화": {"vendor": "블루원", "address": "강원도 춘천시", "equipments": ["STACKER CRANE", "CONVEYOR", "RGV"], "pm": "김영욱 부장"},
    "농심": {"vendor": "농심 엔지니어링", "address": "경기도 안양시", "equipments": ["STACKER CRANE", "CONVEYOR"], "pm": "김재빈"},
    "아모레퍼시픽 물류": {"vendor": "SFA서비스", "address": "경기도 오산시", "equipments": ["STACKER CRANE", "CONVEYOR", "RGV"], "pm": "김선용 차장"},
    "아모레퍼시픽 생산": {"vendor": "SFA서비스", "address": "경기도 오산시", "equipments": ["RGV"], "pm": "김선용"},
    "BGF 광주": {"vendor": "SFA서비스", "address": "경기도 광주시", "equipments": ["STACKER CRANE", "CONVEYOR"], "pm": "김선용"},
    "BGF 진천": {"vendor": "SFA서비스", "address": "충청북도 진천군", "equipments": ["STACKER CRANE"], "pm": "김선용"},
    "SK하이닉스": {"vendor": "SFA서비스", "address": "경기도 이천시", "equipments": ["STACKER CRANE", "CONVEYOR"], "pm": "정용욱 차장"},
    "이마트24": {"vendor": "SFA서비스", "address": "경기도 평택시", "equipments": ["STACKER CRANE"], "pm": "박영대 차장"},
    "한국오츠카 A": {"vendor": "SFA서비스", "address": "경기도 화성시", "equipments": ["STACKER CRANE", "CONVEYOR", "RGV"], "pm": "윤정현 과장"},
    "한국오츠카 B": {"vendor": "SFA서비스", "address": "경기도 화성시", "equipments": ["STACKER CRANE", "CONVEYOR", "RGV"], "pm": "윤정현"},
    "동아ST": {"vendor": "SFA서비스", "address": "인천광역시 연수구", "equipments": ["STACKER CRANE", "CONVEYOR", "RGV"], "pm": "박영대"},
    "포스코퓨처엠": {"vendor": "SFA서비스", "address": "충청북도 세종시", "equipments": ["STACKER CRANE", "CONVEYOR", "RGV"], "pm": "정용욱"},
    "BGF 진천(외자)": {"vendor": "DLS", "address": "충청북도 진천군", "equipments": ["RGV"], "pm": "허민재 수석"},
    "다이소(외자)": {"vendor": "DLS", "address": "부산광역시 강서구", "equipments": ["RGV"], "pm": "이창준 수석"},
    "아모레퍼시픽(외자)": {"vendor": "DLS", "address": "경기도 오산시", "equipments": ["RGV"], "pm": "허민재 수석"},
    "한미약품 팔탄공장 WMS 1": {"vendor": "MXRobotics", "address": "경기도 화성시", "equipments": ["STACKER CRANE"], "pm": "김대호 선임"},
    "한미약품 팔탄공장 WMS 2": {"vendor": "MXRobotics", "address": "경기도 화성시", "equipments": ["CONVEYOR"], "pm": "김대호"},
    "한미약품 팔탄공장 PMS": {"vendor": "MXRobotics", "address": "경기도 화성시", "equipments": ["CONVEYOR", "RGV", "LIFT"], "pm": "김대호"},
    "한미약품 팔탄공장 APS": {"vendor": "MXRobotics", "address": "경기도 화성시", "equipments": ["CONVEYOR"], "pm": "김대호"},
    "한미약품 평택공장 바이오": {"vendor": "MXRobotics", "address": "경기도 평택시", "equipments": ["STACKER CRANE", "CONVEYOR", "RGV"], "pm": "남재준"},
    "한미약품 평택공장 고형제": {"vendor": "MXRobotics", "address": "경기도 평택시", "equipments": ["STACKER CRANE", "CONVEYOR"], "pm": "남재준 선임"},
    "녹십자(구)": {"vendor": "MXRobotics", "address": "충청북도 청주시 청원구", "equipments": ["STACKER CRANE", "CONVEYOR", "RGV"], "pm": "남재준"},
    "녹십자(혈장)": {"vendor": "MXRobotics", "address": "충청북도 청주시 청원구", "equipments": ["STACKER CRANE", "CONVEYOR", "RGV"], "pm": "남재준"},
    "명문제약": {"vendor": "MXRobotics", "address": "경기도 화성시", "equipments": ["STACKER CRANE", "CONVEYOR", "RGV"], "pm": "남재준"},
    "태준제약": {"vendor": "MXRobotics", "address": "경기도 용인시", "equipments": ["STACKER CRANE", "CONVEYOR", "RGV"], "pm": "남재준"},
    "동국제약": {"vendor": "MXRobotics", "address": "충청북도 진천군", "equipments": ["STACKER CRANE", "CONVEYOR", "RGV"], "pm": "남재준"},
    "대원제약": {"vendor": "MXRobotics", "address": "충청북도 진천군", "equipments": ["STACKER CRANE", "CONVEYOR"], "pm": "남재준"},
    "제일약품": {"vendor": "MXRobotics", "address": "경기도 용인시", "equipments": ["STACKER CRANE", "CONVEYOR", "RGV"], "pm": "남재준"},
    "다이소 남사": {"vendor": "MXRobotics", "address": "경기도 용인시", "equipments": ["STACKER CRANE", "CONVEYOR", "RGV"], "pm": "남재준"},
    "삼양사 아산": {"vendor": "MXRobotics", "address": "충청남도 아산시", "equipments": ["RGV"], "pm": "남재준"},
    "피코이노베이션": {"vendor": "MXRobotics", "address": "경기도 평택시", "equipments": ["STACKER CRANE", "CONVEYOR", "RGV"], "pm": "남재준"},
    "메디톡스": {"vendor": "MXRobotics", "address": "충청북도 청주시 흥덕구", "equipments": ["STACKER CRANE", "CONVEYOR"], "pm": "남재준"},
    "연우": {"vendor": "MXRobotics", "address": "인천광역시", "equipments": ["STACKER CRANE", "CONVEYOR", "LIFT"], "pm": "김대호"},
    "기아자동차 화성": {"vendor": "MXRobotics", "address": "경기도 화성시", "equipments": ["STACKER CRANE"], "pm": "김대호"},
    "유한킴벌리 충주": {"vendor": "MXRobotics", "address": "충청북도 충주시", "equipments": ["STACKER CRANE", "CONVEYOR", "RGV"], "pm": "김주영 책임"},
    "유한킴벌리 김천": {"vendor": "MXRobotics", "address": "경상북도 김천시", "equipments": ["STACKER CRANE", "CONVEYOR", "RGV"], "pm": "김주영"},
    "HK이노엔": {"vendor": "MXRobotics", "address": "충청북도 청주시 흥덕구", "equipments": ["STACKER CRANE", "CONVEYOR", "RGV"], "pm": "채우석 선임"},
    "중외제약": {"vendor": "MXRobotics", "address": "충청남도 당진시", "equipments": ["STACKER CRANE", "CONVEYOR", "RGV"], "pm": "김주영"},
    "대웅제약": {"vendor": "MXRobotics", "address": "충청북도 청주시 흥덕구", "equipments": ["STACKER CRANE", "CONVEYOR", "RGV"], "pm": "채우석"},
    "사조동아원": {"vendor": "MXRobotics", "address": "충청남도 당진시", "equipments": ["STACKER CRANE"], "pm": "채우석"},
    "아트라스 대전": {"vendor": "MXRobotics", "address": "대전광역시 유성구", "equipments": ["STACKER CRANE", "CONVEYOR", "RGV"], "pm": "채우석"},
    "유니메드": {"vendor": "MXRobotics", "address": "충청북도 청주시 흥덕구", "equipments": ["STACKER CRANE", "CONVEYOR", "RGV"], "pm": "채우석"},
    "한국야금": {"vendor": "MXRobotics", "address": "충청북도 진천군", "equipments": ["STACKER CRANE", "CONVEYOR"], "pm": "김주영"},
    "한국타이어 금산": {"vendor": "MXRobotics", "address": "충청남도 금산군", "equipments": ["STACKER CRANE"], "pm": "김주영"},
    "SK케미칼": {"vendor": "MXRobotics", "address": "울산광역시 남구", "equipments": ["STACKER CRANE", "CONVEYOR", "RGV"], "pm": "공민수 선임"},
    "현대모비스": {"vendor": "MXRobotics", "address": "울산광역시 남구", "equipments": ["STACKER CRANE", "CONVEYOR"], "pm": "엄인형 책임"},
    "UPP": {"vendor": "MXRobotics", "address": "울산광역시 남구", "equipments": ["STACKER CRANE", "CONVEYOR", "RGV"], "pm": "한동현 선임"},
    "카펙발레오": {"vendor": "MXRobotics", "address": "경상북도 성주군", "equipments": ["STACKER CRANE", "CONVEYOR"], "pm": "김준현 책임"},
    "GGK": {"vendor": "MXRobotics", "address": "인천광역시 영종구", "equipments": ["STACKER CRANE", "CONVEYOR", "RGV"], "pm": "김대호"},
    "알피바이오": {"vendor": "MXRobotics", "address": "경기도 화성시 만세구", "equipments": ["STACKER CRANE", "CONVEYOR", "RGV"], "pm": "남재준"},
    "코오롱글로텍": {"vendor": "MXRobotics", "address": "충청남도 천안시", "equipments": ["STACKER CRANE", "CONVEYOR", "RGV"], "pm": "채우석"}
}

DATA_FILE = "site_memory.json"
DRAFT_FILE = "draft_memory.json"

def load_memory():
    if os.path.exists(DATA_FILE):
        with open(DATA_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}

def save_memory(data):
    with open(DATA_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=4)

def load_draft():
    if os.path.exists(DRAFT_FILE):
        with open(DRAFT_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {"contents": ""}

def save_draft(contents):
    with open(DRAFT_FILE, "w", encoding="utf-8") as f:
        json.dump({"contents": contents}, f, ensure_ascii=False)

site_db = copy(BASE_SITE_DB)
site_db.update(load_memory())

# ==========================================
# 1. 작업 분류 (공사 / 점검)
# ==========================================
st.markdown("### 📋 작업 분류 선택")
task_type = st.radio("보고서 종류", ["점검", "공사"], horizontal=True)

st.divider()

# ==========================================
# 2. 검색 및 현장 선택 (자동 완성 연동)
# ==========================================
st.markdown("### 🏢 현장 선택 및 자동 완성")

search_mode = st.radio("선택 방식", ["🏢 계약 업체별 필터링", "🔍 전체 현장 통합 검색"], horizontal=True)

vendors = sorted(list(set(info["vendor"] for info in site_db.values())))

if search_mode == "🏢 계약 업체별 필터링":
    col_v, col_s = st.columns(2)
    with col_v:
        selected_vendor = st.selectbox("계약 업체 선택", ["전체"] + vendors)
    
    if selected_vendor == "전체":
        filtered_sites = list(site_db.keys())
    else:
        filtered_sites = [s for s, info in site_db.items() if info["vendor"] == selected_vendor]
else:
    col_s = st.container()
    filtered_sites = list(site_db.keys())

site_options = filtered_sites + ["직접 입력..."]

if "target_site" not in st.session_state:
    st.session_state.target_site = site_options[0]
    first_info = site_db.get(site_options[0], {})
    st.session_state.addr_val = first_info.get("address", "")
    st.session_state.manager_val = first_info.get("pm", "남재준")
    st.session_state.equip_val = first_info.get("equipments", ["STACKER CRANE"])

def on_site_select_change():
    chosen = st.session_state.site_dropdown
    if chosen in site_db:
        st.session_state.addr_val = site_db[chosen].get("address", "")
        st.session_state.manager_val = site_db[chosen].get("pm", "")
        st.session_state.equip_val = site_db[chosen].get("equipments", ["STACKER CRANE"])
    elif chosen == "직접 입력...":
        st.session_state.addr_val = ""
        st.session_state.manager_val = ""
        st.session_state.equip_val = ["STACKER CRANE"]

with col_s:
    chosen_site = st.selectbox(
        "현장명 선택 (타이핑 검색 가능)", 
        site_options, 
        key="site_dropdown", 
        on_change=on_site_select_change
    )

if chosen_site == "직접 입력...":
    site_name = st.text_input("새로운 현장명을 입력하세요 (필수)")
else:
    site_name = chosen_site

address = st.text_input("현장 주소", key="addr_val", placeholder="현장 주소를 입력하세요")

st.divider()

# ==========================================
# 3. 작업자, 일정 및 대상 설비
# ==========================================
col1, col2 = st.columns(2)
with col1:
    author = st.text_input("작성자", value="지창현")
with col2:
    manager = st.text_input("담당 PM / 책임자", key="manager_val")

date_range = st.date_input("작업 일자 (기간 선택)", [])
date_str = ""
if len(date_range) == 2:
    start, end = date_range
    if start.month == end.month:
        date_str = f"{start.strftime('%y. %m. %d')} ~ {end.strftime('%d')}"
    else:
        date_str = f"{start.strftime('%y. %m. %d')} ~ {end.strftime('%m. %d')}"

workers = st.text_input("작업자명 및 인원", placeholder="예: 최진명 차장 외 6명")

st.markdown("#### ⚙️ 점검 대상 설비 (현장 선택 시 자동 체크)")
equipments = st.multiselect(
    "설비 목록", 
    ["STACKER CRANE", "CONVEYOR", "RGV", "LIFT"], 
    key="equip_val"
)

DEFAULT_TEXTS = {
    "STACKER CRANE": [
        "1. STACKER CRANE 점검 공통사항",
        "  1) 승강부,주행부,FORK부 구동 MOTOR 및 감속기 발열 상태 및 OIL 누유 상태 점검",
        "  2) CARRIAGE INNER ROLLER, GUIDE ROLLER 구름 상태 및 마모 상태 점검",
        "  3) FORK CHAIN TENSION 점검 및 C/F BEARING, MC GUIDE GREASE 도포",
        "  4) STC(기상반) 단자대 풀림 상태 CHECK 및 재조임",
        "  5) 주행부 구동 WHEEL,종동 WHEEL, GUIDE ROLLER 구름 상태 및 마모 상태 점검"
    ],
    "CONVEYOR": [
        "1. CONVEYOR 점검 공통사항",
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
# 4. 스마트 메모장 (자동 임시 저장)
# ==========================================
st.divider()
col_a, col_b = st.columns([7, 3])
with col_a:
    st.markdown(f"**{task_type} 내용 입력 (2번 항목부터 자동 넘버링 및 색상 분류)**")
with col_b:
    if st.button("🗑️ 메모 초기화"):
        st.session_state.contents = ""
        save_draft("")
        st.rerun()

if "contents" not in st.session_state:
    st.session_state.contents = load_draft().get("contents", "")

def update_draft():
    save_draft(st.session_state.contents)

contents = st.text_area(
    "S/C, CV, RGV, LIFT 등의 키워드가 포함되면 설비별로 자동 분류됩니다.", 
    height=150,
    key="contents",
    on_change=update_draft
)

# ==========================================
# 5. 사진 대장 업로드 (공사/점검 명칭 동적 반영)
# ==========================================
st.divider()
st.markdown(f"### 📷 현장 {task_type} 사진 대장 (선택사항)")
st.caption(f"1페이지당 4칸(좌상단, 우상단, 좌하단, 우하단) 순서대로 엑셀에 깔끔하게 배치됩니다.")

if "photo_blocks" not in st.session_state:
    st.session_state.photo_blocks = 4

photo_data = []

for i in range(0, st.session_state.photo_blocks, 2):
    cols = st.columns(2)
    for j in range(2):
        b_idx = i + j
        if b_idx < st.session_state.photo_blocks:
            with cols[j]:
                pos_label = "좌측" if (b_idx % 2 == 0) else "우측"
                row_label = "상단" if (b_idx % 4 < 2) else "하단"
                page_label = f"{b_idx // 4 + 1}페이지"
                
                with st.container(border=True):
                    st.markdown(f"**[{b_idx+1}번 칸] ({page_label} {row_label} {pos_label})**")
                    photos = st.file_uploader(f"사진 등록 (최대 2장)", type=['png', 'jpg', 'jpeg'], accept_multiple_files=True, key=f"p_{b_idx}", label_visibility="collapsed")
                    if photos:
                        p_cols = st.columns(2)
                        for p_i, p_f in enumerate(photos[:2]):
                            with p_cols[p_i]:
                                try:
                                    img_p = PILImage.open(p_f)
                                    st.image(img_p, use_container_width=True)
                                    p_f.seek(0)
                                except:
                                    st.caption("미리보기 오류")
                    desc = st.text_area(f"{task_type} 설명", key=f"d_{b_idx}", height=65, placeholder=f"{task_type} 사진 설명을 입력하세요", label_visibility="collapsed")
                    photo_data.append((b_idx, photos, desc))

if st.button("➕ 사진 칸 2개 추가"):
    st.session_state.photo_blocks += 2
    st.rerun()

# --- 병합 셀 안전 처리 함수 ---
def get_safe_cell(ws, row, col):
    cell = ws.cell(row=row, column=col)
    if type(cell).__name__ == 'MergedCell':
        for mr in ws.merged_cells.ranges:
            if mr.min_row <= row <= mr.max_row and mr.min_col <= col <= mr.max_col:
                return ws.cell(row=mr.min_row, column=mr.min_col)
    return cell

def sort_rules(text):
    is_need_replace = 1 if "교체 필요" in text else 0
    match = re.search(r'(#)?(\d+)호기', text)
    ho_number = int(match.group(2)) if match else 9999
    return (is_need_replace, ho_number)

def ensure_page_exists(ws, target_page, copied_pages):
    if target_page in copied_pages:
        return
    start_row_target = target_page * 38 + 1
    for r in range(1, 39):
        if ws.row_dimensions[r].height is not None:
            ws.row_dimensions[start_row_target + r - 1].height = ws.row_dimensions[r].height
        for c in range(1, ws.max_column + 1):
            src = ws.cell(row=r, column=c)
            tgt = ws.cell(row=start_row_target + r - 1, column=c)
            tgt.value = src.value
            if src.has_style:
                tgt.font = copy(src.font)
                tgt.border = copy(src.border)
                tgt.fill = copy(src.fill)
                tgt.alignment = copy(src.alignment)
                tgt.number_format = src.number_format
    
    from openpyxl.worksheet.cell_range import CellRange
    new_merges = []
    for mc in ws.merged_cells.ranges:
        if 1 <= mc.min_row <= 38:
            new_range = CellRange(min_col=mc.min_col, min_row=mc.min_row + target_page * 38, max_col=mc.max_col, max_row=mc.max_row + target_page * 38)
            new_merges.append(new_range)
    for nm in new_merges:
        try:
            ws.merge_cells(str(nm))
        except:
            pass
    copied_pages.add(target_page)

def write_equipment_block(ws, defaults, user_lines, row_idx, copied_pages):
    block_size = len(defaults) + len(user_lines) + 1
    current_page = (row_idx - 1) // 38
    data_end_row = current_page * 38 + 38

    if row_idx + block_size - 1 > data_end_row and block_size <= 27:
        current_page += 1
        ensure_page_exists(ws, current_page, copied_pages)
        row_idx = current_page * 38 + 12

    lines_to_write = []
    for i, df_text in enumerate(defaults):
        lines_to_write.append((df_text, i == 0, "000000", 'left'))

    for idx, text in enumerate(user_lines):
        color, bold = "000000", False
        if "교체 필요" in text:
            color, bold = "FF0000", True
        elif "조치" in text or "교체" in text:
            color, bold = "0000FF", True
        lines_to_write.append((f"{idx + 2}. {text}", bold, color, 'left'))

    for text, is_bold, color, align in lines_to_write:
        current_page = (row_idx - 1) // 38
        data_end_row = current_page * 38 + 38
        if row_idx > data_end_row:
            current_page += 1
            ensure_page_exists(ws, current_page, copied_pages)
            row_idx = current_page * 38 + 12

        cell = get_safe_cell(ws, row_idx, 2)
        cell.value = text
        sz = cell.font.size if cell.font and cell.font.size else 11
        cell.font = Font(name='굴림체', size=sz, bold=is_bold, color=color)
        cell.alignment = Alignment(horizontal=align, vertical='center')
        row_idx += 1

    row_idx += 1
    return row_idx

# ==========================================
# 6. 보고서 생성 실행
# ==========================================
st.divider()
if st.button(f"🚀 {task_type}보고서 생성하기", use_container_width=True):
    if not site_name or not date_str or not author or not manager:
        st.warning("작성자, 담당 PM, 현장명, 작업 일자는 필수입니다.")
    elif not equipments:
        st.warning("점검 진행 설비를 최소 1개 이상 선택해주세요.")
    else:
        with st.spinner("보고서를 생성하고 있습니다..."):
            try:
                site_db[site_name] = {
                    "vendor": site_db.get(site_name, {}).get("vendor", "MXRobotics"),
                    "address": address,
                    "pm": manager,
                    "equipments": equipments
                }
                save_memory(site_db)

                template_filename = f"template_mxr_{task_type}.xlsx"
                if not os.path.exists(template_filename):
                    template_filename = 'template_mxr_점검.xlsx'

                wb_report = openpyxl.load_workbook(template_filename)
                ws_report = wb_report.active

                # 헤더 정보 기입
                get_safe_cell(ws_report, 5, 3).value = site_name
                get_safe_cell(ws_report, 6, 3).value = address
                get_safe_cell(ws_report, 9, 3).value = f"{site_name} 정기 {task_type}"
                get_safe_cell(ws_report, 10, 3).value = date_str
                get_safe_cell(ws_report, 10, 9).value = workers
                get_safe_cell(ws_report, 6, 8).value = author
                get_safe_cell(ws_report, 8, 3).value = manager

                for r in range(12, 39):
                    cell = get_safe_cell(ws_report, r, 2)
                    cell.value = None
                    sz = cell.font.size if cell.font and cell.font.size else 11
                    cell.font = Font(name='굴림체', size=sz, color="000000")

                copied_pages = {0}

                raw_lines = [l.strip() for l in contents.split('\n') if l.strip()]
                sorted_lines = sorted(raw_lines, key=sort_rules)

                sc_lines, cv_lines, rgv_lines, lift_lines = [], [], [], []
                for line in sorted_lines:
                    u_line = line.upper()
                    if "RGV" in u_line:
                        rgv_lines.append(line)
                    elif "LIFT" in u_line or "리프트" in u_line:
                        lift_lines.append(line)
                    elif "CV" in u_line or "CONVEYOR" in u_line or "컨베이어" in u_line:
                        cv_lines.append(line)
                    elif "S/C" in u_line or "STC" in u_line or "크레인" in u_line or "호기" in u_line:
                        sc_lines.append(line)
                    else:
                        sc_lines.append(line)

                current_row = 12
                if "STACKER CRANE" in equipments:
                    current_row = write_equipment_block(ws_report, DEFAULT_TEXTS["STACKER CRANE"], sc_lines, current_row, copied_pages)
                if "CONVEYOR" in equipments:
                    current_row = write_equipment_block(ws_report, DEFAULT_TEXTS["CONVEYOR"], cv_lines, current_row, copied_pages)
                if "RGV" in equipments:
                    current_row = write_equipment_block(ws_report, DEFAULT_TEXTS["RGV"], rgv_lines, current_row, copied_pages)
                if "LIFT" in equipments:
                    current_row = write_equipment_block(ws_report, DEFAULT_TEXTS["LIFT"], lift_lines, current_row, copied_pages)

                output_report = io.BytesIO()
                wb_report.save(output_report)
                output_report.seek(0)

                # ==========================================
                # 사진 대장 처리 (2x2 그리드 정밀 좌표 배치)
                # ==========================================
                output_photo = None
                photo_data.sort(key=lambda x: x[0])
                has_photo_data = any(len(p) > 0 or d.strip() for _, p, d in photo_data)

                if has_photo_data and os.path.exists('photo_template.xlsx'):
                    wb_photo = openpyxl.load_workbook('photo_template.xlsx')
                    ws_photo = wb_photo.active

                    # 기존 더미 이미지 완벽 제거
                    if hasattr(ws_photo, '_images'):
                        ws_photo._images.clear()

                    PHOTO_PAGE_ROWS = 40  # 표준 A4 1페이지(4칸) 행 길이 기준 오프셋

                    # 기존 더미 텍스트 초기화
                    for p_i in range(5):
                        p_offset = p_i * PHOTO_PAGE_ROWS
                        for r_desc in [26, 45]:
                            for c_col in [2, 4]:
                                cell = get_safe_cell(ws_photo, p_offset + r_desc, c_col)
                                cell.value = None

                    # 사진 대장 헤더 명칭 동적 반영
                    get_safe_cell(ws_photo, 5, 2).value = f"{site_name} 자동화 창고 {task_type} 사진"
                    get_safe_cell(ws_photo, 10, 9).value = f"1. 현장명 : {site_name}"
                    get_safe_cell(ws_photo, 14, 9).value = f"3. 작업일자 : {date_str}"
                    get_safe_cell(ws_photo, 15, 9).value = f"4. 작업인원 : {workers}"

                    # 1페이지당 4칸 (0: 좌상, 1: 우상, 2: 좌하, 3: 우하)
                    # 0번 칸: B10 (설명: B26)
                    # 1번 칸: D10 (설명: D26)
                    # 2번 칸: B29 (설명: B45)
                    # 3번 칸: D29 (설명: D45)
                    FRAME_CONFIG = {
                        0: {"col_str": "B", "col_num": 2, "img_row": 10, "desc_row": 26},
                        1: {"col_str": "D", "col_num": 4, "img_row": 10, "desc_row": 26},
                        2: {"col_str": "B", "col_num": 2, "img_row": 29, "desc_row": 45},
                        3: {"col_str": "D", "col_num": 4, "img_row": 29, "desc_row": 45},
                    }

                    for b_idx, photos, desc in photo_data:
                        if not photos and not desc:
                            continue

                        page_num = b_idx // 4
                        pos_in_page = b_idx % 4
                        cfg = FRAME_CONFIG[pos_in_page]

                        target_img_row = (page_num * PHOTO_PAGE_ROWS) + cfg["img_row"]
                        target_desc_row = (page_num * PHOTO_PAGE_ROWS) + cfg["desc_row"]
                        target_col_str = cfg["col_str"]
                        target_col_num = cfg["col_num"]

                        # 설명 텍스트 기입
                        if desc:
                            d_cell = get_safe_cell(ws_photo, target_desc_row, target_col_num)
                            d_cell.value = desc
                            sz = d_cell.font.size if d_cell.font and d_cell.font.size else 11
                            d_cell.font = Font(name='굴림체', size=sz, color="000000")
                            d_cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

                        # 이미지 기입 (1장일 때와 2장일 때 너비 자동 배분)
                        num_imgs = min(len(photos), 2)
                        for img_i, p_file in enumerate(photos[:2]):
                            p_file.seek(0)
                            img_pil = PILImage.open(p_file)

                            # 2장일 때는 가로로 나란히 들어가도록 폭을 165px로 축소
                            target_w = 340 if num_imgs == 1 else 165
                            target_h = 280
                            img_pil.thumbnail((target_w, target_h))

                            img_byte_arr = io.BytesIO()
                            img_pil.save(img_byte_arr, format='PNG')
                            img_byte_arr.seek(0)
                            xl_img = Image(img_byte_arr)

                            # 2장 중 2번째 장은 같은 칸 내에서 살짝 우측에 배치되도록 오프셋 적용
                            if num_imgs == 2 and img_i == 1:
                                xl_img.left = 135  # 이미지 내부 오프셋(pt)
                            
                            ws_photo.add_image(xl_img, f"{target_col_str}{target_img_row}")

                    output_photo = io.BytesIO()
                    wb_photo.save(output_photo)
                    output_photo.seek(0)

                st.success(f"🎉 {task_type}보고서 작성이 완료되었습니다!")
                c1, c2 = st.columns(2)
                with c1:
                    st.download_button(
                        label=f"📥 [{task_type} 보고서] 다운로드",
                        data=output_report,
                        file_name=f"(MXR_{task_type}보고서){site_name}_{date_str[:8].replace('. ', '')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                if output_photo:
                    with c2:
                        st.download_button(
                            label=f"🖼️ [{task_type} 사진] 다운로드",
                            data=output_photo,
                            file_name=f"(MXR_{task_type}사진){site_name}_{date_str[:8].replace('. ', '')}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True
                        )

            except Exception as e:
                st.error(f"오류가 발생했습니다: {e}")
