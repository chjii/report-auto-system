import streamlit as st
import openpyxl
from openpyxl.styles import Font, Alignment
from openpyxl.drawing.image import Image
from openpyxl.drawing.spreadsheet_drawing import OneCellAnchor, AnchorMarker
from openpyxl.drawing.xdr import XDRPositiveSize2D
from PIL import Image as PILImage
import io
import re
import json
import os
from copy import copy

st.set_page_config(page_title="자동창고 보고서 생성기", layout="wide")
st.title("📝 현장 보고서 자동 생성기")

# ==========================================
# 0. 마스터 DB 및 기본 설정
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

SC_PARTS_CONFIG = {
    "CARRIAGE부": ["1) FORK C.F BEARING CLEANING", "2) FORK C.F BEARING GREASE 도포", "3) FORK CHAIN TENSION 확인", ""],
    "주행부": ["1) 주행 WHEEL 상태 확인", "2) 주행 GUIDE ROLLER 상태 확인", "3) 주행 MOTOR BRAKE GAP 확인", ""],
    "승강부": ["1) 외측 GUIDE ROLLER 상태 확인", "2) 내측 GUIDE ROLLER 상태 확인", "3) 승강 MOTOR BRAKE GAP 확인", ""],
    "상부부": ["1) 상부 GUIDE ROLLER 상태 확인", "2) 상부 BRAKE ROLLER 상태 확인", "", ""],
    "집전부": ["1) 집전기 ROLLER 및 SHOE 상태 확인", "2) 집전기 CLEANING", "3) 기상반 PANEL 단자대 REBOLTING", "4) 각 SENSOR 단자대 REBOLTING"]
}

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
# 1. 작업 분류 선택
# ==========================================
st.markdown("### 📋 작업 분류 선택")
task_type = st.radio("보고서 종류", ["점검", "공사"], horizontal=True)

st.divider()

# ==========================================
# 2. 현장 선택 및 자동 세팅
# ==========================================
st.markdown("### 🏢 현장 선택 및 자동 세팅")
site_options = list(site_db.keys()) + ["직접 입력..."]

if "site_dropdown" not in st.session_state:
    st.session_state.site_dropdown = site_options[0]
    st.session_state.addr_val = site_db[site_options[0]].get("address", "")
    st.session_state.manager_val = site_db[site_options[0]].get("pm", "")
    st.session_state.equip_val = site_db[site_options[0]].get("equipments", ["STACKER CRANE"])

def on_site_change():
    chosen = st.session_state.site_dropdown
    if chosen in site_db:
        st.session_state.addr_val = site_db[chosen].get("address", "")
        st.session_state.manager_val = site_db[chosen].get("pm", "")
        st.session_state.equip_val = site_db[chosen].get("equipments", ["STACKER CRANE"])
    elif chosen == "직접 입력...":
        st.session_state.addr_val = ""
        st.session_state.manager_val = ""
        st.session_state.equip_val = ["STACKER CRANE"]

col_s1, col_s2 = st.columns([1, 2])
with col_s1:
    chosen_site = st.selectbox("현장명 선택 (타이핑 검색 가능)", site_options, key="site_dropdown", on_change=on_site_change)

with col_s2:
    if chosen_site == "직접 입력...":
        site_name = st.text_input("새로운 현장명 입력 (필수)")
    else:
        site_name = chosen_site
    address = st.text_input("현장 주소", key="addr_val")

col1, col2, col3 = st.columns(3)
with col1:
    author = st.text_input("작성자", value="지창현")
with col2:
    manager = st.text_input("담당 PM / 책임자", key="manager_val")
with col3:
    date_range = st.date_input("작업 일자", [])

date_str = ""
if len(date_range) == 2:
    start, end = date_range
    if start.month == end.month:
        date_str = f"{start.strftime('%y. %m. %d')} ~ {end.strftime('%d')}"
    else:
        date_str = f"{start.strftime('%y. %m. %d')} ~ {end.strftime('%m. %d')}"

workers = st.text_input("작업자명 및 인원", placeholder="예: 최진명 차장 외 6명")

equipments = []
if task_type == "점검":
    st.markdown("#### ⚙️ 점검 대상 설비")
    equipments = st.multiselect("설비 목록", ["STACKER CRANE", "CONVEYOR", "RGV", "LIFT"], key="equip_val")

# ==========================================
# 3. 작업 내용 메모장 (점검 모드 전용)
# ==========================================
if task_type == "점검":
    st.divider()
    col_m1, col_m2 = st.columns([8, 2])
    with col_m1:
        st.markdown(f"**점검 상세 내용 (2번부터 자동 넘버링 및 색상 분류)**")
    with col_m2:
        if st.button("🗑️ 메모 초기화"):
            st.session_state.contents = ""
            save_draft("")
            st.rerun()

    if "contents" not in st.session_state:
        st.session_state.contents = load_draft().get("contents", "")

    def update_draft():
        save_draft(st.session_state.contents)

    contents = st.text_area("S/C, CV, RGV, LIFT 키워드가 포함되면 설비별로 자동 분류됩니다.", height=130, key="contents", on_change=update_draft)
else:
    contents = ""

# ==========================================
# 4. 현장 사진 대장 (공사 / 점검)
# ==========================================
st.divider()
st.markdown(f"### 📷 {task_type} 사진 대장")

photo_upload_data = []

if task_type == "공사":
    st.caption("각 번호별 [공사작업 내역]과 좌/우 작업사진 및 설명을 입력하세요. 원본 사진은 절대 잘리지 않으며 가로세로 비율을 유지한 채 틀 안에 맞춤 삽입됩니다.")
    
    if "const_items_count" not in st.session_state:
        st.session_state.const_items_count = 2

    for c_idx in range(st.session_state.const_items_count):
        with st.container(border=True):
            st.markdown(f"#### 🔨 NO {c_idx + 1}")
            task_title = st.text_input(
                f"공사작업 내역 #{c_idx + 1}", 
                placeholder="예: 공사 자재 확인, 기존 DOORLOCK 철거 등",
                key=f"const_title_{c_idx}"
            )
            
            p_col1, p_col2 = st.columns(2)
            with p_col1:
                st.markdown("**[좌측 사진 (작업 전)]**")
                p_left = st.file_uploader(f"좌측 사진 등록", type=['png', 'jpg', 'jpeg'], key=f"c_img_l_{c_idx}")
                if p_left:
                    try:
                        im_l = PILImage.open(p_left)
                        st.image(im_l, use_container_width=True)
                        p_left.seek(0)
                    except:
                        pass
                d_left = st.text_input(f"좌측 사진 설명", placeholder="예: 작업 전 상태", key=f"c_desc_l_{c_idx}")

            with p_col2:
                st.markdown("**[우측 사진 (작업 후)]**")
                p_right = st.file_uploader(f"우측 사진 등록", type=['png', 'jpg', 'jpeg'], key=f"c_img_r_{c_idx}")
                if p_right:
                    try:
                        im_r = PILImage.open(p_right)
                        st.image(im_r, use_container_width=True)
                        p_right.seek(0)
                    except:
                        pass
                d_right = st.text_input(f"우측 사진 설명", placeholder="예: 작업 완료 후 상태", key=f"c_desc_r_{c_idx}")

            photo_upload_data.append({
                "no": c_idx + 1,
                "title": task_title,
                "p_left": p_left,
                "d_left": d_left,
                "p_right": p_right,
                "d_right": d_right
            })

    col_btn1, col_btn2 = st.columns([2, 8])
    with col_btn1:
        if st.button("➕ 공사 작업 항목 추가 (NO 추가)"):
            st.session_state.const_items_count += 1
            st.rerun()
    with col_btn2:
        if st.session_state.const_items_count > 1:
            if st.button("➖ 마지막 항목 제거"):
                st.session_state.const_items_count -= 1
                st.rerun()

else:
    # 점검 모드: STACKER CRANE 부위별 표준 탭
    if "STACKER CRANE" in equipments:
        st.info("💡 STACKER CRANE 부위별 표준 점검 항목입니다. 사진 업로드 시 해당 항목 텍스트와 함께 엑셀에 들어갑니다.")
        tabs = st.tabs(list(SC_PARTS_CONFIG.keys()))
        for t_idx, (part_name, default_items) in enumerate(SC_PARTS_CONFIG.items()):
            with tabs[t_idx]:
                p_cols = st.columns(2)
                for item_idx, def_val in enumerate(default_items):
                    target_col = p_cols[item_idx % 2]
                    with target_col:
                        with st.container(border=True):
                            custom_desc = st.text_input(
                                f"항목명 #{item_idx+1}", 
                                value=def_val, 
                                key=f"sc_desc_{t_idx}_{item_idx}", 
                                placeholder="점검 내용 직접 입력"
                            )
                            photos = st.file_uploader(
                                f"사진 등록 (최대 2장)", 
                                type=['png', 'jpg', 'jpeg'], 
                                accept_multiple_files=True, 
                                key=f"sc_photo_{t_idx}_{item_idx}"
                            )
                            if photos:
                                preview_cols = st.columns(2)
                                for prv_i, prv_f in enumerate(photos[:2]):
                                    with preview_cols[prv_i]:
                                        try:
                                            im = PILImage.open(prv_f)
                                            st.image(im, use_container_width=True)
                                            prv_f.seek(0)
                                        except:
                                            pass
                            photo_upload_data.append({"photos": photos, "desc": custom_desc})

# --- 엑셀 안전 처리 함수 ---
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
# 💡 비율 유지 무자름 스케일링 앵커 함수
# ==========================================
def add_scaled_photo(ws, file_obj, col_idx, row_idx, max_w_px, max_h_px, offset_x_emu=180000, offset_y_emu=180000):
    file_obj.seek(0)
    pil_img = PILImage.open(file_obj)
    orig_w, orig_h = pil_img.size

    # 원본 비율을 100% 유지하면서 지정 박스 안에 꽉 차도록 계산 (자르지 않음)
    ratio = min(max_w_px / orig_w, max_h_px / orig_h)
    target_w = int(orig_w * ratio)
    target_h = int(orig_h * ratio)

    # 중앙 정렬을 위한 추가 오프셋 계산 (남는 여백을 균등 배분)
    extra_pad_x = int((max_w_px - target_w) / 2 * 9525)
    extra_pad_y = int((max_h_px - target_h) / 2 * 9525)

    b_arr = io.BytesIO()
    pil_img.save(b_arr, format='PNG')
    b_arr.seek(0)
    xl_img = Image(b_arr)

    # 1px = 9525 EMU
    marker = AnchorMarker(
        col=col_idx, 
        colOff=offset_x_emu + extra_pad_x, 
        row=row_idx, 
        rowOff=offset_y_emu + extra_pad_y
    )
    size = XDRPositiveSize2D(int(target_w * 9525), int(target_h * 9525))
    xl_img.anchor = OneCellAnchor(_from=marker, ext=size)
    ws.add_image(xl_img)

# ==========================================
# 5. 생성 및 다운로드 실행
# ==========================================
st.divider()
btn_label = f"🚀 [공사 사진대장] 생성하기" if task_type == "공사" else f"🚀 [점검 보고서 & 사진대장] 생성하기"

if st.button(btn_label, use_container_width=True):
    if not site_name or not date_str or not author or not manager:
        st.warning("작성자, 담당 PM, 현장명, 작업 일자를 확인해주세요.")
    elif task_type == "점검" and not equipments:
        st.warning("점검 진행 설비를 최소 1개 이상 선택해주세요.")
    else:
        with st.spinner("엑셀 파일을 생성 중입니다..."):
            try:
                site_db[site_name] = {
                    "vendor": site_db.get(site_name, {}).get("vendor", "MXRobotics"),
                    "address": address,
                    "pm": manager,
                    "equipments": equipments if task_type == "점검" else site_db.get(site_name, {}).get("equipments", ["STACKER CRANE"])
                }
                save_memory(site_db)

                output_report = None
                output_photo = None

                # --------------------------------------------------
                # A. 점검 보고서 생성 (점검 모드일 때만 실행)
                # --------------------------------------------------
                if task_type == "점검":
                    template_filename = f"template_mxr_점검.xlsx"
                    wb_report = openpyxl.load_workbook(template_filename)
                    ws_report = wb_report.active

                    get_safe_cell(ws_report, 5, 3).value = site_name
                    get_safe_cell(ws_report, 6, 3).value = address
                    get_safe_cell(ws_report, 9, 3).value = f"{site_name} 정기 점검"
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
                        if "RGV" in u_line: rgv_lines.append(line)
                        elif "LIFT" in u_line or "리프트" in u_line: lift_lines.append(line)
                        elif "CV" in u_line or "CONVEYOR" in u_line or "컨베이어" in u_line: cv_lines.append(line)
                        elif "S/C" in u_line or "STC" in u_line or "크레인" in u_line or "호기" in u_line: sc_lines.append(line)
                        else: sc_lines.append(line)

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

                # --------------------------------------------------
                # B. 사진 대장 생성 (무자름·비율유지 스케일링 적용)
                # --------------------------------------------------
                if os.path.exists('photo_template.xlsx'):
                    wb_photo = openpyxl.load_workbook('photo_template.xlsx')
                    ws_photo = wb_photo.active

                    if hasattr(ws_photo, '_images'):
                        ws_photo._images.clear()

                    get_safe_cell(ws_photo, 5, 2).value = f"{site_name} 자동화 창고 {task_type} 사진"
                    get_safe_cell(ws_photo, 10, 9).value = f"1. 현장명 : {site_name}"
                    get_safe_cell(ws_photo, 14, 9).value = f"3. 작업일자 : {date_str}"
                    get_safe_cell(ws_photo, 15, 9).value = f"4. 작업인원 : {workers}"

                    if task_type == "공사":
                        ITEM_HEIGHT = 19
                        num_items = len(photo_upload_data)
                        EMU_5MM = 180000   # 5mm 여백 (180,000 EMU)
                        EMU_2_5MM = 90000  # 2.5mm 중앙 여백 (90,000 EMU)

                        # D27:H36 박스의 유효 픽셀 크기 (여백 5mm 제외 기준)
                        BOX_FULL_W_PX = 320
                        BOX_HALF_W_PX = 155
                        BOX_H_PX = 200

                        for idx, item in enumerate(photo_upload_data):
                            base_r = 10 + (idx * ITEM_HEIGHT)
                            desc_r = base_r + 16

                            # 1) NO 번호
                            no_cell = get_safe_cell(ws_photo, base_r, 1)
                            no_cell.value = str(item["no"])
                            no_cell.font = Font(name='굴림체', size=11, bold=True)
                            no_cell.alignment = Alignment(horizontal='center', vertical='center')

                            # 2) 공사작업 내역 (두 번째 열)
                            title_cell = get_safe_cell(ws_photo, base_r, 2)
                            title_cell.value = item["title"]
                            title_cell.font = Font(name='굴림체', size=10, bold=True)
                            title_cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

                            # 3) 설명 텍스트 기입
                            d_left_cell = get_safe_cell(ws_photo, desc_r, 2)
                            d_left_cell.value = item["d_left"]
                            d_left_cell.font = Font(name='굴림체', size=9)
                            d_left_cell.alignment = Alignment(horizontal='center', vertical='center')

                            d_right_cell = get_safe_cell(ws_photo, desc_r, 4)
                            d_right_cell.value = item["d_right"]
                            d_right_cell.font = Font(name='굴림체', size=9)
                            d_right_cell.alignment = Alignment(horizontal='center', vertical='center')

                            has_left = item["p_left"] is not None
                            has_right = item["p_right"] is not None
                            row_start_idx = 26 + (idx * ITEM_HEIGHT) # row 27 -> index 26

                            # 케이스 A: 1장만 등록되었을 때 (D27:H36 전체 영역에 원본 비율 유지 채움)
                            if (has_left and not has_right) or (has_right and not has_left):
                                single_file = item["p_left"] if has_left else item["p_right"]
                                add_scaled_photo(
                                    ws_photo, 
                                    single_file, 
                                    col_idx=3, # D열
                                    row_idx=row_start_idx, 
                                    max_w_px=BOX_FULL_W_PX, 
                                    max_h_px=BOX_H_PX,
                                    offset_x_emu=EMU_5MM,
                                    offset_y_emu=EMU_5MM
                                )

                            # 케이스 B: 2장 모두 등록되었을 때 (좌/우 5:5 비율 유지 채움)
                            elif has_left and has_right:
                                # 좌측 사진 (D열 시작, 우측 중앙여백 2.5mm 적용)
                                add_scaled_photo(
                                    ws_photo, 
                                    item["p_left"], 
                                    col_idx=3, # D열
                                    row_idx=row_start_idx, 
                                    max_w_px=BOX_HALF_W_PX, 
                                    max_h_px=BOX_H_PX,
                                    offset_x_emu=EMU_5MM,
                                    offset_y_emu=EMU_5MM
                                )

                                # 우측 사진 (F열 시작, 좌측 중앙여백 2.5mm 적용)
                                add_scaled_photo(
                                    ws_photo, 
                                    item["p_right"], 
                                    col_idx=5, # F열
                                    row_idx=row_start_idx, 
                                    max_w_px=BOX_HALF_W_PX, 
                                    max_h_px=BOX_H_PX,
                                    offset_x_emu=EMU_2_5MM,
                                    offset_y_emu=EMU_5MM
                                )

                        # 미사용 하단 템플릿 행 완전 삭제
                        last_used_row = 10 + (num_items * ITEM_HEIGHT)
                        max_clean_row = max(ws_photo.max_row, last_used_row + 100)
                        for r_del in range(last_used_row, max_clean_row + 1):
                            for c_del in range(1, ws_photo.max_column + 1):
                                cell = ws_photo.cell(row=r_del, column=c_del)
                                if type(cell).__name__ != 'MergedCell':
                                    cell.value = None

                    else:
                        # 점검 모드: 2x2 사진대장 기입
                        PHOTO_PAGE_ROWS = 85
                        for p_i in range(5):
                            p_offset = p_i * PHOTO_PAGE_ROWS
                            for b in range(4):
                                desc_r = p_offset + 10 + (b * 19) + 16
                                cell = get_safe_cell(ws_photo, desc_r, 2)
                                cell.value = None

                        valid_photos = [x for x in photo_upload_data if (x.get("photos") and len(x["photos"]) > 0) or (x.get("desc") and x["desc"].strip())]

                        for b_idx, item in enumerate(valid_photos):
                            page_num = b_idx // 4
                            pos_in_page = b_idx % 4
                            base_row = (page_num * PHOTO_PAGE_ROWS) + 10 + (pos_in_page * 19)
                            desc_row = base_row + 16

                            if item.get("desc"):
                                d_cell = get_safe_cell(ws_photo, desc_row, 2)
                                d_cell.value = item["desc"]
                                d_cell.font = Font(name='굴림체', size=10, bold=True)
                                d_cell.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)

                            if item.get("photos"):
                                for img_i, p_file in enumerate(item["photos"][:2]):
                                    p_file.seek(0)
                                    col = 'B' if img_i == 0 else 'D'
                                    add_scaled_photo(
                                        ws_photo,
                                        p_file,
                                        col_idx=(1 if img_i == 0 else 3),
                                        row_idx=base_row - 1,
                                        max_w_px=320,
                                        max_h_px=220
                                    )

                    output_photo = io.BytesIO()
                    wb_photo.save(output_photo)
                    output_photo.seek(0)

                st.success(f"🎉 작성이 완료되었습니다!")

                # 공사 모드: [공사 사진대장] 1개만 다운로드
                if task_type == "공사":
                    st.download_button(
                        label=f"📥 [MXR_공사사진대장] 다운로드",
                        data=output_photo,
                        file_name=f"(MXR_공사사진대장){site_name}_{date_str[:8].replace('. ', '')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                else:
                    c1, c2 = st.columns(2)
                    with c1:
                        st.download_button(
                            label=f"📥 [MXR_점검보고서] 다운로드",
                            data=output_report,
                            file_name=f"(MXR_점검보고서){site_name}_{date_str[:8].replace('. ', '')}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True
                        )
                    if output_photo:
                        with c2:
                            st.download_button(
                                label=f"🖼️ [MXR_점검사진대장] 다운로드",
                                data=output_photo,
                                file_name=f"(MXR_점검사진대장){site_name}_{date_str[:8].replace('. ', '')}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                use_container_width=True
                            )

            except Exception as e:
                st.error(f"오류가 발생했습니다: {e}")
