import datetime
import io
import math
import openpyxl
import pandas as pd
import streamlit as st

# ==================== 1. ตั้งค่าหน้าจอโปรแกรม ====================
st.set_page_config(
    page_title="ระบบจัดตารางสอบ - มทร.ตะวันออก วิทยาเขตจันทบุรี",
    page_icon="🏫",
    layout="wide",
)

USER_CREDENTIALS = {"monthira": "123456", "registry_staff": "rmutto456"}

# Session States
if "logged_in" not in st.session_state:
    st.session_state["logged_in"] = False
if "username" not in st.session_state:
    st.session_state["username"] = ""

# --- เพิ่มตัวแปร Session State สำหรับจดจำข้อมูลป้องกันการจัดชนกัน ---
if "history_schedule" not in st.session_state:
    st.session_state["history_schedule"] = pd.DataFrame()
if "show_confirm_clear" not in st.session_state:
    st.session_state["show_confirm_clear"] = False

def init_memory():
    st.session_state["occ_students"] = {}
    st.session_state["occ_daily_count"] = {}
    st.session_state["occ_heavy"] = {}
    st.session_state["occ_rooms"] = {}
    st.session_state["occ_invigs"] = {}
    st.session_state["occ_workload"] = {}
    st.session_state["history_schedule"] = pd.DataFrame()
    st.session_state["show_confirm_clear"] = False

if "occ_rooms" not in st.session_state:
    init_memory()

# 1. ฐานข้อมูลห้องสอบ (เริ่มต้น)
if "df_rooms" not in st.session_state:
    st.session_state["df_rooms"] = pd.DataFrame([
        {"อาคาร": "อาคารเรียนรวม 36 ปี", "รหัสห้อง": "36-301", "ความจุสอบ": 50, "ประเภท": "ห้องทฤษฎี", "สถานะ": "ใช้งานได้"},
        {"อาคาร": "อาคารเรียนรวม 36 ปี", "รหัสห้อง": "ห้องประชุมใหญ่ 36 ปี", "ความจุสอบ": 250, "ประเภท": "ห้องทฤษฎี", "สถานะ": "ใช้งานได้"},
        {"อาคาร": "อาคารเรียนรวม 36 ปี", "รหัสห้อง": "หอประชุมจันทบุรี", "ความจุสอบ": 500, "ประเภท": "ห้องทฤษฎี", "สถานะ": "ใช้งานได้"},
        {"อาคาร": "อาคารปฏิบัติการไอที", "รหัสห้อง": "IT-LAB1", "ความจุสอบ": 80, "ประเภท": "ห้องปฏิบัติการคอมพิวเตอร์", "สถานะ": "ใช้งานได้"},
    ])

# 2. ฐานข้อมูลบุคลากรคุมสอบสำรอง (อัปเดต คณะวิศวกรรมศาสตร์)
if "df_staff_pool" not in st.session_state:
    st.session_state["df_staff_pool"] = pd.DataFrame([
        {"คณะ": "คณะเทคโนโลยีอุตสาหกรรมการเกษตร", "ชื่อ-นามสกุล": "ดร.สมเกียรติ มั่นคง", "ตำแหน่ง": "อาจารย์", "ประเภท": "อาจารย์ในคณะ"},
        {"คณะ": "คณะเทคโนโลยีอุตสาหกรรมการเกษตร", "ชื่อ-นามสกุล": "นายวิชัย สำรองดี", "ตำแหน่ง": "เจ้าหน้าที่บริหารงานทั่วไป", "ประเภท": "เจ้าหน้าที่สำรองส่วนกลาง"},
        {"คณะ": "คณะเทคโนโลยีสังคม", "ชื่อ-นามสกุล": "อ.ประเสริฐ นามดี", "ตำแหน่ง": "อาจารย์", "ประเภท": "อาจารย์ในคณะ"},
        {"คณะ": "คณะเทคโนโลยีสังคม", "ชื่อ-นามสกุล": "นางสาวนภา ใจเย็น", "ตำแหน่ง": "นักวิชาการศึกษา", "ประเภท": "เจ้าหน้าที่สำรองส่วนกลาง"},
        {"คณะ": "คณะวิศวกรรมศาสตร์", "ชื่อ-นามสกุล": "ดร.กิตติศักดิ์ มณี", "ตำแหน่ง": "อาจารย์", "ประเภท": "อาจารย์ในคณะ"},
        {"คณะ": "คณะวิศวกรรมศาสตร์", "ชื่อ-นามสกุล": "นายสมพร งามดี", "ตำแหน่ง": "นักวิชาการศึกษา", "ประเภท": "เจ้าหน้าที่สำรองส่วนกลาง"},
    ])

def logout():
    st.session_state["logged_in"] = False
    st.session_state["username"] = ""
    st.rerun()

def get_column_value(row, possible_names, default_val=""):
    for name in possible_names:
        if name in row.index and pd.notna(row[name]):
            return row[name]
    return default_val

def parse_int_safe(val, default=0):
    try:
        if pd.notna(val) and str(val).strip() != "":
            return int(float(str(val).strip()))
    except (ValueError, TypeError):
        pass
    return default

def create_test_subject_excel():
    test_data = {
        "คณะ": ["หมวดวิชาศึกษาทั่วไป (GE)", "หมวดวิชาศึกษาทั่วไป (GE)", "หมวดวิชาศึกษาทั่วไป (GE)", "คณะเทคโนโลยีอุตสาหกรรมการเกษตร", "คณะเทคโนโลยีสังคม", "คณะวิศวกรรมศาสตร์"],
        "รหัสวิชา": ["GE-101", "GE-101", "GE-101", "05-300-201", "02-303-101", "03-101-102"],
        "ชื่อวิชา": ["การพัฒนาคุณภาพชีวิต", "การพัฒนาคุณภาพชีวิต", "การพัฒนาคุณภาพชีวิต", "แคลคูลัสสำหรับวิศวกร", "การเขียนโปรแกรมคอมพิวเตอร์", "วิศวกรรมวัสดุเบื้องต้น"],
        "กลุ่มเรียน": ["GE-01", "GE-02", "GE-03", "AG-101", "IT-101", "ENG-101"],
        "จำนวนผู้เข้าสอบ": [60, 50, 40, 35, 45, 30],
        "ชื่อผู้สอน": ["ดร.อนันต์ เรียนดี", "ดร.อนันต์ เรียนดี", "อ.วิภา นามสว่าง", "ดร.วิชัย คำนวณตรง", "ดร.สมชาย ใจดี", "ดร.ศิริพร อัญมณี"],
        "สังกัดสาขา": ["ศึกษาทั่วไป", "ศึกษาทั่วไป", "ศึกษาทั่วไป", "วิศวกรรมเกษตร", "เทคโนโลยีสารสนเทศ", "วิศวกรรมวัสดุ"],
        "ประเภทการสอบ": ["ทฤษฎี", "ทฤษฎี", "ทฤษฎี", "ทฤษฎี", "ปฏิบัติคอมพิวเตอร์", "ทฤษฎี"],
        "ชั่วโมงสอบ_M": [2.0, 2.0, 2.0, 2.0, 2.0, 2.0],
        "ชั่วโมงสอบ_F": [3.0, 3.0, 3.0, 3.0, 3.0, 3.0],
        "วิชาคำนวณ": ["NO", "NO", "NO", "YES", "NO", "NO"],
    }
    df = pd.DataFrame(test_data)
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Exam_Subjects")
    return buffer.getvalue()

def generate_time_slots(start_date, end_date, daily_slots):
    slots = []
    current_date = start_date
    while current_date <= end_date:
        if current_date.weekday() < 5:
            date_str = current_date.strftime("%d/%m/%Y")
            for slot_name in daily_slots:
                slots.append({"full_slot": f"{date_str} ({slot_name})", "date_str": date_str, "time_str": slot_name})
        current_date += datetime.timedelta(days=1)
    return slots

# ==================== 2. อัลกอริทึมจัดตารางสอบ ====================
def auto_schedule_exams_combined(df_subjects, df_rooms, df_staff_pool, slots_m, slots_f, faculty_rule_map=None, proctor_mode="จัดในคณะก่อน หากไม่พอค่อยข้ามคณะ (Priority)"):
    if faculty_rule_map is None:
        faculty_rule_map = {"คณะเทคโนโลยีอุตสาหกรรมการเกษตร": 2, "คณะเทคโนโลยีสังคม": 2, "คณะวิศวกรรมศาสตร์": 2}

    student_group_occupancy = st.session_state["occ_students"]
    student_group_daily_count = st.session_state["occ_daily_count"]
    student_group_daily_heavy = st.session_state["occ_heavy"]
    room_occupancy = st.session_state["occ_rooms"]
    invigilator_occupancy = st.session_state["occ_invigs"]
    invigilator_workload = st.session_state["occ_workload"]

    unassigned_warnings = []
    
    grouped_tasks = []
    for (code, name), group in df_subjects.groupby([
        df_subjects.apply(lambda r: get_column_value(r, ["รหัสวิชา"], ""), axis=1),
        df_subjects.apply(lambda r: get_column_value(r, ["ชื่อวิชา"], ""), axis=1),
    ]):
        grouped_tasks.append({"rows": group.to_dict("records")})

    results = []

    def assign_invigilators(faculty_name, instructor, total_students, slot_str):
        min_rule = faculty_rule_map.get(faculty_name, 2)
        used_invs = invigilator_occupancy.get(slot_str, set())
        assigned = []
        calc_req = math.ceil(total_students / 40)
        req_count = max(min_rule, calc_req)

        # จัดกลุ่มบุคลากรแยกตามคณะ
        fac_staff_df = df_staff_pool[df_staff_pool["คณะ"] == faculty_name]
        other_staff_df = df_staff_pool[df_staff_pool["คณะ"] != faculty_name]

        same_fac_teachers = list(fac_staff_df[fac_staff_df["ประเภท"] == "อาจารย์ในคณะ"]["ชื่อ-นามสกุล"].unique())
        same_fac_backup = list(fac_staff_df[fac_staff_df["ประเภท"] == "เจ้าหน้าที่สำรองส่วนกลาง"]["ชื่อ-นามสกุล"].unique())

        other_fac_teachers = list(other_staff_df[other_staff_df["ประเภท"] == "อาจารย์ในคณะ"]["ชื่อ-นามสกุล"].unique())
        other_fac_backup = list(other_staff_df[other_staff_df["ประเภท"] == "เจ้าหน้าที่สำรองส่วนกลาง"]["ชื่อ-นามสกุล"].unique())

        # เพิ่มอาจารย์ผู้สอนวิชานั้นก่อนเสมอถ้าว่าง
        if instructor not in used_invs:
            assigned.append(instructor)

        # ลำดับการเลือกตามเงื่อนไข (Proctor Mode)
        if proctor_mode == "คณะตนเองเท่านั้น (Strict)":
            pool = same_fac_teachers + same_fac_backup
        elif proctor_mode == "ข้ามคณะได้อิสระ (Any Faculty)":
            pool = same_fac_teachers + other_fac_teachers + same_fac_backup + other_fac_backup
        else: # จัดในคณะก่อน หากไม่พอค่อยข้ามคณะ (Priority)
            pool = same_fac_teachers + same_fac_backup + other_fac_teachers + other_fac_backup

        pool = list(dict.fromkeys(pool)) # ลบรายชื่อซ้ำ
        pool.sort(key=lambda t: invigilator_workload.get(t, 0)) # เรียงตามภาระงานคุมสอบเดิม

        for t in pool:
            if len(assigned) >= req_count: break
            if t not in assigned and t not in used_invs:
                assigned.append(t)

        return assigned

    for task in grouped_tasks:
        rows = task["rows"]
        groups_list = [str(get_column_value(pd.Series(r), ["กลุ่มเรียน", "กลุ่ม"], "")).strip() for r in rows]
        total_students = sum(parse_int_safe(get_column_value(pd.Series(r), ["จำนวนผู้เข้าสอบ", "จำนวนนักศึกษา"], 0)) for r in rows)
        first_r = pd.Series(rows[0])
        subj_code = get_column_value(first_r, ["รหัสวิชา"], "SUBJ")
        subj_name = get_column_value(first_r, ["ชื่อวิชา"], "วิชาไม่ระบุ")
        faculty_name = get_column_value(first_r, ["คณะ"], "หมวดวิชาศึกษาทั่วไป (GE)")
        exam_type = str(get_column_value(first_r, ["ประเภทการสอบ"], "ทฤษฎี")).strip()
        is_heavy = str(get_column_value(first_r, ["วิชาคำนวณ"], "NO")).strip().upper() in ["YES", "Y", "TRUE", "1", "คำนวณ"]
        instructor = get_column_value(first_r, ["ชื่อผู้สอน"], "อาจารย์ผู้สอน")

        active_rooms = df_rooms[df_rooms["สถานะ"] == "ใช้งานได้"]
        target_type = "ห้องปฏิบัติการคอมพิวเตอร์" if ("ปฏิบัติ" in exam_type or "คอม" in exam_type) else "ห้องทฤษฎี"
        valid_rooms = active_rooms[(active_rooms["ประเภท"] == target_type) & (active_rooms["ความจุสอบ"] >= total_students)].sort_values(by="ความจุสอบ")

        def find_best_slot(slots_list, hrs_val):
            if hrs_val <= 0: return None, None, None
            if valid_rooms.empty: return None, None, "NO_ROOM"

            for slot_obj in slots_list:
                slot_str = slot_obj["full_slot"]
                date_str = slot_obj["date_str"]

                if any(g in student_group_occupancy.get(slot_str, set()) for g in groups_list): continue
                if any(student_group_daily_count.get((g, date_str), 0) >= 2 for g in groups_list): continue
                if is_heavy and any(student_group_daily_heavy.get((g, date_str), False) for g in groups_list): continue

                invig_list = assign_invigilators(faculty_name, instructor, total_students, slot_str)
                used_rms = room_occupancy.get(slot_str, set())
                avail_rm = None
                for _, rm in valid_rooms.iterrows():
                    if rm["รหัสห้อง"] not in used_rms:
                        avail_rm = rm["รหัสห้อง"]
                        break

                if avail_rm:
                    for g in groups_list:
                        student_group_occupancy.setdefault(slot_str, set()).add(g)
                        student_group_daily_count[(g, date_str)] = student_group_daily_count.get((g, date_str), 0) + 1
                        if is_heavy: student_group_daily_heavy[(g, date_str)] = True

                    room_occupancy.setdefault(slot_str, set()).add(avail_rm)
                    for inv in invig_list:
                        invigilator_occupancy.setdefault(slot_str, set()).add(inv)
                        invigilator_workload[inv] = invigilator_workload.get(inv, 0) + 1

                    return slot_str, avail_rm, ", ".join(invig_list)
            return None, None, "NO_SLOT"

        hrs_m = float(get_column_value(first_r, ["ชั่วโมงสอบ_M"], 2.0))
        m_slot, m_rm, m_inv = find_best_slot(slots_m, hrs_m)
        hrs_f = float(get_column_value(first_r, ["ชั่วโมงสอบ_F"], 3.0))
        f_slot, f_rm, f_inv = find_best_slot(slots_f, hrs_f)

        if not m_slot and hrs_m > 0:
            unassigned_warnings.append(f"❌ **[{subj_code}] {subj_name} [กลางภาค]**: จัดไม่ได้ (ห้องไม่พอ/เวลาชน)")
        if not f_slot and hrs_f > 0:
            unassigned_warnings.append(f"❌ **[{subj_code}] {subj_name} [ปลายภาค]**: จัดไม่ได้ (ห้องไม่พอ/เวลาชน)")

        for r in rows:
            r_ser = pd.Series(r)
            results.append({
                "คณะ": get_column_value(r_ser, ["คณะ"], faculty_name),
                "รหัสวิชา": subj_code,
                "ชื่อวิชา": subj_name,
                "กลุ่มเรียน": get_column_value(r_ser, ["กลุ่มเรียน"], ""),
                "จำนวนผู้เข้าสอบ": parse_int_safe(get_column_value(r_ser, ["จำนวนผู้เข้าสอบ"], 0)),
                "ชื่อผู้สอน": get_column_value(r_ser, ["ชื่อผู้สอน"], instructor),
                "วันเวลาสอบ_กลางภาค": f"{m_slot} [{m_rm}]" if m_slot else "จัดไม่ได้",
                "ผู้คุมสอบ_กลางภาค": m_inv if m_slot else "-",
                "วันเวลาสอบ_ปลายภาค": f"{f_slot} [{f_rm}]" if f_slot else "จัดไม่ได้",
                "ผู้คุมสอบ_ปลายภาค": f_inv if f_slot else "-",
            })

    return pd.DataFrame(results), list(set(unassigned_warnings))

# ==================== 3. Streamlit Main UI ====================
if not st.session_state["logged_in"]:
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.title("ระบบจัดตารางสอบอัตโนมัติ")
        st.subheader("งานส่งเสริมวิชาการและงานทะเบียน วิทยาเขตจันทบุรี")
        username_input = st.text_input("ชื่อผู้ใช้งาน")
        password_input = st.text_input("รหัสผ่าน", type="password")
        if st.button("เข้าสู่ระบบ 🔐", use_container_width=True):
            if username_input in USER_CREDENTIALS and USER_CREDENTIALS[username_input] == password_input:
                st.session_state["logged_in"] = True
                st.session_state["username"] = username_input
                st.rerun()
            else:
                st.error("ชื่อผู้ใช้งานหรือรหัสผ่านไม่ถูกต้อง")
else:
    header_col1, header_col2 = st.columns([8, 2])
    with header_col1:
        st.title("🏫 ระบบจัดการตารางสอบ (สอบรวมห้องใหญ่)")
    with header_col2:
        st.write(f"ผู้ใช้งาน: **{st.session_state['username']}**")
        if st.button("ออกจากระบบ 🚪", on_click=logout, use_container_width=True): pass

    st.markdown("---")

    menu_selection = st.sidebar.radio(
        "📌 เลือกเมนูการทำงาน:",
        [
            "1. จัดตารางสอบประจำเทอม",
            "2. จัดการห้องสอบ (อัปเดตทุกเทอม)",
            "3. จัดการบุคลากรคุมสอบ/เจ้าหน้าที่สำรอง",
            "4. ประวัติตารางสอบที่จัดแล้ว (กันชน)", 
        ],
    )

    if menu_selection == "1. จัดตารางสอบประจำเทอม":
        st.header("🗓️ จัดตารางสอบอัตโนมัติ (สอบรวมห้องใหญ่)")
        st.info("💡 ระบบจะดึงข้อมูล 'ประวัติตารางสอบ' ที่จัดไปแล้วมาประมวลผล เพื่อป้องกันการใช้ห้องสอบหรือผู้คุมสอบซ้ำในวัน-เวลาเดียวกัน")

        st.subheader("⚙️ กำหนดเงื่อนไขจำนวนผู้คุมสอบขั้นต่ำตามสังกัดคณะ")
        fac_col1, fac_col2, fac_col3 = st.columns(3)
        with fac_col1: rule_agri = st.number_input("🌾 เทคโนฯ เกษตร (คน/ห้อง)", 1, 5, 2)
        with fac_col2: rule_soc = st.number_input("💼 เทคโนฯ สังคม (คน/ห้อง)", 1, 5, 2)
        with fac_col3: rule_eng = st.number_input("⚙️ คณะวิศวกรรมศาสตร์ (คน/ห้อง)", 1, 5, 2)
        faculty_rule_map = {
            "คณะเทคโนโลยีอุตสาหกรรมการเกษตร": rule_agri, 
            "คณะเทคโนโลยีสังคม": rule_soc, 
            "คณะวิศวกรรมศาสตร์": rule_eng
        }

        st.sidebar.markdown("---")
        st.sidebar.subheader("👥 เงื่อนไขการจัดผู้คุมสอบ")
        proctor_mode = st.sidebar.radio(
            "สังกัดคณะของผู้คุมสอบ:",
            [
                "จัดในคณะก่อน หากไม่พอค่อยข้ามคณะ (Priority)",
                "คณะตนเองเท่านั้น (Strict)",
                "ข้ามคณะได้อิสระ (Any Faculty)"
            ],
            index=0
        )

        st.sidebar.subheader("🗓️ กำหนดช่วงเวลาสอบ")
        m_start = st.sidebar.date_input("เริ่มกลางภาค", datetime.date(2026, 8, 24))
        m_end = st.sidebar.date_input("สิ้นสุดกลางภาค", datetime.date(2026, 8, 28))
        f_start = st.sidebar.date_input("เริ่มปลายภาค", datetime.date(2026, 10, 26))
        f_end = st.sidebar.date_input("สิ้นสุดปลายภาค", datetime.date(2026, 11, 1))

        # ปุ่มดาวน์โหลดตัวอย่างไฟล์
        test_file_data = create_test_subject_excel()
        st.sidebar.download_button(
            label="📥 ดาวน์โหลดไฟล์ตัวอย่าง (.xlsx)",
            data=test_file_data,
            file_name="test_subjects_multisec.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

        daily_slots = ["09:00 - 12:00", "13:30 - 16:30"]
        uploaded_file = st.file_uploader("นำเข้าไฟล์รายวิชาสอบ (.xlsx)", type=["xlsx"])

        if uploaded_file is not None:
            df_uploaded = pd.read_excel(uploaded_file)
            st.dataframe(df_uploaded, use_container_width=True)

            if st.button("เริ่มประมวลผลจัดตารางสอบ ⚡", type="primary"):
                slots_m = generate_time_slots(m_start, m_end, daily_slots)
                slots_f = generate_time_slots(f_start, f_end, daily_slots)

                df_new_result, warnings = auto_schedule_exams_combined(
                    df_uploaded, 
                    st.session_state["df_rooms"], 
                    st.session_state["df_staff_pool"], 
                    slots_m, 
                    slots_f, 
                    faculty_rule_map,
                    proctor_mode
                )

                # อัปเดตประวัติรวมใน Session State
                st.session_state["history_schedule"] = pd.concat([st.session_state["history_schedule"], df_new_result], ignore_index=True)

                st.success("✅ จัดตารางสอบเรียบร้อยแล้ว! ข้อมูลถูกบันทึกลงในหน่วยความจำเพื่อป้องกันการจัดชนในไฟล์ถัดไปแล้ว")
                if warnings:
                    st.warning("⚠️ **ข้อสังเกต:**")
                    for w in warnings: st.write(w)
                st.dataframe(df_new_result, use_container_width=True)

    elif menu_selection == "2. จัดการห้องสอบ (อัปเดตทุกเทอม)":
        st.header("🏫 จัดการข้อมูลห้องสอบประจำภาคเรียน")
        edited_rooms = st.data_editor(st.session_state["df_rooms"], num_rows="dynamic", use_container_width=True)
        if st.button("บันทึกการปรับปรุงข้อมูลห้องสอบ 💾", type="primary"):
            st.session_state["df_rooms"] = edited_rooms
            st.success("บันทึกเรียบร้อย!")

    elif menu_selection == "3. จัดการบุคลากรคุมสอบ/เจ้าหน้าที่สำรอง":
        st.header("👥 รายชื่อเจ้าหน้าที่คุมสอบเสริม")
        edited_staff = st.data_editor(st.session_state["df_staff_pool"], num_rows="dynamic", use_container_width=True)
        if st.button("บันทึกรายชื่อ 💾", type="primary"):
            st.session_state["df_staff_pool"] = edited_staff
            st.success("บันทึกเรียบร้อย!")
            
    elif menu_selection == "4. ประวัติตารางสอบที่จัดแล้ว (กันชน)":
        st.header("💾 ประวัติข้อมูลการจัดตารางสอบทั้งหมด")
        st.caption("ข้อมูลในตารางนี้คือวิชาที่จัดสอบไปแล้ว ซึ่งระบบจะใช้เป็นฐานข้อมูลในการกันไม่ให้ห้องสอบหรือกรรมการซ้ำกับไฟล์ใหม่ที่คุณกำลังจะอัปโหลด")
        
        if not st.session_state["history_schedule"].empty:
            st.dataframe(st.session_state["history_schedule"], use_container_width=True)
            
            st.markdown("---")
            st.subheader("🗑️ ล้างข้อมูลประวัติการจัดตารางสอบ")
            
            # --- ยืนยันชั้นที่ 2 สำหรับการลบข้อมูลป้องกันการกดผิด ---
            if not st.session_state.get("show_confirm_clear", False):
                if st.button("🚨 ล้างประวัติทั้งหมด", type="secondary"):
                    st.session_state["show_confirm_clear"] = True
                    st.rerun()
            else:
                st.warning("⚠️ **ยืนยันการลบข้อมูล (ชั้นที่ 2):** การล้างประวัติจะลบข้อมูลตารางสอบที่จัดไปแล้วทั้งหมดออกจากระบบ และคืนค่าห้องสอบ/กรรมการให้ว่าง คุณแน่ใจหรือไม่ว่าต้องการดำเนินการต่อ?")
                col_confirm, col_cancel = st.columns([2, 2])
                with col_confirm:
                    if st.button("🔥 ยืนยันการลบข้อมูลทั้งหมด (Confirm Delete)", type="primary"):
                        init_memory()
                        st.success("✅ ล้างข้อมูลประวัติและคืนค่าห้องสอบ/กรรมการเรียบร้อยแล้ว!")
                        st.rerun()
                with col_cancel:
                    if st.button("❌ ยกเลิก (Cancel)"):
                        st.session_state["show_confirm_clear"] = False
                        st.rerun()
        else:
            st.info("ยังไม่มีข้อมูลประวัติการจัดตารางสอบในขณะนี้")
