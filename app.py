import datetime
import io
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
import pandas as pd
import streamlit as st

# ==================== 1. ตั้งค่าหน้าจอโปรแกรม ====================
st.set_page_config(
    page_title="ระบบจัดตารางสอบ - มทร.ตะวันออก วิทยาเขตจันทบุรี",
    page_icon="🏫",
    layout="wide",
)

USER_CREDENTIALS = {"monthira": "123456", "registry_staff": "rmutto456"}

# Session States สำหรับเก็บ Master Data ที่อัปเดตได้ตลอดเวลา
if "logged_in" not in st.session_state:
    st.session_state["logged_in"] = False
if "username" not in st.session_state:
    st.session_state["username"] = ""

if "history_schedule" not in st.session_state:
    st.session_state["history_schedule"] = pd.DataFrame()

# 1. ฐานข้อมูลห้องสอบ (เริ่มต้น)
if "df_rooms" not in st.session_state:
    st.session_state["df_rooms"] = pd.DataFrame([
        {
            "อาคาร": "อาคารเรียนรวม 36 ปี",
            "รหัสห้อง": "36-301",
            "ความจุสอบ": 40,
            "ประเภท": "ห้องทฤษฎี",
            "สถานะ": "ใช้งานได้",
        },
        {
            "อาคาร": "อาคารเรียนรวม 36 ปี",
            "รหัสห้อง": "36-302",
            "ความจุสอบ": 80,
            "ประเภท": "ห้องทฤษฎี",
            "สถานะ": "ใช้งานได้",
        },
        {
            "อาคาร": "อาคารปฏิบัติการไอที",
            "รหัสห้อง": "IT-201",
            "ความจุสอบ": 50,
            "ประเภท": "ห้องปฏิบัติการคอมพิวเตอร์",
            "สถานะ": "ใช้งานได้",
        },
    ])

# 2. ฐานข้อมูลบุคลากรคุมสอบสำรอง/อาจารย์นอกตาราง (เริ่มต้น)
if "df_staff_pool" not in st.session_state:
    st.session_state["df_staff_pool"] = pd.DataFrame([
        {
            "คณะ": "คณะเทคโนโลยีอุตสาหกรรมการเกษตร",
            "ชื่อ-นามสกุล": "ดร.สมเกียรติ มั่นคง",
            "ตำแหน่ง": "อาจารย์ (ไม่ได้ลงสอนเทอมนี้)",
            "ประเภท": "อาจารย์ในคณะ",
        },
        {
            "คณะ": "คณะเทคโนโลยีอุตสาหกรรมการเกษตร",
            "ชื่อ-นามสกุล": "นายวิชัย สำรองดี",
            "ตำแหน่ง": "เจ้าหน้าที่บริหารงานทั่วไป",
            "ประเภท": "เจ้าหน้าที่สำรองส่วนกลาง",
        },
        {
            "คณะ": "คณะเทคโนโลยีสังคม",
            "ชื่อ-นามสกุล": "อ.ประเสริฐ นามดี",
            "ตำแหน่ง": "อาจารย์ (ไม่ได้ลงสอนเทอมนี้)",
            "ประเภท": "อาจารย์ในคณะ",
        },
        {
            "คณะ": "คณะเทคโนโลยีสังคม",
            "ชื่อ-นามสกุล": "นางสาวนภา ใจเย็น",
            "ตำแหน่ง": "นักวิชาการศึกษา",
            "ประเภท": "เจ้าหน้าที่สำรองส่วนกลาง",
        },
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


def generate_time_slots(start_date, end_date, daily_slots):
    slots = []
    current_date = start_date
    while current_date <= end_date:
        if current_date.weekday() < 5:
            date_str = current_date.strftime("%d/%m/%Y")
            for slot_name in daily_slots:
                slots.append({
                    "full_slot": f"{date_str} ({slot_name})",
                    "date_str": date_str,
                    "time_str": slot_name,
                })
        current_date += datetime.timedelta(days=1)
    return slots


# ==================== 2. อัลกอริทึมจัดตารางสอบ ====================
def auto_schedule_exams_advanced(
    df_subjects,
    df_rooms,
    df_staff_pool,
    slots_m,
    slots_f,
    existing_schedule=None,
    faculty_rule_map=None,
):
    if faculty_rule_map is None:
        faculty_rule_map = {}

    student_group_occupancy = {}
    student_group_daily_count = {}
    student_group_daily_heavy = {}
    room_occupancy = {}
    invigilator_occupancy = {}
    invigilator_workload = {}

    unassigned_warnings = []

    # โหลดประวัติการจัดสอบเดิม
    if existing_schedule is not None and not existing_schedule.empty:
        for _, ex_row in existing_schedule.iterrows():
            grp = str(ex_row.get("กลุ่มเรียน", "")).strip()
            for phase in ["M", "F"]:
                info = str(ex_row.get(f"วันเวลาสอบ_{phase}", ""))
                if info and " [ห้อง " in info:
                    slot_str = info.split(" [ห้อง ")[0].strip()
                    rm_str = (
                        info.split(" [ห้อง ")[1].replace("]", "").strip()
                    )
                    invs = str(ex_row.get(f"ผู้คุมสอบ_{phase}", "")).split(",")
                    date_part = (
                        slot_str.split(" ")[0] if " " in slot_str else slot_str
                    )

                    student_group_occupancy.setdefault(
                        slot_str, set()
                    ).add(grp)
                    student_group_daily_count[
                        (grp, date_part)
                    ] = student_group_daily_count.get((grp, date_part), 0) + 1

                    if rm_str:
                        room_occupancy.setdefault(slot_str, set()).add(
                            rm_str
                        )

                    for inv in invs:
                        inv = inv.strip()
                        if inv:
                            invigilator_occupancy.setdefault(
                                slot_str, set()
                            ).add(inv)
                            invigilator_workload[inv] = (
                                invigilator_workload.get(inv, 0) + 1
                            )

    grouped_subjects = {}
    for idx, row in df_subjects.iterrows():
        subj_name = str(
            get_column_value(row, ["ชื่อวิชา", "รายวิชา"], "ไม่ระบุวิชา")
        ).strip()
        instructor = str(
            get_column_value(
                row,
                ["ชื่อผู้สอน", "ผู้สอน", "อาจารย์ผู้สอน", "อาจารย์"],
                "อาจารย์ผู้สอน",
            )
        ).strip()

        group_key = f"{subj_name}||{instructor}"
        if group_key not in grouped_subjects:
            grouped_subjects[group_key] = []
        grouped_subjects[group_key].append(row)

    results = []

    # ฟังก์ชันคำนวณผู้คุมสอบยืดหยุ่นตามกฎคณะ
    def assign_invigilators(
        faculty_name, instructor, total_students, slot_str, subj_teachers
    ):
        rule = faculty_rule_map.get(faculty_name, 1)
        used_invs = invigilator_occupancy.get(slot_str, set())
        assigned = []

        # ดึงรายชื่ออาจารย์นอกตาราง และ เจ้าหน้าที่สำรอง ของคณะนี้จาก Master Data
        fac_staff_df = df_staff_pool[df_staff_pool["คณะ"] == faculty_name]
        extra_teachers = list(
            fac_staff_df[fac_staff_df["ประเภท"] == "อาจารย์ในคณะ"][
                "ชื่อ-นามสกุล"
            ].unique()
        )
        backup_staffs = list(
            fac_staff_df[
                fac_staff_df["ประเภท"] == "เจ้าหน้าที่สำรองส่วนกลาง"
            ]["ชื่อ-นามสกุล"].unique()
        )

        # ---------------- Rule 1 ----------------
        if rule == 1:
            req_count = 1
            if instructor not in used_invs:
                assigned.append(instructor)

        # ---------------- Rule 2 ----------------
        elif rule == 2:
            req_count = 2
            if instructor not in used_invs:
                assigned.append(instructor)

            # ดึงอาจารย์คนอื่นในคณะ (ทั้งที่มีสอนและไม่มีสอนในเทอมนี้)
            pool = list(set(subj_teachers + extra_teachers))
            pool.sort(key=lambda t: invigilator_workload.get(t, 0))

            for t in pool:
                if t not in assigned and t not in used_invs:
                    assigned.append(t)
                    if len(assigned) == req_count:
                        break

            # ถ้าอาจารย์ไม่พอ ดึงเจ้าหน้าที่สำรองส่วนกลาง
            if len(assigned) < req_count:
                for staff in backup_staffs:
                    if staff not in assigned and staff not in used_invs:
                        assigned.append(staff)
                        if len(assigned) == req_count:
                            break

        # ---------------- Rule 3 ----------------
        elif rule == 3:
            if total_students <= 30:
                req_count = 1
            elif total_students <= 60:
                req_count = 2
            else:
                req_count = 3

            # ผู้สอนคุมได้ถ้าว่าง (ไม่บังคับ)
            if instructor not in used_invs:
                assigned.append(instructor)

            # ดึงอาจารย์ทั้งหมดในคณะ
            pool = list(set(subj_teachers + extra_teachers))
            pool.sort(key=lambda t: invigilator_workload.get(t, 0))

            for t in pool:
                if t not in assigned and t not in used_invs:
                    assigned.append(t)
                    if len(assigned) == req_count:
                        break

            # ถ้าอาจารย์ไม่พอ ดึงเจ้าหน้าที่สำรองส่วนกลาง
            if len(assigned) < req_count:
                for staff in backup_staffs:
                    if staff not in assigned and staff not in used_invs:
                        assigned.append(staff)
                        if len(assigned) == req_count:
                            break

        if len(assigned) >= req_count:
            return assigned
        return None

    # วนลูปจัดวิชา
    for group_key, rows in grouped_subjects.items():
        total_students = 0
        groups_list = []
        for r in rows:
            grp = str(
                get_column_value(r, ["กลุ่มเรียน", "กลุ่ม", "Sec"], "")
            ).strip()
            groups_list.append(grp)
            try:
                total_students += int(
                    get_column_value(
                        r, ["จำนวนผู้เข้าสอบ", "จำนวนนักศึกษา", "ลง"], 0
                    )
                )
            except ValueError:
                pass

        first_row = rows[0]
        subj_code_display = get_column_value(
            first_row, ["รหัสวิชา", "รหัสวิชาสอบ"], "SUBJ"
        )
        subj_name_display = get_column_value(
            first_row, ["ชื่อวิชา", "รายวิชา"], "ไม่ระบุวิชา"
        )
        faculty_name = get_column_value(
            first_row, ["คณะ", "สังกัดคณะ"], "คณะเทคโนโลยีอุตสาหกรรมการเกษตร"
        )
        exam_type = str(
            get_column_value(first_row, ["ประเภทการสอบ", "ประเภท"], "ทฤษฎี")
        ).strip()
        is_heavy = (
            str(
                get_column_value(
                    first_row, ["วิชาคำนวณ", "วิชาหนัก", "Heavy"], "NO"
                )
            )
            .strip()
            .upper()
            in ["YES", "Y", "TRUE", "1", "คำนวณ"]
        )
        instructor = get_column_value(
            first_row,
            ["ชื่อผู้สอน", "ผู้สอน", "อาจารย์ผู้สอน", "อาจารย์"],
            "อาจารย์ผู้สอน",
        )

        active_rooms = df_rooms[df_rooms["สถานะ"] == "ใช้งานได้"]
        target_type = (
            "ห้องปฏิบัติการคอมพิวเตอร์"
            if ("ปฏิบัติ" in exam_type or "คอม" in exam_type)
            else "ห้องทฤษฎี"
        )
        valid_rooms = active_rooms[
            (active_rooms["ประเภท"] == target_type)
            & (active_rooms["ความจุสอบ"] >= total_students)
        ]

        subj_teachers = list(
            df_subjects[
                get_column_value(
                    df_subjects,
                    ["ชื่อผู้สอน", "ผู้สอน"],
                    df_subjects.columns[0],
                )
            ]
            .dropna()
            .unique()
        )

        def find_best_slot(slots_list, hrs_val):
            if hrs_val <= 0:
                return None, None, None
            if valid_rooms.empty:
                return None, None, "NO_ROOM"

            for slot_obj in slots_list:
                slot_str = slot_obj["full_slot"]
                date_str = slot_obj["date_str"]

                # 1. Hard Constraint: ห้ามสอบซ้ำเวลา
                conflict = any(
                    g in student_group_occupancy.get(slot_str, set())
                    for g in groups_list
                )
                if conflict:
                    continue

                # 2. Soft Constraint: ไม่เกิน 2 วิชา/วัน
                if (
                    max(
                        student_group_daily_count.get((g, date_str), 0)
                        for g in groups_list
                    )
                    >= 2
                ):
                    continue

                # 3. Soft Constraint: เลี่ยงวิชาหนักซ้ำวัน
                if is_heavy and any(
                    student_group_daily_heavy.get((g, date_str), False)
                    for g in groups_list
                ):
                    continue

                # 4. Hard Constraint: จัดผู้คุมสอบ
                invig_list = assign_invigilators(
                    faculty_name,
                    instructor,
                    total_students,
                    slot_str,
                    subj_teachers,
                )
                if not invig_list:
                    continue

                # 5. หาห้องสอบ
                used_rms = room_occupancy.get(slot_str, set())
                avail_rm = None
                for _, rm in valid_rooms.iterrows():
                    if rm["รหัสห้อง"] not in used_rms:
                        avail_rm = rm["รหัสห้อง"]
                        break

                if avail_rm:
                    for g in groups_list:
                        student_group_occupancy.setdefault(
                            slot_str, set()
                        ).add(g)
                        student_group_daily_count[(g, date_str)] = (
                            student_group_daily_count.get((g, date_str), 0) + 1
                        )
                        if is_heavy:
                            student_group_daily_heavy[(g, date_str)] = True

                    room_occupancy.setdefault(slot_str, set()).add(avail_rm)
                    for inv in invig_list:
                        invigilator_occupancy.setdefault(
                            slot_str, set()
                        ).add(inv)
                        invigilator_workload[inv] = (
                            invigilator_workload.get(inv, 0) + 1
                        )

                    return slot_str, avail_rm, ", ".join(invig_list)

            return None, None, "NO_SLOT"

        # กลางภาค
        hrs_m = float(
            get_column_value(
                first_row, ["ชั่วโมงสอบ_M", "ชั่วโมงสอบกลางภาค"], 1.5
            )
        )
        m_slot, m_rm, m_inv = find_best_slot(slots_m, hrs_m)
        if not m_slot and hrs_m > 0:
            unassigned_warnings.append(
                f"❌ **[{subj_code_display}] {subj_name_display} (กลางภาค)**: จัดลงตารางไม่ได้ (ห้องเต็ม/ผู้คุมสอบไม่เพียงพอ)"
            )

        # ปลายภาค
        hrs_f = float(
            get_column_value(
                first_row, ["ชั่วโมงสอบ_F", "ชั่วโมงสอบปลายภาค"], 2.0
            )
        )
        f_slot, f_rm, f_inv = find_best_slot(slots_f, hrs_f)
        if not f_slot and hrs_f > 0:
            unassigned_warnings.append(
                f"❌ **[{subj_code_display}] {subj_name_display} (ปลายภาค)**: จัดลงตารางไม่ได้ (ห้องเต็ม/ผู้คุมสอบไม่เพียงพอ)"
            )

        for r in rows:
            results.append({
                "คณะ": faculty_name,
                "รหัสวิชา": get_column_value(
                    r, ["รหัสวิชา", "รหัสวิชาสอบ"], "SUBJ"
                ),
                "ชื่อวิชา": get_column_value(
                    r, ["ชื่อวิชา", "รายวิชา"], "ไม่ระบุวิชา"
                ),
                "ชื่อผู้สอน": instructor,
                "สังกัดสาขา": get_column_value(r, ["สังกัดสาขา", "สาขา"], ""),
                "กลุ่มเรียน": str(
                    get_column_value(r, ["กลุ่มเรียน", "กลุ่ม", "Sec"], "")
                ).strip(),
                "จำนวนผู้เข้าสอบ": int(
                    get_column_value(
                        r, ["จำนวนผู้เข้าสอบ", "จำนวนนักศึกษา", "ลง"], 0
                    )
                ),
                "ชั่วโมงสอบ_M": hrs_m if m_slot else "",
                "ชั่วโมงสอบ_F": hrs_f if f_slot else "",
                "วันเวลาสอบ_M": (
                    f"{m_slot} [ห้อง {m_rm}]" if m_slot else "ไม่มีสอบ"
                ),
                "วันเวลาสอบ_F": (
                    f"{f_slot} [ห้อง {f_rm}]" if f_slot else "ไม่มีสอบ"
                ),
                "ผู้คุมสอบ_M": m_inv if m_slot else "",
                "ผู้คุมสอบ_F": f_inv if f_slot else "",
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
            if (
                username_input in USER_CREDENTIALS
                and USER_CREDENTIALS[username_input] == password_input
            ):
                st.session_state["logged_in"] = True
                st.session_state["username"] = username_input
                st.rerun()
            else:
                st.error("ชื่อผู้ใช้งานหรือรหัสผ่านไม่ถูกต้อง")

else:
    header_col1, header_col2 = st.columns([8, 2])
    with header_col1:
        st.title("🏫 ระบบจัดการตารางสอบ (Dynamic Master Data)")
        st.caption("มหาวิทยาลัยเทคโนโลยีราชมงคลตะวันออก วิทยาเขตจันทบุรี")
    with header_col2:
        st.write(f"ผู้ใช้งาน: **{st.session_state['username']}**")
        if st.button("ออกจากระบบ 🚪", on_click=logout, use_container_width=True):
            pass

    st.markdown("---")

    menu_selection = st.sidebar.radio(
        "📌 เลือกเมนูการทำงาน:",
        [
            "1. จัดตารางสอบประจำเทอม",
            "2. จัดการห้องสอบ (อัปเดตทุกเทอม)",
            "3. จัดการบุคลากรคุมสอบ/เจ้าหน้าที่สำรอง (อัปเดตทุกเทอม)",
        ],
    )

    # ------------------ เมนูที่ 1: จัดตารางสอบ ------------------
    if menu_selection == "1. จัดตารางสอบประจำเทอม":
        st.header("🗓️ จัดตารางสอบอัตโนมัติ")

        st.sidebar.markdown("---")
        st.sidebar.header("⚙️ ตั้งค่าเงื่อนไขผู้คุมสอบประจำคณะ")
        rule_fac1 = st.sidebar.selectbox(
            "คณะเทคโนโลยีอุตสาหกรรมการเกษตร",
            [1, 2, 3],
            format_func=lambda x: f"กฎคณะที่ {x}",
            index=0,
        )
        rule_fac2 = st.sidebar.selectbox(
            "คณะเทคโนโลยีสังคม",
            [1, 2, 3],
            format_func=lambda x: f"กฎคณะที่ {x}",
            index=1,
        )
        faculty_rule_map = {
            "คณะเทคโนโลยีอุตสาหกรรมการเกษตร": rule_fac1,
            "คณะเทคโนโลยีสังคม": rule_fac2,
        }

        st.sidebar.markdown("---")
        st.sidebar.subheader("🗓️ กำหนดช่วงเวลาสอบ")
        m_start = st.sidebar.date_input(
            "เริ่มกลางภาค", datetime.date(2026, 8, 24)
        )
        m_end = st.sidebar.date_input(
            "สิ้นสุดกลางภาค", datetime.date(2026, 8, 28)
        )
        f_start = st.sidebar.date_input(
            "เริ่มปลายภาค", datetime.date(2026, 10, 26)
        )
        f_end = st.sidebar.date_input(
            "สิ้นสุดปลายภาค", datetime.date(2026, 11, 1)
        )

        daily_slots = [
            "09:00 - 11:00",
            "11:00 - 13:00",
            "13:30 - 15:30",
            "15:30 - 17:30",
        ]

        uploaded_file = st.file_uploader(
            "นำเข้าไฟล์รายวิชาสอบ (.xlsx)", type=["xlsx"]
        )

        if uploaded_file is not None:
            df_uploaded = pd.read_excel(uploaded_file)
            st.write("📋 **รายการวิชาสอบที่นำเข้า:**")
            st.dataframe(df_uploaded)

            if st.button("เริ่มประมวลผลจัดตารางสอบ ⚡", type="primary"):
                slots_m = generate_time_slots(m_start, m_end, daily_slots)
                slots_f = generate_time_slots(f_start, f_end, daily_slots)

                df_new_result, warnings = auto_schedule_exams_advanced(
                    df_uploaded,
                    st.session_state["df_rooms"],
                    st.session_state["df_staff_pool"],
                    slots_m,
                    slots_f,
                    existing_schedule=st.session_state["history_schedule"],
                    faculty_rule_map=faculty_rule_map,
                )

                st.session_state["history_schedule"] = pd.concat(
                    [st.session_state["history_schedule"], df_new_result],
                    ignore_index=True,
                )

                st.balloons()
                st.success("✅ ประมวลผลตารางสอบเรียบร้อยแล้ว!")

                if warnings:
                    st.error("🚨 **ตรวจพบข้อผิดพลาดที่ไม่สามารถจัดลงตารางได้:**")
                    for w in warnings:
                        st.write(w)

                st.subheader("📊 ตารางสอบรวม")
                st.dataframe(st.session_state["history_schedule"])

    # ------------------ เมนูที่ 2: จัดการห้องสอบ ------------------
    elif menu_selection == "2. จัดการห้องสอบ (อัปเดตทุกเทอม)":
        st.header("🏫 จัดการข้อมูลห้องสอบประจำภาคเรียน")
        st.caption(
            "สามารถแก้ไข ลบ หรือเพิ่มห้องสอบใหม่ในตารางด้านล่างได้ทันที ข้อมูลจะถูกนำไปใช้คำนวณตารางสอบทันที"
        )

        # ใช้ Data Editor ให้กดแก้บนตารางได้แบบ Excel
        edited_rooms = st.data_editor(
            st.session_state["df_rooms"],
            num_rows="dynamic",
            use_container_width=True,
            key="room_editor",
        )

        if st.button("บันทึกการปรับปรุงข้อมูลห้องสอบ 💾", type="primary"):
            st.session_state["df_rooms"] = edited_rooms
            st.success("บันทึกข้อมูลห้องสอบเรียบร้อยแล้ว!")

    # ------------------ เมนูที่ 3: จัดการบุคลากรคุมสอบ ------------------
    elif (
        menu_selection
        == "3. จัดการบุคลากรคุมสอบ/เจ้าหน้าที่สำรอง (อัปเดตทุกเทอม)"
    ):
        st.header(
            "👥 จัดการรายชื่ออาจารย์นอกตารางสอบ & เจ้าหน้าที่สำรองส่วนกลาง"
        )
        st.info(
            "💡 ข้อมูลส่วนนี้จะถูกดึงไปใช้คุมสอบอัตโนมัติ สำหรับ **คณะที่ 2 และ คณะที่ 3** (หรือคณะที่ 1 หากอนาคตเปลี่ยนกฎ)"
        )

        edited_staff = st.data_editor(
            st.session_state["df_staff_pool"],
            num_rows="dynamic",
            use_container_width=True,
            column_config={
                "ประเภท": st.column_config.SelectboxColumn(
                    "ประเภทบุคลากร",
                    options=[
                        "อาจารย์ในคณะ",
                        "เจ้าหน้าที่สำรองส่วนกลาง",
                    ],
                    required=True,
                ),
                "คณะ": st.column_config.SelectboxColumn(
                    "สังกัดคณะ",
                    options=[
                        "คณะเทคโนโลยีอุตสาหกรรมการเกษตร",
                        "คณะเทคโนโลยีสังคม",
                    ],
                    required=True,
                ),
            },
            key="staff_editor",
        )

        if st.button("บันทึกการปรับปรุงรายชื่อบุคลากร 💾", type="primary"):
            st.session_state["df_staff_pool"] = edited_staff
            st.success("บันทึกข้อมูลรายชื่อบุคลากรเรียบร้อยแล้ว!")
