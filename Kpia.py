import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import os
import json

st.set_page_config(page_title="شركة العلمين فليكس للطباعة - نظام KPIs المتقدم", page_icon="⚙️", layout="wide")

EXCEL_FILE = "maintenance_kpi_database.xlsx"
MACHINES_CONFIG_FILE = "machines_config.json"

DEFAULT_MACHINES_BY_FACTORY = {
    "مصنع الطباعة": [
        "9 colors press", "8 colors press", "Comexi lam", "Nord1 lam", "Nord2 lam", "Nord3 lam",
        "Bemic slit", "Kesheng slit", "Rewinder mc", "Cheeter mc", "Slave machin",
        "Rewinder mc sm", "Wax mc", "Core cutter old", "Core cutter new"
    ],
    "مصنع السلندرات": [
        "Old engrave mc", "New engrave mc", "Chinese engrave mc", "Old cfm", "New cfm",
        "Finish mc", "Prova mc", "German Chrome tank", "Chinese chrome tank",
        "Copper Chinese tank", "German copper tank", "German nakil tank"
    ],
    "محطة السولفنت": [],
    "الخدمات (Chiller/Dryer/Compressors)": [
        "Big chiller", "Cylinders chiller", "Ink cooling conditioning chiller",
        "Keasir big air compressor", "Keasir small air compressor", "Air dryer"
    ]
}

def load_machines():
    if os.path.exists(MACHINES_CONFIG_FILE):
        try:
            with open(MACHINES_CONFIG_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
            for k, v in DEFAULT_MACHINES_BY_FACTORY.items():
                if k not in data:
                    data[k] = v
            return data
        except Exception:
            return {k: list(v) for k, v in DEFAULT_MACHINES_BY_FACTORY.items()}
    else:
        fresh = {k: list(v) for k, v in DEFAULT_MACHINES_BY_FACTORY.items()}
        save_machines(fresh)
        return fresh

def save_machines(data):
    with open(MACHINES_CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

if "last_notification_number" not in st.session_state:
    st.session_state["last_notification_number"] = None

def load_data():
    if os.path.exists(EXCEL_FILE):
        df = pd.read_excel(EXCEL_FILE)
        if not df.empty and "التاريخ" in df.columns:
            df["التاريخ_dt"] = pd.to_datetime(df["التاريخ"], errors='coerce')
        return df
    return pd.DataFrame(columns=[
        "رقم الإخطار", "التاريخ", "المصنع/القسم", "رقم الوردية", "رقم/اسم الماكينة", 
        "اسم مشغل الماكينة", "اسم القائم بالصيانة", "تخصص العطل", "طبيعة الصيانة",
        "حالة الماكينة النهائية", "كود/اسم قطعة الغيار", "تكلفة قطعة الغيار (جنيه)",
        "وقت البداية", "وقت النهاية", "مدة العطل (ساعة)", "السبب الرئيسي", 
        "عطل صيانة", "تأخير مشتريات", "صيانة مخططة", "خطأ مشغل", "تلف سلندر حفر",
        "الساعات التشغيلية المتاحة", "توافرية الصيانة (%)", "ملاحظات"
    ])

def save_data(df):
    if "التاريخ_dt" in df.columns:
        df = df.drop(columns=["التاريخ_dt"], errors="ignore")
    df.to_excel(EXCEL_FILE, index=False)

st.title("🏭 شركة العلمين فليكس للطباعة")
st.subheader("⚙️ منظومة إدارة ومتابعة الصيانة الشاملة (CMMS & KPIs)")

tab1, tab2, tab3 = st.tabs(["📝 إدخال إخطار عطل", "📊 قاعدة البيانات والتحليلات", "📈 المؤشرات الهندسية (MTBF / MTTR)"])

with tab1:
    st.header("🚨 إخطار عطل جديد")

    if st.session_state.get("last_notification_number"):
        st.success(f"✅ آخر إخطار تم حفظه في هذه الجلسة: رقم **{st.session_state['last_notification_number']}**")

    st.subheader("⏱️ تفاصيل التاريخ والأوقات (حساب فوري)")
    col_t0, col_t1, col_t2, col_t3 = st.columns(4)
    with col_t0:
        date_input = st.date_input("التاريخ", datetime.now())
    with col_t1:
        t1_input = st.time_input("وقت بداية العطل", datetime.strptime("08:00", "%H:%M").time())
    with col_t2:
        t2_input = st.time_input("وقت نهاية العطل", datetime.strptime("09:30", "%H:%M").time())
        
    t1_dt_live = datetime.combine(datetime.today(), t1_input)
    t2_dt_live = datetime.combine(datetime.today(), t2_input)
    if t2_dt_live < t1_dt_live:
        t2_dt_live += timedelta(days=1)
    live_duration = round((t2_dt_live - t1_dt_live).total_seconds() / 3600.0, 2)
    
    with col_t3:
        st.metric("مدة العطل المحسوبة حالياً", f"{live_duration} ساعة")

    _preview_df = load_data()
    if not _preview_df.empty and "التاريخ_dt" in _preview_df.columns:
        _monthly_count_preview = _preview_df[
            (_preview_df["التاريخ_dt"].dt.month == date_input.month) &
            (_preview_df["التاريخ_dt"].dt.year == date_input.year)
        ].shape[0]
    else:
        _monthly_count_preview = 0
    preview_notification_number = f"{date_input.strftime('%m-%Y')}-{_monthly_count_preview + 1:03d}"
    st.info(f"🔢 **رقم الإخطار الذي سيتم تسجيله:** {preview_notification_number}")

    st.subheader("🏭 المصنع / القسم والماكينة")
    col_f1, col_f2 = st.columns(2)

    machines_data = load_machines()

    with col_f1:
        factory_site = st.selectbox("المصنع / القسم", [
            "مصنع الطباعة", 
            "مصنع السلندرات", 
            "محطة السولفنت", 
            "الخدمات (Chiller/Dryer/Compressors)"
        ])
    with col_f2:
        _machine_options = machines_data.get(factory_site, [])
        if _machine_options:
            machine_id = st.selectbox("رقم / اسم الماكينة", _machine_options, key="machine_select")
        else:
            machine_id = None
            st.selectbox("رقم / اسم الماكينة", ["— لا توجد ماكينات مضافة بعد —"], disabled=True)

    with st.expander(f"➕ إضافة ماكينة جديدة إلى قائمة: {factory_site}"):
        col_add1, col_add2 = st.columns([3, 1])
        with col_add1:
            new_machine_name = st.text_input("اسم / رقم الماكينة الجديدة", key="new_machine_input", label_visibility="collapsed", placeholder="اكتب اسم الماكينة الجديدة هنا")
        with col_add2:
            add_machine_clicked = st.button("➕ إضافة", key="add_machine_btn", use_container_width=True)
        if add_machine_clicked:
            cleaned_name = new_machine_name.strip()
            if not cleaned_name:
                st.warning("يرجى إدخال اسم الماكينة أولاً.")
            elif cleaned_name in machines_data.get(factory_site, []):
                st.warning("هذه الماكينة موجودة بالفعل في القائمة.")
            else:
                machines_data.setdefault(factory_site, []).append(cleaned_name)
                save_machines(machines_data)
                st.success(f"✅ تمت إضافة '{cleaned_name}' إلى قائمة {factory_site}.")
                st.rerun()

    st.divider()

    with st.form("kpi_form"):
        st.subheader("📌 البيانات الأساسية للوردية والأفراد")
        col1, col2, col3 = st.columns(3)
        with col1:
            st.caption(f"📅 التاريخ المحدد: **{date_input.strftime('%Y-%m-%d')}**")
            st.caption(f"🏭 المصنع/القسم: **{factory_site}**")
            st.caption(f"⚙️ الماكينة: **{machine_id if machine_id else 'لم يتم الاختيار'}**")
        with col2:
            shift_num = st.selectbox("رقم الوردية", ["الوردية الأولى (1)", "الوردية الثانية (2)", "الوردية الثالثة (3)"])
        with col3:
            operator_name = st.text_input("اسم مشغل الماكينة", "")
            technician_name = st.text_input("اسم القائم بالصيانة", "")

        st.divider()

        st.subheader("🛠️ التصنيف الهندسي وحالة الماكينة")
        col_f1, col_f2, col_f3 = st.columns(3)
        with col_f1:
            # تم إضافة بند تخصص سلندرات الحفر هنا
            fault_type = st.selectbox("تخصص العطل", [
                "كهرباء", 
                "ميكانيكا", 
                "تحكم وآليات PLC", 
                "هيدروليك ونيوماتيك", 
                "كروت إلكترونية",
                "سلندرات حفر وتجهيز"
            ])
            maint_nature = st.selectbox("طبيعة الصيانة", ["عطل طارئ (Emergency)", "صيانة وقائية (PM)", "تحسين وتطوير (Modification)"])
        with col_f2:
            final_status = st.selectbox("حالة الماكينة عند المغادرة", ["تشغيل كلي", "تشغيل جزئي مؤقت", "متوقفة بانتظار قطع غيار / سلندر"])
            spare_part_code = st.text_input("كود / اسم قطعة الغيار المستهلكة", "بدون / Spare Part Code")
        with col_f3:
            spare_part_cost = st.number_input("تكلفة قطع الغيار التقديرية (جنيه)", min_value=0.0, value=0.0, step=50.0)
            # تم إضافة خيار تلف / عيب سلندر حفر هنا
            cause_cat = st.selectbox("السبب الرئيسي للعطل", [
                "عطل صيانة", 
                "تلف / عيب سلندر حفر",
                "تأخير مشتريات", 
                "صيانة مخططة", 
                "خطأ مشغل"
            ])

        st.divider()

        st.subheader("📋 الساعات التشغيلية والملاحظات")
        col_n1, col_n2 = st.columns(2)
        with col_n1:
            free_hrs = st.number_input("الساعات التشغيلية المتاحة للوردية (ساعة)", min_value=0.0, value=8.0, step=0.5)
        with col_n2:
            notes = st.text_area("ملاحظات / أرقام طلبات الشراء ورقم السلندر والتفاصيل", "")

        submit = st.form_submit_button("حفظ الإخطار في Excel 💾", use_container_width=True)

        if submit:
            if not machine_id:
                st.error("⚠️ لا يمكن الحفظ: لا توجد ماكينة مختارة لهذا القسم. أضف ماكينة أولاً من قسم '➕ إضافة ماكينة جديدة' بالأعلى.")
                st.stop()
            try:
                df_curr = load_data()
                str_date = str(date_input)
                str_t1 = t1_input.strftime("%H:%M")

                if not df_curr.empty:
                    duplicate_check = df_curr[
                        (df_curr["التاريخ"] == str_date) & 
                        (df_curr["المصنع/القسم"] == factory_site) & 
                        (df_curr["رقم الوردية"] == shift_num) & 
                        (df_curr["رقم/اسم الماكينة"] == machine_id) & 
                        (df_curr["وقت البداية"] == str_t1)
                    ]
                    
                    if not duplicate_check.empty:
                        st.error(f"⚠️ **تم رفض الحفظ!** تم تسجيل إخطار عطل سابق لنفس الماكينة ({machine_id}) في نفس التاريخ والوردية ووقت البداية ({str_t1}).")
                        st.stop()

                duration = live_duration

                maint_h = duration if cause_cat == "عطل صيانة" else 0.0
                cyl_h = duration if cause_cat == "تلف / عيب سلندر حفر" else 0.0
                proc_h = duration if cause_cat == "تأخير مشتريات" else 0.0
                pm_h = duration if cause_cat == "صيانة مخططة" else 0.0
                op_h = duration if cause_cat == "خطأ مشغل" else 0.0

                total_planned = free_hrs + duration
                eff_planned = total_planned - (proc_h + op_h + cyl_h)
                availability = (free_hrs / eff_planned * 100) if eff_planned > 0 else 100.0

                month_year_str = date_input.strftime("%m-%Y")
                if not df_curr.empty and "التاريخ_dt" in df_curr.columns:
                    monthly_count = df_curr[
                        (df_curr["التاريخ_dt"].dt.month == date_input.month) &
                        (df_curr["التاريخ_dt"].dt.year == date_input.year)
                    ].shape[0]
                else:
                    monthly_count = 0
                notification_number = f"{month_year_str}-{monthly_count + 1:03d}"

                new_row = {
                    "رقم الإخطار": notification_number,
                    "التاريخ": str_date,
                    "المصنع/القسم": factory_site,
                    "رقم الوردية": shift_num,
                    "رقم/اسم الماكينة": machine_id,
                    "اسم مشغل الماكينة": operator_name,
                    "اسم القائم بالصيانة": technician_name,
                    "تخصص العطل": fault_type,
                    "طبيعة الصيانة": maint_nature,
                    "حالة الماكينة النهائية": final_status,
                    "كود/اسم قطعة الغيار": spare_part_code,
                    "تكلفة قطعة الغيار (جنيه)": spare_part_cost,
                    "وقت البداية": str_t1,
                    "وقت النهاية": t2_input.strftime("%H:%M"),
                    "مدة العطل (ساعة)": duration,
                    "السبب الرئيسي": cause_cat,
                    "عطل صيانة": maint_h,
                    "تلف سلندر حفر": cyl_h,
                    "تأخير مشتريات": proc_h,
                    "صيانة مخططة": pm_h,
                    "خطأ مشغل": op_h,
                    "الساعات التشغيلية المتاحة": free_hrs,
                    "توافرية الصيانة (%)": round(availability, 2),
                    "ملاحظات": notes
                }
                save_data(pd.concat([df_curr, pd.DataFrame([new_row])], ignore_index=True))
                st.session_state["last_notification_number"] = notification_number
                st.success(f"تم الحفظ بنجاح! رقم الإخطار: {notification_number} | مدة العطل: {duration} ساعة | نسبة التوافرية: {availability:.2f}%")
                st.rerun()
            except Exception as e:
                st.error(f"حدث خطأ أثناء إدخال البيانات: {e}")

def filter_by_date_range(df, key_prefix=""):
    if df.empty or "التاريخ_dt" not in df.columns:
        return df
    
    st.sidebar.header("🗓️ تحديد الفترة الزمنية للتقارير")
    filter_type = st.sidebar.radio(
        "طريقة الفلترة:",
        ["شهري", "فترة مخصصة (من - إلى)", "الكل"],
        key=f"{key_prefix}_filter_type"
    )
    
    if filter_type == "شهري":
        df['سنة_شهر'] = df['التاريخ_dt'].dt.to_period('M')
        available_months = sorted(df['سنة_شهر'].dropna().unique().astype(str), reverse=True)
        if available_months:
            selected_month = st.sidebar.selectbox(
                "اختر الشهر/السنة:",
                available_months,
                key=f"{key_prefix}_selected_month"
            )
            return df[df['سنة_شهر'].astype(str) == selected_month]
    elif filter_type == "فترة مخصصة (من - إلى)":
        min_date = df['التاريخ_dt'].min().date() if not df['التاريخ_dt'].dropna().empty else datetime.now().date()
        max_date = df['التاريخ_dt'].max().date() if not df['التاريخ_dt'].dropna().empty else datetime.now().date()
        start_date = st.sidebar.date_input("من تاريخ:", min_date, key=f"{key_prefix}_start_date")
        end_date = st.sidebar.date_input("إلى تاريخ:", max_date, key=f"{key_prefix}_end_date")
        return df[(df['التاريخ_dt'].dt.date >= start_date) & (df['التاريخ_dt'].dt.date <= end_date)]
    
    return df

df_filtered_shared = filter_by_date_range(load_data(), key_prefix="shared")

with tab2:
    st.subheader("📊 البيانات المفلترة حسب الفترة المحددة")
    st.dataframe(df_filtered_shared.drop(columns=["التاريخ_dt", "سنة_شهر"], errors="ignore"), use_container_width=True)

    st.divider()
    st.subheader("💾 نسخ احتياطي واستعادة قاعدة البيانات")

    col_bk1, col_bk2 = st.columns(2)

    with col_bk1:
        st.markdown("**⬇️ تنزيل نسخة احتياطية**")
        if os.path.exists(EXCEL_FILE):
            with open(EXCEL_FILE, "rb") as f:
                backup_bytes = f.read()
            backup_filename = f"نسخة_احتياطية_{datetime.now().strftime('%Y-%m-%d_%H%M')}.xlsx"
            st.download_button(
                label="⬇️ تنزيل نسخة احتياطية من قاعدة البيانات",
                data=backup_bytes,
                file_name=backup_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
            st.caption(f"آخر تعديل على الملف: {datetime.fromtimestamp(os.path.getmtime(EXCEL_FILE)).strftime('%Y-%m-%d %H:%M')} | عدد السجلات الحالية: {len(load_data())}")
        else:
            st.info("لا توجد قاعدة بيانات محفوظة بعد لعمل نسخة احتياطية منها.")

    with col_bk2:
        st.markdown("**⬆️ استعادة نسخة احتياطية**")
        restore_file = st.file_uploader("اختر ملف Excel (.xlsx) للاستعادة", type=["xlsx"], key="restore_uploader")
        if restore_file is not None:
            try:
                preview_restore_df = pd.read_excel(restore_file)
                st.write(f"عدد السجلات في الملف المرفوع: **{len(preview_restore_df)}** سجل")
                st.dataframe(preview_restore_df.head(5), use_container_width=True)

                confirm_restore = st.checkbox("⚠️ أؤكد استبدال قاعدة البيانات الحالية بالكامل بهذا الملف", key="confirm_restore_checkbox")
                if st.button("♻️ استعادة النسخة الاحتياطية الآن", type="primary", disabled=not confirm_restore, use_container_width=True):
                    restore_df_to_save = preview_restore_df.drop(columns=["التاريخ_dt"], errors="ignore")
                    restore_df_to_save.to_excel(EXCEL_FILE, index=False)
                    st.success(f"✅ تم استعادة قاعدة البيانات بنجاح.")
                    st.rerun()
            except Exception as e:
                st.error(f"حدث خطأ أثناء قراءة الملف: {e}")

    st.divider()

    if not df_filtered_shared.empty:
        col_chart1, col_chart2, col_chart3 = st.columns(3)
        with col_chart1:
            st.subheader("ساعات التوقف حسب المصنع")
            st.bar_chart(df_filtered_shared.groupby("المصنع/القسم")["مدة العطل (ساعة)"].sum())
        with col_chart2:
            st.subheader("ساعات التوقف حسب تخصص العطل")
            st.bar_chart(df_filtered_shared.groupby("تخصص العطل")["مدة العطل (ساعة)"].sum())
        with col_chart3:
            st.subheader("تكلفة قطع الغيار حسب القسم (جنيه)")
            st.bar_chart(df_filtered_shared.groupby("المصنع/القسم")["تكلفة قطعة الغيار (جنيه)"].sum())

with tab3:
    st.header("📈 مؤشرات الأداء الهندسية والاعتمادية ونظام البونص والجزاءات")
    df_kpi = df_filtered_shared
    
    if not df_kpi.empty:
        breakdown_df = df_kpi[df_kpi["السبب الرئيسي"] == "عطل صيانة"]
        cyl_df = df_kpi[df_kpi["السبب الرئيسي"] == "تلف / عيب سلندر حفر"]
        
        total_downtime = breakdown_df["مدة العطل (ساعة)"].sum()
        total_cyl_downtime = cyl_df["مدة العطل (ساعة)"].sum()
        total_failures = len(breakdown_df)
        
        mech_downtime = breakdown_df[breakdown_df["تخصص العطل"] == "ميكانيكا"]["مدة العطل (ساعة)"].sum()
        elec_downtime = breakdown_df[breakdown_df["تخصص العطل"].isin(["كهرباء", "تحكم وآليات PLC", "كروت إلكترونية"])]["مدة العطل (ساعة)"].sum()
        
        unique_shifts_df = df_kpi.drop_duplicates(subset=["التاريخ", "المصنع/القسم", "رقم الوردية"])
        total_operating_hrs = unique_shifts_df["الساعات التشغيلية المتاحة"].sum()
        
        actual_uptime = max(total_operating_hrs - total_downtime, 0.0)
        total_cost = df_kpi["تكلفة قطعة الغيار (جنيه)"].sum()
        
        mttr = round(total_downtime / total_failures, 2) if total_failures > 0 else 0.0
        mtbf = round(actual_uptime / total_failures, 2) if total_failures > 0 else 0.0
        overall_avail = round((actual_uptime / total_operating_hrs) * 100, 2) if total_operating_hrs > 0 else 100.0

        col_m1, col_m2, col_m3, col_m4, col_m5 = st.columns(5)
        col_m1.metric("معدل وقت الإصلاح (MTTR)", f"{mttr} ساعة")
        col_m2.metric("معدل التشغيل بين الأعطال (MTBF)", f"{mtbf} ساعة")
        col_m3.metric("عدد أعطال الصيانة", f"{total_failures} عطل")
        col_m4.metric("التوافرية الإجمالية", f"{overall_avail}%")
        col_m5.metric("توقفات تلف السلندرات", f"{total_cyl_downtime:.2f} ساعة")

        st.divider()

        st.subheader("🎯 عداد رصيد البونص والجزاءات (حد 60 ساعة مسموحة للصيانة)")
        
        ALLOWED_BONUS_HOURS = 60.0
        remaining_bonus = ALLOWED_BONUS_HOURS - total_downtime
        
        b_col1, b_col2, b_col3 = st.columns(3)
        with b_col1:
            st.metric("إجمالي أعطال الصيانة التراكمية", f"{total_downtime:.2f} ساعة")
        with b_col2:
            st.metric("الحد المسموح (البونص)", f"{ALLOWED_BONUS_HOURS} ساعة")
        with b_col3:
            if remaining_bonus >= 0:
                st.metric("رصيد البونص المتبقي", f"{remaining_bonus:.2f} ساعة", delta=f"{remaining_bonus:.2f} ساعة متبقية", delta_color="normal")
            else:
                st.metric("ساعات التجاوز (الجزاءات)", f"{abs(remaining_bonus):.2f} ساعة", delta=f"-{abs(remaining_bonus):.2f} ساعة تجاوز", delta_color="inverse")

        st.divider()
        
        st.markdown("##### ⚡ 🛠️ التنبيهات المنفصلة حسب التخصص والسلندرات:")
        col_warn_m, col_warn_e, col_warn_c = st.columns(3)
        
        with col_warn_m:
            st.markdown("**قسم الميكانيكا**")
            st.metric("توقفات الميكانيكا", f"{mech_downtime:.2f} ساعة")
            if mech_downtime > 35.0:
                st.error(f"⚠️ **تنبيه ميكانيكا:** ارتفاع ملحوظ في الأعطال الميكانيكية.")
            else:
                st.success(f"✅ **الميكانيكا مستقرة:** {mech_downtime:.2f} ساعة.")

        with col_warn_e:
            st.markdown("**قسم الكهرباء والتحكم**")
            st.metric("توقفات الكهرباء والـ PLC", f"{elec_downtime:.2f} ساعة")
            if elec_downtime > 25.0:
                st.error(f"⚠️ **تنبيه كهرباء وتحكم:** ارتفاع الأعطال الكهربائية.")
            else:
                st.success(f"✅ **الكهرباء مستقرة:** {elec_downtime:.2f} ساعة.")

        with col_warn_c:
            st.markdown("**قسم السلندرات والحفر**")
            st.metric("توقفات السلندرات", f"{total_cyl_downtime:.2f} ساعة")
            if total_cyl_downtime > 15.0:
                st.error(f"⚠️ **تنبيه قسم السلندرات:** إجمالي توقفات تلف الحفر بلغت {total_cyl_downtime:.2f} ساعة. يلزم مراجعة اختبار الكروم والتلميع Finish MC.")
            else:
                st.success(f"✅ **السلندرات مستقرة:** {total_cyl_downtime:.2f} ساعة.")

        st.divider()
        st.subheader("🎯 تحليل باريتو للأعطال للفترة المحددة (Pareto 80/20)")
        pareto_df = df_kpi.groupby("رقم/اسم الماكينة")["مدة العطل (ساعة)"].sum().reset_index()
        pareto_df = pareto_df.sort_values(by="مدة العطل (ساعة)", ascending=False)
        st.bar_chart(pareto_df.set_index("رقم/اسم الماكينة"))
    else:
        st.info("لا توجد بيانات مسجلة في الفترة الزمنية المحددة.")
