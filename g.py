import streamlit as st
import pandas as pd
import os

# 頁面配置
st.set_page_config(page_title="信班英文考卷登記系統", page_icon="📝", layout="wide")

EXCEL_FILE = "scores.xlsx"

# 初始化或讀取 Excel 檔案
def load_data():
    if os.path.exists(EXCEL_FILE):
        try:
            return pd.read_excel(EXCEL_FILE)
        except Exception:
            pass
    # 預設資料格式
    return pd.DataFrame(columns=["座號", "姓名", "考卷名稱", "成績", "登記時間"])

def save_data(df):
    df.to_excel(EXCEL_FILE, index=False)

# 初始化 Session State
if "df" not in st.session_state:
    st.session_state.df = load_data()

st.title("📝 信班英文考卷登記系統")

# 分頁分流
menu = st.sidebar.radio("請選擇功能區塊", ["學生成績登記", "教師審核與管理"])

# ----------------- 1. 學生登記區 -----------------
if menu == "學生成績登記":
    st.header("📋 學生登記成績")
    
    with st.form("score_form", clear_on_submit=True):
        col1, col2 = st.columns(2)
        with col1:
            seat_no = st.number_input("座號", min_value=1, max_value=50, step=1)
            student_name = st.text_input("姓名")
        with col2:
            exam_name = st.text_input("考卷名稱（例如：L1 單字小考）")
            score = st.number_input("分數", min_value=0, max_value=100, step=1)
            
        submitted = st.form_submit_button("送出登記")
        
        if submitted:
            if not student_name or not exam_name:
                st.error("請完整填寫姓名與考卷名稱！")
            else:
                now_str = pd.Timestamp.now().strftime("%Y-%m-%d %H:%M:%S")
                new_row = pd.DataFrame([{
                    "座號": seat_no,
                    "姓名": student_name,
                    "考卷名稱": exam_name,
                    "成績": score,
                    "登記時間": now_str
                }])
                
                # 更新並儲存
                st.session_state.df = pd.concat([st.session_state.df, new_row], ignore_index=True)
                save_data(st.session_state.df)
                st.success(f"✅ {student_name} 的成績已順利登記！")

    st.subheader("📊 已登記歷史紀錄")
    if not st.session_state.df.empty:
        st.dataframe(st.session_state.df, use_container_width=True)
    else:
        st.info("目前尚無登記資料。")

# ----------------- 2. 教師管理區 -----------------
elif menu == "教師審核與管理":
    st.header("🔒 教師審核與資料管理")
    
    pwd = st.text_input("請輸入教師驗證碼", type="password")
    
    if pwd == "0112":
        st.success("驗證成功！歡迎進入管理後台。")
        
        df = st.session_state.df
        
        st.subheader("📑 考卷繳交狀態與缺交提醒")
        if not df.empty:
            exams = df["考卷名稱"].unique().tolist()
            selected_exam = st.selectbox("請選擇要查看的考卷名稱", exams)
            
            if selected_exam:
                exam_df = df[df["考卷名稱"] == selected_exam]
                submitted_seats = set(exam_df["座號"].tolist())
                
                # 假設全班 1 ~ 40 號
                all_seats = set(range(1, 41))
                missing_seats = sorted(list(all_seats - submitted_seats))
                
                st.write(f"### 🎯 【{selected_exam}】成績列表")
                st.dataframe(exam_df, use_container_width=True)
                
                if missing_seats:
                    st.error(f"🚨 **未繳交/未登記座號警告：** {', '.join(map(str, missing_seats))}")
                else:
                    st.success("🎉 本考卷全班已完整登記！")
        else:
            st.info("目前無任何登記紀錄可供審核。")

        st.divider()
        st.subheader("📥 Excel 資料匯出與匯入")
        
        col_exp, col_imp = st.columns(2)
        
        with col_exp:
            if not df.empty:
                # 提供 Excel 下載
                from io import BytesIO
                output = BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df.to_excel(writer, index=False, sheet_name='成績單')
                processed_data = output.getvalue()
                
                st.download_button(
                    label="📥 下載完整成績表 Excel",
                    data=processed_data,
                    file_name="信班英文成績表.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        
        with col_imp:
            uploaded_file = st.file_uploader("上傳並覆蓋成績 Excel 檔", type=["xlsx"])
            if uploaded_file is not None:
                try:
                    imported_df = pd.read_excel(uploaded_file)
                    st.session_state.df = imported_df
                    save_data(imported_df)
                    st.success("✅ 成績資料已由 Excel 覆蓋更新！請重新整理頁面。")
                except Exception as e:
                    st.error(f"讀取 Excel 失敗：{e}")
                    
    elif pwd != "":
        st.error("驗證碼錯誤，請重新輸入！")