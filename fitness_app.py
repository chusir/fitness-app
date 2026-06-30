import streamlit as st
import pandas as pd
import os
import plotly.express as px

# 設定網頁
st.set_page_config(page_title="體適能測試 App", layout="wide")
st.title("🏆 小學體適能測試管理 App (進階版)")

# 總資料庫檔案路徑
MASTER_DB = 'master_fitness_data.xlsx'

# 載入資料函數
def load_data():
    if os.path.exists(MASTER_DB):
        return pd.read_excel(MASTER_DB)
    else:
        # 為了相容你的 Excel 格式，建立包含學年的基礎欄位
        return pd.DataFrame(columns=[
            'Academic_Year', 'Class', 'Name', 'Gender (M/F)', 'Age', 
            '1-min Sit-ups', 'Sit-and-Reach', 'Run/Walk', 'Total Score', 'Certificate'
        ])

# 儲存資料到總表
df = load_data()

# 左側選單
menu = st.sidebar.selectbox("功能選單", [
    "📂 批量匯入班級 Excel", 
    "📝 手動新增單筆資料", 
    "📈 學生歷年趨勢查詢", 
    "📊 全校數據概覽"
])

# ==========================================
# 功能 1：批量匯入班級 Excel (新功能！)
# ==========================================
if menu == "📂 批量匯入班級 Excel":
    st.header("批量匯入班級成績")
    st.info("請上傳各班的體適能 Excel 檔案 (例如: 1A.xlsx, 1B.xlsx)。系統會自動讀取「Primary」分頁的數據。")
    
    # 讓老師設定這批檔案的學年
    current_year = st.text_input("請輸入這批資料的學年 (例如：2025-2026)", value="2025-2026")
    
    # 支援多檔案上傳
    uploaded_files = st.file_uploader("請將 Excel 檔案拖曳至此", type=['xlsx'], accept_multiple_files=True)
    
    if st.button("開始匯入資料"):
        if uploaded_files:
            new_data_frames = []
            
            for file in uploaded_files:
                try:
                    # 只讀取你的 Excel 中的 'Primary' 分頁
                    temp_df = pd.read_excel(file, sheet_name='Primary')
                    
                    # 過濾掉空白行或範例行 (假設真實學生的姓名不能為空，且不等於 'E.g.')
                    temp_df = temp_df[temp_df['Name'].notna()]
                    temp_df = temp_df[temp_df['Name'] != 'E.g.']
                    
                    # 抓取我們需要的關鍵欄位 (依照你提供的 CSV 欄位名稱)
                    # 為了避免格式錯誤，我們使用欄位名稱的部分匹配來抓取
                    clean_df = pd.DataFrame()
                    clean_df['Academic_Year'] = [current_year] * len(temp_df)
                    clean_df['Class'] = temp_df['Class']
                    clean_df['Name'] = temp_df['Name']
                    clean_df['Gender (M/F)'] = temp_df['Gender\n(M/F)'] if 'Gender\n(M/F)' in temp_df.columns else temp_df['Gender (M/F)']
                    clean_df['Age'] = temp_df['Age']
                    
                    # 抓取測驗項目 (對應你的 Excel 欄位)
                    clean_df['1-min Sit-ups'] = temp_df['1-min Sit-ups \n(Repetitions)'].fillna(0)
                    clean_df['Sit-and-Reach'] = temp_df['Sit-and-Reach \n(cm)'].fillna(0)
                    clean_df['Run/Walk'] = temp_df['6/9-min Run/Walk \n(m)'].fillna(0)
                    
                    # 成績與獎項
                    clean_df['Total Score'] = temp_df['Total Score '].fillna(0)
                    clean_df['Certificate'] = temp_df['Certificate'].fillna('None')
                    
                    new_data_frames.append(clean_df)
                    st.success(f"✅ 成功讀取檔案: {file.name}")
                    
                except Exception as e:
                    st.error(f"❌ 讀取 {file.name} 時發生錯誤。請確認檔案內有「Primary」分頁。錯誤細節: {e}")
            
            # 將所有新資料合併到總資料庫中
            if new_data_frames:
                all_new_data = pd.concat(new_data_frames, ignore_index=True)
                df = pd.concat([df, all_new_data], ignore_index=True)
                
                # 儲存回總資料庫 Excel
                df.to_excel(MASTER_DB, index=False)
                st.balloons()
                st.success(f"🎉 匯入完成！共新增了 {len(all_new_data)} 筆學生資料至總資料庫。")
        else:
            st.warning("請先上傳至少一個 Excel 檔案！")

# ==========================================
# 功能 2：手動新增 (保留給個別補測的學生)
# ==========================================
elif menu == "📝 手動新增單筆資料":
    st.header("新增單筆學生成績")
    with st.form("input_form"):
        col1, col2, col3 = st.columns(3)
        with col1:
            academic_year = st.text_input("學年", value="2025-2026")
            s_class = st.text_input("班級", value="1A")
            s_name = st.text_input("學生姓名")
        with col2:
            gender = st.selectbox("性別", ["M", "F"])
            age = st.number_input("年齡", 6, 12, 8)
        with col3:
            situps = st.number_input("仰臥起坐 (次)", 0, 100, 0)
            sitreach = st.number_input("坐前伸 (cm)", 0.0, 50.0, 0.0)
            run = st.number_input("耐力跑 (米)", 0, 3000, 0)
            total_score = st.number_input("總分", 0, 30, 0)
            certificate = st.selectbox("獎項", ["Gold", "Silver", "Bronze", "None"])
            
        submit = st.form_submit_button("儲存成績")
        if submit:
            new_row = {
                'Academic_Year': academic_year, 'Class': s_class, 'Name': s_name, 
                'Gender (M/F)': gender, 'Age': age, '1-min Sit-ups': situps,
                'Sit-and-Reach': sitreach, 'Run/Walk': run, 
                'Total Score': total_score, 'Certificate': certificate
            }
            df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
            df.to_excel(MASTER_DB, index=False)
            st.success(f"已儲存 {s_name} ({academic_year}) 的資料！")

# ==========================================
# 功能 3：學生歷年趨勢查詢 (升級多維度圖表)
# ==========================================
elif menu == "📈 學生歷年趨勢查詢":
    st.header("學生歷年能力變化趨勢")
    
    if df.empty:
        st.warning("資料庫目前是空的，請先匯入資料。")
    else:
        # 搜尋學生
        student_list = df['Name'].dropna().unique()
        search_name = st.selectbox("請選擇或輸入學生姓名:", [""] + list(student_list))
        
        if search_name:
            student_records = df[df['Name'] == search_name].sort_values(by='Academic_Year')
            
            if not student_records.empty:
                st.subheader(f"👤 {search_name} 的歷史紀錄")
                st.dataframe(student_records)
                
                # 選擇要觀看的測試項目圖表
                st.write("---")
                st.subheader("📊 成長曲線圖")
                metric_to_plot = st.radio(
                    "請選擇要觀察的項目：", 
                    ['Total Score', '1-min Sit-ups', 'Sit-and-Reach', 'Run/Walk'],
                    horizontal=True
                )
                
                # 繪製趨勢圖 (X軸為學年，Y軸為選擇的項目)
                fig = px.line(
                    student_records, 
                    x="Academic_Year", 
                    y=metric_to_plot, 
                    markers=True,
                    title=f"{search_name} - {metric_to_plot} 歷年趨勢",
                    labels={"Academic_Year": "學年", metric_to_plot: "成績 / 數值"}
                )
                
                # 美化圖表
                fig.update_traces(line=dict(color='#bef264', width=3), marker=dict(size=10))
                fig.update_layout(template="plotly_dark")
                
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.warning("找不到該學生資料")

# ==========================================
# 功能 4：全校數據概覽
# ==========================================
elif menu == "📊 全校數據概覽":
    st.header("全校體適能總資料庫")
    if not df.empty:
        # 可以按年份篩選
        years = df['Academic_Year'].dropna().unique()
        selected_year = st.selectbox("篩選學年:", ["顯示全部"] + list(years))
        
        display_df = df if selected_year == "顯示全部" else df[df['Academic_Year'] == selected_year]
        
        st.dataframe(display_df, use_container_width=True)
        
        # 獎項分佈圓餅圖
        st.subheader(f"🏆 獎項分佈 ({selected_year})")
        award_counts = display_df['Certificate'].value_counts().reset_index()
        award_counts.columns = ['Certificate', 'Count']
        
        fig_pie = px.pie(award_counts, names='Certificate', values='Count', hole=0.4)
        st.plotly_chart(fig_pie)
    else:
        st.info("目前尚無數據，請至「批量匯入」上傳各班檔案。")