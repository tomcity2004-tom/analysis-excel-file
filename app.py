import streamlit as st
import pandas as pd
import numpy as np
import re
from datetime import datetime
import io

st.set_page_config(page_title="Excel 資料完整性整理工具", layout="wide")

st.title("Excel 資料完整性整理工具")
st.markdown("上傳 Excel 檔，自動將資料分為「完整」與「不完整」，並產生統計摘要。")

# =============================================
# 輔助函數（與原本邏輯相同）
# =============================================

def is_missing(value):
    if pd.isna(value):
        return True
    s = str(value).strip().lower()
    missing_keywords = [
        '', 'null', 'none', 'n/a', 'na', 'nil', 'unknown',
        'not available', '#na', '#n/a', '-', '---', 'nan'
    ]
    return s in missing_keywords

def is_invalid_date(value):
    if is_missing(value):
        return True
    s = str(value).strip()
    if any(x in s for x in ['00-00', '00/00', '0000', 'invalid']):
        return True
    formats = [
        '%Y-%m-%d', '%Y/%m/%d', '%Y.%m.%d',
        '%d/%m/%Y', '%d-%m-%Y', '%d.%m.%Y',
        '%m/%d/%Y', '%Y%m%d'
    ]
    for fmt in formats:
        try:
            datetime.strptime(s, fmt)
            return False
        except:
            continue
    return True

def is_invalid_phone(value):
    if is_missing(value):
        return True
    s = str(value).strip()
    cleaned = re.sub(r'[^\d+]', '', s)
    if not (8 <= len(cleaned) <= 15):
        return True
    if re.search(r'[a-zA-Z]', s):
        return True
    return False

def is_invalid_email(value):
    if is_missing(value):
        return True
    s = str(value).strip()
    pattern = r'^[\w\.-]+@[\w\.-]+\.\w+$'
    return not bool(re.match(pattern, s))

def is_incomplete_row(row):
    critical_cols = [
        'First_Name', 'Last_Name', 'Email_Address', 'Phone_Number',
        'DOB', 'Department', 'Employment_Status', 'Salary_Monthly'
    ]
    for col in critical_cols:
        if col in row.index and is_missing(row[col]):
            return True

    date_cols = ['DOB', 'Hire_Date', 'Contract_End_Date', 'Registration_Date', 'Last_Purchase_Date']
    for col in date_cols:
        if col in row.index and is_invalid_date(row[col]):
            return True

    phone_cols = ['Phone_Number', 'Emergency_Contact_Phone']
    for col in phone_cols:
        if col in row.index and is_invalid_phone(row[col]):
            return True

    if 'Email_Address' in row.index and is_invalid_email(row['Email_Address']):
        return True

    if 'Age' in row.index:
        age = row['Age']
        if not is_missing(age):
            try:
                age_num = float(str(age).strip())
                if age_num < 15 or age_num > 80:
                    return True
            except:
                return True

    if 'Salary_Monthly' in row.index:
        salary = row['Salary_Monthly']
        if not is_missing(salary):
            try:
                sal_num = float(str(salary).replace(',', '').replace('$', ''))
                if sal_num < 0 or sal_num > 200000:
                    return True
            except:
                return True

    abnormal_values = ['special chars', 'test row', 'duplicate suspected', 'bad', 'invalid_mac']
    for col in row.index:
        val = str(row[col]).lower()
        if any(ab in val for ab in abnormal_values):
            return True

    return False

def get_value_counts_summary(series, top_n=10):
    if series.dropna().empty:
        return "無資料"
    vc = series.fillna('(空白/缺失)').astype(str).str.strip().value_counts().head(top_n)
    return '; '.join([f"{k}: {v}" for k, v in vc.items()])

# =============================================
# 主介面
# =============================================

uploaded_file = st.file_uploader("請上傳 Excel 檔案（.xlsx）", type=["xlsx"])

if uploaded_file is not None:
    try:
        df = pd.read_excel(uploaded_file)
        st.success(f"成功讀取檔案，共 {len(df)} 筆資料、{len(df.columns)} 個欄位")

        with st.spinner("正在分析資料完整性..."):
            df['Is_Incomplete'] = df.apply(is_incomplete_row, axis=1)

            df_complete = df[df['Is_Incomplete'] == False].drop(columns=['Is_Incomplete'])
            df_incomplete = df[df['Is_Incomplete'] == True].drop(columns=['Is_Incomplete'])

            total = len(df)
            complete_count = len(df_complete)
            incomplete_count = len(df_incomplete)
            complete_pct = round(complete_count / total * 100, 2) if total > 0 else 0
            incomplete_pct = round(incomplete_count / total * 100, 2) if total > 0 else 0

        # 顯示統計數字
        col1, col2, col3 = st.columns(3)
        col1.metric("總筆數", total)
        col2.metric("完整資料", f"{complete_count} 筆", f"{complete_pct}%")
        col3.metric("不完整資料", f"{incomplete_count} 筆", f"{incomplete_pct}%")

        # 分頁顯示資料
        tab1, tab2, tab3 = st.tabs(["完整資料", "不完整資料", "統計摘要"])

        with tab1:
            st.dataframe(df_complete, use_container_width=True)

        with tab2:
            st.dataframe(df_incomplete, use_container_width=True)

        with tab3:
            # 整體摘要
            summary_overall = pd.DataFrame({
                '項目': ['總筆數', '完整資料筆數', '不完整資料筆數', '完整資料佔比(%)', '不完整資料佔比(%)'],
                '數值': [total, complete_count, incomplete_count, complete_pct, incomplete_pct]
            })
            st.subheader("整體摘要")
            st.dataframe(summary_overall, use_container_width=True)

            # 缺失值統計
            missing_stats = []
            for col in df.columns:
                if col == 'Is_Incomplete':
                    continue
                missing_count = df[col].apply(is_missing).sum()
                missing_pct = round(missing_count / total * 100, 2) if total > 0 else 0
                missing_stats.append({
                    '欄位名稱': col,
                    '缺失筆數': missing_count,
                    '缺失率(%)': missing_pct
                })
            df_missing_stats = pd.DataFrame(missing_stats).sort_values('缺失筆數', ascending=False)
            st.subheader("各欄位缺失值統計")
            st.dataframe(df_missing_stats, use_container_width=True)

            # 分類欄位分佈
            categorical_cols = ['Department', 'Employment_Status', 'Gender', 'Currency', 'Education_Level', 'VIP_Tier']
            cat_summary_rows = []
            for col in categorical_cols:
                if col not in df.columns:
                    continue
                cat_summary_rows.append({
                    '欄位': col,
                    '完整資料分佈': get_value_counts_summary(df_complete[col]) if col in df_complete.columns else '',
                    '不完整資料分佈': get_value_counts_summary(df_incomplete[col]) if col in df_incomplete.columns else ''
                })
            df_cat_summary = pd.DataFrame(cat_summary_rows)
            st.subheader("主要分類欄位分佈")
            st.dataframe(df_cat_summary, use_container_width=True)

        # =============================================
        # 產生可下載的 Excel（三個 Sheet）
        # =============================================
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_complete.to_excel(writer, sheet_name='完整資料', index=False)
            df_incomplete.to_excel(writer, sheet_name='不完整資料', index=False)

            # 統計 Sheet
            summary_overall.to_excel(writer, sheet_name='統計', index=False, startrow=0)
            df_missing_stats.to_excel(writer, sheet_name='統計', index=False, startrow=8)
            df_cat_summary.to_excel(writer, sheet_name='統計', index=False, startrow=8 + len(df_missing_stats) + 3)

        output.seek(0)

        st.download_button(
            label="下載整理後的 Excel 檔",
            data=output,
            file_name=f"整理後資料_{uploaded_file.name}",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"處理檔案時發生錯誤：{str(e)}")
else:
    st.info("請先上傳 Excel 檔案。")