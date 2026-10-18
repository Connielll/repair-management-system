import streamlit as st
import pandas as pd
import io
from datetime import datetime
from xlsxwriter.utility import xl_col_to_name

# 設定頁面資訊
st.set_page_config(page_title="維修案件管理系統 (全自動整理版)", layout="wide")
st.title("🛠️ 維修案件管理系統 - 全自動整理版")

st.markdown("""
### 系統修正與優化說明：
1. **以今日檔案為準**：確保今日原始檔的所有案件都會完整顯示，不被歷史檔過濾。
2. **精確排序**：維持 181 個標準欄位順序。
3. **歷史備註帶入**：自動從歷史檔提取對應的備註 (Account, Status_update, Remark)。
4. **自動狀態更新 (精準空白修正版)**：
   * 當 `WorkFlow` 為 **Back** 且 `Back_Receive_DT` **確實有日期內容**時，若 `Status_update` 為空，自動填入「`今日日期+已接收`」。*(若日期欄為空白，則絕對保持空白)*
   * 當 `WorkFlow` 為 **Regular** 且 `Receive_DT` **確實有日期內容**時，若 `Status_update` 為空，自動填入「`今日日期+已接收`」。*(若日期欄為空白，則絕對保持空白)*
""")

# 側邊欄：檔案上傳
with st.sidebar:
    st.header("📂 檔案上傳區")
    history_file = st.file_uploader("1. (選填) 上傳歷史總表", type=['xlsx'], key="history")
    daily_file = st.file_uploader("2. (必填) 上傳今日原始檔案", type=['xlsx', 'csv'], key="daily")


def preprocess_raw_data(df):
    """181 欄位精確排序清單"""
    priority_columns = [
        "RMA_No", "Product_Name", "Barcode", "Account", "Status_update", "Remark", "WholeSet_SN", "WholeSet_PN", "Analysis", "Second_Repair",
        "Warranty_DT", "Warranty_Type", "Quotation_Price", "TAT_Now", "WorkFlow", "Internal_Repair_Status", "External_Repair_Status", "Complete", "Problem_Description", "Repair_Comment",
        "Back_Repair_Comment", "Back_Complete", "Bill_Company", "Bill_ID", "Front_RC", "Back_RC", "RMA_Type", "Category_Type", "Source", "Org_RMANO",
        "Customer_Ref", "Back_Warranty_Type", "Order_DT", "RMS", "BackRMS", "Customer_Send_Date", "Receive_DT", "Repair_DT", "Pre_Ship_DT", "QC_Complete_DT",
        "temperature", "humidity", "No_Charge_Reason", "Back_temperature", "Back_humidity", "Back_No_Charge_Reason", "Back_Customer_Send_Date", "Back_Receive_DT", "Back_Repair_DT", "Back_Pre_Ship_DT",
        "Back_QC_Complete_DT", "Back_Ship_DT", "Back_Billing_DT", "Ship_DT", "BillingDT", "Quotation_DT", "Quotation_SignBack_DT", "Customer_Receive_DT", "Back_Customer_Receive_DT", "BACK_TAT_Now",
        "Failure_Code", "Failure_Description", "Repair_Minutes", "I_Repair_Comment", "Final_Test_Remark", "ship_Company", "Ship_ID", "Product_Type", "Global_Service", "Order_BY",
        "RBU", "Grade", "Receive_BY", "Repair_BY", "Ship_by", "Check_BY", "Back_Receive_BY", "Back_Repair_BY", "Back_Ship_By", "Back_Check_BY",
        "Responsible", "Second_Repair_Reason", "Supplier_Remark", "Receive_Remark", "Forwarder", "Tracking_Code", "Curr", "Service_Charge", "billing", "Back_Category_Type",
        "Back_Product_Type", "Back_Source", "Back_Internal_Repair_Status", "Back_External_Repair_status", "Back_Failure_Code", "Back_Failure_Description", "Back_Repair_Minutes", "Back_I_Repair_Comment", "Back_Final_Test_Remark", "Back_Responsible",
        "Back_Second_Repair", "Back_Second_Repair_Reason", "Back_Supplier_Remark", "Back_Receive_Remark", "Back_Forwarder", "Back_Tracking_Code", "Back_Curr", "Back_Service_Charge", "Back_billing", "Quality_Alert",
        "Back_Quality_Alert", "Cost_Waive", "Cost_Waive_Text", "Cost_Sales_Name", "Cost_Amount", "Cost_Waive_Remark", "Back_Cost_Waive", "Back_Cost_Waive_Text", "Back_Cost_Sales_Name", "Back_Cost_Sales_Amount",
        "Back_Cost_Waive_Remark", "SAP_Cost_Currency", "SAP_Labor_Cost", "SAP_Material_Cost", "Back_SAP_Cost_Currency", "Back_SAP_Labor_Cost", "Back_SAP_Material_Cost", "SubCon", "Back_SubCon", "Quotation_TAT",
        "Sign_TAT", "MES_Shipping_DT", "Second_Repair_Responsible", "Back_Second_Repair_Responsible", "BGA_Rework", "Back_BGA_Rework", "AASC_Invoice", "PO_NO", "Cost_Center", "Back_Cost_Center",
        "Country", "Ship_Contact", "Ship_Address", "Shipping_Instr", "Ship_Contact_Phone", "Rec_Pending_Remark", "Rec_Back_Pending_Remark", "AGS_PartNo", "Back_AGS_PartNo", "Estimated_Date",
        "Delay_Reason_Code", "Delay_Reason_Remark", "Back_Estimated_Date", "Back_Delay_Reason_Code", "Back_Delay_Reason_Remark", "Cell_Phone", "SMS_Flag", "Receive_Product_Type", "Back_Receive_Product_Type", "Cust_SN",
        "Cust_PN", "Rework_DT", "Rework_by", "Rework_Check_by", "Hard_Repair_Back_Reason", "Hard_Repair_Back_Remark", "Contact_email_addr", "Pending_PO_TAT", "Receive_Forwarder", "Receive_Tracking_Code",
        "Back_Receive_Forwarder", "Back_Receive_Tracking_Code", "Back_Reason", "SalesInfo", "Invoice_No", "Working_Date_Type", "Back_Working_Date_Type", "Ship_Remark", "Back_Ship_Remark", "Customer_PO",
        "One_Shipment", "SAP_END_Customer_ID", "SAP_END_Customer_Name", "CA", "PA", "BACK_CA", "BACK_PA", "Refurbished_Cust_PN", "QFS_No", "Internal_Warranty_type",
        "Back_Internal_Warranty_Type", "Is_Expedited_Repair"
    ]

    for col in ["Account", "Status_update", "Remark"]:
        if col not in df.columns:
            df[col] = ""

    existing_priority_cols = [c for c in priority_columns if c in df.columns]
    remaining_cols = [c for c in df.columns if c not in priority_columns]
    final_cols = existing_priority_cols + remaining_cols
    df = df[final_cols]

    if 'TAT_Now' in df.columns:
        df['TAT_Now'] = pd.to_numeric(df['TAT_Now'], errors='coerce')
        df = df.sort_values(by='TAT_Now', ascending=True)

    return df


def build_history_map(history_file):
    """建立歷史資料 RMA 對應表"""
    h_map = {}
    try:
        xls = pd.ExcelFile(history_file)
        for sheet in xls.sheet_names:
            df_h = pd.read_excel(xls, sheet_name=sheet)
            if 'RMA_No' in df_h.columns:
                df_h['RMA_No'] = df_h['RMA_No'].astype(str).str.strip()
                for _, row in df_h.iterrows():
                    rma = str(row['RMA_No'])
                    h_map[rma] = {
                        'Account': row.get('Account', ''),
                        'Status_update': row.get('Status_update', ''),
                        'Remark': row.get('Remark', '')
                    }
    except Exception as e:
        st.error(f"歷史檔讀取錯誤: {e}")
    return h_map


if daily_file is not None:
    h_map = {}
    if history_file is not None:
        with st.spinner("讀取歷史備註中..."):
            h_map = build_history_map(history_file)

    if daily_file.name.endswith('.csv'):
        df_today = pd.read_csv(daily_file)
    else:
        df_today = pd.read_excel(daily_file)

    df_main = preprocess_raw_data(df_today)

    # 比對並填入歷史備註
    if h_map:
        match_count = 0
        df_main['RMA_No'] = df_main['RMA_No'].astype(str).str.strip()
        for idx, row in df_main.iterrows():
            rma = row['RMA_No']
            if rma in h_map:
                df_main.at[idx, 'Account'] = h_map[rma]['Account']
                df_main.at[idx, 'Status_update'] = h_map[rma]['Status_update']
                df_main.at[idx, 'Remark'] = h_map[rma]['Remark']
                match_count += 1
        if match_count > 0:
            st.toast(f"✅ 成功比對並帶入 {match_count} 筆歷史備註！")

    # 調整後的自動填寫狀態邏輯 (嚴格排除所有空白形式)
    today_str = f"{datetime.now().month}/{datetime.now().day}"
    auto_status_text = f"{today_str}已接收"
    auto_fill_count = 0
    check_cols = ['WorkFlow', 'Status_update']

    if all(c in df_main.columns for c in check_cols):
        for idx, row in df_main.iterrows():
            wf = str(row['WorkFlow']).strip()
            su = str(row['Status_update']).strip()

            # 先確認目前 Status_update 是否為真正的空白
            if su == "" or su == "nan" or pd.isna(row['Status_update']):
                # 條件一：Back
                if wf == "Back" and 'Back_Receive_DT' in df_main.columns:
                    val_br_dt = row['Back_Receive_DT']
                    str_br_dt = str(val_br_dt).strip()
                    # 排除各式各樣的空白 (pandas 空值、空字串、nan字串、NaT字串)
                    if not pd.isna(val_br_dt) and str_br_dt != "" and str_br_dt != "nan" and str_br_dt != "NaT":
                        df_main.at[idx, 'Status_update'] = auto_status_text
                        auto_fill_count += 1
                    else:
                        df_main.at[idx, 'Status_update'] = ""  # 強制保持空白
                # 條件二：Regular (支援大小寫混用處理)
                elif (wf == "Regular" or wf == "regular") and 'Receive_DT' in df_main.columns:
                    val_r_dt = row['Receive_DT']
                    str_r_dt = str(val_r_dt).strip()
                    # 排除各式各樣的空白
                    if not pd.isna(val_r_dt) and str_r_dt != "" and str_r_dt != "nan" and str_r_dt != "NaT":
                        df_main.at[idx, 'Status_update'] = auto_status_text
                        auto_fill_count += 1
                    else:
                        df_main.at[idx, 'Status_update'] = ""  # 強制保持空白
            else:
                # 如果原本歷史檔帶入的內容是 "nan" 字串，順便清乾淨
                if su == "nan":
                    df_main.at[idx, 'Status_update'] = ""

        if auto_fill_count > 0:
            st.toast(f"⚡ 偵測到新進案件！已自動為 {auto_fill_count} 筆案件填入「{auto_status_text}」")

    today_sheet = datetime.now().strftime("%Y-%m-%d")
    st.subheader(f"📋 今日整理結果 (共 {len(df_main)} 筆案件)")
    edited_df = st.data_editor(df_main, num_rows="dynamic", use_container_width=True, height=600, key="editor", hide_index=True)

    # --- 產出 Excel ---
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook = writer.book
        cell_fmt = workbook.add_format({'font_name': 'Microsoft JhengHei', 'font_size': 11, 'valign': 'top'})
        head_fmt = workbook.add_format({'font_name': 'Microsoft JhengHei', 'font_size': 11, 'bold': True, 'valign': 'top', 'fg_color': '#D7E4BC', 'border': 1})
        red_fmt = workbook.add_format({'font_name': 'Microsoft JhengHei', 'font_size': 11, 'font_color': 'red', 'bold': True})

        edited_df.to_excel(writer, index=False, sheet_name=today_sheet, startrow=1, header=False)
        worksheet = writer.sheets[today_sheet]

        for col_num, value in enumerate(edited_df.columns.values):
            worksheet.write(0, col_num, value, head_fmt)

        worksheet.set_column(0, len(edited_df.columns) - 1, None, cell_fmt)

        try:
            col_map = {name: i for i, name in enumerate(edited_df.columns)}
            if all(k in col_map for k in ['Product_Name', 'Barcode', 'Warranty_Type']):
                cp = xl_col_to_name(col_map['Product_Name'])
                cb = xl_col_to_name(col_map['Barcode'])
                cw = xl_col_to_name(col_map['Warranty_Type'])
                formula = f'=OR(${cp}2="", ${cb}2="", ${cw}2="")'
                worksheet.conditional_format(1, 0, len(edited_df), 0, {'type': 'formula', 'criteria': formula, 'format': red_fmt})
        except Exception:
            pass

        for i, col in enumerate(edited_df.columns):
            c_len = max(len(str(col)) + 4, 10)
            if len(edited_df) > 0:
                v_len = len(str(edited_df.iloc[0, i]))
                c_len = max(c_len, v_len)
            worksheet.set_column(i, i, min(c_len, 60), cell_fmt)

    output.seek(0)
    st.download_button(
        label="📥 下載整理後的 Excel",
        data=output,
        file_name=f"維修總表_{datetime.now().strftime('%Y%m%d')}.xlsx"
    )
elif daily_file is None:
    st.info("ℹ️ 請上傳檔案開始作業。")
