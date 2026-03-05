import streamlit as st
import pandas as pd
import json
import os
from datetime import date

# 嘗試引入 PDF 產生模組，若無則顯示警告
try:
    import pdf_gen_travel
except ImportError:
    pdf_gen_travel = None
    st.warning("⚠️ 找不到 `pdf_gen_travel.py`，匯出 PDF 功能將無法使用。")

def render_new_form():
    st.subheader("📝 新增出差申請與報支單")
    st.caption("帶有 * 號的欄位為必填項目。簽核欄位已預設隱藏，將於印出 PDF 時由主管簽署。")

    # --- 判斷是否為「編輯模式」 ---
    is_editing = False
    edit_id = st.session_state.get('edit_target_id')
    if edit_id:
        st.info(f"✏️ 正在編輯草稿單號：{edit_id}")
        is_editing = True
        # 這裡未來會從資料庫撈取真實資料，目前我們先清空或給預設值以防報錯
        # travel_record = fetch_from_db(edit_id)

    # --- 區塊 1：基本資料 ---
    st.markdown("##### 👤 基本資料與行程")
    col1, col2, col3 = st.columns(3)
    with col1:
        traveler = st.text_input("出差人 *", placeholder="請輸入姓名")
    with col2:
        project_id = st.text_input("計畫編號")
    with col3:
        budget_source = st.text_input("預算來源")

    col4, col5 = st.columns([2, 1])
    with col4:
        purpose = st.text_input("出差事由 *")
    with col5:
        location = st.text_input("出差地點 *")

    col6, col7 = st.columns(2)
    with col6:
        start_date = st.date_input("出差起始日期", value=date.today())
    with col7:
        end_date = st.date_input("出差結束日期", value=date.today())

    st.divider()

    # --- 區塊 2：交通方式 (動態顯示) ---
    st.markdown("##### 🚗 交通方式")
    transport_options = st.multiselect(
        "請選擇交通工具 (可多選)", 
        ["公務車", "計程車", "私車公用", "高鐵", "飛機", "派車", "其他"]
    )

    private_car_km, private_car_plate = 0.0, ""
    official_car_plate, other_transport = "", ""

    if "私車公用" in transport_options:
        c1, c2 = st.columns(2)
        with c1:
            private_car_km = st.number_input("私車公里數", min_value=0.0, step=1.0)
        with c2:
            private_car_plate = st.text_input("私車車號")
            
    if "公務車" in transport_options:
        official_car_plate = st.text_input("公務車車號")
        
    if "其他" in transport_options:
        other_transport = st.text_input("其他交通工具說明")

    st.divider()

    # --- 區塊 3：差旅費明細 (可編輯表格) ---
    st.markdown("##### 💰 差旅費報支單明細")
    if 'travel_expenses' not in st.session_state:
        st.session_state.travel_expenses = pd.DataFrame([
            {"日期": str(date.today()), "起訖地點": "", "車別": "", "交通費": 0, "膳雜費": 0, "住宿費": 0, "其它": 0, "單據編號": ""}
        ])

    edited_df = st.data_editor(
        st.session_state.travel_expenses,
        num_rows="dynamic",
        use_container_width=True,
        hide_index=True,
        column_config={
            "交通費": st.column_config.NumberColumn("交通費", min_value=0, step=1),
            "膳雜費": st.column_config.NumberColumn("膳雜費", min_value=0, step=1),
            "住宿費": st.column_config.NumberColumn("住宿費", min_value=0, step=1),
            "其它": st.column_config.NumberColumn("其它", min_value=0, step=1),
        }
    )

    total_expense = edited_df["交通費"].sum() + edited_df["膳雜費"].sum() + edited_df["住宿費"].sum() + edited_df["其它"].sum()
    st.markdown(f"<h4 style='text-align: right; color: #E65100;'>總計新台幣： {total_expense:,} 元</h4>", unsafe_allow_html=True)

    st.divider()

    # --- 區塊 4：操作按鈕 ---
    b1, b2, b3, b4 = st.columns(4)
    
    if b1.button("💾 儲存草稿", use_container_width=True):
        st.success("草稿已儲存！(開發中)")
        # 儲存後可以選擇清除 edit_target_id
        if 'edit_target_id' in st.session_state:
            del st.session_state.edit_target_id
        
    if b2.button("🚀 送出", type="primary", use_container_width=True):
        if not traveler or not purpose or not location:
            st.error("⚠️ 請確認已填寫所有必填欄位（出差人、事由、地點）")
        else:
            expense_json = edited_df.to_json(orient="records", force_ascii=False)
            # TODO: 儲存至資料庫的邏輯放這裡
            st.success("✅ 表單已成功送出！")
            
            # 清除編輯狀態
            if 'edit_target_id' in st.session_state:
                del st.session_state.edit_target_id
                
            st.session_state.current_view = 'submitted_list'
            st.rerun()

    # PDF 下載區塊處理
    with b3:
        if st.button("📥 產生 PDF", use_container_width=True):
            if not traveler:
                st.error("⚠️ 請先填寫出差人才能產生 PDF。")
            elif pdf_gen_travel is None:
                st.error("⚠️ 找不到 pdf_gen_travel 模組。")
            else:
                expense_json = edited_df.to_json(orient="records", force_ascii=False)
                record = {
                    "traveler": traveler,
                    "project_id": project_id,
                    "purpose": purpose,
                    "location": location,
                    "start_date": str(start_date),
                    "end_date": str(end_date),
                    "transport_options": transport_options,
                    "expense_json": expense_json,
                    "total_expense": total_expense
                }
                
                os.makedirs("output", exist_ok=True)
                template_path = "templates/travel_bg.pdf" # 確保你的資料夾有這個底圖
                output_path = f"output/travel_{traveler}.pdf"
                
                if not os.path.exists(template_path):
                    st.error(f"❌ 找不到底圖檔案: `{template_path}`，請確認 templates 資料夾是否存在。")
                else:
                    success = pdf_gen_travel.generate_pdf_travel(record, template_path, output_path)
                    if success:
                        st.session_state.pdf_ready_path = output_path
                        st.success("🎉 PDF 產生成功！請點擊下方按鈕下載。")
                    else:
                        st.error("❌ 產生 PDF 失敗。")
                        
        # 當 PDF 產生成功時，顯示 Streamlit 原生下載按鈕
        if st.session_state.get('pdf_ready_path'):
            if os.path.exists(st.session_state.pdf_ready_path):
                with open(st.session_state.pdf_ready_path, "rb") as file:
                    st.download_button(
                        label="⬇️ 點此下載 PDF",
                        data=file,
                        file_name=os.path.basename(st.session_state.pdf_ready_path),
                        mime="application/pdf",
                        use_container_width=True
                    )

    if b4.button("📋 查看送出列表", use_container_width=True):
        st.session_state.current_view = 'submitted_list'
        st.rerun()


def render_draft_list():
    st.subheader("📄 草稿列表")
    st.caption("在此繼續編輯您的草稿，或是直接送出/刪除。")

    mock_drafts = [
        {"id": "DR-TRV-001", "date": "2026-03-01", "traveler": "王小明", "purpose": "台北總公司開會", "amount": 2500},
        {"id": "DR-TRV-002", "date": "2026-03-02", "traveler": "李大華", "purpose": "台中廠區視察", "amount": 1200}
    ]

    if not mock_drafts:
        st.info("目前沒有任何草稿。")
        return

    st.markdown("---")
    h_col1, h_col2, h_col3, h_col4, h_col5, h_col6 = st.columns([1.5, 1.5, 1.5, 3, 1.5, 3])
    h_col1.markdown("**單號**"); h_col2.markdown("**建立日期**"); h_col3.markdown("**出差人**")
    h_col4.markdown("**出差事由**"); h_col5.markdown("**預估金額**"); h_col6.markdown("**操作**")
    st.markdown("---")

    for item in mock_drafts:
        col1, col2, col3, col4, col5, col6 = st.columns([1.5, 1.5, 1.5, 3, 1.5, 3])
        col1.write(item["id"])
        col2.write(item["date"])
        col3.write(item["traveler"])
        col4.write(item["purpose"])
        col5.write(f"${item['amount']:,}")
        
        with col6:
            b1, b2, b3 = st.columns(3)
            if b1.button("✏️ 編輯", key=f"trv_edit_{item['id']}", use_container_width=True):
                st.session_state.edit_target_id = item["id"]
                st.session_state.current_view = 'new_form'
                st.rerun()
                
            if b2.button("🚀 送出", key=f"trv_submit_{item['id']}", type="primary", use_container_width=True):
                st.success(f"草稿 {item['id']} 已送出！")
                
            if b3.button("🗑️ 刪除", key=f"trv_del_{item['id']}", use_container_width=True):
                st.warning(f"草稿 {item['id']} 已刪除！")
                
        st.markdown("<hr style='margin: 0px; padding: 0px; border-top: 1px solid #f0f2f6;'>", unsafe_allow_html=True)


def render_submitted_list():
    st.subheader("📤 已送出表單列表")
    st.caption("查詢歷史紀錄。")

    mock_submitted = [
        {"id": "TR-20260228-01", "date": "2026-02-28", "traveler": "陳經理", "purpose": "高雄拜訪", "amount": 4500, "status": "待簽核"},
        {"id": "TR-20260225-05", "date": "2026-02-25", "traveler": "林專員", "purpose": "新竹受訓", "amount": 800, "status": "已結案"}
    ]

    if not mock_submitted:
        st.info("目前沒有已送出的表單。")
        return

    st.markdown("---")
    h_col1, h_col2, h_col3, h_col4, h_col5, h_col6, h_col7 = st.columns([1.5, 1.5, 1.5, 2.5, 1.5, 1.5, 2.5])
    h_col1.markdown("**單號**"); h_col2.markdown("**送出日期**"); h_col3.markdown("**出差人**")
    h_col4.markdown("**出差事由**"); h_col5.markdown("**總金額**"); h_col6.markdown("**狀態**"); h_col7.markdown("**操作**")
    st.markdown("---")

    for item in mock_submitted:
        col1, col2, col3, col4, col5, col6, col7 = st.columns([1.5, 1.5, 1.5, 2.5, 1.5, 1.5, 2.5])
        col1.write(item["id"])
        col2.write(item["date"])
        col3.write(item["traveler"])
        col4.write(item["purpose"])
        col5.write(f"${item['amount']:,}")
        col6.write(f"🟢 {item['status']}" if item['status'] == '已結案' else f"🟡 {item['status']}")
        
        with col7:
            b1, b2 = st.columns(2)
            if b1.button("👁️ 檢視", key=f"trv_view_{item['id']}", use_container_width=True):
                st.toast(f"檢視 {item['id']} 詳細內容")
                
            if b2.button("📥 下載 PDF", key=f"trv_pdf_{item['id']}", type="primary", use_container_width=True):
                st.info(f"功能開發中，目前請至新增表單頁面產生 PDF。")
                
        st.markdown("<hr style='margin: 0px; padding: 0px; border-top: 1px solid #f0f2f6;'>", unsafe_allow_html=True)


def run_app(view_mode='new_form'):
    if view_mode == 'new_form':
        render_new_form()
    elif view_mode == 'draft_list':
        render_draft_list()
    elif view_mode == 'submitted_list':
        render_submitted_list()

if __name__ == "__main__":
    run_app()