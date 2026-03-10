import streamlit as st
import pandas as pd
from datetime import date

def render_new_form():
    st.subheader("📝 新增支出報帳單")
    st.caption("帶有 * 號的欄位為必填項目。")

    # --- 區塊 1：基本資料 ---
    st.markdown("##### 👤 報帳基本資訊")
    col1, col2, col3 = st.columns(3)
    with col1:
        applicant = st.text_input("申請人 *", placeholder="請輸入姓名")
    with col2:
        department = st.text_input("申請部門")
    with col3:
        apply_date = st.date_input("申請日期", value=date.today())

    # --- 區塊 2：付款方式與對象 (互斥選擇) ---
    st.markdown("##### 💳 付款資訊")
    payment_type = st.radio(
        "付款對象 (單選) *", 
        ["員工代墊款 (匯入員工帳戶)", "借支充抵 (沖銷原借支)", "廠商付款 (匯入廠商帳戶)"],
        horizontal=True
    )
    
    col4, col5 = st.columns(2)
    with col4:
        if "廠商付款" in payment_type:
            vendor_name = st.text_input("廠商名稱 *")
        else:
            payee_name = st.text_input("受款人姓名 *", value=applicant) # 預設帶入申請人
    with col5:
        project_id = st.text_input("歸屬計畫編號 / 專案名稱")

    st.divider()

    # --- 區塊 3：支出明細 (可編輯表格) ---
    st.markdown("##### 🛒 支出項目明細")
    if 'expense_items' not in st.session_state:
        st.session_state.expense_items = pd.DataFrame([
            {"憑證日期": str(date.today()), "發票/收據號碼": "", "品名/用途說明": "", "數量": 1, "單價": 0, "小計": 0}
        ])

    edited_exp_df = st.data_editor(
        st.session_state.expense_items,
        num_rows="dynamic",
        use_container_width=True,
        hide_index=True,
        column_config={
            "數量": st.column_config.NumberColumn("數量", min_value=1, step=1),
            "單價": st.column_config.NumberColumn("單價", min_value=0, step=1),
            "小計": st.column_config.NumberColumn("小計", min_value=0, step=1, help="請自行填寫或確認金額"),
        }
    )

    total_amount = edited_exp_df["小計"].sum()
    st.markdown(f"<h4 style='text-align: right; color: #1565C0;'>總計新台幣： {total_amount:,} 元</h4>", unsafe_allow_html=True)

    st.divider()

    # --- 區塊 4：操作按鈕 ---
    b1, b2, b3, b4 = st.columns(4)
    if b1.button("💾 儲存草稿", key="exp_save_draft", use_container_width=True):
        st.success("草稿已儲存！")
        
    if b2.button("🚀 送出", key="exp_submit", type="primary", use_container_width=True):
        if not applicant:
            st.error("⚠️ 請填寫申請人！")
        else:
            st.success("✅ 支出報帳單已成功送出！")
            st.session_state.current_view = 'submitted_list'
            st.rerun()

    if b3.button("📥 下載PDF", key="exp_dl_pdf", use_container_width=True):
        st.info("正在產生 PDF...")

    if b4.button("📋 查看送出列表", key="exp_view_list", use_container_width=True):
        st.session_state.current_view = 'submitted_list'
        st.rerun()

def render_draft_list():
    st.subheader("📄 支出草稿列表")
    mock_drafts = [
        {"id": "DR-EXP-001", "date": "2026-03-01", "applicant": "王小明", "type": "員工代墊款", "amount": 1500},
        {"id": "DR-EXP-002", "date": "2026-03-02", "applicant": "陳採購", "type": "廠商付款", "amount": 25000}
    ]

    if not mock_drafts:
        st.info("目前沒有任何草稿。")
        return

    st.markdown("---")
    h_col1, h_col2, h_col3, h_col4, h_col5, h_col6 = st.columns([1.5, 1.5, 1.5, 3, 1.5, 3])
    h_col1.markdown("**單號**"); h_col2.markdown("**建立日期**"); h_col3.markdown("**申請人**")
    h_col4.markdown("**付款類型**"); h_col5.markdown("**預估金額**"); h_col6.markdown("**操作**")
    st.markdown("---")

    for item in mock_drafts:
        col1, col2, col3, col4, col5, col6 = st.columns([1.5, 1.5, 1.5, 3, 1.5, 3])
        col1.write(item["id"]); col2.write(item["date"]); col3.write(item["applicant"])
        col4.write(item["type"]); col5.write(f"${item['amount']:,}")
        
        with col6:
            b1, b2, b3 = st.columns(3)
            if b1.button("✏️ 編輯", key=f"exp_edit_{item['id']}", use_container_width=True):
                st.session_state.edit_target_id = item["id"]
                st.session_state.current_view = 'new_form'
                st.rerun()
            if b2.button("🚀 送出", key=f"exp_submit_{item['id']}", type="primary", use_container_width=True):
                st.success(f"草稿 {item['id']} 已送出！")
            if b3.button("🗑️ 刪除", key=f"exp_del_{item['id']}", use_container_width=True):
                st.warning(f"草稿 {item['id']} 已刪除！")
        st.markdown("<hr style='margin: 0px; padding: 0px; border-top: 1px solid #f0f2f6;'>", unsafe_allow_html=True)

def render_submitted_list():
    st.subheader("📤 已送出支出列表")
    mock_submitted = [
        {"id": "EXP-20260228-01", "date": "2026-02-28", "applicant": "林專員", "type": "廠商付款", "amount": 12000, "status": "待會計審核"},
    ]

    if not mock_submitted:
        st.info("目前沒有已送出的表單。")
        return

    st.markdown("---")
    h_col1, h_col2, h_col3, h_col4, h_col5, h_col6, h_col7 = st.columns([1.5, 1.5, 1.5, 2.5, 1.5, 1.5, 2.5])
    h_col1.markdown("**單號**"); h_col2.markdown("**送出日期**"); h_col3.markdown("**申請人**")
    h_col4.markdown("**付款類型**"); h_col5.markdown("**總金額**"); h_col6.markdown("**狀態**"); h_col7.markdown("**操作**")
    st.markdown("---")

    for item in mock_submitted:
        col1, col2, col3, col4, col5, col6, col7 = st.columns([1.5, 1.5, 1.5, 2.5, 1.5, 1.5, 2.5])
        col1.write(item["id"]); col2.write(item["date"]); col3.write(item["applicant"])
        col4.write(item["type"]); col5.write(f"${item['amount']:,}")
        col6.write(f"🟡 {item['status']}")
        
        with col7:
            b1, b2 = st.columns(2)
            if b1.button("👁️ 檢視", key=f"exp_view_{item['id']}", use_container_width=True):
                st.toast(f"檢視 {item['id']} 詳細內容")
            if b2.button("📥 下載 PDF", key=f"exp_pdf_{item['id']}", type="primary", use_container_width=True):
                st.info(f"產生 {item['id']} PDF...")
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