import streamlit as st

# 入口頁：不顯示左側欄（包含 Streamlit 頁面導覽）
st.markdown(
    """
    <style>
    [data-testid="stSidebar"] {display:none !important;}
    [data-testid="collapsedControl"] {display:none !important;}
    </style>
    """,
    unsafe_allow_html=True,
)

st.title("💼 報帳管理入口網")
st.caption("請選擇要操作的系統。")

st.write("")
c1, c2 = st.columns(2, gap="large")

with c1:
    st.subheader("💰 支出報帳系統")
    st.write("一般性採購、廠商付款、零用金報銷等。")
    if st.button("進入 支出報帳", use_container_width=True, type="primary"):
        st.switch_page("expense.py")

with c2:
    st.subheader("🚆 出差報帳系統")
    st.write("國內出差申請、交通費與膳雜費明細報銷等。")
    if st.button("進入 出差報帳", use_container_width=True, type="primary"):
        st.switch_page("apps/travel_old.py")

st.divider()
st.info("若雲端寫入失敗，請確認系統目錄下 data/config.json（支出）與 data/travel_config.json（出差）中的 Sheet ID / Apps Script URL 是否正確，或請聯絡管理者。")
