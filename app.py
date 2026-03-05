import streamlit as st

# 企業報帳管理系統（多頁）
st.set_page_config(
    page_title="企業報帳管理系統",
    page_icon="💼",
    layout="wide",
    initial_sidebar_state="expanded",
)

home = st.Page("pages/home.py", title="回到入口", icon="🏠", default=True)
expense = st.Page("expense.py", title="支出報帳", icon="💰")
travel = st.Page("apps/travel_old.py", title="出差報帳", icon="🚆")

pg = st.navigation([home, expense, travel])
pg.run()
