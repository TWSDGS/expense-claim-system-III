import streamlit as st

st.set_page_config(page_title="企業報帳管理系統", page_icon="💼", layout="wide")

home_page = st.Page("pages/home.py", title="入口", icon="🏠", default=True)
expense_page = st.Page("expense.py", title="支出報帳", icon="💰")
travel_page = st.Page("apps/travel_old.py", title="出差報帳", icon="🚆")

pg = st.navigation([home_page, expense_page, travel_page])
pg.run()
