import streamlit as st

st.set_page_config(page_title="Main App", page_icon="📁")

st.title("Main Dashboard")
st.write("Choose an app to open:")

col1, col2 = st.columns(2)

with col1:
    st.page_link("pages/product_price_checker.py", label="Go to App 1", icon="🟩", use_container_width=True)

with col2:
    st.page_link("pages/insert_product_picture.py", label="Go to App 2", icon="🟦", use_container_width=True)