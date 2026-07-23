import streamlit as st

def show_logo():
    col1, col2 = st.columns([1, 9])
    with col1:
        st.image("assets/unloop-logo.png", width=180)
