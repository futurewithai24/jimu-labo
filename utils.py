import streamlit as st
import base64
from pathlib import Path

_logo_b64 = base64.b64encode(Path("assets/unloop-logo.png").read_bytes()).decode()

def show_logo():
    st.markdown(
        f'<img src="data:image/png;base64,{_logo_b64}" width="140" style="display:block;margin-bottom:8px;border-radius:10px;">',
        unsafe_allow_html=True
    )
