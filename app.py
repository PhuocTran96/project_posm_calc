import streamlit as st
import pandas as pd
import math
import logging
import base64
from io import BytesIO
from datetime import datetime

try:
    from cost_calc_improved import (
        calculate_posm_report,
    )
except ImportError:
    st.error("Error: Could not import calculation logic from cost_calc_improved.py. Make sure the file exists in the same directory.")
    st.stop()

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

st.set_page_config(
    page_title="POSM Cost Calculator",
    page_icon="https://res.cloudinary.com/dwnnet2w1/image/upload/v1744772802/logo-blue-2022_cd9wmm.png",
    layout="wide"
)

def add_bg_from_local():
    with open("bg.gif", "rb") as f:
        data = base64.b64encode(f.read()).decode("utf-8")
    
    st.markdown(
        f"""
        <style>
        .stApp {{
            background-image: linear-gradient(rgba(0, 0, 0, 0.7), rgba(0, 0, 0, 0.7)), url("data:image/gif;base64,{data}");
            background-size: cover;
            background-position: center;
            background-repeat: repeat;
        }}
        .stMarkdown, .stHeader, h1, h2, h3 {{
            color: white;
            text-shadow: 0px 0px 3px rgba(0,0,0,0.9);
        }}
        .stTextInput, .stFileUploader, .stDataFrame, .stAlert {{
            background-color: rgba(255, 255, 255, 0.85) !important;
            border-radius: 10px !important;
            padding: 10px !important;
        }}
        .stButton > button {{
            border-radius: 10px !important;
            box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
        }}
        .stButton {{
            display: inline-block !important;
            background-color: transparent !important;
        }}
        div[data-testid="stHorizontalBlock"] {{
            background-color: transparent !important;
        }}
        div.block-container {{
            background-color: transparent !important;
        }}
        </style>
        """,
        unsafe_allow_html=True
    )

# Gọi hàm để thêm ảnh nền
add_bg_from_local()

# Add a logo next to the title
col1, col2 = st.columns([1, 10])
with col1:
    with open("linkedin.gif", "rb") as f:
        data = base64.b64encode(f.read()).decode("utf-8")
    
    st.markdown(
        f"""
        <a href="https://www.linkedin.com/in/phuoctran1996" target="_blank">
            <img src="data:image/gif;base64,{data}" width="80">
        </a>
        """,
        unsafe_allow_html=True,
    )
with col2:
    st.markdown(
        """
        <h1 style="color: #FFD700; text-shadow: 0px 0px 3px rgba(0,0,0,0.9); font-weight: bold;">
        POSM Cost Calculator App
        </h1>
        """,
        unsafe_allow_html=True
    )

st.info("Upload the required Excel files below to calculate POSM costs and allocations.")

# --- File Uploaders ---
col1, col2, col3, col4 = st.columns(4)
with col1:
    uploaded_fact_display = st.file_uploader("1. Upload fact_display.xlsx", type="xlsx", key="fact_display")

with col2:
    uploaded_dim_storelist = st.file_uploader("2. Upload dim_storelist.xlsx", type="xlsx", key="dim_storelist")

with col3:
    uploaded_dim_model = st.file_uploader("3. Upload dim_model.xlsx", type="xlsx", key="dim_model")

with col4:
    uploaded_dim_posm = st.file_uploader("4. Upload dim_posm.xlsx (must contain 'posm' and 'price' sheets)", type="xlsx", key="dim_posm")

# --- Calculation Trigger ---
if st.button("🚀 Calculate Report", type="primary"):
    if uploaded_fact_display and uploaded_dim_storelist and uploaded_dim_model and uploaded_dim_posm:
        try:
            # Load data from uploaded files into DataFrames
            with st.spinner('Loading data...'):
                fact_display_df = pd.read_excel(uploaded_fact_display)
                dim_storelist_df = pd.read_excel(uploaded_dim_storelist)
                if "Store name" in dim_storelist_df.columns:
                     dim_storelist_df = dim_storelist_df.rename(columns={"Store name": "shop"})
                elif "shop" not in dim_storelist_df.columns:
                     st.error("Error: 'dim_storelist.xlsx' must contain either 'Store name' or 'shop' column.")
                     st.stop()

                dim_model_df = pd.read_excel(uploaded_dim_model)

                # Read both sheets from dim_posm.xlsx
                dim_posm_sheets = pd.read_excel(uploaded_dim_posm, sheet_name=None)
                dim_posm_df = dim_posm_sheets.get('posm')
                price_posm_df = dim_posm_sheets.get('price')

                if dim_posm_df is None or price_posm_df is None:
                    st.error("Error: 'dim_posm.xlsx' must contain both 'posm' and 'price' sheets.")
                    st.stop()

            # Perform calculation using the imported function
            with st.spinner('Calculating report... Please wait.'):
                posm_summary_df, address_summary_df = calculate_posm_report(
                    fact_display_df, dim_storelist_df, dim_model_df, dim_posm_df, price_posm_df
                )

            if posm_summary_df is not None and address_summary_df is not None:
                st.success("✅ Calculation Complete!")

                # --- Display Results ---
                st.subheader("📊 POSM Summary")
                st.dataframe(posm_summary_df, use_container_width=True)

                st.subheader("📍 Address Summary by POSM")
                st.dataframe(address_summary_df, use_container_width=True)

                # --- Download Button ---
                output = BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    posm_summary_df.to_excel(writer, sheet_name='POSM Summary', index=False)
                    address_summary_df.to_excel(writer, sheet_name='Address Summary by POSM', index=False)
                
                output.seek(0) 

                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                download_filename = f'posm_cost_report_{timestamp}.xlsx'

                st.download_button(
                    label="📥 Download Report as Excel",
                    data=output,
                    file_name=download_filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.error("❌ Calculation failed. Check the logs or input files for errors.")

        except Exception as e:
            st.error(f"An error occurred: {e}")
            logging.error(f"Streamlit App Error: {e}", exc_info=True)

    else:
        st.warning("⚠️ Please upload all required Excel files.")

st.markdown("---")
st.caption("Developed by Phước ADMIN | [LinkedIn](https://www.linkedin.com/in/phuoctran1996) | 2025")
st.caption("This app is designed to help you calculate POSM costs and allocations efficiently. If you have any questions or feedback, please reach out!")
