import streamlit as st
import pandas as pd
import math
import logging
from io import BytesIO
from datetime import datetime


try:
    from cost_calc_improved import (
        calculate_posm_report,
    )
except ImportError:
    st.error("Error: Could not import calculation logic from cost_calc_improved.py. Make sure the file exists in the same directory.")
    st.stop() # Stop execution if import fails

# Configure logging (optional for Streamlit, but good practice)
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

st.set_page_config(
    page_title="POSM Cost Calculator",
    page_icon="https://res.cloudinary.com/dwnnet2w1/image/upload/v1744772802/logo-blue-2022_cd9wmm.png",
    layout="wide"
)

st.markdown(
    """
    <style>
    .stTitle {
        margin-top: -20px;  /* Adjust this value to reduce or increase the spacing */
    }
    </style>
    """,
    unsafe_allow_html=True
)

# Add a logo next to the title
col1, col2 = st.columns([1, 10])
with col1:
    # Use HTML to display the animated GIF
    import base64
    with open("linkedin.gif", "rb") as f:
        data = base64.b64encode(f.read()).decode("utf-8")
    
    st.markdown(
        f"""
        <img src="data:image/gif;base64,{data}" width="80">
        """,
        unsafe_allow_html=True,
    )
with col2:
    st.title("POSM Cost Calculator App")

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
    # Check if all files are uploaded
    if uploaded_fact_display and uploaded_dim_storelist and uploaded_dim_model and uploaded_dim_posm:
        try:
            # Load data from uploaded files into DataFrames
            with st.spinner('Loading data...'):
                fact_display_df = pd.read_excel(uploaded_fact_display)
                # Ensure 'shop' column exists after renaming
                dim_storelist_df = pd.read_excel(uploaded_dim_storelist)
                if "Store name" in dim_storelist_df.columns:
                     dim_storelist_df = dim_storelist_df.rename(columns={"Store name": "shop"})
                elif "shop" not in dim_storelist_df.columns:
                     st.error("Error: 'dim_storelist.xlsx' must contain either 'Store name' or 'shop' column.")
                     st.stop()

                dim_model_df = pd.read_excel(uploaded_dim_model)

                # Read both sheets from dim_posm.xlsx
                dim_posm_sheets = pd.read_excel(uploaded_dim_posm, sheet_name=None) # Read all sheets
                dim_posm_df = dim_posm_sheets.get('posm')
                price_posm_df = dim_posm_sheets.get('price')

                if dim_posm_df is None or price_posm_df is None:
                    st.error("Error: 'dim_posm.xlsx' must contain both 'posm' and 'price' sheets.")
                    st.stop() # Stop if sheets are missing

            # Perform calculation using the imported function
            with st.spinner('Calculating report... Please wait.'):
                final_df, province_summary_df = calculate_posm_report(
                    fact_display_df, dim_storelist_df, dim_model_df, dim_posm_df, price_posm_df
                )

            if final_df is not None and province_summary_df is not None:
                st.success("✅ Calculation Complete!")

                # --- Display Results ---
                st.subheader("POSM Summary")
                st.dataframe(final_df)

                st.subheader("Province Summary")
                st.dataframe(province_summary_df)

                # --- Download Button ---
                # Create an in-memory Excel file
                output = BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    final_df.to_excel(writer, sheet_name='POSM Summary', index=False)
                    province_summary_df.to_excel(writer, sheet_name='Province Summary', index=False)
                
                # Important: Seek back to the beginning of the stream
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
st.caption("Developed by Cline")
