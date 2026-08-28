import streamlit as st
import pandas as pd
from io import BytesIO

# Page configuration
st.set_page_config(page_title="Household Excel Manager", page_icon="📊", layout="wide")

st.title("📊 Household Staff Excel Manager")
st.markdown("Upload your Excel file, add custom column filters, select specific columns for your output, sort, edit, and download.")

# --- 1. FILE UPLOADER ---
uploaded_file = st.file_uploader("Upload your Excel spreadsheet", type=["xlsx", "xls"])

if uploaded_file is not None:
    try:
        # Read the Excel file into a Pandas DataFrame
        df = pd.read_excel(uploaded_file)
        st.success("File uploaded successfully!")
        
        # --- 2. DYNAMIC MULTI-COLUMN FILTER BUILDER ---
        st.subheader("1. Dynamic Column Filters (Use as many as you want)")
        st.markdown("Add as many column filters as you need. Each filter lets you select a column and pick its specific values.")

        # Initialize session state to keep track of added filter rows
        if "num_filters" not in st.session_state:
            st.session_state.num_filters = 1

        # Buttons to add or remove filter rows dynamically
        col_btn1, col_btn2 = st.columns([1, 5])
        with col_btn1:
            if st.button("➕ Add Filter"):
                st.session_state.num_filters += 1
        with col_btn2:
            if st.button("🔄 Reset Filters") and st.session_state.num_filters > 1:
                st.session_state.num_filters = 1

        filtered_df = df.copy()
        columns_list = df.columns.tolist()

        # Loop through and render each active filter row
        for i in range(st.session_state.num_filters):
            st.markdown(f"**Filter #{i+1}**")
            f_col1, f_col2 = st.columns(2)
            
            with f_col1:
                selected_column = st.selectbox(f"Select Column {i+1}:", options=columns_list, key=f"filter_col_{i}")
            
            with f_col2:
                if selected_column:
                    unique_vals = df[selected_column].dropna().unique().tolist()
                    unique_vals_str = [str(v) for v in unique_vals]
                    
                    selected_vals = st.multiselect(
                        f"Choose values from **{selected_column}** (leave empty for none):", 
                        options=unique_vals_str, 
                        key=f"filter_vals_{i}"
                    )
                    
                    if selected_vals:
                        filtered_df = filtered_df[filtered_df[selected_column].astype(str).isin(selected_vals)]
            
            st.markdown("---")

        st.markdown(f"**Result:** Showing {len(filtered_df)} of {len(df)} rows after applying your filters.")

        # --- 3. SELECT SPECIFIC COLUMNS FOR OUTPUT ---
        st.subheader("2. Choose Output Columns")
        st.markdown("Select which columns you want to include in your final output file (uncheck any columns you want to drop).")
        
        selected_output_columns = st.multiselect(
            "Columns to keep:",
            options=filtered_df.columns.tolist(),
            default=filtered_df.columns.tolist(),
            key="output_columns_selector"
        )

        # Restrict dataframe to only selected columns if any are chosen
        if selected_output_columns:
            output_df = filtered_df[selected_output_columns]
        else:
            output_df = filtered_df  # Fallback if someone clears it completely

        # --- 4. OPTIONAL SORTING SECTION ---
        st.subheader("3. Sort Data (Optional)")
        sort_col1, sort_col2 = st.columns(2)
        
        sort_options = ["--- None (Keep Original Order) ---"] + output_df.columns.tolist()
        
        with sort_col1:
            sort_column = st.selectbox("Sort by column:", options=sort_options, key="sort_column_choice")
        with sort_col2:
            sort_order = st.radio("Order:", options=["Ascending (A-Z / Low-High)", "Descending (Z-A / High-Low)"], key="sort_order_choice")
            ascending_bool = True if "Ascending" in sort_order else False

        if sort_column != "--- None (Keep Original Order) ---":
            sorted_df = output_df.sort_values(by=sort_column, ascending=ascending_bool)
        else:
            sorted_df = output_df

        # --- 5. EDIT DATA & COLUMNS ---
        st.subheader("4. Edit Spreadsheet & Modify Data")
        st.info("💡 *Tip: You can edit cell values or add/remove rows directly in the table below.*")
        
        edited_df = st.data_editor(sorted_df, num_rows="dynamic", key="data_editor")

        # --- 6. DOWNLOAD OUTPUT FILE ---
        st.subheader("5. Download Result Sheet")
        
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            edited_df.to_excel(writer, index=False)
        processed_data = output.getvalue()

        st.download_button(
            label="📥 Download Custom Excel File",
            data=processed_data,
            file_name="custom_filtered_household_staff.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"An error occurred while processing the file: {e}")
else:
    st.warning("Please upload an Excel file to get started.")