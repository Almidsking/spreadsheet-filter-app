import streamlit as st
import pandas as pd

st.set_page_config(page_title="Universal Spreadsheet Filter", layout="wide")

st.title("📊 Universal Spreadsheet Filter")

uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx", "xls"])

if uploaded_file:
    header_row = st.number_input(
        "Which row contains column headers? (First row = 1)",
        min_value=1,
        step=1
    )

    try:
        df = pd.read_excel(uploaded_file, header=header_row - 1)
        st.success("File loaded successfully")

        col1, col2, col3 = st.columns(3)

        with col1:
            column = st.selectbox("Column", df.columns)

        with col2:
            operator = st.selectbox(
                "Operator",
                ['==', '!=', '>', '<', '>=', '<=', 'contains']
            )

        with col3:
            series = df[column]

            if pd.api.types.is_bool_dtype(series):
                value = st.selectbox("Value", sorted(series.dropna().unique()))
            else:
                value = st.text_input("Value")

        if st.button("Apply Filter"):
            if value == "":
                st.error("Please enter a value")
            else:
                try:
                    if pd.api.types.is_numeric_dtype(series):
                        value = float(value)

                    if operator == '==':
                        filtered = df[series == value]
                    elif operator == '!=':
                        filtered = df[series != value]
                    elif operator == '>':
                        filtered = df[series > value]
                    elif operator == '<':
                        filtered = df[series < value]
                    elif operator == '>=':
                        filtered = df[series >= value]
                    elif operator == '<=':
                        filtered = df[series <= value]
                    elif operator == 'contains':
                        filtered = df[series.astype(str).str.contains(str(value), case=False)]
                    else:
                        st.error("Invalid operator")

                    # Strip time from datetime columns
                    for dt_col in filtered.select_dtypes(include=['datetime64[ns]']).columns:
                        filtered[dt_col] = filtered[dt_col].dt.date

                    st.success(f"Rows matched: {len(filtered)}")
                    st.dataframe(filtered)

                    # Download button
                    output = filtered.to_excel(index=False)
                    st.download_button(
                        "⬇️ Download Filtered Excel",
                        data=output,
                        file_name="filtered_output.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

                except Exception as e:
                    st.error(f"Filtering failed: {e}")

    except Exception as e:
        st.error(f"Failed to load file: {e}")
