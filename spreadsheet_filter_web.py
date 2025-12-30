import streamlit as st
import pandas as pd

st.set_page_config(page_title="Universal Spreadsheet Filter", layout="wide")
st.title("📊 Universal Spreadsheet Filter")

# --- Session State ---
if "filters" not in st.session_state:
    st.session_state.filters = []

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

        st.subheader("Filter Conditions")

        # Add new filter
        if st.button("➕ Add Filter"):
            st.session_state.filters.append({
                "column": df.columns[0],
                "operator": "==",
                "value": ""
            })

        logic = st.radio("Combine filters using:", ["AND", "OR"])

        for i, f in enumerate(st.session_state.filters):
            c1, c2, c3, c4 = st.columns([3, 2, 3, 1])

            with c1:
                f["column"] = st.selectbox(
                    "Column",
                    df.columns,
                    key=f"col_{i}"
                )

            with c2:
                f["operator"] = st.selectbox(
                    "Operator",
                    ['==', '!=', '>', '<', '>=', '<=', 'contains'],
                    key=f"op_{i}"
                )

            with c3:
                series = df[f["column"]]
                if pd.api.types.is_bool_dtype(series):
                    f["value"] = st.selectbox(
                        "Value",
                        sorted(series.dropna().unique()),
                        key=f"val_{i}"
                    )
                else:
                    f["value"] = st.text_input(
                        "Value",
                        f["value"],
                        key=f"val_{i}"
                    )

            with c4:
                if st.button("❌", key=f"del_{i}"):
                    st.session_state.filters.pop(i)
                    st.experimental_rerun()

        if st.button("Apply Filters"):
            if not st.session_state.filters:
                st.error("Add at least one filter.")
            else:
                try:
                    masks = []

                    for f in st.session_state.filters:
                        col = f["column"]
                        op = f["operator"]
                        val = f["value"]
                        series = df[col]

                        if pd.api.types.is_numeric_dtype(series):
                            val = float(val)

                        if op == '==':
                            mask = series == val
                        elif op == '!=':
                            mask = series != val
                        elif op == '>':
                            mask = series > val
                        elif op == '<':
                            mask = series < val
                        elif op == '>=':
                            mask = series >= val
                        elif op == '<=':
                            mask = series <= val
                        elif op == 'contains':
                            mask = series.astype(str).str.contains(str(val), case=False)
                        else:
                            continue

                        masks.append(mask)

                    final_mask = masks[0]
                    for m in masks[1:]:
                        final_mask = (
                            final_mask & m if logic == "AND"
                            else final_mask | m
                        )

                    filtered = df[final_mask]

                    for dt_col in filtered.select_dtypes(include=['datetime64[ns]']).columns:
                        filtered[dt_col] = filtered[dt_col].dt.date

                    st.success(f"Rows matched: {len(filtered)}")
                    st.dataframe(filtered)

                    st.download_button(
                        "⬇️ Download Filtered Excel",
                        filtered.to_excel(index=False),
                        file_name="filtered_output.xlsx"
                    )

                except Exception as e:
                    st.error(f"Filtering failed: {e}")

    except Exception as e:
        st.error(f"Failed to load file: {e}")
