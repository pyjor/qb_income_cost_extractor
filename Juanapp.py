import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Project Summary Extractor", layout="wide")
st.title("📊 Extract Contract Sales & COGS from Excel Files")

# Function to extract data from each Excel file
def extract_with_month_from_b6(file_obj, file_name):
    df = pd.read_excel(file_obj, sheet_name=0, header=None)

    # Project names from row 5 (index 4), columns B onward
    project_names = df.iloc[4, 1:].fillna("").astype(str).str.strip()

    # Only use columns with valid project names that include ()
    valid_cols = [
        i for i, name in enumerate(project_names, start=1)
        if name and "total" not in name.lower() and "(" in name and ")" in name
    ]

    # Extract month from B6 (index 5, col 1)
    month_cell = df.iloc[5, 1]
    month = str(month_cell).strip() if pd.notna(month_cell) else file_name

    # Find rows with specific labels
    sales_row_index = df[df.iloc[:, 0].astype(str).str.strip() == "61100 Contract Sales"].index
    cogs_row_index = df[df.iloc[:, 0].astype(str).str.strip() == "Total Cost of Goods Sold"].index

    if not sales_row_index.empty and not cogs_row_index.empty:
        sales_row = df.iloc[sales_row_index[0], valid_cols].fillna(0).astype(float)
        cogs_row = df.iloc[cogs_row_index[0], valid_cols].fillna(0).astype(float)

        # First income, then cost
        sales_df = pd.DataFrame({
            'Project': [project_names[i] + " - Income" for i in valid_cols],
            month: sales_row.values
        })

        cogs_df = pd.DataFrame({
            'Project': [project_names[i] + " - Cost" for i in valid_cols],
            month: cogs_row.values
        })

        return pd.concat([sales_df, cogs_df], axis=0).reset_index(drop=True)

    return pd.DataFrame()

# Upload files via Streamlit UI
uploaded_files = st.file_uploader("Upload one or more Excel files", type=["xlsx"], accept_multiple_files=True)

if uploaded_files:
    all_data = []
    for file in uploaded_files:
        result = extract_with_month_from_b6(file, file.name)
        all_data.append(result)

    combined_df = pd.concat(all_data, axis=0)
    pivot_df = combined_df.pivot_table(index="Project", aggfunc='first')
    pivot_df = pivot_df.fillna(0)

    st.success("✅ Data processed successfully!")
    st.dataframe(pivot_df)

    # Export to Excel
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        pivot_df.to_excel(writer, index=True, sheet_name='Project Summary')
    st.download_button("📥 Download Excel File", data=buffer.getvalue(), file_name="project_summary.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
else:
    st.info("Please upload at least one Excel file to begin.")
