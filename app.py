import streamlit as st
import pandas as pd
import io
from openpyxl.styles import numbers

st.set_page_config(page_title="Excel Combiner Tool", layout="centered")

# Title & Instructions
st.title("📊 Excel File Combiner")
st.markdown("""
Easily merge multiple `.xlsx` files into one.
- The second column in each file will be formatted as **Text**.
- A preview will appear before download.
""")

# File uploader
uploaded_files = st.file_uploader(
    label="🔼 Upload Excel Files",
    type="xlsx",
    accept_multiple_files=True,
    help="You can upload more than one Excel file."
)

# Process files
if uploaded_files:
    st.info(f"Processing {len(uploaded_files)} file(s)...")
    combined_df = pd.DataFrame()

    for file in uploaded_files:
        df = pd.read_excel(file)
        combined_df = pd.concat([combined_df, df], ignore_index=True)

    # Format column B as text
    if combined_df.shape[1] >= 2:
        combined_df.iloc[:, 1] = combined_df.iloc[:, 1].astype(str)

    # Preview table
    st.success(f"✅ Combined {len(uploaded_files)} file(s) successfully!")
    st.markdown("### 🧾 Preview of Combined Data")
    st.dataframe(combined_df, use_container_width=True)

    # Convert to Excel with column B as text
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        combined_df.to_excel(writer, index=False, sheet_name="Combined")
        worksheet = writer.sheets["Combined"]
        for row in range(2, worksheet.max_row + 1):
            cell = worksheet.cell(row=row, column=2)
            cell.number_format = numbers.FORMAT_TEXT

    # Download button
    st.download_button(
        label="📥 Download Combined Excel",
        data=output.getvalue(),
        file_name="Combined_File.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

else:
    st.warning("📂 Please upload at least one Excel file to begin.")
