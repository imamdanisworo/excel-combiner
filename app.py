import streamlit as st
import pandas as pd
import io
from openpyxl import load_workbook
from openpyxl.styles import numbers

st.title("Excel File Combiner")

uploaded_files = st.file_uploader("Upload multiple Excel files (.xlsx)", type="xlsx", accept_multiple_files=True)

if uploaded_files:
    all_data = pd.DataFrame()

    for file in uploaded_files:
        df = pd.read_excel(file)
        all_data = pd.concat([all_data, df], ignore_index=True)

    # Convert Column B (second column) to string
    if all_data.shape[1] >= 2:
        all_data.iloc[:, 1] = all_data.iloc[:, 1].astype(str)

    st.success(f"✅ {len(uploaded_files)} files combined successfully!")
    st.dataframe(all_data)

    # Save to Excel
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        all_data.to_excel(writer, index=False, sheet_name="Combined")

        # Access the worksheet to format Column B as text
        workbook = writer.book
        worksheet = writer.sheets["Combined"]
        for row in range(2, worksheet.max_row + 1):  # Start at row 2 to skip header
            cell = worksheet.cell(row=row, column=2)  # Column B is column 2
            cell.number_format = numbers.FORMAT_TEXT

    st.download_button(
        "📥 Download Combined Excel",
        data=output.getvalue(),
        file_name="Merged.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
