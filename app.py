import streamlit as st
import pandas as pd
import io

st.title("Excel File Combiner")

uploaded_files = st.file_uploader("Upload multiple Excel files (.xlsx)", type="xlsx", accept_multiple_files=True)

if uploaded_files:
    all_data = pd.DataFrame()

    for file in uploaded_files:
        df = pd.read_excel(file)
        all_data = pd.concat([all_data, df], ignore_index=True)

    st.success(f"? {len(uploaded_files)} files combined successfully!")
    st.dataframe(all_data)

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        all_data.to_excel(writer, index=False)
    st.download_button("?? Download Combined Excel", data=output.getvalue(), file_name="Merged.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
