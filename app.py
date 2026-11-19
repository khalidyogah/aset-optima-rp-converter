import io
import re
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Rp to Number Converter", layout="centered")
st.title("Excel Rupiah (Rp) to Full Number Converter")


def clean_rp(val):
    if pd.isna(val):
        return val
    if isinstance(val, (int, float)):
        return int(val) if float(val).is_integer() else val

    text = str(val).strip()
    if not re.search(r"rp", text, re.IGNORECASE):
        return val

    # Remove 'Rp', spaces, and thousand-separator dots; convert decimal comma to dot
    cleaned = re.sub(r"rp|\s|\.", "", text, flags=re.IGNORECASE).replace(",", ".")
    try:
        num = float(cleaned)
        return int(num) if num.is_integer() else num
    except ValueError:
        return val


uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)

    # Auto-detect columns containing 'Rp'
    detected_cols = [
        col
        for col in df.columns
        if df[col].astype(str).str.contains(r"\bRp\b", case=False, regex=True).any()
    ]

    selected_cols = st.multiselect(
        "Select columns to convert",
        options=df.columns.tolist(),
        default=detected_cols,
    )

    result_df = df.copy()
    for col in selected_cols:
        result_df[col] = result_df[col].apply(clean_rp)

    st.dataframe(result_df, use_container_width=True)

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        result_df.to_excel(writer, index=False)

    st.download_button(
        label="Download Converted Excel",
        data=output.getvalue(),
        file_name=f"converted_{uploaded_file.name}",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )