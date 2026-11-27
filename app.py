import io
import re
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Excel Data Cleaner", layout="centered")
st.title("Excel Data Cleaner")


def clean_rp(val):
    if pd.isna(val):
        return val
    if isinstance(val, (int, float)):
        return int(val) if float(val).is_integer() else val

    text = str(val).strip()
    if not re.search(r"rp", text, re.IGNORECASE):
        return val

    cleaned = re.sub(r"rp|\s|\.", "", text, flags=re.IGNORECASE).replace(",", ".")
    try:
        num = float(cleaned)
        return int(num) if num.is_integer() else num
    except ValueError:
        return val


def clean_text(val):
    if pd.isna(val) or not isinstance(val, str):
        return val
    # 1. Delete all ' and "
    text = val.replace("'", "").replace('"', "")
    # 2. Replace 'ags' (case-insensitive) with 'Agu'
    text = re.sub(r"ags", "Agu", text, flags=re.IGNORECASE)
    return text


uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)

    detected_rp_cols = [
        col
        for col in df.columns
        if df[col].astype(str).str.contains(r"\bRp\b", case=False, regex=True).any()
    ]

    selected_rp_cols = st.multiselect(
        "Select columns to convert from Rp to number",
        options=df.columns.tolist(),
        default=detected_rp_cols,
    )

    apply_text_clean = st.checkbox(
        "Delete all ' and \" & replace 'ags' with 'Agu' across all columns",
        value=True,
    )

    result_df = df.copy()

    for col in selected_rp_cols:
        result_df[col] = result_df[col].apply(clean_rp)

    if apply_text_clean:
        for col in result_df.columns:
            result_df[col] = result_df[col].apply(clean_text)

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