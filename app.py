import streamlit as st
import pandas as pd
import numpy as np
import os

st.set_page_config(
    page_title="Insurance Premium Calculator",
    page_icon="💰",
    layout="wide"
)

st.title("💰 Aviva GCL Insurance Premium Calculator")
st.markdown("Select plan details below")

FILE_MAP = {
    ("Single Life", "Home Loan"): "Aviva Single HomeLoan.xlsx",
    ("Single Life", "LAP"):       "Aviva Single Lap.xlsx",
    ("Joint Life",  "Home Loan"): "Aviva Joint Homeloan.xlsx",
    ("Joint Life",  "LAP"):       "Aviva Joint Lap.xlsx",
}

def load_rate_table(life_type, loan_type):
    fname = FILE_MAP[(life_type, loan_type)]
    if not os.path.exists(fname):
        raise FileNotFoundError(
            f"File not found: '{fname}' — Please make sure this file is in the GitHub repo."
        )

    raw = pd.read_excel(fname, sheet_name="Sheet1", header=None)

    header_row = None
    for i, row in raw.iterrows():
        for val in row.values:
            if isinstance(val, str) and "AGE" in val.upper():
                header_row = i
                break
        if header_row is not None:
            break

    if header_row is None:
        raise ValueError("Could not find AGE/TERM header row in the file.")

    df = pd.read_excel(fname, sheet_name="Sheet1", header=header_row)
    df.columns = [str(c).strip() for c in df.columns]

    age_col = df.columns[0]
    df = df.dropna(subset=[age_col])
    df[age_col] = pd.to_numeric(df[age_col], errors='coerce')
    df = df.dropna(subset=[age_col])
    df[age_col] = df[age_col].astype(int)
    df = df.set_index(age_col)

    tenure_map = {}
    for col in df.columns:
        try:
            tenure_map[int(float(col))] = col
        except Exception:
            pass

    return df, tenure_map


def get_rate(df, tenure_map, age, tenure):
    if age not in df.index:
        raise ValueError(f"Age {age} not found in rate table.")
    if tenure not in tenure_map:
        raise ValueError(f"Tenure {tenure} yrs not found in rate table.")
    return float(df.loc[age, tenure_map[tenure]])


# ============================================
# DROPDOWNS
# ============================================

col1, col2 = st.columns(2)
with col1:
    life_type = st.selectbox("Select Life Type", ["Single Life", "Joint Life"])
with col2:
    loan_type = st.selectbox("Select Loan Type", ["Home Loan", "LAP"])

st.divider()

# ============================================
# MANUAL SECTION
# ============================================

st.subheader("🔢 Manual Rate Lookup")

col3, col4 = st.columns(2)
with col3:
    age = st.number_input("Enter Age", min_value=18, max_value=65, value=30, step=1)

# tenure limits based on loan type
max_tenure = 25 if loan_type == "Home Loan" else 10

with col4:
    tenure = st.number_input("Enter Tenure", min_value=2, max_value=max_tenure, value=5, step=1)
    st.caption("📅 Tenure is in Years")

if st.button("Get Rate", type="primary"):
    try:
        df_rates, tenure_map = load_rate_table(life_type, loan_type)
        rate = get_rate(df_rates, tenure_map, age, tenure)
        st.success(f"✅ {life_type} | {loan_type} | Age {age} | Tenure {tenure} yrs")
        st.metric("Rate", f"₹ {rate:,.2f}")
    except Exception as e:
        st.error(f"Error: {e}")

st.divider()

# ============================================
# EXCEL UPLOAD SECTION
# ============================================

st.subheader("📂 Upload Member Data for Bulk Rate Lookup")
st.markdown("Your Excel must have at least: **Name**, **Age**, **Tenure** (in years)")
st.warning("⚠️ Please make sure you have selected **Life Type** and **Loan Type** above before uploading your Excel file.")

uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx"])

if uploaded_file is not None:
    try:
        df = pd.read_excel(uploaded_file)
        df.columns = df.columns.str.strip()

        st.subheader("Uploaded Data Preview")
        st.dataframe(df.head())

        # Mandatory columns check
        name_col = next((c for c in df.columns if "name" in c.lower()), None)
        age_col = next((c for c in df.columns if "age" in c.lower()), None)
        tenure_col = next((c for c in df.columns if "tenure" in c.lower()), None)

        if not name_col or not age_col or not tenure_col:
            raise ValueError("Excel must contain mandatory columns: Name, Age, Tenure")

        # Convert Age and Tenure to numeric
        df[age_col] = pd.to_numeric(df[age_col], errors='coerce')
        df[tenure_col] = pd.to_numeric(df[tenure_col], errors='coerce')

        # Auto-detect months vs years
        if df[tenure_col].dropna().median() > 30:
            st.info("ℹ️ Tenure values look like months — auto-converting to years.")
            df[tenure_col] = (df[tenure_col] / 12).round(0).astype('Int64')
        else:
            df[tenure_col] = df[tenure_col].round(0).astype('Int64')

        df[age_col] = df[age_col].round(0).astype('Int64')

        # Apply tenure limits based on loan type
        max_tenure = 25 if loan_type == "Home Loan" else 10
        df[tenure_col] = df[tenure_col].clip(lower=2, upper=max_tenure)

        df_rates, tenure_map = load_rate_table(life_type, loan_type)

        rates = []
        remarks = []

        for _, row in df.iterrows():
            try:
                r_age = int(row[age_col])
                r_tenure = int(row[tenure_col])
                r = get_rate(df_rates, tenure_map, r_age, r_tenure)
                rates.append(round(r, 2))
                remarks.append("✅")
            except Exception as e:
                rates.append(None)
                remarks.append(f"❌ {e}")

        df['Rate'] = rates
        df['Status'] = remarks

        st.subheader("Rate Lookup Output")
        st.dataframe(df, use_container_width=True)

        output_file = "Rate_Output.xlsx"
        df.to_excel(output_file, index=False)

        with open(output_file, "rb") as file:
            st.download_button(
                label="⬇ Download Output Excel",
                data=file,
                file_name=output_file,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"Error: {e}")
