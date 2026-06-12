import streamlit as st
import pandas as pd
import numpy as np

st.set_page_config(
    page_title="Insurance Premium Calculator",
    page_icon="💰",
    layout="wide"
)

st.title("💰 Group Term Life Insurance Premium Calculator")
st.markdown("Select plan details and upload your Excel file for premium calculation")

GST_RATE = 0.18

# ============================================
# FILE MAPPING
# ============================================

FILE_MAP = {
    "Single Life": {
        "file": "Aviva Single life.xlsx",
        "Home Loan": "Home Loan",
        "LAP": "Lap"
    },
    "Joint Life": {
        "file": "Aviva Joint life.xlsx",
        "Home Loan": "Homeloan",
        "LAP": "Lap"
    }
}

# ============================================
# RATE LOOKUP FUNCTION
# ============================================

def load_rate_table(life_type, loan_type):
    file_info = FILE_MAP[life_type]
    file_name = file_info["file"]
    sheet_name = file_info[loan_type]

    # Read all rows to find the AGE/TENURE header row
    raw = pd.read_excel(file_name, sheet_name=sheet_name, header=None)

    header_row = None
    for i, row in raw.iterrows():
        for val in row.values:
            if isinstance(val, str) and "AGE" in val.upper():
                header_row = i
                break
        if header_row is not None:
            break

    if header_row is None:
        raise ValueError(f"Could not find AGE/TENURE header in sheet '{sheet_name}'")

    df = pd.read_excel(file_name, sheet_name=sheet_name, header=header_row)
    df.columns = [str(c).strip() for c in df.columns]

    # First column is Age
    age_col = df.columns[0]
    df = df.dropna(subset=[age_col])
    df[age_col] = pd.to_numeric(df[age_col], errors='coerce')
    df = df.dropna(subset=[age_col])
    df[age_col] = df[age_col].astype(int)
    df = df.set_index(age_col)

    # Tenure columns - keep only numeric ones
    numeric_cols = {}
    for col in df.columns:
        try:
            numeric_cols[int(float(col))] = col
        except:
            pass

    return df, numeric_cols


def get_rate(df, numeric_cols, age, tenure):
    if age not in df.index:
        raise ValueError(f"Age {age} not found in rate table")
    if tenure not in numeric_cols:
        raise ValueError(f"Tenure {tenure} not found in rate table")
    col_name = numeric_cols[tenure]
    return float(df.loc[age, col_name])


# ============================================
# UI - DROPDOWNS AND INPUTS
# ============================================

col1, col2 = st.columns(2)

with col1:
    life_type = st.selectbox(
        "Select Life Type",
        ["Single Life", "Joint Life"]
    )

with col2:
    loan_type = st.selectbox(
        "Select Loan Type",
        ["Home Loan", "LAP"]
    )

col3, col4 = st.columns(2)

with col3:
    age = st.number_input(
        "Enter Age",
        min_value=18,
        max_value=70,
        value=30,
        step=1
    )

with col4:
    tenure = st.number_input(
        "Enter Tenure",
        min_value=1,
        max_value=30,
        value=5,
        step=1
    )
    st.caption("📅 Tenure is in **Years**")

# ============================================
# GET RATE BUTTON
# ============================================

if st.button("Get Rate"):
    try:
        df_rates, numeric_cols = load_rate_table(life_type, loan_type)
        rate = get_rate(df_rates, numeric_cols, age, tenure)
        st.success(f"✅ Rate for Age **{age}**, Tenure **{tenure} years** ({life_type} | {loan_type}): **₹ {rate:,.2f}** per lakh")
    except Exception as e:
        st.error(f"Error: {e}")

st.divider()

# ============================================
# FILE UPLOAD
# ============================================

st.subheader("Upload Member Data")
uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx"])

if uploaded_file is not None:
    try:
        df = pd.read_excel(uploaded_file)
        df.columns = df.columns.str.strip()

        st.subheader("Uploaded Data Preview")
        st.dataframe(df.head())

        required_columns = [
            'Loan Account No.',
            'Name of Primary Loan borrower',
            'Mobile No',
            'Sum Assured'
        ]

        for col in required_columns:
            if col not in df.columns:
                df[col] = ""

        df['Sum Assured'] = pd.to_numeric(df['Sum Assured'], errors='coerce').fillna(0)

        if 'MAIN MEMBER AGE' not in df.columns:
            df['MAIN MEMBER AGE'] = age

        if 'Loan Outstanding Amount' not in df.columns:
            df['Loan Outstanding Amount'] = 0

        # Load rate table once
        df_rates, numeric_cols = load_rate_table(life_type, loan_type)

        def calc_premium(row):
            try:
                member_age = int(row['MAIN MEMBER AGE']) if row['MAIN MEMBER AGE'] != 0 else age
                r = get_rate(df_rates, numeric_cols, member_age, tenure)
                return (row['Sum Assured'] / 100000) * r
            except:
                # Fallback to selected age if member age lookup fails
                try:
                    r = get_rate(df_rates, numeric_cols, age, tenure)
                    return (row['Sum Assured'] / 100000) * r
                except:
                    return 0

        df['Premium Excl GST'] = df.apply(calc_premium, axis=1)
        df['Premium + GST'] = df['Premium Excl GST'] * (1 + GST_RATE)

        output_columns = [
            'Loan Account No.',
            'Name of Primary Loan borrower',
            'Mobile No',
            'MAIN MEMBER AGE',
            'Sum Assured',
            'Premium Excl GST',
            'Premium + GST'
        ]

        final_df = df[output_columns]

        st.subheader("Portfolio Summary")
        c1, c2, c3 = st.columns(3)

        with c1:
            st.metric("Total Members", len(final_df))
        with c2:
            st.metric("Total Sum Assured", f"₹ {final_df['Sum Assured'].sum():,.0f}")
        with c3:
            st.metric("Total Premium (incl. GST)", f"₹ {final_df['Premium + GST'].sum():,.2f}")

        st.subheader("Premium Calculation Output")
        st.dataframe(final_df, use_container_width=True)

        output_file = "Premium_Output.xlsx"
        final_df.to_excel(output_file, index=False)

        with open(output_file, "rb") as file:
            st.download_button(
                label="⬇ Download Output Excel",
                data=file,
                file_name=output_file,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"Error: {e}")
