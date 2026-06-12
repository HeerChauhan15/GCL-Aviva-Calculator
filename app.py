import streamlit as st
import pandas as pd
import numpy as np
import os

st.set_page_config(
    page_title="Insurance Premium Calculator",
    page_icon="💰",
    layout="wide"
)

st.title("💰 Group Term Life Insurance Premium Calculator")
st.markdown("Select plan details and upload your Excel file for premium calculation")

GST_RATE = 0.18

# Map selections to file names (files must be in same folder as app.py)
FILE_MAP = {
    ("Single Life", "Home Loan"): "Aviva_Single_HomeLoan.xlsx",
    ("Single Life", "LAP"):       "Aviva_Single_Lap.xlsx",
    ("Joint Life",  "Home Loan"): "Aviva_Joint_Homeloan.xlsx",
    ("Joint Life",  "LAP"):       "Aviva_Joint_Lap.xlsx",
}

def load_rate_table(life_type, loan_type):
    fname = FILE_MAP[(life_type, loan_type)]
    if not os.path.exists(fname):
        raise FileNotFoundError(
            f"File not found: '{fname}' — Please make sure this file is in the same folder as app.py"
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
        except:
            pass

    return df, tenure_map


def get_rate(df, tenure_map, age, tenure):
    if age not in df.index:
        raise ValueError(f"Age {age} not found in rate table. Available ages: {sorted(df.index.tolist())}")
    if tenure not in tenure_map:
        raise ValueError(f"Tenure {tenure} not found in rate table. Available tenures: {sorted(tenure_map.keys())}")
    return float(df.loc[age, tenure_map[tenure]])


col1, col2 = st.columns(2)
with col1:
    life_type = st.selectbox("Select Life Type", ["Single Life", "Joint Life"])
with col2:
    loan_type = st.selectbox("Select Loan Type", ["Home Loan", "LAP"])

col3, col4 = st.columns(2)
with col3:
    age = st.number_input("Enter Age", min_value=18, max_value=70, value=30, step=1)
with col4:
    tenure = st.number_input("Enter Tenure", min_value=2, max_value=30, value=5, step=1)
    st.caption("📅 Tenure is in Years")

if st.button("Get Rate", type="primary"):
    try:
        df_rates, tenure_map = load_rate_table(life_type, loan_type)
        rate = get_rate(df_rates, tenure_map, age, tenure)
        premium_excl_gst = rate
        premium_incl_gst = rate * (1 + GST_RATE)

        st.success(f"✅ Rate found for Age {age} | Tenure {tenure} years | {life_type} | {loan_type}")

        m1, m2, m3 = st.columns(3)
        with m1:
            st.metric("Rate per Lakh (Excl GST)", f"₹ {premium_excl_gst:,.2f}")
        with m2:
            st.metric("GST (18%)", f"₹ {premium_excl_gst * GST_RATE:,.2f}")
        with m3:
            st.metric("Rate per Lakh (Incl GST)", f"₹ {premium_incl_gst:,.2f}")
    except Exception as e:
        st.error(f"Error: {e}")

st.divider()

st.subheader("📂 Upload Member Data for Bulk Calculation")
uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx"])

if uploaded_file is not None:
    try:
        df = pd.read_excel(uploaded_file)
        df.columns = df.columns.str.strip()

        st.subheader("Uploaded Data Preview")
        st.dataframe(df.head())

        required_columns = ['Loan Account No.', 'Name of Primary Loan borrower', 'Mobile No', 'Sum Assured']
        for col in required_columns:
            if col not in df.columns:
                df[col] = ""

        df['Sum Assured'] = pd.to_numeric(df['Sum Assured'], errors='coerce').fillna(0)

        if 'MAIN MEMBER AGE' not in df.columns:
            df['MAIN MEMBER AGE'] = age
        if 'Loan Outstanding Amount' not in df.columns:
            df['Loan Outstanding Amount'] = 0

        df_rates, tenure_map = load_rate_table(life_type, loan_type)

        def calc_premium(row):
            try:
                member_age = int(row['MAIN MEMBER AGE']) if pd.notna(row['MAIN MEMBER AGE']) and row['MAIN MEMBER AGE'] != 0 else age
                r = get_rate(df_rates, tenure_map, member_age, tenure)
                return (row['Sum Assured'] / 100000) * r
            except:
                try:
                    r = get_rate(df_rates, tenure_map, age, tenure)
                    return (row['Sum Assured'] / 100000) * r
                except:
                    return 0

        df['Premium Excl GST'] = df.apply(calc_premium, axis=1)
        df['Premium + GST'] = df['Premium Excl GST'] * (1 + GST_RATE)

        output_columns = [
            'Loan Account No.', 'Name of Primary Loan borrower', 'Mobile No',
            'MAIN MEMBER AGE', 'Sum Assured', 'Premium Excl GST', 'Premium + GST'
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
