import streamlit as st
import pandas as pd

# ============================================
# PAGE CONFIG
# ============================================

st.set_page_config(
    page_title="Aviva Rate Calculator",
    page_icon="📊",
    layout="centered"
)

# ============================================
# TITLE
# ============================================

st.title("📊 Aviva Rate Calculator")

st.markdown(
    "Select Life Type, Loan Type, Age and Tenure to fetch the applicable rate."
)

# ============================================
# FILE PATHS
# ============================================

SINGLE_LIFE_FILE = "Aviva Single life.xlsx"
JOINT_LIFE_FILE = "Aviva Joint life.xlsx"

# ============================================
# INPUTS
# ============================================

life_type = st.selectbox(
    "Life Type",
    ["Single Life", "Joint Life"]
)

loan_type = st.selectbox(
    "Loan Type",
    ["Home Loan", "LAP"]
)

age = st.number_input(
    "Age",
    min_value=18,
    max_value=100,
    step=1
)

st.markdown("**Tenure (in Years)**")

tenure = st.number_input(
    "Tenure",
    min_value=1,
    step=1
)

# ============================================
# CALCULATE BUTTON
# ============================================

if st.button("Get Rate"):

    try:

        # ============================================
        # SELECT FILE
        # ============================================

        if life_type == "Single Life":
            file_path = SINGLE_LIFE_FILE
        else:
            file_path = JOINT_LIFE_FILE

        # ============================================
        # SELECT SHEET
        # ============================================

        if loan_type == "Home Loan":
            sheet_name = "Home Loan"
        else:
            sheet_name = "LAP"

        # ============================================
        # READ SHEET
        # ============================================

        df = pd.read_excel(
            file_path,
            sheet_name=sheet_name
        )

        # ============================================
        # CLEAN DATA
        # ============================================

        df.columns = [str(col).strip() for col in df.columns]

        # First column assumed to be AGE
        age_column = df.columns[0]

        df[age_column] = pd.to_numeric(
            df[age_column],
            errors="coerce"
        )

        # ============================================
        # AGE VALIDATION
        # ============================================

        if age not in df[age_column].values:

            st.error(
                "Age/Tenure not available in rate card."
            )

        else:

            row = df[
                df[age_column] == age
            ]

            tenure_column = str(tenure)

            # ============================================
            # TENURE VALIDATION
            # ============================================

            if tenure_column not in df.columns:

                st.error(
                    "Age/Tenure not available in rate card."
                )

            else:

                rate = row.iloc[0][tenure_column]

                st.success("Rate Found Successfully")

                st.metric(
                    label="Applicable Rate",
                    value=f"{rate}"
                )

    except Exception as e:

        st.error(
            f"Error: {e}"
        )