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


def find_column(df, target):
    """
    Find a column matching `target` exactly (case-insensitive, ignoring spaces).
    e.g. target='Age1' matches a column literally named 'Age1' / 'age 1' / ' AGE1 '
    but will NOT match 'Age' or 'Age2'.
    """
    target_norm = target.strip().lower().replace(" ", "")
    for col in df.columns:
        col_norm = str(col).strip().lower().replace(" ", "")
        if col_norm == target_norm:
            return col
    return None


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
if loan_type == "Home Loan":
    min_tenure, max_tenure = 5, 25
else:
    min_tenure, max_tenure = 2, 10

with col4:
    tenure = st.number_input(
        "Enter Tenure",
        min_value=min_tenure,
        max_value=max_tenure,
        value=min_tenure,
        step=1
    )
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

if life_type == "Single Life":
    st.markdown("Your Excel must have at least: **Name**, **Age**, **Tenure** (in years)")
else:
    st.markdown("Your Excel must have at least: **Name1**, **Age1**, **Tenure1**, **Name2**, **Age2**, **Tenure2** (in years)")

st.warning("⚠️ Please make sure you have selected **Life Type** and **Loan Type** above before uploading your Excel file.")

uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx"])

if uploaded_file is not None:
    try:
        df = pd.read_excel(uploaded_file)
        df.columns = [str(c).strip() for c in df.columns]

        st.subheader("Uploaded Data Preview")
        st.dataframe(df.head())

        # Apply tenure limits based on loan type
        if loan_type == "Home Loan":
            min_t, max_t = 5, 25
        else:
            min_t, max_t = 2, 10

        df_rates, tenure_map = load_rate_table(life_type, loan_type)

        # ============================================
        # SINGLE LIFE
        # ============================================
        if life_type == "Single Life":
            name_col = find_column(df, "Name")
            age_col = find_column(df, "Age")
            tenure_col = find_column(df, "Tenure")

            if not name_col or not age_col or not tenure_col:
                raise ValueError("Excel must contain mandatory columns: Name, Age, Tenure")

            df[age_col] = pd.to_numeric(df[age_col], errors='coerce')
            df[tenure_col] = pd.to_numeric(df[tenure_col], errors='coerce')

            # Auto-detect months vs years
            if df[tenure_col].dropna().median() > 30:
                st.info("ℹ️ Tenure values look like months — auto-converting to years.")
                df[tenure_col] = (df[tenure_col] / 12).round(0).astype('Int64')
            else:
                df[tenure_col] = df[tenure_col].round(0).astype('Int64')

            df[age_col] = df[age_col].round(0).astype('Int64')
            df[tenure_col] = df[tenure_col].clip(lower=min_t, upper=max_t)

            premiums = []
            statuses = []
            for _, row in df.iterrows():
                try:
                    r_age = int(row[age_col])
                    r_tenure = int(row[tenure_col])
                    r = get_rate(df_rates, tenure_map, r_age, r_tenure)
                    premiums.append(round(r, 2))
                    statuses.append("✅")
                except Exception as e:
                    premiums.append(None)
                    statuses.append(f"❌ {e}")

            df["Premium"] = premiums
            df["Status"] = statuses

            # Reorder: Name, Age, Tenure, Premium, ...extras (Status kept as an extra)
            core_cols = [name_col, age_col, tenure_col, "Premium"]
            extra_cols = [c for c in df.columns if c not in core_cols]
            df = df[core_cols + extra_cols]

            # Grand total
            total_premium = pd.to_numeric(pd.Series(premiums), errors='coerce').sum()

            st.metric("💰 Grand Total Premium", f"₹ {total_premium:,.2f}")

            st.subheader("Rate Lookup Output")
            st.dataframe(df, use_container_width=True)

            # Build output with TOTAL PREMIUM row at bottom
            total_row = {c: "" for c in df.columns}
            total_row[name_col] = "TOTAL PREMIUM"
            total_row["Premium"] = round(total_premium, 2)
            df_out = pd.concat([df, pd.DataFrame([total_row])], ignore_index=True)

            output_file = "Rate_Output.xlsx"
            df_out.to_excel(output_file, index=False)

            with open(output_file, "rb") as file:
                st.download_button(
                    label="⬇ Download Output Excel",
                    data=file,
                    file_name=output_file,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

        # ============================================
        # JOINT LIFE
        # ============================================
        else:
            name1_col = find_column(df, "Name1")
            age1_col = find_column(df, "Age1")
            tenure1_col = find_column(df, "Tenure1")
            name2_col = find_column(df, "Name2")
            age2_col = find_column(df, "Age2")
            tenure2_col = find_column(df, "Tenure2")

            missing = [n for n, c in [
                ("Name1", name1_col), ("Age1", age1_col), ("Tenure1", tenure1_col),
                ("Name2", name2_col), ("Age2", age2_col), ("Tenure2", tenure2_col)
            ] if c is None]

            if missing:
                raise ValueError(
                    "Excel must contain mandatory columns: Name1, Age1, Tenure1, Name2, Age2, Tenure2. "
                    f"Missing: {', '.join(missing)}"
                )

            for c in [age1_col, tenure1_col, age2_col, tenure2_col]:
                df[c] = pd.to_numeric(df[c], errors='coerce')

            # Auto-detect months vs years (checked per tenure column independently)
            for tcol, label in [(tenure1_col, "Tenure1"), (tenure2_col, "Tenure2")]:
                if df[tcol].dropna().median() > 30:
                    st.info(f"ℹ️ {label} values look like months — auto-converting to years.")
                    df[tcol] = (df[tcol] / 12).round(0).astype('Int64')
                else:
                    df[tcol] = df[tcol].round(0).astype('Int64')

            df[age1_col] = df[age1_col].round(0).astype('Int64')
            df[age2_col] = df[age2_col].round(0).astype('Int64')
            df[tenure1_col] = df[tenure1_col].clip(lower=min_t, upper=max_t)
            df[tenure2_col] = df[tenure2_col].clip(lower=min_t, upper=max_t)

            premium1_list = []
            premium2_list = []
            total_list = []
            statuses = []

            for _, row in df.iterrows():
                row_status = "✅"
                p1 = None
                p2 = None
                try:
                    r_age1 = int(row[age1_col])
                    r_tenure1 = int(row[tenure1_col])
                    p1 = round(get_rate(df_rates, tenure_map, r_age1, r_tenure1), 2)
                except Exception as e:
                    row_status = f"❌ Person1: {e}"

                try:
                    r_age2 = int(row[age2_col])
                    r_tenure2 = int(row[tenure2_col])
                    p2 = round(get_rate(df_rates, tenure_map, r_age2, r_tenure2), 2)
                except Exception as e:
                    row_status = (row_status + f" | Person2: {e}") if row_status != "✅" else f"❌ Person2: {e}"

                premium1_list.append(p1)
                premium2_list.append(p2)
                total_list.append(round(p1 + p2, 2) if (p1 is not None and p2 is not None) else None)
                statuses.append(row_status)

            df["Premium1"] = premium1_list
            df["Premium2"] = premium2_list
            df["Total Premium"] = total_list
            df["Status"] = statuses

            # Reorder: Name1, Age1, Tenure1, Premium1, Name2, Age2, Tenure2, Premium2, Total Premium, ...extras
            core_cols = [
                name1_col, age1_col, tenure1_col, "Premium1",
                name2_col, age2_col, tenure2_col, "Premium2",
                "Total Premium"
            ]
            extra_cols = [c for c in df.columns if c not in core_cols]
            df = df[core_cols + extra_cols]

            grand_total = pd.to_numeric(pd.Series(total_list), errors='coerce').sum()

            st.metric("💰 Grand Total Premium", f"₹ {grand_total:,.2f}")

            st.subheader("Rate Lookup Output")
            st.dataframe(df, use_container_width=True)

            # Build output with TOTAL PREMIUM row at bottom
            total_row = {c: "" for c in df.columns}
            total_row[name1_col] = "TOTAL PREMIUM"
            total_row["Total Premium"] = round(grand_total, 2)
            df_out = pd.concat([df, pd.DataFrame([total_row])], ignore_index=True)

            output_file = "Rate_Output.xlsx"
            df_out.to_excel(output_file, index=False)

            with open(output_file, "rb") as file:
                st.download_button(
                    label="⬇ Download Output Excel",
                    data=file,
                    file_name=output_file,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

    except Exception as e:
        st.error(f"Error: {e}")
