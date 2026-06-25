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

# Extra inputs for Joint Life
if life_type == "Joint Life":
    st.markdown("### 👥 Second Person Details")
    col5, col6 = st.columns(2)
    with col5:
        age2 = st.number_input("Enter Age (Person 2)", min_value=18, max_value=65, value=30, step=1)
    with col6:
        tenure2 = st.number_input(
            "Enter Tenure (Person 2)",
            min_value=min_tenure,
            max_value=max_tenure,
            value=min_tenure,
            step=1
        )
        st.caption("📅 Tenure is in Years")

if st.button("Get Rate", type="primary"):
    try:
        df_rates, tenure_map = load_rate_table(life_type, loan_type)

        if life_type == "Single Life":
            rate = get_rate(df_rates, tenure_map, age, tenure)
            st.success(f"✅ {life_type} | {loan_type} | Age {age} | Tenure {tenure} yrs")
            st.metric("Rate", f"₹ {rate:,.2f}")

        else:  # Joint Life
            rate1 = get_rate(df_rates, tenure_map, age, tenure)
            rate2 = get_rate(df_rates, tenure_map, age2, tenure2)
            total_rate = rate1 + rate2

            st.success(f"✅ {life_type} | {loan_type}")
            st.write(f"Age1: {age} | Tenure1: {tenure} yrs → ₹ {rate1:,.2f}")
            st.write(f"Age2: {age2} | Tenure2: {tenure2} yrs → ₹ {rate2:,.2f}")
            st.metric("Rate", f"₹ {total_rate:,.2f}")

    except Exception as e:
        st.error(f"Error: {e}")
