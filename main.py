# streamlit_app.py
import pandas as pd
import streamlit as st
import io

# --------------------------
# Data Processing Functions
# --------------------------
def preprocess_df(df):

    # --------------------------
    # Sanitize & filter rows
    # --------------------------
    df['Hours'] = pd.to_numeric(df['Hours'], errors='coerce')

    # Remove rows with empty or zero hours + unwanted activities
    df = df[
        (~df['Activity #'].astype(str).str.contains('Lunch', na=False)) &
        (~df['Activity Name'].str.contains('Break|Work across border|Overtime', case=False, na=False)) &
        (df['Hours'].notna()) &
        (df['Hours'] != 0)
    ].reset_index(drop=True)

    # --------------------------
    # Determine existing billable column name (if any)
    # --------------------------
    possible_billable_cols = ['Billable', 'Bill.', 'Invoiceable', 'Inv.']
    billable_col = next((col for col in possible_billable_cols if col in df.columns), None)

    def get_billable(row):
        return row[billable_col] if billable_col else None

    # --------------------------
    # Build the final dataframe
    # --------------------------
    final_df = pd.DataFrame({
        'Date': pd.to_datetime(df['Date']).dt.date,
        'PersonName': df['Employee Name'],
        'PersonId': df['Employee #'].astype(str),
        'Billable': df.apply(get_billable, axis=1),
        'Hours': df['Hours'],
        'Project': df['Activity Name'],
        'ProjectId': df['Activity #'].astype(str) + df['Project #'].astype(str),
        'Client': df['Project Name'],
        'ClientId': df['Project #'].astype(str),
        'Task': '',
        'TaskId': '',
        'Description': ''
    })

    # --------------------------
    # Project name → ID overrides
    # --------------------------
    project_code_map = {
        "Sales existing customer": "4100",
        "Customer specific emission factors": "5100",
        "Knowledge transfer": "4400",
        "Customer Success Management": "6800"
    }

    for project_name, code in project_code_map.items():
        final_df.loc[
            final_df['Project'].str.contains(project_name, case=False, na=False),
            'ProjectId'
        ] = code

    # --------------------------
    # Default billable rule
    # --------------------------
    final_df['Billable'] = final_df['Billable'].apply(
        lambda x: "TRUE" if x > 0 else "FALSE"
    )

    # --------------------------
    # Override: mapped projects are NON-billable
    # --------------------------
    final_df.loc[
        final_df['Project'].str.contains(
            '|'.join(project_code_map.keys()),
            case=False,
            na=False
        ),
        'Billable'
    ] = "FALSE"

    # --------------------------
    # Business rules
    # --------------------------
    non_billable_projects = {"1002", "1001", "4"}
    non_billable_activities = {"44", "14", "15", "16", "17", "41"}

    # Extract Activity # (first numeric block of ProjectId)
    activity_ids = final_df['ProjectId'].str.extract(r'^(\d+)')[0]

    final_df.loc[
        final_df['ClientId'].isin(non_billable_projects) |
        activity_ids.isin(non_billable_activities),
        'Billable'
    ] = "FALSE"

    return final_df


# --------------------------
# Streamlit UI
# --------------------------
st.set_page_config(page_title="Excel Filter App", layout="wide")
st.title("Excel Filter and Preview Tool")

uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx", "xls"])
if uploaded_file:
    df = pd.read_excel(uploaded_file)
    final_df = preprocess_df(df)

    # --------------------------
    # Person selection
    # --------------------------
    st.subheader("Select Persons")
    all_persons = (
        final_df[['PersonId', 'PersonName']]
        .drop_duplicates()
        .sort_values('PersonId')
    )

    unchecked_ids = {2112, 4014, 5009, 2102, 2100, 2107, 4131, 4013, 1007, 1020}
    force_checked_ids = {1008, 1145}

    selected_persons = []
    for _, row in all_persons.iterrows():
        pid = int(row['PersonId'])
        default_checked = not (
            (str(pid).startswith(('1', '9')) or pid in unchecked_ids)
            and pid not in force_checked_ids
        )

        if st.checkbox(f"{row['PersonName']} ({row['PersonId']})", value=default_checked):
            selected_persons.append(row['PersonId'])

    filtered_df = final_df[final_df['PersonId'].isin(selected_persons)]

    # --------------------------
    # Preview
    # --------------------------
    st.subheader("Preview (first 100 rows)")
    st.dataframe(filtered_df.head(100), use_container_width=True)

    # --------------------------
    # Export
    # --------------------------
    
    # Create a buffer to hold Excel data
    output = io.BytesIO()
    
    # Export DataFrame to Excel
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        filtered_df.to_excel(writer, index=False, sheet_name='Filtered Data')
    
    # Reset pointer to start
    output.seek(0)
    
    # Download button for Excel
    st.download_button(
        label="Export Full Filtered Data to Excel",
        data=output,
        file_name="filtered_data.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


