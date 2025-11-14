# streamlit_app.py
import pandas as pd
import streamlit as st
import io
# --------------------------
# Data Processing Functions
# --------------------------
def preprocess_df(df):
    df['Hours'] = pd.to_numeric(df['Hours'], errors='coerce')
    df = df[
        (~df['Activity #'].str.contains('Lunch', na=False)) &
        (~df['Activity Name'].str.contains('Break|Work across border|Overtime', case=False, na=False)) &
        (df['Hours'].notna()) &
        (df['Hours'] != 0)
    ].reset_index(drop=True)

    def get_billable(row):
        for col in ['Billable', 'Bill.', 'Invoiceable', 'Inv.']:
            if col in df.columns:
                return row[col]
        return None

    final_df = pd.DataFrame({
        'Date': pd.to_datetime(df['Date']).dt.date,
        'PersonName': df['Employee Name'],
        'PersonId': df['Employee #'].astype(str),
        'Billable': df.apply(get_billable, axis=1),
        'Hours': df['Hours'],
        'Project': df['Activity Name'],
        'ProjectId': df['Activity #'].astype(str) + df['Project #'].astype(str),
        'Client': df['Project Name'],
        'ClientId': df['Project #'],
        'Task': '',
        'TaskId': '',
        'Description': ''
    })

    project_code_map = {
        "Sales existing customer": "4100",
        "Customer specific emission factors": "5100",
        "Knowledge transfer": "4400",
        "Customer Success Management": "6800"
    }
    for project_name, code in project_code_map.items():
        final_df.loc[final_df['Project'].str.contains(project_name, case=False, na=False), 'ProjectId'] = code

    final_df['Billable'] = final_df['Hours'].apply(lambda x: "TRUE" if x > 0 else "FALSE")
    return final_df

# --------------------------
# Streamlit UI
# --------------------------
st.set_page_config(page_title="Excel Filter App", layout="wide")
st.title("Excel Filter and Preview Tool")

# Upload Excel file
uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx", "xls"])
if uploaded_file:
    df = pd.read_excel(uploaded_file)
    final_df = preprocess_df(df)

    # Person selection
    st.subheader("Select Persons")
    all_persons = final_df[['PersonId', 'PersonName']].drop_duplicates().sort_values('PersonId')
    unchecked_ids = {2112, 4014, 5009, 2102, 2100, 2107, 4131, 4013, 1007, 1020}
    force_checked_ids = {1008, 1145}

    selected_persons = []
    for _, row in all_persons.iterrows():
        pid = int(row['PersonId'])
        default_checked = not ((str(pid).startswith(('1','9')) or pid in unchecked_ids) and pid not in force_checked_ids)
        if st.checkbox(f"{row['PersonName']} ({row['PersonId']})", value=default_checked):
            selected_persons.append(row['PersonId'])

    filtered_df = final_df[final_df['PersonId'].isin(selected_persons)]

    # Preview first 100 rows
    st.subheader("Preview (first 100 rows)")
    st.dataframe(filtered_df.head(100), use_container_width=True)

    # Export full dataset
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        filtered_df.to_excel(writer, index=False)
    
    processed_data = output.getvalue()
    
    st.download_button(
        label="Export Full Filtered Data to Excel",
        data=processed_data,
        file_name="filtered_data.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

