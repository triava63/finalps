import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime
import requests
from io import BytesIO
import office365.runtime.auth.authentication_context as auth_ctx

def init_session_state():
    if 'authenticated' not in st.session_state:
        st.session_state.authenticated = False

init_session_state()

def authenticate_sharepoint(username, password):
    try:
        site_url = "https://share.amazon.com/sites/MMAA-HOTO"
        auth_context = auth_ctx.AuthenticationContext(url=site_url)
        if auth_context.acquire_token_for_user(username, password):
            return True
        return False
    except Exception as e:
        st.error(f"Authentication Error: {e}")
        return False

def get_sharepoint_excel():
    try:
        file_id = "e659cda6-67fe-45a3-a405-0a25b2dd0c26"
        site_url = "https://share.amazon.com/sites/MMAA-HOTO"
        file_url = f"{site_url}/_layouts/15/download.aspx?SourceUrl=/sites/MMAA-HOTO/Shared%20Documents/PS%20TPH%20Check%20Sheet/VLOOKUP.xlsx"
        
        response = requests.get(file_url)
        if response.status_code == 200:
            return pd.ExcelFile(BytesIO(response.content))
        else:
            st.error(f"Failed to fetch SharePoint file: {response.status_code}")
            return None
    except Exception as e:
        st.error(f"Error accessing SharePoint file: {e}")
        return None

def get_tracking_ids_from_sheet(df, selected_date):
    selected_date = pd.to_datetime(selected_date)
    
    column_pairs = [
        {'date': 'DATE', 'tracking': 'TRACKING ID1'},
        {'date': 'DATE.1', 'tracking': 'TRACKING ID2'},
        {'date': 'DATE.2', 'tracking': 'TRACKING ID3'},
        {'date': 'DATE.3', 'tracking': 'TRACKING ID4'}
    ]
    
    all_tracking_ids = []
    
    for pair in column_pairs:
        date_col = pair['date']
        tracking_col = pair['tracking']
        
        if date_col in df.columns and tracking_col in df.columns:
            df[date_col] = pd.to_datetime(df[date_col])
            date_mask = df[date_col].dt.date == selected_date.date()
            tracking_ids = df[date_mask][tracking_col].astype(str).str.strip()
            tracking_ids = tracking_ids[tracking_ids.notna() & (tracking_ids != 'nan')].unique().tolist()
            all_tracking_ids.extend(tracking_ids)
    
    return list(set(all_tracking_ids))

def process_excel_sheet(excel_file, sheet_name, selected_date):
    try:
        df = pd.read_excel(excel_file, sheet_name=sheet_name)
        tracking_ids = []
        
        night_shift_row = df[df.iloc[:, 0] == 'NIGHT SHIFT'].index
        
        if len(night_shift_row) > 0:
            day_df = df.iloc[:night_shift_row[0]]
            day_ids = get_tracking_ids_from_sheet(day_df, selected_date)
            tracking_ids.extend(day_ids)
            
            night_df = df.iloc[night_shift_row[0]+1:]
            night_ids = get_tracking_ids_from_sheet(night_df, selected_date)
            tracking_ids.extend(night_ids)
            
            st.write(f"Day shift tracking IDs found: {len(day_ids)}")
            st.write(f"Night shift tracking IDs found: {len(night_ids)}")
        else:
            tracking_ids.extend(get_tracking_ids_from_sheet(df, selected_date))
        
        return list(set(tracking_ids))
    except Exception as e:
        st.error(f"Error processing sheet {sheet_name}: {e}")
        return []

def compare_tracking_ids(sharepoint_tracking_ids, df_container):
    uploaded_tracking_ids = df_container['container_label'].astype(str).str.strip().unique().tolist()
    
    matched_ids = [tid for tid in sharepoint_tracking_ids if tid in uploaded_tracking_ids]
    unmatched_ids = [tid for tid in sharepoint_tracking_ids if tid not in uploaded_tracking_ids]
    
    total_ids = len(sharepoint_tracking_ids)
    matched_count = len(matched_ids)
    compliance_percentage = (matched_count / total_ids * 100) if total_ids > 0 else 0
    
    return {
        'total_ids': total_ids,
        'matched_count': matched_count,
        'unmatched_count': len(unmatched_ids),
        'compliance_percentage': compliance_percentage,
        'matched_ids': matched_ids,
        'unmatched_ids': unmatched_ids
    }

def create_compliance_chart(results):
    fig = px.pie(
        values=[results['matched_count'], results['unmatched_count']],
        names=['Matched', 'Unmatched'],
        title=f'Compliance Chart (Total Tracking IDs: {results["total_ids"]})',
        color_discrete_sequence=['#00CC96', '#EF553B']
    )
    fig.update_traces(textposition='inside', textinfo='percent+label')
    return fig

def display_results(results):
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("Total Tracking IDs", results['total_ids'])
    with col2:
        st.metric("Matched IDs", results['matched_count'])
    with col3:
        st.metric("Unmatched IDs", results['unmatched_count'])
    
    st.subheader("Compliance")
    st.progress(results['compliance_percentage'] / 100)
    st.write(f"Compliance: {results['compliance_percentage']:.2f}%")
    
    st.plotly_chart(create_compliance_chart(results))
    
    col1, col2 = st.columns(2)
    with col1:
        st.subheader("✅ Matched Tracking IDs")
        if results['matched_ids']:
            st.write(pd.DataFrame(results['matched_ids'], columns=['Tracking ID']))
        else:
            st.write("No matches found")
    
    with col2:
        st.subheader("❌ Unmatched Tracking IDs")
        if results['unmatched_ids']:
            st.write(pd.DataFrame(results['unmatched_ids'], columns=['Tracking ID']))
        else:
            st.write("No unmatched IDs")

def main():
    st.title("HOTO Tracking ID Matcher")

    with st.sidebar:
        st.header("Authentication")
        if not st.session_state.authenticated:
            username = st.text_input("Amazon Email", help="Enter your @amazon.com email")
            password = st.text_input("Password", type="password")
            if st.button("Login"):
                if authenticate_sharepoint(username, password):
                    st.session_state.authenticated = True
                    st.success("Successfully authenticated!")
                else:
                    st.error("Authentication failed!")
        else:
            st.success("You are logged in!")
            if st.button("Logout"):
                st.session_state.authenticated = False
                st.experimental_rerun()

    if st.session_state.authenticated:
        selected_date = st.date_input("Select Date for Comparison")
        container_file = st.file_uploader("Upload file with container_label column", type=["csv", "xlsx"])

        if container_file:
            try:
                if container_file.name.endswith('.csv'):
                    df_container = pd.read_csv(container_file)
                else:
                    df_container = pd.read_excel(container_file)
                
                st.write("Container file preview:")
                st.dataframe(df_container.head())

                excel_file = get_sharepoint_excel()
                
                if excel_file:
                    date_str = selected_date.strftime("%dTH %B").upper()
                    relevant_sheets = [sheet for sheet in excel_file.sheet_names 
                                    if date_str in sheet.upper()]
                    
                    if not relevant_sheets:
                        st.warning(f"No sheets found for date: {date_str}")
                    else:
                        st.write(f"Processing sheets: {relevant_sheets}")
                        
                        all_tracking_ids = []
                        for sheet in relevant_sheets:
                            tracking_ids = process_excel_sheet(excel_file, sheet, selected_date)
                            all_tracking_ids.extend(tracking_ids)
                        
                        all_tracking_ids = list(set(all_tracking_ids))
                        results = compare_tracking_ids(all_tracking_ids, df_container)
                        display_results(results)

            except Exception as e:
                st.error(f"An error occurred: {e}")
                st.write("Please check your file format and column names")
    else:
        st.warning("Please login using your Amazon credentials in the sidebar.")

if __name__ == "__main__":
    main()
