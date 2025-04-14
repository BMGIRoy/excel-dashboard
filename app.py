import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime

# Set page configuration
st.set_page_config(
    page_title="Excel Dashboard",
    page_icon="📊",
    layout="wide"
)

# App title and description
st.title("Excel Dashboard")
st.markdown("Upload your Excel file to generate visualizations from the Contracts and BillBook sheets.")

# File uploader
uploaded_file = st.file_uploader("Choose an Excel file", type="xlsx")

if uploaded_file is not None:
    # Load the Excel file
    try:
        # Read the excel file
        excel_file = pd.ExcelFile(uploaded_file)
        
        # Check if required sheets exist
        required_sheets = ["Contracts", "BillBook"]
        missing_sheets = [sheet for sheet in required_sheets if sheet not in excel_file.sheet_names]
        
        if missing_sheets:
            st.error(f"The following required sheets are missing: {', '.join(missing_sheets)}")
        else:
            st.success("File uploaded successfully! Found both required sheets.")
            
            # Read the sheets with specified columns and use the first row as headers
            # For Contracts sheet, columns A to O (0 to 14)
            contracts_df = pd.read_excel(excel_file, sheet_name="Contracts", usecols=range(15), header=0)
            
            # For BillBook sheet, columns CM to CX (90 to 101)
            billbook_df = pd.read_excel(excel_file, sheet_name="BillBook", usecols=range(90, 102), header=0)
            
            # Display column names for debugging
            st.write("Contracts sheet columns:", contracts_df.columns.tolist())
            st.write("BillBook sheet columns:", billbook_df.columns.tolist())
            
            # Define column mappings based on user information
            # For Contracts sheet
            business_head_col = "BH"  # Column D is labeled "BH"
            client_name_col = "Client Name" if "Client Name" in contracts_df.columns else "A"
            po_balance_col = "PO Balance" if "PO Balance" in contracts_df.columns else "O"
            po_value_col = "Total PO Value" if "Total PO Value" in contracts_df.columns else "I"
            
            # For BillBook sheet
            bh_col = "Business Head"  # Column CV is labeled "Business Head"
            consultant_col = "Consultant Name" if "Consultant Name" in billbook_df.columns else "CN"
            client_col = "Client Name" if "Client Name" in billbook_df.columns else "CO"
            billed_amount_col = "Billed Amount" if "Billed Amount" in billbook_df.columns else "CQ"
            deductions_col = "Deductions" if "Deductions" in billbook_df.columns else "CR"
            net_amount_col = "Net Amount" if "Net Amount" in billbook_df.columns else "CS"
            month_col = "month" if "month" in billbook_df.columns else "CU"
            quarter_col = "Quarter" if "Quarter" in billbook_df.columns else "CT"
            
            # Create tabs for different analyses
            tab1, tab2 = st.tabs(["Contracts Analysis", "BillBook Analysis"])
            
            with tab1:
                st.header("Contracts Analysis")
                
                try:
                    # Business Head filter for Contracts
                    business_heads_contracts = contracts_df[business_head_col].unique().tolist()
                    selected_bh_contracts = st.selectbox(
                        f"Select Business Head",
                        ["All"] + business_heads_contracts
                    )
                    
                    # Filter data based on selection
                    if selected_bh_contracts != "All":
                        filtered_contracts_df = contracts_df[contracts_df[business_head_col] == selected_bh_contracts]
                    else:
                        filtered_contracts_df = contracts_df
                    
                    # Create two columns for charts
                    col1, col2 = st.columns(2)
                    
                    with col1:
                        # PO Balance by Client
                        po_balance_by_client = filtered_contracts_df.groupby(client_name_col)[po_balance_col].sum().reset_index()
                        po_balance_by_client = po_balance_by_client.sort_values(po_balance_col, ascending=False)
                        
                        fig1 = px.bar(
                            po_balance_by_client,
                            x=client_name_col,
                            y=po_balance_col,
                            title="PO Balance by Client",
                            color_discrete_sequence=["#3498db"]
                        )
                        fig1.update_layout(xaxis_title="Client", yaxis_title="PO Balance")
                        st.plotly_chart(fig1, use_container_width=True)
                    
                    with col2:
                        # PO Balance by Business Head
                        if selected_bh_contracts == "All":
                            po_balance_by_bh = contracts_df.groupby(business_head_col)[po_balance_col].sum().reset_index()
                            po_balance_by_bh = po_balance_by_bh.sort_values(po_balance_col, ascending=False)
                            
                            fig2 = px.bar(
                                po_balance_by_bh,
                                x=business_head_col,
                                y=po_balance_col,
                                title="PO Balance by Business Head",
                                color_discrete_sequence=["#2ecc71"]
                            )
                            fig2.update_layout(xaxis_title="Business Head", yaxis_title="PO Balance")
                            st.plotly_chart(fig2, use_container_width=True)
                        else:
                            # Show Total PO Value vs PO Balance for selected Business Head
                            st.subheader(f"Total PO Value vs PO Balance for {selected_bh_contracts}")
                            total_po_value = filtered_contracts_df[po_value_col].sum()
                            total_po_balance = filtered_contracts_df[po_balance_col].sum()
                            
                            fig2 = go.Figure()
                            fig2.add_trace(go.Bar(
                                x=["Total PO Value", "PO Balance"],
                                y=[total_po_value, total_po_balance],
                                marker_color=["#2ecc71", "#e74c3c"]
                            ))
                            st.plotly_chart(fig2, use_container_width=True)
                    
                    # Summary metrics
                    st.subheader("Summary Metrics")
                    metric1, metric2, metric3 = st.columns(3)
                    
                    with metric1:
                        total_po_value = filtered_contracts_df[po_value_col].sum()
                        st.metric("Total PO Value", f"${total_po_value:,.2f}")
                    
                    with metric2:
                        total_po_balance = filtered_contracts_df[po_balance_col].sum()
                        st.metric("Total PO Balance", f"${total_po_balance:,.2f}")
                    
                    with metric3:
                        if total_po_value > 0:
                            balance_ratio = (total_po_balance / total_po_value) * 100
                            st.metric("Balance to Value Ratio", f"{balance_ratio:.2f}%")
                        else:
                            st.metric("Balance to Value Ratio", "N/A")
                except Exception as e:
                    st.error(f"Error in Contracts Analysis: {e}")
                    st.write("Please check your Excel file format and column names.")
            
            with tab2:
                st.header("BillBook Analysis")
                
                try:
                    # Time period selector
                    time_period = st.radio("Select Time Period", ["Monthly", "Quarterly"], horizontal=True)
                    
                    # Business Head filter for BillBook
                    business_heads_billbook = billbook_df[bh_col].unique().tolist()
                    selected_bh_billbook = st.selectbox(
                        "Select Business Head (BillBook)",
                        ["All"] + business_heads_billbook
                    )
                    
                    # Filter data based on selection
                    if selected_bh_billbook != "All":
                        filtered_billbook_df = billbook_df[billbook_df[bh_col] == selected_bh_billbook]
                    else:
                        filtered_billbook_df = billbook_df
                    
                    # Create two columns for charts
                    col1, col2 = st.columns(2)
                    
                    with col1:
                        # Billed Amount Trend
                        if time_period == "Monthly":
                            # Group by month
                            time_trend = filtered_billbook_df.groupby(month_col)[billed_amount_col].sum().reset_index()
                            x_axis = month_col
                            title = "Monthly Billed Amount Trend"
                        else:
                            # Group by quarter
                            time_trend = filtered_billbook_df.groupby(quarter_col)[billed_amount_col].sum().reset_index()
                            x_axis = quarter_col
                            title = "Quarterly Billed Amount Trend"
                        
                        fig3 = px.line(
                            time_trend,
                            x=x_axis,
                            y=billed_amount_col,
                            title=title,
                            markers=True,
                            color_discrete_sequence=["#9b59b6"]
                        )
                        fig3.update_layout(xaxis_title=x_axis, yaxis_title="Billed Amount")
                        st.plotly_chart(fig3, use_container_width=True)
                    
                    with col2:
                        # Billed Amount by Consultant
                        billed_by_consultant = filtered_billbook_df.groupby(consultant_col)[billed_amount_col].sum().reset_index()
                        billed_by_consultant = billed_by_consultant.sort_values(billed_amount_col, ascending=False).head(10)
                        
                        fig4 = px.bar(
                            billed_by_consultant,
                            x=consultant_col,
                            y=billed_amount_col,
                            title="Top 10 Consultants by Billed Amount",
                            color_discrete_sequence=["#f39c12"]
                        )
                        fig4.update_layout(xaxis_title="Consultant", yaxis_title="Billed Amount")
                        st.plotly_chart(fig4, use_container_width=True)
                    
                    # Summary metrics
                    st.subheader("Summary Metrics")
                    metric1, metric2, metric3 = st.columns(3)
                    
                    with metric1:
                        total_billed = filtered_billbook_df[billed_amount_col].sum()
                        st.metric("Total Billed Amount", f"${total_billed:,.2f}")
                    
                    with metric2:
                        total_deductions = filtered_billbook_df[deductions_col].sum()
                        st.metric("Total Deductions", f"${total_deductions:,.2f}")
                    
                    with metric3:
                        total_net = filtered_billbook_df[net_amount_col].sum()
                        st.metric("Total Net Amount", f"${total_net:,.2f}")
                except Exception as e:
                    st.error(f"Error in BillBook Analysis: {e}")
                    st.write("Please check your Excel file format and column names.")
    
    except Exception as e:
        st.error(f"Error processing the file: {e}")
        st.write("Please check your Excel file format and try again.")
else:
    # Instructions when no file is uploaded
    st.info("Please upload an Excel file to begin.")
    
    # Example of expected data format
    st.subheader("Expected Data Format")
    
    st.markdown("""
    ### Contracts Sheet (Columns A to O):
    - **A**: Client Name
    - **B**: Type of Work
    - **C**: PO No.
    - **D**: BH (Business Head)
    - **I**: Total PO Value
    - **O**: PO Balance
    
    ### BillBook Sheet (Columns CM to CX):
    - **CM**: Date
    - **CN**: Consultant Name
    - **CO**: Client Name
    - **CP**: Days
    - **CQ**: Billed Amount
    - **CR**: Deductions
    - **CS**: Net Amount
    - **CT**: Quarter
    - **CU**: Month
    - **CV**: Business Head
    - **CW**: Team Size
    """)

# Add footer
st.markdown("---")
st.markdown("© 2023 Excel Dashboard | Created with Streamlit")
