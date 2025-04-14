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
            
            # Read the sheets with specified columns
            contracts_df = pd.read_excel(excel_file, sheet_name="Contracts", usecols=range(15))  # Columns A to O
            billbook_df = pd.read_excel(excel_file, sheet_name="BillBook", usecols=range(90, 102))  # Columns CM to CX
            
            # Create tabs for different analyses
            tab1, tab2 = st.tabs(["Contracts Analysis", "BillBook Analysis"])
            
            with tab1:
                st.header("Contracts Analysis")
                
                # Business Head filter for Contracts
                business_heads_contracts = contracts_df["Business Head"].unique().tolist()
                selected_bh_contracts = st.selectbox(
                    "Select Business Head (Contracts)",
                    ["All"] + business_heads_contracts
                )
                
                # Filter data based on selection
                if selected_bh_contracts != "All":
                    filtered_contracts_df = contracts_df[contracts_df["Business Head"] == selected_bh_contracts]
                else:
                    filtered_contracts_df = contracts_df
                
                # Create two columns for charts
                col1, col2 = st.columns(2)
                
                with col1:
                    # PO Balance by Client
                    po_balance_by_client = filtered_contracts_df.groupby("Client Name")["PO Balance"].sum().reset_index()
                    po_balance_by_client = po_balance_by_client.sort_values("PO Balance", ascending=False)
                    
                    fig1 = px.bar(
                        po_balance_by_client,
                        x="Client Name",
                        y="PO Balance",
                        title="PO Balance by Client",
                        color_discrete_sequence=["#3498db"]
                    )
                    fig1.update_layout(xaxis_title="Client", yaxis_title="PO Balance")
                    st.plotly_chart(fig1, use_container_width=True)
                
                with col2:
                    # PO Balance by Business Head
                    if selected_bh_contracts == "All":
                        po_balance_by_bh = contracts_df.groupby("Business Head")["PO Balance"].sum().reset_index()
                        po_balance_by_bh = po_balance_by_bh.sort_values("PO Balance", ascending=False)
                        
                        fig2 = px.bar(
                            po_balance_by_bh,
                            x="Business Head",
                            y="PO Balance",
                            title="PO Balance by Business Head",
                            color_discrete_sequence=["#2ecc71"]
                        )
                        fig2.update_layout(xaxis_title="Business Head", yaxis_title="PO Balance")
                        st.plotly_chart(fig2, use_container_width=True)
                    else:
                        # Show Total PO Value vs PO Balance for selected Business Head
                        st.subheader(f"Total PO Value vs PO Balance for {selected_bh_contracts}")
                        total_po_value = filtered_contracts_df["Total PO Value"].sum()
                        total_po_balance = filtered_contracts_df["PO Balance"].sum()
                        
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
                    total_po_value = filtered_contracts_df["Total PO Value"].sum()
                    st.metric("Total PO Value", f"${total_po_value:,.2f}")
                
                with metric2:
                    total_po_balance = filtered_contracts_df["PO Balance"].sum()
                    st.metric("Total PO Balance", f"${total_po_balance:,.2f}")
                
                with metric3:
                    if total_po_value > 0:
                        balance_ratio = (total_po_balance / total_po_value) * 100
                        st.metric("Balance to Value Ratio", f"{balance_ratio:.2f}%")
                    else:
                        st.metric("Balance to Value Ratio", "N/A")
            
            with tab2:
                st.header("BillBook Analysis")
                
                # Time period selector
                time_period = st.radio("Select Time Period", ["Monthly", "Quarterly"], horizontal=True)
                
                # Business Head filter for BillBook
                business_heads_billbook = billbook_df["Business Head"].unique().tolist()
                selected_bh_billbook = st.selectbox(
                    "Select Business Head (BillBook)",
                    ["All"] + business_heads_billbook
                )
                
                # Filter data based on selection
                if selected_bh_billbook != "All":
                    filtered_billbook_df = billbook_df[billbook_df["Business Head"] == selected_bh_billbook]
                else:
                    filtered_billbook_df = billbook_df
                
                # Create two columns for charts
                col1, col2 = st.columns(2)
                
                with col1:
                    # Billed Amount Trend
                    if time_period == "Monthly":
                        # Group by month
                        time_trend = filtered_billbook_df.groupby("month")["Billed Amount"].sum().reset_index()
                        x_axis = "month"
                        title = "Monthly Billed Amount Trend"
                    else:
                        # Group by quarter
                        time_trend = filtered_billbook_df.groupby("Quarter")["Billed Amount"].sum().reset_index()
                        x_axis = "Quarter"
                        title = "Quarterly Billed Amount Trend"
                    
                    fig3 = px.line(
                        time_trend,
                        x=x_axis,
                        y="Billed Amount",
                        title=title,
                        markers=True,
                        color_discrete_sequence=["#9b59b6"]
                    )
                    fig3.update_layout(xaxis_title=x_axis, yaxis_title="Billed Amount")
                    st.plotly_chart(fig3, use_container_width=True)
                
                with col2:
                    # Billed Amount by Consultant
                    billed_by_consultant = filtered_billbook_df.groupby("Consultant Name")["Billed Amount"].sum().reset_index()
                    billed_by_consultant = billed_by_consultant.sort_values("Billed Amount", ascending=False).head(10)
                    
                    fig4 = px.bar(
                        billed_by_consultant,
                        x="Consultant Name",
                        y="Billed Amount",
                        title="Top 10 Consultants by Billed Amount",
                        color_discrete_sequence=["#f39c12"]
                    )
                    fig4.update_layout(xaxis_title="Consultant", yaxis_title="Billed Amount")
                    st.plotly_chart(fig4, use_container_width=True)
                
                # Summary metrics
                st.subheader("Summary Metrics")
                metric1, metric2, metric3 = st.columns(3)
                
                with metric1:
                    total_billed = filtered_billbook_df["Billed Amount"].sum()
                    st.metric("Total Billed Amount", f"${total_billed:,.2f}")
                
                with metric2:
                    total_deductions = filtered_billbook_df["Deductions"].sum()
                    st.metric("Total Deductions", f"${total_deductions:,.2f}")
                
                with metric3:
                    total_net = filtered_billbook_df["Net Amount"].sum()
                    st.metric("Total Net Amount", f"${total_net:,.2f}")
    
    except Exception as e:
        st.error(f"Error processing the file: {e}")
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
    - **D**: Business Head
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
