import streamlit as st
import PyPDF2
import pandas as pd
import io
import re
from openai import OpenAI
import tempfile
import base64
from io import BytesIO
import time
import os
from typing import Dict, List, Tuple, Any

# Set page configuration
st.set_page_config(
    page_title="Financial Statement Extractor",
    page_icon="📊",
    layout="wide"
)

# Create a container for the header
header = st.container()
with header:
    st.title("Financial Statement Extractor")
    st.markdown(
        """
        Upload a financial PDF document, enter your OpenAI API key, and extract the following:
        - Company Name
        - Balance Sheet
        - Profit & Loss Statement
        - Statement of Cash Flows
        
        The extracted data will be formatted and saved to an Excel file for download.
        """
    )
    st.markdown("---")

# Function to extract text from PDF
def extract_text_from_pdf(pdf_file):
    pdf_reader = PyPDF2.PdfReader(pdf_file)
    text = ""
    for page_num in range(len(pdf_reader.pages)):
        text += pdf_reader.pages[page_num].extract_text()
    return text

# Function to extract financial information using GPT-4o
def extract_financial_info(text, api_key):
    client = OpenAI(api_key=api_key)
    
    prompt = f"""
    I have a financial report in text format. Please extract the following information:
    
    1. Company Name
    2. Balance Sheet (with all assets, liabilities, and equity items)
    3. Profit & Loss Statement (with all revenue, expenses, and profit items)
    4. Statement of Cash Flows (with operating, investing, and financing activities)
    
    Please format the response in a structured JSON format with the following keys:
    - company_name: The name of the company
    - balance_sheet: A dictionary with 'assets', 'liabilities', and 'equity' as keys, each containing a list of items with 'name' and 'amount'
    - profit_loss: A dictionary with 'revenue', 'expenses', and 'profit' as keys, each containing a list of items with 'name' and 'amount'
    - cash_flows: A dictionary with 'operating', 'investing', and 'financing' as keys, each containing a list of items with 'name' and 'amount'
    
    Here is the text from the financial report:
    {text[:15000]}  # Limiting to 15000 characters to avoid token limits
    """
    
    response = client.chat.completions.create(
        model="o3-mini",
        messages=[
            {"role": "system", "content": "You are a financial analyst AI that extracts structured financial information from reports."},
            {"role": "user", "content": prompt}
        ],
        response_format={"type": "json_object"}
    )
    
    return response.choices[0].message.content

# Function to create Excel file with multiple sheets
def create_excel_file(financial_data):
    import json
    
    # Parse the JSON string if it's not already a dictionary
    if isinstance(financial_data, str):
        financial_data = json.loads(financial_data)
    
    # Create a BytesIO object to store the Excel file
    excel_buffer = BytesIO()
    
    # Create an Excel writer object
    with pd.ExcelWriter(excel_buffer, engine='xlsxwriter') as writer:
        workbook = writer.book
        
        # Create formats
        header_format = workbook.add_format({
            'bold': True,
            'font_size': 12,
            'bg_color': '#4F81BD',
            'font_color': 'white',
            'border': 1,
            'align': 'center',
            'valign': 'vcenter'
        })
        
        subheader_format = workbook.add_format({
            'bold': True,
            'font_size': 11,
            'bg_color': '#D0D8E8',
            'border': 1
        })
        
        item_format = workbook.add_format({
            'border': 1
        })
        
        amount_format = workbook.add_format({
            'border': 1,
            'num_format': '#,##0.00'
        })
        
        title_format = workbook.add_format({
            'bold': True,
            'font_size': 14,
            'font_color': '#1F497D',
            'align': 'center',
            'valign': 'vcenter'
        })
        
        # Create Summary sheet
        summary_df = pd.DataFrame({
            'Information': ['Company Name'],
            'Value': [financial_data.get('company_name', 'Not Available')]
        })
        
        summary_sheet = writer.book.add_worksheet('Summary')
        summary_sheet.write(0, 0, 'Financial Statement Summary', title_format)
        summary_sheet.merge_range('A1:B1', 'Financial Statement Summary', title_format)
        summary_sheet.write(2, 0, 'Information', header_format)
        summary_sheet.write(2, 1, 'Value', header_format)
        summary_sheet.write(3, 0, 'Company Name', item_format)
        summary_sheet.write(3, 1, financial_data.get('company_name', 'Not Available'), item_format)
        summary_sheet.write(4, 0, 'Report Date', item_format)
        summary_sheet.write(4, 1, 'As extracted', item_format)
        
        # Set column widths
        summary_sheet.set_column('A:A', 20)
        summary_sheet.set_column('B:B', 40)
        
        # Create Balance Sheet
        balance_sheet = financial_data.get('balance_sheet', {})
        
        # Create a DataFrame for the balance sheet
        bs_rows = []
        
        # Add Assets
        bs_rows.append(['ASSETS', ''])
        for asset in balance_sheet.get('assets', []):
            bs_rows.append([asset.get('name', ''), asset.get('amount', '')])
        
        # Add Liabilities
        bs_rows.append(['', ''])
        bs_rows.append(['LIABILITIES', ''])
        for liability in balance_sheet.get('liabilities', []):
            bs_rows.append([liability.get('name', ''), liability.get('amount', '')])
        
        # Add Equity
        bs_rows.append(['', ''])
        bs_rows.append(['EQUITY', ''])
        for equity in balance_sheet.get('equity', []):
            bs_rows.append([equity.get('name', ''), equity.get('amount', '')])
        
        bs_df = pd.DataFrame(bs_rows, columns=['Item', 'Amount'])
        
        # Write Balance Sheet to Excel
        bs_sheet = writer.book.add_worksheet('Balance Sheet')
        bs_sheet.write(0, 0, 'Balance Sheet', title_format)
        bs_sheet.merge_range('A1:B1', 'Balance Sheet', title_format)
        bs_sheet.write(2, 0, 'Item', header_format)
        bs_sheet.write(2, 1, 'Amount', header_format)
        
        row = 3
        for i, (item, amount) in enumerate(bs_rows):
            if item in ['ASSETS', 'LIABILITIES', 'EQUITY']:
                bs_sheet.write(row, 0, item, subheader_format)
                bs_sheet.write(row, 1, '', subheader_format)
            else:
                bs_sheet.write(row, 0, item, item_format)
                bs_sheet.write(row, 1, amount, amount_format)
            row += 1
        
        # Set column widths
        bs_sheet.set_column('A:A', 40)
        bs_sheet.set_column('B:B', 20)
        
        # Create Profit & Loss Statement
        profit_loss = financial_data.get('profit_loss', {})
        
        # Create a DataFrame for the profit & loss
        pl_rows = []
        
        # Add Revenue
        pl_rows.append(['REVENUE', ''])
        for revenue in profit_loss.get('revenue', []):
            pl_rows.append([revenue.get('name', ''), revenue.get('amount', '')])
        
        # Add Expenses
        pl_rows.append(['', ''])
        pl_rows.append(['EXPENSES', ''])
        for expense in profit_loss.get('expenses', []):
            pl_rows.append([expense.get('name', ''), expense.get('amount', '')])
        
        # Add Profit
        pl_rows.append(['', ''])
        pl_rows.append(['PROFIT/LOSS', ''])
        for profit in profit_loss.get('profit', []):
            pl_rows.append([profit.get('name', ''), profit.get('amount', '')])
        
        # Write Profit & Loss to Excel
        pl_sheet = writer.book.add_worksheet('Profit & Loss')
        pl_sheet.write(0, 0, 'Profit & Loss Statement', title_format)
        pl_sheet.merge_range('A1:B1', 'Profit & Loss Statement', title_format)
        pl_sheet.write(2, 0, 'Item', header_format)
        pl_sheet.write(2, 1, 'Amount', header_format)
        
        row = 3
        for i, (item, amount) in enumerate(pl_rows):
            if item in ['REVENUE', 'EXPENSES', 'PROFIT/LOSS']:
                pl_sheet.write(row, 0, item, subheader_format)
                pl_sheet.write(row, 1, '', subheader_format)
            else:
                pl_sheet.write(row, 0, item, item_format)
                pl_sheet.write(row, 1, amount, amount_format)
            row += 1
        
        # Set column widths
        pl_sheet.set_column('A:A', 40)
        pl_sheet.set_column('B:B', 20)
        
        # Create Cash Flow Statement
        cash_flows = financial_data.get('cash_flows', {})
        
        # Create a DataFrame for the cash flows
        cf_rows = []
        
        # Add Operating Activities
        cf_rows.append(['OPERATING ACTIVITIES', ''])
        for op in cash_flows.get('operating', []):
            cf_rows.append([op.get('name', ''), op.get('amount', '')])
        
        # Add Investing Activities
        cf_rows.append(['', ''])
        cf_rows.append(['INVESTING ACTIVITIES', ''])
        for inv in cash_flows.get('investing', []):
            cf_rows.append([inv.get('name', ''), inv.get('amount', '')])
        
        # Add Financing Activities
        cf_rows.append(['', ''])
        cf_rows.append(['FINANCING ACTIVITIES', ''])
        for fin in cash_flows.get('financing', []):
            cf_rows.append([fin.get('name', ''), fin.get('amount', '')])
        
        # Write Cash Flows to Excel
        cf_sheet = writer.book.add_worksheet('Cash Flows')
        cf_sheet.write(0, 0, 'Statement of Cash Flows', title_format)
        cf_sheet.merge_range('A1:B1', 'Statement of Cash Flows', title_format)
        cf_sheet.write(2, 0, 'Item', header_format)
        cf_sheet.write(2, 1, 'Amount', header_format)
        
        row = 3
        for i, (item, amount) in enumerate(cf_rows):
            if item in ['OPERATING ACTIVITIES', 'INVESTING ACTIVITIES', 'FINANCING ACTIVITIES']:
                cf_sheet.write(row, 0, item, subheader_format)
                cf_sheet.write(row, 1, '', subheader_format)
            else:
                cf_sheet.write(row, 0, item, item_format)
                cf_sheet.write(row, 1, amount, amount_format)
            row += 1
        
        # Set column widths
        cf_sheet.set_column('A:A', 40)
        cf_sheet.set_column('B:B', 20)
    
    # Reset the buffer position to the beginning
    excel_buffer.seek(0)
    
    return excel_buffer

# Function to create a download link for the Excel file
def get_excel_download_link(excel_buffer, filename="financial_statements.xlsx"):
    b64 = base64.b64encode(excel_buffer.read()).decode()
    return f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="{filename}" class="download-button">Download Excel File</a>'

# Main application
def main():
    # Create sidebar for input fields
    with st.sidebar:
        st.header("Inputs")
        
        # File uploader for PDF
        uploaded_file = st.file_uploader("Upload Financial PDF", type="pdf")
        
        # Input field for OpenAI API key
        api_key = st.text_input("Enter OpenAI API Key", type="password")
        
        # Checkbox to use sample data (for testing)
        use_sample = st.checkbox("Use sample data (for testing)")
        
        # Submit button
        submit_button = st.button("Extract Financial Data")
    
    # Main content area
    main_area = st.container()
    
    with main_area:
        # Check if the form is submitted
        if submit_button:
            if use_sample:
                # Use sample data for testing
                with st.spinner("Generating sample data..."):
                    time.sleep(2)  # Simulate processing time
                    
                    sample_data = {
                        "company_name": "ABC Corporation Ltd.",
                        "balance_sheet": {
                            "assets": [
                                {"name": "Cash and Cash Equivalents", "amount": "$1,234,567"},
                                {"name": "Accounts Receivable", "amount": "$987,654"},
                                {"name": "Inventory", "amount": "$765,432"},
                                {"name": "Property, Plant and Equipment", "amount": "$4,321,098"},
                                {"name": "Intangible Assets", "amount": "$1,543,210"}
                            ],
                            "liabilities": [
                                {"name": "Accounts Payable", "amount": "$654,321"},
                                {"name": "Short-term Debt", "amount": "$432,109"},
                                {"name": "Long-term Debt", "amount": "$2,109,876"},
                                {"name": "Deferred Tax Liabilities", "amount": "$123,456"}
                            ],
                            "equity": [
                                {"name": "Common Stock", "amount": "$1,000,000"},
                                {"name": "Retained Earnings", "amount": "$4,532,199"},
                                {"name": "Treasury Stock", "amount": "$(500,000)"}
                            ]
                        },
                        "profit_loss": {
                            "revenue": [
                                {"name": "Sales Revenue", "amount": "$12,345,678"},
                                {"name": "Service Revenue", "amount": "$1,234,567"},
                                {"name": "Other Revenue", "amount": "$234,567"}
                            ],
                            "expenses": [
                                {"name": "Cost of Goods Sold", "amount": "$7,654,321"},
                                {"name": "Selling, General & Administrative", "amount": "$2,345,678"},
                                {"name": "Research & Development", "amount": "$1,234,567"},
                                {"name": "Depreciation & Amortization", "amount": "$543,210"},
                                {"name": "Interest Expense", "amount": "$321,098"}
                            ],
                            "profit": [
                                {"name": "Gross Profit", "amount": "$5,925,924"},
                                {"name": "Operating Income", "amount": "$1,802,369"},
                                {"name": "Income Before Tax", "amount": "$1,481,271"},
                                {"name": "Net Income", "amount": "$1,111,271"}
                            ]
                        },
                        "cash_flows": {
                            "operating": [
                                {"name": "Net Income", "amount": "$1,111,271"},
                                {"name": "Depreciation & Amortization", "amount": "$543,210"},
                                {"name": "Changes in Working Capital", "amount": "$(123,456)"},
                                {"name": "Net Cash from Operating Activities", "amount": "$1,531,025"}
                            ],
                            "investing": [
                                {"name": "Capital Expenditures", "amount": "$(876,543)"},
                                {"name": "Acquisitions", "amount": "$(432,109)"},
                                {"name": "Net Cash used in Investing Activities", "amount": "$(1,308,652)"}
                            ],
                            "financing": [
                                {"name": "Dividends Paid", "amount": "$(234,567)"},
                                {"name": "Debt Repayment", "amount": "$(123,456)"},
                                {"name": "Share Repurchases", "amount": "$(87,654)"},
                                {"name": "Net Cash used in Financing Activities", "amount": "$(445,677)"}
                            ]
                        }
                    }
                    
                    st.success("Sample data generated successfully!")
                    
                    # Display the extracted information
                    st.subheader("Extracted Information")
                    
                    col1, col2 = st.columns(2)
                    
                    with col1:
                        st.markdown(f"**Company Name:** {sample_data['company_name']}")
                    
                    # Create tabs for different financial statements
                    tab1, tab2, tab3 = st.tabs(["Balance Sheet", "Profit & Loss", "Cash Flows"])
                    
                    with tab1:
                        # Display Balance Sheet
                        st.markdown("### Balance Sheet")
                        
                        # Assets
                        st.markdown("#### Assets")
                        for asset in sample_data['balance_sheet']['assets']:
                            st.write(f"{asset['name']}: {asset['amount']}")
                        
                        # Liabilities
                        st.markdown("#### Liabilities")
                        for liability in sample_data['balance_sheet']['liabilities']:
                            st.write(f"{liability['name']}: {liability['amount']}")
                        
                        # Equity
                        st.markdown("#### Equity")
                        for equity in sample_data['balance_sheet']['equity']:
                            st.write(f"{equity['name']}: {equity['amount']}")
                    
                    with tab2:
                        # Display Profit & Loss
                        st.markdown("### Profit & Loss Statement")
                        
                        # Revenue
                        st.markdown("#### Revenue")
                        for revenue in sample_data['profit_loss']['revenue']:
                            st.write(f"{revenue['name']}: {revenue['amount']}")
                        
                        # Expenses
                        st.markdown("#### Expenses")
                        for expense in sample_data['profit_loss']['expenses']:
                            st.write(f"{expense['name']}: {expense['amount']}")
                        
                        # Profit
                        st.markdown("#### Profit/Loss")
                        for profit in sample_data['profit_loss']['profit']:
                            st.write(f"{profit['name']}: {profit['amount']}")
                    
                    with tab3:
                        # Display Cash Flows
                        st.markdown("### Statement of Cash Flows")
                        
                        # Operating Activities
                        st.markdown("#### Operating Activities")
                        for op in sample_data['cash_flows']['operating']:
                            st.write(f"{op['name']}: {op['amount']}")
                        
                        # Investing Activities
                        st.markdown("#### Investing Activities")
                        for inv in sample_data['cash_flows']['investing']:
                            st.write(f"{inv['name']}: {inv['amount']}")
                        
                        # Financing Activities
                        st.markdown("#### Financing Activities")
                        for fin in sample_data['cash_flows']['financing']:
                            st.write(f"{fin['name']}: {fin['amount']}")
                    
                    # Create Excel file and provide download link
                    excel_buffer = create_excel_file(sample_data)
                    
                    st.markdown("---")
                    st.subheader("Download Excel File")
                    st.markdown("Click the button below to download the formatted Excel file:")
                    
                    # Display download button
                    st.download_button(
                        label="Download Excel File",
                        data=excel_buffer,
                        file_name="financial_statements_sample.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                
            elif uploaded_file is not None and api_key:
                # Process the PDF and extract financial information
                with st.spinner("Processing PDF and extracting financial information..."):
                    try:
                        # Extract text from PDF
                        pdf_text = extract_text_from_pdf(uploaded_file)
                        
                        # Use OpenAI to extract financial information
                        financial_info = extract_financial_info(pdf_text, api_key)
                        
                        # Parse the JSON response
                        import json
                        financial_data = json.loads(financial_info)
                        
                        st.success("Financial information extracted successfully!")
                        
                        # Display the extracted information
                        st.subheader("Extracted Information")
                        
                        col1, col2 = st.columns(2)
                        
                        with col1:
                            st.markdown(f"**Company Name:** {financial_data.get('company_name', 'Not Available')}")
                        
                        # Create tabs for different financial statements
                        tab1, tab2, tab3 = st.tabs(["Balance Sheet", "Profit & Loss", "Cash Flows"])
                        
                        with tab1:
                            # Display Balance Sheet
                            st.markdown("### Balance Sheet")
                            
                            balance_sheet = financial_data.get('balance_sheet', {})
                            
                            # Assets
                            st.markdown("#### Assets")
                            for asset in balance_sheet.get('assets', []):
                                st.write(f"{asset.get('name', '')}: {asset.get('amount', '')}")
                            
                            # Liabilities
                            st.markdown("#### Liabilities")
                            for liability in balance_sheet.get('liabilities', []):
                                st.write(f"{liability.get('name', '')}: {liability.get('amount', '')}")
                            
                            # Equity
                            st.markdown("#### Equity")
                            for equity in balance_sheet.get('equity', []):
                                st.write(f"{equity.get('name', '')}: {equity.get('amount', '')}")
                        
                        with tab2:
                            # Display Profit & Loss
                            st.markdown("### Profit & Loss Statement")
                            
                            profit_loss = financial_data.get('profit_loss', {})
                            
                            # Revenue
                            st.markdown("#### Revenue")
                            for revenue in profit_loss.get('revenue', []):
                                st.write(f"{revenue.get('name', '')}: {revenue.get('amount', '')}")
                            
                            # Expenses
                            st.markdown("#### Expenses")
                            for expense in profit_loss.get('expenses', []):
                                st.write(f"{expense.get('name', '')}: {expense.get('amount', '')}")
                            
                            # Profit
                            st.markdown("#### Profit/Loss")
                            for profit in profit_loss.get('profit', []):
                                st.write(f"{profit.get('name', '')}: {profit.get('amount', '')}")
                        
                        with tab3:
                            # Display Cash Flows
                            st.markdown("### Statement of Cash Flows")
                            
                            cash_flows = financial_data.get('cash_flows', {})
                            
                            # Operating Activities
                            st.markdown("#### Operating Activities")
                            for op in cash_flows.get('operating', []):
                                st.write(f"{op.get('name', '')}: {op.get('amount', '')}")
                            
                            # Investing Activities
                            st.markdown("#### Investing Activities")
                            for inv in cash_flows.get('investing', []):
                                st.write(f"{inv.get('name', '')}: {inv.get('amount', '')}")
                            
                            # Financing Activities
                            st.markdown("#### Financing Activities")
                            for fin in cash_flows.get('financing', []):
                                st.write(f"{fin.get('name', '')}: {fin.get('amount', '')}")
                        
                        # Create Excel file and provide download link
                        excel_buffer = create_excel_file(financial_data)
                        
                        st.markdown("---")
                        st.subheader("Download Excel File")
                        st.markdown("Click the button below to download the formatted Excel file:")
                        
                        # Display download button
                        st.download_button(
                            label="Download Excel File",
                            data=excel_buffer,
                            file_name=f"{financial_data.get('company_name', 'financial_statements').replace(' ', '_')}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                        
                    except Exception as e:
                        st.error(f"An error occurred: {str(e)}")
                        st.error("Please check your API key and try again, or use a different PDF file.")
            
            else:
                if not uploaded_file:
                    st.warning("Please upload a PDF file or use the sample data option.")
                if not api_key and not use_sample:
                    st.warning("Please enter your OpenAI API key or use the sample data option.")

# Run the application
if __name__ == "__main__":
    main()
