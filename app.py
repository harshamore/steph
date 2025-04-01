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
import json
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
    try:
        pdf_reader = PyPDF2.PdfReader(pdf_file)
        text = ""
        total_pages = len(pdf_reader.pages)
        
        # Show progress bar for PDF extraction
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        for page_num in range(total_pages):
            try:
                page_text = pdf_reader.pages[page_num].extract_text()
                text += page_text + "\n\n--- Page Break ---\n\n"  # Add page breaks to help with context
                
                # Update progress
                progress = (page_num + 1) / total_pages
                progress_bar.progress(progress)
                status_text.text(f"Extracting text from page {page_num + 1}/{total_pages}")
                
            except Exception as e:
                st.warning(f"Error extracting text from page {page_num + 1}: {str(e)}")
                continue
        
        progress_bar.progress(1.0)
        status_text.text("PDF text extraction complete!")
        time.sleep(0.5)  # Give time for the user to see the completion
        status_text.empty()
        progress_bar.empty()
        
        return text
    except Exception as e:
        st.error(f"Failed to extract text from PDF: {str(e)}")
        st.info("Please try a different PDF file or check if the file is correctly formatted.")
        return ""

# Function to split text into chunks for processing
def split_text_into_chunks(text, max_chunk_size=12000, overlap=1000):
    chunks = []
    start = 0
    text_length = len(text)
    
    while start < text_length:
        end = min(start + max_chunk_size, text_length)
        
        # If this is not the first chunk, include some overlap
        if start > 0:
            start = start - overlap
        
        chunks.append(text[start:end])
        start = end
    
    return chunks

# Function to extract financial information using GPT-4o
def extract_financial_info(text, api_key):
    client = OpenAI(api_key=api_key)
    
    # Split the text into manageable chunks
    text_chunks = split_text_into_chunks(text)
    
    all_responses = []
    financial_data = {
        "company_name": None,
        "balance_sheet": {"assets": [], "liabilities": [], "equity": []},
        "profit_loss": {"revenue": [], "expenses": [], "profit": []},
        "cash_flows": {"operating": [], "investing": [], "financing": []}
    }
    
    # First, extract the company name from the first chunk
    company_name_prompt = f"""
    Extract ONLY the company name from this financial report text. 
    Return ONLY a JSON object with a single key 'company_name' and the value as the company name.
    
    Text:
    {text_chunks[0][:5000]}
    """
    
    try:
        company_response = client.chat.completions.create(
            model="gpt-4o",
            messages=[
                {"role": "system", "content": "You are a financial analyst that extracts company names from reports. Return only valid JSON."},
                {"role": "user", "content": company_name_prompt}
            ],
            response_format={"type": "json_object"}
        )
        
        company_data = json.loads(company_response.choices[0].message.content)
        financial_data["company_name"] = company_data.get("company_name", "Unknown Company")
    except Exception as e:
        st.warning(f"Could not extract company name: {str(e)}")
        financial_data["company_name"] = "Unknown Company"
    
    # Process each financial statement type separately with specific prompts
    
    # 1. Extract Balance Sheet
    balance_sheet_prompt = """
    Extract the Balance Sheet data from this financial report text.
    Focus ONLY on the Balance Sheet section (also sometimes called Statement of Financial Position).
    
    Return a JSON with the following structure:
    {
      "assets": [
        {"name": "Asset Name 1", "amount": "$XXX,XXX"},
        {"name": "Asset Name 2", "amount": "$XXX,XXX"}
      ],
      "liabilities": [
        {"name": "Liability Name 1", "amount": "$XXX,XXX"},
        {"name": "Liability Name 2", "amount": "$XXX,XXX"}
      ],
      "equity": [
        {"name": "Equity Item 1", "amount": "$XXX,XXX"},
        {"name": "Equity Item 2", "amount": "$XXX,XXX"}
      ]
    }
    
    Common assets include: Cash, Accounts Receivable, Inventory, Property/Plant/Equipment, Investments, Intangible Assets
    Common liabilities include: Accounts Payable, Short-term Debt, Long-term Debt, Accrued Expenses
    Common equity items include: Common Stock, Retained Earnings, Additional Paid-in Capital
    
    Look for sections that specifically mention "ASSETS", "LIABILITIES", and "EQUITY" or "SHAREHOLDERS' EQUITY".
    
    Text:
    """
    
    # 2. Extract Profit & Loss Statement
    profit_loss_prompt = """
    Extract the Profit & Loss Statement data from this financial report text.
    Focus ONLY on the Profit & Loss section (also sometimes called Income Statement, Statement of Operations, or Statement of Earnings).
    
    Return a JSON with the following structure:
    {
      "revenue": [
        {"name": "Revenue Item 1", "amount": "$XXX,XXX"},
        {"name": "Revenue Item 2", "amount": "$XXX,XXX"}
      ],
      "expenses": [
        {"name": "Expense Item 1", "amount": "$XXX,XXX"},
        {"name": "Expense Item 2", "amount": "$XXX,XXX"}
      ],
      "profit": [
        {"name": "Gross Profit", "amount": "$XXX,XXX"},
        {"name": "Operating Income", "amount": "$XXX,XXX"},
        {"name": "Net Income", "amount": "$XXX,XXX"}
      ]
    }
    
    Common revenue items include: Sales, Service Revenue, Interest Income
    Common expenses include: Cost of Goods Sold, Operating Expenses, SG&A, R&D, Interest Expense
    Common profit items include: Gross Profit, Operating Income, Income Before Tax, Net Income
    
    Look for sections that specifically mention "REVENUE", "INCOME", "EXPENSES", "EARNINGS", or "PROFIT".
    
    Text:
    """
    
    # 3. Extract Cash Flow Statement
    cash_flow_prompt = """
    Extract the Statement of Cash Flows data from this financial report text.
    Focus ONLY on the Cash Flow section.
    
    Return a JSON with the following structure:
    {
      "operating": [
        {"name": "Operating Item 1", "amount": "$XXX,XXX"},
        {"name": "Operating Item 2", "amount": "$XXX,XXX"}
      ],
      "investing": [
        {"name": "Investing Item 1", "amount": "$XXX,XXX"},
        {"name": "Investing Item 2", "amount": "$XXX,XXX"}
      ],
      "financing": [
        {"name": "Financing Item 1", "amount": "$XXX,XXX"},
        {"name": "Financing Item 2", "amount": "$XXX,XXX"}
      ]
    }
    
    Common operating items include: Net Income, Depreciation, Changes in Working Capital
    Common investing items include: Capital Expenditures, Acquisitions, Investment Purchases/Sales
    Common financing items include: Debt Issuance/Repayment, Dividends, Share Repurchases
    
    Look for sections that specifically mention "CASH FLOWS", "OPERATING ACTIVITIES", "INVESTING ACTIVITIES", or "FINANCING ACTIVITIES".
    
    Text:
    """
    
    # Process each statement type with each chunk
    financial_statements = [
        {"name": "balance_sheet", "prompt": balance_sheet_prompt, "sections": ["assets", "liabilities", "equity"]},
        {"name": "profit_loss", "prompt": profit_loss_prompt, "sections": ["revenue", "expenses", "profit"]},
        {"name": "cash_flows", "prompt": cash_flow_prompt, "sections": ["operating", "investing", "financing"]}
    ]
    
    for statement in financial_statements:
        statement_name = statement["name"]
        statement_prompt = statement["prompt"]
        
        # Try to find the statement in each chunk
        for i, chunk in enumerate(text_chunks):
            chunk_prompt = statement_prompt + chunk
            
            try:
                response = client.chat.completions.create(
                    model="gpt-4o",
                    messages=[
                        {"role": "system", "content": f"You are a financial analyst that extracts {statement_name.replace('_', ' ')} data from reports. Return only valid JSON."},
                        {"role": "user", "content": chunk_prompt}
                    ],
                    response_format={"type": "json_object"}
                )
                
                statement_data = json.loads(response.choices[0].message.content)
                
                # Check if we got meaningful data
                has_data = False
                for section in statement["sections"]:
                    if section in statement_data and len(statement_data[section]) > 0:
                        has_data = True
                        # Add the data to our consolidated results
                        financial_data[statement_name][section].extend(statement_data[section])
                
                # If we found data, we can stop processing this statement
                if has_data:
                    break
                    
            except Exception as e:
                st.warning(f"Error extracting {statement_name} from chunk {i+1}: {str(e)}")
                continue
    
    # Convert to a clean JSON string
    return json.dumps(financial_data, indent=2)

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
                    elif uploaded_file is not None and api_key:
                # Process the PDF and extract financial information
                try:
                    # Extract text from PDF
                    with st.expander("Step 1: Extracting text from PDF", expanded=True):
                        st.info("Extracting text content from your PDF file...")
                        pdf_text = extract_text_from_pdf(uploaded_file)
                        
                        if not pdf_text:
                            st.error("Failed to extract text from the PDF. Please try a different file.")
                            st.stop()
                        
                        # Show a sample of the extracted text
                        st.success("Text extraction complete!")
                        with st.expander("Preview extracted text"):
                            st.text_area("Sample of extracted text", value=pdf_text[:1500] + "...", height=200)
                    
                    # Use OpenAI to extract financial information
                    with st.expander("Step 2: Analyzing financial data with GPT-4o", expanded=True):
                        st.info("Using GPT-4o to identify and extract financial statements...")
                        
                        progress_bar = st.progress(0)
                        status_text = st.empty()
                        
                        status_text.text("Extracting company information...")
                        progress_bar.progress(0.1)
                        
                        financial_info = extract_financial_info(pdf_text, api_key)
                        
                        # Parse the JSON response
                        financial_data = json.loads(financial_info)
                        
                        progress_bar.progress(1.0)
                        status_text.text("Analysis complete!")
                        time.sleep(0.5)
                        status_text.empty()
                        progress_bar.empty()
                        
                        # Check if we have data in each section
                        has_balance_sheet = any(len(financial_data.get('balance_sheet', {}).get(section, [])) > 0 
                                               for section in ['assets', 'liabilities', 'equity'])
                        has_profit_loss = any(len(financial_data.get('profit_loss', {}).get(section, [])) > 0 
                                             for section in ['revenue', 'expenses', 'profit'])
                        has_cash_flows = any(len(financial_data.get('cash_flows', {}).get(section, [])) > 0 
                                            for section in ['operating', 'investing', 'financing'])
                        
                        # Success or warning based on extraction results
                        if has_balance_sheet and has_profit_loss and has_cash_flows:
                            st.success("All financial statements extracted successfully!")
                        else:
                            warning_msg = "Partial extraction: "
                            if not has_balance_sheet:
                                warning_msg += "Balance Sheet was not found. "
                            if not has_profit_loss:
                                warning_msg += "Profit & Loss was not found. "
                            if not has_cash_flows:
                                warning_msg += "Cash Flow Statement was not found. "
                            st.warning(warning_msg + "This may affect the quality of the Excel output.")
                    
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
