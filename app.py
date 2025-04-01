import streamlit as st
import openai
import PyPDF2
import pandas as pd
from io import BytesIO

st.title("PDF Extraction & Excel Export App")

# 1. Ask user to upload PDF file
uploaded_pdf = st.file_uploader("Upload a PDF file", type="pdf")

# 2. Ask user to enter OpenAI key
openai_key = st.text_input("Enter your OpenAI API key", type="password")
if openai_key:
    openai.api_key = openai_key

def extract_text_from_pdf(pdf_file):
    pdf_reader = PyPDF2.PdfReader(pdf_file)
    text = ""
    for page in pdf_reader.pages:
        text += page.extract_text() + "\n"
    return text

def query_gpt4(text, prompt):
    # Use GPT-4 to extract required details.
    # Adjust the prompt as needed if the PDF content is long.
    messages = [
        {"role": "system", "content": "You are an expert financial document analyzer."},
        {"role": "user", "content": f"Extract the following details from the provided PDF text:\n"
                                    f"a) Company Name\n"
                                    f"b) Balance Sheet\n"
                                    f"c) Profit & Loss Statement\n"
                                    f"d) Statement of Cash Flows\n\n"
                                    f"PDF Text:\n{text}\n\n"
                                    f"Please return the answer as a JSON object with keys 'company_name', 'balance_sheet', 'profit_loss', 'cash_flows'."}
    ]
    try:
        response = openai.ChatCompletion.create(
            model="o3-mini",
            messages=messages,
            temperature=0
        )
        return response.choices[0].message.content
    except Exception as e:
        st.error(f"Error during OpenAI API call: {e}")
        return None

if uploaded_pdf and openai_key:
    st.info("Extracting text from PDF...")
    # 3. Extract text from PDF file
    pdf_text = extract_text_from_pdf(uploaded_pdf)
    
    st.info("Extracting financial data using GPT-4...")
    extraction_result = query_gpt4(pdf_text, "extract details")
    
    if extraction_result:
        st.subheader("Extraction Result")
        st.text_area("GPT-4 Output", extraction_result, height=300)
        
        try:
            # Convert the GPT-4 JSON string output into a dictionary.
            import json
            result_dict = json.loads(extraction_result)
        except Exception as e:
            st.error("Failed to parse GPT-4 output as JSON. Please ensure the PDF contains the expected content.")
            result_dict = {}
        
        if result_dict:
            # 4. Transfer the extracted text to an Excel sheet with good formatting.
            # We'll create an Excel file with separate sheets for each extracted section.
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                # Company name as a single value
                df_company = pd.DataFrame({"Company Name": [result_dict.get("company_name", "Not Found")]})
                df_company.to_excel(writer, sheet_name="Company Info", index=False)
                
                # Balance Sheet
                df_bs = pd.DataFrame({"Balance Sheet": [result_dict.get("balance_sheet", "Not Found")]})
                df_bs.to_excel(writer, sheet_name="Balance Sheet", index=False)
                
                # Profit & Loss
                df_pl = pd.DataFrame({"Profit & Loss": [result_dict.get("profit_loss", "Not Found")]})
                df_pl.to_excel(writer, sheet_name="Profit & Loss", index=False)
                
                # Cash Flows
                df_cf = pd.DataFrame({"Cash Flows": [result_dict.get("cash_flows", "Not Found")]})
                df_cf.to_excel(writer, sheet_name="Cash Flows", index=False)
                
                writer.save()
                processed_data = output.getvalue()
            
            # 5. Provide the download option to excel file.
            st.download_button(
                label="Download Excel file",
                data=processed_data,
                file_name="extracted_financial_data.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
