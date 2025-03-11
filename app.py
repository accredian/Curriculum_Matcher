
__import__('pysqlite3')
import sys
sys.modules['sqlite3'] = sys.modules.pop('pysqlite3')
import streamlit as st
import os
import json
from google.oauth2 import service_account
import gspread
from crewai import Agent, Task, Crew, Process
from crewai.tools import BaseTool
import pdfplumber
from typing import Any

# Set OpenAI environment variables from Streamlit secrets
os.environ["OPENAI_API_KEY"] = st.secrets["general"]["OPENAI_API_KEY"]
os.environ["OPENAI_MODEL_NAME"] = st.secrets["general"]["OPENAI_MODEL_NAME"]


class PDFExtractorTool(BaseTool):
    """
    Tool to extract curriculum items from a PDF file.
    """
    name: str = "PDF Extractor"
    description: str = "Extracts curriculum items from a PDF file."
    pdf_file: Any  # File-like object from Streamlit uploader

    def _run(self):
        with pdfplumber.open(self.pdf_file) as pdf:
            text = ""
            for page in pdf.pages:
                text += page.extract_text() or ""
        start = text.find("Curriculum:") + len("Curriculum:")
        end = text.find("End of Curriculum", start)
        if end == -1:
            end = len(text)
        curriculum_text = text[start:end].strip()
        curriculum_items = [item.strip() for item in curriculum_text.split("\n") if item.strip()]
        print("Extracted from PDF:", curriculum_items)
        return curriculum_items

def fetch_gsheet_data(credentials, spreadsheet_name: str, worksheet_name: str):
    """
    Fetch data from a Google Sheet using provided credentials.
    """
    try:
        client = gspread.authorize(credentials)
        spreadsheet = client.open(spreadsheet_name)
        sheet = spreadsheet.worksheet(worksheet_name)
        data = sheet.get_all_values()
        return data
    except Exception as e:
        st.error(f"Error fetching data from Google Sheets: {e}")
        return None


def main():
    st.title("Curriculum Comparison App")

    # File uploaders and input fields
    uploaded_pdf = st.file_uploader("Upload PDF Brochure", type="pdf")
    spreadsheet_name = st.text_input("Spreadsheet Name", value="Master_Curriculums")
    worksheet_name = st.text_input("Worksheet Name", value="CyberSecurity")

    if st.button("Run Comparison"):
        if not (uploaded_pdf and spreadsheet_name and worksheet_name):
            st.error("Please provide all required inputs: PDF, spreadsheet name, and worksheet name.")
            return

        with st.spinner("Processing..."):
            # Load credentials from uploaded JSON
            try:
                credentials_info = st.secrets["google_sheets"]
                credentials = service_account.Credentials.from_service_account_info(
                    credentials_info,
                    scopes=[
                        "https://www.googleapis.com/auth/spreadsheets",
                        "https://www.googleapis.com/auth/drive"
                    ]
                )
            except Exception as e:
                st.error(f"Error loading credentials: {e}")
                return

            # Fetch Google Sheet data
            gsheet_data = fetch_gsheet_data(credentials, spreadsheet_name, worksheet_name)
            if gsheet_data is None:
                return

            # Initialize PDF extraction tool
            pdf_tool = PDFExtractorTool(pdf_file=uploaded_pdf)

            # Define Agents
            pdf_extractor = Agent(
                role="PDF Curriculum Extractor",
                goal="Extract curriculum items from the PDF brochure",
                backstory="Expert in document parsing and retrieval-augmented generation",
                tools=[pdf_tool],
                verbose=True
            )

            gsheet_extractor = Agent(
                role="Curriculum Extractor and Formatter",
                goal="Extract curriculum (terms, topics, modules) and structure it properly from the provided Google Sheet data.",
                backstory="Specialist in accessing, processing, and structuring data",
                verbose=True
            )

            comparator = Agent(
                role="Curriculum Comparator Expert",
                goal=(
                    "Compare the brochure curriculum against the Google Sheet curriculum, using the latter as the primary reference. "
                    "Identify discrepancies including capitalization, grammar, spelling, missing or extra items, and topic counts within modules."
                ),
                backstory="Proficient in data analysis and comparison",
                verbose=True
            )

            # Define Tasks
            task1 = Task(
                description="Extract program syllabus items (terms, modules, and topics) from the PDF brochure, ignoring unrelated content.",
                agent=pdf_extractor,
                expected_output="List of program syllabus items from the PDF"
            )

            task2 = Task(
                description="Extract curriculum items (terms, modules, and topics) from the provided Google Sheet data.",
                agent=gsheet_extractor,
                expected_output="Curriculum formatted as Terms, Modules, and Topics"
            )

            task3 = Task(
                description=(""" Compare the brochure curriculum against the curriculum extrqacted by gsheet_extractor agent, using the latter as the primary reference. Conduct a thorough review to identify even the smallest discrepancies, including:
                     - Capitalization inconsistencies (e.g., title case vs. sentence case differences).
                     - Grammar and spelling errors across module names, topic descriptions, and overall content.
                     - Missing topics or modules that are present in the Google Sheet but absent in the brochure.
                     - Mismatched module or topic names, even if they appear similar but have slight variations.
                     - Extra or unnecessary words added to module or topic names in the brochure that deviate from the original curriculum.
                     - Also match the number of topics withing the modules"""),
                agent=comparator,
                expected_output="Matching values, PDF-only items, Google Sheet-only items, and discrepancy in details and the changes that needs to be done",
                context=[task1, task2]
            )

            # Create and run Crew
            crew = Crew(
                agents=[pdf_extractor, gsheet_extractor, comparator],
                tasks=[task1, task2, task3],
                process=Process.sequential
            )

            try:
                result = crew.kickoff(inputs={"Gsheet_data": gsheet_data})
                crew_output_dict = result.__dict__
                st.subheader("Comparison Results")
                st.markdown(result)
            except Exception as e:
                st.error(f"Error during processing: {e}")

    st.sidebar.header("Setup Instructions")
    st.sidebar.write("""
    1. Create a `secrets.toml` file in the `.streamlit` directory with your OpenAI credentials.
    2. Ensure you have the correct credentials JSON for Google Sheets.
    3. First Add the Curriculum to the master_curriculum sheet below.
    """)
    st.sidebar.page_link(
    "https://docs.google.com/spreadsheets/d/1hf3UCRKxpSOSblxWVbW0dQhRoM5ZSFP3AG0Vz7iNkUE/edit?usp=sharing", 
    label="Master_Curriculums", 
    icon="📊"  # Excel sheet icon
)

if __name__ == "__main__":
    main()
