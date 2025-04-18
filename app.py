















import streamlit as st
from bs4 import BeautifulSoup
from selenium import webdriver
import pandas as pd
import os
import json
import time

# Function to sanitize sheet names for Excel
def sanitize_sheet_name(name):
    invalid_chars = ['/', '\\', '?', '*', '[', ']']
    for char in invalid_chars:
        name = name.replace(char, '_')
    return name[:31]  # Excel sheet name limit is 31 characters

# Function to extract table data
def extract_table_data(table):
    """
    Extracts data from a table and returns it as a list of rows.
    :param table: The BeautifulSoup object representing the table.
    :return: List of rows containing data.
    """
    data = []
    rows = table.find_all('tr')
    for row in rows:
        cols = row.find_all('td')  # Find all table cells in the row
        cols = [col.text.strip() for col in cols]  # Extract text from each cell
        if cols:  # Ignore empty rows
            data.append(cols)
    return data

# Main function to extract data from a single URL
def extract_hospital_data(url, sections_to_extract=None):
    driver = webdriver.Chrome()
    driver.get(url)
    time.sleep(5)  # Wait for the page to load
    soup = BeautifulSoup(driver.page_source, 'html.parser')

    hospital_data = {}

    # Default sections to extract
    sections = {
        "basicInfo": "Basic Information",
        "administrativeInfo": "Administrative Information",
        "contactInformation": "Contact Information",
        "facilityInfo": "Facility Infrastructure",
        "departmentInfos": "Department Info",
        "buildingInfos": "Building Info",
        "bedInfo": "Bed Info",
        "landInfo": "Land Information",
        "permissionInfo": "Permission_Approval Information",
        "availableMajorServices": "Available Major Services"
    }

    # Override sections if provided
    if sections_to_extract:
        sections = {k: v for k, v in sections.items() if v in sections_to_extract}

    for section_id, section_name in sections.items():
        try:
            if section_name == "Basic Information":
                basic_info_table = soup.find('table', {'class': 'facility-registry-table'})
                if basic_info_table:
                    rows = basic_info_table.find_all('tr')
                    data = [[col.text.strip() for col in row.find_all('td')] for row in rows if row.find_all('td')]
                    hospital_data[section_name] = data
                else:
                    hospital_data[section_name] = {"Error": "Basic Information table not found."}
            else:
                section_div = soup.find('div', id=section_id)
                if section_div:
                    table = section_div.find('table', class_='facility-registry-table')
                    if table:
                        empty_message = table.find('td', class_='dataTables_empty')
                        if empty_message:
                            hospital_data[section_name] = {"Message": empty_message.text.strip()}
                        else:
                            hospital_data[section_name] = extract_table_data(table)
                    else:
                        hospital_data[section_name] = {"Error": "Table not found."}
                else:
                    hospital_data[section_name] = {"Error": "Section not found."}
        except Exception as e:
            hospital_data[section_name] = {"Error": str(e)}

    driver.quit()
    return hospital_data

# Function to extract HRM Status data
def extract_hrm_status_data(url):
    driver = webdriver.Chrome()
    driver.get(url)
    time.sleep(5)  # Wait for the page to load
    soup = BeautifulSoup(driver.page_source, 'html.parser')

    hospital_data = {}

    # Locate the main container div
    posts_table_wrapper = soup.find('div', {'class': 'dataTables_wrapper'})
    if posts_table_wrapper:
        # Locate the table within the wrapper
        table = posts_table_wrapper.find('table', {'id': 'postsTable'})
        if table:
            # Extract data from the table
            hospital_data["HRM Status"] = extract_table_data(table)
        else:
            hospital_data["HRM Status"] = {"Error": "HRM Status table not found."}
    else:
        hospital_data["HRM Status"] = {"Error": "Div containing the HRM Status table not found."}

    driver.quit()
    return hospital_data

# Function to save all data for one hospital to an Excel file
def save_hospital_to_excel(hospital_id, hospital_data, output_dir="output"):
    os.makedirs(output_dir, exist_ok=True)
    output_file = os.path.join(output_dir, f"{hospital_id}.xlsx")

    if not hospital_data:
        st.error(f"No data extracted for {hospital_id}. Skipping Excel file creation.")
        return

    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        has_visible_sheets = False  # Track whether at least one sheet is added
        for section_name, section_data in hospital_data.items():
            sanitized_name = sanitize_sheet_name(section_name)
            if isinstance(section_data, dict):  # For error messages or empty tables
                if "Error" in section_data or "Message" in section_data:
                    continue  # Skip adding sheets for errors or empty messages
                df = pd.DataFrame(list(section_data.items()), columns=["Field", "Value"])
            else:  # For normal table data
                df = pd.DataFrame(section_data)

            if not df.empty:  # Only add non-empty sheets
                df.to_excel(writer, sheet_name=sanitized_name, index=False)
                has_visible_sheets = True

        if not has_visible_sheets:
            st.error(f"No visible sheets to save for {hospital_id}. Skipping Excel file creation.")
            return

    st.success(f"All data for {hospital_id} saved to {output_file}")

# Streamlit App
st.title("Hospital Data Extractor")

# Input options
uploaded_file = st.file_uploader("Upload a JSON file with hospital URLs", type=["json"])

# Output Directory
output_dir = "output"
os.makedirs(output_dir, exist_ok=True)

if uploaded_file is not None:
    # Load the JSON file
    hospitals = json.load(uploaded_file)
    st.write(f"Found {len(hospitals)} hospitals in the uploaded file.")

    if st.button("Extract Data"):
        total_hospitals = len(hospitals)
        progress_bar = st.progress(0)
        status_text = st.empty()

        for i, hospital in enumerate(hospitals):
            hospital_id = hospital.get("id", f"hospital_{i+1}")
            hospital_urls = hospital.get("urls")  # Expecting a list of URLs
            if not hospital_urls or len(hospital_urls) != 2:
                st.error(f"Incomplete URLs provided for hospital ID: {hospital_id}. Skipping...")
                continue

            # Initialize combined data for this hospital
            combined_hospital_data = {}

            for j, url in enumerate(hospital_urls):
                status_text.text(f"Processing {hospital_id} - URL {j+1}/{len(hospital_urls)}...")
                if j == 0:  # Detailed Information URL
                    hospital_data = extract_hospital_data(url)
                elif j == 1:  # HRM Status URL
                    hospital_data = extract_hrm_status_data(url)

                combined_hospital_data.update(hospital_data)

            # Save the combined data to Excel
            save_hospital_to_excel(hospital_id, combined_hospital_data, output_dir)
            progress_bar.progress((i + 1) / total_hospitals)

        st.success("Data extraction complete!")
        st.markdown(f"Download your files from the `{output_dir}` folder.")
else:
    st.info("Please upload a JSON file to proceed.")