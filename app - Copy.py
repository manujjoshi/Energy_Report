import streamlit as st
import xlwings as xw
import os
from io import BytesIO
from PIL import ImageGrab
import pandas as pd
import requests

###################################################################
# # --- Configuration ---
# EXCEL_PATH = 'Energy Report.xlsm'
# # file_id = "1o-9owvuE0p4tS6RgIuSCz9MuWYz7fC_i"  # replace with your actual ID
# # url = f"https://drive.google.com/uc?id={file_id}"
# # output_path = 'Energy Reports.xlsm'

# # response = requests.get(url)
# # with open(output_path, 'wb') as f:
# #     f.write(response.content)

###################################################################

EXCEL_PATH = 'Energy Reports.xlsm'
st.set_page_config(layout="wide", page_title="Energy Report")
st.title("Energy Report Data Entry")

BUILDING_TYPE_OPTIONS = [
    "Convenience Store",
    "Data center",
    "Data center estimates",
    "Hospital (general medical and surgical)",
    "Hotel",
    "K-12 school",
    "Medical office",
    "Multifamily housing",
    "bank branch, financial",
    "Parking",
    "dormitory",
    "and wholesale",
    "Senior living community",
    "Single-family homes",
    "Supermarket/grocery store",
    "Swimming pool",
    "Vehicle dealerships",
    "distribution center, non-",
    "Wastewater treatment plant",
    "Worship facility"
]

# --- Cell Mappings ---
# This dictionary maps each form field to the *exact cell* where its value should be written.
# It will NOT modify any other cell, including the keyword/label cells to the left.
INPUT_MAPPING = {
    'Entity': 'D3',
    'Facility': 'D4',
    'Address': 'D6',
    'State/Province': 'D9',
    'Zip Code': 'D11',
    'Gross Floor Area': ['E13', 'L12'],
    'Building Type': 'L4',
    'State Code': 'L9',
    'Percent of the building Cooled': 'L13',
    'Heating Degree Days': 'L14',
    'Cooling Degree Days': 'L15',
    'High School': 'L16',
    'Number of workers': 'L17',
    'Open weekends': 'L18',
    'Presence of cooking': 'L19'
}

# --- Functions ---
def update_excel_file(input_data):
    """
    Writes inputs directly to the main Excel file.
    Using xlwings ensures that when the file is saved, Excel's calculation engine
    updates all formulas and linked cells across all tabs (e.g., Report, Calculation, etc.),
    preserving the workbook's interconnected functionality.
    """
    if not os.path.exists(EXCEL_PATH):
        st.error(f"Excel file not found at {EXCEL_PATH}")
        return False

    try:
        with xw.App(visible=False) as app:
            wb = app.books.open(EXCEL_PATH)
            ws = wb.sheets['Form']

            # This loop writes the new values to the Form tab.
            for key, cell in INPUT_MAPPING.items():
                value = input_data.get(key)
                if isinstance(cell, list):
                    for c in cell:
                        ws.range(c).value = value
                else:
                    ws.range(cell).value = value
            
            # This is the crucial step. Saving the workbook triggers Excel
            # to recalculate all dependent formulas, lookups, and linked cells
            # across the entire workbook before the file is closed.
            wb.save()
            wb.close()
        return True
    except Exception as e:
        st.error(f"An error occurred: {e}")
        st.info("Please ensure Excel is not open and the file is not locked.")
        return False

def capture_excel_range(filepath, sheet_name, range_str):
    app = xw.App(visible=False)
    wb = app.books.open(filepath)
    sheet = wb.sheets[sheet_name]
    rng = sheet.range(range_str)
    rng.api.CopyPicture(Format=2)  # bitmap
    img = ImageGrab.grabclipboard()
    wb.close()
    app.quit()

    buf = BytesIO()
    img.save(buf, format="PNG")
    buf.seek(0)
    return buf

def extract_bill_data(bill_path):
    df = pd.read_excel(bill_path, sheet_name=0, header=None)
    # Find the header row for electrical table
    header_row_candidates = df.apply(lambda row: row.astype(str).str.strip().str.lower().str.contains('month').any(), axis=1)
    if not header_row_candidates.any():
        raise ValueError("Could not find a row containing 'Month' in the uploaded bill. Please check the file format.")
    month_row = header_row_candidates[header_row_candidates].index[0]
    # Get the header row as a list of strings
    header_row = df.iloc[month_row].astype(str).str.strip().str.lower().tolist()
    # Find column indices for electrical
    col_map = {}
    for idx, col in enumerate(header_row):
        if 'kw/h' in col or 'kwh' in col:
            col_map['kwh'] = idx
        elif col == 'kw':
            col_map['kw'] = idx
        elif col == 'cost' and 'electric' in col:
            col_map['electric_cost'] = idx
        elif col == 'cost':
            # If there are two 'cost' columns, the first is for electric, second for gas
            if 'electric_cost' not in col_map:
                col_map['electric_cost'] = idx
            else:
                col_map['gas_cost'] = idx
        elif 'mcf' in col:
            col_map['mcf'] = idx

    # Extract 12 months of data for electrical and gas
    data_rows = df.iloc[month_row+1:month_row+13]
    # Electrical: D41 (kWh), E41 (kW), F41 (kW), H41 (Electric Cost)
    electrical = pd.DataFrame({
        'kwh': data_rows.iloc[:, col_map['kwh']].values,
        'kw1': data_rows.iloc[:, col_map['kw']].values,
        'kw2': data_rows.iloc[:, col_map['kw']].values,  # If you have two kW columns, adjust here
        'electric_cost': data_rows.iloc[:, col_map['electric_cost']].values
    })
    # Gas: K41 (mcf), L41 (cost)
    gas = pd.DataFrame({
        'mcf': data_rows.iloc[:, col_map['mcf']].values,
        'gas_cost': data_rows.iloc[:, col_map['gas_cost']].values if 'gas_cost' in col_map else [None]*12
    })
    return electrical, gas

def update_template_with_bill(electrical, gas, template_path):
    with xw.App(visible=False) as app:
        wb = app.books.open(template_path)
        ws = wb.sheets['Form']

        # Electrical: D41 (kWh), E41 (kW), F41 (kW), H41 (Electric Cost)
        ws.range('D41').options(index=False, header=False).value = electrical['kwh'].values.reshape(-1, 1)
        ws.range('E41').options(index=False, header=False).value = electrical['kw1'].values.reshape(-1, 1)
        ws.range('F41').options(index=False, header=False).value = electrical['kw2'].values.reshape(-1, 1)
        ws.range('H41').options(index=False, header=False).value = electrical['electric_cost'].values.reshape(-1, 1)

        # Gas: K41 (mcf), L41 (cost)
        ws.range('K41').options(index=False, header=False).value = gas['mcf'].values.reshape(-1, 1)
        ws.range('L41').options(index=False, header=False).value = gas['gas_cost'].values.reshape(-1, 1)

        wb.save()
        wb.close()

# --- Input Form ---
with st.expander("Enter New Data (Update Template)", expanded=True):
    with st.form("data_entry_form"):
        col1, col2 = st.columns(2)
        with col1:
            st.subheader("General Information")
            entity = st.text_input("Entity")
            facility = st.text_input("Facility")
            address = st.text_area("Address")
            state_province = st.text_input("State/Province")
            zip_code = st.text_input("Zip Code")
            gross_floor_area = st.number_input("Gross Floor Area (Sq.Ft)", min_value=0)

        with col2:
            st.subheader("Building & Climate Details")
            building_type = st.selectbox("Building Type", BUILDING_TYPE_OPTIONS)
            state_code = st.text_input("State Code")
            percent_cooled = st.number_input("Percent of the building Cooled (%)", min_value=0, max_value=100)
            heating_degree_days = st.number_input("Heating Degree Days", min_value=0)
            cooling_degree_days = st.number_input("Cooling Degree Days", min_value=0)
            presence_of_cooking = st.number_input("Presence of cooking")
            
        st.subheader("Operations")
        c1, c2, c3, c4 = st.columns(4)
        with c1:
            high_school = st.selectbox("High School", ["Yes", "No"])
        with c2:
            open_weekends = st.selectbox("Open weekends", ["Yes", "No", ""])
        with c3:
            cooking = st.selectbox("Cooking", ["Yes", "No"])
        
        number_of_workers = st.number_input("Number of workers", min_value=0)

        submitted = st.form_submit_button("Update Excel File")

if submitted:
    with st.spinner("Updating Excel file..."):
        input_data = {
            'Entity': entity, 'Facility': facility, 'Address': address, 'State/Province': state_province, 'Zip Code': zip_code,
            'Building Type': building_type, 'State Code': state_code, 'Gross Floor Area': gross_floor_area,
            'Percent of the building Cooled': percent_cooled / 100.0, 'Heating Degree Days': heating_degree_days,
            'Cooling Degree Days': cooling_degree_days, 'High School': high_school, 'Number of workers': number_of_workers,
            'Open weekends': open_weekends, 'Presence of cooking': presence_of_cooking,
        }
        
        if update_excel_file(input_data):
            st.success("Excel file has been updated successfully!")

uploaded_file = st.file_uploader("📁 Upload your Excel bill (.xlsm or .xlsx)", type=["xlsm", "xlsx"])
if uploaded_file:
    temp_path = "temp.xlsm"
    with open(temp_path, "wb") as f:
        f.write(uploaded_file.getbuffer())
    try:
        # Extract data from bill
        electrical, gas = extract_bill_data(temp_path)
        # Update template with bill data
        update_template_with_bill(electrical, gas, EXCEL_PATH)
        # Show two snapshots from the Report tab
        st.subheader("📸 Snapshot: Columns A to AA")
        img1 = capture_excel_range(EXCEL_PATH, "Report", "A1:AA50")
        st.image(img1)
        st.download_button("⬇️ Download A–AA Snapshot", img1, file_name="table1_snapshot.png", mime="image/png")

        st.subheader("📸 Snapshot: Columns AB to BH")
        img2 = capture_excel_range(EXCEL_PATH, "Report", "AB1:BH50")
        st.image(img2)
        st.download_button("⬇️ Download AB–BH Snapshot", img2, file_name="table2_snapshot.png", mime="image/png")

    except Exception as e:
        st.error(f"❌ Error capturing snapshot: {e}")
    # Download processed Report tab as Excel
    # ... code to extract and download Report tab only ... 