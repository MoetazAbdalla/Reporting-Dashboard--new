import dash
from dash import dcc, html, Input, Output, dash_table
import dash_bootstrap_components as dbc
import pandas as pd
import base64
import io

app = dash.Dash(__name__, external_stylesheets=[dbc.themes.BOOTSTRAP])

# Tuition fees data
tuitionFees = {
    "Medicine (English)": {"listFee": 40000, "advanceFee": 36000, "Deposit": 20000},
    "Medicine (30% English)": {"listFee": 30000, "advanceFee": 27000, "Deposit": 15000},
    "Dentistry (English)": {"listFee": 32000, "advanceFee": 28800, "Deposit": 16000},
    "Dentistry (30% English 70% Turkish)": {"listFee": 30000, "advanceFee": 27000, "Deposit": 15000},
    "Pharmacy (English)": {"listFee": 18000, "advanceFee": 16200, "Deposit": 9000},
    "Pharmacy (Turkish)": {"listFee": 14000, "advanceFee": 12600, "Deposit": 7000},
    "Law (30% English)": {"listFee": 10000, "advanceFee": 9000, "Deposit": 5000},
    "Nursing (English)": {"listFee": 7000, "advanceFee": 6300, "Deposit": 3500},
    "Nursing (Turkish)": {"listFee": 7000, "advanceFee": 6300, "Deposit": 3500},
    "Physiotherapy and Rehabilitation (English)": {"listFee": 7000, "advanceFee": 6300, "Deposit": 3500},
    "Physiotherapy and Rehabilitation (Turkish)": {"listFee": 7000, "advanceFee": 6300, "Deposit": 3500},
    "Nutrition and Dietetics (Turkish)": {"listFee": 5500, "advanceFee": 4950, "Deposit": 2750},
    "Health Management (English)": {"listFee": 5500, "advanceFee": 4950, "Deposit": 2750},
    "Health Management (Turkish)": {"listFee": 5500, "advanceFee": 4950, "Deposit": 2750},
    "Audiology (Turkish)": {"listFee": 5500, "advanceFee": 4950, "Deposit": 2750},
    "Orthotics and Prosthetics (Turkish)": {"listFee": 5500, "advanceFee": 4950, "Deposit": 2750},
    "Child Development (Turkish)": {"listFee": 5500, "advanceFee": 4950, "Deposit": 2750},
    "Midwifery (Turkish)": {"listFee": 5500, "advanceFee": 4950, "Deposit": 2750},
    "Ergotherapy (Turkish)": {"listFee": 5500, "advanceFee": 4950, "Deposit": 2750},
    "Social Services (Turkish)": {"listFee": 5500, "advanceFee": 4950, "Deposit": 2750},
    "English Teaching 100% English": {"listFee": 5000, "advanceFee": 4500, "Deposit": 2500},
    "Speech and Language Therapy (English)": {"listFee": 5500, "advanceFee": 4950, "Deposit": 2750},
    "Speech and Language Therapy (Turkish)": {"listFee": 5500, "advanceFee": 4950, "Deposit": 2750},
    "Electrical-Electronic Engineering (English)": {"listFee": 6500, "advanceFee": 5850, "Deposit": 3250},
    "Biomedical Engineering (English)": {"listFee": 7000, "advanceFee": 6300, "Deposit": 3500},
    "Industrial Engineering (English)": {"listFee": 5500, "advanceFee": 4950, "Deposit": 2750},
    "Computer Engineering (English)": {"listFee": 7000, "advanceFee": 6300, "Deposit": 3500},
    "Civil Engineering (English)": {"listFee": 6500, "advanceFee": 5850, "Deposit": 3250},
    "Civil Engineering (Turkish)": {"listFee": 5000, "advanceFee": 4500, "Deposit": 2500},
    "Business Administration (English)": {"listFee": 5500, "advanceFee": 4950, "Deposit": 2750},
    "Economics and Finance (English)": {"listFee": 5500, "advanceFee": 4950, "Deposit": 2750},
    "International Trade and Finance (English)": {"listFee": 5500, "advanceFee": 4950, "Deposit": 2750},
    "International Trade and Finance (Turkish)": {"listFee": 5500, "advanceFee": 4950, "Deposit": 2750},
    "Management Information Systems (English)": {"listFee": 5500, "advanceFee": 4950, "Deposit": 2750},
    "Management Information Systems (Turkish)": {"listFee": 5500, "advanceFee": 4950, "Deposit": 2750},
    "Logistic Management (English)": {"listFee": 5500, "advanceFee": 4950, "Deposit": 2750},
    "Logistic Management (Turkish)": {"listFee": 5500, "advanceFee": 4950, "Deposit": 2750},
    "Human Resources Management (Turkish)": {"listFee": 5500, "advanceFee": 4950, "Deposit": 2750},
    "Aviation Management (Turkish)": {"listFee": 5500, "advanceFee": 4950, "Deposit": 2750},
    "Banking and Insurance (Turkish)": {"listFee": 5500, "advanceFee": 4950, "Deposit": 2750},
    "Psychology (English)": {"listFee": 5500, "advanceFee": 4950, "Deposit": 2750},
    "Psychology (Turkish)": {"listFee": 5500, "advanceFee": 4950, "Deposit": 2750},
    "Political Science and International Relations (English)": {"listFee": 5500, "advanceFee": 4950, "Deposit": 2750},
    "Political Science and Public Administration (English)": {"listFee": 5500, "advanceFee": 4950, "Deposit": 2750},
    "Political Science and Public Administration (Turkish)": {"listFee": 5500, "advanceFee": 4950, "Deposit": 2750},
    "Architecture (English)": {"listFee": 5000, "advanceFee": 4500, "Deposit": 2500},
    "Architecture (Turkish)": {"listFee": 5000, "advanceFee": 4500, "Deposit": 2500},
    "Industrial Design (Turkish)": {"listFee": 5000, "advanceFee": 4500, "Deposit": 2500},
    "Interior Architecture and Environmental Design (English)": {"listFee": 5000, "advanceFee": 4500, "Deposit": 2500},
    "Interior Architecture and Environmental Design (Turkish)": {"listFee": 5000, "advanceFee": 4500, "Deposit": 2500},
    "Visual Communication Design (Turkish)": {"listFee": 5000, "advanceFee": 4500, "Deposit": 2500},
    "(Turkish) Music Art (Turkish)": {"listFee": 5000, "advanceFee": 4500, "Deposit": 2500},
    "Gastronomy and Culinary Arts (Turkish)": {"listFee": 5000, "advanceFee": 4500, "Deposit": 2500},
    "Urban Design and Landscape Architecture (Turkish)": {"listFee": 5000, "advanceFee": 4500, "Deposit": 2500},
    "Psychological Counselling and Guidance (English)": {"listFee": 5000, "advanceFee": 4500, "Deposit": 2500},
    "Psychological Counselling and Guidance (Turkish)": {"listFee": 5000, "advanceFee": 4500, "Deposit": 2500},
    "(English) Teaching (English)": {"listFee": 5000, "advanceFee": 4500, "Deposit": 2500},
    "Primary Mathematics Teaching (Turkish)": {"listFee": 5000, "advanceFee": 4500, "Deposit": 2500},
    "Special Education Teaching (Turkish)": {"listFee": 5000, "advanceFee": 4500, "Deposit": 2500},
    "Preschool Teaching (Turkish)": {"listFee": 5000, "advanceFee": 4500, "Deposit": 2500},
    "Journalism (Turkish)": {"listFee": 5000, "advanceFee": 4500, "Deposit": 2500},
    "Public Relations and Advertising (English)": {"listFee": 5000, "advanceFee": 4500, "Deposit": 2500},
    "Public Relations and Advertising (Turkish)": {"listFee": 5000, "advanceFee": 4500, "Deposit": 2500},
    "Media and Visual Arts (Turkish)": {"listFee": 5000, "advanceFee": 4500, "Deposit": 2500},
    "New Media and Communication (Turkish)": {"listFee": 5000, "advanceFee": 4500, "Deposit": 2500},
    "Radio Television and Cinema (Turkish)": {"listFee": 5000, "advanceFee": 4500, "Deposit": 2500},
    "Nutrition and Dietetics (Turkish)": {"listFee": 5000, "advanceFee": 4500, "Deposit": 2500},
    "Child Development (Turkish) (Haliç Campus) (Turkish)": {"listFee": 5000, "advanceFee": 4500, "Deposit": 2500},
    "Midwifery (Haliç Campus) (Turkish)": {"listFee": 5000, "advanceFee": 4500, "Deposit": 2500},
    "Physiotherapy and Rehabilitation (Haliç Campus) (Turkish)": {"listFee": 5000, "advanceFee": 4500, "Deposit": 2500},
    "Audiology (Turkish)": {"listFee": 5000, "advanceFee": 4500, "Deposit": 2500},
    "Social Services (Turkish)": {"listFee": 5000, "advanceFee": 4500, "Deposit": 2500},
    "Justice (Turkish)": {"listFee": 3500, "advanceFee": 2800, "Deposit": 1750},
    "Oral and Dental Health (Turkish)": {"listFee": 3500, "advanceFee": 3150, "Deposit": 1750},
    "Operating Room Services (Turkish)": {"listFee": 3500, "advanceFee": 3150, "Deposit": 1750},
    "Anesthesia (English)": {"listFee": 4000, "advanceFee": 3600, "Deposit": 2000},
    "Anesthesia (Turkish)": {"listFee": 3500, "advanceFee": 3150, "Deposit": 1750},
    "Dental Prosthetics Technology (Turkish)": {"listFee": 4000, "advanceFee": 3600, "Deposit": 2000},
    "Child Development (Turkish)": {"listFee": 3500, "advanceFee": 3150, "Deposit": 1750},
    "Dental Prosthesis Technology (Turkish)": {"listFee": 4000, "advanceFee": 3600, "Deposit": 2000},
    "Dialysis (Turkish)": {"listFee": 3500, "advanceFee": 3150, "Deposit": 1750},
    "Pharmacy Services (Turkish)": {"listFee": 3500, "advanceFee": 3150, "Deposit": 1750},
    "Electroneurophysiology (Turkish)": {"listFee": 3500, "advanceFee": 3150, "Deposit": 1750},
    "Physiotherapy (English)": {"listFee": 3500, "advanceFee": 3150, "Deposit": 1750},
    "Physiotherapy (Turkish)": {"listFee": 3500, "advanceFee": 3150, "Deposit": 1750},
    "First and Emergency Aid (English)": {"listFee": 4000, "advanceFee": 3600, "Deposit": 2000},
    "First and Emergency Aid (Turkish)": {"listFee": 3500, "advanceFee": 3150, "Deposit": 1750},
    "Occupational Health and Safety (Turkish)": {"listFee": 3500, "advanceFee": 3150, "Deposit": 1750},
    "Audiometry (Turkish)": {"listFee": 3500, "advanceFee": 3150, "Deposit": 1750},
    "Opticianry (Turkish)": {"listFee": 3500, "advanceFee": 3150, "Deposit": 1750},
    "Pathology Laboratory Techniques (Turkish)": {"listFee": 3500, "advanceFee": 3150, "Deposit": 1750},
    "Radiotherapy (Turkish)": {"listFee": 3500, "advanceFee": 3150, "Deposit": 1750},
    "Prosthetics and Orthotics (Turkish)": {"listFee": 3500, "advanceFee": 3150, "Deposit": 1750},
    "Management of Health Institutions (Turkish)": {"listFee": 3500, "advanceFee": 3150, "Deposit": 1750},
    "Medical Documentation and Secretary (Turkish)": {"listFee": 3500, "advanceFee": 3150, "Deposit": 1750},
    "Medical Imaging Techniques (Turkish)": {"listFee": 3500, "advanceFee": 3150, "Deposit": 1750},
    "Medical Laboratory Techniques (Turkish)": {"listFee": 3500, "advanceFee": 3150, "Deposit": 1750},
    "Oral and Dental Health (Turkish)": {"listFee": 3500, "advanceFee": 3150, "Deposit": 1750},
    "Operation Room Service (Turkish)": {"listFee": 3500, "advanceFee": 3150, "Deposit": 1750},
    "Anesthesia (Turkish)": {"listFee": 3500, "advanceFee": 3150, "Deposit": 1750},
    "Computer Programming (Turkish)": {"listFee": 3500, "advanceFee": 3150, "Deposit": 1750},
    "Child Development (Turkish)": {"listFee": 3500, "advanceFee": 3150, "Deposit": 1750},
    "Biomedical Device Technology (Turkish)": {"listFee": 3500, "advanceFee": 3150, "Deposit": 1750},
    "Dental Prosthesis Technology (Haliç) ": {"listFee": 3500, "advanceFee": 3150, "Deposit": 1750},
    "Dialysis (Turkish)": {"listFee": 3500, "advanceFee": 3150, "Deposit": 1750},
    "Electroneurophysiology (Turkish)": {"listFee": 3500, "advanceFee": 3150, "Deposit": 1750},
    "physiotherapy (Turkish)": {"listFee": 3500, "advanceFee": 3150, "Deposit": 1750},
    "Interior Design (Turkish)": {"listFee": 3500, "advanceFee": 3150, "Deposit": 1750},
    "First and emergency Aid (Turkish)": {"listFee": 3500, "advanceFee": 3150, "Deposit": 1750},
    "Construction Technology (Turkish)": {"listFee": 3500, "advanceFee": 3150, "Deposit": 1750},
    "Occiptional Health and Safety (Turkish)": {"listFee": 3500, "advanceFee": 3150, "Deposit": 1750},
    "Architectural Restoration (Turkish)": {"listFee": 3500, "advanceFee": 3150, "Deposit": 1750},
    "Audiometry (Turkish)": {"listFee": 3500, "advanceFee": 3150, "Deposit": 1750},
    "Opticianry (Turkish)": {"listFee": 3500, "advanceFee": 3150, "Deposit": 1750},
    "Civil Aviation Transportation Management (Turkish)": {"listFee": 3500, "advanceFee": 3150, "Deposit": 1750},
    "Management of Health Institutions (Turkish)": {"listFee": 3500, "advanceFee": 3150, "Deposit": 1750},
    "Medical Documentation and Secretary (Turkish)": {"listFee": 3500, "advanceFee": 3150, "Deposit": 1750},
    "Medical Imaging Techniques (Turkish)": {"listFee": 3500, "advanceFee": 3150, "Deposit": 1750},
    "Medical Laboratory Techniques (Turkish)": {"listFee": 3500, "advanceFee": 3150, "Deposit": 1750},
    "Foreign Trade": {"listFee": 3500, "advanceFee": 3150, "Deposit": 1750},
    "Banking and Insurance (Turkish)": {"listFee": 3500, "advanceFee": 3150, "Deposit": 1750},
    "Foreign Trade (Turkish)": {"listFee": 3500, "advanceFee": 3150, "Deposit": 1750},
    "Public Relations and Publicity (Turkish)": {"listFee": 3500, "advanceFee": 3150, "Deposit": 1750},
    "Human Resources Management (Turkish)": {"listFee": 3500, "advanceFee": 3150, "Deposit": 1750},
    "Business Administration (Turkish)": {"listFee": 3500, "advanceFee": 3150, "Deposit": 1750},
    "Logistics (Turkish)": {"listFee": 3500, "advanceFee": 3150, "Deposit": 1750},
    "Finance (Turkish)": {"listFee": 3500, "advanceFee": 3150, "Deposit": 1750},
    "Accounting and Taxation (Turkish)": {"listFee": 3500, "advanceFee": 3150, "Deposit": 1750},
    "Radio and Television Programming (Turkish)": {"listFee": 3500, "advanceFee": 3150, "Deposit": 1750},
    "Civil Aviation Cabin Services": {"listFee": 3500, "advanceFee": 3150, "Deposit": 1750},
    "Social Services (Turkish)": {"listFee": 3500, "advanceFee": 3150, "Deposit": 1750},
    "Sports Management (Turkish)": {"listFee": 3500, "advanceFee": 3150, "Deposit": 1750},
    "Applied English and Translation (English)": {"listFee": 3500, "advanceFee": 3150, "Deposit": 1750},
    "English Language Teaching (English)": {"listFee": 4500, "advanceFee": 4050, "Deposit": 1750},
    "Pre-School Teaching (Turkish)": {"listFee": 1200, "advanceFee": 1200, "Deposit": 1750},

}

specialCountries = {
    "Afghanistan", "Albania", "Algeria", "Azerbaijan", "Bangladesh", "Belarus",
    "Bosnia and Herzegovina", "Brazil", "Bulgaria", "Cameroon", "Chile", "Comoros",
    "Congo", "Croatia", "Democratic Republic of Congo", "Djibouti", "Egypt",
    "Georgia", "Greece", "India", "Indonesia", "Israel", "Jordan", "Kazakhstan",
    "Kenya", "Kosovo", "Kyrgyzstan", "Lebanon", "Malaysia", "Mali", "Mauritania",
    "Moldova", "Mongolia", "Montenegro", "Morocco", "Myanmar", "Nigeria", "North Macedonia",
    "Pakistan", "Palestine", "Peru", "Romania", "Russia", "Senegal", "Serbia",
    "Sierra Leone", "Somalia", "Sudan", "Tajikistan", "Tanzania", "Tunisia",
    "Ukraine", "Uzbekistan"
}

includedTurkishCountries = {
    "Bahrain", "United Arab Emirates", "Iraq", "Qatar", "Comoros", "Kuwait",
    "Libya", "Saudi Arabia", "Syria", "Oman", "Yemen"
}

excludedTurkishPrograms = ["Medicine", "Dentistry", "Pharmacy", "Physiotherapy", "Law"]
campaigns = {
    "Biomedical Engineering English": 15,
    "Electrical-Electronic Engineering English": 15,
    "Nursing English": 15,
    "Architecture English": 15,
    "Medicine Turkish": 0,
    "Pharmacy English": 0
}

# Layout
app.layout = dbc.Container([
    html.H1("University Program Revenue Calculator", className="text-center mb-4"),

    # File upload for dynamic columns
    dbc.Row([
        dbc.Col(html.Label("Upload Excel File:", className="fw-bold"), width="auto"),
        dbc.Col(
            dcc.Upload(
                id="upload-file",
                children=html.Button("Choose File"),
                multiple=False,
                className="mb-2"
            ),
            width=8
        ),
    ], justify="start", className="mb-3"),

    # Placeholder for displaying dynamic table from uploaded file
    html.Div(id="dynamic-table"),
], fluid=True)


# Callback to process the uploaded file, aggregate data, and display the table
@app.callback(
    Output("dynamic-table", "children"),
    Input("upload-file", "contents")
)
def display_uploaded_file(contents):
    if contents is None:
        return html.Div("Please upload an Excel file.")

    # Decode the uploaded file
    content_type, content_string = contents.split(',')
    decoded = base64.b64decode(content_string)
    try:
        # Read the file into a DataFrame
        df = pd.read_excel(io.BytesIO(decoded))

        # Check if necessary columns exist in the uploaded file
        required_columns = ["Program Name", "Nationality"]
        if not all(col in df.columns for col in required_columns):
            return html.Div("The uploaded file must contain 'Program Name' and 'Nationality' columns.")

        # Count the number of students per program and nationality
        df["Number of Students"] = 1  # Assign 1 to each row for counting
        grouped_df = df.groupby(["Program Name", "Nationality"], as_index=False)["Number of Students"].sum()

        # Add listFee, advanceFee, and language columns with numeric conversion
        grouped_df["Gross Tuition Fee ($)"] = grouped_df["Program Name"].apply(
            lambda x: tuitionFees.get(x, {}).get("listFee", "N/A")
        ).apply(pd.to_numeric, errors='coerce')

        grouped_df["Advance Fee ($)"] = grouped_df["Program Name"].apply(
            lambda x: tuitionFees.get(x, {}).get("advanceFee", "N/A")
        ).apply(pd.to_numeric, errors='coerce')

        grouped_df["Language"] = grouped_df["Program Name"].apply(
            lambda x: "English" if "(English)" in x else "Turkish" if "(Turkish)" in x else "Mixed"
        )

        # Apply discount logic and revenue calculations
        def calculate_final_fee(row):
            gross_tuition = row["Gross Tuition Fee ($)"]
            discount_percentage = 15 if row["Nationality"] in specialCountries else 0
            discount_percentage = min(discount_percentage, 50)  # Cap discount at 50%
            final_fee = gross_tuition - (gross_tuition * discount_percentage / 100) if pd.notnull(gross_tuition) else 0
            return final_fee

        grouped_df["Final Fee ($)"] = grouped_df.apply(calculate_final_fee, axis=1)
        grouped_df["Total Revenue ($)"] = grouped_df["Final Fee ($)"] * grouped_df["Number of Students"]
        grouped_df["Advance Fee Revenue ($)"] = grouped_df["Advance Fee ($)"] * grouped_df["Number of Students"]
        # Normalize program names for a case-insensitive match with tuitionFees
        grouped_df["Gross Tuition Fee ($)"] = grouped_df["Program Name"].apply(
            lambda x: tuitionFees.get(x.strip(), {}).get("listFee", "N/A")
        )
        grouped_df["Advance Fee ($)"] = grouped_df["Program Name"].apply(
            lambda x: tuitionFees.get(x.strip(), {}).get("advanceFee", "N/A")
        )
        grouped_df["Deposit ($)"] = grouped_df["Program Name"].apply(
            lambda x: tuitionFees.get(x.strip(), {}).get("Deposit", "N/A")
        )

        # Total calculations
        total_revenue_sum = grouped_df["Total Revenue ($)"].sum()
        total_students_count = grouped_df["Number of Students"].sum()

        # Add these totals as a summary row
        summary_row = pd.DataFrame({
            "Program Name": ["Total"],
            "Nationality": [""],
            "Language": [""],
            "Gross Tuition Fee ($)": [""],
            "Number of Students": [total_students_count],
            "Advance Fee ($)": [""],
            "Final Fee ($)": [""],
            "Total Revenue ($)": [total_revenue_sum],
            "Advance Fee Revenue ($)": [""]
        })

        grouped_df = pd.concat([grouped_df, summary_row], ignore_index=True)

        # Arrange the columns in the desired order
        displayed_columns = ["Program Name", "Nationality", "Language", "Number of Students", "Gross Tuition Fee ($)", "Advance Fee ($)",
                              "Deposit ($)", "Total Revenue ($)", "Total Revenue ($)", "Total Revenue ($)", "Total Revenue ($)", "Advance Fee Revenue ($)"]
        grouped_df = grouped_df[displayed_columns]

        # Create a Dash DataTable with the processed data
        table = dash_table.DataTable(
            data=grouped_df.to_dict('records'),
            columns=[{"name": col, "id": col} for col in grouped_df.columns],
            style_table={'overflowX': 'auto'},
            style_cell={'textAlign': 'center'},
            style_header={
                'backgroundColor': 'blue',
                'color': 'white',
                'fontWeight': 'bold',
                'textAlign': 'center'
            },
            style_data={
                'color': 'black',
                'backgroundColor': 'white'
            }
        )

        return table

    except Exception as e:
        return html.Div(f"Error processing file: {e}")


if __name__ == "__main__":
    app.run_server(debug=True)