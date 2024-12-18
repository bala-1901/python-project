from flask import Flask, render_template, request, redirect, url_for
import openpyxl

app = Flask(__name__)

# Path to your Excel file
EXCEL_FILE = 'Excel_Sample.xlsx'

def load_or_create_excel():
    """Load the Excel file or create one if it doesn't exist."""
    try:
        # Try to open an existing Excel file
        workbook = openpyxl.load_workbook(EXCEL_FILE)
    except FileNotFoundError:
        # Create a new Excel workbook if it doesn't exist
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        # sheet.append(["Name", "Age", "Email"])  # Add headers

        sheet.append(["Company Identifier", "Currency", "Start Date","Date Range","OnePage Profile PitchBook","OnePage Profile FactSet"])  # Add headers
       		# Start Date	Date Range			Top 20 Investors overview	Relationship Snapshot	PrivateCompany Investor

        workbook.save(EXCEL_FILE)
    return workbook

@app.route('/')
def index():
    """Render the form for data input."""
    return render_template('index.html')

@app.route('/submit', methods=['POST'])
def submit():
    """Handle form submission and write data to Excel."""
    # Retrieve form data
    # name = request.form['name']
    # age = request.form['age']
    # email = request.form['email']

    companyidentifier = request.form['Company_Identifier']
    currency = request.form['Currency']
    startdate = request.form['Start_Date']
    daterange = request.form['Date_Range']
    onepage_profile_pitchbook = 'OnePage_Profile_PitchBook' in request.form 
    onepage_profile_factset = 'OnePage_Profile_FactSet' in request.form 

    print(onepage_profile_pitchbook)
    print(onepage_profile_factset)

    #		Top 20 Investors overview	Relationship Snapshot	PrivateCompany Investor

    
    # Load or create the Excel workbook
    workbook = load_or_create_excel()
    sheet = workbook.active
    
    # Append new row with data to the Excel file
    sheet.append([companyidentifier, currency, startdate, daterange, onepage_profile_pitchbook, onepage_profile_factset])

    # sheet.append([name, age, email])
    
    # Save the workbook with the new data
    workbook.save(EXCEL_FILE)
    
    return redirect(url_for('index'))

if __name__ == '__main__':
    app.run(debug=True)
