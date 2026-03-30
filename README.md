# HKIE CPD Certificate Generator

A Streamlit-based web application for automating the generation of personalized CPD (Continuing Professional Development) certificates for HKIE (Hong Kong Institution of Engineers) events.

## Features

- **Event Information Scraping**: Automatically fetch event details from HKIE website
- **Flexible Data Import**: Support for CSV and Excel file formats
- **Zoom Attendance Verification**: Cross-reference registration data with Zoom attendee reports
- **Batch Certificate Generation**: Generate certificates for multiple attendees at once
- **Multiple Output Formats**: 
  - Word documents (.docx) - Unencrypted
  - PDF files (.pdf) - Encrypted with attendee's email as password
- **ZIP Packaging**: Download all generated certificates in a single ZIP file

## Requirements

### Python Packages
streamlit pandas requests beautifulsoup4 docxtpl pikepdf pywin32 (Windows only)

### System Requirements
- **PDF Conversion**: Windows OS with Microsoft Word installed (for PDF export feature)
- **Word Generation**: Cross-platform support (Windows, macOS, Linux)

## Installation

1. Clone the repository:
```bash
git clone <repository-url>
cd CPD-Cert
```

2. Install dependencies:
```bash 
pip install streamlit pandas requests beautifulsoup4 docxtpl pikepdf
```

3. (Windows only) Install pywin32 for PDF conversion:
```bash 
pip install pywin32
```

## Usage

1. Run the application:
```bash 
streamlit run app.py
```
2. Follow the 4-step process in the web interface:

### Step 1: Get Event Information

- Enter the HKIE event URL
- Click "Fetch Event Info" to automatically extract event title and details
- Or manually enter the information

### Step 2: Upload Files

- __Registration Form__ (Required): Upload the registration Excel/CSV file
- __Certificate Template__ (Required): Upload the Word template (.docx)
- __Zoom Report__ (Optional): Upload Zoom attendee report for verification

### Step 3: Process Data

- The system automatically maps columns from the registration file
- If Zoom verification is enabled, it matches attendees by email
- Displays matched and unmatched records for review

### Step 4: Generate Certificates

- Select output format (Word or encrypted PDF)
- Click "Start Generation"
- Download the ZIP file containing all certificates

## Data Format Requirements

### Registration File Columns

The registration file should contain the following columns (auto-detected):

- `First Name` or `名字`
- `Last Name` or `姓氏`
- `Email` or `電郵`
- `Membership No` or `會員編號` (Optional)
- `Salutation` or `稱呼` (Optional)

### Word Template Variables

Use these placeholders in your Word template:

- `{{ name }}` - Full name with salutation
- `{{ membership_no }}` - Membership number
- `{{ event_title }}` - Event title
- `{{ event_details }}` - Event date and time

## Functions

### `normalize_name(name)`

Normalizes names for comparison by:

- Converting to lowercase
- Removing common titles (ir, mr, ms, miss, dr, prof)
- Removing special characters
- Collapsing multiple spaces

### `parse_zoom_report(file_obj)`

Parses Zoom attendee reports (CSV or Excel):

- Detects "Attendee Details" section automatically
- Handles trailing commas in CSV files
- Aggregates total session time for attendees with multiple joins
- Returns (DataFrame, error_message) tuple

## Notes

- __Email Matching__: Zoom verification uses email addresses as the primary matching key
- __PDF Encryption__: When generating PDFs, the attendee's email address is used as the password
- __Membership Number__: If not found in the registration file, it will be left blank in the certificate
- __Name Formatting__: The system automatically splits "Full Name" into "First Name" and "Last Name" if needed

## Troubleshooting

### "Cannot find Attendee Details section in Zoom report"

- Ensure the Zoom report is in the standard format exported from Zoom
- Check that the file contains "Attendee Details" or "User Name" column headers

### PDF conversion not working

- PDF export requires Windows OS with Microsoft Word installed
- On other platforms, use the Word document format instead

### Missing membership numbers

- Verify that the registration file contains a column with "Membership" or "會員編號" in the header
- The system will show a warning if this column is not found

##
