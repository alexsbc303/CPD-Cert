# HKIE CPD Certificate Generator

This is a Streamlit application designed to automate the process of generating CPD (Continuing Professional Development) certificates for HKIE (The Hong Kong Institution of Engineers) events.

It streamlines the workflow by fetching event details, verifying attendance against Zoom reports, and batch-generating encrypted PDF or Word certificates.

## Features

-   **Event Information Retrieval**: Automatically scrapes event title, date, and time from the HKIE ITD website URL.
-   **Attendance Verification**: Cross-checks the Registration list (Excel/CSV) against the Zoom Attendee Report to confirm actual attendance.
    -   Supports matching by **Email** or **Name**.
    -   Normalizes names (removes titles like Ir, Mr, Dr) for better matching accuracy.
-   **Batch Certificate Generation**:
    -   Uses a Microsoft Word (`.docx`) template.
    -   Supports placeholders for Name, Membership Number, Event Title, and Date/Time.
-   **Secure Output**:
    -   Generates certificates in **Word (.docx)** or **PDF (.pdf)** format.
    -   **PDF Encryption**: Automatically encrypts PDF files using the attendee's **Email Address** as the password.
-   **User-Friendly Interface**: Web-based UI built with Streamlit.

## Prerequisites

-   **Operating System**: Windows (Required for PDF conversion via Microsoft Word COM automation).
-   **Software**: Microsoft Word must be installed on the machine running the script.
-   **Python**: Python 3.x

## Installation

1.  Clone the repository:
    ```bash
    git clone https://github.com/alexsbc303/HKIE-CPD-Cert.git
    cd HKIE-CPD-Cert
    ```

2.  Install the required Python packages:
    ```bash
    pip install -r requirements.txt
    ```
    *Common dependencies include: `streamlit`, `pandas`, `requests`, `beautifulsoup4`, `docxtpl`, `pikepdf`, `pywin32`.*

## Usage

1.  **Start the Application**:
    ```bash
    streamlit run app.py
    ```

2.  **Step 1: Get Event Info**:
    -   Enter the HKIE event URL (e.g., `http://it.hkie.org.hk/...`).
    -   Click **"抓取活動資訊" (Fetch Info)** to auto-fill the Event Title and Date/Time. You can also edit these manually.

3.  **Step 2: Upload Files**:
    -   **Registration File**: Upload the Excel/CSV file containing applicant details.
        -   *Required Columns*: First Name, Last Name, Email.
        -   *Optional Columns*: Membership No, Salutation.
    -   **Certificate Template**: Upload the `.docx` template file.
    -   **Zoom Report** (Optional): Upload the Zoom attendee report to verify attendance.

4.  **Step 3: Process List**:
    -   The app will map columns and cross-check attendees.
    -   Review the preview of the matched list.

5.  **Step 4: Generate & Download**:
    -   Select output format: **Word** or **PDF (Encrypted)**.
    -   Click **"開始生成" (Start Generation)**.
    -   Download the resulting `.zip` file containing all certificates.

## Template Configuration

The Word (`.docx`) template should use Jinja2-style placeholders:

-   `{{ name }}` : Attendee's full name (with Salutation).
-   `{{ membership_no }}` : HKIE Membership Number.
-   `{{ event_title }}` : Title of the event.
-   `{{ event_details }}` : Date and time of the event.

## Notes

-   **PDF Generation**: The PDF conversion relies on `win32com` to control a local instance of Microsoft Word. This feature only works on Windows environments with Word installed.
-   **Encryption**: If PDF output is selected, files are encrypted using `pikepdf`. The password is set to the attendee's **Email Address**.
