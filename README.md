# Salesforce-Report-Metadata-Extractor

This script automates the extraction of metadata from Salesforce reports and exports the information into an Excel file. It's designed to help Salesforce admins and analysts quickly gather insights into report configurations, usage, and structure without manually navigating the Salesforce UI.

## Features

- Queries Salesforce to fetch report metadata based on specified criteria.
- Processes and organizes metadata into structured data.
- Exports the processed data into an Excel file for easy analysis and documentation.

## Prerequisites

- Python 3.x
- `openpyxl` library for Excel file creation (`pip install openpyxl`)
- `simple_salesforce` library for Salesforce API interaction (`pip install simple-salesforce`)
- Salesforce credentials with access to the desired reports.

This script assumes that you have the necessary permissions to access and query report metadata in Salesforce. Please ensure compliance with your organization's Salesforce usage policies.

## Setup

1. **Salesforce Credentials**: Ensure your Salesforce credentials are correctly set up in a `creds.py` file. This file should define variables for `SALESFORCE_USERNAME`, `SALESFORCE_PASSWORD`, `SALESFORCE_SECURITY_TOKEN`, `SALESFORCE_SANDBOX` (boolean), and `SALESFORCE_API_VERSION`.

2. **Install Dependencies**: Run the following command to install necessary Python libraries:

   ```sh
   pip install simple-salesforce openpyxl
   ```
3. **OPTIONAL STEP:** Specify Where Clause: Modify the `QUERY_WHERE_CLAUSE` variable in the script to filter the reports you want to process. Leave it as an empty string to fetch all available reports.

## Usage

```sh
python reportExtractor.py
```
The script will:

Authenticate with Salesforce using the credentials provided in creds.py.
Query Salesforce for reports based on the specified criteria.
Extract metadata for each report and process it.
Export the processed metadata into an Excel file named Report_Of_Reports.xlsx

## Customization

You can customize the script by modifying the `REPORT_FIELDS` list to include or exclude specific fields from the report metadata. Additionally, the `QUERY_WHERE_CLAUSE` can be adjusted to filter reports more granularly based on your needs.

## Output
 
 - Report URL
 - Report ID
 - Report Name
 - Folder Name
 - Report Type
 - Report Format
 - Field Labels
 - Field API Names
 - Filters
 - Created By
 - Created Date
 - Last Modified By
 - Last Modified Date
 - Last Run Date
 - Last View Date

Each row in the Excel file represents a single Salesforce report, providing a comprehensive overview of its configuration and metadata.

## License

MIT

