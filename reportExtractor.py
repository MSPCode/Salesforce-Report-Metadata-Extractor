import json
from collections import OrderedDict
from openpyxl import Workbook
from simple_salesforce import Salesforce
import creds  # Assuming Salesforce credentials are stored here now

REPORT_FIELDS = [
    'Id', 'Name', 'DeveloperName', 'CreatedDate', 'CreatedBy.Name', 'LastRunDate',
    'LastModifiedBy.Name', 'LastModifiedDate', 'LastViewedDate', 'Description'
]

QUERY_WHERE_CLAUSE = ""  # SPECIFY WHERE CLAUSE HERE

class CustomSalesforce(Salesforce):
    def __init__(self):
        super().__init__(
            username=creds.SALESFORCE_USERNAME, 
            password=creds.SALESFORCE_PASSWORD,
            security_token = '',
            domain='test' if creds.SALESFORCE_SANDBOX else None, 
            version=creds.SALESFORCE_API_VERSION
        )

    def describe_report(self, report_id):
        url = f"{self.base_url}analytics/reports/{report_id}/describe"
        result = self._call_salesforce('GET', url)
        return result.json(object_pairs_hook=OrderedDict) if result.status_code == 200 else None

def get_column_labels(report_metadata, field_apis):
    detail_column_info = report_metadata.get('detailColumnInfo', {})
    return [(detail_column_info[api]['label'], api.split('.')[-1]) for api in field_apis]

def get_folder_name(report_json, folder_names_by_id):
    folder_id = report_json['reportMetadata']['folderId']
    return folder_names_by_id.get(
        folder_id, 
        'My Personal Custom Reports' if folder_id.startswith('005') else 
        'Unfiled Custom Reports' if folder_id.startswith('00D') else folder_id
    )

def get_filters(filter_list):
    """
    Extracts and formats filter information from a given list of filters.
    """
    filter_strings = []
    for filter_info in filter_list:
        column = filter_info.get('column', '').split('.')[-1]  # Assuming columns might be prefixed and we need the last part
        operator = filter_info.get('operator', '')
        value = filter_info.get('value', '')
        filter_strings.append(f"{column} {operator} {value}")
    return ' | '.join(filter_strings)

def get_filter_names(report_metadata):
    """
    Extracts filter names from report metadata, accommodating structures with or without 'blocks'.
    """
    if 'blocks' in report_metadata:
        filters = []
        for block in report_metadata['blocks']:
            if 'reportFilters' in block:
                filters.append(get_filters(block['reportFilters']))
        return ', '.join(filters)
    else:
        return get_filters(report_metadata.get('reportFilters', []))
    
def safe_slice(date_string, slice_pos):
    """Returns a sliced string if not None, else an empty string."""
    return date_string[:slice_pos] if date_string else ''


def fetch_and_process_reports(sfdc):
    data = []
    query = f"SELECT {','.join(REPORT_FIELDS)} FROM Report {QUERY_WHERE_CLAUSE}"
    result = sfdc.query_all(query)
    folder_names_by_id = {f['Id']: f['Name'] for f in sfdc.query_all("SELECT Id, Name FROM Folder")['records']}
    
    for record in result['records']:
        report_metadata = sfdc.describe_report(record['Id'])
        if not report_metadata:
            continue  # Skip if no metadata found

        # Assemble the base URL for report links
        report_url = f"{sfdc.sf_instance}/{record['Id']}"
        folder_name = get_folder_name(report_metadata, folder_names_by_id)
        field_data = get_column_labels(report_metadata['reportExtendedMetadata'], report_metadata['reportMetadata']['detailColumns'])

        # Extracting additional details from the report metadata
        report_name = report_metadata['reportMetadata'].get('name', '')
        report_type = report_metadata['reportMetadata'].get('reportType', {}).get('label', '')
        report_format = report_metadata['reportMetadata'].get('reportFormat', '')
        created_by_name = record.get('CreatedBy', {}).get('Name', '')
        last_modified_by_name = record.get('LastModifiedBy', {}).get('Name', '')
        created_date = safe_slice(record.get('CreatedDate'), -9)
        last_modified_date = safe_slice(record.get('LastModifiedDate'), -9)
        last_run_date = safe_slice(record.get('LastRunDate'), -9)
        last_view_date = safe_slice(record.get('LastViewedDate'), -9)

        #print statement
        print("Processing: ", report_name)
        # Extracting additional details including filters
        filters = get_filter_names(report_metadata['reportMetadata'])

        # Construct a structured data row for each report
        data.append({
            'Report_url': report_url,
            'Report_ID': record['Id'],
            'Report_Name': report_name,
            'Folder_Name': folder_name,
            'Report_Type': report_type,
            'Report_Format': report_format,
            'Field_Labels': [fd[0] for fd in field_data],  # Extracting label part of field_data
            'Field_API_Names': [fd[1] for fd in field_data],  # Extracting API name part of field_data
            'Filters': filters,
            'Created_By': created_by_name,
            'Created_Date': created_date,
            'Last_Modified_By': last_modified_by_name,
            'Last_Modified_Date': last_modified_date,
            'Last_Run_Date': last_run_date,
            'Last_View_Date': last_view_date
        })


    return data

def save_to_excel(data):
    wb = Workbook()
    ws = wb.active
    headers = ['Report_url', 'Report_ID', 'Report_Name', 'Folder_Name', 'Report_Type', 'Report_Format', 'Field_Labels', 'Field_API_Names', 'Filters', 'Created_By', 'Created_Date', 'Last_Modified_By', 'Last_Modified_Date', 'Last_Run_Date', 'Last_View_Date']
    ws.append(headers)

    for row in data:
        # Convert lists to strings
        field_labels_str = ', '.join(row.get('Field_Labels', []))
        field_api_names_str = ', '.join(row.get('Field_API_Names', []))

        ws.append([
            row.get('Report_url', ''),
            row.get('Report_ID', ''),
            row.get('Report_Name', ''),
            row.get('Folder_Name', ''),
            row.get('Report_Type', ''),
            row.get('Report_Format', ''),
            field_labels_str,  # Use the joined string version
            field_api_names_str,  # Use the joined string version
            row.get('Filters', ''),
            row.get('Created_By', ''),
            row.get('Created_Date', ''),
            row.get('Last_Modified_By', ''),
            row.get('Last_Modified_Date', ''),
            row.get('Last_Run_Date', ''),
            row.get('Last_View_Date', '')
        ])
    wb.save('Report_Of_Reports.xlsx')

def main():
    sfdc = CustomSalesforce()
    report_data = fetch_and_process_reports(sfdc)
    save_to_excel(report_data)

if __name__ == "__main__":
    main()
