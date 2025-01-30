import requests
import pandas as pd
import json
import urllib3
from datetime import datetime

# Disable SSL warnings
urllib3.disable_warnings()

class EndpointsReport:
    def __init__(self, ndi_ip):
        self.base_url = f"https://{ndi_ip}"
        self.session = requests.Session()
        self.session.verify = False
        self.session.headers.update({'Connection': 'close'})

    def login(self, domain, username, password):
        """Login to NDI"""
        login_url = f"{self.base_url}/login"
        credentials = {
            "domain": domain,
            "userName": username,
            "userPasswd": password
        }
        response = self.session.post(login_url, json=credentials, timeout=(5, 30))
        response.raise_for_status()

    def get_all_endpoints(self, site_name="INPUT_FABRIC_NAME_HERE"):
        """Get all endpoints data from NDI"""
        # First get total count
        url = f"{self.base_url}/sedgeapi/v1/cisco-nir/api/api/v1/endpoints?siteName={site_name}&count=1"
        
        print("Getting total count...")
        try:
            response = self.session.get(url, timeout=(5, 30))
            response.raise_for_status()
            data = response.json()
            
            total_count = data.get('totalItemsCount', 0)
            print(f"Total endpoints available: {total_count}")
            
            if total_count == 0:
                print("No endpoints found for the given site name")
                return {"entries": [], "totalItemsCount": 0}
            
            # Now fetch all records with the total count
            url = f"{self.base_url}/sedgeapi/v1/cisco-nir/api/api/v1/endpoints?siteName={site_name}&count={total_count}"
            print(f"Fetching all endpoints with URL: {url}")
            
            response = self.session.get(url, timeout=(5, 30))
            response.raise_for_status()
            data = response.json()
            
            all_entries = data.get('entries', [])
            print(f"Fetched {len(all_entries)} endpoints")
            return {"entries": all_entries, "totalItemsCount": total_count}
            
        except requests.exceptions.RequestException as e:
            print(f"Error details:")
            print(f"Status code: {e.response.status_code if hasattr(e, 'response') else 'N/A'}")
            print(f"Response content: {e.response.text if hasattr(e, 'response') else 'N/A'}")
            raise

    def format_value(self, value):
        """Format a value for Excel output"""
        if isinstance(value, (list, tuple)):
            # Convert each item in the list
            formatted_items = []
            for item in value:
                if isinstance(item, dict):
                    # For dictionaries, try to get meaningful values
                    if 'name' in item:
                        formatted_items.append(str(item['name']))
                    elif 'value' in item:
                        formatted_items.append(str(item['value']))
                    else:
                        # If no specific field to extract, use the whole dict
                        formatted_items.append(json.dumps(item))
                else:
                    formatted_items.append(str(item))
            return ', '.join(formatted_items)
        elif isinstance(value, dict):
            # For dictionaries, try to get meaningful values
            if 'name' in value:
                return str(value['name'])
            elif 'value' in value:
                return str(value['value'])
            else:
                return json.dumps(value)
        else:
            return str(value)

    def process_endpoints_data(self, data):
        """Process endpoints data and create DataFrame"""
        entries = data.get('entries', [])
        print(f"Processing {len(entries)} endpoints...")
        
        # First, create the DataFrame
        df = pd.DataFrame(entries)
        
        # Process each column
        for column in df.columns:
            df[column] = df[column].apply(self.format_value)
        
        return df

    def generate_report(self, site_name="ACI-ODC"):
        """Generate endpoints report"""
        try:
            # Get all endpoints data
            print("Fetching all endpoints data...")
            endpoints_data = self.get_all_endpoints(site_name)
            
            if not endpoints_data.get('entries'):
                print("No data to process")
                return None
            
            # Process data
            print("Processing data...")
            df = self.process_endpoints_data(endpoints_data)
            
            # Generate timestamp for filename
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            output_filename = f'endpoints_report_{timestamp}.xlsx'
            
            print("Generating Excel report...")
            # Create Excel writer with xlsxwriter engine
            with pd.ExcelWriter(output_filename, engine='xlsxwriter') as writer:
                df.to_excel(writer, sheet_name='Endpoints', index=False)
                
                # Get workbook and worksheet objects
                workbook = writer.book
                worksheet = writer.sheets['Endpoints']
                
                # Define formats
                header_format = workbook.add_format({
                    'bold': True,
                    'text_wrap': True,
                    'valign': 'vcenter',
                    'align': 'center',
                    'bg_color': '#366092',
                    'font_color': 'white',
                    'border': 1,
                    'border_color': '#D9D9D9',
                    'font_size': 11
                })
                
                data_format = workbook.add_format({
                    'text_wrap': True,
                    'valign': 'vcenter',
                    'align': 'left',
                    'border': 1,
                    'border_color': '#D9D9D9',
                    'font_size': 10
                })
                
                # Set row height for header
                worksheet.set_row(0, 30)
                
                # Write headers with formatting
                for col_num, value in enumerate(df.columns.values):
                    # Make header text more readable
                    header_text = value.replace('_', ' ').title()
                    worksheet.write(0, col_num, header_text, header_format)
                
                # Write data with formatting
                for row in range(len(df)):
                    for col in range(len(df.columns)):
                        worksheet.write(row + 1, col, df.iloc[row, col], data_format)
                
                # Adjust column widths based on content
                for idx, col in enumerate(df.columns):
                    # Get maximum length in the column
                    max_length = max(
                        df[col].astype(str).apply(len).max(),
                        len(col.replace('_', ' ').title())  # Account for formatted header
                    )
                    # Set width with some padding, but cap at 50
                    width = min(max_length + 3, 50)
                    worksheet.set_column(idx, idx, width)
                
                # Freeze the header row
                worksheet.freeze_panes(1, 0)
                
                # Add alternating row colors
                for row in range(1, len(df) + 1):
                    if row % 2 == 0:
                        worksheet.set_row(row, None, None, {'level': 1, 'hidden': False})
                    else:
                        worksheet.set_row(row, None, workbook.add_format({'bg_color': '#F2F2F2'}))
            
            print(f"Report has been generated: {output_filename}")
            print(f"Total endpoints processed: {len(df)}")
            return output_filename
            
        except Exception as e:
            print(f"Error occurred while generating report: {e}")
            if hasattr(e, 'response'):
                print(f"Response content: {e.response.text}")
            return None

def main():
    # NDI details
    ndi_ip = "NDI_IP"
    domain = "METHOD"
    username = "NDI_USERNAME"
    password = "NDI PASSWORD"
    
    # Create report instance
    report = EndpointsReport(ndi_ip)
    
    try:
        # Login to NDI
        print("Logging in to NDI...")
        report.login(domain, username, password)
        
        # Generate report
        print("Generating report...")
        site_name = input("Enter site name: ") or "INPUT_FABRIC_NAME_HERE"
        output_file = report.generate_report(site_name)
        
        if output_file:
            print("Script completed successfully!")
            
    except Exception as e:
        print(f"Error: {str(e)}")
        if hasattr(e, 'response'):
            print(f"Response content: {e.response.text}")

if __name__ == "__main__":
    main()
