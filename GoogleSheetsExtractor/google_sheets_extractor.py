import json
import os
from dotenv import load_dotenv
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

# Scopes required for reading Google Sheets
SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']
load_dotenv()
# Replace with your Google Sheets ID
SPREADSHEET_ID = os.getenv("SPREADSHEET_ID")

# Sheet and column configuration
SHEET_NAME = 'India vs Eng Women'
VIDEO_LINK_COLUMN = 'VIDEO LINK'

class GoogleSheetsExtractor:
    def __init__(self, spreadsheet_id, sheet_name, column_name):
        self.spreadsheet_id = spreadsheet_id
        self.sheet_name = sheet_name
        self.column_name = column_name
        self.service = None
        self.authenticate()
    
    def authenticate(self):
        """Authenticate with Google Sheets API using service account"""
        try:
            # Load service account credentials
            creds = Credentials.from_service_account_file(
                'service-account-key.json', 
                scopes=SCOPES
            )
            
            self.service = build('sheets', 'v4', credentials=creds)
            print("✅ Successfully authenticated with Google Sheets API")
            
        except FileNotFoundError:
            print("❌ Error: 'service-account-key.json' not found!")
            print("Please download your service account key and rename it to 'service-account-key.json'")
            raise
        except Exception as e:
            print(f"❌ Authentication error: {e}")
            raise
    
    def get_sheet_data(self):
        """Get all data from the specified sheet"""
        try:
            sheet = self.service.spreadsheets()
            result = sheet.values().get(
                spreadsheetId=self.spreadsheet_id,
                range=f'{self.sheet_name}!A:Z'  # Get all columns
            ).execute()
            
            values = result.get('values', [])
            
            if not values:
                print('No data found in the sheet.')
                return None
            
            return values
        
        except Exception as e:
            print(f'Error getting sheet data: {e}')
            return None
    
    def find_column_index(self, headers, column_name):
        """Find the index of the specified column"""
        try:
            return headers.index(column_name)
        except ValueError:
            # If exact match not found, try partial match
            for i, header in enumerate(headers):
                if column_name.lower() in header.lower():
                    return i
            return -1
    
    def extract_sharepoint_links(self):
        """Extract SharePoint links from the specified column"""
        data = self.get_sheet_data()
        
        if not data:
            return []
        
        # Get headers (first row)
        headers = data[0]
        print(f"Available columns: {headers}")
        
        # Find the column index
        column_index = self.find_column_index(headers, self.column_name)
        
        if column_index == -1:
            print(f"Column '{self.column_name}' not found!")
            print(f"Available columns: {headers}")
            return []
        
        print(f"Found '{self.column_name}' at column index {column_index}")
        
        # Extract SharePoint links
        sharepoint_links = []
        
        for row_index, row in enumerate(data[1:], start=2):  # Skip header, start from row 2
            if column_index < len(row):
                cell_value = row[column_index]
                
                if cell_value and isinstance(cell_value, str):
                    # Check if it's a SharePoint link
                    if 'sharepoint.com' in cell_value.lower():
                        sharepoint_links.append({
                            'row': row_index,
                            'link': cell_value.strip()
                        })
        
        return sharepoint_links
    
    def save_links_to_file(self, links, filename='sharepoint_links.json'):
        """Save extracted links to a JSON file"""
        try:
            with open(filename, 'w') as f:
                json.dump(links, f, indent=2)
            print(f"Links saved to {filename}")
        except Exception as e:
            print(f"Error saving links: {e}")

def main():
    # Replace these with your actual values
    spreadsheet_id = SPREADSHEET_ID
    sheet_name = SHEET_NAME
    column_name = VIDEO_LINK_COLUMN
    
    print("Starting Google Sheets SharePoint Link Extractor...")
    print(f"Spreadsheet ID: {spreadsheet_id}")
    print(f"Sheet Name: {sheet_name}")
    print(f"Column Name: {column_name}")
    
    # Create extractor instance
    extractor = GoogleSheetsExtractor(spreadsheet_id, sheet_name, column_name)
    
    # Extract links
    links = extractor.extract_sharepoint_links()
    
    if links:
        print(f"\nFound {len(links)} SharePoint links:")
        for i, link_info in enumerate(links, 1):
            print(f"{i}. Row {link_info['row']}: {link_info['link'][:80]}...")
        
        # Save to file
        extractor.save_links_to_file(links)
        
        # Also save just the URLs to a text file for easy copy-paste
        with open('sharepoint_urls.txt', 'w') as f:
            for link_info in links:
                f.write(link_info['link'] + '\n')
        
        print(f"\nLinks also saved to 'sharepoint_urls.txt' for easy access")
        
    else:
        print("No SharePoint links found!")

if __name__ == '__main__':
    main()