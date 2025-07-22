import json
import os
from dotenv import load_dotenv
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

# Scopes required for reading Google Sheets
SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']
load_dotenv()

# Replace with your Google Sheets ID for the new sheet
SPREADSHEET_ID = os.getenv("SPREADSHEET_ID")  # Add this to your .env file

# Sheet and column configuration
SHEET_NAME = 'Sheet1'  # Update this to your actual sheet name
LINK_COLUMN = 'Link'

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
    
    def classify_link_type(self, url):
        """Classify the type of link (transfernow or wetransfer)"""
        url_lower = url.lower()
        
        if 'transfernow.net' in url_lower:
            return 'transfernow'
        elif 'wetransfer.com' in url_lower:
            return 'wetransfer'
        else:
            return 'unknown'
    
    def split_cell_links(self, cell_value):
        """Split cell value by various delimiters to extract individual links"""
        if not cell_value:
            return []
        
        # Split by common delimiters
        delimiters = ['\n', '\r\n', '\r', '|', ';', ',']
        links = [cell_value]
        
        for delimiter in delimiters:
            new_links = []
            for link in links:
                new_links.extend([l.strip() for l in link.split(delimiter) if l.strip()])
            links = new_links
        
        # Filter out empty strings and ensure we have valid URLs
        valid_links = []
        for link in links:
            link = link.strip()
            if link and ('http://' in link or 'https://' in link):
                valid_links.append(link)
        
        return valid_links

    def extract_transfer_links(self):
        """Extract transfer links from the specified column"""
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
        
        # Extract transfer links
        transfer_links = []
        link_counter = 1
        
        for row_index, row in enumerate(data[1:], start=2):  # Skip header, start from row 2
            if column_index < len(row):
                cell_value = row[column_index]
                
                if cell_value and isinstance(cell_value, str):
                    # Split the cell value to handle multiple links
                    individual_links = self.split_cell_links(cell_value)
                    
                    for link_url in individual_links:
                        link_type = self.classify_link_type(link_url)
                        
                        # Only process transfernow and wetransfer links
                        if link_type in ['transfernow', 'wetransfer']:
                            transfer_links.append({
                                'id': f"link_{link_counter}",
                                'row': row_index,
                                'original_cell': cell_value[:100] + "..." if len(cell_value) > 100 else cell_value,
                                'url': link_url,
                                'type': link_type,
                                'status': 'pending',
                                'processed': 0  # New field: 0 = not processed, 1 = processed
                            })
                            link_counter += 1
        
        return transfer_links
    
    def save_links_to_file(self, links, filename='transfer_links.json'):
        """Save extracted links to a JSON file"""
        try:
            # Create a structured JSON with metadata
            output_data = {
                'metadata': {
                    'total_links': len(links),
                    'transfernow_count': len([l for l in links if l['type'] == 'transfernow']),
                    'wetransfer_count': len([l for l in links if l['type'] == 'wetransfer']),
                    'last_updated': __import__('datetime').datetime.now().isoformat()
                },
                'links': links
            }
            
            with open(filename, 'w') as f:
                json.dump(output_data, f, indent=2)
            print(f"✅ Links saved to {filename}")
            
            # Print summary
            print(f"\nSummary:")
            print(f"- Total links: {output_data['metadata']['total_links']}")
            print(f"- TransferNow links: {output_data['metadata']['transfernow_count']}")
            print(f"- WeTransfer links: {output_data['metadata']['wetransfer_count']}")
            
        except Exception as e:
            print(f"❌ Error saving links: {e}")

def main():
    # Configuration - update these values
    spreadsheet_id = SPREADSHEET_ID
    sheet_name = SHEET_NAME
    column_name = LINK_COLUMN
    
    if not spreadsheet_id:
        print("❌ Error: Please set NEW_SPREADSHEET_ID in your .env file")
        return
    
    print("🚀 Starting Google Sheets Transfer Link Extractor...")
    print(f"📊 Spreadsheet ID: {spreadsheet_id}")
    print(f"📋 Sheet Name: {sheet_name}")
    print(f"📝 Column Name: {column_name}")
    print("-" * 60)
    
    try:
        # Create extractor instance
        extractor = GoogleSheetsExtractor(spreadsheet_id, sheet_name, column_name)
        
        # Extract links
        links = extractor.extract_transfer_links()
        
        if links:
            print(f"\n✅ Found {len(links)} transfer links:")
            current_row = None
            for i, link_info in enumerate(links, 1):
                link_type = link_info['type'].upper()
                url_preview = link_info['url'][:60] + "..." if len(link_info['url']) > 60 else link_info['url']
                
                # Show row grouping for multiple links from same cell
                if current_row != link_info['row']:
                    if current_row is not None:
                        print()  # Add spacing between rows
                    print(f"  📍 Row {link_info['row']}:")
                    current_row = link_info['row']
                
                print(f"    {i}. [{link_type}] {url_preview}")
            
            # Save to file
            extractor.save_links_to_file(links, 'transfer_links.json')
            
            # Also save just the URLs to a text file for reference
            with open('transfer_urls.txt', 'w') as f:
                f.write("Transfer Links Extracted:\n")
                f.write("=" * 50 + "\n\n")
                current_row = None
                for link_info in links:
                    if current_row != link_info['row']:
                        if current_row is not None:
                            f.write("\n")
                        f.write(f"Row {link_info['row']}:\n")
                        current_row = link_info['row']
                    f.write(f"  [{link_info['type'].upper()}] {link_info['url']}\n")
            
            print(f"\n📄 URLs also saved to 'transfer_urls.txt' for reference")
            
        else:
            print("❌ No transfer links found!")
            print("Make sure your sheet contains TransferNow or WeTransfer links in the specified column.")
            
    except Exception as e:
        print(f"❌ An error occurred: {e}")

if __name__ == '__main__':
    main()