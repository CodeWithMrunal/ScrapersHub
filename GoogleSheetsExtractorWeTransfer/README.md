# Transfer Link Scraper Setup Guide

## 📁 File Structure
Make sure you have these files in your project directory:

```
project_folder/
├── google_sheets_extractor.py      # Updated extractor for transfer links
├── transfer_scraper.py             # Updated scraper with WeTransfer support
├── main_runner.py                  # Main orchestration script
├── service-account-key.json        # Your Google Service Account key
├── .env                           # Environment configuration
└── requirements.txt               # Python dependencies
```

## 🔧 Setup Steps

### 1. Install Dependencies
```bash
pip install selenium google-auth google-auth-oauthlib google-auth-httplib2 google-api-python-client python-dotenv
```

### 2. Configure Environment Variables
Create a `.env` file with your Google Sheets ID:

```bash
# Your new spreadsheet ID (the one with Transfer/WeTransfer links)
NEW_SPREADSHEET_ID=your_spreadsheet_id_here

# Optional: Keep your original for SharePoint links
SPREADSHEET_ID=your_original_spreadsheet_id_here
```

To find your Spreadsheet ID:
- Open your Google Sheet
- Look at the URL: `https://docs.google.com/spreadsheets/d/[SPREADSHEET_ID]/edit`
- Copy the long string between `/d/` and `/edit`

### 3. Update Google Sheets Extractor Settings
In `google_sheets_extractor.py`, update these variables if needed:
```python
SHEET_NAME = 'Sheet1'  # Your actual sheet name
LINK_COLUMN = 'Link'   # Your actual column name
```

### 4. Set up Chrome WebDriver
Make sure you have Google Chrome installed. Selenium will automatically download the appropriate ChromeDriver.

## 🚀 Usage

### Option 1: Run Everything at Once
```bash
python main_runner.py
```
This will:
1. Extract links from Google Sheets
2. Ask for confirmation
3. Download all files
4. Generate a summary report

### Option 2: Run Steps Separately

**Step 1: Extract Links**
```bash
python google_sheets_extractor.py
```

**Step 2: Download Files**
```bash
python transfer_scraper.py
```

## 📊 Output Files

The system will create:
- `transfer_links.json` - Extracted links with metadata
- `transfer_urls.txt` - Simple list of URLs for reference
- `Downloads/Link_X/` - Downloaded files for each link
- `Downloads/Link_X/extracted/` - Extracted zip contents
- `scraping_report.txt` - Final summary report

## 🔍 Google Sheets Format

Your Google Sheet should have:
- **Sheet name**: `Sheet1` (or update in config)
- **Column name**: `Link` (or update in config)
- **Link format**: Full URLs to TransferNow or WeTransfer

Example:
```
| Link |
|------|
| https://www.transfernow.net/dl/20250710txoVHZ23 |
| https://wetransfer.com/downloads/abc123def456 |
```

## 🛠 Troubleshooting

### Google Sheets Authentication Issues
- Ensure `service-account-key.json` is in the project root
- Make sure your service account has access to the Google Sheet
- Check that the Google Sheets API is enabled in your Google Cloud Console

### Download Issues
- Check your internet connection
- Ensure you have sufficient disk space
- Some transfers may have expired - check the links manually
- WeTransfer links may have additional verification steps

### Browser Issues
- Make sure Google Chrome is installed
- If Chrome is in a non-standard location, you may need to specify the path
- Try running without headless mode to see what's happening

### Common Error Fixes
1. **"Column not found"**: Check your sheet name and column name in the config
2. **"No links found"**: Verify your links are TransferNow or WeTransfer URLs
3. **"Download timeout"**: Large files may take time, increase timeout in code
4. **"Button not found"**: Website layouts may change, check browser output

## 📝 Customization

### Timeouts
Adjust timeouts in `transfer_scraper.py`:
```python
# Download completion timeout (default: 300 seconds)
self.wait_for_downloads_to_complete(timeout=600)  # 10 minutes

# WebDriver wait timeout (default: 20 seconds)
self.wait = WebDriverWait(self.driver, 30)  # 30 seconds
```

### Download Directory
Change the base download directory in `transfer_scraper.py`:
```python
BASE_DOWNLOAD_DIR = "/path/to/your/downloads"
```

### Browser Options
Enable headless mode in `transfer_scraper.py`:
```python
chrome_options.add_argument("--headless")
```

## 🆘 Support

If you encounter issues:
1. Check the console output for error messages
2. Look at `scraping_report.txt` for detailed results
3. Try running individual links manually to test
4. Verify that the transfer links are still valid