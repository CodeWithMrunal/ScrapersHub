#!/usr/bin/env python3
"""
Main runner script for the Transfer Link Scraping System
This script orchestrates the entire process:
1. Extract links from Google Sheets
2. Process and download files from those links
"""

import os
import sys
import subprocess
import json
from datetime import datetime

def check_dependencies():
    """Check if required files and dependencies exist"""
    required_files = [
        'google_sheets_extractor.py',
        'transfer_scraper.py',
        'service-account-key.json',
        '.env'
    ]
    
    missing_files = []
    for file in required_files:
        if not os.path.exists(file):
            missing_files.append(file)
    
    if missing_files:
        print("❌ Missing required files:")
        for file in missing_files:
            print(f"  - {file}")
        print("\nPlease ensure all required files are present.")
        return False
    
    return True

def run_extraction():
    """Run the Google Sheets extraction"""
    print("🚀 Step 1: Extracting links from Google Sheets...")
    print("-" * 50)
    
    try:
        result = subprocess.run([sys.executable, 'google_sheets_extractor.py'], 
                              capture_output=True, text=True, timeout=120)
        
        if result.returncode == 0:
            print("✅ Link extraction completed successfully!")
            print(result.stdout)
            return True
        else:
            print("❌ Link extraction failed!")
            print("Error output:", result.stderr)
            return False
            
    except subprocess.TimeoutExpired:
        print("❌ Link extraction timed out!")
        return False
    except Exception as e:
        print(f"❌ Error running extraction: {e}")
        return False

def check_extracted_links():
    """Check if links were successfully extracted"""
    try:
        with open('transfer_links.json', 'r') as f:
            data = json.load(f)
        
        links = data.get('links', [])
        metadata = data.get('metadata', {})
        
        if not links:
            print("❌ No links were extracted!")
            return False
        
        print(f"📊 Found {len(links)} links:")
        print(f"  - TransferNow: {metadata.get('transfernow_count', 0)}")
        print(f"  - WeTransfer: {metadata.get('wetransfer_count', 0)}")
        
        return True
        
    except FileNotFoundError:
        print("❌ transfer_links.json not found!")
        return False
    except Exception as e:
        print(f"❌ Error checking extracted links: {e}")
        return False

def run_scraping():
    """Run the transfer link scraping"""
    print("\n🚀 Step 2: Scraping and downloading files...")
    print("-" * 50)
    
    try:
        # Run the scraper
        result = subprocess.run([sys.executable, 'transfer_scraper.py'], 
                              text=True, timeout=7200)  # 2 hour timeout
        
        if result.returncode == 0:
            print("✅ File scraping completed!")
            return True
        else:
            print("❌ File scraping completed with some errors.")
            print("Check the output above for details.")
            return True  # Return True as partial success is still useful
            
    except subprocess.TimeoutExpired:
        print("⚠️ File scraping timed out after 2 hours!")
        print("Some downloads may still be in progress.")
        return False
    except Exception as e:
        print(f"❌ Error running scraper: {e}")
        return False

def generate_summary_report():
    """Generate a summary report of the entire process"""
    try:
        with open('transfer_links.json', 'r') as f:
            data = json.load(f)
        
        links = data.get('links', [])
        metadata = data.get('metadata', {})
        
        # Count statuses and processed
        completed = len([l for l in links if l.get('status') == 'completed'])
        failed = len([l for l in links if l.get('status') == 'failed'])
        error = len([l for l in links if l.get('status') == 'error'])
        pending = len([l for l in links if l.get('status') == 'pending'])
        processed = len([l for l in links if l.get('processed', 0) == 1])
        unprocessed = len([l for l in links if l.get('processed', 0) == 0])
        
        # Group by rows to show multiple links per cell
        rows = {}
        for link in links:
            row_num = link['row']
            if row_num not in rows:
                rows[row_num] = []
            rows[row_num].append(link)
        
        # Generate report
        report = f"""
{'='*80}
TRANSFER LINK SCRAPER - FINAL REPORT
{'='*80}
Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}

EXTRACTION SUMMARY:
- Total links found: {len(links)}
- Links from {len(rows)} Google Sheet rows
- TransferNow links: {metadata.get('transfernow_count', 0)}
- WeTransfer links: {metadata.get('wetransfer_count', 0)}

PROCESSING SUMMARY:
- ✅ Successfully completed: {completed}
- ❌ Failed: {failed}
- ⚠️ Errors: {error}
- ⏳ Pending: {pending}
- 🔄 Total processed: {processed}
- 📋 Unprocessed: {unprocessed}

DOWNLOAD LOCATION:
- Files downloaded to: {os.path.join(os.getcwd(), 'Downloads')}

DETAILED RESULTS (Grouped by Google Sheet Row):
"""
        
        for row_num in sorted(rows.keys()):
            row_links = rows[row_num]
            report += f"\n📍 Google Sheet Row {row_num} ({len(row_links)} links):\n"
            
            for link in row_links:
                status_emoji = {
                    'completed': '✅',
                    'failed': '❌',
                    'error': '⚠️',
                    'pending': '⏳',
                    'processing': '🔄'
                }.get(link.get('status', 'pending'), '❓')
                
                processed_emoji = '🔄' if link.get('processed', 0) == 1 else '📋'
                
                report += f"  {status_emoji}{processed_emoji} [{link['type'].upper()}] {link['id']}: {link['url'][:60]}...\n"
                if link.get('error'):
                    report += f"      Error: {link['error']}\n"
        
        report += f"\n{'='*80}\n"
        report += "Legend: ✅=Success ❌=Failed ⚠️=Error ⏳=Pending 🔄=Processed 📋=Unprocessed\n"
        report += f"{'='*80}\n"
        
        # Save report to file
        with open('scraping_report.txt', 'w') as f:
            f.write(report)
        
        print(report)
        print(f"📄 Full report saved to: scraping_report.txt")
        
    except Exception as e:
        print(f"⚠️ Could not generate summary report: {e}")

def main():
    """Main execution function"""
    print("🎯 TRANSFER LINK SCRAPING SYSTEM")
    print("=" * 80)
    print("This system will:")
    print("1. Extract transfer links from your Google Sheet")
    print("2. Download files from TransferNow and WeTransfer links")
    print("3. Extract any zip files found")
    print("4. Generate a summary report")
    print("=" * 80)
    
    # Check dependencies
    print("\n🔍 Checking dependencies...")
    if not check_dependencies():
        return 1
    print("✅ All dependencies found!")
    
    # Step 1: Extract links from Google Sheets
    if not run_extraction():
        print("❌ Process stopped due to extraction failure.")
        return 1
    
    # Check if links were extracted
    if not check_extracted_links():
        print("❌ Process stopped - no links to process.")
        return 1
    
    # Ask user confirmation before scraping
    print("\n" + "⚠️ " * 20)
    print("WARNING: The scraping process will start downloading files.")
    print("This may take a long time and use significant bandwidth.")
    print("Make sure you have enough disk space and a stable internet connection.")
    print("⚠️ " * 20)
    
    user_input = input("\nDo you want to continue with the download process? (y/N): ").lower().strip()
    if user_input not in ['y', 'yes']:
        print("❌ Process cancelled by user.")
        return 0
    
    # Step 2: Scrape and download files
    scraping_success = run_scraping()
    
    # Step 3: Generate summary report
    print("\n🚀 Step 3: Generating summary report...")
    print("-" * 50)
    generate_summary_report()
    
    if scraping_success:
        print("\n🎉 Process completed successfully!")
        print("Check the Downloads folder for your files.")
    else:
        print("\n⚠️ Process completed with some issues.")
        print("Check the report above and Downloads folder for partial results.")
    
    return 0

if __name__ == "__main__":
    try:
        exit_code = main()
        sys.exit(exit_code)
    except KeyboardInterrupt:
        print("\n\n❌ Process interrupted by user (Ctrl+C)")
        print("Any downloads in progress may continue in the background.")
        sys.exit(1)
    except Exception as e:
        print(f"\n❌ Unexpected error: {e}")
        sys.exit(1)