import os
import time
import json
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.common.exceptions import TimeoutException, NoSuchElementException
import zipfile
import glob

class TransferScraper:
    def __init__(self, download_directory):
        self.download_directory = download_directory
        self.driver = None
        self.wait = None
        
    def setup_chrome_driver(self):
        """Setup Chrome driver with download preferences"""
        chrome_options = Options()
        
        # Set download preferences
        prefs = {
            "download.default_directory": self.download_directory,
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "safebrowsing.enabled": True
        }
        chrome_options.add_experimental_option("prefs", prefs)
        
        # Optional: Run in headless mode (uncomment if you don't want to see the browser)
        # chrome_options.add_argument("--headless")
        
        # Disable notifications
        chrome_options.add_argument("--disable-notifications")
        
        # Create driver
        self.driver = webdriver.Chrome(options=chrome_options)
        self.wait = WebDriverWait(self.driver, 20)
        return self.driver

    def wait_for_downloads_to_complete(self, timeout=300):
        """Wait for all downloads to complete with progress monitoring"""
        print("⏳ Waiting for downloads to complete...")
        start_time = time.time()
        
        while time.time() - start_time < timeout:
            # Check for any .crdownload files (Chrome's temporary download files)
            crdownload_files = [f for f in os.listdir(self.download_directory) if f.endswith('.crdownload')]
            
            if not crdownload_files:
                print("✅ All downloads completed!")
                return True
            
            # Show progress for each downloading file
            print(f"📥 Still downloading... {len(crdownload_files)} files remaining")
            for crdownload_file in crdownload_files:
                file_path = os.path.join(self.download_directory, crdownload_file)
                try:
                    file_size = os.path.getsize(file_path) / (1024 * 1024)  # Size in MB
                    print(f"  - {crdownload_file}: {file_size:.1f} MB downloaded")
                except:
                    print(f"  - {crdownload_file}: downloading...")
            
            time.sleep(5)
        
        print("⚠️ Download timeout reached!")
        return False

    def extract_zip_files(self):
        """Extract all zip files in the download directory"""
        zip_files = glob.glob(os.path.join(self.download_directory, "*.zip"))
        
        if not zip_files:
            print("ℹ️ No zip files found to extract")
            return False
        
        for zip_file in zip_files:
            print(f"📦 Extracting: {os.path.basename(zip_file)}")
            try:
                with zipfile.ZipFile(zip_file, 'r') as zip_ref:
                    # Create extraction folder
                    extract_folder = os.path.join(self.download_directory, "extracted")
                    os.makedirs(extract_folder, exist_ok=True)
                    
                    zip_ref.extractall(extract_folder)
                    print(f"✅ Extracted to: {extract_folder}")
                    
                    # List extracted files
                    extracted_files = zip_ref.namelist()
                    print(f"📁 Extracted {len(extracted_files)} files:")
                    for file in extracted_files:
                        print(f"  - {file}")
                    
            except zipfile.BadZipFile:
                print(f"❌ Error: {zip_file} is not a valid zip file")
            except Exception as e:
                print(f"❌ Error extracting {zip_file}: {e}")
        
        return True

    def download_transfernow_files(self, url):
        """Download files from TransferNow"""
        print(f"🚀 Processing TransferNow URL: {url}")
        
        try:
            self.driver.get(url)
            
            # Wait for page to load
            print("⏳ Looking for 'Download all' button...")
            
            # Try multiple possible selectors for the download button
            download_button_selectors = [
                "//button[contains(text(), 'Download all')]",
                "//a[contains(text(), 'Download all')]",
                "//button[contains(@class, 'download') and contains(text(), 'Download')]",
                "//a[contains(@class, 'download') and contains(text(), 'Download')]",
                "//*[contains(text(), 'Download all')]"
            ]
            
            download_button = None
            for selector in download_button_selectors:
                try:
                    download_button = self.wait.until(EC.element_to_be_clickable((By.XPATH, selector)))
                    print(f"✅ Found download button using selector: {selector}")
                    break
                except TimeoutException:
                    continue
            
            if not download_button:
                print("❌ Could not find download button. Page content:")
                page_text = self.driver.find_element(By.TAG_NAME, "body").text
                print(page_text[:1000])
                return False
            
            # Click the download button
            print("🖱️ Clicking 'Download all' button...")
            self.driver.execute_script("arguments[0].click();", download_button)
            
            # Handle potential download dialog
            print("🔍 Checking for download dialog...")
            time.sleep(3)
            
            # Check if there's a confirmation dialog and handle it
            self.handle_confirmation_dialog()
            
            return self.monitor_download_progress()
            
        except Exception as e:
            print(f"❌ Error processing TransferNow: {e}")
            return False

    def download_wetransfer_files(self, url):
        """Download files from WeTransfer"""
        print(f"🚀 Processing WeTransfer URL: {url}")
        
        try:
            self.driver.get(url)
            
            # Wait for page to load
            print("⏳ Looking for 'Download' button...")
            
            # WeTransfer specific selectors for the Download button
            # We want the "Download" button, not the "Scan and download" button
            download_button_selectors = [
                "//button[contains(text(), 'Download') and not(contains(text(), 'Scan'))]",
                "//a[contains(text(), 'Download') and not(contains(text(), 'Scan'))]",
                "//button[contains(@class, 'download') and contains(text(), 'Download') and not(contains(text(), 'Scan'))]",
                "//a[contains(@class, 'download') and contains(text(), 'Download') and not(contains(text(), 'Scan'))]",
                "//button[@data-testid='download-button']",
                "//button[contains(@class, 'button--download')]"
            ]
            
            download_button = None
            for selector in download_button_selectors:
                try:
                    download_button = self.wait.until(EC.element_to_be_clickable((By.XPATH, selector)))
                    button_text = download_button.text.lower()
                    # Make sure we don't click the "Scan and download" button
                    if 'scan' not in button_text:
                        print(f"✅ Found download button: '{download_button.text}'")
                        break
                except TimeoutException:
                    continue
            
            if not download_button:
                print("❌ Could not find download button. Available buttons:")
                buttons = self.driver.find_elements(By.TAG_NAME, "button")
                for btn in buttons:
                    try:
                        if btn.text:
                            print(f"  - Button: '{btn.text}'")
                    except:
                        pass
                return False
            
            # Click the download button
            print(f"🖱️ Clicking '{download_button.text}' button...")
            self.driver.execute_script("arguments[0].click();", download_button)
            
            # WeTransfer may have additional steps or dialogs
            print("⏳ Waiting for download to start...")
            time.sleep(5)
            
            # Handle potential dialogs
            self.handle_confirmation_dialog()
            
            # WeTransfer might redirect or show additional UI elements
            # Wait a bit more for the actual download to start
            time.sleep(10)
            
            return self.monitor_download_progress()
            
        except Exception as e:
            print(f"❌ Error processing WeTransfer: {e}")
            return False

    def handle_confirmation_dialog(self):
        """Handle potential confirmation dialogs"""
        try:
            confirm_selectors = [
                "//button[contains(text(), 'Allow')]",
                "//button[contains(text(), 'OK')]",
                "//button[contains(text(), 'Yes')]",
                "//button[contains(text(), 'Download')]",
                "//button[contains(text(), 'Continue')]",
                "//button[contains(text(), 'Accept')]"
            ]
            
            for selector in confirm_selectors:
                try:
                    confirm_button = self.driver.find_element(By.XPATH, selector)
                    print(f"🔘 Found confirmation button: {confirm_button.text}")
                    confirm_button.click()
                    print("✅ Clicked confirmation button")
                    time.sleep(2)
                    break
                except NoSuchElementException:
                    continue
        except Exception as e:
            print(f"ℹ️ No confirmation dialog found: {e}")

    def monitor_download_progress(self):
        """Monitor download progress and return success status"""
        # Wait a bit more for download to start
        time.sleep(5)
        
        # Check if download started
        print("🔍 Checking if downloads have started...")
        files_before = set(os.listdir(self.download_directory))
        time.sleep(10)  # Wait 10 seconds
        files_after = set(os.listdir(self.download_directory))
        
        new_files = files_after - files_before
        if new_files or any(f.endswith('.crdownload') for f in files_after):
            print("✅ Download started successfully!")
            
            # Wait for downloads to complete
            if self.wait_for_downloads_to_complete():
                print("🎉 All downloads completed successfully!")
                
                # List downloaded files
                final_files = [f for f in os.listdir(self.download_directory) if not f.endswith('.crdownload')]
                print(f"📁 Downloaded {len(final_files)} files:")
                for file in final_files:
                    file_path = os.path.join(self.download_directory, file)
                    file_size = os.path.getsize(file_path) / (1024 * 1024)  # Size in MB
                    print(f"  - {file} ({file_size:.1f} MB)")
                
                # Extract zip files if any
                print("\n📦 Checking for zip files to extract...")
                self.extract_zip_files()
                
                return True
            else:
                print("⚠️ Download timed out!")
                return False
        else:
            print("❌ Download may not have started. Check manually.")
            return False

    def process_link(self, link_data):
        """Process a single link based on its type"""
        url = link_data['url']
        link_type = link_data['type']
        
        print(f"\n{'='*60}")
        print(f"Processing Link ID: {link_data['id']}")
        print(f"Type: {link_type.upper()}")
        print(f"URL: {url}")
        print(f"{'='*60}")
        
        if link_type == 'transfernow':
            return self.download_transfernow_files(url)
        elif link_type == 'wetransfer':
            return self.download_wetransfer_files(url)
        else:
            print(f"❌ Unsupported link type: {link_type}")
            return False

    def close(self):
        """Close the browser"""
        if self.driver:
            print("🚪 Closing browser...")
            time.sleep(3)  # Keep browser open briefly to see results
            self.driver.quit()

def load_links_from_json(filename='transfer_links.json'):
    """Load links from the JSON file created by the Google Sheets extractor"""
    try:
        with open(filename, 'r') as f:
            data = json.load(f)
        
        if 'links' in data:
            return data['links'], data.get('metadata', {})
        else:
            # Fallback for older format
            return data, {}
            
    except FileNotFoundError:
        print(f"❌ Error: {filename} not found!")
        print("Please run the Google Sheets extractor first to generate the links file.")
        return [], {}
    except json.JSONDecodeError:
        print(f"❌ Error: Invalid JSON in {filename}")
        return [], {}
    except Exception as e:
        print(f"❌ Error loading links: {e}")
        return [], {}

def update_link_status(filename, link_id, status, processed=None, error_message=None):
    """Update the status and processed field of a link in the JSON file"""
    try:
        with open(filename, 'r') as f:
            data = json.load(f)
        
        # Find and update the link
        for link in data['links']:
            if link['id'] == link_id:
                link['status'] = status
                link['processed_at'] = datetime.now().isoformat()
                if processed is not None:
                    link['processed'] = processed
                if error_message:
                    link['error'] = error_message
                break
        
        # Save updated data
        with open(filename, 'w') as f:
            json.dump(data, f, indent=2)
            
    except Exception as e:
        print(f"⚠️ Warning: Could not update link status: {e}")

def main():
    # Configuration
    LINKS_FILE = "transfer_links.json"
    BASE_DOWNLOAD_DIR = os.path.join(os.getcwd(), "Downloads")
    
    print("🚀 Starting Transfer Link Scraper")
    print(f"📂 Base Download Directory: {BASE_DOWNLOAD_DIR}")
    print(f"📄 Links File: {LINKS_FILE}")
    print("=" * 80)
    
    # Load links from JSON file
    links, metadata = load_links_from_json(LINKS_FILE)
    
    if not links:
        print("❌ No links to process!")
        return
    
    # Filter out already processed links
    unprocessed_links = [link for link in links if link.get('processed', 0) == 0]
    processed_count = len(links) - len(unprocessed_links)
    
    print(f"📊 Link Status:")
    print(f"  - Total links: {len(links)}")
    print(f"  - Already processed: {processed_count}")
    print(f"  - To be processed: {len(unprocessed_links)}")
    
    if metadata:
        print(f"📈 Link Types:")
        print(f"  - TransferNow: {metadata.get('transfernow_count', 0)}")
        print(f"  - WeTransfer: {metadata.get('wetransfer_count', 0)}")
    
    if not unprocessed_links:
        print("✅ All links have already been processed!")
        return
    
    # Process each unprocessed link
    successful_downloads = 0
    failed_downloads = 0
    
    for i, link_data in enumerate(unprocessed_links, 1):
        print(f"\n🔄 Processing link {i}/{len(unprocessed_links)} (ID: {link_data['id']})")
        print(f"📍 From Row {link_data['row']} in Google Sheet")
        
        # Mark as being processed
        update_link_status(LINKS_FILE, link_data['id'], 'processing', processed=0)
        
        # Create a specific download directory for this link
        link_download_dir = os.path.join(BASE_DOWNLOAD_DIR, f"Link_{link_data['id']}")
        os.makedirs(link_download_dir, exist_ok=True)
        
        # Create scraper instance
        scraper = TransferScraper(link_download_dir)
        
        try:
            # Setup browser
            scraper.setup_chrome_driver()
            
            # Process the link
            success = scraper.process_link(link_data)
            
            if success:
                print(f"✅ Successfully processed link {i}")
                successful_downloads += 1
                update_link_status(LINKS_FILE, link_data['id'], 'completed', processed=1)
            else:
                print(f"❌ Failed to process link {i}")
                failed_downloads += 1
                update_link_status(LINKS_FILE, link_data['id'], 'failed', processed=1, error_message='Download failed')
                
        except Exception as e:
            print(f"❌ Error processing link {i}: {e}")
            failed_downloads += 1
            update_link_status(LINKS_FILE, link_data['id'], 'error', processed=1, error_message=str(e))
            
        finally:
            scraper.close()
            
        # Add delay between downloads to be respectful
        if i < len(unprocessed_links):
            print("⏳ Waiting 10 seconds before next download...")
            time.sleep(10)
    
    # Final summary
    print("\n" + "=" * 80)
    print("📊 FINAL SUMMARY")
    print("=" * 80)
    print(f"✅ Successful downloads: {successful_downloads}")
    print(f"❌ Failed downloads: {failed_downloads}")
    print(f"📁 Download directory: {BASE_DOWNLOAD_DIR}")
    print("=" * 80)

if __name__ == "__main__":
    main()