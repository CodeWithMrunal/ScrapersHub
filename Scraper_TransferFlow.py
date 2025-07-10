import os
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.common.exceptions import TimeoutException, NoSuchElementException
import zipfile
import glob

def setup_chrome_driver(download_directory):
    """Setup Chrome driver with download preferences"""
    chrome_options = Options()
    
    # Set download preferences
    prefs = {
        "download.default_directory": download_directory,
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
    driver = webdriver.Chrome(options=chrome_options)
    return driver

def wait_for_downloads_to_complete(download_directory, timeout=300):
    """Wait for all downloads to complete with progress monitoring"""
    print("Waiting for downloads to complete...")
    start_time = time.time()
    
    while time.time() - start_time < timeout:
        # Check for any .crdownload files (Chrome's temporary download files)
        crdownload_files = [f for f in os.listdir(download_directory) if f.endswith('.crdownload')]
        
        if not crdownload_files:
            print("All downloads completed!")
            return True
        
        # Show progress for each downloading file
        print(f"Still downloading... {len(crdownload_files)} files remaining")
        for crdownload_file in crdownload_files:
            file_path = os.path.join(download_directory, crdownload_file)
            try:
                file_size = os.path.getsize(file_path) / (1024 * 1024)  # Size in MB
                print(f"  - {crdownload_file}: {file_size:.1f} MB downloaded")
            except:
                print(f"  - {crdownload_file}: downloading...")
        
        time.sleep(5)
    
    print("Download timeout reached!")
    return False
def extract_zip_files(download_directory):
    """Extract all zip files in the download directory"""
    zip_files = glob.glob(os.path.join(download_directory, "*.zip"))
    
    if not zip_files:
        print("No zip files found to extract")
        return False
    
    for zip_file in zip_files:
        print(f"Extracting: {os.path.basename(zip_file)}")
        try:
            with zipfile.ZipFile(zip_file, 'r') as zip_ref:
                # Create extraction folder
                extract_folder = os.path.join(download_directory, "extracted")
                os.makedirs(extract_folder, exist_ok=True)
                
                zip_ref.extractall(extract_folder)
                print(f"Extracted to: {extract_folder}")
                
                # List extracted files
                extracted_files = zip_ref.namelist()
                print(f"Extracted {len(extracted_files)} files:")
                for file in extracted_files:
                    print(f"  - {file}")
                
        except zipfile.BadZipFile:
            print(f"Error: {zip_file} is not a valid zip file")
        except Exception as e:
            print(f"Error extracting {zip_file}: {e}")
    
    return True

def download_transfernow_files(url, download_directory):
    """Main function to download files from TransferNow"""
    
    # Create download directory if it doesn't exist
    os.makedirs(download_directory, exist_ok=True)
    print(f"Download directory: {download_directory}")
    
    # Setup driver
    driver = setup_chrome_driver(download_directory)
    
    try:
        print(f"Opening URL: {url}")
        driver.get(url)
        
        # Wait for page to load
        wait = WebDriverWait(driver, 20)
        
        # Wait for the download all button to be present and clickable
        print("Looking for 'Download all' button...")
        
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
                download_button = wait.until(EC.element_to_be_clickable((By.XPATH, selector)))
                print(f"Found download button using selector: {selector}")
                break
            except TimeoutException:
                continue
        
        if not download_button:
            print("Could not find download button. Let's check what's available on the page:")
            # Print page source for debugging
            page_text = driver.find_element(By.TAG_NAME, "body").text
            print("Page content:")
            print(page_text[:1000])  # Print first 1000 characters
            return False
        
        # Click the download button
        print("Clicking 'Download all' button...")
        driver.execute_script("arguments[0].click();", download_button)
        
        # Handle potential download dialog
        print("Checking for download dialog...")
        time.sleep(3)  # Wait a bit for any dialogs to appear
        
        # Check if there's a confirmation dialog and handle it
        try:
            # Look for common dialog buttons
            confirm_selectors = [
                "//button[contains(text(), 'Allow')]",
                "//button[contains(text(), 'OK')]",
                "//button[contains(text(), 'Yes')]",
                "//button[contains(text(), 'Download')]",
                "//button[contains(text(), 'Continue')]"
            ]
            
            for selector in confirm_selectors:
                try:
                    confirm_button = driver.find_element(By.XPATH, selector)
                    print(f"Found confirmation button: {confirm_button.text}")
                    confirm_button.click()
                    print("Clicked confirmation button")
                    break
                except NoSuchElementException:
                    continue
        except Exception as e:
            print(f"No confirmation dialog found or error handling it: {e}")
        
        # Wait a bit more for download to start
        time.sleep(5)
        
        # Check if download started by looking for files in download directory
        print("Checking if downloads have started...")
        files_before = set(os.listdir(download_directory))
        time.sleep(10)  # Wait 10 seconds
        files_after = set(os.listdir(download_directory))
        
        new_files = files_after - files_before
        if new_files or any(f.endswith('.crdownload') for f in files_after):
            print("Download started successfully!")
            
            # Wait for downloads to complete
            # Wait for downloads to complete
            if wait_for_downloads_to_complete(download_directory):
                print("All downloads completed successfully!")
                
                # List downloaded files
                final_files = [f for f in os.listdir(download_directory) if not f.endswith('.crdownload')]
                print(f"Downloaded {len(final_files)} files:")
                for file in final_files:
                    file_path = os.path.join(download_directory, file)
                    file_size = os.path.getsize(file_path) / (1024 * 1024)  # Size in MB
                    print(f"  - {file} ({file_size:.1f} MB)")
                
                # Extract zip files if any
                print("\nChecking for zip files to extract...")
                extract_zip_files(download_directory)
                
                return True
            else:
                print("Download timed out!")
                return False
        else:
            print("Download may not have started. Check manually.")
            return False
            
    except Exception as e:
        print(f"An error occurred: {e}")
        return False
    
    finally:
        # Keep browser open for a few seconds to see the result
        print("Keeping browser open for 5 seconds...")
        time.sleep(5)
        driver.quit()

if __name__ == "__main__":
    # Configuration
    URL = "https://www.transfernow.net/dl/20250710txoVHZ23"
    DOWNLOAD_DIR = os.path.join(os.getcwd(), "TransferFlow_Videos")  # Creates 'downloads' folder in current directory
    
    # You can change this to any directory you prefer
    # DOWNLOAD_DIR = "/path/to/your/download/directory"
    
    print("=== TransferNow File Downloader ===")
    print(f"URL: {URL}")
    print(f"Download Directory: {DOWNLOAD_DIR}")
    print("=" * 50)
    
    success = download_transfernow_files(URL, DOWNLOAD_DIR)
    
    if success:
        print("\n✅ Download completed successfully!")
    else:
        print("\n❌ Download failed or incomplete.")
    
    print("\nScript finished.")