import os
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import TimeoutException, NoSuchElementException
import zipfile
import glob

# CONFIGURATION - Update these values
WETRANSFER_URL = "https://we.tl/t-UL1BUR0T7q"  # Replace with your WeTransfer link
DOWNLOAD_FOLDER = "WeTransfer_Downloads"  # Folder name where files will be saved

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
    
    # Disable notifications and popups
    chrome_options.add_argument("--disable-notifications")
    chrome_options.add_argument("--disable-popup-blocking")
    chrome_options.add_argument("--disable-blink-features=AutomationControlled")
    
    # Optional: Uncomment to run in headless mode (no browser window)
    # chrome_options.add_argument("--headless")
    
    # Create driver
    driver = webdriver.Chrome(options=chrome_options)
    
    # Make browser look more human-like
    driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
    
    return driver

def wait_for_downloads_to_complete(download_directory, timeout=600):
    """Wait for all downloads to complete with progress monitoring"""
    print("⏳ Waiting for downloads to complete...")
    start_time = time.time()
    
    while time.time() - start_time < timeout:
        # Check for Chrome's temporary download files
        temp_files = []
        for ext in ['.crdownload', '.tmp', '.part']:
            temp_files.extend([f for f in os.listdir(download_directory) if f.endswith(ext)])
        
        if not temp_files:
            print("✅ All downloads completed!")
            return True
        
        # Show progress
        print(f"📥 Still downloading... {len(temp_files)} files in progress")
        for temp_file in temp_files:
            file_path = os.path.join(download_directory, temp_file)
            try:
                file_size = os.path.getsize(file_path) / (1024 * 1024)  # Size in MB
                print(f"  - {temp_file}: {file_size:.1f} MB")
            except:
                print(f"  - {temp_file}: downloading...")
        
        time.sleep(10)
    
    print("⚠️ Download timeout reached!")
    return False

def extract_archives(download_directory):
    """Extract any zip files found in the download directory"""
    archive_files = []
    for ext in ['*.zip', '*.rar', '*.7z']:
        archive_files.extend(glob.glob(os.path.join(download_directory, ext)))
    
    if not archive_files:
        print("ℹ️ No archive files found to extract")
        return
    
    for archive_file in archive_files:
        print(f"📦 Extracting: {os.path.basename(archive_file)}")
        try:
            if archive_file.endswith('.zip'):
                with zipfile.ZipFile(archive_file, 'r') as zip_ref:
                    extract_folder = os.path.join(download_directory, "extracted")
                    os.makedirs(extract_folder, exist_ok=True)
                    zip_ref.extractall(extract_folder)
                    
                    extracted_files = zip_ref.namelist()
                    print(f"✅ Extracted {len(extracted_files)} files to: {extract_folder}")
                    for file in extracted_files:
                        print(f"  - {file}")
            else:
                print(f"⚠️ Unsupported archive format: {archive_file}")
                
        except Exception as e:
            print(f"❌ Error extracting {archive_file}: {e}")

def download_wetransfer_files(url, download_directory):
    """Main function to download files from WeTransfer"""
    
    # Create download directory
    os.makedirs(download_directory, exist_ok=True)
    download_path = os.path.abspath(download_directory)
    
    print(f"🚀 WeTransfer Downloader")
    print(f"📄 URL: {url}")
    print(f"📂 Download Directory: {download_path}")
    print("=" * 60)
    
    # Setup driver
    driver = setup_chrome_driver(download_path)
    
    try:
        print("🌐 Opening WeTransfer link...")
        driver.get(url)
        
        # Wait for page to load and handle redirects (we.tl links redirect)
        wait = WebDriverWait(driver, 30)
        time.sleep(5)  # Give time for any redirects
        
        print(f"🔍 Current page: {driver.current_url}")
        print(f"📄 Page title: {driver.title}")
        
        # Sometimes WeTransfer shows an age verification, terms, or cookie consent page first
        print("🔍 Checking for any initial dialogs or consent pages...")
        time.sleep(3)
        
        # Look for common dialog buttons that might appear first
        initial_buttons = [
            "//button[contains(text(), 'I agree')]",
            "//button[contains(text(), 'Accept')]",
            "//button[contains(text(), 'Accept all')]",
            "//button[contains(text(), 'Continue')]",
            "//button[contains(text(), 'Proceed')]",
            "//button[contains(text(), 'Allow')]",
            "//a[contains(text(), 'Continue')]",
            "//button[contains(@class, 'accept')]",
            "//button[contains(@class, 'consent')]"
        ]
        
        for selector in initial_buttons:
            try:
                button = driver.find_element(By.XPATH, selector)
                if button.is_displayed():
                    print(f"🔘 Found initial dialog button: '{button.text}' - clicking...")
                    driver.execute_script("arguments[0].click();", button)
                    time.sleep(5)  # Wait after clicking
                    break
            except NoSuchElementException:
                continue
        
        # Wait for the main transfer page to load
        print("⏳ Waiting for transfer page to fully load...")
        time.sleep(5)
        
        # Now look for the main download button
        print("🔍 Looking for 'Download' button...")
        
        # Extended WeTransfer download button selectors for different layouts
        download_selectors = [
            # Standard download buttons
            "//button[contains(text(), 'Download') and not(contains(text(), 'Scan'))]",
            "//a[contains(text(), 'Download') and not(contains(text(), 'Scan'))]",
            
            # Specific WeTransfer selectors
            "//button[@data-testid='download-button']",
            "//button[contains(@class, 'download') and not(contains(text(), 'Scan'))]",
            "//button[contains(@class, 'Download') and not(contains(text(), 'Scan'))]",
            "//*[@role='button'][contains(text(), 'Download') and not(contains(text(), 'Scan'))]",
            
            # More generic selectors
            "//button[contains(@class, 'Button') and contains(text(), 'Download')]",
            "//div[contains(@class, 'download')]/button",
            "//button[contains(@class, 'primary') and contains(text(), 'Download')]",
            
            # For we.tl short links that might have different layouts
            "//button[text()='Download']",
            "//a[text()='Download']",
            
            # Try without text matching - just look for download-related classes
            "//button[contains(@class, 'download')]",
            "//a[contains(@class, 'download')]"
        ]
        
        download_button = None
        for i, selector in enumerate(download_selectors):
            try:
                print(f"  Trying selector {i+1}/{len(download_selectors)}...")
                potential_buttons = driver.find_elements(By.XPATH, selector)
                
                for button in potential_buttons:
                    if button.is_displayed() and button.is_enabled():
                        button_text = button.text.lower() if button.text else ""
                        
                        # Skip scan buttons and make sure it's a download button
                        if 'scan' not in button_text and ('download' in button_text or 'download' in button.get_attribute('class').lower()):
                            download_button = button
                            print(f"✅ Found download button: '{button.text or button.get_attribute('class')}'")
                            break
                
                if download_button:
                    break
                    
            except Exception as e:
                print(f"  Selector {i+1} failed: {e}")
                continue
        
        if not download_button:
            print("❌ Could not find download button!")
            print("🔍 Available buttons on the page:")
            buttons = driver.find_elements(By.TAG_NAME, "button")
            links = driver.find_elements(By.TAG_NAME, "a")
            
            for element in buttons + links:
                try:
                    if element.text and element.is_displayed():
                        print(f"  - {element.tag_name.upper()}: '{element.text}'")
                except:
                    pass
            
            print("\n📄 Page title:", driver.title)
            print("🌐 Current URL:", driver.current_url)
            return False
        
        # Scroll to button and click it
        print(f"🖱️ Clicking '{download_button.text}' button...")
        driver.execute_script("arguments[0].scrollIntoView(true);", download_button)
        time.sleep(2)
        driver.execute_script("arguments[0].click();", download_button)
        
        # Wait for any additional dialogs or redirects
        print("⏳ Waiting for download to start...")
        time.sleep(10)
        
        # Check for any confirmation dialogs
        confirmation_selectors = [
            "//button[contains(text(), 'Allow')]",
            "//button[contains(text(), 'Save')]",
            "//button[contains(text(), 'OK')]",
            "//button[contains(text(), 'Yes')]",
            "//button[contains(text(), 'Continue')]"
        ]
        
        for selector in confirmation_selectors:
            try:
                confirm_button = driver.find_element(By.XPATH, selector)
                if confirm_button.is_displayed():
                    print(f"🔘 Found confirmation button: '{confirm_button.text}' - clicking...")
                    confirm_button.click()
                    time.sleep(3)
                    break
            except NoSuchElementException:
                continue
        
        # Monitor download progress
        print("🔍 Checking if download has started...")
        
        # Check files before and after waiting
        files_before = set(os.listdir(download_directory))
        time.sleep(15)  # Wait for download to start
        files_after = set(os.listdir(download_directory))
        
        new_files = files_after - files_before
        temp_files = [f for f in files_after if f.endswith(('.crdownload', '.tmp', '.part'))]
        
        if new_files or temp_files:
            print("✅ Download started successfully!")
            
            # Wait for downloads to complete
            if wait_for_downloads_to_complete(download_directory):
                print("🎉 Download completed successfully!")
                
                # List final files
                final_files = [f for f in os.listdir(download_directory) 
                              if not f.endswith(('.crdownload', '.tmp', '.part'))]
                
                if final_files:
                    print(f"📁 Downloaded {len(final_files)} files:")
                    for file in final_files:
                        file_path = os.path.join(download_directory, file)
                        file_size = os.path.getsize(file_path) / (1024 * 1024)  # MB
                        print(f"  - {file} ({file_size:.1f} MB)")
                    
                    # Extract any archives
                    print("\n📦 Checking for archives to extract...")
                    extract_archives(download_directory)
                    
                    return True
                else:
                    print("⚠️ No files found in download directory")
                    return False
            else:
                print("❌ Download timed out!")
                return False
        else:
            print("❌ Download may not have started")
            print("🔍 Please check manually or verify the WeTransfer link is valid")
            print(f"Current page: {driver.current_url}")
            return False
            
    except Exception as e:
        print(f"❌ An error occurred: {e}")
        return False
    
    finally:
        print("\n⏳ Keeping browser open for 5 seconds to see results...")
        time.sleep(5)
        print("🚪 Closing browser...")
        driver.quit()

def main():
    """Main execution function"""
    print("🎯 WeTransfer Standalone Downloader")
    print("=" * 50)
    
    # Validate URL - support both wetransfer.com and we.tl formats
    url_lower = WETRANSFER_URL.lower()
    if not ("wetransfer.com" in url_lower or "we.tl" in url_lower):
        print("❌ Error: Please update WETRANSFER_URL with a valid WeTransfer link")
        print(f"Current URL: {WETRANSFER_URL}")
        print("💡 Supported formats:")
        print("  - https://wetransfer.com/downloads/...")
        print("  - https://we.tl/t-...")
        return
    
    # Create full download path
    download_path = os.path.join(os.getcwd(), DOWNLOAD_FOLDER)
    
    print(f"📄 WeTransfer URL: {WETRANSFER_URL}")
    print(f"📂 Download Folder: {download_path}")
    
    # Confirm before proceeding
    user_input = input("\n🚀 Start download? (y/N): ").lower().strip()
    if user_input not in ['y', 'yes']:
        print("❌ Download cancelled by user")
        return
    
    # Start download
    success = download_wetransfer_files(WETRANSFER_URL, download_path)
    
    print("\n" + "=" * 50)
    if success:
        print("✅ Download completed successfully!")
        print(f"📁 Files saved to: {download_path}")
    else:
        print("❌ Download failed or incomplete")
        print("💡 Tips:")
        print("  - Check if the WeTransfer link is still valid")
        print("  - Try running the script again")
        print("  - Check your internet connection")
        print("  - Some we.tl links may redirect multiple times")
    print("=" * 50)

if __name__ == "__main__":
    main()