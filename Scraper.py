import selenium
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from selenium.webdriver.common.action_chains import ActionChains
import time
import os
import re
from urllib.parse import urlparse

class SharePointVideoDownloader:
    def __init__(self, download_folder=None):
        self.download_folder = download_folder or os.path.join(os.getcwd(), "downloads")
        self.setup_driver()
        
    def setup_driver(self):
        """Setup Chrome driver with download preferences"""
        chrome_options = Options()
        
        # Create download directory if it doesn't exist
        if not os.path.exists(self.download_folder):
            os.makedirs(self.download_folder)
        
        # Set download preferences
        prefs = {
            "download.default_directory": self.download_folder,
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "safebrowsing.enabled": True,
            "profile.default_content_setting_values.notifications": 2,
            "profile.default_content_settings.popups": 0
        }
        
        chrome_options.add_experimental_option("prefs", prefs)
        chrome_options.add_argument("--no-sandbox")
        chrome_options.add_argument("--disable-dev-shm-usage")
        chrome_options.add_argument("--disable-blink-features=AutomationControlled")
        chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
        chrome_options.add_experimental_option('useAutomationExtension', False)
        
        self.driver = webdriver.Chrome(options=chrome_options)
        self.driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
        self.wait = WebDriverWait(self.driver, 20)
    
    def navigate_to_google_sheets(self, sheet_url):
        """Navigate to Google Sheets and select the correct sheet"""
        try:
            print("Navigating to Google Sheets...")
            self.driver.get(sheet_url)
            time.sleep(3)
            
            # Wait for the page to load
            self.wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))
            
            # Look for the side panel and click on "India vs Eng Women"
            self.select_sheet_tab()
            
        except Exception as e:
            print(f"Error navigating to Google Sheets: {e}")
            return False
        return True
    
    def select_sheet_tab(self):
        """Select the 'India vs Eng Women' sheet tab"""
        try:
            # Wait a bit for the sheet to fully load
            time.sleep(5)
            
            # Try different selectors for sheet tabs
            possible_selectors = [
                "//div[contains(text(), 'India vs Eng Women')]",
                "//span[contains(text(), 'India vs Eng Women')]",
                "//div[@class='docs-sheet-tab-name'][contains(text(), 'India vs Eng Women')]",
                "//*[contains(text(), 'India') and contains(text(), 'Eng') and contains(text(), 'Women')]"
            ]
            
            for selector in possible_selectors:
                try:
                    sheet_tab = self.driver.find_element(By.XPATH, selector)
                    sheet_tab.click()
                    print("Successfully clicked on 'India vs Eng Women' tab")
                    time.sleep(3)
                    return True
                except:
                    continue
            
            print("Could not find 'India vs Eng Women' tab. Proceeding with current sheet...")
            return True
            
        except Exception as e:
            print(f"Error selecting sheet tab: {e}")
            return False
    
    def extract_video_links(self):
        """Extract SharePoint video links by clicking on cells in column D"""
        try:
            print("Extracting video links by clicking cells in column D...")
            time.sleep(5)
            
            video_links = []
            
            # First, try to find all cells in the sheet
            all_cells = self.driver.find_elements(By.XPATH, "//div[contains(@class, 'cell') or contains(@role, 'gridcell')]")
            print(f"Found {len(all_cells)} total cells")
            
            # Also try alternative selectors for Google Sheets cells
            if not all_cells:
                all_cells = self.driver.find_elements(By.XPATH, "//td | //div[@role='gridcell']")
            
            processed_cells = 0
            
            for i, cell in enumerate(all_cells):
                try:
                    # Check if cell contains SharePoint text
                    cell_text = cell.text
                    if 'setindia-my.sharepoint.com' in cell_text:
                        print(f"Found SharePoint text in cell {i+1}")
                        
                        # Method 1: Try hovering over the cell
                        try:
                            from selenium.webdriver.common.action_chains import ActionChains
                            actions = ActionChains(self.driver)
                            actions.move_to_element(cell).perform()
                            time.sleep(1)
                            
                            # Look for clickable link that appears on hover
                            clickable_links = self.driver.find_elements(By.XPATH, "//a[contains(@href, 'setindia-my.sharepoint.com')]")
                            for link in clickable_links:
                                href = link.get_attribute('href')
                                if href and href not in video_links:
                                    video_links.append(href)
                                    print(f"Found clickable link on hover: {href[:80]}...")
                        except Exception as e:
                            print(f"Hover method failed for cell {i+1}: {e}")
                        
                        # Method 2: Try clicking the cell
                        try:
                            cell.click()
                            time.sleep(1)
                            
                            # Look for clickable links after clicking
                            clickable_links = self.driver.find_elements(By.XPATH, "//a[contains(@href, 'setindia-my.sharepoint.com')]")
                            for link in clickable_links:
                                href = link.get_attribute('href')
                                if href and href not in video_links:
                                    video_links.append(href)
                                    print(f"Found clickable link after click: {href[:80]}...")
                            
                            # Click somewhere else to deselect
                            self.driver.find_element(By.TAG_NAME, "body").click()
                            
                        except Exception as e:
                            print(f"Click method failed for cell {i+1}: {e}")
                        
                        processed_cells += 1
                        
                        # Limit processing to avoid timeout
                        if processed_cells > 20:
                            break
                            
                except Exception as e:
                    continue
            
            # Fallback: If no clickable links found, use page source method
            if not video_links:
                print("No clickable links found, falling back to page source method...")
                page_source = self.driver.page_source
                sharepoint_pattern = r'https://setindia-my\.sharepoint\.com[^\s<>"\'&]*'
                sharepoint_matches = re.findall(sharepoint_pattern, page_source)
                
                for match in sharepoint_matches:
                    clean_link = match.rstrip('.,;')
                    if clean_link not in video_links and len(clean_link) > 50:
                        video_links.append(clean_link)
                        print(f"Found SharePoint link via page source: {clean_link[:80]}...")
            
            print(f"Total SharePoint video links found: {len(video_links)}")
            return video_links
            
        except Exception as e:
            print(f"Error extracting video links: {e}")
            return []
    
    def download_video_from_sharepoint(self, sharepoint_url, video_index):
        """Download video from SharePoint URL"""
        try:
            print(f"Processing video {video_index + 1}: {sharepoint_url}")
            
            # Open SharePoint link in new tab
            self.driver.execute_script("window.open('');")
            self.driver.switch_to.window(self.driver.window_handles[-1])
            
            self.driver.get(sharepoint_url)
            time.sleep(5)
            
            # Wait for page to load
            self.wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))
            
            # Look for download button - try multiple selectors
            download_selectors = [
                "//button[contains(@aria-label, 'Download')]",
                "//button[contains(text(), 'Download')]",
                "//span[contains(text(), 'Download')]/../..",
                "//div[@data-automationid='DownloadCommand']",
                "//button[@data-automationid='downloadCommand']",
                "//*[@title='Download']",
                "//i[contains(@class, 'ms-Icon--Download')]/.."
            ]
            
            download_button = None
            for selector in download_selectors:
                try:
                    download_button = self.wait.until(EC.element_to_be_clickable((By.XPATH, selector)))
                    break
                except:
                    continue
            
            if not download_button:
                print(f"Could not find download button for video {video_index + 1}")
                self.driver.close()
                self.driver.switch_to.window(self.driver.window_handles[0])
                return False
            
            # Click download button
            download_button.click()
            print(f"Clicked download button for video {video_index + 1}")
            time.sleep(3)
            
            # Handle the dialog box about video-only download
            self.handle_download_dialogs()
            
            # Wait for download to start
            time.sleep(5)
            
            # Close the SharePoint tab
            self.driver.close()
            self.driver.switch_to.window(self.driver.window_handles[0])
            
            print(f"Successfully initiated download for video {video_index + 1}")
            return True
            
        except Exception as e:
            print(f"Error downloading video {video_index + 1}: {e}")
            # Close tab if it's still open
            if len(self.driver.window_handles) > 1:
                self.driver.close()
                self.driver.switch_to.window(self.driver.window_handles[0])
            return False
    
    def handle_download_dialogs(self):
        """Handle various download dialog boxes"""
        try:
            # Wait for and handle the first dialog about video-only download
            dialog_selectors = [
                "//button[contains(text(), 'Download')]",
                "//button[contains(@aria-label, 'Download')]",
                "//div[contains(@class, 'dialog')]//button[contains(text(), 'Download')]",
                "//div[contains(@class, 'modal')]//button[contains(text(), 'Download')]"
            ]
            
            # Try to find and click the dialog download button
            for selector in dialog_selectors:
                try:
                    dialog_button = self.wait.until(EC.element_to_be_clickable((By.XPATH, selector)))
                    dialog_button.click()
                    print("Clicked download button in dialog")
                    time.sleep(2)
                    break
                except:
                    continue
            
            # Handle potential "Allow download" dialog
            allow_selectors = [
                "//button[contains(text(), 'Allow')]",
                "//button[contains(@aria-label, 'Allow')]",
                "//button[contains(text(), 'Yes')]",
                "//button[contains(text(), 'OK')]"
            ]
            
            for selector in allow_selectors:
                try:
                    allow_button = WebDriverWait(self.driver, 5).until(
                        EC.element_to_be_clickable((By.XPATH, selector))
                    )
                    allow_button.click()
                    print("Clicked allow button in dialog")
                    time.sleep(2)
                    break
                except:
                    continue
                    
        except TimeoutException:
            print("No additional dialogs found or they timed out")
        except Exception as e:
            print(f"Error handling dialogs: {e}")
    
    def run(self, sheet_url):
        """Main method to run the entire process"""
        try:
            print("Starting SharePoint video downloader...")
            print(f"Download folder: {self.download_folder}")
            
            # Navigate to Google Sheets
            if not self.navigate_to_google_sheets(sheet_url):
                return False
            
            # Extract video links
            video_links = self.extract_video_links()
            
            if not video_links:
                print("No SharePoint video links found in the sheet")
                return False
            
            # Download each video
            successful_downloads = 0
            for i, link in enumerate(video_links):
                print(f"\n--- Processing video {i + 1} of {len(video_links)} ---")
                if self.download_video_from_sharepoint(link, i):
                    successful_downloads += 1
                    # Wait between downloads to avoid overwhelming the server
                    time.sleep(10)
                else:
                    print(f"Failed to download video {i + 1}")
            
            print(f"\n--- Download Summary ---")
            print(f"Total videos found: {len(video_links)}")
            print(f"Successful downloads: {successful_downloads}")
            print(f"Failed downloads: {len(video_links) - successful_downloads}")
            
            return True
            
        except Exception as e:
            print(f"Error in main process: {e}")
            return False
        finally:
            self.cleanup()
    
    def cleanup(self):
        """Clean up resources"""
        try:
            self.driver.quit()
            print("Browser closed successfully")
        except:
            pass

# Usage
def main():
    # Configuration
    SHEET_URL = "https://docs.google.com/spreadsheets/d/1dHItb5n2rNc_6v7LZ6XJeMeQzsO33aNzz5bFmNRkeg0/edit?usp=sharing"
    DOWNLOAD_FOLDER = os.path.join(os.getcwd(), "SharePoint_Videos")  # Change this to your desired folder
    
    # Create downloader instance
    downloader = SharePointVideoDownloader(download_folder=DOWNLOAD_FOLDER)
    
    # Run the process
    downloader.run(SHEET_URL)

if __name__ == "__main__":
    main()