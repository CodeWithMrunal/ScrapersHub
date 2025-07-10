import json
import time
import os
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import TimeoutException, NoSuchElementException

class SimpleSharePointDownloader:
    def __init__(self, download_folder="downloads", headless=False):
        self.download_folder = os.path.abspath(download_folder)
        self.setup_driver(headless)
        
        # Create download folder if it doesn't exist
        os.makedirs(self.download_folder, exist_ok=True)
        print(f"Downloads will be saved to: {self.download_folder}")
    
    def setup_driver(self, headless=False):
        """Setup Chrome driver with download preferences"""
        chrome_options = Options()
        
        if headless:
            chrome_options.add_argument("--headless")
        
        chrome_options.add_argument("--no-sandbox")
        chrome_options.add_argument("--disable-dev-shm-usage")
        chrome_options.add_argument("--disable-gpu")
        chrome_options.add_argument("--window-size=1920,1080")
        
        # Set download preferences
        prefs = {
            "download.default_directory": self.download_folder,
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "safebrowsing.enabled": True,
            "safebrowsing.disable_download_protection": True,
            "profile.default_content_setting_values.notifications": 2
        }
        chrome_options.add_experimental_option("prefs", prefs)
        
        try:
            self.driver = webdriver.Chrome(options=chrome_options)
            print("Chrome driver initialized successfully")
        except Exception as e:
            print(f"Error initializing Chrome driver: {e}")
            print("Make sure you have Chrome and ChromeDriver installed")
            raise
    
    def handle_authentication(self):
        """Handle SharePoint authentication if needed"""
        try:
            # Check if we're on a login page
            login_indicators = [
                "Sign in",
                "Enter your password", 
                "Stay signed in",
                "Microsoft",
                "office.com",
                "login.microsoftonline.com"
            ]
            
            page_text = self.driver.page_source.lower()
            current_url = self.driver.current_url.lower()
            
            for indicator in login_indicators:
                if indicator.lower() in page_text or indicator.lower() in current_url:
                    print(f"🔐 Authentication required - detected: {indicator}")
                    print(f"Current URL: {self.driver.current_url}")
                    print("Please complete authentication manually in the browser...")
                    
                    # Wait for user to complete authentication
                    input("Press Enter after you've signed in and can see the video...")
                    return True
            
            return False
            
        except Exception as e:
            print(f"Error checking authentication: {e}")
            return False
    
    def maintain_session(self):
        """Maintain SharePoint session between downloads"""
        try:
            # Ensure browser stays focused
            self.ensure_browser_focus()
            
            # Simulate user activity
            self.simulate_user_activity()
            
            # Keep the same browser tab open
            # Don't navigate away completely
            time.sleep(2)
            
            # Clear any popups or overlays that might interfere
            try:
                # Close any notification bars
                close_buttons = self.driver.find_elements(By.CSS_SELECTOR, "button[aria-label*='Close'], button[aria-label*='Dismiss']")
                for btn in close_buttons:
                    try:
                        btn.click()
                    except:
                        pass
            except:
                pass
                
        except Exception as e:
            print(f"Session maintenance error: {e}")

    def ensure_browser_focus(self):
        """Ensure browser window is focused and active"""
        try:
            # Bring browser window to front
            self.driver.maximize_window()
            
            # Switch to the current window (ensures focus)
            self.driver.switch_to.window(self.driver.current_window_handle)
            
            # Click somewhere on the page to ensure it's active
            try:
                body = self.driver.find_element(By.TAG_NAME, "body")
                self.driver.execute_script("arguments[0].click();", body)
            except:
                pass
                
            # Small delay to ensure focus is established
            time.sleep(1)
            
            print("✅ Browser window focused")
            
        except Exception as e:
            print(f"Error focusing browser: {e}")

    def simulate_user_activity(self):
        """Simulate user activity to keep SharePoint active"""
        try:
            # Execute JavaScript to simulate user presence
            self.driver.execute_script("""
                // Simulate mouse movement
                var event = new MouseEvent('mousemove', {
                    view: window,
                    bubbles: true,
                    cancelable: true,
                    clientX: 100,
                    clientY: 100
                });
                document.dispatchEvent(event);
                
                // Simulate page visibility
                Object.defineProperty(document, 'hidden', {
                    value: false,
                    writable: true
                });
                
                // Trigger focus event
                window.focus();
                
                // Simulate scroll to activate page
                window.scrollTo(0, 100);
                window.scrollTo(0, 0);
            """)
            
            print("✅ Simulated user activity")
            
        except Exception as e:
            print(f"Error simulating activity: {e}")

    def download_video(self, url, row_number=None):
        """Download a single video from SharePoint URL"""
        try:
            print(f"\nProcessing {'Row ' + str(row_number) + ': ' if row_number else ''}{url}")
            
            # Navigate to the URL
            self.driver.get(url)
            print("Page loaded, waiting for content...")

            # Wait for page to load
            WebDriverWait(self.driver, 30).until(
                EC.presence_of_element_located((By.TAG_NAME, "body"))
            )

            # Ensure browser is focused and active
            self.ensure_browser_focus()
            
            # Wait additional time for SharePoint to fully load
            time.sleep(5)

            # Check if authentication is needed
            # if self.handle_authentication():
            #     # Wait a bit more after authentication
            #     time.sleep(5)

            # Wait for the actual video player interface to load
            # Wait for the actual video player interface to load
            print("⏳ Waiting for video player to load...")
            video_player_loaded = False

            # Simulate user activity to keep SharePoint active
            self.simulate_user_activity()

            # Multiple attempts to load video player
            for attempt in range(3):
                try:
                    print(f"Attempt {attempt + 1} to load video player...")
                    
                    # Ensure browser focus before each attempt
                    self.ensure_browser_focus()
                    
                    # Wait for video-specific elements to appear
                    WebDriverWait(self.driver, 15).until(
                        EC.any_of(
                            # Video player elements
                            EC.presence_of_element_located((By.CSS_SELECTOR, "video")),
                            EC.presence_of_element_located((By.CSS_SELECTOR, "[data-automationid='downloadButton']")),
                            EC.presence_of_element_located((By.CSS_SELECTOR, "button[aria-label*='Download']")),
                            EC.presence_of_element_located((By.CSS_SELECTOR, ".ms-Button[aria-label*='Download']")),
                            # SharePoint video player containers
                            EC.presence_of_element_located((By.CSS_SELECTOR, "[data-automationid='videoPlayer']")),
                            EC.presence_of_element_located((By.CSS_SELECTOR, ".od-VideoPlayer")),
                            EC.presence_of_element_located((By.CSS_SELECTOR, "[class*='videoPlayer']")),
                            EC.presence_of_element_located((By.CSS_SELECTOR, "[class*='VideoPlayer']"))
                        )
                    )
                    print("✅ Video player interface loaded")
                    video_player_loaded = True
                    break
                    
                except TimeoutException:
                    print(f"⚠️ Video player not loaded on attempt {attempt + 1}")
                    if attempt < 2:  # Don't refresh on last attempt
                        print("Refreshing page and trying again...")
                        self.driver.refresh()
                        time.sleep(5)
                    else:
                        print("⚠️ Video player not detected after all attempts")
                        # Debug: Check what's actually on the page
                        page_title = self.driver.title
                        current_url = self.driver.current_url
                        print(f"Page title: {page_title}")
                        print(f"Current URL: {current_url}")
                        
                        # Check if it's an error page or redirect
                        if "error" in page_title.lower() or "not found" in page_title.lower():
                            print("❌ Error page detected")
                            return False
                        elif "sign in" in page_title.lower() or "login" in current_url:
                            print("🔐 Authentication required")
                            if self.handle_authentication():
                                time.sleep(5)
                                video_player_loaded = True
                            else:
                                return False

            # Additional wait for video content to fully load
            if video_player_loaded:
                time.sleep(3)
            else:
                # Last attempt - wait a bit more and check for any SharePoint content
                print("⏳ Giving SharePoint more time to load...")
                time.sleep(10)
            
            # Try multiple selectors for the download button
            download_selectors = [
                # Most specific SharePoint selectors first
                "button[data-automationid='downloadButton']",
                "[data-automationid='downloadButton']",
                "button[aria-label*='Download']",
                "button[title*='Download']",
                ".ms-Button[aria-label*='Download']",
                ".od-Button[aria-label*='Download']",
                "button[aria-label='Download']",
                "button.ms-Button[aria-label*='Download']",
                # More general selectors
                "button[name='Download']",
                ".ms-CommandBar button[aria-label*='Download']",
                "[role='menuitem'][aria-label*='Download']",
                ".ms-ContextualMenu button[aria-label*='Download']",
                # Fallback selectors
                "button:contains('Download')",
                "a[href*='download']"
            ]
            
            download_button = None
            
            for selector in download_selectors:
                try:
                    if ":contains" in selector:
                        # Use XPath for text-based search
                        download_button = WebDriverWait(self.driver, 5).until(
                            EC.element_to_be_clickable((By.XPATH, f"//button[contains(text(), 'Download')]"))
                        )
                    else:
                        download_button = WebDriverWait(self.driver, 5).until(
                            EC.element_to_be_clickable((By.CSS_SELECTOR, selector))
                        )
                    
                    print(f"Found download button with selector: {selector}")
                    break
                except TimeoutException:
                    continue
            
            if not download_button:
                print("❌ Download button not found. Trying alternative methods...")
                
                # Debug: Show all buttons on the page
                try:
                    all_buttons = self.driver.find_elements(By.TAG_NAME, "button")
                    print(f"Found {len(all_buttons)} total buttons on page:")
                    for i, button in enumerate(all_buttons[:10]):  # Show first 10 buttons
                        btn_text = button.get_attribute("aria-label") or button.text or button.get_attribute("title") or "No text"
                        btn_class = button.get_attribute("class") or "No class"
                        print(f"  Button {i+1}: '{btn_text}' (class: {btn_class[:50]})")
                        
                    # Also check for any elements with "download" in them
                    print("\n🔍 Searching for any elements containing 'download':")
                    download_elements = self.driver.find_elements(By.XPATH, "//*[contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'download') or contains(translate(@aria-label, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'download')]")
                    print(f"Found {len(download_elements)} elements with 'download':")
                    for i, elem in enumerate(download_elements[:5]):
                        elem_text = elem.text or elem.get_attribute("aria-label") or elem.get_attribute("title") or "No text"
                        elem_tag = elem.tag_name
                        print(f"  Element {i+1}: <{elem_tag}> '{elem_text}'")

                except Exception as e:
                    print(f"Error listing buttons: {e}")
                
                # Try to find any button that might be the download button
                try:
                    # Look for buttons with download-related text
                    buttons = self.driver.find_elements(By.TAG_NAME, "button")
                    for button in buttons:
                        button_text = button.get_attribute("aria-label") or button.text or button.get_attribute("title") or ""
                        if "download" in button_text.lower():
                            download_button = button
                            print(f"Found download button with text: {button_text}")
                            break
                except Exception as e:
                    print(f"Error finding buttons: {e}")
                
                if not download_button:
                    print("❌ Could not find download button. Trying page refresh...")
                    # Try refreshing the page once
                    self.driver.refresh()
                    time.sleep(10)
                    
                    # Try one more time after refresh
                    for selector in download_selectors[:5]:  # Try top 5 selectors
                        try:
                            if ":contains" in selector:
                                download_button = WebDriverWait(self.driver, 5).until(
                                    EC.element_to_be_clickable((By.XPATH, f"//button[contains(text(), 'Download')]"))
                                )
                            else:
                                download_button = WebDriverWait(self.driver, 5).until(
                                    EC.element_to_be_clickable((By.CSS_SELECTOR, selector))
                                )
                            print(f"Found download button after refresh with: {selector}")
                            break
                        except TimeoutException:
                            continue
                    
                    if not download_button:
                        print("❌ Still could not find download button after refresh")
                        
                        # Offer manual intervention
                        print("\n🔧 Manual intervention option:")
                        print("The browser window is still open. You can:")
                        print("1. Manually navigate to the video")
                        print("2. Click download manually") 
                        print("3. Press Enter here when download starts")
                        
                        try:
                            user_input = input("Press Enter if you manually started the download, or 'skip' to skip this video: ").strip().lower()
                            if user_input == 'skip':
                                return False
                            else:
                                print("✅ Manual download initiated")
                                time.sleep(3)
                                return True
                        except KeyboardInterrupt:
                            return False
            
            # Click the download button
            try:
                # Scroll to button if needed
                self.driver.execute_script("arguments[0].scrollIntoView();", download_button)
                time.sleep(1)
                
                # Click using JavaScript to avoid interception
                self.driver.execute_script("arguments[0].click();", download_button)
                print("✅ Clicked download button")
                
            except Exception as e:
                print(f"Error clicking download button: {e}")
                return False
            
            # Handle potential dialog boxes
            self.handle_download_dialogs()
            
            # Wait for download to start (reduced time since you said it downloads quickly)
            print("⏳ Waiting for download to start...")
            # time.sleep(3)

            # Check if download actually started by looking for download indicators
            try:
                # Sometimes SharePoint shows a progress indicator
                WebDriverWait(self.driver, 5).until(
                    EC.any_of(
                        EC.presence_of_element_located((By.CSS_SELECTOR, "[aria-label*='progress']")),
                        EC.presence_of_element_located((By.CSS_SELECTOR, ".ms-ProgressIndicator")),
                        EC.presence_of_element_located((By.CSS_SELECTOR, "[role='progressbar']"))
                    )
                )
                print("✅ Download progress detected")
            except TimeoutException:
                print("ℹ️ No download progress indicator found (download may have completed quickly)")
            
            return True
            
        except Exception as e:
            print(f"❌ Error downloading video: {e}")
            return False
    
    def handle_download_dialogs(self):
        """Handle the download dialog boxes"""
        try:
            # First dialog - "The download is video only..."
            print("Looking for download confirmation dialog...")
            
            # Wait for dialog to appear first
            try:
                WebDriverWait(self.driver, 10).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, "[role='dialog'], .ms-Dialog, .od-Dialog"))
                )
                print("✅ Dialog appeared")
                
                # Wait a moment for dialog content to load
                time.sleep(2)
                
                # Log dialog content for debugging
                try:
                    dialog = self.driver.find_element(By.CSS_SELECTOR, "[role='dialog'], .ms-Dialog, .od-Dialog")
                    print(f"Dialog text: {dialog.text[:200]}")  # First 200 characters
                except:
                    print("Could not find dialog element for debugging")
                
                # Debug: List all buttons in the dialog
                try:
                    dialog_buttons = self.driver.find_elements(By.XPATH, "//div[@role='dialog']//button | //div[contains(@class, 'Dialog')]//button")
                    print(f"Found {len(dialog_buttons)} buttons in dialog:")
                    for i, btn in enumerate(dialog_buttons):
                        btn_text = btn.text or btn.get_attribute('aria-label') or 'No text'
                        btn_class = btn.get_attribute('class') or 'No class'
                        print(f"  Button {i+1}: '{btn_text}' (class: {btn_class})")
                except Exception as e:
                    print(f"Could not list dialog buttons: {e}")
                
            except TimeoutException:
                print("No dialog appeared")
                return
            
            # Wait specifically for the download video button to appear
            try:
                WebDriverWait(self.driver, 15).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, "button[aria-label='Download video']"))
                )
                print("✅ Found 'Download video' button in dialog")
            except TimeoutException:
                print("⚠️ 'Download video' button not found, trying alternative selectors...")
            
            # Look for download button specifically within the dialog
            dialog_selectors = [
                "button[aria-label='Download video']",
                "button.ms-Button--primary[aria-label='Download video']",
                "//button[@aria-label='Download video']",
                "//button[contains(@class, 'ms-Button--primary') and @aria-label='Download video']",
                "//button[contains(@class, 'ms-Button--primary')]//span[text()='Download']",
                "//button[@data-is-focusable='true' and @aria-label='Download video']",
                "//div[@role='dialog']//button[contains(text(), 'Download')]",
                "//div[contains(@class, 'Dialog')]//button[contains(text(), 'Download')]",
                "//div[contains(@class, 'ms-Dialog')]//button[contains(text(), 'Download')]",
                "//div[@role='dialog']//button[@data-automationid='primaryButton']",
                "//div[@role='dialog']//button[contains(@class, 'ms-Button--primary')]"
            ]
            
            for selector in dialog_selectors:
                try:
                    if selector.startswith("//"):
                        # XPath selector
                        dialog_button = WebDriverWait(self.driver, 10).until(
                            EC.element_to_be_clickable((By.XPATH, selector))
                        )
                    else:
                        # CSS selector
                        dialog_button = WebDriverWait(self.driver, 10).until(
                            EC.element_to_be_clickable((By.CSS_SELECTOR, selector))
                        )
                    
                    self.driver.execute_script("arguments[0].click();", dialog_button)
                    print("✅ Clicked download button in dialog")
                    time.sleep(2)
                    break
                    
                except TimeoutException:
                    continue
            
            # Second dialog - "Allow download" (if it appears)
            try:
                allow_selectors = [
                    "//div[@role='dialog']//button[contains(text(), 'Allow')]",
                    "//div[@role='dialog']//button[contains(text(), 'OK')]", 
                    "//div[@role='dialog']//button[contains(text(), 'Yes')]",
                    "//div[contains(@class, 'Dialog')]//button[contains(@class, 'primary')]",
                    "//button[@data-automationid='primaryButton' and contains(text(), 'Allow')]",
                    "//button[contains(@class, 'ms-Button--primary') and contains(text(), 'Allow')]"
                ]
                
                for selector in allow_selectors:
                    try:
                        allow_button = WebDriverWait(self.driver, 5).until(
                            EC.element_to_be_clickable((By.XPATH, selector))
                        )
                        
                        self.driver.execute_script("arguments[0].click();", allow_button)
                        print("✅ Clicked allow button in second dialog")
                        break
                        
                    except TimeoutException:
                        continue
                        
            except Exception as e:
                print("ℹ️ No second dialog appeared (this is normal)")
                
        except Exception as e:
            print(f"ℹ️ No dialogs appeared or error handling dialogs: {e}")
    
    def test_multiple_links(self, links_data):
        """Test if the issue is with multiple links"""
        print("🧪 Testing first 3 links with longer delays...")
        
        for i, link_info in enumerate(links_data[:3]):
            url = link_info.get('link', '')
            row = link_info.get('row', 'Unknown')
            
            print(f"\n🔗 Testing link {i+1}: Row {row}")
            success = self.download_video(url, row)
            
            if success:
                print("✅ Success!")
            else:
                print("❌ Failed!")
                
            # Long delay between tests
            if i < 2:
                print("⏳ Waiting 30 seconds before next test...")
                time.sleep(30)
    
    def download_from_file(self, filename="sharepoint_links.json"):
        """Download all videos from a JSON file containing links"""
        try:
            with open(filename, 'r') as f:
                links_data = json.load(f)
            
            print(f"Found {len(links_data)} links to process")
            
            successful_downloads = 0
            failed_downloads = 0
            
            for i, link_info in enumerate(links_data, 1):
                print(f"\n{'='*60}")
                print(f"Processing link {i}/{len(links_data)}")
                
                url = link_info.get('link', '')
                row = link_info.get('row', 'Unknown')
                
                if self.download_video(url, row):
                    successful_downloads += 1
                    print(f"✅ Successfully processed link from row {row}")
                else:
                    failed_downloads += 1
                    print(f"❌ Failed to process link from row {row}")
                
                # Maintain session between downloads
                self.maintain_session()
                
                # Wait between downloads (increased time to avoid rate limiting)
                if i < len(links_data):
                    print("⏳ Waiting 15 seconds before next download...")
                    time.sleep(15)
            
            print(f"\n{'='*60}")
            print(f"📊 SUMMARY:")
            print(f"✅ Successful downloads: {successful_downloads}")
            print(f"❌ Failed downloads: {failed_downloads}")
            print(f"📂 Downloads saved to: {self.download_folder}")
            
        except FileNotFoundError:
            print(f"❌ File '{filename}' not found!")
            print("Please run the Google Sheets extractor first to generate the links file.")
        except json.JSONDecodeError:
            print(f"❌ Error reading JSON file '{filename}'")
        except Exception as e:
            print(f"❌ Error: {e}")
    
    def download_from_text_file(self, filename="sharepoint_urls.txt"):
        """Download all videos from a text file containing URLs (one per line)"""
        try:
            with open(filename, 'r') as f:
                urls = [line.strip() for line in f if line.strip()]
            
            print(f"Found {len(urls)} URLs to process")
            
            successful_downloads = 0
            failed_downloads = 0
            
            for i, url in enumerate(urls, 1):
                print(f"\n{'='*60}")
                print(f"Processing URL {i}/{len(urls)}")
                
                if self.download_video(url):
                    successful_downloads += 1
                    print(f"✅ Successfully processed URL {i}")
                else:
                    failed_downloads += 1
                    print(f"❌ Failed to process URL {i}")
                
                # Maintain session between downloads
                self.maintain_session()
                
                # Wait between downloads (increased time to avoid rate limiting)
                if i < len(urls):
                    print("⏳ Waiting 15 seconds before next download...")
                    time.sleep(15)
            
            print(f"\n{'='*60}")
            print(f"📊 SUMMARY:")
            print(f"✅ Successful downloads: {successful_downloads}")
            print(f"❌ Failed downloads: {failed_downloads}")
            print(f"📂 Downloads saved to: {self.download_folder}")
            
        except FileNotFoundError:
            print(f"❌ File '{filename}' not found!")
        except Exception as e:
            print(f"❌ Error: {e}")
    
    def close(self):
        """Close the browser"""
        if hasattr(self, 'driver'):
            self.driver.quit()
            print("Browser closed")

def main():
    print("🚀 SharePoint Video Downloader")
    print("=" * 50)
    
    # Create downloader instance
    downloader = SimpleSharePointDownloader(
        download_folder="downloads",
        headless=False  # Set to True to run without showing browser
    )
    
    try:
        # Check which files exist
        json_file = "sharepoint_links.json"
        text_file = "sharepoint_urls.txt"
        
        if os.path.exists(json_file):
            print(f"📂 Found {json_file}, starting downloads...")
            downloader.download_from_file(json_file)
        elif os.path.exists(text_file):
            print(f"📂 Found {text_file}, starting downloads...")
            downloader.download_from_text_file(text_file)
        else:
            print("❌ No link files found!")
            print("Please run the Google Sheets extractor first to generate:")
            print("- sharepoint_links.json")
            print("- sharepoint_urls.txt")
            
    except KeyboardInterrupt:
        print("\n⏹️ Download interrupted by user")
    except Exception as e:
        print(f"❌ Unexpected error: {e}")
    finally:
        downloader.close()

if __name__ == "__main__":
    main()