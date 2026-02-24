#!/usr/bin/env python3
"""
Download all 4 counties' files from SharePoint sharing link
Uses Selenium to automate the download from SharePoint's web interface.

Requirements:
    pip install selenium requests
"""

import os
import time
import glob
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options


def setup_chrome_driver(download_dir):
    """Setup Chrome with automatic download settings"""
    chrome_options = Options()
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--window-size=1400,900")
    chrome_options.add_argument("--disable-blink-features=AutomationControlled")

    # Configure download settings
    prefs = {
        "download.default_directory": download_dir,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": False,  # Disable safe browsing to avoid download warnings
        "safebrowsing.disable_download_protection": True,
        "profile.default_content_settings.popups": 0,
        "profile.default_content_setting_values.automatic_downloads": 1,
    }
    chrome_options.add_experimental_option("prefs", prefs)
    chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
    chrome_options.add_experimental_option("useAutomationExtension", False)

    driver = webdriver.Chrome(options=chrome_options)

    # Enable downloads
    driver.execute_cdp_cmd("Page.setDownloadBehavior", {
        "behavior": "allow",
        "downloadPath": download_dir
    })

    return driver


def wait_for_download_complete(download_dir, timeout=60):
    """Wait for any download to complete in the download directory"""
    print("⏳ Waiting for download to complete...")

    end_time = time.time() + timeout
    while time.time() < end_time:
        # Check for .crdownload files (Chrome temp download files)
        downloading = glob.glob(os.path.join(download_dir, "*.crdownload"))
        if not downloading:
            # Check if any .png files were downloaded
            png_files = glob.glob(os.path.join(download_dir, "*.png"))
            if png_files:
                # Get the most recently modified file
                latest_file = max(png_files, key=os.path.getmtime)
                file_size = os.path.getsize(latest_file)
                if file_size > 0:
                    return latest_file
        time.sleep(0.5)

    return None


def _count_file_items(driver):
    """Count visible file items on the SharePoint page."""
    return driver.execute_script("""
        const items = document.querySelectorAll(
            '[data-automationid="row"], [role="row"], .ms-List-cell, '
            + '.od-ItemContent, [data-automationid="DetailsRow"]'
        );
        return items.length;
    """)


def _find_scrollable_container(driver):
    """Find the actual scrollable container that holds the SharePoint file list.

    Starts from a file row element and walks up the DOM tree to find the
    nearest ancestor with overflow scroll/auto.
    """
    container = driver.execute_script("""
        // First, find any file row element as an anchor point
        const row = document.querySelector(
            '[data-automationid="row"], [data-automationid="DetailsRow"], '
            + '[role="row"], .ms-List-cell, .od-ItemContent'
        );
        if (!row) return null;

        // Walk up from the row to find the scrollable ancestor
        let el = row.parentElement;
        while (el && el !== document.body && el !== document.documentElement) {
            const style = window.getComputedStyle(el);
            const overflowY = style.overflowY;
            // Check if this element is scrollable
            if ((overflowY === 'auto' || overflowY === 'scroll' || overflowY === 'overlay')
                && el.scrollHeight > el.clientHeight) {
                return el;
            }
            el = el.parentElement;
        }

        // Fallback: try known SharePoint selectors
        const selectors = [
            '[data-is-scrollable="true"]',
            '.ms-ScrollablePane--contentContainer',
            '[role="presentation"][style*="overflow"]',
        ];
        for (const sel of selectors) {
            const candidates = document.querySelectorAll(sel);
            for (const c of candidates) {
                if (c.scrollHeight > c.clientHeight + 50) {
                    return c;
                }
            }
        }

        return null;
    """)
    return container


def scroll_to_load_all(driver, max_scrolls=20, wait_time=2.0):
    """Scroll down the page incrementally to trigger lazy-loading of all files.

    SharePoint uses a virtualized list inside a deeply nested scrollable
    container. We find that exact container by walking up from a file row
    element, then scroll it directly via JavaScript.
    """
    print("📜 Scrolling to load all files...")

    # Find the scrollable container that holds the file list
    container = _find_scrollable_container(driver)

    if container:
        # Log container info for debugging
        info = driver.execute_script("""
            const el = arguments[0];
            return {
                tag: el.tagName,
                className: (el.className || '').substring(0, 100),
                scrollHeight: el.scrollHeight,
                clientHeight: el.clientHeight,
                scrollTop: el.scrollTop
            };
        """, container)
        print(f"   Found scrollable container: <{info['tag']}> "
              f"class='{info['className'][:50]}' "
              f"scrollHeight={info['scrollHeight']} clientHeight={info['clientHeight']}")
    else:
        print("   ⚠️ Could not find scrollable container, will try window scroll")

    consecutive_same = 0

    for i in range(max_scrolls):
        prev_count = _count_file_items(driver)

        if container:
            # Scroll the identified container directly
            driver.execute_script("""
                arguments[0].scrollTop += 600;
            """, container)
        else:
            # Fallback: scroll window
            driver.execute_script("window.scrollBy(0, 600);")

        time.sleep(wait_time)

        new_count = _count_file_items(driver)
        print(f"   Scroll {i+1}/{max_scrolls}: {prev_count} -> {new_count} items")

        if new_count == prev_count:
            consecutive_same += 1
            if consecutive_same >= 3:
                print(f"   ✅ All content loaded ({new_count} items)")
                break
        else:
            consecutive_same = 0

    # Scroll back to top
    if container:
        driver.execute_script("arguments[0].scrollTop = 0;", container)
    else:
        driver.execute_script("window.scrollTo(0, 0);")
    time.sleep(1)


def _switch_to_tiles_view(driver):
    """Switch SharePoint file list to Tiles view."""
    print("\n📜 Changing view to Tiles to load all files...")
    try:
        view_changed = driver.execute_script("""
            const buttons = document.querySelectorAll('button');
            for (const btn of buttons) {
                const text = (btn.textContent || '').trim();
                if (text === 'All Documents' && btn.offsetWidth > 0 && btn.offsetHeight > 0) {
                    btn.click();
                    return true;
                }
            }
            return false;
        """)

        if view_changed:
            print("   ✅ Clicked view selector")
            time.sleep(1)

            tiles_clicked = driver.execute_script("""
                const tilesSelectors = [
                    'button[aria-label*="Tiles"]',
                    'button[title*="Tiles"]',
                    'button[name="Tiles"]',
                    '[data-automationid="Tiles"]',
                    'li[role="menuitem"]',
                    'button.ms-ContextualMenu-link'
                ];
                for (const selector of tilesSelectors) {
                    const elements = document.querySelectorAll(selector);
                    for (const el of elements) {
                        const text = (el.textContent || '').trim();
                        const aria = el.getAttribute('aria-label') || '';
                        const title = el.getAttribute('title') || '';
                        if (text.toLowerCase().includes('tiles') ||
                            aria.toLowerCase().includes('tiles') ||
                            title.toLowerCase().includes('tiles')) {
                            if (el.offsetWidth > 0 && el.offsetHeight > 0) {
                                el.click();
                                return true;
                            }
                        }
                    }
                }
                return false;
            """)

            if tiles_clicked:
                print("   ✅ Changed to Tiles view")
                time.sleep(3)
            else:
                print("   ⚠️  Could not find Tiles option")
        else:
            print("   ⚠️  Could not find view selector button")
    except Exception as e:
        print(f"   ⚠️  View change failed: {e}")


def find_and_click_file(driver, target_filename):
    """Find a file by name on the current SharePoint page and click it.

    Uses two strategies: exact match then partial match.

    Returns:
        True if file was found and clicked, False otherwise.
    """
    print(f"\n🔎 Looking for file: {target_filename}")

    # Strategy 1: Exact match
    print("   Strategy 1: Trying JavaScript exact match search...")
    try:
        clicked = driver.execute_script("""
            const fileName = arguments[0];
            const matches = [];
            const allElements = document.querySelectorAll('*');

            for (const el of allElements) {
                const text = (el.textContent || '').trim();
                const aria = el.getAttribute('aria-label') || '';
                const title = el.getAttribute('title') || '';
                const name = el.getAttribute('name') || '';

                if (text === fileName || aria.includes(fileName) ||
                    title.includes(fileName) || name === fileName) {
                    const rect = el.getBoundingClientRect();
                    if (rect.width > 0 && rect.height > 0) {
                        matches.push({
                            element: el,
                            text: text.substring(0, 100),
                            width: rect.width,
                            height: rect.height
                        });
                    }
                }
            }

            if (matches.length > 0) {
                try {
                    matches[0].element.scrollIntoView({block: 'center'});
                    matches[0].element.click();
                    return {success: true, count: matches.length};
                } catch (e) {
                    return {success: false, count: matches.length, error: e.toString()};
                }
            }
            return null;
        """, target_filename)

        if clicked and clicked['success']:
            print(f"   ✅ File found and clicked (exact match, {clicked['count']} match(es))")
            return True
        elif clicked:
            print(f"   ⚠️ Found but click failed: {clicked.get('error', 'unknown')}")
    except Exception as e:
        print(f"   Exact match failed: {e}")

    # Strategy 2: Partial match
    print("   Strategy 2: Trying JavaScript partial match search...")
    try:
        clicked = driver.execute_script("""
            const fileName = arguments[0];
            const matches = [];
            const allElements = document.querySelectorAll('*');

            for (const el of allElements) {
                const text = (el.textContent || '').trim();
                if (text.includes(fileName)) {
                    const rect = el.getBoundingClientRect();
                    if (rect.width > 0 && rect.height > 0 && rect.width < 500) {
                        matches.push({
                            element: el,
                            text: text.substring(0, 100),
                            width: rect.width,
                            height: rect.height
                        });
                    }
                }
            }

            if (matches.length > 0) {
                try {
                    matches[0].element.scrollIntoView({block: 'center'});
                    matches[0].element.click();
                    return {success: true, count: matches.length};
                } catch (e) {
                    return {success: false, count: matches.length, error: e.toString()};
                }
            }
            return null;
        """, target_filename)

        if clicked and clicked['success']:
            print(f"   ✅ File found and clicked (partial match, {clicked['count']} match(es))")
            return True
        elif clicked:
            print(f"   ⚠️ Found but click failed: {clicked.get('error', 'unknown')}")
    except Exception as e:
        print(f"   Partial match failed: {e}")

    print(f"   ❌ Could not find '{target_filename}'")
    return False


def find_and_click_download(driver):
    """Find and click the Download button on the SharePoint toolbar.

    Returns:
        True if download was triggered, False otherwise.
    """
    print("\n🔎 Looking for Download button...")

    try:
        download_button = driver.execute_script("""
            const buttons = document.querySelectorAll('button');

            for (const btn of buttons) {
                const aria = btn.getAttribute('aria-label') || '';
                const title = btn.getAttribute('title') || '';
                const text = (btn.textContent || '').trim();

                if (aria.toLowerCase().includes('download') ||
                    title.toLowerCase().includes('download') ||
                    text.toLowerCase().includes('download')) {

                    const rect = btn.getBoundingClientRect();
                    if (rect.top < 200 && rect.left < 500 &&
                        rect.width > 0 && rect.height > 0) {
                        return btn;
                    }
                }
            }

            // Fallback: any visible download button
            for (const btn of buttons) {
                const aria = btn.getAttribute('aria-label') || '';
                const title = btn.getAttribute('title') || '';
                if ((aria.toLowerCase().includes('download') ||
                     title.toLowerCase().includes('download')) &&
                    btn.offsetWidth > 0 && btn.offsetHeight > 0) {
                    return btn;
                }
            }
            return null;
        """)

        if not download_button:
            print("   ❌ Could not find Download button")
            return False

        print("   ✅ Found Download button")

        from selenium.webdriver.common.action_chains import ActionChains
        actions = ActionChains(driver)

        try:
            actions.move_to_element(download_button).perform()
            time.sleep(0.5)
            download_button.click()
            print("   ✅ Download button clicked!")
        except Exception:
            driver.execute_script("arguments[0].click();", download_button)
            print("   ✅ Download button clicked (via JavaScript)!")

        time.sleep(2)

        # Handle potential download menu
        try:
            menu_clicked = driver.execute_script("""
                const menus = document.querySelectorAll('[role="menu"], [role="menubar"], .ms-ContextualMenu');
                for (const menu of menus) {
                    if (menu.offsetWidth > 0 && menu.offsetHeight > 0) {
                        const menuItems = menu.querySelectorAll('[role="menuitem"], button, a');
                        for (const item of menuItems) {
                            const aria = (item.getAttribute('aria-label') || '').toLowerCase();
                            if (aria.includes('download this file') &&
                                item.offsetWidth > 0 && item.offsetHeight > 0) {
                                item.click();
                                return true;
                            }
                        }
                        for (const item of menuItems) {
                            const text = (item.textContent || '').trim().toLowerCase();
                            const aria2 = (item.getAttribute('aria-label') || '').toLowerCase();
                            if ((text.includes('download') || aria2.includes('download')) &&
                                item.offsetWidth > 0 && item.offsetHeight > 0) {
                                item.click();
                                return true;
                            }
                        }
                    }
                }
                return false;
            """)
            if menu_clicked:
                print("   ✅ Clicked download menu item")
                time.sleep(1)
        except Exception:
            pass

        return True

    except Exception as e:
        print(f"   ❌ Error: {e}")
        return False


def download_all_files_from_sharepoint(sharepoint_url, target_filenames, download_dir):
    """Download multiple files from a SharePoint sharing link in a single browser session.

    Opens one browser, scrolls to load all lazy-loaded content, then downloads
    each file sequentially.

    Args:
        sharepoint_url: SharePoint sharing URL
        target_filenames: List of filenames to download
        download_dir: Directory to save downloads

    Returns:
        List of paths to successfully downloaded files
    """
    os.makedirs(download_dir, exist_ok=True)
    downloaded_files = []
    driver = None

    try:
        print("🚀 Starting Chrome browser (single session for all files)...")
        driver = setup_chrome_driver(download_dir)

        print(f"📂 Opening SharePoint link...")
        driver.get(sharepoint_url)

        print("⏳ Waiting for SharePoint to load...")
        time.sleep(8)

        # Check for login
        try:
            page_title = driver.title
            print(f"   Page title: {page_title}")
            if "sign in" in page_title.lower() or "login" in page_title.lower():
                print("\n⚠️  SharePoint login required!")
                print("   Please log in to SharePoint in the browser window.")
                print("   You have 30 seconds to complete the login...")
                time.sleep(30)
        except Exception as e:
            print(f"   Could not read page title: {e}")

        # Change view to Tiles
        _switch_to_tiles_view(driver)

        # Scroll to load all lazy-loaded files
        scroll_to_load_all(driver)

        # Download each file
        for i, filename in enumerate(target_filenames):
            print(f"\n{'='*60}")
            print(f"📥 [{i+1}/{len(target_filenames)}] Downloading: {filename}")
            print(f"{'='*60}")

            # Find and click the file
            if not find_and_click_file(driver, filename):
                print(f"❌ Could not find '{filename}' — skipping")
                continue

            time.sleep(2)

            # Click download
            if not find_and_click_download(driver):
                print(f"❌ Could not trigger download for '{filename}' — skipping")
                try:
                    driver.execute_script("document.body.click();")
                except Exception:
                    pass
                time.sleep(1)
                continue

            # Wait for download
            downloaded_file = wait_for_download_complete(download_dir, timeout=60)

            if downloaded_file:
                print(f"✅ Downloaded: {os.path.basename(downloaded_file)}")
                downloaded_files.append(downloaded_file)
            else:
                print(f"⚠️ Download may not have completed for '{filename}'")

            # Navigate back to file list for next file
            if i < len(target_filenames) - 1:
                print("\n🔙 Returning to file list...")
                try:
                    driver.back()
                    time.sleep(3)

                    # Verify we're back on the file list
                    item_count = driver.execute_script("""
                        const items = document.querySelectorAll(
                            '[data-automationid="row"], [role="row"], .ms-List-cell, .od-ItemContent'
                        );
                        return items.length;
                    """)

                    if item_count == 0:
                        print("   ⚠️ Back navigation failed, re-navigating...")
                        driver.get(sharepoint_url)
                        time.sleep(8)
                        _switch_to_tiles_view(driver)
                        scroll_to_load_all(driver)
                    else:
                        print(f"   ✅ Back on file list ({item_count} items)")

                except Exception as e:
                    print(f"   ⚠️ Navigation error: {e}, re-navigating...")
                    driver.get(sharepoint_url)
                    time.sleep(8)
                    _switch_to_tiles_view(driver)
                    scroll_to_load_all(driver)

        return downloaded_files

    except Exception as e:
        print(f"\n❌ Error: {e}")
        import traceback
        traceback.print_exc()
        return downloaded_files

    finally:
        if driver:
            try:
                print("\n🔒 Closing browser...")
                driver.quit()
            except Exception:
                pass


def download_file_from_sharepoint(sharepoint_url, target_filename, download_dir):
    """
    Download a specific file from SharePoint sharing link

    Args:
        sharepoint_url: SharePoint sharing URL
        target_filename: Name of file to download (e.g., "Alameda.png")
        download_dir: Directory to save downloads

    Returns:
        Path to downloaded file or None
    """
    os.makedirs(download_dir, exist_ok=True)

    driver = None
    try:
        print("🚀 Starting Chrome browser...")
        driver = setup_chrome_driver(download_dir)

        print(f"📂 Opening SharePoint link...")
        driver.get(sharepoint_url)

        # Wait for page to load
        print("⏳ Waiting for SharePoint to load...")
        time.sleep(8)

        # Debug: Check what we see on the page
        print("\n🔍 Checking page content...")
        try:
            page_title = driver.title
            print(f"   Page title: {page_title}")

            # Check if login is required
            if "sign in" in page_title.lower() or "login" in page_title.lower():
                print("\n⚠️  SharePoint login required!")
                print("   Please log in to SharePoint in the browser window.")
                print("   You have 10 seconds to complete the login...")
                time.sleep(10)
        except Exception as e:
            print(f"   Could not read page title: {e}")

        # Change view to "Tiles" to load all files at once
        print("\n📜 Changing view to Tiles to load all files...")
        try:
            # Find and click the "All Documents" button
            view_changed = driver.execute_script("""
                // Look for button with text "All Documents"
                const buttons = document.querySelectorAll('button');

                for (const btn of buttons) {
                    const text = (btn.textContent || '').trim();
                    if (text === 'All Documents' && btn.offsetWidth > 0 && btn.offsetHeight > 0) {
                        console.log('Found All Documents button:', btn);
                        btn.click();
                        return true;
                    }
                }

                return false;
            """)

            if view_changed:
                print("   ✅ Clicked view selector")
                time.sleep(1)

                # Now find and click "Tiles" option
                tiles_clicked = driver.execute_script("""
                    // Look for "Tiles" option in the dropdown
                    const tilesSelectors = [
                        'button[aria-label*="Tiles"]',
                        'button[title*="Tiles"]',
                        'button[name="Tiles"]',
                        '[data-automationid="Tiles"]',
                        'li[role="menuitem"]',
                        'button.ms-ContextualMenu-link'
                    ];

                    for (const selector of tilesSelectors) {
                        const elements = document.querySelectorAll(selector);
                        for (const el of elements) {
                            const text = (el.textContent || '').trim();
                            const aria = el.getAttribute('aria-label') || '';
                            const title = el.getAttribute('title') || '';

                            if (text.toLowerCase().includes('tiles') ||
                                aria.toLowerCase().includes('tiles') ||
                                title.toLowerCase().includes('tiles')) {
                                if (el.offsetWidth > 0 && el.offsetHeight > 0) {
                                    console.log('Found Tiles option:', el);
                                    el.click();
                                    return true;
                                }
                            }
                        }
                    }

                    return false;
                """)

                if tiles_clicked:
                    print("   ✅ Changed to Tiles view")
                    time.sleep(3)  # Wait for view to load

                    # After view change, we need to search again since elements are recreated
                    # So we'll set file_found = False and file_element = None to trigger new search
                else:
                    print("   ⚠️  Could not find Tiles option")

            else:
                print("   ⚠️  Could not find view selector button")

        except Exception as e:
            print(f"   ⚠️  View change failed: {e}")
            import traceback
            traceback.print_exc()

        # Try to find the file in various ways
        print(f"\n🔎 Looking for file: {target_filename}")

        file_found = False

        # Strategy 1: Use JavaScript exact match search and click directly
        if not file_found:
            print("   Strategy 1: Trying JavaScript exact match search...")
            try:
                clicked = driver.execute_script("""
                    const fileName = arguments[0];
                    const matches = [];

                    // Search all elements
                    const allElements = document.querySelectorAll('*');

                    for (const el of allElements) {
                        const text = (el.textContent || '').trim();
                        const aria = el.getAttribute('aria-label') || '';
                        const title = el.getAttribute('title') || '';
                        const name = el.getAttribute('name') || '';

                        if (text === fileName || aria.includes(fileName) ||
                            title.includes(fileName) || name === fileName) {
                            const rect = el.getBoundingClientRect();
                            if (rect.width > 0 && rect.height > 0) {
                                matches.push({
                                    element: el,
                                    text: text.substring(0, 100),
                                    aria: aria.substring(0, 100),
                                    title: title.substring(0, 100),
                                    name: name.substring(0, 100),
                                    width: rect.width,
                                    height: rect.height
                                });
                            }
                        }
                    }

                    if (matches.length > 0) {
                        // Try to click the first match
                        try {
                            matches[0].element.scrollIntoView({block: 'center'});
                            matches[0].element.click();
                            return {success: true, matches: matches};
                        } catch (e) {
                            return {success: false, matches: matches, error: e.toString()};
                        }
                    }

                    return null;
                """, target_filename)

                if clicked:
                    print(f"   📊 Found {len(clicked['matches'])} exact match(es)")

                    if clicked['success']:
                        file_found = True
                        print(f"   ✅ File clicked successfully!")
                    else:
                        print(f"   ⚠️ Found but failed to click: {clicked.get('error', 'unknown error')}")
            except Exception as e:
                print(f"   JavaScript exact match failed: {e}")

        # Strategy 2: Fallback partial match
        if not file_found:
            print("   Strategy 2: Trying JavaScript partial match search...")
            try:
                clicked = driver.execute_script("""
                    const fileName = arguments[0];
                    const matches = [];

                    // Search all elements
                    const allElements = document.querySelectorAll('*');

                    for (const el of allElements) {
                        const text = (el.textContent || '').trim();
                        if (text.includes(fileName)) {
                            const rect = el.getBoundingClientRect();
                            if (rect.width > 0 && rect.height > 0 && rect.width < 500) {
                                matches.push({
                                    element: el,
                                    text: text.substring(0, 100),
                                    width: rect.width,
                                    height: rect.height
                                });
                            }
                        }
                    }

                    if (matches.length > 0) {
                        // Try to click the first match
                        try {
                            matches[0].element.scrollIntoView({block: 'center'});
                            matches[0].element.click();
                            return {success: true, matches: matches};
                        } catch (e) {
                            return {success: false, matches: matches, error: e.toString()};
                        }
                    }

                    return null;
                """, target_filename)

                if clicked:
                    print(f"   📊 Found {len(clicked['matches'])} partial match(es)")

                    if clicked['success']:
                        file_found = True
                        print(f"   ✅ File clicked successfully!")
                    else:
                        print(f"   ⚠️ Found but failed to click: {clicked.get('error', 'unknown error')}")
            except Exception as e:
                print(f"   JavaScript partial match failed: {e}")

        if not file_found:
            print(f"\n❌ Could not locate '{target_filename}' on the page")
            print("\n💡 The browser window will stay open for 60 seconds.")
            print("   You can manually:")
            print("   1. Navigate to the file")
            print("   2. Click on it to select it")
            print("   3. Click the Download button")
            time.sleep(60)

            # Check if user downloaded manually
            downloaded_file = wait_for_download_complete(download_dir, timeout=5)
            if downloaded_file:
                return downloaded_file

            return None

        # File was already clicked in the JavaScript search
        print(f"\n✅ File '{target_filename}' clicked successfully")
        time.sleep(2)

        # Now look for the Download button in the upper left area
        print("\n🔎 Looking for Download button in upper left area...")

        download_clicked = False

        # Try to find and hover over the download button
        try:
            download_button = driver.execute_script("""
                // Look for Download button with tooltip "Download this file to your device"
                const buttons = document.querySelectorAll('button');

                for (const btn of buttons) {
                    const aria = btn.getAttribute('aria-label') || '';
                    const title = btn.getAttribute('title') || '';
                    const text = (btn.textContent || '').trim();

                    // Check if it's the download button
                    if (aria.toLowerCase().includes('download') ||
                        title.toLowerCase().includes('download') ||
                        text.toLowerCase().includes('download')) {

                        const rect = btn.getBoundingClientRect();

                        // Check if it's in the upper left area (top 200px, left 500px)
                        if (rect.top < 200 && rect.left < 500 &&
                            rect.width > 0 && rect.height > 0) {
                            console.log('Found download button:', btn);
                            console.log('Position:', rect.top, rect.left);
                            console.log('aria-label:', aria);
                            console.log('title:', title);
                            return btn;
                        }
                    }
                }

                // Fallback: any visible download button
                for (const btn of buttons) {
                    const aria = btn.getAttribute('aria-label') || '';
                    const title = btn.getAttribute('title') || '';

                    if ((aria.toLowerCase().includes('download') ||
                         title.toLowerCase().includes('download')) &&
                        btn.offsetWidth > 0 && btn.offsetHeight > 0) {
                        console.log('Found download button (fallback):', btn);
                        return btn;
                    }
                }

                return null;
            """)

            if download_button:
                print("   ✅ Found Download button")

                # Hover over the button to show tooltip
                from selenium.webdriver.common.action_chains import ActionChains
                actions = ActionChains(driver)

                try:
                    print("   🖱️  Hovering over Download button...")
                    actions.move_to_element(download_button).perform()
                    time.sleep(1)  # Wait for tooltip to appear

                    print("   🖱️  Clicking Download button...")
                    download_button.click()
                    download_clicked = True
                    print("   ✅ Download button clicked!")
                    time.sleep(3)  # Wait for download to start

                    # Check if any confirmation dialog appeared
                    try:
                        alert_text = driver.execute_script("""
                            const dialogs = document.querySelectorAll('[role="dialog"], [role="alertdialog"]');
                            if (dialogs.length > 0) {
                                return dialogs[0].textContent.substring(0, 200);
                            }
                            return null;
                        """)
                        if alert_text:
                            print(f"   📋 Dialog detected: {alert_text}")
                    except Exception:
                        pass

                except Exception as e:
                    print(f"   ⚠️ Direct click failed: {e}")
                    # Try JavaScript click
                    driver.execute_script("arguments[0].click();", download_button)
                    download_clicked = True
                    print("   ✅ Download button clicked (via JavaScript)!")
                    time.sleep(3)

                    # Check if a download menu appeared
                    try:
                        menu_items = driver.execute_script("""
                            const menus = document.querySelectorAll('[role="menu"], [role="menubar"], .ms-ContextualMenu');
                            if (menus.length > 0) {
                                const items = [];
                                for (const menu of menus) {
                                    if (menu.offsetWidth > 0 && menu.offsetHeight > 0) {
                                        const menuItems = menu.querySelectorAll('[role="menuitem"], button, a');
                                        for (const item of menuItems) {
                                            if (item.offsetWidth > 0 && item.offsetHeight > 0) {
                                                items.push({
                                                    text: (item.textContent || '').trim().substring(0, 100),
                                                    aria: (item.getAttribute('aria-label') || '').substring(0, 100),
                                                    title: (item.getAttribute('title') || '').substring(0, 100)
                                                });
                                            }
                                        }
                                    }
                                }
                                return items.length > 0 ? items : null;
                            }
                            return null;
                        """)

                        if menu_items:
                            print(f"\n   📋 Download menu detected with {len(menu_items)} items")

                            # Try to click the specific download menu item
                            # Priority 1: "Download this file to your device" (aria-label)
                            # Priority 2: Any item with "download" in text/aria
                            clicked_menu = driver.execute_script("""
                                const menus = document.querySelectorAll('[role="menu"], [role="menubar"], .ms-ContextualMenu');

                                // First pass: Look for "Download this file to your device"
                                for (const menu of menus) {
                                    if (menu.offsetWidth > 0 && menu.offsetHeight > 0) {
                                        const menuItems = menu.querySelectorAll('[role="menuitem"], button, a');
                                        for (const item of menuItems) {
                                            const aria = (item.getAttribute('aria-label') || '').toLowerCase();
                                            if (aria.includes('download this file') &&
                                                item.offsetWidth > 0 && item.offsetHeight > 0) {
                                                console.log('Found specific download item:', aria);
                                                item.click();
                                                return true;
                                            }
                                        }
                                    }
                                }

                                // Second pass: Any download button/link
                                for (const menu of menus) {
                                    if (menu.offsetWidth > 0 && menu.offsetHeight > 0) {
                                        const menuItems = menu.querySelectorAll('button, a');
                                        for (const item of menuItems) {
                                            const text = (item.textContent || '').trim().toLowerCase();
                                            const aria = (item.getAttribute('aria-label') || '').toLowerCase();
                                            const dataId = (item.getAttribute('data-id') || '').toLowerCase();

                                            if ((text.includes('download') || aria.includes('download') || dataId.includes('download')) &&
                                                item.offsetWidth > 0 && item.offsetHeight > 0) {
                                                console.log('Found download item (fallback):', text, aria, dataId);
                                                item.click();
                                                return true;
                                            }
                                        }
                                    }
                                }
                                return false;
                            """)

                            if clicked_menu:
                                print("   ✅ Clicked download menu item!")
                                time.sleep(2)
                    except Exception as menu_err:
                        print(f"   ⚠️ Menu check failed: {menu_err}")
            else:
                print("   ❌ Could not find Download button")

        except Exception as e:
            print(f"   ❌ Error finding download button: {e}")

        if not download_clicked:
            print("\n⚠️  Could not find or click Download button")
            print("   Trying to trigger download via keyboard shortcut...")

            # Try using keyboard shortcut (Ctrl+S or Cmd+S)
            try:
                from selenium.webdriver.common.keys import Keys
                from selenium.webdriver.common.action_chains import ActionChains

                actions = ActionChains(driver)
                # Use Ctrl on Windows/Linux, Cmd on Mac
                if os.uname().sysname == 'Darwin':
                    actions.key_down(Keys.COMMAND).send_keys('s').key_up(Keys.COMMAND).perform()
                else:
                    actions.key_down(Keys.CONTROL).send_keys('s').key_up(Keys.CONTROL).perform()

                print("   ⌨️  Sent save keyboard shortcut")
                time.sleep(2)
            except Exception as e:
                print(f"   ❌ Keyboard shortcut failed: {e}")

        # Wait for download to complete
        downloaded_file = wait_for_download_complete(download_dir, timeout=60)

        if downloaded_file:
            file_size = os.path.getsize(downloaded_file)
            print(f"\n✅ Download successful!")
            print(f"   File: {downloaded_file}")
            print(f"   Size: {file_size:,} bytes")
            return downloaded_file
        else:
            print("\n❌ Download did not complete or file not found")
            print("\n💡 Keeping browser open for 30 seconds for manual intervention...")
            time.sleep(30)

            # Check one more time
            downloaded_file = wait_for_download_complete(download_dir, timeout=2)
            if downloaded_file:
                print(f"✅ Found downloaded file: {downloaded_file}")
                return downloaded_file

            return None

    except Exception as e:
        print(f"\n❌ Error: {e}")
        import traceback
        traceback.print_exc()
        return None

    finally:
        if driver:
            try:
                print("\n🔒 Closing browser...")
                driver.quit()
            except Exception:
                pass


def main():
    SHAREPOINT_URL = "https://carorg.sharepoint.com/:f:/s/CAR-RE-PublicProducts/ElrCKkQh6_ZMpe5RIZgOohoB33WDC9L1NlkigRWlqWwvGg?e=Vi1XJZ"
    TARGET_FILES = [
        "Santa Clara.png",
        "San Mateo.png",
        "Alameda.png",
        "San Francisco.png"
    ]
    DOWNLOAD_DIR = os.path.join(os.getcwd(), "downloads")

    print("=" * 70)
    print("🚀 SharePoint File Downloader - All Counties (Batch Mode)")
    print("=" * 70)
    print(f"Target files: {len(TARGET_FILES)} counties")
    print(f"Download directory: {DOWNLOAD_DIR}")
    print("=" * 70)

    downloaded_files = download_all_files_from_sharepoint(
        SHAREPOINT_URL,
        TARGET_FILES,
        DOWNLOAD_DIR
    )

    print("\n" + "=" * 70)
    print("📊 DOWNLOAD SUMMARY")
    print("=" * 70)
    print(f"Target: {len(TARGET_FILES)} files")
    print(f"Success: {len(downloaded_files)} files")

    if downloaded_files:
        print("\n✅ Downloaded files:")
        for f in downloaded_files:
            size = os.path.getsize(f)
            print(f"   • {os.path.basename(f)} ({size:,} bytes)")

    if len(downloaded_files) == len(TARGET_FILES):
        print("\n🎉 ALL FILES DOWNLOADED SUCCESSFULLY!")
    elif downloaded_files:
        print("\n⚠️  Some files failed to download")
    else:
        print("\n❌ No files were downloaded")
    print("=" * 70)


if __name__ == "__main__":
    main()
