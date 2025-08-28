"""
Final Working County Downloader
Successfully downloads county reports from CAR.org Power BI table.

Usage: python final_working_downloader.py
"""

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
import time
import requests
import os
import re
from datetime import datetime

def setup_work_directory():
    """Setup work directory from environment or create new one"""
    try:
        # Try to get work directory from environment (set by workflow_startup.py)
        work_dir = os.getenv('SALA_WORK_DIR')
        
        if not work_dir:
            # Create new directory based on current month
            current_month = datetime.now().strftime("%B %Y")
            downloads_path = os.path.expanduser("~/Downloads")
            work_dir = os.path.join(downloads_path, f"SALA Report {current_month}")
        
        # Create directory if it doesn't exist
        if not os.path.exists(work_dir):
            os.makedirs(work_dir)
            print(f"📁 Created work directory: {work_dir}")
        else:
            print(f"📁 Using work directory: {work_dir}")
        
        return work_dir
        
    except Exception as e:
        print(f"❌ Error setting up work directory: {e}")
        return os.getcwd()

def setup_driver():
    """Setup Chrome driver with minimal options"""
    chrome_options = Options()
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    return webdriver.Chrome(options=chrome_options)

def get_sharepoint_links_from_powerbi(driver):
    """Extract both county and city SharePoint links from current Power BI view"""
    # Find and switch to Power BI iframe
    iframes = driver.find_elements(By.TAG_NAME, "iframe")
    for iframe in iframes:
        src = iframe.get_attribute("src")
        if src and "app.powerbi.com" in src:
            driver.switch_to.frame(iframe)
            break
    else:
        return {"county": [], "city": []}
    
    # Extract SharePoint image links from both county and city columns
    js_extract = """
    const links = document.querySelectorAll('a[href*="carorg.sharepoint.com/:i:"]');
    const countyLinks = [];
    const cityLinks = [];
    
    links.forEach(link => {
        const cell = link.closest('[role="gridcell"]');
        const columnIndex = cell ? cell.getAttribute('column-index') : null;
        const ariaColIndex = cell ? cell.getAttribute('aria-colindex') : null;
        
        // County links: column-index="2" aria-colindex="4"
        if (columnIndex === '2' && ariaColIndex === '4') {
            countyLinks.push(link.href);
        }
        // City links: column-index="3" aria-colindex="5"  
        else if (columnIndex === '3' && ariaColIndex === '5') {
            cityLinks.push(link.href);
        }
    });
    
    return {
        county: countyLinks,
        city: cityLinks
    };
    """
    
    try:
        result = driver.execute_script(js_extract)
        driver.switch_to.default_content()
        return result
    except Exception:
        driver.switch_to.default_content()
        return {"county": [], "city": []}

def download_image_file(sharing_url, filename, work_dir):
    """Download image file using proven SharePoint method"""
    headers = {
        'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/139.0.0.0 Safari/537.36'
    }
    
    try:
        # Get SharePoint viewer page
        response = requests.get(sharing_url, headers=headers, allow_redirects=True)
        if response.status_code != 200:
            return False
        
        # Extract download URL using proven regex pattern
        html = response.text
        pattern = r'downloadUrl["\']?\s*[:=]\s*["\']([^"\']*)["\']'
        matches = re.findall(pattern, html, re.IGNORECASE)
        
        if not matches:
            return False
        
        # Clean up the URL (handle Unicode escapes)
        download_url = matches[0].replace('\\/', '/').replace('\\u0026', '&')
        try:
            download_url = download_url.encode().decode('unicode_escape')
        except:
            pass
        
        # Download the file
        download_response = requests.get(download_url, headers=headers, stream=True)
        download_response.raise_for_status()
        
        # Save the file to work directory
        filepath = os.path.join(work_dir, filename)
        with open(filepath, 'wb') as f:
            for chunk in download_response.iter_content(chunk_size=8192):
                if chunk:
                    f.write(chunk)
        
        file_size = os.path.getsize(filepath)
        print(f"✅ Downloaded: {filepath} ({file_size:,} bytes)")
        return filepath
        
    except Exception as e:
        print(f"❌ Download failed for {filename}: {e}")
        return False

def main():
    """Main execution function"""
    print("🚀 Final Working County & City Downloader")
    print("="*50)
    print("Downloads both county and city level images")
    print("Format: County 1.png, City 1(1).png = County 1, City 1")
    print("="*50)
    
    # Setup work directory
    work_directory = setup_work_directory()
    
    # Target counties with their order numbers
    counties = [
        ("Santa Clara", 1),
        ("San Mateo", 2),
        ("Alameda", 3),
        ("San Francisco", 4)
    ]
    
    driver = setup_driver()
    downloaded_files = []
    
    try:
        # Open Power BI page and wait for loading
        print("📍 Opening Power BI page...")
        driver.get("https://www.car.org/marketdata/interactive/buyersguide")
        time.sleep(3)  # Wait for Power BI to load
        print("✅ Page loaded")
        
        # Process each county
        for county_name, county_order in counties:
            print(f"\n{'='*50}")
            print(f"🏛️  COUNTY {county_order}/{len(counties)}: {county_name}")
            print(f"{'='*50}")
            
            # Manual county selection
            print(f"👆 Please select '{county_name}' from the county dropdown in the browser")
            print(f"   Then press Enter when ready...")
            
            try:
                input()
            except:
                pass
            
            # Extract SharePoint links from current view
            links_data = get_sharepoint_links_from_powerbi(driver)
            
            print(f"📊 Found {len(links_data['county'])} county links and {len(links_data['city'])} city links")
            
            # Download county level images (if any)
            if links_data['county']:
                print(f"📥 Downloading county level image...")
                county_filename = f"{county_order}.png"
                if download_image_file(links_data['county'][0], county_filename, work_directory):
                    downloaded_files.append(county_filename)
            else:
                print(f"❌ No county level link found")
            
            # Download city level images (if any) 
            if links_data['city']:
                print(f"📥 Downloading {len(links_data['city'])} city level images...")
                
                # Sort city links alphabetically by getting the city names
                # For now, download in the order they appear (which should be alphabetical)
                for city_index, city_url in enumerate(links_data['city'], 1):
                    city_filename = f"{county_order}({city_index}).png"
                    if download_image_file(city_url, city_filename, work_directory):
                        downloaded_files.append(city_filename)
                        print(f"   ✅ City {city_index} downloaded")
                    else:
                        print(f"   ❌ City {city_index} failed")
            else:
                print(f"❌ No city level links found")
            
            print(f"✅ Completed {county_name}")
        
        # Summary
        print(f"\n{'='*60}")
        print(f"📊 FINAL SUMMARY")
        print(f"{'='*60}")
        print(f"Target counties: {len(counties)}")
        print(f"Successfully downloaded: {len(downloaded_files)} files")
        
        if downloaded_files:
            print(f"\n✅ Downloaded files:")
            # Sort files for better display
            county_files = [f for f in downloaded_files if '(' not in f]
            city_files = [f for f in downloaded_files if '(' in f]
            
            if county_files:
                print(f"  🏛️  County level files:")
                for filename in sorted(county_files):
                    if os.path.exists(filename):
                        file_size = os.path.getsize(filename)
                        print(f"    📁 {filename} ({file_size:,} bytes)")
            
            if city_files:
                print(f"  🏙️  City level files:")
                for filename in sorted(city_files):
                    if os.path.exists(filename):
                        file_size = os.path.getsize(filename)
                        print(f"    📁 {filename} ({file_size:,} bytes)")
        
        print(f"\n🎉 All done!")
        
    except Exception as e:
        print(f"❌ Error: {e}")
    finally:
        print(f"\nClosing browser...")
        driver.quit()

if __name__ == "__main__":
    main()