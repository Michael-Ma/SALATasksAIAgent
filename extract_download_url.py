import requests
import re
import json
from urllib.parse import unquote

def extract_download_url_from_sharepoint(sharing_url):
    """Extract the actual download URL from SharePoint viewer page"""
    
    headers = {
        'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/139.0.0.0 Safari/537.36',
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
        'Accept-Language': 'en-US,en;q=0.5',
        'DNT': '1',
    }
    
    session = requests.Session()
    session.headers.update(headers)
    
    try:
        # Get the viewer page
        response = session.get(sharing_url, allow_redirects=True)
        html = response.text
        
        print(f"✓ Got viewer page: {response.url}")
        
        # Method 1: Look for downloadUrl in JavaScript/JSON data
        download_patterns = [
            r'"downloadUrl":"([^"]*)"',
            r'"downloadUrlNoAuth":"([^"]*)"', 
            r'downloadUrl["\']?\s*[:=]\s*["\']([^"\']*)["\']',
            r'window\.SPClientContext\s*=\s*({[^}]*downloadUrl[^}]*})',
        ]
        
        for pattern in download_patterns:
            matches = re.findall(pattern, html, re.IGNORECASE)
            if matches:
                print(f"✓ Found download URL pattern: {pattern}")
                for match in matches:
                    # Clean up the URL (handle escaped characters)
                    clean_url = match.replace('\\/', '/').replace('\\u0026', '&')
                    # Handle Unicode escapes like \u002f
                    try:
                        clean_url = clean_url.encode().decode('unicode_escape')
                    except:
                        pass
                    if clean_url.startswith('http'):
                        print(f"📥 Download URL: {clean_url}")
                        return clean_url
        
        # Method 2: Look for direct image URLs in the page
        image_patterns = [
            r'https://[^"\']*\.sharepoint\.com[^"\']*\.png[^"\']*',
            r'https://[^"\']*/_layouts/15/download\.aspx[^"\']*',
        ]
        
        for pattern in image_patterns:
            matches = re.findall(pattern, html)
            if matches:
                print(f"✓ Found image URL pattern: {pattern}")
                for match in matches:
                    clean_url = unquote(match)
                    print(f"🖼️  Image URL: {clean_url}")
                    return clean_url
        
        # Method 3: Parse any JSON objects that might contain file info
        json_patterns = [
            r'window\.g_fileInfo\s*=\s*({.*?});',
            r'SPClientContext\s*=\s*({.*?});',
            r'_spPageContextInfo\s*=\s*({.*?});'
        ]
        
        for pattern in json_patterns:
            matches = re.findall(pattern, html, re.DOTALL)
            for match in matches:
                try:
                    data = json.loads(match)
                    # Look for download-related fields
                    for key, value in data.items():
                        if 'download' in key.lower() and isinstance(value, str) and value.startswith('http'):
                            print(f"📥 Found in JSON ({key}): {value}")
                            return value
                except:
                    continue
        
        print("❌ Could not find download URL in page")
        
        # Debug: Save HTML for manual inspection
        with open("sharepoint_debug.html", "w", encoding='utf-8') as f:
            f.write(html)
        print("💾 Saved HTML to sharepoint_debug.html for manual inspection")
        
        return None
        
    except Exception as e:
        print(f"❌ Error: {e}")
        return None

def download_from_url(download_url, filename):
    """Download file from extracted URL"""
    headers = {
        'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/139.0.0.0 Safari/537.36',
    }
    
    try:
        response = requests.get(download_url, headers=headers, stream=True)
        response.raise_for_status()
        
        total_size = 0
        with open(filename, 'wb') as f:
            for chunk in response.iter_content(chunk_size=8192):
                if chunk:
                    f.write(chunk)
                    total_size += len(chunk)
        
        file_size = total_size
        print(f"✅ Downloaded: {filename} ({file_size} bytes)")
        return True
        
    except Exception as e:
        print(f"❌ Download failed: {e}")
        return False

def main():
    sharing_url = "https://carorg.sharepoint.com/:i:/s/CAR-RE-PublicProducts/ERwqkypR2rtLg8AHUtC2yUAB-ljnk_3GVb4bxlTZE2pZpg"
    
    print("Extracting Download URL from SharePoint")
    print("=" * 50)
    
    download_url = extract_download_url_from_sharepoint(sharing_url)
    
    if download_url:
        print(f"\n🎯 Found download URL!")
        print(f"URL: {download_url}")
        
        # Try to download
        success = download_from_url(download_url, "extracted_download.png")
        if success:
            print("✅ Successfully downloaded using extracted URL!")
        else:
            print("❌ Download failed - URL might need authentication")
    else:
        print("❌ Could not extract download URL")

if __name__ == "__main__":
    main()