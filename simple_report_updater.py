"""
Simple Report Updater
Sends the report and PNG images to GPT and gets back the complete updated report.
Increments all months by 1 (7月 -> 8月).
"""

import os
from openai import OpenAI
import base64
from datetime import datetime
from PIL import Image
import io

class SimpleReportUpdater:
    def __init__(self, api_key: str):
        """Initialize with OpenAI API key"""
        self.client = OpenAI(api_key=api_key)
        self.work_directory = self.setup_work_directory()

    def setup_work_directory(self):
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

    def encode_image_to_base64(self, image_path: str, max_size: int = 2000) -> str:
        """Convert image to base64 with optional resizing to reduce payload size"""
        print(f"\n🔍 DEBUG: Attempting to read file:")
        print(f"   File path: {image_path}")
        print(f"   File exists: {os.path.exists(image_path)}")
        print(f"   Is absolute path: {os.path.isabs(image_path)}")
        print(f"   Current working directory: {os.getcwd()}")

        try:
            # Get file stats for debugging
            if os.path.exists(image_path):
                file_stat = os.stat(image_path)
                print(f"   Original file size: {file_stat.st_size:,} bytes")
                print(f"   File permissions: {oct(file_stat.st_mode)}")
                print(f"   File owner UID: {file_stat.st_uid}")
                print(f"   Current process UID: {os.getuid()}")

            # Open and resize image to reduce payload
            print(f"   Attempting to open and resize image...")
            img = Image.open(image_path)
            print(f"   ✅ Image opened: {img.size[0]}x{img.size[1]} pixels")

            # Resize if image is too large
            if img.size[0] > max_size or img.size[1] > max_size:
                # Calculate new size maintaining aspect ratio
                ratio = min(max_size / img.size[0], max_size / img.size[1])
                new_size = (int(img.size[0] * ratio), int(img.size[1] * ratio))
                img = img.resize(new_size, Image.Resampling.LANCZOS)
                print(f"   🔽 Resized to: {new_size[0]}x{new_size[1]} pixels")

            # Convert to bytes
            buffer = io.BytesIO()
            img.save(buffer, format='PNG', optimize=True)
            data = buffer.getvalue()
            print(f"   ✅ Compressed size: {len(data):,} bytes (saved {file_stat.st_size - len(data):,} bytes)")

            # Encode to base64
            encoded = base64.b64encode(data).decode('utf-8')
            print(f"   ✅ Encoded to base64: {len(encoded):,} characters")
            return encoded

        except PermissionError as e:
            print(f"\n❌ PERMISSION ERROR: {e}")
            print(f"📁 File: {image_path}")
            print(f"📁 Basename: {os.path.basename(image_path)}")
            print(f"📁 Directory: {os.path.dirname(image_path)}")
            print(f"\n💡 SOLUTION:")
            print(f"   macOS is blocking Python from reading files in the Downloads folder.")
            print(f"   To fix this:")
            print(f"   1. Go to: System Settings > Privacy & Security > Full Disk Access")
            print(f"   2. Click the '+' button")
            print(f"   3. Add your Terminal app (or Python executable)")
            print(f"   4. Restart your terminal and try again")
            print(f"\n   OR move the PNG files to the current directory: {os.getcwd()}")
            raise
        except FileNotFoundError as e:
            print(f"\n❌ FILE NOT FOUND: {e}")
            print(f"📁 Looking for: {image_path}")
            print(f"📁 Basename: {os.path.basename(image_path)}")
            print(f"📁 Directory exists: {os.path.exists(os.path.dirname(image_path))}")
            raise
        except Exception as e:
            print(f"\n❌ ERROR reading file")
            print(f"   File: {image_path}")
            print(f"   Basename: {os.path.basename(image_path)}")
            print(f"   Error type: {type(e).__name__}")
            print(f"   Error message: {e}")
            import traceback
            print(f"\n📋 Stack trace:")
            traceback.print_exc()
            raise

    def copy_files_to_local(self, source_files: list) -> list:
        """Copy files to current directory to avoid permission issues"""
        import shutil
        local_files = []

        print(f"\n📋 Copying files to current directory to avoid permission issues...")
        for source_file in source_files:
            try:
                filename = os.path.basename(source_file)
                local_path = os.path.join(os.getcwd(), filename)

                # Skip if already in current directory
                if os.path.abspath(source_file) == os.path.abspath(local_path):
                    local_files.append(local_path)
                    print(f"   ✅ Already local: {filename}")
                    continue

                # Copy file
                shutil.copy2(source_file, local_path)
                local_files.append(local_path)
                print(f"   ✅ Copied: {filename}")
            except Exception as e:
                print(f"   ❌ Failed to copy {filename}: {e}")
                # Try to use original file if copy fails
                local_files.append(source_file)

        return local_files

    def find_files(self) -> tuple:
        """Find report file and PNG files in work directory"""
        # Find report file - check both current dir and work dir
        report_file = None

        # First try current directory (for the sample)
        if os.path.exists("monthly report sample.txt"):
            report_file = "monthly report sample.txt"

        # Find PNG files in work directory
        png_files = []
        counties = [
            ("Santa Clara", "Santa Clara.png"),
            ("San Mateo", "San Mateo.png"),
            ("Alameda", "Alameda.png"),
            ("San Francisco", "San Francisco.png"),
        ]

        print(f"🔍 Looking for files in: {self.work_directory}")

        for i, (county_name, plain_filename) in enumerate(counties, 1):
            # Try multiple naming conventions
            filename_with_name = f"{i}{county_name}.png"  # 1Santa Clara.png (legacy)
            filename_simple = f"{i}.png"                 # 1.png (legacy)

            candidates = [
                os.path.join(self.work_directory, plain_filename),
                os.path.join(self.work_directory, filename_with_name),
                os.path.join(self.work_directory, filename_simple),
                plain_filename,
                filename_with_name,
                filename_simple,
            ]

            found = None
            for candidate in candidates:
                if os.path.exists(candidate):
                    found = candidate
                    break

            if found:
                png_files.append(found)
                print(f"✅ Found: {found}")
            else:
                print(f"❌ Missing: {plain_filename} (or {filename_with_name}/{filename_simple})")

        # Copy files to local directory to avoid permission issues
        if png_files:
            png_files = self.copy_files_to_local(png_files)

        return report_file, png_files

    def read_report_content(self, report_file: str) -> str:
        """Read the current report content"""
        try:
            with open(report_file, "r", encoding="utf-8") as f:
                return f.read()
        except Exception as e:
            print(f"❌ Error reading report: {e}")
            return ""

    def update_report_with_gpt(self, report_content: str, png_files: list) -> str:
        """Ask GPT to update the entire report"""
        print("🤖 Sending report and images to GPT for complete update...")

        try:
            # Create message content
            message_content = [
                {
                    "type": "text",
                    "text": f"""You are tasked with updating a Chinese real estate monthly report. You MUST extract new data from the provided PNG images and update ALL numerical values in the report.

CRITICAL REQUIREMENTS:
1. EXTRACT ALL DATA from PNG images: median prices, price changes, sales volume changes, days on market, months of inventory
2. UPDATE EVERY NUMBER in the report with the extracted data
3. DO NOT keep any old numbers - everything must be updated with new data from images
4. Keep the exact same text structure and formatting

WHAT YOU MUST UPDATE:
- All median prices (中位成交价) - extract from PNG images
- All percentage changes (比上月上升/下降) - extract from PNG images  
- All sales volume changes (成交量比上月) - extract from PNG images
- All inventory levels (月存量) - extract from PNG images
- All days on market (市场停留时间) - extract from PNG images
- All year-over-year comparisons (比去年同期) - extract from PNG images
- Month references: increment by +1 (6月 → 7月, 7月 → 8月)

SUBTITLE LINE 2 FORMAT:
[市场月报]X月速递：均价XXX万，比上月XXX%，销售量XXX%
- Use Santa Clara County data for this summary line
- Extract: median price, month-over-month price change, month-over-month volume change

DATA EXTRACTION PRIORITY:
1. Look at each PNG image carefully
2. Find the current month's data for each county
3. Extract ALL numerical values shown
4. Replace corresponding values in the report
5. Ensure NO old data remains

FORMAT RULES:
- Use "上升X%" or "下降X%" for changes
- Use "4%" not "4.0%" for whole number percentages
- Use "与上月持平" for no change instead of "0%"

OUTPUT RULES:
- Return ONLY the updated Chinese report content; do not add any preface, notes, or summaries
- Do NOT include lines like "Here's the updated report..." or "Note: ..."
- Preserve the report structure and formatting exactly, with no extra blank lines

CURRENT REPORT:
{report_content}

IMPORTANT: Every number in this report must be updated with fresh data from the PNG images. Do not preserve any old numerical values."""
                }
            ]

            # Add images (encode once and store)
            encoded_images = []
            for i, png_file in enumerate(png_files, 1):
                # Get county name from filename
                if "Santa Clara" in png_file:
                    county_name = "Santa Clara"
                elif "San Mateo" in png_file:
                    county_name = "San Mateo"
                elif "Alameda" in png_file:
                    county_name = "Alameda"
                elif "San Francisco" in png_file:
                    county_name = "San Francisco"
                else:
                    county_name = "Unknown"

                print(f"📸 Processing image {i}/{len(png_files)}: {os.path.basename(png_file)}")

                try:
                    base64_image = self.encode_image_to_base64(png_file)
                    encoded_images.append((county_name, base64_image))
                    print(f"   ✅ Successfully encoded {county_name} image ({len(base64_image):,} chars)")
                except Exception as e:
                    print(f"   ❌ Failed to encode {county_name} image")
                    raise  # Re-raise to be caught by outer exception handler

            # Add encoded images to message
            for county_name, base64_image in encoded_images:
                message_content.extend([
                    {
                        "type": "text",
                        "text": f"\n=== {county_name} County Data ==="
                    },
                    {
                        "type": "image_url",
                        "image_url": {
                            "url": f"data:image/png;base64,{base64_image}"
                        }
                    }
                ])

            print(f"\n📊 Request details:")
            print(f"   - Model: gpt-5.1-chat-latest")
            print(f"   - Images: {len(png_files)} PNG files")
            print(f"   - Max completion tokens: 8000")

            # Calculate approximate payload size
            total_image_size = sum(len(base64_img) for _, base64_img in encoded_images)
            print(f"   - Total encoded image size: {total_image_size:,} characters (~{total_image_size/1024/1024:.2f} MB)")

            # Make API call (single attempt, no retry)
            print(f"\n🔄 Sending request to OpenAI API...")
            response = self.client.chat.completions.create(
                model="gpt-5.1-chat-latest",
                messages=[{"role": "user", "content": message_content}],
                max_completion_tokens=8000,
                timeout=300  # 5 minute timeout
            )

            # Debug response details
            print(f"✅ API call successful!")
            if hasattr(response, 'usage') and response.usage:
                usage = response.usage
                print(f"📈 Token usage:")
                print(f"   - Prompt tokens: {usage.prompt_tokens:,}")
                print(f"   - Completion tokens: {usage.completion_tokens:,}")
                if hasattr(usage, 'reasoning_tokens') and usage.reasoning_tokens:
                    print(f"   - Reasoning tokens: {usage.reasoning_tokens:,}")
                print(f"   - Total tokens: {usage.total_tokens:,}")

            # Get the updated report content
            content = response.choices[0].message.content
            
            if not content:
                print("❌ Empty response from GPT")
                return ""
            
            print(f"📋 Received updated report ({len(content)} characters)")
            print(f"🔍 Preview: {content[:200]}...")
            
            return content

        except PermissionError as e:
            print(f"\n❌ PERMISSION ERROR - Cannot read image files")
            print(f"   {e}")
            print(f"\n🔧 This is a macOS security restriction on the Downloads folder.")
            return ""
        except FileNotFoundError as e:
            print(f"\n❌ FILE NOT FOUND ERROR")
            print(f"   {e}")
            return ""
        except Exception as e:
            print(f"\n❌ Error updating report with GPT")
            print(f"   Error type: {type(e).__name__}")
            print(f"   Error message: {e}")

            # Print more details for connection errors
            if "connection" in str(e).lower() or "timeout" in str(e).lower():
                print(f"\n🌐 This appears to be a network/connection issue:")
                print(f"   - Check your internet connection")
                print(f"   - Verify OpenAI API key is valid")
                print(f"   - Check if OpenAI services are accessible")

            # Print stack trace for debugging
            import traceback
            print(f"\n📋 Full error details:")
            traceback.print_exc()

            return ""

    def save_updated_report(self, updated_content: str) -> bool:
        """Save the updated report to work directory"""
        try:
            output_file = os.path.join(self.work_directory, "monthly_report_updated.txt")
            with open(output_file, "w", encoding="utf-8") as f:
                f.write(updated_content)
            
            print(f"✅ Updated report saved as: {output_file}")
            
            # Show line count and size
            lines = updated_content.split('\n')
            print(f"📄 Report stats:")
            print(f"   - Lines: {len(lines)}")
            print(f"   - Characters: {len(updated_content):,}")
            print(f"📁 Saved to: {self.work_directory}")
            
            return True

        except Exception as e:
            print(f"❌ Error saving updated report: {e}")
            return False

    def run_update(self) -> bool:
        """Main update process"""
        print("🚀 Simple Report Updater")
        print("="*50)
        print("Sends report + images to GPT, gets back complete updated report")
        print("="*50)

        # Find files
        report_file, png_files = self.find_files()
        
        if not report_file:
            print("❌ No report file found. Need 'monthly report sample.txt'")
            return False
        
        if not png_files:
            print("❌ No PNG files found")
            return False

        print(f"📄 Using report: {report_file}")
        print(f"🖼️ Using {len(png_files)} PNG files")

        # Read current report
        report_content = self.read_report_content(report_file)
        if not report_content:
            return False

        # Get updated report from GPT
        updated_content = self.update_report_with_gpt(report_content, png_files)
        if not updated_content:
            print("❌ Could not get updated report from GPT")
            return False

        # Save updated report
        return self.save_updated_report(updated_content)

def main():
    """Main execution"""
    print("🤖 Simple Report Updater")
    print("Gets complete updated report from GPT")
    print()

    # Try to get API key from environment first
    api_key = os.getenv("OPENAI_API_KEY", "").strip()
    
    if api_key:
        print("✅ Found OpenAI API key in environment variable")
    else:
        print("🔑 No OPENAI_API_KEY environment variable found")
        api_key = input("Enter your OpenAI API key: ").strip()
        
        if not api_key:
            print("❌ OpenAI API key required for this script")
            return

    # Initialize and run
    updater = SimpleReportUpdater(api_key)
    success = updater.run_update()

    if success:
        print("\n🎉 Report updated successfully!")
        print("📄 Check 'monthly_report_updated.txt' for the complete updated report")
    else:
        print("\n❌ Failed to update report")

if __name__ == "__main__":
    main()
