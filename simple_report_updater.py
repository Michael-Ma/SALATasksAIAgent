"""
Simple Report Updater
Sends the report and PNG images to GPT and gets back the complete updated report.
Increments all months by 1 (7月 -> 8月).
"""

import os
from openai import OpenAI
import base64
from datetime import datetime

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

    def encode_image_to_base64(self, image_path: str) -> str:
        """Convert image to base64"""
        with open(image_path, "rb") as image_file:
            return base64.b64encode(image_file.read()).decode('utf-8')

    def find_files(self) -> tuple:
        """Find report file and PNG files in work directory"""
        # Find report file - check both current dir and work dir
        report_file = None
        
        # First try current directory (for the sample)
        if os.path.exists("monthly report sample.txt"):
            report_file = "monthly report sample.txt"
        
        # Find PNG files in work directory
        png_files = []
        counties = ["Santa Clara", "San Mateo", "Alameda", "San Francisco"]
        
        print(f"🔍 Looking for files in: {self.work_directory}")
        
        for i, county_name in enumerate(counties, 1):
            # Try both naming conventions
            filename_with_name = f"{i}{county_name}.png"  # 1Santa Clara.png (Step 1)
            filename_simple = f"{i}.png"  # 1.png (Step 3)
            
            # Check work directory for both formats
            filepath_with_name = os.path.join(self.work_directory, filename_with_name)
            filepath_simple = os.path.join(self.work_directory, filename_simple)
            
            if os.path.exists(filepath_with_name):
                png_files.append(filepath_with_name)
                print(f"✅ Found: {filepath_with_name}")
            elif os.path.exists(filepath_simple):
                png_files.append(filepath_simple)
                print(f"✅ Found: {filepath_simple}")
            else:
                # Also check current directory as fallback
                if os.path.exists(filename_with_name):
                    png_files.append(filename_with_name)
                    print(f"✅ Found (fallback): {filename_with_name}")
                elif os.path.exists(filename_simple):
                    png_files.append(filename_simple)
                    print(f"✅ Found (fallback): {filename_simple}")
                else:
                    print(f"❌ Missing: {filename_with_name} or {filename_simple}")
        
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

CURRENT REPORT:
{report_content}

IMPORTANT: Every number in this report must be updated with fresh data from the PNG images. Do not preserve any old numerical values."""
                }
            ]

            # Add images
            for png_file in png_files:
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
                
                base64_image = self.encode_image_to_base64(png_file)
                
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

            print(f"📊 Request details:")
            print(f"   - Model: gpt-5")
            print(f"   - Images: {len(png_files)} PNG files")
            print(f"   - Max completion tokens: 8000")

            # Make API call
            response = self.client.chat.completions.create(
                model="gpt-5",
                messages=[{"role": "user", "content": message_content}],
                max_completion_tokens=8000
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

        except Exception as e:
            print(f"❌ Error updating report with GPT: {e}")
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