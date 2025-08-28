"""
Interactive Workflow Startup Script
Guides through complete report download workflow with navigation controls.
"""

import os
import sys
import subprocess
import webbrowser
from time import sleep
from datetime import datetime

class WorkflowGuide:
    def __init__(self):
        self.current_step = 1
        self.total_steps = 5
        self.completed_steps = set()
        self.work_directory = self.setup_work_directory()
    
    def setup_work_directory(self):
        """Create SALA Report directory for current month"""
        try:
            # Get current month name
            current_month = datetime.now().strftime("%B %Y")  # e.g., "January 2025"
            
            # Create directory path
            downloads_path = os.path.expanduser("~/Downloads")
            work_dir = os.path.join(downloads_path, f"SALA Report {current_month}")
            
            # Create directory if it doesn't exist
            if not os.path.exists(work_dir):
                os.makedirs(work_dir)
                print(f"📁 Created work directory: {work_dir}")
            else:
                print(f"📁 Using existing work directory: {work_dir}")
            
            return work_dir
            
        except Exception as e:
            print(f"❌ Error creating work directory: {e}")
            # Fallback to current directory
            return os.getcwd()
        
    def clear_screen(self):
        """Clear terminal screen"""
        os.system('cls' if os.name == 'nt' else 'clear')
    
    def show_header(self):
        """Display workflow header"""
        print("🚀 CAR.org Report Download Workflow")
        print("="*60)
        print(f"Step {self.current_step} of {self.total_steps}")
        print(f"Progress: {len(self.completed_steps)}/{self.total_steps} steps completed")
        print(f"📁 Work Directory: {self.work_directory}")
        print("="*60)
        print()
    
    def show_navigation(self):
        """Display navigation options"""
        print("\n" + "─"*60)
        print("Navigation / 导航:")
        if self.current_step > 1:
            print("  [b] ← Back to previous step / 返回上一步")
        if self.current_step < self.total_steps:
            print("  [Enter] → Next step / 下一步")
        print("  [r] ↻ Repeat current step / 重复当前步骤")
        print("  [c] ✓ Mark step as completed / 标记步骤为完成")
        print("  [s] 📊 Show progress summary / 显示进度摘要")
        print("  [q] ❌ Quit workflow / 退出工作流程")
        print("─"*60)
    
    def get_user_input(self):
        """Get user navigation choice"""
        while True:
            choice = input("\nEnter your choice / 输入您的选择: ").lower().strip()
            if choice == '':
                return 'n'  # Enter key maps to next step
            elif choice in ['b', 'r', 'c', 's', 'q']:
                return choice
            print("❌ Invalid choice. Please enter: Enter, b, r, c, s, or q")
            print("❌ 无效选择。请输入：回车键、b、r、c、s、或 q")
    
    def show_progress_summary(self):
        """Show completed steps summary"""
        self.clear_screen()
        print("📊 WORKFLOW PROGRESS SUMMARY")
        print("📊 工作流程进度摘要")
        print("="*60)
        
        steps_info = [
            ("Download PNG files from CAR.org and SharePoint", "从CAR.org和SharePoint下载PNG文件"),
            ("Run final_working_downloader.py script", "运行final_working_downloader.py脚本"),
            ("Download county and city images via Power BI", "通过Power BI下载县和市级图像"),
            ("MLSL BI portal tasks (login, reports, email, convert)", "MLSL BI门户任务（登录、报告、邮件、转换）"),
            ("Generate updated monthly report using AI", "使用AI生成更新的月度报告")
        ]
        
        for i, (step_desc_en, step_desc_cn) in enumerate(steps_info, 1):
            status = "✅" if i in self.completed_steps else "⏳"
            print(f"{status} Step {i}: {step_desc_en}")
            print(f"    步骤{i}: {step_desc_cn}")
        
        print(f"\n📈 Progress: {len(self.completed_steps)}/{self.total_steps} steps completed")
        print(f"📈 进度: {len(self.completed_steps)}/{self.total_steps} 步骤已完成")
        
        if len(self.completed_steps) == self.total_steps:
            print("\n🎉 ALL STEPS COMPLETED! Workflow finished successfully!")
            print("🎉 所有步骤已完成！工作流程成功完成！")
        
        input("\nPress Enter to continue / 按Enter继续...")
    
    def step_1(self):
        """Step 1: Download PNG files from CAR.org and SharePoint"""
        self.clear_screen()
        self.show_header()
        
        print("📋 Step 1: Download PNG files from CAR.org and SharePoint")
        print("─"*60)
        print("🎯 Target: Download PNG files for 4 counties:")
        print("   • Santa Clara County")
        print("   • San Mateo County") 
        print("   • San Francisco County")
        print("   • Alameda County")
        print()
        print("🔗 PRIMARY URL: https://www.car.org/marketing/chartsandgraphs/marketupdate")
        print("🔗 ALTERNATIVE: https://carorg.sharepoint.com/:f:/s/CAR-RE-PublicProducts/ElrCKkQh6_ZMpe5RIZgOohoB33WDC9L1NlkigRWlqWwvGg?e=Vi1XJZ")
        print()
        print("📝 ENGLISH INSTRUCTIONS:")
        print("─" * 30)
        print("1. Try the primary CAR.org marketing page first")
        print("2. If needed, use the SharePoint alternative URL")
        print("3. Look for charts/graphs related to the 4 target counties")
        print("4. Download PNG files for each county")
        print("5. Save them with naming format:")
        print("   • 1Santa Clara.png")
        print("   • 2San Mateo.png") 
        print("   • 3San Francisco.png")
        print("   • 4Alameda.png")
        print()
        print("📝 中文说明:")
        print("─" * 30)
        print("1. 首先尝试主要的CAR.org营销页面")
        print("2. 如需要，使用SharePoint备用网址")
        print("3. 寻找与4个目标县相关的图表/图形")
        print("4. 下载每个县的PNG文件")
        print("5. 使用以下命名格式保存:")
        print("   • 1Santa Clara.png")
        print("   • 2San Mateo.png") 
        print("   • 3San Francisco.png")
        print("   • 4Alameda.png")
        print()
        
        if input("🔗 Open CAR.org marketing page? (y/n): ").lower() == 'y':
            try:
                webbrowser.open("https://www.car.org/marketing/chartsandgraphs/marketupdate")
                print("✅ Opened CAR.org marketing page in browser")
                print("✅ 已在浏览器中打开CAR.org营销页面")
            except Exception as e:
                print(f"❌ Error opening browser: {e}")
                print("Please manually open: https://www.car.org/marketing/chartsandgraphs/marketupdate")
        
        print("\n🔗 Need SharePoint alternative?")
        print("🔗 需要SharePoint备选方案吗？")
        if input("Open SharePoint folder? (y/n): ").lower() == 'y':
            try:
                webbrowser.open("https://carorg.sharepoint.com/:f:/s/CAR-RE-PublicProducts/ElrCKkQh6_ZMpe5RIZgOohoB33WDC9L1NlkigRWlqWwvGg?e=Vi1XJZ")
                print("✅ Opened SharePoint folder in browser")
                print("✅ 已在浏览器中打开SharePoint文件夹")
            except Exception as e:
                print(f"❌ Error opening browser: {e}")
                print("Please manually open the SharePoint URL above")
        
        return False  # Don't exit workflow
    
    def step_2(self):
        """Step 2: Prepare final working downloader"""
        self.clear_screen()
        self.show_header()
        
        print("📋 Step 2: Prepare final_working_downloader.py script")
        print("📋 步骤2: 准备final_working_downloader.py脚本")
        print("─"*60)
        print("🎯 Goal: Set up automated Power BI download script")
        print("🎯 目标: 设置自动化Power BI下载脚本")
        print()
        print("📝 ENGLISH PREPARATION CHECKLIST:")
        print("─" * 30)
        print("✓ Ensure final_working_downloader.py is in current directory")
        print("✓ Chrome browser is installed and updated")
        print("✓ Internet connection is stable")
        print("✓ No other Chrome instances accessing CAR.org")
        print()
        print("📝 中文准备清单:")
        print("─" * 30)
        print("✓ 确保final_working_downloader.py在当前目录中")
        print("✓ Chrome浏览器已安装并更新")
        print("✓ 网络连接稳定")
        print("✓ 没有其他Chrome实例访问CAR.org")
        print()
        
        # Check if script exists
        script_exists = os.path.exists("final_working_downloader.py")
        print(f"📄 Script status: {'✅ Found' if script_exists else '❌ Missing'}")
        print(f"📄 脚本状态: {'✅ 找到' if script_exists else '❌ 缺失'}")
        
        if script_exists:
            print("\n🔧 ENGLISH - Ready to run the downloader script!")
            print("The script will:")
            print("• Open CAR.org Power BI page automatically")
            print("• Wait for you to manually select each county")
            print("• Automatically extract and download both county and city images")
            print("• Save files with format: 1.png, 1(1).png, 1(2).png, etc.")
            print("\n🔧 中文 - 准备运行下载器脚本!")
            print("脚本将:")
            print("• 自动打开CAR.org Power BI页面")
            print("• 等待您手动选择每个县")
            print("• 自动提取并下载县和市级图像")
            print("• 保存文件格式: 1.png, 1(1).png, 1(2).png等")
        else:
            print("\n❌ ENGLISH - final_working_downloader.py not found!")
            print("Please ensure the script is in the current directory.")
            print("\n❌ 中文 - 找不到final_working_downloader.py!")
            print("请确保脚本在当前目录中。")
        
        return False  # Don't exit workflow
    
    def step_3(self):
        """Step 3: Run final working downloader"""
        self.clear_screen()
        self.show_header()
        
        print("📋 Step 3: Run final_working_downloader.py")
        print("📋 步骤3: 运行final_working_downloader.py")
        print("─"*60)
        print("🚀 Execute the automated Power BI downloader")
        print("🚀 执行自动化Power BI下载器")
        print()
        print("📝 ENGLISH WORKFLOW:")
        print("─" * 30)
        print("1. Script opens CAR.org Power BI page")
        print("2. You manually select counties from dropdown:")
        print("   → Santa Clara (files: 1.png, 1(1).png, 1(2).png...)")
        print("   → San Mateo (files: 2.png, 2(1).png, 2(2).png...)")
        print("   → Alameda (files: 3.png, 3(1).png, 3(2).png...)")
        print("   → San Francisco (files: 4.png, 4(1).png, 4(2).png...)")
        print("3. Script automatically downloads all county + city images")
        print()
        print("📝 中文工作流程:")
        print("─" * 30)
        print("1. 脚本打开CAR.org Power BI页面")
        print("2. 您手动从下拉菜单中选择县:")
        print("   → Santa Clara (文件: 1.png, 1(1).png, 1(2).png...)")
        print("   → San Mateo (文件: 2.png, 2(1).png, 2(2).png...)")
        print("   → Alameda (文件: 3.png, 3(1).png, 3(2).png...)")
        print("   → San Francisco (文件: 4.png, 4(1).png, 4(2).png...)")
        print("3. 脚本自动下载所有县和市级图像")
        print()
        
        print("🚀 Starting final_working_downloader.py automatically...")
        print("🚀 自动启动final_working_downloader.py...")
        try:
            print("▶️ Starting downloader script...")
            print("▶️ 启动下载器脚本...")
            print(f"📁 Files will be saved to: {self.work_directory}")
            print(f"📁 文件将保存到: {self.work_directory}")
            
            # Set environment variable for work directory
            env = os.environ.copy()
            env['SALA_WORK_DIR'] = self.work_directory
            
            result = subprocess.run([sys.executable, "final_working_downloader.py"], 
                                  cwd=os.getcwd(), env=env)
            if result.returncode == 0:
                print("✅ Script completed successfully!")
                print("✅ 脚本成功完成!")
            else:
                print("⚠️ Script finished with warnings/errors")
                print("⚠️ 脚本完成但有警告/错误")
        except Exception as e:
            print(f"❌ Error running script: {e}")
            print(f"❌ 运行脚本时出错: {e}")
            print("You can manually run: python final_working_downloader.py")
            print("您可以手动运行: python final_working_downloader.py")
        
        return False  # Don't exit workflow
    
    def step_4(self):
        """Step 4: MLSL BI Portal Tasks (Combined)"""
        self.clear_screen()
        self.show_header()
        
        print("📋 Step 4: MLSL BI Portal Tasks")
        print("─"*60)
        print("🎯 Complete all MLSL BI portal related tasks")
        print()
        
        print("🌐 ENGLISH INSTRUCTIONS:")
        print("─" * 30)
        print("1. 🔗 Open: https://mlsl.aculist.com/BI")
        print("2. 👤 Login as: Sunnie (use saved password)")
        print("3. 🔍 Search for any ZIP-related content")
        print("4. 📂 Navigate to 'Favorite Reports' section")
        print("5. 📋 Select ALL reports in Favorite Reports")
        print("6. 📧 Send all reports to your email address")
        print("7. 📥 Download PDF attachments from your email")
        print("8. 🔄 Convert PDFs to PNG format")
        print("9. 📝 Rename files using format: 6(1).png, 6(2).png, 6(3).png...")
        print("10. 📁 Save PNG files to work directory")
        print()
        
        print("🇨🇳 中文说明:")
        print("─" * 30)
        print("1. 🔗 打开网址: https://mlsl.aculist.com/BI")
        print("2. 👤 用户名: Sunnie 登录 (使用保存的密码)")
        print("3. 🔍 搜索任何与ZIP相关的内容")
        print("4. 📂 导航到收藏报告部分")
        print("5. 📋 选择收藏报告中的所有报告")
        print("6. 📧 将所有报告发送到您的邮箱")
        print("7. 📥 从邮箱下载PDF附件")
        print("8. 🔄 将PDF转换为PNG格式")
        print("9. 📝 重命名文件格式: 6(1).png, 6(2).png, 6(3).png...")
        print("10. 📁 将PNG文件保存到工作目录")
        print()
        
        print("💡 TIPS / 提示:")
        print("• Look for 'Reports', 'Favorites', or 'My Reports' sections")
        print("• 寻找'报告'、'收藏'或'我的报告'部分")
        print("• Use online converters like PDF to PNG for conversion")
        print("• 使用在线转换器如PDF转PNG进行转换")
        print()
        
        # Auto-open the URL
        try:
            webbrowser.open("https://mlsl.aculist.com/BI")
            print("✅ Opened MLSL BI portal in browser")
            print("✅ 已在浏览器中打开MLSL BI门户")
        except Exception as e:
            print(f"❌ Error opening browser: {e}")
            print("Please manually open: https://mlsl.aculist.com/BI")
        
        print("\n⏳ Complete all the above tasks, then return here to continue")
        print("⏳ 完成上述所有任务，然后回到这里继续")
        
        return False  # Don't exit workflow
    
    def step_5(self):
        """Step 5: Generate updated monthly report using AI"""
        self.clear_screen()
        self.show_header()
        
        print("📋 Step 5: Generate Updated Monthly Report Using AI")
        print("📋 步骤5: 使用AI生成更新的月度报告")
        print("─"*60)
        print("🤖 Use AI to create updated monthly report from downloaded data")
        print("🤖 使用AI从下载的数据创建更新的月度报告")
        print()
        print("📝 ENGLISH REPORT GENERATION PROCESS:")
        print("─" * 30)
        print("1. 📊 Script analyzes all downloaded PNG files")
        print("2. 🔄 Compares current data with previous month's report")
        print("3. 📝 Updates all numbers and percentages automatically")
        print("4. 🗓️ Increments all months by +1 (7月 → 8月)")
        print("5. 🇨🇳 Uses proper Chinese terms for changes (上升/下降)")
        print("6. 💾 Saves complete updated report to work directory")
        print()
        print("📝 中文报告生成过程:")
        print("─" * 30)
        print("1. 📊 脚本分析所有下载的PNG文件")
        print("2. 🔄 将当前数据与上月报告进行比较")
        print("3. 📝 自动更新所有数字和百分比")
        print("4. 🗓️ 所有月份增加+1 (7月 → 8月)")
        print("5. 🇨🇳 使用正确的中文变化术语 (上升/下降)")
        print("6. 💾 将完整的更新报告保存到工作目录")
        print()
        print("🎯 REQUIREMENTS / 要求:")
        print("• OpenAI API key (set as OPENAI_API_KEY environment variable)")
        print("• OpenAI API密钥 (设置为OPENAI_API_KEY环境变量)")
        print("• PNG files in work directory (from previous steps)")
        print("• 工作目录中的PNG文件 (来自之前的步骤)")
        print("• Original monthly report sample file")
        print("• 原始月度报告样本文件")
        print()
        
        print("🚀 Starting simple_report_updater.py automatically...")
        print("🚀 自动启动simple_report_updater.py...")
        try:
            print("▶️ Starting AI report updater...")
            print("▶️ 启动AI报告更新器...")
            print(f"📁 Using work directory: {self.work_directory}")
            print(f"📁 使用工作目录: {self.work_directory}")
            
            # Set environment variable for work directory
            env = os.environ.copy()
            env['SALA_WORK_DIR'] = self.work_directory
            
            result = subprocess.run([sys.executable, "simple_report_updater.py"], 
                                  cwd=os.getcwd(), env=env)
            if result.returncode == 0:
                print("✅ AI report generation completed successfully!")
                print("✅ AI报告生成成功完成!")
                print(f"📄 Generated: {self.work_directory}/monthly_report_updated.txt")
                print(f"📄 已生成: {self.work_directory}/monthly_report_updated.txt")
                
                # Ask permission to replace the sample file
                self.offer_to_replace_sample_report()
                
                # Mark step 5 as completed and show final success message
                self.completed_steps.add(5)
                print("\n" + "="*60)
                print("🎉 WORKFLOW COMPLETED SUCCESSFULLY!")
                print("🎉 工作流程成功完成!")
                print("="*60)
                print("All 5 steps have been completed:")
                print("所有5个步骤已完成:")
                print("✅ Downloaded PNG files from CAR.org and SharePoint")
                print("✅ 从CAR.org和SharePoint下载了PNG文件")
                print("✅ Ran final_working_downloader.py script")
                print("✅ 运行了final_working_downloader.py脚本")
                print("✅ Downloaded county and city images via Power BI")
                print("✅ 通过Power BI下载了县和市级图像")
                print("✅ Completed MLSL BI portal tasks")
                print("✅ 完成了MLSL BI门户任务")
                print("✅ Generated updated monthly report using AI")
                print("✅ 使用AI生成了更新的月度报告")
                print()
                print("🏁 All report download and generation tasks completed!")
                print("🏁 所有报告下载和生成任务已完成!")
                print(f"📁 Final files saved to: {self.work_directory}")
                print(f"📁 最终文件保存到: {self.work_directory}")
                
                # Exit the workflow
                return True  # Signal to exit workflow
                
            else:
                print("⚠️ Report generation finished with warnings/errors")
                print("⚠️ 报告生成完成但有警告/错误")
                
        except Exception as e:
            print(f"❌ Error running AI report updater: {e}")
            print(f"❌ 运行AI报告更新器时出错: {e}")
            print("You can manually run: python simple_report_updater.py")
            print("您可以手动运行: python simple_report_updater.py")
            
        
        return False  # Don't exit, continue with navigation
    
    
    def offer_to_replace_sample_report(self):
        """Offer to replace sample report with updated version"""
        try:
            updated_file = os.path.join(self.work_directory, "monthly_report_updated.txt")
            sample_file = "monthly report sample.txt"
            
            if not os.path.exists(updated_file):
                print("❌ Updated report not found, skipping replacement")
                print("❌ 找不到更新的报告，跳过替换")
                return
            
            if not os.path.exists(sample_file):
                print("❌ Sample report not found, skipping replacement")
                print("❌ 找不到样本报告，跳过替换")
                return
            
            print("\n" + "="*60)
            print("📄 REPORT REPLACEMENT PREVIEW")
            print("📄 报告替换预览")
            print("="*60)
            
            # Read both files
            with open(sample_file, 'r', encoding='utf-8') as f:
                old_content = f.read()
            
            with open(updated_file, 'r', encoding='utf-8') as f:
                new_content = f.read()
            
            # Show detailed line-by-line diff preview
            old_lines = old_content.split('\n')
            new_lines = new_content.split('\n')
            
            print("🔍 Line-by-line differences preview:")
            print("🔍 逐行差异预览:")
            print("─" * 80)
            
            # Find meaningful differences (skip empty lines and minor changes)
            differences_shown = 0
            max_differences = 10  # Limit to avoid overwhelming output
            
            for i, (old_line, new_line) in enumerate(zip(old_lines, new_lines)):
                if old_line.strip() != new_line.strip() and old_line.strip() and new_line.strip():
                    if differences_shown >= max_differences:
                        print("  ... (more differences not shown)")
                        print("  ... (更多差异未显示)")
                        break
                    
                    print(f"Line {i+1} / 第{i+1}行:")
                    print(f"  OLD / 旧: {old_line}")
                    print(f"  NEW / 新: {new_line}")
                    print("─" * 40)
                    differences_shown += 1
            
            # Handle case where new file has more lines
            if len(new_lines) > len(old_lines):
                print(f"New file has {len(new_lines) - len(old_lines)} additional lines")
                print(f"新文件多出 {len(new_lines) - len(old_lines)} 行")
            elif len(old_lines) > len(new_lines):
                print(f"New file has {len(old_lines) - len(new_lines)} fewer lines")
                print(f"新文件少了 {len(old_lines) - len(new_lines)} 行")
            
            if differences_shown == 0:
                print("No significant content differences found")
                print("未发现重要内容差异")
            
            print(f"\n📊 File comparison / 文件比较:")
            print(f"  Old file: {len(old_lines)} lines, {len(old_content)} chars")
            print(f"  旧文件: {len(old_lines)} 行, {len(old_content)} 字符")
            print(f"  New file: {len(new_lines)} lines, {len(new_content)} chars")
            print(f"  新文件: {len(new_lines)} 行, {len(new_content)} 字符")
            
            # Ask for permission
            print("\n🔄 Replace 'monthly report sample.txt' with updated version?")
            print("🔄 用更新版本替换'monthly report sample.txt'吗?")
            if input("Enter 'yes' to replace / 输入'yes'进行替换: ").lower() in ['yes', 'y']:
                # Make backup first
                backup_file = "monthly report sample.txt.backup"
                with open(backup_file, 'w', encoding='utf-8') as f:
                    f.write(old_content)
                print(f"📄 Backup saved as: {backup_file}")
                print(f"📄 备份保存为: {backup_file}")
                
                # Replace the file
                with open(sample_file, 'w', encoding='utf-8') as f:
                    f.write(new_content)
                
                print("✅ Successfully replaced monthly report sample.txt")
                print("✅ 成功替换monthly report sample.txt")
                print("📄 Old version backed up as monthly report sample.txt.backup")
                print("📄 旧版本已备份为monthly report sample.txt.backup")
            else:
                print("⏭️ Skipped file replacement")
                print("⏭️ 跳过文件替换")
                
        except Exception as e:
            print(f"❌ Error during file replacement: {e}")
            print(f"❌ 文件替换过程中出错: {e}")
    
    def run_workflow(self):
        """Main workflow execution"""
        steps = {
            1: self.step_1,
            2: self.step_2,
            3: self.step_3,
            4: self.step_4,
            5: self.step_5
        }
        
        print("🚀 Starting CAR.org Report Download Workflow")
        print("🚀 开始CAR.org报告下载工作流程")
        print("Use navigation controls to move between steps")
        print("使用导航控制在步骤之间移动")
        sleep(2)
        
        while True:
            # Execute current step
            if self.current_step in steps:
                should_exit = steps[self.current_step]()
                if should_exit:
                    break  # Exit the workflow
            
            # Show navigation and get user input
            self.show_navigation()
            choice = self.get_user_input()
            
            # Handle navigation
            if choice == 'q':
                print("\n👋 Exiting workflow. Progress saved.")
                print("\n👋 退出工作流程。进度已保存。")
                break
            elif choice == 'b' and self.current_step > 1:
                self.current_step -= 1
            elif choice == 'n' and self.current_step < self.total_steps:
                self.current_step += 1
            elif choice == 'r':
                continue  # Repeat current step
            elif choice == 'c':
                self.completed_steps.add(self.current_step)
                print(f"✅ Step {self.current_step} marked as completed!")
                print(f"✅ 步骤{self.current_step}已标记为完成!")
                if self.current_step < self.total_steps:
                    self.current_step += 1
                sleep(1)
            elif choice == 's':
                self.show_progress_summary()
            

def main():
    """Main execution"""
    try:
        workflow = WorkflowGuide()
        workflow.run_workflow()
    except KeyboardInterrupt:
        print("\n\n👋 Workflow interrupted. You can restart anytime by running this script again.")
        print("\n👋 工作流程中断。您可以随时通过再次运行此脚本重新启动。")
    except Exception as e:
        print(f"\n❌ Unexpected error: {e}")
        print(f"\n❌ 意外错误: {e}")

if __name__ == "__main__":
    main()