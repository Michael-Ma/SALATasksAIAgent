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
        self.total_steps = 4  # Reduced from 5 to 4 (combined old steps 2 and 3)
        self.work_directory = self.setup_work_directory()
        # Include full sharing token so API fallback can enumerate
        self.sharepoint_folder_url = "https://carorg.sharepoint.com/:f:/s/CAR-RE-PublicProducts/ElrCKkQh6_ZMpe5RIZgOohoB33WDC9L1NlkigRWlqWwvGg?e=Vi1XJZ"

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
                print(f"📁 已创建工作目录: {work_dir}")
            else:
                print(f"📁 使用现有工作目录: {work_dir}")

            return work_dir

        except Exception as e:
            print(f"❌ 创建工作目录时出错: {e}")
            # Fallback to current directory
            return os.getcwd()

    def clear_screen(self):
        """Do not clear screen to retain logs for debugging"""
        return

    def show_step_menu(self):
        """Display all steps and allow jumping to specific step"""
        self.clear_screen()
        print("="*60)
        print("🚀 CAR.org 报告下载工作流程")
        print("="*60)
        print(f"📁 工作目录: {self.work_directory}")
        print("="*60)
        print()

        steps_info = [
            "MLSL BI门户任务（登录、报告、邮件、转换）",
            "从SharePoint下载4个县的PNG文件",
            "运行Power BI下载器获取所有4个县和市级图像",
            "使用AI生成更新的月度报告"
        ]

        print("📋 工作流程步骤:")
        print("─"*60)
        for i, step_desc in enumerate(steps_info, 1):
            marker = "👉" if i == self.current_step else "  "
            print(f"{marker} [{i}] {step_desc}")

        print()
        print("─"*60)
        print("导航选项:")
        print("  [1-4] 跳转到指定步骤")
        print("  [Enter] 继续当前步骤")
        print("  [n] 下一步")
        print("  [b] 上一步")
        print("  [q] 退出")
        print("─"*60)

    def get_step_choice(self):
        """Get user's step choice"""
        while True:
            choice = input("\n输入您的选择: ").lower().strip()

            if choice == '':
                return self.current_step  # Continue with current step
            elif choice == 'n':
                if self.current_step < self.total_steps:
                    return self.current_step + 1
                else:
                    print("❌ 已经是最后一步")
                    continue
            elif choice == 'b':
                if self.current_step > 1:
                    return self.current_step - 1
                else:
                    print("❌ 已经是第一步")
                    continue
            elif choice == 'q':
                return 'q'
            elif choice.isdigit() and 1 <= int(choice) <= self.total_steps:
                return int(choice)
            else:
                print(f"❌ 无效选择。请输入 1-{self.total_steps}, n, b, 或 q")

    def show_header(self):
        """Display current step header"""
        self.clear_screen()
        print("="*60)
        print(f"步骤 {self.current_step}/{self.total_steps}")
        print(f"📁 工作目录: {self.work_directory}")
        print("="*60)
        print()

    def step_1(self):
        """Step 1: Download PNG files from CAR.org and SharePoint"""
        self.show_header()

        print("📋 步骤 1: 从SharePoint下载4个县的PNG文件")
        print("─"*60)
        print("🎯 目标: 下载所有4个县的PNG文件")
        print("   • Santa Clara County")
        print("   • San Mateo County")
        print("   • Alameda County")
        print("   • San Francisco County")
        print()
        print("🔗 SharePoint网址: https://carorg.sharepoint.com/:f:/s/CAR-RE-PublicProducts/ElrCKkQh6_ZMpe5RIZgOohoB33WDC9L1NlkigRWlqWwvGg?e=Vi1XJZ")
        print()

        # Offer automated SharePoint download attempt
        if input("🤖  尝试自动从SharePoint下载所有4个县的PNG文件? (y/n): ").lower().strip() == 'y':
            if self.auto_download_step1():
                input("\n⏸️  自动下载完成。按Enter继续...")
                return
            else:
                print("\n⚠️ 自动下载未成功，继续手动步骤。")

        if input("🔗 打开CAR.org营销页面? (y/n): ").lower() == 'y':
            try:
                webbrowser.open("https://www.car.org/marketing/chartsandgraphs/marketupdate")
                print("✅ 已在浏览器中打开CAR.org营销页面")
            except Exception as e:
                print(f"❌ 打开浏览器时出错: {e}")
                print("请手动打开: https://www.car.org/marketing/chartsandgraphs/marketupdate")

        print("\n🔗 需要SharePoint备选方案吗？")
        if input("打开SharePoint文件夹? (y/n): ").lower() == 'y':
            try:
                webbrowser.open("https://carorg.sharepoint.com/:f:/s/CAR-RE-PublicProducts/ElrCKkQh6_ZMpe5RIZgOohoB33WDC9L1NlkigRWlqWwvGg?e=Vi1XJZ")
                print("✅ 已在浏览器中打开SharePoint文件夹")
            except Exception as e:
                print(f"❌ 打开浏览器时出错: {e}")
                print("请手动打开SharePoint网址")

        input("\n⏸️  完成后按Enter继续...")

    def auto_download_step1(self):
        """Attempt to auto-download the four county PNGs from the SharePoint folder"""
        try:
            from auto_download_county_file_sharepoint import download_file_from_sharepoint
        except ImportError:
            print("❌ 未找到download_alameda_sharepoint模块")
            return False

        # Try to download all 4 county PNG files
        target_files = [
            "Santa Clara.png",
            "San Mateo.png",
            "Alameda.png",
            "San Francisco.png"
        ]

        downloaded_files = []
        for filename in target_files:
            print(f"\n{'='*60}")
            print(f"📥 正在下载: {filename}")
            print(f"{'='*60}")

            result = download_file_from_sharepoint(
                self.sharepoint_folder_url,
                filename,
                self.work_directory
            )

            if result:
                downloaded_files.append(result)
                print(f"✅ 成功下载: {filename}")
            else:
                print(f"❌ 下载失败: {filename}")

        # Summary
        print(f"\n{'='*60}")
        print(f"📊 下载摘要")
        print(f"{'='*60}")
        print(f"目标: {len(target_files)} 个文件")
        print(f"成功: {len(downloaded_files)} 个文件")

        if downloaded_files:
            print(f"\n✅ 已下载文件:")
            for f in downloaded_files:
                print(f"   • {os.path.basename(f)}")

        return len(downloaded_files) > 0

    def step_2(self):
        """Step 2: Run Power BI downloader (manual by default, optional auto county selection)"""
        self.show_header()

        print("📋 步骤 2: 运行Power BI下载器获取县和市级图像")
        print("─"*60)
        print("🚀 下载Power BI中的县和市级数据")
        print()

        # Check scripts
        manual_exists = os.path.exists("final_working_downloader.py")
        auto_exists = os.path.exists("final_working_downloader_auto.py")

        if not manual_exists:
            print("❌ 找不到final_working_downloader.py!")
            print("请确保脚本在当前目录中。")
            input("\n⏸️  按Enter继续...")
            return

        print(f"✅ 手动脚本: {'已找到' if manual_exists else '缺失'}")
        print(f"✅ 自动脚本: {'已找到' if auto_exists else '缺失'}")
        print()
        print("📝 工作流程:")
        print("─" * 30)
        print("1. 脚本打开CAR.org Power BI页面")
        print("2. 选择模式:")
        print("   • 手动: 您从下拉菜单中选择每个县")
        print("   • 自动: 脚本尝试自动选择县（如DOM变化可能失效）")
        print("3. 脚本提取并下载县和市级图像")
        print("4. 保存文件格式: 1.png, 1(1).png, 1(2).png等")
        print()
        print("🎯 目标县:")
        print("   → Santa Clara (文件: 1.png, 1(1).png, 1(2).png...)")
        print("   → San Mateo (文件: 2.png, 2(1).png, 2(2).png...)")
        print("   → Alameda (文件: 3.png, 3(1).png, 3(2).png...)")
        print("   → San Francisco (文件: 4.png, 4(1).png, 4(2).png...)")
        print()

        mode = 'm'
        if auto_exists:
            mode = input("选择下载模式 [M 手动 / A 自动] (默认M): ").strip().lower() or 'm'
        use_auto = auto_exists and mode == 'a'
        script_name = "final_working_downloader_auto.py" if use_auto else "final_working_downloader.py"

        print(f"🚀 启动下载器脚本 ({'自动选择县' if use_auto else '手动选择县'})...")
        try:
            print("▶️  正在启动...")
            print(f"📁 文件将保存到: {self.work_directory}")

            # Set environment variable for work directory
            env = os.environ.copy()
            env['SALA_WORK_DIR'] = self.work_directory

            result = subprocess.run([sys.executable, script_name],
                                  cwd=os.getcwd(), env=env)
            if result.returncode == 0:
                print("✅ 下载脚本成功完成!")
            else:
                print("⚠️  脚本完成但有警告/错误")
        except Exception as e:
            print(f"❌ 运行脚本时出错: {e}")
            print(f"您可以手动运行: python {script_name}")

        input("\n⏸️  完成后按Enter继续...")

    def step_3(self):
        """Step 3: MLSL BI Portal Tasks"""
        self.show_header()

        print("📋 步骤 3: MLSL BI门户任务")
        print("─"*60)
        print("🎯 完成所有MLSL BI门户相关任务")
        print()

        print("📝 操作步骤 (摘要):")
        print("─" * 30)
        print("1. 🔗 打开 https://mlsl.aculist.com/BI 完成登录")
        print("2. 📧 下载ZIP报告PDF到工作目录")
        print("3. 🤖 如需自动转换，运行内置转换器并删除PDF")

        print("💡 提示:")
        print("• 寻找'报告'、'收藏'或'我的报告'部分")
        print("• 使用在线转换器如PDF转PNG进行转换")
        print()
        
        # Optional: open the MLSL BI portal (only if user requests)
        if input("🔗 需要现在打开 MLSL BI 门户吗? (y/n): ").lower().strip() == 'y':
            try:
                webbrowser.open("https://mlsl.aculist.com/BI")
                print("✅ 已在浏览器中打开MLSL BI门户")
            except Exception as e:
                print(f"❌ 打开浏览器时出错: {e}")
                print("请手动打开: https://mlsl.aculist.com/BI")

        # Optional: auto convert and rename ZIP PDFs to PNGs
        if input("🤖 需要自动将下载的 ZIP PDF 转成 6(x).png 并重命名吗? (y/n): ").lower().strip() == 'y':
            try:
                env = os.environ.copy()
                env['SALA_WORK_DIR'] = self.work_directory
                print("▶️  正在运行 convert_zip_pdfs_to_png.py ...")
                result = subprocess.run([sys.executable, "convert_zip_pdfs_to_png.py"],
                                        cwd=os.getcwd(), env=env)
                if result.returncode == 0:
                    print("✅ PDF 转 PNG 完成。")
                else:
                    print("⚠️ PDF 转换脚本返回非零状态，请检查输出。")
            except Exception as e:
                print(f"❌ 运行 PDF 转换脚本时出错: {e}")

        input("\n⏸️  完成所有任务后按Enter继续...")

    def step_4(self):
        """Step 4: Generate updated monthly report using AI"""
        self.show_header()

        print("📋 步骤 4: 使用AI生成更新的月度报告")
        print("─"*60)
        print("🤖 使用AI从下载的数据创建更新的月度报告")
        print()
        print("📝 报告生成过程:")
        print("─" * 30)
        print("1. 📊 脚本分析所有下载的PNG文件")
        print("2. 🔄 将当前数据与上月报告进行比较")
        print("3. 📝 自动更新所有数字和百分比")
        print("4. 🗓️  所有月份增加+1 (7月 → 8月)")
        print("5. 🇨🇳 使用正确的中文变化术语 (上升/下降)")
        print("6. 💾 将完整的更新报告保存到工作目录")
        print()
        print("🎯 要求:")
        print("• OpenAI API密钥 (设置为OPENAI_API_KEY环境变量)")
        print("• 工作目录中的PNG文件 (来自之前的步骤)")
        print("• 原始月度报告样本文件")
        print()

        print("🚀 自动启动simple_report_updater.py...")
        try:
            print("▶️  启动AI报告更新器...")
            print(f"📁 使用工作目录: {self.work_directory}")

            # Set environment variable for work directory
            env = os.environ.copy()
            env['SALA_WORK_DIR'] = self.work_directory

            result = subprocess.run([sys.executable, "simple_report_updater.py"],
                                  cwd=os.getcwd(), env=env)
            if result.returncode == 0:
                print("\n✅ AI报告生成成功完成!")
                print(f"📄 已生成: {self.work_directory}/monthly_report_updated.txt")

                # Ask permission to replace the sample file
                self.offer_to_replace_sample_report()

                print("\n" + "="*60)
                print("🎉 工作流程成功完成!")
                print("="*60)
                print("所有4个步骤已完成:")
                print("✅ 从CAR.org和SharePoint下载了PNG文件")
                print("✅ 通过Power BI下载了县和市级图像")
                print("✅ 完成了MLSL BI门户任务")
                print("✅ 使用AI生成了更新的月度报告")
                print()
                print("🏁 所有报告下载和生成任务已完成!")
                print(f"📁 最终文件保存到: {self.work_directory}")

                return True  # Signal workflow completion

            else:
                print("⚠️  报告生成完成但有警告/错误")

        except Exception as e:
            print(f"❌ 运行AI报告更新器时出错: {e}")
            print("您可以手动运行: python simple_report_updater.py")

        input("\n⏸️  按Enter继续...")
        return False

    def offer_to_replace_sample_report(self):
        """Offer to replace sample report with updated version"""
        try:
            updated_file = os.path.join(self.work_directory, "monthly_report_updated.txt")
            sample_file = "monthly report sample.txt"

            if not os.path.exists(updated_file):
                print("❌ 找不到更新的报告，跳过替换")
                return

            if not os.path.exists(sample_file):
                print("❌ 找不到样本报告，跳过替换")
                return

            print("\n" + "="*60)
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

            print("🔍 逐行差异预览:")
            print("─" * 80)

            # Find meaningful differences (skip empty lines and minor changes)
            differences_shown = 0
            max_differences = 10  # Limit to avoid overwhelming output

            for i, (old_line, new_line) in enumerate(zip(old_lines, new_lines)):
                if old_line.strip() != new_line.strip() and old_line.strip() and new_line.strip():
                    if differences_shown >= max_differences:
                        print("  ... (更多差异未显示)")
                        break

                    print(f"第{i+1}行:")
                    print(f"  旧: {old_line}")
                    print(f"  新: {new_line}")
                    print("─" * 40)
                    differences_shown += 1

            # Handle case where new file has more lines
            if len(new_lines) > len(old_lines):
                print(f"新文件多出 {len(new_lines) - len(old_lines)} 行")
            elif len(old_lines) > len(new_lines):
                print(f"新文件少了 {len(old_lines) - len(new_lines)} 行")

            if differences_shown == 0:
                print("未发现重要内容差异")

            print(f"\n📊 文件比较:")
            print(f"  旧文件: {len(old_lines)} 行, {len(old_content)} 字符")
            print(f"  新文件: {len(new_lines)} 行, {len(new_content)} 字符")

            # Ask for permission
            print("\n🔄 用更新版本替换'monthly report sample.txt'吗?")
            if input("输入'yes'进行替换: ").lower() in ['yes', 'y']:
                # Make backup first
                backup_file = "monthly report sample.txt.backup"
                with open(backup_file, 'w', encoding='utf-8') as f:
                    f.write(old_content)
                print(f"📄 备份保存为: {backup_file}")

                # Replace the file
                with open(sample_file, 'w', encoding='utf-8') as f:
                    f.write(new_content)

                print("✅ 成功替换monthly report sample.txt")
                print("📄 旧版本已备份为monthly report sample.txt.backup")
            else:
                print("⏭️  跳过文件替换")

        except Exception as e:
            print(f"❌ 文件替换过程中出错: {e}")

    def run_workflow(self):
        """Main workflow execution"""
        steps = {
            1: self.step_3,  # MLSL BI Portal (moved to first)
            2: self.step_1,  # SharePoint downloads
            3: self.step_2,  # Power BI downloader
            4: self.step_4   # AI report generation
        }

        print("🚀 开始CAR.org报告下载工作流程")
        sleep(1)

        while True:
            # Show step menu
            self.show_step_menu()

            # Get user's choice
            choice = self.get_step_choice()

            if choice == 'q':
                print("\n👋 退出工作流程")
                break

            # Update current step
            self.current_step = choice

            # Execute the step
            if self.current_step in steps:
                completed = steps[self.current_step]()
                if completed:
                    # Workflow finished
                    break

def main():
    """Main execution"""
    try:
        workflow = WorkflowGuide()
        workflow.run_workflow()
    except KeyboardInterrupt:
        print("\n\n👋 工作流程中断。您可以随时通过再次运行此脚本重新启动。")
    except Exception as e:
        print(f"\n❌ 意外错误: {e}")

if __name__ == "__main__":
    main()
