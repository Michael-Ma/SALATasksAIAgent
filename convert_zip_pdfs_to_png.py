"""
Convert ZIP PDF reports to sequential PNGs.
- Searches for ZIPTrends...*.pdf in SALA work directory (or current dir)
- Converts each page to PNG using PyMuPDF
- Saves as 6(1).png, 6(2).png, ... matching MLSL naming pattern
"""

import os
import re
import sys
from typing import List

import fitz  # PyMuPDF


def find_pdfs(work_dir: str) -> List[str]:
    """Find ZIPTrends PDFs in the work directory, sorted."""
    candidates = []
    for name in os.listdir(work_dir):
        if not name.lower().endswith(".pdf"):
            continue
        if "ziptrends" in name.lower():
            candidates.append(os.path.join(work_dir, name))
    return sorted(candidates)


def convert_pdf_to_pngs(pdf_path: str, work_dir: str, start_index: int) -> int:
    """
    Convert a single PDF to PNG(s).
    Returns number of files written.
    """
    written = 0
    with fitz.open(pdf_path) as doc:
        for page_num, page in enumerate(doc, 1):
            pix = page.get_pixmap(dpi=200)
            filename = f"6({start_index + written}).png"
            out_path = os.path.join(work_dir, filename)
            pix.save(out_path)
            written += 1
            print(f"   ✅ {os.path.basename(pdf_path)} page {page_num} -> {filename}")
    return written


def main():
    work_dir = os.getenv("SALA_WORK_DIR") or os.getcwd()
    print(f"📁 工作目录: {work_dir}")

    if not os.path.isdir(work_dir):
        print("❌ 工作目录无效")
        sys.exit(1)

    pdfs = find_pdfs(work_dir)
    if not pdfs:
        print("❌ 未找到 ZIPTrends PDF 文件")
        sys.exit(1)

    print(f"🔍 找到 {len(pdfs)} 个PDF文件，开始转换...")

    total_written = 0
    for pdf in pdfs:
        written = convert_pdf_to_pngs(pdf, work_dir, start_index=total_written + 1)
        total_written += written
        if written > 0:
            try:
                os.remove(pdf)
                print(f"   🗑️ 已删除原PDF: {os.path.basename(pdf)}")
            except Exception as e:
                print(f"   ⚠️ 删除PDF失败 {os.path.basename(pdf)}: {e}")

    if total_written:
        print(f"🎉 转换完成，共生成 {total_written} 个PNG，命名 6(1).png 起。")
        sys.exit(0)
    else:
        print("⚠️ 未生成任何PNG。")
        sys.exit(1)


if __name__ == "__main__":
    main()
