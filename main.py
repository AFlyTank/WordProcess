#GH 2025/10/24/4:01 version-1
import os
import re
import shutil
import tkinter as tk
from tkinter import filedialog, messagebox
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx2pdf import convert  # 关键：docx2pdf依赖
from docx.shared import Pt

def select_folder():
    """选择要处理的文件夹"""
    root = tk.Tk()
    root.withdraw()
    folder_path = filedialog.askdirectory(title="选择要处理的文件夹")
    return folder_path


def get_all_word_files(folder_path):
    """获取文件夹中所有Word文件（包括子文件夹）"""
    word_files = []
    for root, dirs, files in os.walk(folder_path):
        for file in files:
            if file.lower().endswith(('.docx',)):
                word_files.append(os.path.join(root, file))
    return word_files

def remove_header_footer(doc):
    """删除页眉页脚所有内容（包括段落标记，彻底清空）"""
    for section in doc.sections:
        # 处理页眉：彻底删除所有段落（含标记）
        header = section.header
        # 反向遍历删除，避免索引混乱
        for para in reversed(header.paragraphs):
            # 获取段落XML元素的父节点，直接删除段落
            p_element = para._element
            parent = p_element.getparent()
            if parent is not None:
                parent.remove(p_element)
            # 清空段落引用，避免残留
            para._p = None
            para._element = None

        # 处理页脚：彻底删除所有段落（含标记）
        footer = section.footer
        for para in reversed(footer.paragraphs):
            p_element = para._element
            parent = p_element.getparent()
            if parent is not None:
                parent.remove(p_element)
            para._p = None
            para._element = None


def add_custom_header(doc):
    """添加页眉：学为人师，行为师范（左对齐，华文行楷，小四）"""
    for section in doc.sections:
        header = section.header  # 获取当前节的页眉

        # 新建一个段落（页眉已被清空，直接添加）
        para = header.add_paragraph()

        # 设置段落左对齐
        para.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT

        # 添加页眉文字
        run = para.add_run("学为人师，行为世范")

        # 设置字体为华文行楷（兼容Windows）
        run.font.name = "华文行楷"
        # 强制指定中文字体（关键）
        run._element.rPr.rFonts.set(qn('w:eastAsia'), "华文行楷")

        # 设置字号为小四（12磅）
        run.font.size = Pt(12)

def add_centered_page_number(doc):
    """在清空后的页脚添加居中页码（无空行/无格式）"""
    for section in doc.sections:
        footer = section.footer

        # 此时页脚已彻底清空，直接新建唯一段落
        p = footer.add_paragraph()

        # 清除段落默认格式（关键：解决首行缩进/空行）
        para_format = p.paragraph_format
        para_format.first_line_indent = 0  # 取消首行缩进
        para_format.left_indent = 0        # 取消左缩进
        para_format.right_indent = 0       # 取消右缩进
        para_format.space_before = 0       # 取消段前距
        para_format.space_after = 0        # 取消段后距
        para_format.alignment = WD_ALIGN_PARAGRAPH.CENTER  # 强制居中

        # 添加页码字段（保持原有XML操作逻辑）
        run = p.add_run()
        fld_char_begin = OxmlElement('w:fldChar')
        fld_char_begin.set(qn('w:fldCharType'), 'begin')
        run._r.append(fld_char_begin)

        instr_text = OxmlElement('w:instrText')
        instr_text.text = "PAGE"
        run._r.append(instr_text)

        fld_char_end = OxmlElement('w:fldChar')
        fld_char_end.set(qn('w:fldCharType'), 'end')
        run._r.append(fld_char_end)

import re

import re


def replace_patterns_in_paragraph(paragraph):
    """修正版：按字符位置精准还原，解决报错问题"""
    # 1. 记录所有文本run的原始信息（字符起止位置、文本内容）
    text_runs = []  # 存储：(run对象, 原始文本, 字符起始位置, 字符结束位置)
    char_pos = 0  # 全局字符计数器

    for run in paragraph.runs:
        if not run.text:  # 跳过非文本run（公式、图片）
            continue
        text = run.text
        start = char_pos
        end = char_pos + len(text)
        text_runs.append((run, text, start, end))
        char_pos = end  # 更新全局位置

    if not text_runs:
        return  # 无文本内容直接返回

    # 2. 合并所有文本，执行替换并记录被删除的字符位置
    all_text = ''.join([t[1] for t in text_runs])
    english_pattern = r'\(([12][0-9]).{3,}?\)'
    chinese_pattern = r'（([12][0-9]).{3,}?）'

    # 记录所有被替换的字符区间（start, end）
    replaced_ranges = []

    # 修正：用列表.append()记录，避免变量作用域问题
    def mark_replaced(match):
        replaced_ranges.append((match.start(), match.end()))
        return ""

    # 先替换中文括号，再替换英文括号
    new_text = re.sub(chinese_pattern, mark_replaced, all_text)
    new_text = re.sub(english_pattern, mark_replaced, new_text)

    # 3. 计算替换后每个位置的字符是否保留（生成保留掩码）
    keep_mask = [True] * len(all_text)  # True=保留，False=删除
    for start, end in replaced_ranges:
        for i in range(start, end):
            if i < len(keep_mask):
                keep_mask[i] = False  # 标记被删除的字符

    # 4. 按原始run的字符范围，提取保留的字符并更新run
    for run, original_text, start, end in text_runs:
        kept_chars = []
        for i in range(start, end):  # 遍历当前run包含的字符位置
            if i < len(keep_mask) and keep_mask[i]:
                # 计算在原始文本中的索引（i - 起始位置 = 原始文本中的下标）
                original_idx = i - start
                kept_chars.append(original_text[original_idx])
        # 更新当前run的文本
        run.text = ''.join(kept_chars)


def process_word_file(file_path, root_folder):
    """处理单个Word文件（保留图片和公式）"""
    try:
        # 关键：使用原文档的XML结构，避免重新生成导致元素丢失
        doc = Document(file_path)

        # 1. 删除页眉页脚内容并添加页码
        remove_header_footer(doc)
        add_centered_page_number(doc)
        add_custom_header(doc)

        # 2. 替换文本（保留非文本元素）
        # 处理普通段落
        for para in doc.paragraphs:
            replace_patterns_in_paragraph(para)

        # 处理表格中的段落
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for para in cell.paragraphs:
                        replace_patterns_in_paragraph(para)

        # 保存到processed文件夹
        relative_path = os.path.relpath(file_path, root_folder)
        processed_root = os.path.join(root_folder, "processed")
        target_path = os.path.join(processed_root, relative_path)
        os.makedirs(os.path.dirname(target_path), exist_ok=True)
        doc.save(target_path)
        return True, target_path
    except Exception as e:
        return False, f"处理文件 {file_path} 时出错: {str(e)}"


def copy_non_word_files(folder_path):
    """复制非Word文件到processed文件夹，PDF文件特殊处理"""
    processed_root = os.path.join(folder_path, "processed")
    for root, dirs, files in os.walk(folder_path):
        if root.startswith(processed_root):
            continue
        for file in files:
            if not file.lower().endswith(('.docx',)):  # 只处理非Word文件
                src_path = os.path.join(root, file)
                relative_path = os.path.relpath(src_path, folder_path)
                target_path = os.path.join(processed_root, relative_path)
                os.makedirs(os.path.dirname(target_path), exist_ok=True)

                # 新增：判断是否为PDF
                if file.lower().endswith('.pdf'):
                    # 生成对应Word文件名（同路径同名称）
                    word_filename = os.path.splitext(file)[0] + '.docx'
                    word_relative_path = os.path.join(os.path.relpath(root, folder_path), word_filename)
                    word_target_path = os.path.join(processed_root, word_relative_path)

                    # 检查目标位置是否有同名Word文件
                    if os.path.exists(word_target_path):
                        # 有则用Word生成PDF，不复制原PDF
                        pdf_target_path = os.path.splitext(word_target_path)[0] + '.pdf'
                        try:
                            convert(word_target_path, pdf_target_path)  # 生成PDF
                            continue  # 跳过复制原PDF
                        except:
                            pass  # 生成失败则 fallback 到复制原PDF
                # 非PDF文件 或 无对应Word 或 生成失败，正常复制
                shutil.copy2(src_path, target_path)


def main():
    root = tk.Tk()
    root.title("Word文件处理器")
    root.geometry("500x300")

    folder_var = tk.StringVar()
    folder_var.set("等待选择文件夹...")
    folder_label = tk.Label(root, textvariable=folder_var, wraplength=450)
    folder_label.pack(pady=10)

    def select_and_display_folder():
        folder = select_folder()
        if folder:
            folder_var.set(f"已选择文件夹: {folder}")

    select_btn = tk.Button(root, text="选择文件夹", command=select_and_display_folder)
    select_btn.pack(pady=5)

    status_var = tk.StringVar()
    status_var.set("等待处理...")
    status_label = tk.Label(root, textvariable=status_var, wraplength=450)
    status_label.pack(pady=20)

    def process_files():
        folder_path = folder_var.get().replace("已选择文件夹: ", "")
        if not folder_path or folder_path == "等待选择文件夹...":
            messagebox.showwarning("警告", "请先选择文件夹")
            return

        processed_folder = os.path.join(folder_path, "processed")
        if os.path.exists(processed_folder):
            if not messagebox.askyesno("提示", "processed文件夹已存在，是否清空后重新处理？"):
                return
            shutil.rmtree(processed_folder)
        os.makedirs(processed_folder, exist_ok=True)

        word_files = get_all_word_files(folder_path)
        total_word = len(word_files)
        success_word = 0
        error_messages = []

        for i, file_path in enumerate(word_files):
            status_var.set(f"正在处理Word文件 {i + 1}/{total_word}: {os.path.basename(file_path)}")
            root.update_idletasks()

            success, result = process_word_file(file_path, folder_path)
            if success:
                success_word += 1
            else:
                error_messages.append(result)

        # status_var.set("正在复制非Word文件...")
        # root.update_idletasks()
        # copy_non_word_files(folder_path)

        result_msg = f"处理完成！\nWord文件处理: 成功 {success_word}/{total_word}\n"
        result_msg += f"所有文件已保存到: {processed_folder}\n"
        if error_messages:
            result_msg += "\n错误信息:\n" + "\n".join(error_messages)
        messagebox.showinfo("处理结果", result_msg)
        status_var.set("处理完成")

    process_btn = tk.Button(root, text="开始处理", command=process_files)
    process_btn.pack(pady=10)

    exit_btn = tk.Button(root, text="退出", command=lambda: [root.destroy(), os._exit(0)])
    exit_btn.pack(pady=10)

    root.mainloop()


if __name__ == "__main__":
    main()