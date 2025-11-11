import os
import re
import shutil
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx2pdf import convert
from docx.shared import Pt


def select_folder():
    """选择要处理的文件夹"""
    root = tk.Tk()
    root.withdraw()  # 隐藏主窗口
    folder_path = filedialog.askdirectory(title="选择要处理的文件夹")
    return folder_path if folder_path else None  # 未选择时返回None


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
        for para in reversed(header.paragraphs):
            p_element = para._element
            parent = p_element.getparent()
            if parent is not None:
                parent.remove(p_element)
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
        header = section.header
        para = header.add_paragraph()
        para.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
        run = para.add_run("学为人师，行为世范")
        run.font.name = "华文行楷"
        run._element.rPr.rFonts.set(qn('w:eastAsia'), "华文行楷")
        run.font.size = Pt(12)


def add_centered_page_number(doc):
    """在清空后的页脚添加居中页码（无空行/无格式）"""
    for section in doc.sections:
        footer = section.footer
        p = footer.add_paragraph()
        para_format = p.paragraph_format
        para_format.first_line_indent = 0
        para_format.left_indent = 0
        para_format.right_indent = 0
        para_format.space_before = 0
        para_format.space_after = 0
        para_format.alignment = WD_ALIGN_PARAGRAPH.CENTER

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


def replace_patterns_in_paragraph(paragraph):
    """修正版：按字符位置精准还原，解决报错问题"""
    text_runs = []
    char_pos = 0

    for run in paragraph.runs:
        if not run.text:
            continue
        text = run.text
        start = char_pos
        end = char_pos + len(text)
        text_runs.append((run, text, start, end))
        char_pos = end

    if not text_runs:
        return

    all_text = ''.join([t[1] for t in text_runs])
    english_pattern = r'\(([12][0-9]).{3,}?\)'
    chinese_pattern = r'（([12][0-9]).{3,}?）'
    replaced_ranges = []

    def mark_replaced(match):
        replaced_ranges.append((match.start(), match.end()))
        return ""

    re.sub(chinese_pattern, mark_replaced, all_text)
    re.sub(english_pattern, mark_replaced, all_text)

    keep_mask = [True] * len(all_text)
    for start, end in replaced_ranges:
        for i in range(start, end):
            if i < len(keep_mask):
                keep_mask[i] = False

    for run, original_text, start, end in text_runs:
        kept_chars = []
        for i in range(start, end):
            if i < len(keep_mask) and keep_mask[i]:
                original_idx = i - start
                kept_chars.append(original_text[original_idx])
        run.text = ''.join(kept_chars)


def process_word_file(file_path, root_folder, options):
    """处理单个Word文件（根据选项执行相应功能）"""
    try:
        doc = Document(file_path)

        # 根据选项处理页眉页脚
        if options['remove_header_footer']:
            remove_header_footer(doc)

        if options['add_page_number']:
            add_centered_page_number(doc)

        if options['add_custom_header']:
            add_custom_header(doc)

        # 根据选项替换文本
        if options['replace_patterns']:
            for para in doc.paragraphs:
                replace_patterns_in_paragraph(para)
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for para in cell.paragraphs:
                            replace_patterns_in_paragraph(para)

        # 确定保存路径
        if options['use_processed_folder']:
            relative_path = os.path.relpath(file_path, root_folder)
            processed_root = os.path.join(root_folder, "processed")
            target_path = os.path.join(processed_root, relative_path)
            os.makedirs(os.path.dirname(target_path), exist_ok=True)
        else:
            # 备份原文件
            backup_path = f"{file_path}.bak"
            if not os.path.exists(backup_path):
                shutil.copy2(file_path, backup_path)
            target_path = file_path

        doc.save(target_path)
        return True, target_path
    except Exception as e:
        return False, f"处理文件 {file_path} 时出错: {str(e)}"


def copy_non_word_files(folder_path, use_processed_folder):
    """复制非Word文件到指定位置"""
    if use_processed_folder:
        processed_root = os.path.join(folder_path, "processed")
    else:
        processed_root = folder_path

    for root, dirs, files in os.walk(folder_path):
        if use_processed_folder and root.startswith(processed_root):
            continue
        for file in files:
            if not file.lower().endswith(('.docx',)):
                src_path = os.path.join(root, file)
                relative_path = os.path.relpath(src_path, folder_path)
                target_path = os.path.join(processed_root, relative_path)
                os.makedirs(os.path.dirname(target_path), exist_ok=True)

                if file.lower().endswith('.pdf') and use_processed_folder:
                    word_filename = os.path.splitext(file)[0] + '.docx'
                    word_relative_path = os.path.join(os.path.relpath(root, folder_path), word_filename)
                    word_target_path = os.path.join(processed_root, word_relative_path)

                    if os.path.exists(word_target_path):
                        pdf_target_path = os.path.splitext(word_target_path)[0] + '.pdf'
                        try:
                            convert(word_target_path, pdf_target_path)
                            continue
                        except:
                            pass
                shutil.copy2(src_path, target_path)


def main():
    root = tk.Tk()
    root.title("Word文件处理器")
    root.geometry("600x500")

    # 创建选项变量
    options = {
        'remove_header_footer': tk.BooleanVar(value=True),
        'add_custom_header': tk.BooleanVar(value=True),
        'add_page_number': tk.BooleanVar(value=True),
        'replace_patterns': tk.BooleanVar(value=True),
        'use_processed_folder': tk.BooleanVar(value=True),
        'copy_non_word': tk.BooleanVar(value=True)
    }

    folder_var = tk.StringVar()
    folder_var.set("等待选择文件夹...")

    # 创建UI组件
    frame = ttk.Frame(root, padding="20")
    frame.pack(fill=tk.BOTH, expand=True)

    # 文件夹选择区域
    ttk.Label(frame, text="选择处理文件夹:", font=("Arial", 10, "bold")).pack(anchor=tk.W, pady=(0, 5))
    ttk.Label(frame, textvariable=folder_var, wraplength=550).pack(anchor=tk.W, pady=(0, 15))

    # 修复：只调用一次select_folder()并缓存结果
    def select_and_display_folder():
        folder = select_folder()
        if folder:  # 只有选择了文件夹才更新显示
            folder_var.set(f"已选择文件夹: {folder}")

    select_btn = ttk.Button(frame, text="选择文件夹", command=select_and_display_folder)
    select_btn.pack(anchor=tk.W, pady=(0, 20))

    # 功能选项区域
    ttk.Label(frame, text="处理选项:", font=("Arial", 10, "bold")).pack(anchor=tk.W, pady=(0, 5))

    options_frame = ttk.Frame(frame)
    options_frame.pack(fill=tk.X, pady=(0, 15))

    # 第一列选项
    col1 = ttk.Frame(options_frame)
    col1.pack(side=tk.LEFT, fill=tk.X, expand=True)

    ttk.Checkbutton(
        col1, text="删除页眉页脚", variable=options['remove_header_footer']
    ).pack(anchor=tk.W, pady=2)

    ttk.Checkbutton(
        col1, text="添加自定义页眉", variable=options['add_custom_header']
    ).pack(anchor=tk.W, pady=2)

    # 第二列选项
    col2 = ttk.Frame(options_frame)
    col2.pack(side=tk.LEFT, fill=tk.X, expand=True)

    ttk.Checkbutton(
        col2, text="添加居中页码", variable=options['add_page_number']
    ).pack(anchor=tk.W, pady=2)

    ttk.Checkbutton(
        col2, text="替换指定文本模式", variable=options['replace_patterns']
    ).pack(anchor=tk.W, pady=2)

    # 保存选项
    ttk.Label(frame, text="保存选项:", font=("Arial", 10, "bold")).pack(anchor=tk.W, pady=(0, 5))

    ttk.Checkbutton(
        frame,
        text="使用processed文件夹保存(不勾选则替换原文件，原文件会备份为.bak)",
        variable=options['use_processed_folder']
    ).pack(anchor=tk.W, pady=2)

    ttk.Checkbutton(
        frame, text="复制非Word文件到目标位置", variable=options['copy_non_word']
    ).pack(anchor=tk.W, pady=(2, 20))

    # 状态区域
    status_var = tk.StringVar()
    status_var.set("等待处理...")
    ttk.Label(frame, textvariable=status_var, wraplength=550).pack(anchor=tk.W, pady=(0, 10))

    # 处理按钮
    def process_files():
        folder_path = folder_var.get().replace("已选择文件夹: ", "")
        if not folder_path or folder_path == "等待选择文件夹...":
            messagebox.showwarning("警告", "请先选择文件夹")
            return

        # 检查是否至少选择了一个处理选项
        if not any([
            options['remove_header_footer'].get(),
            options['add_custom_header'].get(),
            options['add_page_number'].get(),
            options['replace_patterns'].get()
        ]):
            if not messagebox.askyesno("提示", "未选择任何处理选项，是否继续?"):
                return

        # 准备处理选项
        current_options = {
            key: var.get() for key, var in options.items()
        }

        # 处理processed文件夹
        if current_options['use_processed_folder']:
            processed_folder = os.path.join(folder_path, "processed")
            if os.path.exists(processed_folder):
                if not messagebox.askyesno("提示", "processed文件夹已存在，是否清空后重新处理？"):
                    return
                shutil.rmtree(processed_folder)
            os.makedirs(processed_folder, exist_ok=True)

        # 处理Word文件
        word_files = get_all_word_files(folder_path)
        total_word = len(word_files)
        success_word = 0
        error_messages = []

        for i, file_path in enumerate(word_files):
            status_var.set(f"正在处理Word文件 {i + 1}/{total_word}: {os.path.basename(file_path)}")
            root.update_idletasks()

            success, result = process_word_file(file_path, folder_path, current_options)
            if success:
                success_word += 1
            else:
                error_messages.append(result)

        # 复制非Word文件
        if current_options['copy_non_word']:
            status_var.set("正在复制非Word文件...")
            root.update_idletasks()
            copy_non_word_files(folder_path, current_options['use_processed_folder'])

        # 显示结果
        result_msg = f"处理完成！\nWord文件处理: 成功 {success_word}/{total_word}\n"
        if current_options['use_processed_folder']:
            result_msg += f"所有文件已保存到: {os.path.join(folder_path, 'processed')}\n"
        else:
            result_msg += f"文件已替换原文件，原文件备份为.bak格式\n"

        if error_messages:
            result_msg += "\n错误信息:\n" + "\n".join(error_messages)

        messagebox.showinfo("处理结果", result_msg)
        status_var.set("处理完成")

    process_btn = ttk.Button(frame, text="开始处理", command=process_files)
    process_btn.pack(pady=10)

    exit_btn = ttk.Button(frame, text="退出", command=lambda: [root.destroy(), os._exit(0)])
    exit_btn.pack(pady=5)

    root.mainloop()


if __name__ == "__main__":
    main()