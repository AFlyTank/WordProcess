import os
import re
import shutil
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import win32com.client
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx2pdf import convert
from docx.shared import Pt


# ------------------------------
# 通用工具函数
# ------------------------------
def select_folder():
    root = tk.Tk()
    root.withdraw()
    folder_path = filedialog.askdirectory(title="选择文件夹")
    return folder_path if folder_path else None


def get_all_files_by_ext(folder_path, exts):
    exts = [ext.lower() for ext in exts]
    files = []
    for root, _, filenames in os.walk(folder_path):
        for filename in filenames:
            if any(filename.lower().endswith(ext) for ext in exts):
                files.append(os.path.join(root, filename))
    return files


# ------------------------------
# 主功能区：Word处理功能
# ------------------------------
def remove_header_footer(doc):
    for section in doc.sections:
        header = section.header
        for para in reversed(header.paragraphs):
            p_element = para._element
            parent = p_element.getparent()
            if parent:
                parent.remove(p_element)
            para._p = None
            para._element = None

        footer = section.footer
        for para in reversed(footer.paragraphs):
            p_element = para._element
            parent = p_element.getparent()
            if parent:
                parent.remove(p_element)
            para._p = None
            para._element = None


def add_custom_header(doc):
    for section in doc.sections:
        header = section.header
        para = header.add_paragraph()
        para.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
        run = para.add_run("学为人师，行为世范")
        run.font.name = "华文行楷"
        run._element.rPr.rFonts.set(qn('w:eastAsia'), "华文行楷")
        run.font.size = Pt(12)


def add_centered_page_number(doc):
    for section in doc.sections:
        footer = section.footer
        p = footer.add_paragraph()
        para_format = p.paragraph_format
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


def process_word_file(file_path, keep_backup):
    try:
        if keep_backup and not os.path.exists(f"{file_path}.bak"):
            shutil.copy2(file_path, f"{file_path}.bak")

        doc = Document(file_path)

        if options['remove_header_footer'].get():
            remove_header_footer(doc)
        if options['add_page_number'].get():
            add_centered_page_number(doc)
        if options['add_custom_header'].get():
            add_custom_header(doc)

        if options['replace_patterns'].get():
            for para in doc.paragraphs:
                replace_patterns_in_paragraph(para)
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for para in cell.paragraphs:
                            replace_patterns_in_paragraph(para)

        doc.save(file_path)
        return True, file_path
    except Exception as e:
        return False, f"处理失败：{os.path.basename(file_path)}，错误：{str(e)}"


# ------------------------------
# 辅助功能区：格式转换功能
# ------------------------------
def batch_convert_doc_to_docx(root_dir, keep_source):
    root_dir = os.path.normpath(root_dir)
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0

    total = 0
    converted = 0
    deleted = 0
    skipped_convert = 0
    skipped_delete = 0
    error_details = []

    for foldername, _, filenames in os.walk(root_dir):
        for filename in filenames:
            lower_name = filename.lower()
            if lower_name.endswith(".doc") and not lower_name.endswith(".docx"):
                total += 1
                doc_path = os.path.join(foldername, filename)
                docx_name = f"{os.path.splitext(filename)[0]}.docx"
                docx_path = os.path.join(foldername, docx_name)

                if not os.path.exists(docx_path):
                    try:
                        doc = word.Documents.Open(os.path.abspath(doc_path))
                        doc.SaveAs2(os.path.abspath(docx_path), FileFormat=12)
                        doc.Close()
                        converted += 1
                    except Exception as e:
                        error_details.append(f"转换失败：{filename} - {str(e).split(',')[0]}")
                        continue
                else:
                    skipped_convert += 1

                if not keep_source:
                    if os.path.exists(doc_path):
                        try:
                            os.remove(doc_path)
                            deleted += 1
                        except Exception as e:
                            skipped_delete += 1
                            error_details.append(f"删除失败：{filename} - {str(e)}")

    word.Quit()

    result_msg = (f"处理完成！\n总doc文件：{total}\n"
                  f"新转换docx：{converted} | 已存在docx：{skipped_convert}\n")
    if not keep_source:
        result_msg += f"已删除源文件：{deleted} | 删除失败：{skipped_delete}\n"
    else:
        result_msg += "已保留所有源文件\n"

    if error_details:
        result_msg += "\n错误详情：\n" + "\n".join(error_details[:5])
    messagebox.showinfo("DOC转DOCX结果", result_msg)


def batch_convert_docx_to_pdf(root_dir, use_separate_folder, apply_to_all_var):
    """
    批量转换docx到pdf，处理文件冲突：
    - 存在同名文件时询问是否替换
    - 支持"应用于所有"选项，一次性决定后续所有冲突
    """
    root_dir = os.path.normpath(root_dir)
    docx_files = get_all_files_by_ext(root_dir, ['.docx'])
    total = len(docx_files)
    success = 0
    replaced = 0  # 替换的文件数
    skipped = 0  # 跳过的文件数
    fail = 0
    error_details = []

    # 冲突处理状态变量
    apply_to_all = False
    global_decision = None  # 全局决定：True=替换所有，False=跳过所有

    if use_separate_folder:
        output_root = os.path.join(root_dir, "docx2pdf")
        os.makedirs(output_root, exist_ok=True)

    for i, docx_path in enumerate(docx_files):
        # 确定PDF保存路径
        if use_separate_folder:
            relative_path = os.path.relpath(docx_path, root_dir)
            pdf_path = os.path.join(output_root, f"{os.path.splitext(relative_path)[0]}.pdf")
            os.makedirs(os.path.dirname(pdf_path), exist_ok=True)
        else:
            pdf_path = f"{os.path.splitext(docx_path)[0]}.pdf"

        # 检查文件是否存在
        if os.path.exists(pdf_path):
            # 处理冲突
            if apply_to_all:
                # 使用全局决定
                if global_decision:
                    # 替换现有文件
                    pass
                else:
                    # 跳过
                    skipped += 1
                    continue
            else:
                # 弹出询问对话框
                filename = os.path.basename(pdf_path)
                dialog = tk.Toplevel()
                dialog.title("文件已存在")
                dialog.geometry("400x150")
                dialog.transient(root)  # 设置为主窗口的子窗口
                dialog.grab_set()  # 模态窗口，阻止操作主窗口

                ttk.Label(dialog, text=f"文件 '{filename}' 已存在，是否替换？").pack(pady=10, padx=10)

                # 应用于所有选项
                apply_var = tk.BooleanVar(value=False)
                ttk.Checkbutton(dialog, text="对后续所有文件应用此选择", variable=apply_var).pack(anchor=tk.W, padx=20)

                decision = None  # 存储用户决定

                def on_replace():
                    nonlocal decision
                    decision = True
                    dialog.destroy()

                def on_skip():
                    nonlocal decision
                    decision = False
                    dialog.destroy()

                btn_frame = ttk.Frame(dialog)
                btn_frame.pack(pady=15)
                ttk.Button(btn_frame, text="替换", command=on_replace).pack(side=tk.LEFT, padx=5)
                ttk.Button(btn_frame, text="跳过", command=on_skip).pack(side=tk.LEFT, padx=5)

                dialog.wait_window()  # 等待对话框关闭

                if decision is None:
                    # 对话框被关闭，视为跳过
                    skipped += 1
                    continue

                # 记录是否应用于所有
                if apply_var.get():
                    apply_to_all = True
                    global_decision = decision

                if not decision:
                    skipped += 1
                    continue

        try:
            convert(docx_path, pdf_path)
            success += 1
            if os.path.exists(pdf_path) and (apply_to_all and global_decision or not apply_to_all):
                replaced += 1
        except Exception as e:
            error_details.append(f"{os.path.basename(docx_path)}：{str(e)}")
            fail += 1

    # 生成结果信息
    result_msg = (f"转换完成！\n总文件：{total}\n"
                  f"成功转换：{success - replaced} | 替换现有：{replaced} | 跳过：{skipped} | 失败：{fail}\n")
    if use_separate_folder:
        result_msg += f"PDF文件保存于：{output_root}"
    else:
        result_msg += "PDF文件保存于原docx文件位置"

    if error_details:
        result_msg += f"\n\n错误详情：\n" + "\n".join(error_details[:5])
    messagebox.showinfo("DOCX转PDF结果", result_msg)


# ------------------------------
# 主界面
# ------------------------------
def main():
    global options, root
    root = tk.Tk()
    root.title("Word文件处理工具")
    root.geometry("700x750")
    root.resizable(False, False)

    # 变量定义
    folder_var = tk.StringVar(value="等待选择文件夹...")
    options = {
        # 主功能选项
        'remove_header_footer': tk.BooleanVar(value=True),
        'add_custom_header': tk.BooleanVar(value=True),
        'add_page_number': tk.BooleanVar(value=True),
        'replace_patterns': tk.BooleanVar(value=True),
        'keep_backup': tk.BooleanVar(value=True),
        # 辅助功能选项
        'keep_source_doc': tk.BooleanVar(value=True),
        'docx2pdf_separate_folder': tk.BooleanVar(value=False),
        'docx2pdf_apply_to_all': tk.BooleanVar(value=False)  # 应用于所有选项
    }

    # 主框架
    main_frame = ttk.Frame(root, padding=15)
    main_frame.pack(fill=tk.BOTH, expand=True)

    # 文件夹选择区
    ttk.Label(main_frame, text="工作文件夹:", font=("Arial", 10, "bold")).pack(anchor=tk.W, pady=(0, 5))
    ttk.Label(main_frame, textvariable=folder_var, wraplength=650).pack(anchor=tk.W, pady=(0, 10))

    def select_folder_action():
        folder = select_folder()
        if folder:
            folder_var.set(f"已选择：{folder}")

    ttk.Button(main_frame, text="选择文件夹", command=select_folder_action).pack(anchor=tk.W, pady=(0, 20))

    # ------------------------------
    # 主功能区：Word处理
    # ------------------------------
    ttk.Separator(main_frame, orient="horizontal").pack(fill=tk.X, pady=10)
    ttk.Label(main_frame, text="【主功能区：Word文件处理】", font=("Arial", 11, "bold")).pack(anchor=tk.W, pady=(0, 10))

    # 处理内容选项
    ttk.Label(main_frame, text="处理内容选项:", font=("Arial", 10, "bold")).pack(anchor=tk.W, pady=(0, 5))
    options_frame = ttk.Frame(main_frame)
    options_frame.pack(fill=tk.X, pady=(0, 10))

    col1 = ttk.Frame(options_frame)
    col1.pack(side=tk.LEFT, fill=tk.X, expand=True)
    ttk.Checkbutton(col1, text="删除页眉页脚", variable=options['remove_header_footer']).pack(anchor=tk.W, pady=2)
    ttk.Checkbutton(col1, text="添加自定义页眉", variable=options['add_custom_header']).pack(anchor=tk.W, pady=2)

    col2 = ttk.Frame(options_frame)
    col2.pack(side=tk.LEFT, fill=tk.X, expand=True)
    ttk.Checkbutton(col2, text="添加居中页码", variable=options['add_page_number']).pack(anchor=tk.W, pady=2)
    ttk.Checkbutton(col2, text="替换指定文本模式", variable=options['replace_patterns']).pack(anchor=tk.W, pady=2)

    # 保存选项
    ttk.Label(main_frame, text="文件保存选项:", font=("Arial", 10, "bold")).pack(anchor=tk.W, pady=(0, 5))
    ttk.Checkbutton(
        main_frame,
        text="保留原文件为.bak备份（不勾选则直接替换源文件）",
        variable=options['keep_backup']
    ).pack(anchor=tk.W, pady=(0, 15))

    # 主功能处理按钮
    status_var = tk.StringVar(value="就绪")
    ttk.Label(main_frame, textvariable=status_var, wraplength=650).pack(anchor=tk.W, pady=(0, 10))

    def process_word_files_action():
        folder_path = folder_var.get().replace("已选择：", "")
        if not folder_path or folder_path == "等待选择文件夹...":
            messagebox.showwarning("警告", "请先选择文件夹")
            return

        if not any([
            options['remove_header_footer'].get(),
            options['add_custom_header'].get(),
            options['add_page_number'].get(),
            options['replace_patterns'].get()
        ]):
            if not messagebox.askyesno("提示", "未选择任何处理选项，是否继续？"):
                return

        keep_backup = options['keep_backup'].get()
        word_files = get_all_files_by_ext(folder_path, ['.docx'])
        total = len(word_files)
        success = 0
        errors = []

        for i, file_path in enumerate(word_files):
            status_var.set(f"正在处理 {i + 1}/{total}：{os.path.basename(file_path)}")
            root.update_idletasks()
            res, msg = process_word_file(file_path, keep_backup)
            if res:
                success += 1
            else:
                errors.append(msg)

        result = f"处理完成！成功：{success}/{total}\n"
        if keep_backup:
            result += "原文件已备份为.bak格式，处理后的文件已替换原文件"
        else:
            result += "已直接替换原文件（未保留备份）"

        if errors:
            result += f"\n\n错误列表：\n" + "\n".join(errors[:5])
        messagebox.showinfo("处理结果", result)
        status_var.set("就绪")

    ttk.Button(main_frame, text="开始处理Word文件", command=process_word_files_action).pack(pady=(0, 15))

    # ------------------------------
    # 辅助功能区：格式转换
    # ------------------------------
    ttk.Separator(main_frame, orient="horizontal").pack(fill=tk.X, pady=10)
    ttk.Label(main_frame, text="【辅助功能区：格式转换】", font=("Arial", 11, "bold")).pack(anchor=tk.W, pady=(0, 10))

    # DOC转DOCX
    ttk.Label(main_frame, text="DOC转DOCX选项:", font=("Arial", 10, "bold")).pack(anchor=tk.W, pady=(0, 5))
    ttk.Checkbutton(
        main_frame,
        text="保留源文件（.doc）",
        variable=options['keep_source_doc']
    ).pack(anchor=tk.W, pady=(0, 5))

    def convert_doc_action():
        folder = folder_var.get().replace("已选择：", "")
        if not folder or folder == "等待选择文件夹...":
            messagebox.showwarning("警告", "请先选择文件夹")
            return
        status_var.set("正在进行DOC转DOCX...")
        root.update_idletasks()
        batch_convert_doc_to_docx(folder, options['keep_source_doc'].get())
        status_var.set("就绪")

    ttk.Button(main_frame, text="批量转换DOC→DOCX", command=convert_doc_action).pack(pady=(0, 10))

    # DOCX转PDF
    ttk.Label(main_frame, text="DOCX转PDF选项:", font=("Arial", 10, "bold")).pack(anchor=tk.W, pady=(0, 5))
    ttk.Checkbutton(
        main_frame,
        text="保存到docx2pdf文件夹（不勾选则保存到原位置）",
        variable=options['docx2pdf_separate_folder']
    ).pack(anchor=tk.W, pady=(0, 5))

    def convert_pdf_action():
        folder = folder_var.get().replace("已选择：", "")
        if not folder or folder == "等待选择文件夹...":
            messagebox.showwarning("警告", "请先选择文件夹")
            return
        status_var.set("正在进行DOCX转PDF...")
        root.update_idletasks()
        batch_convert_docx_to_pdf(
            folder,
            options['docx2pdf_separate_folder'].get(),
            options['docx2pdf_apply_to_all'].get()
        )
        status_var.set("就绪")

    ttk.Button(main_frame, text="批量转换DOCX→PDF", command=convert_pdf_action).pack(pady=(0, 10))

    # 退出按钮
    ttk.Button(main_frame, text="退出", command=lambda: [root.destroy(), os._exit(0)]).pack(pady=15)

    root.mainloop()


if __name__ == "__main__":
    main()