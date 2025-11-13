import os
import re
import shutil
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import win32com.client  # 用于格式转换
from docx import Document  # 用于docx文档基本操作
from docx.enum.text import WD_ALIGN_PARAGRAPH  # 用于段落对齐设置
from docx.oxml import OxmlElement  # 用于操作XML元素
from docx.oxml.ns import qn  # 用于设置XML命名空间
from docx2pdf import convert  # 用于docx转pdf
from docx.shared import Pt  # 用于设置字体大小


# ------------------------------
# 通用工具函数
# ------------------------------
def select_folder():
    """选择文件夹并返回路径，若取消选择则返回None"""
    root = tk.Tk()
    root.withdraw()  # 隐藏主窗口
    folder_path = filedialog.askdirectory(title="选择文件夹")
    return folder_path if folder_path else None


def get_all_files_by_ext(folder_path, exts):
    """
    获取指定文件夹下所有符合扩展名的文件路径（排除Word临时文件）
    :param folder_path: 文件夹路径
    :param exts: 扩展名列表（如['.docx', '.doc']）
    :return: 符合条件的文件路径列表
    """
    exts = [ext.lower() for ext in exts]  # 统一转为小写便于匹配
    files = []
    # 遍历文件夹及其子文件夹
    for root, _, filenames in os.walk(folder_path):
        for filename in filenames:
            # 过滤Word临时文件（以~$开头的文件）
            if filename.startswith('~$'):
                continue
            # 检查文件是否符合任一扩展名
            if any(filename.lower().endswith(ext) for ext in exts):
                files.append(os.path.join(root, filename))
    return files


# ------------------------------
# 主功能区：Word处理功能
# ------------------------------
def remove_header_footer(doc):
    """删除文档中所有节的页眉页脚内容"""
    for section in doc.sections:
        # 处理页眉
        header = section.header
        # 反向遍历段落（避免删除时索引错乱）
        for para in reversed(header.paragraphs):
            p_element = para._element  # 获取段落的XML元素
            parent = p_element.getparent()  # 获取父节点
            if parent is not None:
                parent.remove(p_element)  # 从父节点中移除当前段落
            para._p = None  # 清除段落引用
            para._element = None

        # 处理页脚（逻辑同页眉）
        footer = section.footer
        for para in reversed(footer.paragraphs):
            p_element = para._element
            parent = p_element.getparent()
            if parent is not None:
                parent.remove(p_element)
            para._p = None
            para._element = None


def add_custom_header(doc):
    """为文档所有节添加自定义页眉（内容：学为人师，行为世范）"""
    for section in doc.sections:
        header = section.header
        para = header.add_paragraph()  # 添加新段落
        para.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT  # 左对齐
        run = para.add_run("学为人师，行为世范")  # 添加文本
        # 设置字体（同时支持英文字体和中文字体）
        run.font.name = "华文行楷"
        run._element.rPr.rFonts.set(qn('w:eastAsia'), "华文行楷")  # 中文字体设置
        run.font.size = Pt(12)  # 字体大小12磅


def add_centered_page_number(doc):
    """为文档所有节添加居中页码，格式为：第X页/共Y页"""
    for section in doc.sections:
        footer = section.footer
        p = footer.add_paragraph()  # 添加页脚段落
        para_format = p.paragraph_format
        para_format.alignment = WD_ALIGN_PARAGRAPH.CENTER  # 居中对齐

        # 添加"第"字
        run = p.add_run("第")
        run.font.name = "宋体"
        run._element.rPr.rFonts.set(qn('w:eastAsia'), "宋体")
        run.font.size = Pt(12)

        # 添加当前页码域（Word中的PAGE域）
        run = p.add_run()
        fld_char_begin = OxmlElement('w:fldChar')
        fld_char_begin.set(qn('w:fldCharType'), 'begin')  # 域开始标记
        run._r.append(fld_char_begin)

        instr_text = OxmlElement('w:instrText')
        instr_text.text = "PAGE"  # 页码域指令
        run._r.append(instr_text)

        fld_char_sep = OxmlElement('w:fldChar')
        fld_char_sep.set(qn('w:fldCharType'), 'separate')  # 域分隔标记
        run._r.append(fld_char_sep)

        fld_char_end = OxmlElement('w:fldChar')
        fld_char_end.set(qn('w:fldCharType'), 'end')  # 域结束标记
        run._r.append(fld_char_end)

        # 添加"页/共"
        run = p.add_run("页/共")
        run.font.name = "宋体"
        run._element.rPr.rFonts.set(qn('w:eastAsia'), "宋体")
        run.font.size = Pt(12)

        # 添加总页数域（Word中的NUMPAGES域）
        run = p.add_run()
        fld_char_begin = OxmlElement('w:fldChar')
        fld_char_begin.set(qn('w:fldCharType'), 'begin')
        run._r.append(fld_char_begin)

        instr_text = OxmlElement('w:instrText')
        instr_text.text = "NUMPAGES"  # 总页数域指令
        run._r.append(instr_text)

        fld_char_sep = OxmlElement('w:fldChar')
        fld_char_sep.set(qn('w:fldCharType'), 'separate')
        run._r.append(fld_char_sep)

        fld_char_end = OxmlElement('w:fldChar')
        fld_char_end.set(qn('w:fldCharType'), 'end')
        run._r.append(fld_char_end)

        # 添加"页"字
        run = p.add_run("页")
        run.font.name = "宋体"
        run._element.rPr.rFonts.set(qn('w:eastAsia'), "宋体")
        run.font.size = Pt(12)


def replace_patterns_in_paragraph(paragraph):
    """
    替换段落中符合特定模式的文本（中英文括号内的特定数字模式）
    :param paragraph: 需要处理的段落对象
    """
    text_runs = []  # 存储段落中所有文本片段（包含run对象、文本内容及位置）
    char_pos = 0  # 字符位置计数器

    for run in paragraph.runs:
        if not run.text:
            continue  # 跳过空文本
        text = run.text
        start = char_pos
        end = char_pos + len(text)
        text_runs.append((run, text, start, end))
        char_pos = end  # 更新位置

    if not text_runs:
        return  # 无文本则直接返回

    # 合并所有文本用于匹配
    all_text = ''.join([t[1] for t in text_runs])
    # 定义需要匹配的模式（中英文括号内的特定数字模式）
    english_pattern = r'\(([12][0-9]).{3,}?\)'
    chinese_pattern = r'（([12][0-9]).{3,}?）'
    replaced_ranges = []  # 存储需要替换的文本范围

    # 标记需要替换的范围
    def mark_replaced(match):
        replaced_ranges.append((match.start(), match.end()))
        return ""  # 替换为空

    # 执行匹配并标记
    re.sub(chinese_pattern, mark_replaced, all_text)
    re.sub(english_pattern, mark_replaced, all_text)

    # 创建保留标记（True表示保留，False表示删除）
    keep_mask = [True] * len(all_text)
    for start, end in replaced_ranges:
        for i in range(start, end):
            if i < len(keep_mask):
                keep_mask[i] = False

    # 根据保留标记更新每个run的文本
    for run, original_text, start, end in text_runs:
        kept_chars = []
        for i in range(start, end):
            if i < len(keep_mask) and keep_mask[i]:
                original_idx = i - start  # 计算在原始文本中的索引
                kept_chars.append(original_text[original_idx])
        run.text = ''.join(kept_chars)  # 更新run的文本


import re
from docx.oxml.ns import qn
from docx.oxml import OxmlElement


def set_outline_level(doc):
    """
    将文档中符合特定格式的段落大纲级别设置为1级，规则如下：
    - 题型+数字/中文数字：段落开头匹配即可
    - 知识点+数字/中文数字：段落开头匹配即可
    - 考点+数字：段落开头匹配即可
    - A/B/C开头的特定模式：段落中仅包含该文本（允许前后有空格、换行等空白字符）
    - 第X章：段落开头匹配即可
    - 第X单元：段落开头匹配即可
    :param doc: 文档对象
    """
    # 定义中文数字（扩展常用范围，支持组合数字）
    chinese_nums = r'(?:一|二|三|四|五|六|七|八|九|十|十一|十二|十三|十四|十五|十六|十七|十八|十九|二十|廿|卅|卌|百|千|万|廿一|廿二|卅一|卅二|卌一|卌二)'

    # 组合所有匹配模式
    pattern = (
        # 原题型模式：段落开头为"题型+数字/中文数字"，后续可接任意内容
        r'题型(?:\d+|' + chinese_nums + r').*'

        # 知识点模式（数字或中文数字）：段落开头为"知识点+数字/中文数字"
        r'|'
        r'知识点(?:\d+|' + chinese_nums + r').*'

        # 考点模式（数字）：段落开头为"考点+数字"
        r'|'
        r'考点\d+.*'

        # A/B/C开头的特定模式：允许前后有空格等空白字符，核心文本必须是这三个之一
        r'|'
        r'^\s*(?:A夯实基础|B能力提升|C综合素养)\s*$'  # \s*匹配任意数量空白字符

        # 新增：第X章模式：段落开头为"第+数字/中文数字+章"
        r'|'
        r'第(?:\d+|' + chinese_nums + r')章.*'

        # 新增：第X单元模式：段落开头为"第+数字/中文数字+单元"
        r'|'
        r'第(?:\d+|' + chinese_nums + r')单元.*'
    )

    for para in doc.paragraphs:
        # 清除段落首尾的空白字符（这里主要是为了统一处理，不影响A/B/C的匹配）
        clean_text = para.text.strip()
        # 检查是否完全匹配模式（fullmatch确保整个段落符合规则）
        if re.search(pattern, clean_text):
            # 获取或创建段落属性元素
            p_pr = para._element.get_or_add_pPr()
            # 移除已有的大纲级别设置（避免重复）
            for elem in p_pr.findall(qn('w:outlineLvl')):
                p_pr.remove(elem)
            # 创建大纲级别元素并设置为1级（Word中0对应1级）
            outline_level = OxmlElement('w:outlineLvl')
            outline_level.set(qn('w:val'), '0')
            p_pr.append(outline_level)

def process_word_file(file_path, keep_backup):
    """
    处理单个Word文件（根据选项执行相应操作）
    :param file_path: 文件路径
    :param keep_backup: 是否保留备份
    :return: (处理结果, 消息)
    """
    try:
        # 保留备份（若未存在备份）
        if keep_backup and not os.path.exists(f"{file_path}.bak"):
            shutil.copy2(file_path, f"{file_path}.bak")

        # 打开文档进行处理
        doc = Document(file_path)

        # 根据选项执行操作
        if options['remove_header_footer'].get():
            remove_header_footer(doc)
        if options['add_page_number'].get():
            add_centered_page_number(doc)
        if options['add_custom_header'].get():
            add_custom_header(doc)

        if options['replace_patterns'].get():
            # 处理普通段落
            for para in doc.paragraphs:
                replace_patterns_in_paragraph(para)
            # 处理表格中的段落
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for para in cell.paragraphs:
                            replace_patterns_in_paragraph(para)

        # 设置题型段落的大纲级别为1级
        if options['set_question_outline'].get():
            set_outline_level(doc)

        # 保存修改
        doc.save(file_path)
        return True, file_path
    except Exception as e:
        return False, f"处理失败：{os.path.basename(file_path)}，错误：{str(e)}"


# ------------------------------
# 辅助功能区：格式转换功能
# ------------------------------
def batch_convert_doc_to_docx(root_dir, keep_source, status_var):
    """
    批量将doc文件转换为docx文件
    :param root_dir: 根目录
    :param keep_source: 是否保留源文件
    :param status_var: 状态变量（用于更新UI）
    """
    root_dir = os.path.normpath(root_dir)
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False  # 隐藏Word窗口
    word.DisplayAlerts = 0  # 禁用提示框

    total = 0  # 总文件数
    converted = 0  # 成功转换数
    deleted = 0  # 成功删除源文件数
    skipped_convert = 0  # 跳过转换数（目标文件已存在）
    skipped_delete = 0  # 跳过删除数（删除失败）
    error_details = []  # 错误详情

    # 统计所有doc文件
    doc_files = []
    for foldername, _, filenames in os.walk(root_dir):
        for filename in filenames:
            lower_name = filename.lower()
            # 匹配.doc但不匹配.docx（避免重复）
            if lower_name.endswith(".doc") and not lower_name.endswith(".docx"):
                doc_files.append(os.path.join(foldername, filename))
    total = len(doc_files)

    # 开始转换
    for i, doc_path in enumerate(doc_files):
        filename = os.path.basename(doc_path)
        status_var.set(f"正在转换DOC→DOCX ({i + 1}/{total})：{filename}")
        root.update_idletasks()  # 更新UI

        # 构建目标docx路径
        docx_name = f"{os.path.splitext(filename)[0]}.docx"
        docx_path = os.path.join(os.path.dirname(doc_path), docx_name)

        # 若目标文件不存在则转换
        if not os.path.exists(docx_path):
            try:
                doc = word.Documents.Open(os.path.abspath(doc_path))
                # 12对应docx格式
                doc.SaveAs2(os.path.abspath(docx_path), FileFormat=12)
                doc.Close()
                converted += 1
            except Exception as e:
                # 记录错误信息（简化错误描述）
                error_details.append(f"转换失败：{filename} - {str(e).split(',')[0]}")
                continue
        else:
            skipped_convert += 1

        # 若不需要保留源文件则删除
        if not keep_source:
            if os.path.exists(doc_path):
                try:
                    os.remove(doc_path)
                    deleted += 1
                except Exception as e:
                    skipped_delete += 1
                    error_details.append(f"删除失败：{filename} - {str(e)}")

    word.Quit()  # 关闭Word应用

    # 构建结果消息
    result_msg = (f"处理完成！\n总doc文件：{total}\n"
                  f"新转换docx：{converted} | 已存在docx：{skipped_convert}\n")
    if not keep_source:
        result_msg += f"已删除源文件：{deleted} | 删除失败：{skipped_delete}\n"
    else:
        result_msg += "已保留所有源文件\n"

    if error_details:
        result_msg += "\n错误详情：\n" + "\n".join(error_details[:5])  # 只显示前5个错误
    messagebox.showinfo("DOC转DOCX结果", result_msg)
    status_var.set("就绪")


def batch_convert_docx_to_pdf(root_dir, use_separate_folder, status_var):
    """
    批量将docx文件转换为pdf文件（改用Word原生接口，保留大纲/目录/书签）
    :param root_dir: 根目录
    :param use_separate_folder: 是否使用单独文件夹保存
    :param status_var: 状态变量（用于更新UI）
    """
    root_dir = os.path.normpath(root_dir)
    # 初始化Word应用（隐藏窗口，禁用提示）
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0

    total = 0  # 总文件数
    success = 0  # 成功转换数
    replaced = 0  # 替换现有文件数
    skipped = 0  # 跳过数
    fail = 0  # 失败数
    error_details = []  # 错误详情

    # 重要：定义文件覆盖确认的全局变量（原代码逻辑）
    apply_to_all = False  # 是否对所有文件应用相同选择
    global_decision = None  # 全局选择（替换/跳过）

    # 获取所有docx文件（排除临时文件）
    docx_files = get_all_files_by_ext(root_dir, ['.docx'])
    total = len(docx_files)

    # 开始转换
    for i, docx_path in enumerate(docx_files):
        filename = os.path.basename(docx_path)
        status_var.set(f"正在转换DOCX→PDF ({i + 1}/{total})：{filename}")
        root.update_idletasks()  # 更新UI

        # 构建目标PDF路径
        if use_separate_folder:
            # 保持相对路径结构
            relative_path = os.path.relpath(docx_path, root_dir)
            pdf_path = os.path.join(root_dir, "docx2pdf", f"{os.path.splitext(relative_path)[0]}.pdf")
            os.makedirs(os.path.dirname(pdf_path), exist_ok=True)  # 创建父目录
        else:
            # 保存到原位置
            pdf_path = f"{os.path.splitext(docx_path)[0]}.pdf"

        # 处理目标文件已存在的情况
        if os.path.exists(pdf_path):
            if apply_to_all:
                # 应用全局选择
                if not global_decision:
                    skipped += 1
                    continue
            else:
                # 弹出对话框询问用户（原代码的回调逻辑）
                dialog = tk.Toplevel()
                dialog.title("文件已存在")
                dialog.geometry("400x150")
                dialog.transient(root)  # 设置为主窗口的子窗口
                dialog.grab_set()  # 模态窗口

                ttk.Label(dialog, text=f"文件 '{os.path.basename(pdf_path)}' 已存在，是否替换？").pack(pady=10, padx=10)

                apply_var = tk.BooleanVar(value=False)
                ttk.Checkbutton(dialog, text="对后续所有文件应用此选择", variable=apply_var).pack(anchor=tk.W, padx=20)

                decision = None  # 用户选择

                # 替换按钮回调（保留原逻辑）
                def on_replace():
                    nonlocal decision
                    decision = True
                    dialog.destroy()

                # 跳过按钮回调（保留原逻辑）
                def on_skip():
                    nonlocal decision
                    decision = False
                    dialog.destroy()

                # 按钮布局（保留原UI）
                btn_frame = ttk.Frame(dialog)
                btn_frame.pack(pady=15)
                ttk.Button(btn_frame, text="替换", command=on_replace).pack(side=tk.LEFT, padx=5)
                ttk.Button(btn_frame, text="跳过", command=on_skip).pack(side=tk.LEFT, padx=5)

                dialog.wait_window()  # 等待对话框关闭

                if decision is None:
                    skipped += 1
                    continue

                # 应用全局选择（保留原逻辑）
                if apply_var.get():
                    apply_to_all = True
                    global_decision = decision

                if not decision:
                    skipped += 1
                    continue

        # 核心：用Word原生接口转换（保留大纲/目录）
        try:
            # 打开docx文件（只读模式避免锁定）
            doc = word.Documents.Open(os.path.abspath(docx_path), ReadOnly=True)
            # 17 = Word的PDF格式常量，CreateBookmarks=1保留大纲书签
            doc.ExportAsFixedFormat(
                OutputFileName=os.path.abspath(pdf_path),
                ExportFormat=17,  # 17对应PDF格式
                IncludeDocProps=True,
                CreateBookmarks=1,  # 1=基于标题创建书签（对应大纲级别）
                DocStructureTags=True  # 保留文档结构（支持PDF目录导航）
            )
            doc.Close(SaveChanges=0)  # 关闭文件，不保存修改
            success += 1
            # 统计替换数（保留原逻辑）
            if os.path.exists(pdf_path) and (apply_to_all and global_decision or not apply_to_all):
                replaced += 1
        except Exception as e:
            error_details.append(f"{filename}：{str(e)}")
            fail += 1
            continue

    # 关闭Word应用
    word.Quit()

    # 构建结果消息（沿用原有格式）
    result_msg = (f"转换完成！\n总文件：{total}\n"
                  f"成功转换：{success - replaced} | 替换现有：{replaced} | 跳过：{skipped} | 失败：{fail}\n")
    if use_separate_folder:
        result_msg += f"PDF文件保存于：{os.path.join(root_dir, 'docx2pdf')}"
    else:
        result_msg += "PDF文件保存于原docx文件位置"

    if error_details:
        result_msg += f"\n\n错误详情：\n" + "\n".join(error_details[:5])  # 只显示前5个错误
    messagebox.showinfo("DOCX转PDF结果", result_msg)
    status_var.set("就绪")


# ------------------------------
# 主界面
# ------------------------------
def main():
    global options, root
    root = tk.Tk()
    root.title("Word文件处理工具")
    root.geometry("700x800")  # 调整高度
    root.resizable(False, False)  # 禁止调整窗口大小

    # 变量定义
    folder_var = tk.StringVar(value="等待选择文件夹...")
    options = {
        # 主功能选项
        'remove_header_footer': tk.BooleanVar(value=True),
        'add_custom_header': tk.BooleanVar(value=True),
        'add_page_number': tk.BooleanVar(value=True),
        'replace_patterns': tk.BooleanVar(value=True),
        'set_question_outline': tk.BooleanVar(value=True),  # 新增：设置题型段落大纲级别
        'keep_backup': tk.BooleanVar(value=False),
        # 辅助功能选项
        'keep_source_doc': tk.BooleanVar(value=False),
        'docx2pdf_separate_folder': tk.BooleanVar(value=False)
    }

    # 主框架
    main_frame = ttk.Frame(root, padding=15)
    main_frame.pack(fill=tk.BOTH, expand=True)

    # 文件夹选择区
    ttk.Label(main_frame, text="工作文件夹:", font=("Arial", 10, "bold")).pack(anchor=tk.W, pady=(0, 5))
    ttk.Label(main_frame, textvariable=folder_var, wraplength=650).pack(anchor=tk.W, pady=(0, 10))

    def select_folder_action():
        """选择文件夹并更新显示"""
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

    # 选项列1
    col1 = ttk.Frame(options_frame)
    col1.pack(side=tk.LEFT, fill=tk.X, expand=True)
    ttk.Checkbutton(col1, text="删除页眉页脚", variable=options['remove_header_footer']).pack(anchor=tk.W, pady=2)
    ttk.Checkbutton(col1, text="添加自定义页眉", variable=options['add_custom_header']).pack(anchor=tk.W, pady=2)
    # 替换为设置大纲级别选项
    ttk.Checkbutton(
        col1,
        text="将题型、知识点、考点等段落大纲级别设置为1级",
        variable=options['set_question_outline']
    ).pack(anchor=tk.W, pady=2)

    # 选项列2
    col2 = ttk.Frame(options_frame)
    col2.pack(side=tk.LEFT, fill=tk.X, expand=True)
    ttk.Checkbutton(
        col2,
        text="添加居中页码（格式：第X页/共Y页）",
        variable=options['add_page_number']
    ).pack(anchor=tk.W, pady=2)
    ttk.Checkbutton(
        col2,
        text="替换指定文本模式",
        variable=options['replace_patterns']
    ).pack(anchor=tk.W, pady=2)

    # 保存选项
    ttk.Label(main_frame, text="文件保存选项:", font=("Arial", 10, "bold")).pack(anchor=tk.W, pady=(0, 5))
    ttk.Checkbutton(
        main_frame,
        text="保留原文件为.bak备份（不勾选则直接替换源文件）",
        variable=options['keep_backup']
    ).pack(anchor=tk.W, pady=(0, 15))

    # 状态显示区（所有功能共用）
    status_var = tk.StringVar(value="就绪")
    ttk.Label(main_frame, textvariable=status_var, wraplength=650).pack(anchor=tk.W, pady=(0, 10))

    def process_word_files_action():
        """处理Word文件的入口函数"""
        folder_path = folder_var.get().replace("已选择：", "")
        if not folder_path or folder_path == "等待选择文件夹...":
            messagebox.showwarning("警告", "请先选择文件夹")
            return

        # 检查是否选择了至少一个处理选项
        if not any([
            options['remove_header_footer'].get(),
            options['add_custom_header'].get(),
            options['add_page_number'].get(),
            options['replace_patterns'].get(),
            options['set_question_outline'].get()  # 更新选项
        ]):
            if not messagebox.askyesno("提示", "未选择任何处理选项，是否继续？"):
                return

        keep_backup = options['keep_backup'].get()
        word_files = get_all_files_by_ext(folder_path, ['.docx'])
        total = len(word_files)
        success = 0
        errors = []

        # 批量处理文件
        for i, file_path in enumerate(word_files):
            status_var.set(f"正在处理Word文件 ({i + 1}/{total})：{os.path.basename(file_path)}")
            root.update_idletasks()
            res, msg = process_word_file(file_path, keep_backup)
            if res:
                success += 1
            else:
                errors.append(msg)

        # 构建结果消息
        result = f"处理完成！成功：{success}/{total}\n"
        if keep_backup:
            result += "原文件已备份为.bak格式，处理后的文件已替换原文件"
        else:
            result += "已直接替换原文件（未保留备份）"

        if errors:
            result += f"\n\n错误列表：\n" + "\n".join(errors[:5])  # 只显示前5个错误
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
        """触发DOC转DOCX功能"""
        folder = folder_var.get().replace("已选择：", "")
        if not folder or folder == "等待选择文件夹...":
            messagebox.showwarning("警告", "请先选择文件夹")
            return
        batch_convert_doc_to_docx(folder, options['keep_source_doc'].get(), status_var)

    ttk.Button(main_frame, text="批量转换DOC→DOCX", command=convert_doc_action).pack(pady=(0, 10))

    # DOCX转PDF
    ttk.Label(main_frame, text="DOCX转PDF选项:", font=("Arial", 10, "bold")).pack(anchor=tk.W, pady=(0, 5))
    ttk.Checkbutton(
        main_frame,
        text="保存到docx2pdf文件夹（不勾选则保存到原位置）",
        variable=options['docx2pdf_separate_folder']
    ).pack(anchor=tk.W, pady=(0, 5))

    def convert_pdf_action():
        """触发DOCX转PDF功能"""
        folder = folder_var.get().replace("已选择：", "")
        if not folder or folder == "等待选择文件夹...":
            messagebox.showwarning("警告", "请先选择文件夹")
            return
        batch_convert_docx_to_pdf(folder, options['docx2pdf_separate_folder'].get(), status_var)

    ttk.Button(main_frame, text="批量转换DOCX→PDF", command=convert_pdf_action).pack(pady=(0, 10))

    # 退出按钮
    ttk.Button(main_frame, text="退出", command=lambda: [root.destroy(), os._exit(0)]).pack(pady=15)

    root.mainloop()


if __name__ == "__main__":
    main()