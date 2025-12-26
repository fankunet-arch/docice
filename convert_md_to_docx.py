#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
专业Markdown到Word文档转换器
- 完全去除Markdown语法符号
- 专业排版设计和美化
- 优化视觉效果
"""

import os
import re
from docx import Document
from docx.shared import Pt, RGBColor, Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml.ns import qn
from docx.oxml import OxmlElement


def set_cell_background(cell, color):
    """设置单元格背景色"""
    shading_elm = OxmlElement('w:shd')
    shading_elm.set(qn('w:fill'), color)
    cell._element.get_or_add_tcPr().append(shading_elm)


def set_cell_border(cell, **kwargs):
    """设置表格单元格边框"""
    tc = cell._element
    tcPr = tc.get_or_add_tcPr()
    tcBorders = OxmlElement('w:tcBorders')

    for edge in ('top', 'left', 'bottom', 'right'):
        if kwargs.get(edge):
            element = OxmlElement(f'w:{edge}')
            element.set(qn('w:val'), 'single')
            element.set(qn('w:sz'), '8')
            element.set(qn('w:space'), '0')
            element.set(qn('w:color'), '4472C4')
            tcBorders.append(element)

    tcPr.append(tcBorders)


def parse_inline_formatting(text):
    """
    解析并清理内联格式，返回格式化片段列表
    完全去除Markdown符号
    """
    segments = []
    i = 0

    while i < len(text):
        # 粗体 **text**
        if text[i:i+2] == '**':
            end = text.find('**', i + 2)
            if end != -1:
                segments.append({
                    'text': text[i+2:end],
                    'bold': True,
                    'italic': False,
                    'code': False
                })
                i = end + 2
                continue

        # 斜体 *text*
        if text[i:i+1] == '*':
            end = text.find('*', i + 1)
            if end != -1:
                segments.append({
                    'text': text[i+1:end],
                    'bold': False,
                    'italic': True,
                    'code': False
                })
                i = end + 1
                continue

        # 代码 `text`
        if text[i:i+1] == '`':
            end = text.find('`', i + 1)
            if end != -1:
                segments.append({
                    'text': text[i+1:end],
                    'bold': False,
                    'italic': False,
                    'code': True
                })
                i = end + 1
                continue

        # 普通文本
        next_pos = len(text)
        for marker in ['**', '*', '`']:
            pos = text.find(marker, i)
            if pos != -1 and pos < next_pos:
                next_pos = pos

        if next_pos > i:
            segments.append({
                'text': text[i:next_pos],
                'bold': False,
                'italic': False,
                'code': False
            })
            i = next_pos
        else:
            i += 1

    return segments


def add_formatted_text(paragraph, text):
    """为段落添加格式化文本"""
    segments = parse_inline_formatting(text)

    for seg in segments:
        run = paragraph.add_run(seg['text'])

        if seg['bold']:
            run.bold = True
            run.font.color.rgb = RGBColor(31, 78, 120)  # 深蓝色

        if seg['italic']:
            run.italic = True

        if seg['code']:
            run.font.name = 'Consolas'
            run.font.size = Pt(10)
            run.font.color.rgb = RGBColor(199, 37, 78)
            run.bold = False


def add_title_with_design(doc, text, level):
    """添加设计精美的标题"""
    para = doc.add_paragraph()
    para.paragraph_format.space_before = Pt(18 if level == 1 else 12)
    para.paragraph_format.space_after = Pt(12 if level == 1 else 8)

    run = para.add_run(text)
    run.bold = True

    # 根据级别设置样式
    if level == 1:
        run.font.size = Pt(24)
        run.font.color.rgb = RGBColor(31, 78, 120)  # 深蓝色
        # 添加底部边框
        para.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
    elif level == 2:
        run.font.size = Pt(18)
        run.font.color.rgb = RGBColor(68, 114, 196)  # 蓝色
    elif level == 3:
        run.font.size = Pt(14)
        run.font.color.rgb = RGBColor(91, 155, 213)  # 浅蓝色
    else:
        run.font.size = Pt(12)
        run.font.color.rgb = RGBColor(127, 127, 127)  # 灰色

    run.font.name = 'Microsoft YaHei'
    run._element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑')

    return para


def add_quote_block(doc, text):
    """添加引用块（带背景色和样式）"""
    para = doc.add_paragraph()

    # 设置缩进和间距
    para.paragraph_format.left_indent = Inches(0.3)
    para.paragraph_format.right_indent = Inches(0.3)
    para.paragraph_format.space_before = Pt(6)
    para.paragraph_format.space_after = Pt(6)

    # 添加文本
    add_formatted_text(para, text)

    # 设置字体为斜体
    for run in para.runs:
        run.italic = True
        run.font.color.rgb = RGBColor(68, 114, 196)

    # 添加背景色（通过shading）
    shading_elm = OxmlElement('w:shd')
    shading_elm.set(qn('w:fill'), 'E7F3FF')
    para._element.get_or_add_pPr().append(shading_elm)

    return para


def add_code_block(doc, code_text):
    """添加代码块（带背景色）"""
    para = doc.add_paragraph(code_text)

    # 设置字体
    for run in para.runs:
        run.font.name = 'Consolas'
        run.font.size = Pt(9)
        run.font.color.rgb = RGBColor(51, 51, 51)

    # 设置缩进和间距
    para.paragraph_format.left_indent = Inches(0.5)
    para.paragraph_format.space_before = Pt(6)
    para.paragraph_format.space_after = Pt(6)

    # 添加背景色
    shading_elm = OxmlElement('w:shd')
    shading_elm.set(qn('w:fill'), 'F5F5F5')
    para._element.get_or_add_pPr().append(shading_elm)

    return para


def parse_table(lines):
    """解析Markdown表格"""
    if len(lines) < 2:
        return None

    # 解析表头
    header = [cell.strip() for cell in lines[0].split('|')[1:-1]]

    # 过滤数据行（跳过分隔行）
    data_rows = []
    for line in lines[2:]:
        if '|' in line and not all(c in '|-: ' for c in line):
            cells = [cell.strip() for cell in line.split('|')[1:-1]]
            data_rows.append(cells)

    return header, data_rows


def add_beautiful_table(doc, header, data_rows):
    """添加美化的表格"""
    if not data_rows:
        return

    table = doc.add_table(rows=1, cols=len(header))
    table.style = 'Table Grid'

    # 添加表头
    header_row = table.rows[0]
    for idx, header_text in enumerate(header):
        cell = header_row.cells[idx]
        cell.text = header_text

        # 表头样式：深蓝色背景、白色粗体文字
        set_cell_background(cell, '4472C4')
        set_cell_border(cell, top=True, left=True, bottom=True, right=True)

        for para in cell.paragraphs:
            for run in para.runs:
                run.font.bold = True
                run.font.color.rgb = RGBColor(255, 255, 255)
                run.font.size = Pt(11)
                run.font.name = 'Microsoft YaHei'
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # 添加数据行
    for row_idx, row_data in enumerate(data_rows):
        row = table.add_row()

        # 交替行背景色
        bg_color = 'D9E2F3' if row_idx % 2 == 0 else 'FFFFFF'

        for col_idx, cell_text in enumerate(row_data):
            if col_idx < len(row.cells):
                cell = row.cells[col_idx]
                cell.text = cell_text

                set_cell_background(cell, bg_color)
                set_cell_border(cell, top=True, left=True, bottom=True, right=True)

                for para in cell.paragraphs:
                    for run in para.runs:
                        run.font.size = Pt(10)
                        run.font.name = 'Microsoft YaHei'

    # 表格后添加空行
    doc.add_paragraph()


def add_separator(doc):
    """添加分隔线"""
    para = doc.add_paragraph()
    para.paragraph_format.space_before = Pt(6)
    para.paragraph_format.space_after = Pt(6)

    # 添加边框作为分隔线
    pPr = para._element.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')

    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single')
    bottom.set(qn('w:sz'), '6')
    bottom.set(qn('w:space'), '1')
    bottom.set(qn('w:color'), 'CCCCCC')

    pBdr.append(bottom)
    pPr.append(pBdr)


def convert_md_to_docx(md_file, docx_file):
    """将Markdown文件转换为专业排版的Word文档"""
    doc = Document()

    # 设置文档默认样式
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Calibri'
    font.size = Pt(11)
    style._element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑')

    # 设置段落间距
    paragraph_format = style.paragraph_format
    paragraph_format.space_after = Pt(8)
    paragraph_format.line_spacing = 1.15

    # 读取Markdown文件
    with open(md_file, 'r', encoding='utf-8') as f:
        lines = f.readlines()

    i = 0
    in_code_block = False
    code_lines = []
    in_table = False
    table_lines = []

    while i < len(lines):
        line = lines[i].rstrip()

        # 代码块处理
        if line.startswith('```'):
            if in_code_block:
                # 结束代码块
                add_code_block(doc, '\n'.join(code_lines))
                code_lines = []
                in_code_block = False
            else:
                # 开始代码块
                in_code_block = True
            i += 1
            continue

        if in_code_block:
            code_lines.append(line)
            i += 1
            continue

        # 表格处理
        if '|' in line and line.strip().startswith('|'):
            if not in_table:
                in_table = True
                table_lines = []
            table_lines.append(line)
            i += 1

            # 检查表格是否结束
            if i >= len(lines) or '|' not in lines[i] or not lines[i].strip().startswith('|'):
                result = parse_table(table_lines)
                if result:
                    header, data_rows = result
                    add_beautiful_table(doc, header, data_rows)
                in_table = False
                table_lines = []
            continue

        # 空行
        if not line.strip():
            doc.add_paragraph()
            i += 1
            continue

        # 分隔线 ---
        if line.strip() == '---':
            add_separator(doc)
            i += 1
            continue

        # 标题（去除#符号）
        if line.startswith('#'):
            level = len(line) - len(line.lstrip('#'))
            title_text = line.lstrip('#').strip()
            add_title_with_design(doc, title_text, level)
            i += 1
            continue

        # 引用块（去除>符号）
        if line.strip().startswith('>'):
            quote_text = line.strip()[1:].strip()
            add_quote_block(doc, quote_text)
            i += 1
            continue

        # 无序列表（去除- * +符号）
        if re.match(r'^[\s]*[-*+]\s', line):
            text = re.sub(r'^[\s]*[-*+]\s', '', line)
            para = doc.add_paragraph(style='List Bullet')
            add_formatted_text(para, text)
            i += 1
            continue

        # 有序列表（保留数字但优化格式）
        if re.match(r'^[\s]*\d+\.\s', line):
            text = re.sub(r'^[\s]*\d+\.\s', '', line)
            para = doc.add_paragraph(style='List Number')
            add_formatted_text(para, text)
            i += 1
            continue

        # 普通段落
        para = doc.add_paragraph()
        add_formatted_text(para, line.strip())
        i += 1

    # 保存文档
    doc.save(docx_file)
    print(f"✓ {os.path.basename(md_file)} -> {os.path.basename(docx_file)}")


def main():
    """主函数"""
    md_files = sorted([f for f in os.listdir('.') if f.endswith('.md')])
    output_dir = 'newdoc'
    os.makedirs(output_dir, exist_ok=True)

    print(f"找到 {len(md_files)} 个Markdown文件")
    print("开始专业转换...\n")

    for md_file in md_files:
        base_name = os.path.splitext(md_file)[0]
        docx_file = os.path.join(output_dir, f"{base_name}.docx")

        try:
            convert_md_to_docx(md_file, docx_file)
        except Exception as e:
            print(f"✗ {md_file} - {str(e)}")

    print(f"\n✅ 转换完成！专业排版的Word文档已保存到 {output_dir}/ 目录")


if __name__ == '__main__':
    main()
