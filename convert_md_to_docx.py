#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Markdown to Word Document Converter
将Markdown文档转换为格式良好的Word文档
"""

import os
import re
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement


def set_cell_border(cell, **kwargs):
    """
    设置表格单元格边框
    """
    tc = cell._element
    tcPr = tc.get_or_add_tcPr()

    tcBorders = OxmlElement('w:tcBorders')
    for edge in ('top', 'left', 'bottom', 'right'):
        edge_data = kwargs.get(edge)
        if edge_data:
            tag = 'w:{}'.format(edge)
            element = OxmlElement(tag)
            element.set(qn('w:val'), 'single')
            element.set(qn('w:sz'), '4')
            element.set(qn('w:space'), '0')
            element.set(qn('w:color'), '000000')
            tcBorders.append(element)

    tcPr.append(tcBorders)


def parse_markdown_line(line):
    """
    解析Markdown行的内联样式（粗体、斜体等）
    """
    # 返回格式化的文本片段列表 [(text, is_bold, is_italic), ...]
    segments = []

    # 简化处理：使用正则表达式识别粗体和斜体
    pattern = r'(\*\*\*([^*]+)\*\*\*|\*\*([^*]+)\*\*|\*([^*]+)\*|`([^`]+)`|[^*`]+)'
    pos = 0

    while pos < len(line):
        # ***粗斜体***
        if line[pos:pos+3] == '***':
            end = line.find('***', pos + 3)
            if end != -1:
                segments.append((line[pos+3:end], True, True, False))
                pos = end + 3
                continue

        # **粗体**
        if line[pos:pos+2] == '**':
            end = line.find('**', pos + 2)
            if end != -1:
                segments.append((line[pos+2:end], True, False, False))
                pos = end + 2
                continue

        # *斜体*
        if line[pos:pos+1] == '*':
            end = line.find('*', pos + 1)
            if end != -1:
                segments.append((line[pos+1:end], False, True, False))
                pos = end + 1
                continue

        # `代码`
        if line[pos:pos+1] == '`':
            end = line.find('`', pos + 1)
            if end != -1:
                segments.append((line[pos+1:end], False, False, True))
                pos = end + 1
                continue

        # 普通文本
        next_special = len(line)
        for char in ['*', '`']:
            idx = line.find(char, pos)
            if idx != -1 and idx < next_special:
                next_special = idx

        if next_special > pos:
            segments.append((line[pos:next_special], False, False, False))
            pos = next_special
        else:
            pos += 1

    return segments


def add_formatted_paragraph(doc, text, style_name=None):
    """
    添加带有内联格式的段落
    """
    para = doc.add_paragraph(style=style_name)
    segments = parse_markdown_line(text)

    for segment_text, is_bold, is_italic, is_code in segments:
        run = para.add_run(segment_text)
        run.bold = is_bold
        run.italic = is_italic

        if is_code:
            run.font.name = 'Consolas'
            run.font.size = Pt(9)
            run.font.color.rgb = RGBColor(199, 37, 78)

    return para


def convert_md_to_docx(md_file, docx_file):
    """
    将Markdown文件转换为Word文档
    """
    # 创建Word文档
    doc = Document()

    # 设置文档默认字体
    doc.styles['Normal'].font.name = 'Calibri'
    doc.styles['Normal']._element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑')
    doc.styles['Normal'].font.size = Pt(11)

    # 读取Markdown文件
    with open(md_file, 'r', encoding='utf-8') as f:
        lines = f.readlines()

    i = 0
    in_code_block = False
    code_block_lines = []
    in_table = False
    table_lines = []
    in_list = False
    list_level = 0

    while i < len(lines):
        line = lines[i].rstrip()

        # 代码块处理
        if line.startswith('```'):
            if in_code_block:
                # 结束代码块
                code_text = '\n'.join(code_block_lines)
                para = doc.add_paragraph(code_text)
                para.style = 'Normal'
                for run in para.runs:
                    run.font.name = 'Consolas'
                    run.font.size = Pt(9)
                    run.font.color.rgb = RGBColor(51, 51, 51)
                para.paragraph_format.left_indent = Inches(0.5)
                para.paragraph_format.space_before = Pt(6)
                para.paragraph_format.space_after = Pt(6)

                code_block_lines = []
                in_code_block = False
            else:
                # 开始代码块
                in_code_block = True
            i += 1
            continue

        if in_code_block:
            code_block_lines.append(line)
            i += 1
            continue

        # 表格处理
        if '|' in line and line.strip().startswith('|'):
            if not in_table:
                in_table = True
                table_lines = []
            table_lines.append(line)
            i += 1

            # 检查是否表格结束
            if i >= len(lines) or '|' not in lines[i] or not lines[i].strip().startswith('|'):
                # 处理表格
                if len(table_lines) > 0:
                    # 过滤掉分隔行
                    header_line = table_lines[0]
                    data_lines = [l for l in table_lines[2:] if not all(c in '|-: ' for c in l)]

                    # 解析表头
                    headers = [h.strip() for h in header_line.split('|')[1:-1]]

                    # 创建表格
                    if data_lines:
                        table = doc.add_table(rows=1, cols=len(headers))
                        table.style = 'Light Grid Accent 1'

                        # 添加表头
                        for idx, header in enumerate(headers):
                            cell = table.rows[0].cells[idx]
                            cell.text = header
                            cell.paragraphs[0].runs[0].font.bold = True
                            set_cell_border(cell, top=True, left=True, bottom=True, right=True)

                        # 添加数据行
                        for data_line in data_lines:
                            cells_data = [c.strip() for c in data_line.split('|')[1:-1]]
                            row = table.add_row()
                            for idx, cell_data in enumerate(cells_data):
                                if idx < len(row.cells):
                                    cell = row.cells[idx]
                                    cell.text = cell_data
                                    set_cell_border(cell, top=True, left=True, bottom=True, right=True)

                        doc.add_paragraph()  # 添加空行

                in_table = False
                table_lines = []
            continue

        # 空行
        if not line.strip():
            if in_list:
                in_list = False
            doc.add_paragraph()
            i += 1
            continue

        # 标题
        if line.startswith('#'):
            if in_list:
                in_list = False

            level = len(line) - len(line.lstrip('#'))
            title_text = line.lstrip('#').strip()

            para = doc.add_heading(title_text, level=min(level, 9))
            i += 1
            continue

        # 无序列表
        if line.strip().startswith(('- ', '* ', '+ ')):
            in_list = True
            text = line.strip()[2:].strip()
            para = add_formatted_paragraph(doc, text)
            para.style = 'List Bullet'
            i += 1
            continue

        # 有序列表
        if re.match(r'^\d+\.\s', line.strip()):
            in_list = True
            text = re.sub(r'^\d+\.\s', '', line.strip())
            para = add_formatted_paragraph(doc, text)
            para.style = 'List Number'
            i += 1
            continue

        # 普通段落
        if in_list:
            in_list = False

        para = add_formatted_paragraph(doc, line.strip())
        i += 1

    # 保存文档
    doc.save(docx_file)
    print(f"✓ 已转换: {os.path.basename(md_file)} -> {os.path.basename(docx_file)}")


def main():
    """
    主函数：转换所有Markdown文件
    """
    # 获取所有.md文件
    md_files = [f for f in os.listdir('.') if f.endswith('.md')]
    md_files.sort()

    # 创建输出目录
    output_dir = 'newdoc'
    os.makedirs(output_dir, exist_ok=True)

    print(f"找到 {len(md_files)} 个Markdown文件")
    print("开始转换...\n")

    # 转换每个文件
    for md_file in md_files:
        # 生成输出文件名
        base_name = os.path.splitext(md_file)[0]
        docx_file = os.path.join(output_dir, f"{base_name}.docx")

        try:
            convert_md_to_docx(md_file, docx_file)
        except Exception as e:
            print(f"✗ 转换失败: {md_file} - {str(e)}")

    print(f"\n转换完成！所有Word文档已保存到 {output_dir}/ 目录")


if __name__ == '__main__':
    main()
