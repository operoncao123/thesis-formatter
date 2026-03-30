#!/usr/bin/env python3
"""
将表格转换为三线表格式（北大博士论文标准）

三线表特征：
1. 只有三条线：顶线（粗10pt）、表头下方线（细6pt）、底线（粗10pt）
2. 无竖线
3. 无内部横线
"""

import sys
import xml.etree.ElementTree as ET
from pathlib import Path

# Word XML命名空间
NAMESPACES = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
}

def convert_to_three_line_table(table_element):
    """
    将表格转换为三线表格式

    参数:
        table_element: XML表格元素
    """
    # 获取表格属性
    tbl_pr = table_element.find('w:tblPr', NAMESPACES)
    if tbl_pr is None:
        tbl_pr = ET.SubElement(table_element, '{%s}tblPr' % NAMESPACES['w'])

    # 移除或更新表格边框
    tbl_borders = tbl_pr.find('w:tblBorders', NAMESPACES)
    if tbl_borders is not None:
        tbl_pr.remove(tbl_borders)

    # 创建新的三线表边框
    tbl_borders = ET.SubElement(tbl_pr, '{%s}tblBorders' % NAMESPACES['w'])

    # 顶线：粗线10pt
    top_border = ET.SubElement(tbl_borders, '{%s}top' % NAMESPACES['w'])
    top_border.set('{%s}val' % NAMESPACES['w'], 'single')
    top_border.set('{%s}color' % NAMESPACES['w'], '000000')
    top_border.set('{%s}sz' % NAMESPACES['w'], '10')
    top_border.set('{%s}space' % NAMESPACES['w'], '0')

    # 左边框：无
    left_border = ET.SubElement(tbl_borders, '{%s}left' % NAMESPACES['w'])
    left_border.set('{%s}val' % NAMESPACES['w'], 'none')
    left_border.set('{%s}color' % NAMESPACES['w'], '000000')
    left_border.set('{%s}sz' % NAMESPACES['w'], '0')
    left_border.set('{%s}space' % NAMESPACES['w'], '0')

    # 底线：粗线10pt
    bottom_border = ET.SubElement(tbl_borders, '{%s}bottom' % NAMESPACES['w'])
    bottom_border.set('{%s}val' % NAMESPACES['w'], 'single')
    bottom_border.set('{%s}color' % NAMESPACES['w'], '000000')
    bottom_border.set('{%s}sz' % NAMESPACES['w'], '10')
    bottom_border.set('{%s}space' % NAMESPACES['w'], '0')

    # 右边框：无
    right_border = ET.SubElement(tbl_borders, '{%s}right' % NAMESPACES['w'])
    right_border.set('{%s}val' % NAMESPACES['w'], 'none')
    right_border.set('{%s}color' % NAMESPACES['w'], '000000')
    right_border.set('{%s}sz' % NAMESPACES['w'], '0')
    right_border.set('{%s}space' % NAMESPACES['w'], '0')

    # 内部横线：无
    inside_h = ET.SubElement(tbl_borders, '{%s}insideH' % NAMESPACES['w'])
    inside_h.set('{%s}val' % NAMESPACES['w'], 'none')
    inside_h.set('{%s}color' % NAMESPACES['w'], '000000')
    inside_h.set('{%s}sz' % NAMESPACES['w'], '0')
    inside_h.set('{%s}space' % NAMESPACES['w'], '0')

    # 内部竖线：无
    inside_v = ET.SubElement(tbl_borders, '{%s}insideV' % NAMESPACES['w'])
    inside_v.set('{%s}val' % NAMESPACES['w'], 'none')
    inside_v.set('{%s}color' % NAMESPACES['w'], '000000')
    inside_v.set('{%s}sz' % NAMESPACES['w'], '0')
    inside_v.set('{%s}space' % NAMESPACES['w'], '0')

    # 处理每一行
    rows = table_element.findall('w:tr', NAMESPACES)
    for i, row in enumerate(rows):
        # 获取该行的所有单元格
        cells = row.findall('w:tc', NAMESPACES)

        for cell in cells:
            # 获取或创建单元格属性
            tc_pr = cell.find('w:tcPr', NAMESPACES)
            if tc_pr is None:
                tc_pr = ET.SubElement(cell, '{%s}tcPr' % NAMESPACES['w'])

            # 移除现有边框
            tc_borders = tc_pr.find('w:tcBorders', NAMESPACES)
            if tc_borders is not None:
                tc_pr.remove(tc_borders)

            # 创建新边框
            tc_borders = ET.SubElement(tc_pr, '{%s}tcBorders' % NAMESPACES['w'])

            if i == 0:
                # 表头行：顶线10pt，底线6pt
                top = ET.SubElement(tc_borders, '{%s}top' % NAMESPACES['w'])
                top.set('{%s}val' % NAMESPACES['w'], 'single')
                top.set('{%s}color' % NAMESPACES['w'], '000000')
                top.set('{%s}sz' % NAMESPACES['w'], '10')
                top.set('{%s}space' % NAMESPACES['w'], '0')

                bottom = ET.SubElement(tc_borders, '{%s}bottom' % NAMESPACES['w'])
                bottom.set('{%s}val' % NAMESPACES['w'], 'single')
                bottom.set('{%s}color' % NAMESPACES['w'], '000000')
                bottom.set('{%s}sz' % NAMESPACES['w'], '6')
                bottom.set('{%s}space' % NAMESPACES['w'], '0')

            elif i == 1:
                # 第一行内容：顶线6pt
                top = ET.SubElement(tc_borders, '{%s}top' % NAMESPACES['w'])
                top.set('{%s}val' % NAMESPACES['w'], 'single')
                top.set('{%s}color' % NAMESPACES['w'], '000000')
                top.set('{%s}sz' % NAMESPACES['w'], '6')
                top.set('{%s}space' % NAMESPACES['w'], '0')

            elif i == len(rows) - 1:
                # 最后一行：底线10pt
                bottom = ET.SubElement(tc_borders, '{%s}bottom' % NAMESPACES['w'])
                bottom.set('{%s}val' % NAMESPACES['w'], 'single')
                bottom.set('{%s}color' % NAMESPACES['w'], '000000')
                bottom.set('{%s}sz' % NAMESPACES['w'], '10')
                bottom.set('{%s}space' % NAMESPACES['w'], '0')

            # 中间行不需要边框

def process_document(doc_path):
    """
    处理文档中的所有表格
    """
    tree = ET.parse(doc_path)
    root = tree.getroot()

    # 查找所有表格
    tables = root.findall('.//w:tbl', NAMESPACES)

    print(f"找到 {len(tables)} 个表格")

    for i, table in enumerate(tables, 1):
        print(f"处理表格 {i}...")
        convert_to_three_line_table(table)

    # 保存修改
    tree.write(doc_path, encoding='utf-8', xml_declaration=True)
    print(f"已保存到 {doc_path}")

def main():
    if len(sys.argv) < 2:
        print("用法: python convert_to_three_line_table.py <document.xml路径>")
        sys.exit(1)

    doc_path = Path(sys.argv[1])

    if not doc_path.exists():
        print(f"错误: 文件不存在: {doc_path}")
        sys.exit(1)

    process_document(doc_path)
    print("\n三线表转换完成！")

if __name__ == '__main__':
    main()
