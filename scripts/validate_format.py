#!/usr/bin/env python3
"""
验证博士论文格式是否符合YOUR_UNIVERSITY要求
"""

import sys
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

# Word XML命名空间
NAMESPACES = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
}

def check_page_settings(doc_root):
    """检查页面设置"""
    issues = []

    # 查找sectPr（节属性）
    sect_pr = doc_root.find('.//w:sectPr', NAMESPACES)
    if sect_pr is None:
        issues.append("❌ 未找到页面设置")
        return issues

    # 检查页面大小 (A4: 12240 x 15840 DXA)
    pg_sz = sect_pr.find('w:pgSz', NAMESPACES)
    if pg_sz is not None:
        width = pg_sz.get('{%s}w' % NAMESPACES['w'])
        height = pg_sz.get('{%s}h' % NAMESPACES['w'])
        if width != '12240' or height != '15840':
            issues.append(f"❌ 页面大小不正确: {width}x{height} (应为12240x15840)")
        else:
            print("✓ 页面大小正确 (A4)")
    else:
        issues.append("❌ 未设置页面大小")

    # 检查页边距
    pg_mar = sect_pr.find('w:pgMar', NAMESPACES)
    if pg_mar is not None:
        top = pg_mar.get('{%s}top' % NAMESPACES['w'])
        bottom = pg_mar.get('{%s}bottom' % NAMESPACES['w'])
        left = pg_mar.get('{%s}left' % NAMESPACES['w'])
        right = pg_mar.get('{%s}right' % NAMESPACES['w'])
        header = pg_mar.get('{%s}header' % NAMESPACES['w'])
        footer = pg_mar.get('{%s}footer' % NAMESPACES['w'])

        # 允许小范围误差 (±10 DXA)
        if not (1691 <= int(top or 0) <= 1711):
            issues.append(f"❌ 上边距不正确: {top} (应为1701)")
        else:
            print("✓ 上边距正确 (3.0 cm)")

        if not (1407 <= int(bottom or 0) <= 1427):
            issues.append(f"❌ 下边距不正确: {bottom} (应为1417)")
        else:
            print("✓ 下边距正确 (2.5 cm)")

        if not (1464 <= int(left or 0) <= 1484):
            issues.append(f"❌ 左边距不正确: {left} (应为1474)")
        else:
            print("✓ 左边距正确 (2.6 cm)")

        if not (1464 <= int(right or 0) <= 1484):
            issues.append(f"❌ 右边距不正确: {right} (应为1474)")
        else:
            print("✓ 右边距正确 (2.6 cm)")

        if not (1124 <= int(header or 0) <= 1144):
            issues.append(f"❌ 页眉距边界不正确: {header} (应为1134)")
        else:
            print("✓ 页眉距边界正确 (2.0 cm)")

        if not (982 <= int(footer or 0) <= 1002):
            issues.append(f"❌ 页脚距边界不正确: {footer} (应为992)")
        else:
            print("✓ 页脚距边界正确 (1.75 cm)")
    else:
        issues.append("❌ 未设置页边距")

    return issues

def check_fonts(doc_root):
    """检查字体设置"""
    issues = []
    allowed_fonts = {
        '宋体',
        'SimSun',
        '黑体',
        'SimHei',
        'Times New Roman',
        'Arial',
        'Cambria Math',
        'Symbol',
        'Wingdings',
    }

    # 检查默认字体
    styles = doc_root.findall('.//w:style', NAMESPACES)

    # 统计字体使用情况
    fonts_used = {}
    for rpr in doc_root.findall('.//w:rPr', NAMESPACES):
        font = rpr.find('w:rFonts', NAMESPACES)
        if font is not None:
            ascii_font = font.get('{%s}ascii' % NAMESPACES['w'])
            eastasia_font = font.get('{%s}eastAsia' % NAMESPACES['w'])
            if ascii_font:
                fonts_used[ascii_font] = fonts_used.get(ascii_font, 0) + 1
            if eastasia_font:
                fonts_used[eastasia_font] = fonts_used.get(eastasia_font, 0) + 1

    print("\n字体使用统计:")
    for font, count in sorted(fonts_used.items(), key=lambda x: x[1], reverse=True)[:10]:
        print(f"  {font}: {count}次")

    # 检查是否使用了宋体和Times New Roman
    if '宋体' not in fonts_used and 'SimSun' not in fonts_used:
        issues.append("⚠️  未检测到宋体使用（正文应使用宋体）")

    if 'Times New Roman' not in fonts_used:
        issues.append("⚠️  未检测到Times New Roman使用（英文数字应使用Times New Roman）")

    unexpected_fonts = sorted(font for font in fonts_used if font not in allowed_fonts)
    if unexpected_fonts:
        issues.append(f"⚠️  检测到非标准字体残留: {', '.join(unexpected_fonts)}")

    return issues

def check_headings(doc_root):
    """检查标题格式"""
    issues = []

    # 查找所有段落
    paragraphs = doc_root.findall('.//w:p', NAMESPACES)

    heading_count = {'Heading1': 0, 'Heading2': 0, 'Heading3': 0}

    for para in paragraphs:
        ppr = para.find('w:pPr', NAMESPACES)
        if ppr is not None:
            pstyle = ppr.find('w:pStyle', NAMESPACES)
            if pstyle is not None:
                style_val = pstyle.get('{%s}val' % NAMESPACES['w'])
                if style_val in heading_count:
                    heading_count[style_val] += 1

    print(f"\n标题统计:")
    print(f"  一级标题 (章): {heading_count['Heading1']}个")
    print(f"  二级标题 (节): {heading_count['Heading2']}个")
    print(f"  三级标题: {heading_count['Heading3']}个")

    if heading_count['Heading1'] == 0:
        issues.append("⚠️  未检测到章标题（Heading1）")

    if heading_count['Heading3'] > 0:
        issues.append(f"⚠️  检测到{heading_count['Heading3']}个三级标题（一般不建议使用三级及以上标题）")

    return issues

def check_figures_tables(doc_root):
    """检查图表"""
    issues = []

    # 统计图片数量
    drawings = doc_root.findall('.//w:drawing', NAMESPACES)
    print(f"\n图片数量: {len(drawings)}个")

    # 查找图注（通常包含"图"字）
    figure_captions = []
    table_captions = []

    for para in doc_root.findall('.//w:p', NAMESPACES):
        text = ''.join(para.itertext())
        if text.strip().startswith('图') and any(c.isdigit() for c in text[:10]):
            figure_captions.append(text.strip()[:50])
        elif text.strip().startswith('表') and any(c.isdigit() for c in text[:10]):
            table_captions.append(text.strip()[:50])

    print(f"图注数量: {len(figure_captions)}个")
    if figure_captions:
        print("  示例:", figure_captions[0] if figure_captions else "无")

    print(f"表注数量: {len(table_captions)}个")
    if table_captions:
        print("  示例:", table_captions[0] if table_captions else "无")

    if len(drawings) > 0 and len(figure_captions) == 0:
        issues.append("⚠️  有图片但未检测到图注")

    return issues

def check_references(doc_root):
    """检查参考文献"""
    issues = []

    # 查找参考文献部分
    in_references = False
    ref_count = 0
    ref_samples = []

    for para in doc_root.findall('.//w:p', NAMESPACES):
        text = ''.join(para.itertext()).strip()

        if '参考文献' in text:
            in_references = True
            continue

        if in_references:
            # 检测是否是参考文献条目（以数字开头）
            if text and (text[0].isdigit() or (len(text) > 1 and text[1].isdigit())):
                ref_count += 1
                if ref_count <= 3:
                    ref_samples.append(text[:100])

            # 如果遇到新的章节标题，停止统计
            if text.startswith('附录') or text.startswith('致谢'):
                break

    print(f"\n参考文献数量: {ref_count}条")
    if ref_samples:
        print("参考文献示例:")
        for i, sample in enumerate(ref_samples, 1):
            print(f"  [{i}] {sample}...")

    if ref_count == 0:
        issues.append("⚠️  未检测到参考文献")

    return issues

def main():
    if len(sys.argv) < 2:
        print("用法: python validate_format.py <论文文件.docx>")
        sys.exit(1)

    docx_path = Path(sys.argv[1])

    if not docx_path.exists():
        print(f"错误: 文件不存在: {docx_path}")
        sys.exit(1)

    print(f"正在验证: {docx_path.name}")
    print("=" * 60)

    all_issues = []

    try:
        # 解压docx文件并读取document.xml
        with zipfile.ZipFile(docx_path, 'r') as zip_ref:
            with zip_ref.open('word/document.xml') as xml_file:
                tree = ET.parse(xml_file)
                root = tree.getroot()

        # 执行各项检查
        print("\n【页面设置检查】")
        all_issues.extend(check_page_settings(root))

        print("\n【字体检查】")
        all_issues.extend(check_fonts(root))

        print("\n【标题检查】")
        all_issues.extend(check_headings(root))

        print("\n【图表检查】")
        all_issues.extend(check_figures_tables(root))

        print("\n【参考文献检查】")
        all_issues.extend(check_references(root))

        # 输出总结
        print("\n" + "=" * 60)
        if all_issues:
            print(f"\n发现 {len(all_issues)} 个问题:")
            for issue in all_issues:
                print(f"  {issue}")
        else:
            print("\n✓ 所有检查项通过！")

        print("\n注意: 此脚本只能检查部分格式要求，完整检查请参考自查清单。")

    except Exception as e:
        print(f"\n错误: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)

if __name__ == '__main__':
    main()
