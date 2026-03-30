#!/usr/bin/env python3
"""
格式化参考文献为Nature格式（顺序编码制）
"""

import re
import sys

def format_reference_nature(ref_text, ref_number):
    """
    将参考文献格式化为Nature格式

    参数:
        ref_text: 原始参考文献文本
        ref_number: 参考文献序号

    返回:
        格式化后的参考文献
    """

    # 移除原有的序号
    ref_text = re.sub(r'^\[\d+\]', '', ref_text).strip()
    ref_text = re.sub(r'^\d+\.', '', ref_text).strip()

    # Nature格式要求:
    # 作者姓名用逗号分隔，最后两个作者用&连接
    # 期刊名斜体
    # 卷号粗体
    # 页码用短横线
    # 年份在括号内

    # 这里返回基本格式，实际格式化需要在XML层面处理斜体和粗体
    return f"{ref_number}. {ref_text}"

def parse_references(text):
    """
    从文本中提取参考文献列表
    """
    references = []

    # 查找参考文献部分
    ref_section = re.search(r'参考文献.*?(?=\n\n|\Z)', text, re.DOTALL)
    if not ref_section:
        return references

    ref_text = ref_section.group(0)

    # 分割各条文献（以数字序号开头）
    ref_items = re.split(r'\n(?=\d+\.|\[\d+\])', ref_text)

    for item in ref_items:
        item = item.strip()
        if item and not item.startswith('参考文献'):
            references.append(item)

    return references

def main():
    if len(sys.argv) < 2:
        print("用法: python format_references.py <输入文件>")
        sys.exit(1)

    input_file = sys.argv[1]

    with open(input_file, 'r', encoding='utf-8') as f:
        text = f.read()

    references = parse_references(text)

    print(f"找到 {len(references)} 条参考文献\n")
    print("格式化后的参考文献（Nature格式）:\n")

    for i, ref in enumerate(references, 1):
        formatted = format_reference_nature(ref, i)
        print(formatted)
        print()

if __name__ == '__main__':
    main()
