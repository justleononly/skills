#!/usr/bin/env python3
"""
IP销售分析报告结构校验器

功能：校验生成的 .docx 报告结构与「新IP分析方法」模板的结构一致性
用法：python3 validate_report.py --report <报告文件.docx>

返回：
- 结构完整性检查结果
- 缺少/多余的章节列表
"""

import argparse
import os
import sys
import re
from docx import Document


# 模板中定义的标准章节结构
EXPECTED_STRUCTURE = [
    ("一级标题", "一、关键结论"),
    ("一级标题", "二、基础销售数据"),
    ("一级标题", "三、销售数据趋势"),
    ("一级标题", "四、付费用户画像"),
    ("二级标题", "4.1 付费用户性别分布"),
    ("二级标题", "4.2 付费用户年龄分布"),
    ("二级标题", "4.3 付费用户性别&年龄交叉分布"),
    ("二级标题", "4.4 付费用户会员等级分布"),
    ("一级标题", "五、分析总结与建议"),
]


def extract_headings(doc_path):
    """从docx中提取所有标题文本"""
    doc = Document(doc_path)
    headings = []
    for para in doc.paragraphs:
        if para.style.name.startswith("Heading"):
            level = 0
            if para.style.name == "Heading 1":
                level = 1
            elif para.style.name == "Heading 2":
                level = 2
            elif para.style.name == "Heading 0":
                level = 0
            text = para.text.strip()
            if text:
                headings.append((level, text))
    return headings


def check_structure(headings, expected):
    """对比实际标题与预期结构"""
    actual_texts = [h[1] for h in headings]
    found = []
    missing = []
    extra = []

    for level, exp_text in expected:
        matched = False
        for h_text in actual_texts:
            if exp_text in h_text or h_text in exp_text:
                matched = True
                found.append((level, exp_text))
                break
        if not matched:
            missing.append((level, exp_text))

    # 检查多余的标题（非标准章节）
    standard_set = set(e[1] for e in expected)
    for h in headings:
        is_standard = False
        for std in standard_set:
            if std in h[1] or h[1] in std:
                is_standard = True
                break
        if not is_standard and h[0] >= 1:
            extra.append(h)

    return found, missing, extra


def main():
    parser = argparse.ArgumentParser(description="IP销售分析报告结构校验器")
    parser.add_argument("--report", required=True, help="生成的.docx报告文件路径")
    args = parser.parse_args()

    if not os.path.isfile(args.report):
        print(f"[错误] 找不到报告文件：{args.report}")
        sys.exit(1)

    print("=" * 60)
    print("  IP销售分析报告 - 结构一致性校验")
    print("=" * 60)
    print(f"📄 报告文件：{args.report}")
    print()

    headings = extract_headings(args.report)
    print(f"📋 检测到 {len(headings)} 个标题章节")
    for level, text in headings:
        prefix = "  " if level >= 2 else ""
        print(f"  {prefix}[Level {level}] {text}")
    print()

    found, missing, extra = check_structure(headings, EXPECTED_STRUCTURE)

    # 缺失检查
    if missing:
        print("❌ 缺少的章节：")
        for level, text in missing:
            label = "一级标题" if level == 1 else "二级标题"
            print(f"  - [{label}] {text}")
    else:
        print("✅ 所有必需章节均已包含，无缺失")
    print()

    # 多余检查
    if extra:
        print("⚠️  检测到额外章节（非模板标准部分，可能为新增内容）：")
        for level, text in extra:
            print(f"  + [{level}级] {text}")
    else:
        print("✅ 无多余章节")
    print()

    # 结构完整率
    total_expected = len(EXPECTED_STRUCTURE)
    total_found = len(found)
    completeness = total_found / total_expected * 100
    print(f"📊 结构完整率：{completeness:.1f}%（{total_found}/{total_expected}）")
    print()

    if completeness >= 100:
        print("🎉 结论：报告结构完整，与模板完全一致！")
        return 0
    elif completeness >= 80:
        print("⚠️  结论：报告结构基本完整，建议补充缺失章节。")
        return 1
    else:
        print("❌ 结论：报告结构缺失较多，请检查模板要求。")
        return 2


if __name__ == "__main__":
    sys.exit(main())
