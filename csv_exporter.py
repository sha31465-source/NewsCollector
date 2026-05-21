#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
CSV导出模块 - NewsCollector V36

功能：
1. 同步生成CSV文件（与TXT配套）
2. UTF-8编码，兼容中文
3. 标准CSV转义（逗号、双引号、换行等）
4. 与TXT同名、同目录、同步保存

作者：Claude Code
版本：V1.0
日期：2026-04-25
"""

import csv
import os
import re
from typing import List, Dict, Any
from datetime import datetime


class CSVExporter:
    """CSV导出器 - 严格遵循CSV标准格式"""

    # CSV表头（固定）
    HEADERS = ['序号', '标题', '来源', '标签', '总结', 'URL', '日期']

    def __init__(self, encoding: str = 'utf-8'):
        """
        初始化CSV导出器

        Args:
            encoding: 文件编码，默认utf-8（支持BOM以兼容Excel）
        """
        self.encoding = encoding

    def escape_csv_field(self, field: Any) -> str:
        """
        CSV字段转义 - 严格遵循RFC 4180标准

        转义规则：
        1. 如果包含双引号，将每个双引号替换为两个双引号（""）
        2. 如果包含逗号、双引号、换行符，用双引号包裹整个字段
        3. 其他情况直接返回原字符串

        Args:
            field: 字段值（任意类型）

        Returns:
            str: 转义后的字符串
        """
        if field is None:
            return ''

        # 转换为字符串
        text = str(field).strip()

        # 处理特殊字符
        # 1. 双引号替换为两个双引号
        text = text.replace('"', '""')

        # 2. 检查是否需要用双引号包裹
        # 包含以下任一字符就需要包裹：逗号、双引号、换行符
        if any(char in text for char in [',', '"', '\n', '\r']):
            text = f'"{text}"'

        return text

    def normalize_text(self, text: str, max_length: int = 10000) -> str:
        """
        文本标准化处理

        处理：
        1. 移除控制字符（保留换行符）
        2. 统一换行符为\\n
        3. 限制最大长度（防止CSV字段过长）

        Args:
            text: 原文本
            max_length: 最大长度

        Returns:
            str: 标准化后的文本
        """
        if not text:
            return ''

        # 移除除换行外的控制字符
        text = re.sub(r'[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]', '', text)

        # 统一换行符
        text = text.replace('\r\n', '\n').replace('\r', '\n')

        # 限制长度
        if len(text) > max_length:
            text = text[:max_length] + '...[截断]'

        return text.strip()

    def extract_item_data(self, item: Dict[str, Any], index: int) -> Dict[str, str]:
        """
        从资讯字典中提取CSV所需数据

        Args:
            item: 资讯字典
            index: 序号

        Returns:
            Dict[str, str]: CSV行数据
        """
        # 提取基本信息
        title = self.normalize_text(item.get('title', '无标题'))

        source = self.normalize_text(item.get('source', '未知'))

        # 标签：优先使用category_name，其次category_key，最后'未分类'
        category = self.normalize_text(
            item.get('category_name') or
            item.get('category_key') or
            item.get('category', '未分类')
        )

        # 总结：优先使用summary字段（AI生成的内容总结），其次使用content摘要
        summary = item.get('summary', '')

        if not summary:
            # 如果没有summary字段，使用content的前300字符作为后备
            content = item.get('content', '')
            if content:
                summary = self.normalize_text(content[:300] + ('...' if len(content) > 300 else ''))
            else:
                summary = '无总结'
        else:
            # 使用summary字段并标准化
            summary = self.normalize_text(summary)

        # URL：优先使用link，其次使用url
        url = self.normalize_text(item.get('link') or item.get('url', ''))

        # 日期：兼容多种字段名
        date_str = item.get('pub_date') or item.get('published_date') or item.get('date', '未知')
        date_str = self.normalize_text(date_str)

        return {
            '序号': str(index),
            '标题': title,
            '来源': source,
            '标签': category,
            '总结': summary,
            'URL': url,
            '日期': date_str
        }

    def generate_csv_content(self, items: List[Dict[str, Any]]) -> str:
        """
        生成CSV内容（字符串形式）

        Args:
            items: 资讯列表

        Returns:
            str: CSV格式的内容
        """
        if not items:
            # 空数据也要写表头
            lines = [','.join(self.HEADERS)]
            return '\n'.join(lines)

        lines = []

        # 写入表头（UTF-8 BOM标识，让Excel正确识别编码）
        lines.append(','.join(self.HEADERS))

        # 写入数据行
        for idx, item in enumerate(items, start=1):
            # 提取数据
            row_data = self.extract_item_data(item, idx)

            # 转义每个字段
            escaped_fields = [self.escape_csv_field(row_data[header]) for header in self.HEADERS]

            # 拼接成一行
            line = ','.join(escaped_fields)
            lines.append(line)

        return '\n'.join(lines)

    def save_csv(self, items: List[Dict[str, Any]], output_path: str,
                 with_bom: bool = True) -> str:
        """
        保存CSV文件

        Args:
            items: 资讯列表
            output_path: 输出文件路径
            with_bom: 是否添加UTF-8 BOM（Excel兼容性）

        Returns:
            str: 实际保存的文件路径
        """
        try:
            # 生成CSV内容
            csv_content = self.generate_csv_content(items)

            # 确保输出目录存在
            os.makedirs(os.path.dirname(output_path), exist_ok=True)

            # 写入文件
            encoding = 'utf-8-sig' if with_bom else self.encoding
            with open(output_path, 'w', encoding=encoding, newline='') as f:
                f.write(csv_content)

            return output_path

        except Exception as e:
            raise IOError(f"CSV文件保存失败: {e}")

    def save_csv_with_python_csv_module(self, items: List[Dict[str, Any]],
                                        output_path: str,
                                         with_bom: bool = True) -> str:
        """
        使用Python标准csv模块保存（更标准的实现）

        Args:
            items: 资讯列表
            output_path: 输出文件路径
            with_bom: 是否添加UTF-8 BOM（Excel兼容性）

        Returns:
            str: 实际保存的文件路径
        """
        try:
            # 确保输出目录存在
            os.makedirs(os.path.dirname(output_path), exist_ok=True)

            # 写入文件
            encoding = 'utf-8-sig' if with_bom else self.encoding
            with open(output_path, 'w', encoding=encoding, newline='') as f:
                writer = csv.writer(f, quoting=csv.QUOTE_MINIMAL)

                # 写入表头
                writer.writerow(self.HEADERS)

                # 写入数据行
                for idx, item in enumerate(items, start=1):
                    row_data = self.extract_item_data(item, idx)
                    row = [row_data[header] for header in self.HEADERS]
                    writer.writerow(row)

            return output_path

        except Exception as e:
            raise IOError(f"CSV文件保存失败: {e}")


def generate_csv_sync(txt_path: str, items: List[Dict[str, Any]]) -> str:
    """
    同步生成CSV文件（与TXT配套）

    根据TXT文件路径自动生成同名CSV文件

    Args:
        txt_path: TXT文件路径
        items: 资讯列表

    Returns:
        str: CSV文件路径

    Example:
        >>> txt_path = "M:\\output\\资讯报告_2026-04-25_1200.txt"
        >>> csv_path = generate_csv_sync(txt_path, items)
        >>> # csv_path = "M:\\output\\资讯报告_2026-04-25_1200.csv"
    """
    # 生成CSV文件路径（将.txt替换为.csv）
    csv_path = txt_path.replace('.txt', '.csv')

    # 创建导出器
    exporter = CSVExporter()

    # 保存CSV（使用标准csv模块，带BOM以兼容Excel）
    exporter.save_csv_with_python_csv_module(items, csv_path, with_bom=True)

    return csv_path


# ============================================
# 便捷函数
# ============================================

def quick_export_csv(items: List[Dict[str, Any]], output_path: str) -> str:
    """
    快速导出CSV（简化版）

    Args:
        items: 资讯列表
        output_path: 输出文件路径（.csv）

    Returns:
        str: 实际保存的文件路径
    """
    exporter = CSVExporter()
    return exporter.save_csv_with_python_csv_module(items, output_path, with_bom=True)


# ============================================
# 测试代码
# ============================================

if __name__ == '__main__':
    # 测试数据
    test_items = [
        {
            'title': 'AI辅助诊断肺癌的新突破',
            'source': 'Nature Medicine',
            'category_key': 'lung_cancer',
            'category_name': '肺癌前沿',
            'analysis': '该研究提出了一种基于深度学习的AI诊断系统，能够准确识别早期肺癌。',
            'content': '研究团队开发了一种新型AI算法...',
            'pub_date': '2026-04-25'
        },
        {
            'title': '含有"逗号"，和"双引号"的测试标题',
            'source': 'Test Source',
            'category_name': '测试分类',
            'analysis': '这是一条测试包含特殊字符的内容：\n第一行\n第二行\n"引号内容"',
            'content': '测试内容',
            'pub_date': '2026-04-25'
        },
        {
            'title': '营养与肺癌康复研究',
            'source': 'JAMA Oncology',
            'category_key': 'nutrition_diet',
            'category_name': '营养与膳食研究',
            'analysis': '研究发现，地中海饮食可降低肺癌复发风险。',
            'content': '这是一段很长的内容...' * 100,
            'pub_date': '2026-04-24'
        }
    ]

    # 测试导出
    exporter = CSVExporter()

    # 方法1：手动生成内容
    csv_content = exporter.generate_csv_content(test_items)
    print("=== CSV内容 ===")
    print(csv_content)
    print("\n")

    # 方法2：保存文件
    test_output = 'M:\\claude\\生成文件\\信息搜集\\测试导出.csv'
    try:
        path = exporter.save_csv_with_python_csv_module(test_items, test_output)
        print(f"✓ CSV文件已保存: {path}")
        print(f"✓ 文件大小: {os.path.getsize(path)} 字节")
    except Exception as e:
        print(f"✗ 保存失败: {e}")

    # 测试同步生成
    txt_file = 'M:\\claude\\生成文件\\信息搜集\\资讯报告_2026-04-25_1200.txt'
    try:
        csv_path = generate_csv_sync(txt_file, test_items)
        print(f"✓ 同步生成CSV: {csv_path}")
    except Exception as e:
        print(f"✗ 同步生成失败: {e}")
