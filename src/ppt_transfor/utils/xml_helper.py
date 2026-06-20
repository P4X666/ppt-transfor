"""XML 辅助工具：直接操作 lxml 元素处理 python-pptx 未暴露的属性。

python-pptx 部分属性（如阴影细节、渐变 stops、自定义几何路径）未直接暴露，
通过 shape._element 直接操作底层 XML。
"""

from __future__ import annotations

from lxml import etree


def get_element_xml(element) -> str:
    """获取元素的 XML 字符串（调试用）。"""
    return etree.tostring(element, pretty_print=True, encoding="unicode")


def find_child(element, tag: str):
    """按标签名查找直接子元素（命名空间自动匹配）。"""
    for child in element:
        if etree.QName(child).localname == tag:
            return child
    return None


def find_children(element, tag: str):
    """按标签名查找所有直接子元素。"""
    return [c for c in element if etree.QName(c).localname == tag]


def ensure_child(element, tag: str, nsmap: dict | None = None):
    """确保子元素存在，不存在则创建并追加。返回该子元素。"""
    child = find_child(element, tag)
    if child is None:
        # 使用 pptx 默认命名空间
        ns = "http://schemas.openxmlformats.org/drawingml/2006/main"
        child = etree.SubElement(element, f"{{{ns}}}{tag}")
    return child


def remove_child(element, tag: str) -> bool:
    """移除指定标签的子元素，返回是否移除成功。"""
    child = find_child(element, tag)
    if child is not None:
        element.remove(child)
        return True
    return False
