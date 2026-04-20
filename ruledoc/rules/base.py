"""
FormatRule 基类 - 规则接口定义

设计原则:
- 接口优先: 先定义抽象方法，后实现
- 默认值合理: 提供通用默认实现
- 验证机制: 包含配置验证方法
"""

import logging
from abc import ABC, abstractmethod
from typing import TYPE_CHECKING, Dict, List, Optional, Tuple, Type

if TYPE_CHECKING:
    from ruledoc.context import ProcessingContext
    from ruledoc.processors.base import PostProcessor


_registered_rules: Dict[str, Type["FormatRule"]] = {}


def register_rule(cls: Type["FormatRule"]) -> Type["FormatRule"]:
    """
    规则注册装饰器

    将规则类注册到全局注册表，支持通过名称加载。

    Args:
        cls: FormatRule 子类

    Returns:
        原类 (装饰器模式)

    Example:
        >>> @register_rule
        ... class YZUFormatRule(FormatRule):
        ...     @property
        ...     def name(self) -> str:
        ...         return 'yzu'
    """
    try:
        instance = cls()
        _registered_rules[instance.name] = cls
        logging.debug(f"规则已注册: {instance.name}")
    except Exception as e:
        logging.warning(f"规则注册失败: {cls.__name__} - {e}")
    return cls


def load_rule(rule_name: str) -> Optional["FormatRule"]:
    """
    加载规则

    从注册表加载规则实例，支持动态导入和别名映射。

    Args:
        rule_name: 规则名称或别名

    Returns:
        规则实例，不存在则返回 None

    Example:
        >>> rule = load_rule('yzu_thesis')
        >>> print(rule.description)
        扬州大学毕业论文格式

        >>> # 向后兼容别名
        >>> rule = load_rule('yzu')
        >>> print(rule.name)
        yzu_thesis
    """
    _RULE_ALIASES = {
        "yzu": "yzu_thesis",
    }

    resolved_name = _RULE_ALIASES.get(rule_name, rule_name)

    if resolved_name in _registered_rules:
        return _registered_rules[resolved_name]()

    try:
        import importlib

        module = importlib.import_module(f".{resolved_name}", package="ruledoc.rules")
        if resolved_name in _registered_rules:
            return _registered_rules[resolved_name]()
    except ImportError:
        pass

    return None


def list_available_rules() -> List[str]:
    """
    列出所有可用规则

    Returns:
        已注册规则名称列表 (按字母排序)

    Example:
        >>> rules = list_available_rules()
        >>> print(rules)
        ['yzu']
    """
    return sorted(_registered_rules.keys())


class FormatRule(ABC):
    """
    格式规则抽象基类

    所有学校格式规则必须继承此类并实现抽象方法。
    规则定义了论文格式的所有可配置项。

    职责:
    - 定义页面设置 (边距、装订线等)
    - 定义字体设置 (正文字体、标题字体等)
    - 定义标题格式 (字号、对齐方式)
    - 检测段落类型 (标题、题注、正文等)
    - 格式化段落 (应用字体、缩进等)

    不负责:
    - 流程控制 (由 Formatter 负责)
    - 文档转换 (由 PostProcessor 负责)

    Example:
        >>> @register_rule
        ... class YZUFormatRule(FormatRule):
        ...     @property
        ...     def name(self) -> str:
        ...         return 'yzu'
        ...
        ...     @property
        ...     def description(self) -> str:
        ...         return '扬州大学毕业论文格式'
        ...
        ...     def get_page_settings(self) -> Dict[str, float]:
        ...         return {'top_margin': 2.2, 'bottom_margin': 2.2}
    """

    DEFAULT_PAGE_SETTINGS = {
        "top_margin": 2.2,
        "bottom_margin": 2.2,
        "left_margin": 2.5,
        "right_margin": 2.0,
        "gutter": 0.5,
    }

    @property
    @abstractmethod
    def name(self) -> str:
        """
        规则名称 (唯一标识)

        Returns:
            规则名称，如 'yzu'
        """
        pass

    @property
    @abstractmethod
    def description(self) -> str:
        """
        规则描述

        Returns:
            规则描述，如 '扬州大学毕业论文格式'
        """
        pass

    @abstractmethod
    def get_page_settings(self) -> Dict[str, float]:
        """
        获取页面设置

        Returns:
            页面设置字典，包含:
            - top_margin: 上边距 (cm)
            - bottom_margin: 下边距 (cm)
            - left_margin: 左边距 (cm)
            - right_margin: 右边距 (cm)
            - gutter: 装订线 (cm)
        """
        pass

    @abstractmethod
    def get_font_settings(self) -> Dict[str, str]:
        """
        获取字体设置

        Returns:
            字体设置字典，包含:
            - body_font: 正文字体
            - heading_font: 标题字体
            - caption_font: 题注字体
        """
        pass

    @abstractmethod
    def get_heading_format(self, level: int) -> Tuple[str, str, float]:
        """
        获取标题格式

        Args:
            level: 标题层级 (1-4)

        Returns:
            元组 (字体名称, 对齐方式, 字号)
            - 字体名称: 如 '黑体'
            - 对齐方式: 'left', 'center', 'right'
            - 字号: Word 字号值
        """
        pass

    def get_style_map(self) -> Dict[str, str]:
        """
        获取样式名映射

        Returns:
            样式名映射字典，键为 Pandoc 样式名，值为 Word 样式名
        """
        return {
            "heading 1": "Heading 1",
            "heading 2": "Heading 2",
            "heading 3": "Heading 3",
            "heading 4": "Heading 4",
        }

    def get_pre_processors(self) -> List["PostProcessor"]:
        """
        获取规则前置处理器

        Returns:
            前置处理器列表 (在通用处理器之前执行)
        """
        return []

    def get_post_processors(self) -> List["PostProcessor"]:
        """
        获取规则后置处理器

        Returns:
            后置处理器列表 (在通用处理器之后执行)
        """
        return []

    def get_processor_config(self) -> Dict:
        """
        获取处理器配置

        Returns:
            处理器配置字典:
            - style: 是否启用 StyleProcessor
            - heading: 是否启用 HeadingProcessor
            - caption: 是否启用 CaptionProcessor
            - custom_order: 自定义处理器顺序 (None 使用默认顺序)
        """
        return {
            "style": True,
            "heading": True,
            "caption": True,
            "custom_order": None,
        }

    def detect_paragraph_type(self, para, ctx: "ProcessingContext") -> str:
        """
        检测段落类型

        Args:
            para: 段落对象
            ctx: 处理上下文

        Returns:
            段落类型:
            - 'heading': 标题
            - 'caption': 题注
            - 'abstract': 摘要
            - 'references': 参考文献
            - 'body': 正文
        """
        return "body"

    def format_paragraph(self, para, para_type: str, ctx: "ProcessingContext") -> None:
        """
        格式化段落

        Args:
            para: 段落对象
            para_type: 段落类型
            ctx: 处理上下文
        """
        pass

    def format_references_section(self, paragraphs, ctx: "ProcessingContext") -> None:
        """
        格式化参考文献章节

        子类可以重写此方法实现自定义的参考文献格式化。

        Args:
            paragraphs: 段落列表
            ctx: 处理上下文
        """
        pass

    def scan_document_structure(self, paragraphs, ctx: "ProcessingContext") -> None:
        """
        扫描文档结构（可选）

        子类可以重写此方法来扫描文档并设置上下文标志。
        默认实现为空，由 Formatter._update_section_flags 处理。

        Args:
            paragraphs: 段落列表
            ctx: 处理上下文
        """
        pass

    def get_header_footer_settings(self) -> Dict:
        """
        获取页眉页脚设置

        Returns:
            页眉页脚设置字典:
            - header_text: 页眉文本
            - header_font: 页眉字体
            - header_font_size: 页眉字号
            - footer_type: 页脚类型 ('page_number', 'none')
        """
        return {
            "header_text": "",
            "header_font": "宋体",
            "header_font_size": 9,
            "footer_type": "page_number",
        }

    def get_validated_page_settings(self) -> Dict[str, float]:
        """
        获取验证后的页面设置

        验证所有页面设置值，无效值使用默认值替代。

        Returns:
            验证后的页面设置字典
        """
        settings = self.get_page_settings()
        validated = {}
        for key, default in self.DEFAULT_PAGE_SETTINGS.items():
            value = settings.get(key, default)
            if not isinstance(value, (int, float)) or value < 0:
                logging.warning(f"[{self.name}] Invalid {key}={value}, using default {default}")
                value = default
            validated[key] = value
        return validated
