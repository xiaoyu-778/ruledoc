"""
RuleDoc 全局配置模块

集中管理所有配置项，支持:
- 默认配置值
- 配置验证
- 配置导出
"""

import os
from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional, Set


@dataclass
class FontConfig:
    """字体配置"""

    chinese: str = "宋体"
    english: str = "Times New Roman"
    heading: str = "黑体"
    abstract_thesis: str = "楷体"
    abstract_design: str = "宋体"
    signature: str = "仿宋"
    code: str = "Consolas"


@dataclass
class FontSizeConfig:
    """字号配置（单位：磅）"""

    title: float = 18.0
    heading_1: float = 15.0
    heading_2: float = 14.0
    heading_3: float = 12.0
    heading_4: float = 12.0
    body: float = 12.0
    caption: float = 10.5
    header_footer: float = 9.0
    code: float = 9.0


@dataclass
class PageConfig:
    """页面配置（单位：厘米）"""

    top_margin: float = 2.2
    bottom_margin: float = 2.2
    left_margin: float = 2.5
    right_margin: float = 2.0
    gutter: float = 0.5
    header_distance: float = 1.2
    footer_distance: float = 1.5


@dataclass
class ParagraphConfig:
    """段落配置"""

    line_spacing: float = 1.5
    first_line_indent_cm: float = 0.74
    heading_space_before_lines: int = 100
    heading_space_after_lines: int = 100


@dataclass
class TabStopConfig:
    """制表位配置（单位：厘米）"""

    center: float = 7.5
    right: float = 14.5


@dataclass
class TableConfig:
    """表格配置"""

    top_border_pt: float = 1.5
    header_border_pt: float = 0.75
    bottom_border_pt: float = 1.5
    cell_space_pt: float = 3.0


@dataclass
class SpecialTitlesConfig:
    """特殊标题配置"""

    abstract: Set[str] = field(default_factory=lambda: {"摘要", "abstract"})
    keywords: Set[str] = field(default_factory=lambda: {"关键词", "keywords"})
    contents: Set[str] = field(default_factory=lambda: {"目录", "contents"})
    references: Set[str] = field(default_factory=lambda: {"参考文献", "references"})
    acknowledgements: Set[str] = field(
        default_factory=lambda: {"致谢", "acknowledgements", "acknowledgement"}
    )
    appendix: Set[str] = field(default_factory=lambda: {"附录", "appendix"})
    introduction: Set[str] = field(default_factory=lambda: {"引言", "introduction"})
    conclusion: Set[str] = field(default_factory=lambda: {"结论", "conclusion"})

    @property
    def all(self) -> Set[str]:
        """获取所有特殊标题"""
        result = set()
        for attr in [
            "abstract",
            "keywords",
            "contents",
            "references",
            "acknowledgements",
            "appendix",
            "introduction",
            "conclusion",
        ]:
            result.update(getattr(self, attr))
        return result


@dataclass
class SignatureKeywordsConfig:
    """署名关键词配置"""

    items: List[str] = field(
        default_factory=lambda: [
            "年级专业",
            "学生姓名",
            "指导教师",
            "学院",
            "学号",
            "届别",
            "专业班级",
            "完成日期",
            "答辩日期",
        ]
    )


@dataclass
class NumberingConfig:
    """编号配置"""

    multilevel_heading_id: int = 100
    reference_list_id: int = 99


@dataclass
class CodeBlockConfig:
    """代码块配置"""

    background_color: str = "F5F5F5"
    left_indent_cm: float = 0.5
    right_indent_cm: float = 0.5


@dataclass
class FileConfig:
    """文件配置"""

    supported_input_formats: Set[str] = field(default_factory=lambda: {".md", ".markdown", ".docx"})
    max_file_size_mb: int = 100
    temp_file_prefix: str = "ruledoc_"


@dataclass
class PandocConfig:
    """Pandoc 配置"""

    version_check_timeout: int = 10
    convert_timeout: int = 60
    download_url: str = "https://pandoc.org/installing.html"


@dataclass
class RuleAliasesConfig:
    """规则别名配置"""

    aliases: Dict[str, str] = field(
        default_factory=lambda: {
            "yzu": "yzu_thesis",
        }
    )


class Config:
    """
    全局配置类

    集中管理所有配置项，提供统一的访问接口。

    Usage:
        config = Config()
        print(config.fonts.chinese)  # '宋体'
        print(config.font_sizes.body)  # 12.0
    """

    _instance: Optional["Config"] = None

    def __new__(cls):
        if cls._instance is None:
            cls._instance = super().__new__(cls)
            cls._instance._init_configs()
        return cls._instance

    def _init_configs(self) -> None:
        """初始化所有配置"""
        self._fonts = FontConfig()
        self._font_sizes = FontSizeConfig()
        self._page = PageConfig()
        self._paragraph = ParagraphConfig()
        self._tab_stops = TabStopConfig()
        self._table = TableConfig()
        self._special_titles = SpecialTitlesConfig()
        self._signature_keywords = SignatureKeywordsConfig()
        self._numbering = NumberingConfig()
        self._code_block = CodeBlockConfig()
        self._file = FileConfig()
        self._pandoc = PandocConfig()
        self._rule_aliases = RuleAliasesConfig()

    @property
    def fonts(self) -> FontConfig:
        """字体配置"""
        return self._fonts

    @property
    def font_sizes(self) -> FontSizeConfig:
        """字号配置"""
        return self._font_sizes

    @property
    def page(self) -> PageConfig:
        """页面配置"""
        return self._page

    @property
    def paragraph(self) -> ParagraphConfig:
        """段落配置"""
        return self._paragraph

    @property
    def tab_stops(self) -> TabStopConfig:
        """制表位配置"""
        return self._tab_stops

    @property
    def table(self) -> TableConfig:
        """表格配置"""
        return self._table

    @property
    def special_titles(self) -> SpecialTitlesConfig:
        """特殊标题配置"""
        return self._special_titles

    @property
    def signature_keywords(self) -> SignatureKeywordsConfig:
        """署名关键词配置"""
        return self._signature_keywords

    @property
    def numbering(self) -> NumberingConfig:
        """编号配置"""
        return self._numbering

    @property
    def code_block(self) -> CodeBlockConfig:
        """代码块配置"""
        return self._code_block

    @property
    def file(self) -> FileConfig:
        """文件配置"""
        return self._file

    @property
    def pandoc(self) -> PandocConfig:
        """Pandoc 配置"""
        return self._pandoc

    @property
    def rule_aliases(self) -> RuleAliasesConfig:
        """规则别名配置"""
        return self._rule_aliases

    def to_dict(self) -> Dict[str, Any]:
        """导出配置为字典"""
        return {
            "fonts": {
                "chinese": self._fonts.chinese,
                "english": self._fonts.english,
                "heading": self._fonts.heading,
                "abstract_thesis": self._fonts.abstract_thesis,
                "abstract_design": self._fonts.abstract_design,
                "signature": self._fonts.signature,
                "code": self._fonts.code,
            },
            "font_sizes": {
                "title": self._font_sizes.title,
                "heading_1": self._font_sizes.heading_1,
                "heading_2": self._font_sizes.heading_2,
                "heading_3": self._font_sizes.heading_3,
                "heading_4": self._font_sizes.heading_4,
                "body": self._font_sizes.body,
                "caption": self._font_sizes.caption,
                "header_footer": self._font_sizes.header_footer,
                "code": self._font_sizes.code,
            },
            "page": {
                "top_margin": self._page.top_margin,
                "bottom_margin": self._page.bottom_margin,
                "left_margin": self._page.left_margin,
                "right_margin": self._page.right_margin,
                "gutter": self._page.gutter,
                "header_distance": self._page.header_distance,
                "footer_distance": self._page.footer_distance,
            },
            "paragraph": {
                "line_spacing": self._paragraph.line_spacing,
                "first_line_indent_cm": self._paragraph.first_line_indent_cm,
            },
            "tab_stops": {
                "center": self._tab_stops.center,
                "right": self._tab_stops.right,
            },
            "table": {
                "top_border_pt": self._table.top_border_pt,
                "header_border_pt": self._table.header_border_pt,
                "bottom_border_pt": self._table.bottom_border_pt,
                "cell_space_pt": self._table.cell_space_pt,
            },
            "special_titles": list(self._special_titles.all),
            "signature_keywords": self._signature_keywords.items,
            "numbering": {
                "multilevel_heading_id": self._numbering.multilevel_heading_id,
                "reference_list_id": self._numbering.reference_list_id,
            },
            "code_block": {
                "background_color": self._code_block.background_color,
            },
            "file": {
                "max_file_size_mb": self._file.max_file_size_mb,
                "temp_file_prefix": self._file.temp_file_prefix,
            },
            "pandoc": {
                "version_check_timeout": self._pandoc.version_check_timeout,
                "convert_timeout": self._pandoc.convert_timeout,
                "download_url": self._pandoc.download_url,
            },
            "rule_aliases": self._rule_aliases.aliases,
        }


def get_config() -> Config:
    """获取全局配置实例"""
    return Config()


DEFAULT_CONFIG = Config()
