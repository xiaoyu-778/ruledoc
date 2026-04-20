"""
StyleProcessor - 样式名映射处理器

职责:
- 将 Pandoc 生成的样式名映射为标准 Word 样式名
- 处理样式不存在的情况
- 支持自定义样式映射

执行顺序:
StyleProcessor 是第一个执行的处理器，在 HeadingProcessor 和 CaptionProcessor 之前。
"""

from typing import TYPE_CHECKING, Dict, Optional, Set

from ruledoc.processors.base import PostProcessor

if TYPE_CHECKING:
    from ruledoc.context import ProcessingContext


class StyleProcessor(PostProcessor):
    """
    样式名映射处理器

    将 Pandoc 生成的样式名 (如 "heading 1") 映射为 Word 标准样式名 (如 "Heading 1")。
    这是后处理器链中的第一个处理器。

    Attributes:
        style_map: 样式名映射字典 (键为小写 Pandoc 样式名，值为 Word 样式名)
        _missing_styles: 记录缺失的样式，避免重复警告
    """

    DEFAULT_STYLE_MAP = {
        "heading 1": "Heading 1",
        "heading 2": "Heading 2",
        "heading 3": "Heading 3",
        "heading 4": "Heading 4",
        "heading 5": "Heading 5",
        "heading 6": "Heading 6",
        "title": "Title",
        "subtitle": "Subtitle",
        "normal": "Normal",
        "first paragraph": "Normal",
        "firstparagraph": "Normal",
        "body text": "Normal",
        "bodytext": "Normal",
        "quote": "Quote",
        "code": "Code",
        "caption": "Caption",
        "figure": "Caption",
        "table": "Caption",
    }

    def __init__(self, style_map: Optional[Dict[str, str]] = None):
        """
        初始化样式处理器

        Args:
            style_map: 自定义样式名映射字典，会与默认映射合并
        """
        self._style_map: Dict[str, str] = {}
        self._missing_styles: Set[str] = set()

        if style_map:
            for key, value in style_map.items():
                self._style_map[key.lower()] = value

        for key, value in self.DEFAULT_STYLE_MAP.items():
            if key not in self._style_map:
                self._style_map[key] = value

    @property
    def style_map(self) -> Dict[str, str]:
        """获取样式映射字典"""
        return self._style_map.copy()

    def add_mapping(self, pandoc_style: str, word_style: str) -> None:
        """
        添加样式映射

        Args:
            pandoc_style: Pandoc 样式名
            word_style: Word 样式名
        """
        self._style_map[pandoc_style.lower()] = word_style

    def process(self, ctx: "ProcessingContext") -> None:
        """
        执行样式映射

        遍历文档中的所有段落，将 Pandoc 生成的样式名映射为 Word 标准样式名。
        如果目标样式不存在，添加警告到上下文。

        Args:
            ctx: 处理上下文
        """
        self._missing_styles.clear()

        for para in ctx.doc.paragraphs:
            self._process_paragraph(para, ctx)

        if self._missing_styles:
            styles_list = ", ".join(sorted(self._missing_styles))
            ctx.add_warning(f"StyleProcessor: 以下样式不存在: {styles_list}")

    def _process_paragraph(self, para, ctx: "ProcessingContext") -> None:
        """
        处理单个段落

        Args:
            para: 段落对象
            ctx: 处理上下文
        """
        if not para.style:
            return

        current_style = para.style.name
        if not current_style:
            return

        style_key = current_style.lower()

        if style_key not in self._style_map:
            return

        target_style = self._style_map[style_key]

        if current_style == target_style:
            return

        try:
            para.style = target_style
        except KeyError:
            if target_style not in self._missing_styles:
                self._missing_styles.add(target_style)
                ctx.add_warning(
                    f"StyleProcessor: 样式 '{target_style}' 不存在，段落 '{para.text[:30]}...' 保持原样式"
                )
        except Exception as e:
            ctx.add_warning(f"StyleProcessor: 应用样式 '{target_style}' 失败: {e}")

    def get_mapping_stats(self) -> Dict[str, int]:
        """
        获取映射统计信息

        Returns:
            包含映射数量的统计字典
        """
        return {
            "total_mappings": len(self._style_map),
            "missing_styles": len(self._missing_styles),
        }
