"""HeadingProcessor - 标题处理器

职责:
- 检测标题层级
- 删除手打编号
- 应用手动编号（1.、1.1、1.1.1 格式）
- 更新 ctx.chapter_num

执行顺序:
HeadingProcessor 是第二个执行的处理器，在 StyleProcessor 之后、CaptionProcessor 之前。
这是关键：CaptionProcessor 需要 ctx.chapter_num 来生成正确的图表编号。

扬州大学毕业论文格式规范：
- 编号格式：1.、1.1、1.1.1（第四种编号方式）
- 特殊标题不编号：摘要、引言、参考文献、致谢、附录等
"""

import re
from typing import TYPE_CHECKING, Dict, List, Optional

from ruledoc.processors.base import PostProcessor

if TYPE_CHECKING:
    from ruledoc.context import ProcessingContext


class HeadingProcessor(PostProcessor):
    """标题处理器

    处理文档中的标题段落:
    - 删除原有手打编号 (如 "第一章 绪论" → "绪论", "1.1 背景" → "背景")
    - 应用手动编号（直接修改文本，不使用 Word 多级列表）
    - 编号格式：一级 "1."、二级 "1.1"、三级 "1.1.1"、四级 "1.1.1.1"
    - 更新上下文中的章节号 (用于 CaptionProcessor)

    注意: 此处理器必须在 CaptionProcessor 之前执行，
    以确保 CaptionProcessor 能获取正确的章节号。

    Attributes:
        remove_manual_numbering: 是否删除原有手打编号
        apply_manual_numbering: 是否应用手动编号
        chinese_numeral_map: 中文数字映射
    """

    CHINESE_NUMERAL_MAP = {
        "一": 1,
        "二": 2,
        "三": 3,
        "四": 4,
        "五": 5,
        "六": 6,
        "七": 7,
        "八": 8,
        "九": 9,
        "十": 10,
        "十一": 11,
        "十二": 12,
        "十三": 13,
        "十四": 14,
        "十五": 15,
    }

    PATTERNS = {
        "chapter_cn": re.compile(r"^(第[一二三四五六七八九十]+[章节部篇])\s*"),
        "section_arabic": re.compile(r"^(\d+(?:\.\d+)*)\s+"),
        "section_arabic_full": re.compile(r"^(\d+\.\d+\.\d+\.\d+)\s+"),  # 匹配四级编号如 2.1.1.1
        "section_cn": re.compile(r"^([一二三四五六七八九十]+[、.．])\s*"),
        "subsection_letter": re.compile(r"^([（(]\s*[a-zA-Z]\s*[)）]|[a-zA-Z]\s*[、.．])\s*"),
    }

    MULTILEVEL_NUM_ID = 100

    def __init__(
        self,
        remove_manual_numbering: bool = True,
        apply_manual_numbering: bool = True,
        custom_patterns: Optional[List[re.Pattern]] = None,
    ):
        """初始化标题处理器

        Args:
            remove_manual_numbering: 是否删除原有手打编号
            apply_manual_numbering: 是否应用手动编号（直接修改文本）
            custom_patterns: 自定义编号匹配模式列表
        """
        self.remove_manual_numbering = remove_manual_numbering
        self.apply_manual_numbering = apply_manual_numbering
        self._custom_patterns = custom_patterns or []

        self._heading_counts: Dict[int, int] = {}
        self._processed_count = 0

    def process(self, ctx: "ProcessingContext") -> None:
        """执行标题处理

        遍历文档中的所有段落，识别标题并处理编号。

        Args:
            ctx: 处理上下文
        """
        self._heading_counts = {1: 0, 2: 0, 3: 0, 4: 0, 5: 0, 6: 0}
        self._processed_count = 0

        for para in ctx.doc.paragraphs:
            if self._is_heading(para):
                self._process_heading(para, ctx)

        ctx.add_warning(
            f"HeadingProcessor: 处理了 {self._processed_count} 个标题，"
            f"当前章节号: {ctx.chapter_num}"
        )

    def _is_heading(self, para) -> bool:
        """检测段落是否为标题

        Args:
            para: 段落对象

        Returns:
            是否为标题
        """
        if not para.style:
            return False
        style_name = para.style.name.lower()
        return style_name.startswith("heading")

    SPECIAL_SKIP_TITLES = {
        "摘要",
        "abstract",
        "关键词",
        "keywords",
        "参考文献",
        "references",
        "目录",
        "contents",
        "致谢",
        "acknowledgements",
        "附录",
        "appendix",
        "引言",
        "introduction",
    }

    SPECIAL_SKIP_EXACT = set()

    def _should_skip_numbering(self, para, level: int, ctx: "ProcessingContext") -> bool:
        """判断是否应跳过多级列表编号（非章节标题不应编号）

        跳过的场景：
        - 第一个 Heading 1（论文标题）
        - 特殊标题（摘要、Abstract、参考文献、致谢、附录等）

        Args:
            para: 段落对象
            level: 标题层级
            ctx: 处理上下文

        Returns:
            是否跳过编号
        """
        text = para.text.strip().lower().replace(" ", "")
        text_original = para.text.strip()

        if text_original.lower() in self.SPECIAL_SKIP_EXACT or text in self.SPECIAL_SKIP_EXACT:
            return True

        if any(kw in text for kw in self.SPECIAL_SKIP_TITLES):
            if len(text_original) < 15:
                return True

        if level == 1:
            if not hasattr(self, "_first_h1_seen"):
                self._first_h1_seen = set()

            doc_id = id(ctx.doc)
            if doc_id not in self._first_h1_seen:
                self._first_h1_seen.add(doc_id)
                return True

        return False

    def _get_heading_level(self, para) -> int:
        """获取标题层级

        Args:
            para: 段落对象

        Returns:
            标题层级 (1-6)，非标题返回 0
        """
        if not para.style:
            return 0
        style_name = para.style.name.lower()

        for i in range(1, 7):
            if f"heading {i}" in style_name:
                return i
        return 0

    def _process_heading(self, para, ctx: "ProcessingContext") -> None:
        """处理单个标题段落

        Args:
            para: 段落对象
            ctx: 处理上下文
        """
        level = self._get_heading_level(para)
        if level == 0:
            return

        text = para.text.strip()
        original_text = text

        if self.remove_manual_numbering:
            text = self._remove_numbering(text, level)

        skip_num = self._should_skip_numbering(para, level, ctx)

        if self.apply_manual_numbering and level <= 4 and not skip_num:
            text = self._apply_manual_numbering(text, level)

        if text != original_text:
            para.text = text
            ctx.add_warning(f"HeadingProcessor: 处理标题 '{original_text[:30]}' → '{text[:30]}'")

        self._update_context(level, ctx, skip_num)
        self._processed_count += 1

    def _remove_numbering(self, text: str, level: int) -> str:
        """删除手打编号

        Args:
            text: 原始文本
            level: 标题层级

        Returns:
            删除编号后的文本
        """
        if level == 1:
            match = self.PATTERNS["chapter_cn"].match(text)
            if match:
                return text[len(match.group(0)) :].strip()

        if level >= 2:
            # 优先匹配四级编号（如 2.1.1.1）
            if level >= 4:
                match = self.PATTERNS["section_arabic_full"].match(text)
                if match:
                    return text[len(match.group(0)) :].strip()

            match = self.PATTERNS["section_arabic"].match(text)
            if match:
                return text[len(match.group(0)) :].strip()

            match = self.PATTERNS["section_cn"].match(text)
            if match:
                return text[len(match.group(0)) :].strip()

        if level >= 3:
            match = self.PATTERNS["subsection_letter"].match(text)
            if match:
                return text[len(match.group(0)) :].strip()

        for pattern in self._custom_patterns:
            match = pattern.match(text)
            if match:
                return text[len(match.group(0)) :].strip()

        return text

    def _update_context(self, level: int, ctx: "ProcessingContext", skip_num: bool = False) -> None:
        """更新处理上下文

        当遇到一级标题时，增加章节号并重置图表计数器。

        Args:
            level: 标题层级
            ctx: 处理上下文
            skip_num: 是否跳过编号（特殊标题）
        """
        if level == 1 and not skip_num:
            ctx.chapter_num += 1
            ctx.reset_chapter_counters()

        ctx.in_abstract = False
        ctx.in_references = False

    def _apply_manual_numbering(self, text: str, level: int) -> str:
        """应用手动编号（直接修改文本）

        编号格式：
        - 一级标题：1.、2.、3.
        - 二级标题：1.1、1.2、2.1
        - 三级标题：1.1.1、1.1.2
        - 四级标题：1.1.1.1

        Args:
            text: 标题文本（已删除原有编号）
            level: 标题层级

        Returns:
            带编号的标题文本
        """
        if level == 1:
            self._heading_counts[1] = self._heading_counts.get(1, 0) + 1
            self._heading_counts[2] = 0
            self._heading_counts[3] = 0
            self._heading_counts[4] = 0
            return f"{self._heading_counts[1]}. {text}"

        elif level == 2:
            self._heading_counts[2] = self._heading_counts.get(2, 0) + 1
            self._heading_counts[3] = 0
            self._heading_counts[4] = 0
            chap = self._heading_counts.get(1, 1)
            return f"{chap}.{self._heading_counts[2]} {text}"

        elif level == 3:
            self._heading_counts[3] = self._heading_counts.get(3, 0) + 1
            self._heading_counts[4] = 0
            chap = self._heading_counts.get(1, 1)
            sec = self._heading_counts.get(2, 1)
            return f"{chap}.{sec}.{self._heading_counts[3]} {text}"

        elif level == 4:
            self._heading_counts[4] = self._heading_counts.get(4, 0) + 1
            chap = self._heading_counts.get(1, 0)
            sec = self._heading_counts.get(2, 0)
            sub = self._heading_counts.get(3, 0)
            # 如果上级标题计数为0，使用1作为默认值
            chap = chap if chap > 0 else 1
            sec = sec if sec > 0 else 1
            sub = sub if sub > 0 else 1
            return f"{chap}.{sec}.{sub}.{self._heading_counts[4]} {text}"

        return text

    def chinese_to_arabic(self, chinese: str) -> int:
        """将中文数字转换为阿拉伯数字

        Args:
            chinese: 中文数字字符串

        Returns:
            阿拉伯数字
        """
        chinese = chinese.strip()
        if chinese in self.CHINESE_NUMERAL_MAP:
            return self.CHINESE_NUMERAL_MAP[chinese]

        if chinese.startswith("十"):
            if len(chinese) == 1:
                return 10
            return 10 + self.CHINESE_NUMERAL_MAP.get(chinese[1], 0)

        if chinese.endswith("十"):
            return self.CHINESE_NUMERAL_MAP.get(chinese[0], 1) * 10

        return 0

    def get_stats(self) -> Dict:
        """获取处理统计信息

        Returns:
            统计信息字典
        """
        return {
            "processed_count": self._processed_count,
            "heading_counts": self._heading_counts.copy(),
        }
