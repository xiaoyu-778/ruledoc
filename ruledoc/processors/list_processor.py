"""ListProcessor - 列表处理器

职责:
- 将自动编号列表转换为手打编号
- 统一列表项字体为 Times New Roman
- 移除自动编号格式，改为文本编号

执行顺序:
ListProcessor 在 StyleProcessor 之后执行，处理列表样式段落。
"""

from typing import TYPE_CHECKING

from ruledoc.processors.base import PostProcessor

if TYPE_CHECKING:
    from ruledoc.context import ProcessingContext


class ListProcessor(PostProcessor):
    """列表处理器

    处理文档中的编号列表:
    - 将 Word 自动编号转换为手打编号 (1., 2., 3.)
    - 统一字体为 Times New Roman
    - 保持列表缩进和间距

    支持的列表样式:
    - List Paragraph
    - List Bullet
    - List Number
    - Compact
    - 其他列表相关样式
    """

    # 列表样式关键词
    LIST_STYLE_KEYWORDS = ["list", "bullet", "number", "compact", "列表", "项目", "编号"]

    # 手打编号格式
    NUMBERING_FORMAT = "{num}. {text}"

    def __init__(self, font_name: str = "Times New Roman"):
        """初始化列表处理器

        Args:
            font_name: 列表项字体名称，默认为 Times New Roman
        """
        self.font_name = font_name
        self._processed_count = 0

    def process(self, ctx: "ProcessingContext") -> None:
        """执行列表处理

        遍历文档中的所有段落，识别列表项并转换为手打编号。

        Args:
            ctx: 处理上下文
        """
        self._processed_count = 0

        # 按顺序处理，维护编号计数
        current_number = 0
        in_list = False

        for para in ctx.doc.paragraphs:
            if self._is_list_paragraph(para):
                if not in_list:
                    # 新列表开始
                    current_number = 1
                    in_list = True
                else:
                    current_number += 1

                self._convert_to_manual_numbering(para, current_number)
                self._processed_count += 1
            else:
                in_list = False
                current_number = 0

        if self._processed_count > 0:
            ctx.add_warning(
                f"ListProcessor: 处理了 {self._processed_count} 个列表项，"
                f"转换为手打编号并应用 {self.font_name} 字体"
            )

    def _is_list_paragraph(self, para) -> bool:
        """检测段落是否为列表项

        通过样式名判断:
        - 包含 'list', 'bullet', 'number', 'compact' 等关键词
        - 排除标题样式

        Args:
            para: 段落对象

        Returns:
            是否为列表项
        """
        if not para.style:
            return False

        style_name = para.style.name.lower() if para.style.name else ""

        # 排除标题样式
        if style_name.startswith("heading"):
            return False

        # 检查是否为列表样式
        for keyword in self.LIST_STYLE_KEYWORDS:
            if keyword in style_name:
                return True

        # 检查是否有自动编号
        p_xml = para._p.xml
        if "<w:numPr>" in p_xml and "w:numId" in p_xml:
            return True

        return False

    def _convert_to_manual_numbering(self, para, number: int) -> None:
        """将自动编号列表项转换为手打编号

        Args:
            para: 段落对象
            number: 列表项编号
        """
        from docx.shared import Pt

        text = para.text.strip()
        if not text:
            return

        # 清除段落内容
        para.clear()

        # 添加手打编号文本
        numbered_text = f"{number}. {text}"
        run = para.add_run(numbered_text)

        # 应用字体设置
        run.font.name = self.font_name

        # 尝试设置字体大小（使用正文大小）
        try:
            run.font.size = Pt(12)
        except Exception:
            pass

        # 移除自动编号属性
        self._remove_numbering(para)

    def _remove_numbering(self, para) -> None:
        """移除段落的自动编号属性

        Args:
            para: 段落对象
        """
        try:
            pPr = para._p.get_or_add_pPr()

            # 查找并移除 numPr 元素
            for numPr in pPr.findall(
                ".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}numPr"
            ):
                pPr.remove(numPr)
        except Exception:
            pass
