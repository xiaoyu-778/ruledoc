"""
ProcessingContext - 处理上下文状态管理

设计原则:
- 状态传递对象，不包含业务逻辑
- 使用 dataclass 简化定义
- 所有处理器共享同一上下文实例
"""

from dataclasses import dataclass, field
from typing import TYPE_CHECKING, Any, List

if TYPE_CHECKING:
    from ruledoc.rules.base import FormatRule


@dataclass
class ProcessingContext:
    """
    处理上下文 - 跨处理器状态共享

    用于在处理器链中传递和共享状态信息。
    每个处理器可以读取和修改上下文中的状态。

    Attributes:
        doc: python-docx Document 对象
        rule: 当前使用的格式规则
        chapter_num: 当前章节号 (由 HeadingProcessor 更新)
        fig_counter: 图计数器 (当前章节内)
        tab_counter: 表格计数器 (当前章节内)
        in_abstract: 是否在摘要部分
        in_references: 是否在参考文献部分
        warnings: 处理过程中的警告列表
        _title_para_id: 内部使用，记录已识别的标题段落ID

    Example:
        >>> ctx = ProcessingContext(doc=document, rule=rule)
        >>> ctx.chapter_num = 1
        >>> ctx.add_warning("发现未识别的段落样式")
    """

    doc: Any
    rule: "FormatRule"

    chapter_num: int = 0
    fig_counter: int = 0
    tab_counter: int = 0

    in_abstract: bool = False
    in_references: bool = False

    warnings: List[str] = field(default_factory=list)

    _title_para_id: int = 0

    def add_warning(self, msg: str) -> None:
        """
        添加警告消息

        Args:
            msg: 警告消息内容
        """
        self.warnings.append(msg)

    def reset_chapter_counters(self) -> None:
        """
        重置章节计数器

        在新章节开始时调用，重置图表计数器。
        """
        self.fig_counter = 0
        self.tab_counter = 0

    def get_next_fig_number(self) -> str:
        """
        获取下一个图编号

        Returns:
            格式为 "章节号-图号" 的字符串，如 "1-1"
        """
        self.fig_counter += 1
        return f"{self.chapter_num}-{self.fig_counter}"

    def get_next_tab_number(self) -> str:
        """
        获取下一个表格编号

        Returns:
            格式为 "章节号-表号" 的字符串，如 "1-1"
        """
        self.tab_counter += 1
        return f"{self.chapter_num}-{self.tab_counter}"
