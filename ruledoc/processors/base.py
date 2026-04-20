"""PostProcessor 基类 - 后处理器接口定义

设计原则:
- 单一职责: 每个处理器只负责一种文档转换
- 接口优先: 先定义抽象方法，后实现
- 上下文传递: 通过 ProcessingContext 共享状态
"""

from abc import ABC, abstractmethod
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from ruledoc.context import ProcessingContext


class PostProcessor(ABC):
    """后处理器抽象基类

    所有文档后处理器必须继承此类并实现 process() 方法。

    处理器职责:
    - 执行单一的文档转换任务
    - 通过上下文读取和更新状态
    - 添加警告消息 (非致命错误)

    处理器不应:
    - 包含学校特定的格式逻辑 (由 FormatRule 提供)
    - 抛出异常 (应添加警告并继续处理)

    Example:
        >>> class MyProcessor(PostProcessor):
        ...     def process(self, ctx: ProcessingContext) -> None:
        ...         for para in ctx.doc.paragraphs:
        ...             # 处理段落
        ...             pass
    """

    @abstractmethod
    def process(self, ctx: "ProcessingContext") -> None:
        """执行处理

        Args:
            ctx: 处理上下文，包含文档和状态信息
        """
        pass

    @property
    def name(self) -> str:
        """处理器名称

        Returns:
            处理器类名
        """
        return self.__class__.__name__
