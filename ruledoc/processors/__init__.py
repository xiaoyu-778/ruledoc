"""后处理器模块

包含所有通用后处理器:
- StyleProcessor: 样式名映射
- HeadingProcessor: 标题处理
- CaptionProcessor: 题注处理
- ListProcessor: 列表处理

执行顺序 (关键):
规则前置处理器 → StyleProcessor → HeadingProcessor → ListProcessor → CaptionProcessor → 规则后置处理器
"""

from ruledoc.processors.base import PostProcessor
from ruledoc.processors.caption_processor import CaptionProcessor
from ruledoc.processors.heading_processor import HeadingProcessor
from ruledoc.processors.list_processor import ListProcessor
from ruledoc.processors.style_processor import StyleProcessor

__all__ = [
    "PostProcessor",
    "StyleProcessor",
    "HeadingProcessor",
    "ListProcessor",
    "CaptionProcessor",
]
