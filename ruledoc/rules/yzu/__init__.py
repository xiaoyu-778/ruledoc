"""
扬州大学 (YZU) 论文格式规则包

提供:
- YZUThesisRule: 毕业论文规则
- YZUDesignRule: 毕业设计报告规则
- YZUFormatRule: 向后兼容别名 (推荐使用 YZUThesisRule)
- ThesisType: 论文类型枚举
- YZUConstants: 常量定义
- YZUMixin: 共享方法 Mixin
"""

from typing import TYPE_CHECKING

from .common import ThesisType, YZUConstants, YZUMixin
from .yzu_design import YZUDesignRule
from .yzu_thesis import YZUThesisRule

if TYPE_CHECKING:
    YZUFormatRule = YZUThesisRule
else:
    YZUFormatRule = YZUThesisRule

__all__ = [
    "YZUThesisRule",
    "YZUDesignRule",
    "YZUFormatRule",
    "ThesisType",
    "YZUConstants",
    "YZUMixin",
]
