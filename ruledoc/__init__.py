"""
RuleDoc - 规则驱动的论文格式化工具

核心设计原则:
- 核心与规则完全解耦
- 接口优先，后实现
- 依赖注入
"""

__version__ = "1.0.0"
__author__ = "xiaoyu-778"
__email__ = "3389357760@qq.com"

from ruledoc.context import ProcessingContext
from ruledoc.exceptions import (
    ConfigurationError,
    PandocNotInstalledError,
    ProcessingError,
    RuleDocError,
    RuleNotFoundError,
)
from ruledoc.processors.base import PostProcessor
from ruledoc.rules.base import FormatRule

__all__ = [
    "RuleDocError",
    "RuleNotFoundError",
    "ProcessingError",
    "ConfigurationError",
    "PandocNotInstalledError",
    "ProcessingContext",
    "FormatRule",
    "PostProcessor",
]
