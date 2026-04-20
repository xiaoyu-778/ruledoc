"""
自定义异常类

设计原则:
- 所有异常继承自 RuleDocError
- 包含上下文信息便于调试
- 错误消息清晰明确
"""

from typing import Dict, Optional


class RuleDocError(Exception):
    """
    RuleDoc 基础异常

    所有 RuleDoc 异常的基类，提供统一的错误处理接口。

    Attributes:
        message: 错误消息
        context: 错误上下文信息 (用于调试)
    """

    def __init__(self, message: str, context: Optional[Dict] = None):
        self.message = message
        self.context = context or {}
        super().__init__(message)

    def __str__(self) -> str:
        if self.context:
            ctx_str = ", ".join(f"{k}={v}" for k, v in self.context.items())
            return f"{self.message} [{ctx_str}]"
        return self.message


class RuleNotFoundError(RuleDocError):
    """
    规则不存在异常

    当请求的规则未注册或无法加载时抛出。

    Example:
        >>> raise RuleNotFoundError("规则 'xyz' 不存在", {'available_rules': ['yzu']})
    """

    pass


class ProcessingError(RuleDocError):
    """
    处理错误异常

    当文档处理过程中发生错误时抛出。

    Example:
        >>> raise ProcessingError("段落处理失败", {'paragraph_index': 42})
    """

    pass


class ConfigurationError(RuleDocError):
    """
    配置错误异常

    当规则配置无效或缺失时抛出。

    Example:
        >>> raise ConfigurationError("无效的页面边距", {'key': 'top_margin', 'value': -1})
    """

    pass


class PandocNotInstalledError(RuleDocError):
    """
    Pandoc 未安装异常

    当系统未安装 Pandoc 时抛出。

    Example:
        >>> raise PandocNotInstalledError("Pandoc 未安装", {'download_url': 'https://pandoc.org/installing.html'})
    """

    pass
