"""异常类单元测试
"""

import pytest

from ruledoc.exceptions import (
    ConfigurationError,
    PandocNotInstalledError,
    ProcessingError,
    RuleDocError,
    RuleNotFoundError,
)


class TestExceptions:
    """异常类测试"""

    def test_ruledoc_error_basic(self):
        """测试基础异常"""
        error = RuleDocError("测试错误")

        assert str(error) == "测试错误"
        assert error.message == "测试错误"
        assert error.context == {}

    def test_ruledoc_error_with_context(self):
        """测试带上下文的异常"""
        error = RuleDocError("测试错误", {"key": "value", "count": 42})

        assert "测试错误" in str(error)
        assert "key=value" in str(error)
        assert "count=42" in str(error)
        assert error.context == {"key": "value", "count": 42}

    def test_rule_not_found_error(self):
        """测试规则不存在异常"""
        error = RuleNotFoundError("规则 'xyz' 不存在", {"available_rules": ["yzu", "yzu_design"]})

        assert isinstance(error, RuleDocError)
        assert "xyz" in error.message
        assert error.context["available_rules"] == ["yzu", "yzu_design"]

    def test_processing_error(self):
        """测试处理错误异常"""
        error = ProcessingError("文档处理失败", {"input_path": "test.docx", "error": "格式错误"})

        assert isinstance(error, RuleDocError)
        assert error.message == "文档处理失败"

    def test_configuration_error(self):
        """测试配置错误异常"""
        error = ConfigurationError("无效的配置", {"key": "top_margin", "value": -1})

        assert isinstance(error, RuleDocError)
        assert "无效的配置" in str(error)

    def test_pandoc_not_installed_error(self):
        """测试 Pandoc 未安装异常"""
        error = PandocNotInstalledError(
            "Pandoc 未安装", {"download_url": "https://pandoc.org/installing.html"}
        )

        assert isinstance(error, RuleDocError)
        assert "Pandoc" in error.message

    def test_exception_inheritance(self):
        """测试异常继承关系"""
        assert issubclass(RuleNotFoundError, RuleDocError)
        assert issubclass(ProcessingError, RuleDocError)
        assert issubclass(ConfigurationError, RuleDocError)
        assert issubclass(PandocNotInstalledError, RuleDocError)

    def test_catch_with_base_class(self):
        """测试使用基类捕获异常"""
        with pytest.raises(RuleDocError):
            raise RuleNotFoundError("测试")

        with pytest.raises(RuleDocError):
            raise ProcessingError("测试")

        with pytest.raises(RuleDocError):
            raise ConfigurationError("测试")

        with pytest.raises(RuleDocError):
            raise PandocNotInstalledError("测试")
