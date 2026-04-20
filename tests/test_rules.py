"""
规则注册和加载测试
"""

from typing import Dict, Tuple

from ruledoc.rules.base import (
    FormatRule,
    list_available_rules,
    load_rule,
)
from ruledoc.rules.yzu import YZUDesignRule, YZUFormatRule, YZUThesisRule


class DummyRule(FormatRule):
    """测试用虚拟规则"""

    @property
    def name(self) -> str:
        return "dummy"

    @property
    def description(self) -> str:
        return "虚拟测试规则"

    def get_page_settings(self) -> Dict[str, float]:
        return {"top_margin": 2.0, "bottom_margin": 2.0}

    def get_font_settings(self) -> Dict[str, str]:
        return {"body_font": "宋体"}

    def get_heading_format(self, level: int) -> Tuple[str, str, int]:
        _ = level
        return ("黑体", "left", 14)


class TestRuleRegistration:
    """规则注册测试"""

    def test_list_available_rules(self):
        """测试列出可用规则"""
        rules = list_available_rules()

        assert isinstance(rules, list)
        assert "yzu_thesis" in rules
        assert "yzu_design" in rules

    def test_load_yzu_thesis_rule(self):
        """测试加载 YZU Thesis 规则"""
        rule = load_rule("yzu_thesis")

        assert rule is not None
        assert rule.name == "yzu_thesis"
        assert "扬州大学" in rule.description

    def test_load_yzu_alias(self):
        """测试向后兼容别名 'yzu' -> 'yzu_thesis'"""
        rule = load_rule("yzu")

        assert rule is not None
        assert rule.name == "yzu_thesis"
        assert "扬州大学" in rule.description

    def test_load_yzu_design_rule(self):
        """测试加载 YZU Design 规则"""
        rule = load_rule("yzu_design")

        assert rule is not None
        assert rule.name == "yzu_design"
        assert "毕业设计报告" in rule.description

    def test_load_nonexistent_rule(self):
        """测试加载不存在的规则"""
        rule = load_rule("nonexistent_rule_xyz")

        assert rule is None

    def test_rule_instance_is_unique(self):
        """测试每次加载返回新实例"""
        rule1 = load_rule("yzu_thesis")
        rule2 = load_rule("yzu_thesis")

        assert rule1 is not None
        assert rule2 is not None
        assert rule1 is not rule2
        assert rule1.name == rule2.name


class TestYZUThesisRule:
    """YZU Thesis 规则测试"""

    def test_page_settings(self):
        """测试页面设置"""
        rule = YZUThesisRule()
        settings = rule.get_validated_page_settings()

        assert settings["top_margin"] == 2.2
        assert settings["bottom_margin"] == 2.2
        assert settings["left_margin"] == 2.5
        assert settings["right_margin"] == 2.0
        assert settings["gutter"] == 0.5

    def test_font_settings(self):
        """测试字体设置"""
        rule = YZUThesisRule()
        settings = rule.get_font_settings()

        assert settings["chinese_font"] == "宋体"
        assert settings["english_font"] == "Times New Roman"
        assert settings["heading_font"] == "黑体"

    def test_heading_format_level_1(self):
        """测试一级标题格式"""
        rule = YZUThesisRule()
        font, align, size = rule.get_heading_format(1)

        assert font == "黑体"
        assert align == "center"
        assert size == 15

    def test_heading_format_level_2(self):
        """测试二级标题格式"""
        rule = YZUThesisRule()
        font, align, size = rule.get_heading_format(2)

        assert font == "黑体"
        assert align == "left"
        assert size == 14

    def test_processor_config(self):
        """测试处理器配置"""
        rule = YZUThesisRule()
        config = rule.get_processor_config()

        assert config["style"] is True
        assert config["heading"] is True
        assert config["caption"] is True
        assert config["custom_order"] is None

    def test_style_map(self):
        """测试样式映射"""
        rule = YZUThesisRule()
        style_map = rule.get_style_map()

        assert "heading 1" in style_map
        assert style_map["heading 1"] == "Heading 1"

    def test_backward_compatible_alias(self):
        """测试向后兼容别名 YZUFormatRule"""
        rule = YZUFormatRule()

        assert isinstance(rule, YZUThesisRule)
        assert rule.name == "yzu_thesis"


class TestYZUDesignRule:
    """YZU Design 规则测试"""

    def test_design_rule_heading_format(self):
        """测试毕业设计报告一级标题格式（应靠左）"""
        rule = YZUDesignRule()
        font, align, size = rule.get_heading_format(1)

        assert font == "黑体"
        assert align == "left"
        assert size == 15

    def test_design_rule_name(self):
        """测试毕业设计报告规则名称"""
        rule = YZUDesignRule()

        assert rule.name == "yzu_design"
        assert "毕业设计报告" in rule.description
