"""后处理器单元测试"""

import pytest

from ruledoc.processors import CaptionProcessor, HeadingProcessor, StyleProcessor
from ruledoc.processors.base import PostProcessor


class TestStyleProcessor:
    """StyleProcessor 测试"""

    def test_init_default_style_map(self):
        """测试默认样式映射"""
        processor = StyleProcessor()

        assert "heading 1" in processor.style_map
        assert processor.style_map["heading 1"] == "Heading 1"

    def test_init_custom_style_map(self):
        """测试自定义样式映射"""
        custom_map = {"custom style": "Custom Style"}
        processor = StyleProcessor(style_map=custom_map)

        assert processor.style_map["custom style"] == "Custom Style"
        assert "heading 1" in processor.style_map

    def test_add_mapping(self):
        """测试添加样式映射"""
        processor = StyleProcessor()
        processor.add_mapping("new style", "New Style")

        assert processor.style_map["new style"] == "New Style"

    def test_name_property(self):
        """测试处理器名称"""
        processor = StyleProcessor()

        assert processor.name == "StyleProcessor"

    def test_get_mapping_stats(self):
        """测试获取映射统计"""
        processor = StyleProcessor()
        stats = processor.get_mapping_stats()

        assert "total_mappings" in stats
        assert "missing_styles" in stats
        assert stats["total_mappings"] > 0


class TestHeadingProcessor:
    """HeadingProcessor 测试"""

    def test_init_default_settings(self):
        """测试默认设置"""
        processor = HeadingProcessor()

        assert processor.remove_manual_numbering is True
        assert processor.apply_manual_numbering is True

    def test_name_property(self):
        """测试处理器名称"""
        processor = HeadingProcessor()

        assert processor.name == "HeadingProcessor"

    def test_chinese_numeral_map(self):
        """测试中文数字映射"""
        processor = HeadingProcessor()

        assert processor.CHINESE_NUMERAL_MAP["一"] == 1
        assert processor.CHINESE_NUMERAL_MAP["十"] == 10
        assert processor.CHINESE_NUMERAL_MAP["十五"] == 15

    def test_chinese_to_arabic(self):
        """测试中文数字转阿拉伯数字"""
        processor = HeadingProcessor()

        assert processor.chinese_to_arabic("一") == 1
        assert processor.chinese_to_arabic("十") == 10
        assert processor.chinese_to_arabic("十一") == 11
        assert processor.chinese_to_arabic("二十") == 20
        assert processor.chinese_to_arabic("十五") == 15

    def test_get_stats(self):
        """测试获取统计信息"""
        processor = HeadingProcessor()
        stats = processor.get_stats()

        assert "processed_count" in stats
        assert "heading_counts" in stats


class TestCaptionProcessor:
    """CaptionProcessor 测试"""

    def test_init_default_settings(self):
        """测试默认设置"""
        processor = CaptionProcessor()

        assert processor._use_style_detection is True
        assert processor._use_seq_field is False  # 默认False，MS Word兼容性更好

    def test_name_property(self):
        """测试处理器名称"""
        processor = CaptionProcessor()

        assert processor.name == "CaptionProcessor"

    def test_caption_styles(self):
        """测试题注样式集合"""
        processor = CaptionProcessor()

        assert "caption" in processor.CAPTION_STYLES
        assert "figure" in processor.CAPTION_STYLES
        assert "table" in processor.CAPTION_STYLES

    def test_get_stats(self):
        """测试获取统计信息"""
        processor = CaptionProcessor()
        stats = processor.get_stats()

        assert "processed_figures" in stats
        assert "processed_tables" in stats
        assert "skipped_captions" in stats
        assert "total_processed" in stats

    def test_add_fig_pattern(self):
        """测试添加图题注模式"""
        import re

        processor = CaptionProcessor()
        pattern = re.compile(r"^图表\s*\d+")

        processor.add_fig_pattern(pattern)

        assert pattern in processor._fig_patterns

    def test_add_tab_pattern(self):
        """测试添加表题注模式"""
        import re

        processor = CaptionProcessor()
        pattern = re.compile(r"^表格\s*\d+")

        processor.add_tab_pattern(pattern)

        assert pattern in processor._tab_patterns


class TestPostProcessorBase:
    """PostProcessor 基类测试"""

    def test_is_abstract(self):
        """测试是否为抽象类"""
        with pytest.raises(TypeError):
            PostProcessor()  # type: ignore[abstract]

    def test_name_property(self):
        """测试 name 属性"""

        class ConcreteProcessor(PostProcessor):
            def process(self, ctx):  # type: ignore[override]
                _ = ctx
                pass

        processor = ConcreteProcessor()

        assert processor.name == "ConcreteProcessor"
