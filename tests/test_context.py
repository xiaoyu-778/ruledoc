"""ProcessingContext 单元测试"""

from unittest.mock import MagicMock

from ruledoc.context import ProcessingContext
from ruledoc.rules.base import FormatRule


class TestProcessingContext:
    """ProcessingContext 测试类"""

    def test_init_default_values(self):
        """测试默认值初始化"""
        mock_doc = MagicMock()
        mock_rule = MagicMock(spec=FormatRule)

        ctx = ProcessingContext(doc=mock_doc, rule=mock_rule)

        assert ctx.doc == mock_doc
        assert ctx.rule == mock_rule
        assert ctx.chapter_num == 0
        assert ctx.fig_counter == 0
        assert ctx.tab_counter == 0
        assert ctx.in_abstract is False
        assert ctx.in_references is False
        assert ctx.warnings == []

    def test_add_warning(self):
        """测试添加警告"""
        mock_doc = MagicMock()
        mock_rule = MagicMock(spec=FormatRule)

        ctx = ProcessingContext(doc=mock_doc, rule=mock_rule)
        ctx.add_warning("测试警告1")
        ctx.add_warning("测试警告2")

        assert len(ctx.warnings) == 2
        assert ctx.warnings[0] == "测试警告1"
        assert ctx.warnings[1] == "测试警告2"

    def test_reset_chapter_counters(self):
        """测试重置章节计数器"""
        mock_doc = MagicMock()
        mock_rule = MagicMock(spec=FormatRule)

        ctx = ProcessingContext(doc=mock_doc, rule=mock_rule)
        ctx.fig_counter = 5
        ctx.tab_counter = 3

        ctx.reset_chapter_counters()

        assert ctx.fig_counter == 0
        assert ctx.tab_counter == 0

    def test_get_next_fig_number(self):
        """测试获取下一个图编号"""
        mock_doc = MagicMock()
        mock_rule = MagicMock(spec=FormatRule)

        ctx = ProcessingContext(doc=mock_doc, rule=mock_rule)
        ctx.chapter_num = 2

        assert ctx.get_next_fig_number() == "2-1"
        assert ctx.get_next_fig_number() == "2-2"
        assert ctx.get_next_fig_number() == "2-3"
        assert ctx.fig_counter == 3

    def test_get_next_tab_number(self):
        """测试获取下一个表格编号"""
        mock_doc = MagicMock()
        mock_rule = MagicMock(spec=FormatRule)

        ctx = ProcessingContext(doc=mock_doc, rule=mock_rule)
        ctx.chapter_num = 3

        assert ctx.get_next_tab_number() == "3-1"
        assert ctx.get_next_tab_number() == "3-2"
        assert ctx.tab_counter == 2

    def test_chapter_num_increment(self):
        """测试章节号递增"""
        mock_doc = MagicMock()
        mock_rule = MagicMock(spec=FormatRule)

        ctx = ProcessingContext(doc=mock_doc, rule=mock_rule)

        ctx.chapter_num = 1
        ctx.reset_chapter_counters()

        ctx.chapter_num = 2
        ctx.reset_chapter_counters()

        assert ctx.chapter_num == 2
        assert ctx.fig_counter == 0
