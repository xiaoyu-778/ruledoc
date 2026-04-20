"""扬州大学毕业设计报告格式规则

继承自 YZUThesisRule，复用大部分格式逻辑，
仅覆盖有差异的部分（一级标题靠左、摘要内容宋体）。
"""

from ruledoc.rules.base import register_rule

from .common import ThesisType
from .yzu_thesis import YZUThesisRule


@register_rule
class YZUDesignRule(YZUThesisRule):
    """扬州大学毕业设计报告格式规则

    继承自 YZUThesisRule，复用大部分格式逻辑，
    仅覆盖有差异的部分（一级标题靠左、摘要内容宋体）。

    命名说明:
        - YZU: Yangzhou University (扬州大学)
        - Design: 毕业设计报告
        - Rule: 规则类后缀
    """

    def __init__(self):
        YZUThesisRule.__init__(self, ThesisType.DESIGN_REPORT)  # type: ignore[arg-type]

    @property
    def name(self) -> str:
        return "yzu_design"

    @property
    def description(self) -> str:
        return "扬州大学毕业设计报告格式"
