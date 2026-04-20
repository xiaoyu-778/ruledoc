"""命令行接口模块

提供 RuleDoc 的命令行入口，支持:
- 参数解析
- 交互式规则选择
- 多种调用方式
- 临时文件清理
"""

import argparse
import sys
from typing import TYPE_CHECKING, Optional

from ruledoc import __version__
from ruledoc.exceptions import ConfigurationError, ProcessingError, RuleNotFoundError
from ruledoc.formatter import Formatter
from ruledoc.pandoc_converter import cleanup_legacy_temp_files
from ruledoc.rules import list_available_rules, load_rule

if TYPE_CHECKING:
    from ruledoc.rules.base import FormatRule


def create_parser() -> argparse.ArgumentParser:
    """创建命令行参数解析器

    Returns:
        配置好的 ArgumentParser
    """
    parser = argparse.ArgumentParser(
        prog="ruledoc",
        description="规则驱动的论文格式化工具",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
示例:
  ruledoc input.md                    # 使用默认规则格式化
  ruledoc input.md --rule yzu         # 使用 YZU 规则
  ruledoc input.md -o output.docx     # 指定输出路径
  ruledoc --list-rules                # 列出所有规则
  ruledoc --cleanup                   # 清理残留临时文件
  python -m ruledoc input.md          # 模块方式调用
        """,
    )

    parser.add_argument("input", nargs="?", help="输入文件路径 (支持 .md, .docx)")

    parser.add_argument("-o", "--output", help="输出文件路径 (默认: {input}_formatted.docx)")

    parser.add_argument("--rule", help="规则名称 (默认: yzu)")

    parser.add_argument("--list-rules", action="store_true", help="列出所有可用规则")

    parser.add_argument("--cleanup", action="store_true", help="清理历史残留的临时文件")

    parser.add_argument("--version", action="version", version=f"%(prog)s {__version__}")

    return parser


def select_rule(rule_name: Optional[str] = None, non_interactive: bool = False) -> "FormatRule":
    """选择格式规则

    如果指定了规则名称，则直接加载。
    如果未指定，则在交互模式下提示用户选择。

    Args:
        rule_name: 规则名称 (可选)
        non_interactive: 非交互模式，直接使用默认规则

    Returns:
        FormatRule 实例

    Raises:
        RuleNotFoundError: 规则不存在
    """
    available_rules = list_available_rules()

    if rule_name:
        rule = load_rule(rule_name)
        if rule:
            return rule
        raise RuleNotFoundError(f"规则 '{rule_name}' 不存在", {"available_rules": available_rules})

    DEFAULT_RULE = "yzu_thesis"

    if non_interactive or not sys.stdin.isatty():
        rule = load_rule(DEFAULT_RULE)
        assert rule is not None, f"默认规则 '{DEFAULT_RULE}' 不存在"
        return rule

    print(f"默认将使用规则: {DEFAULT_RULE}")
    print("输入 'L' 查看其他规则，或按 Enter 继续...")

    while True:
        try:
            choice = input(f"请输入规则名称（直接回车使用 {DEFAULT_RULE}）: ").strip()
        except (EOFError, KeyboardInterrupt):
            print("\n操作已取消")
            sys.exit(1)

        if not choice:
            rule = load_rule(DEFAULT_RULE)
            assert rule is not None, f"默认规则 '{DEFAULT_RULE}' 不存在"
            return rule

        if choice.lower() == "l":
            print(f"可用规则: {', '.join(available_rules)}")
            continue

        rule = load_rule(choice)
        if rule:
            return rule

        print(f"规则 '{choice}' 不存在，可用规则: {', '.join(available_rules)}")


def format_file(input_path: str, rule: "FormatRule", output_path: Optional[str] = None) -> str:
    """格式化文件

    Args:
        input_path: 输入文件路径
        rule: 格式规则
        output_path: 输出文件路径 (可选)

    Returns:
        输出文件路径

    Raises:
        ProcessingError: 处理失败
    """
    formatter = Formatter(input_path, rule, output_path)
    formatter.process()
    return formatter.output_path


def main(argv: Optional[list] = None) -> int:
    """CLI 主入口

    Args:
        argv: 命令行参数 (可选，用于测试)

    Returns:
        退出码 (0=成功, 非0=失败)
    """
    parser = create_parser()
    args = parser.parse_args(argv)

    if args.list_rules:
        available_rules = list_available_rules()
        if available_rules:
            print("可用规则:")
            for name in available_rules:
                rule = load_rule(name)
                if rule:
                    print(f"  {name}: {rule.description}")
        else:
            print("没有可用的规则")
        return 0

    if args.cleanup:
        cleaned = cleanup_legacy_temp_files()
        if cleaned > 0:
            print(f"已清理 {cleaned} 个残留临时文件")
        else:
            print("没有发现残留临时文件")
        return 0

    if not args.input:
        parser.print_help()
        return 1

    try:
        rule = select_rule(args.rule, non_interactive=True)

        print(f"使用规则: {rule.name} ({rule.description})")
        print(f"输入文件: {args.input}")

        output_path = format_file(args.input, rule, args.output)

        print(f"输出文件: {output_path}")
        print("格式化完成!")

        return 0

    except RuleNotFoundError as e:
        print(f"错误: {e.message}", file=sys.stderr)
        if e.context.get("available_rules"):
            print(f"可用规则: {', '.join(e.context['available_rules'])}", file=sys.stderr)
        return 2

    except ConfigurationError as e:
        print(f"配置错误: {e.message}", file=sys.stderr)
        return 3

    except ProcessingError as e:
        print(f"处理错误: {e.message}", file=sys.stderr)
        return 4

    except KeyboardInterrupt:
        print("\n操作已取消", file=sys.stderr)
        return 130

    except Exception as e:
        print(f"未知错误: {e}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    sys.exit(main())
