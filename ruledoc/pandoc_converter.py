"""
Pandoc 转换器模块

封装 Pandoc 调用，实现 Markdown → Word 转换。

职责:
- 检测 Pandoc 可用性
- 执行 Markdown → Word 转换
- 管理临时文件（含自动清理机制）
"""

import atexit
import glob
import logging
import os
import shutil
import subprocess
import tempfile
import time
import weakref
from typing import Optional, Set

from ruledoc.config import get_config
from ruledoc.exceptions import PandocNotInstalledError, ProcessingError

_TEMP_FILE_PREFIX = "ruledoc_"
_temp_registry: Set[str] = set()
_cleanup_registered = False
_logger = logging.getLogger(__name__)


def _register_cleanup():
    """注册全局清理函数"""
    global _cleanup_registered
    if not _cleanup_registered:
        atexit.register(_cleanup_all_temp_files)
        _cleanup_registered = True


def _cleanup_all_temp_files():
    """清理所有注册的临时文件"""
    global _temp_registry
    for temp_path in list(_temp_registry):
        _safe_remove_file(temp_path)
    _temp_registry.clear()


def _safe_remove_file(file_path: str, max_retries: int = 3, delay: float = 0.1) -> bool:
    """
    安全删除文件，带重试机制

    Args:
        file_path: 文件路径
        max_retries: 最大重试次数
        delay: 重试间隔（秒）

    Returns:
        是否成功删除
    """
    if not file_path or not os.path.exists(file_path):
        return True

    for attempt in range(max_retries):
        try:
            os.remove(file_path)
            _logger.debug(f"成功删除临时文件: {file_path}")
            return True
        except OSError as e:
            if attempt < max_retries - 1:
                time.sleep(delay)
                delay *= 2
            else:
                _logger.warning(f"删除临时文件失败: {file_path}, 错误: {e}")
                return False
    return False


def cleanup_legacy_temp_files(temp_dir: Optional[str] = None) -> int:
    """
    清理历史残留的 RuleDoc 临时文件

    Args:
        temp_dir: 临时目录路径，默认使用系统临时目录

    Returns:
        清理的文件数量
    """
    if temp_dir is None:
        temp_dir = tempfile.gettempdir()

    pattern = os.path.join(temp_dir, f"{_TEMP_FILE_PREFIX}*.docx")
    cleaned = 0

    try:
        for file_path in glob.glob(pattern):
            if _safe_remove_file(file_path):
                cleaned += 1
    except Exception as e:
        _logger.warning(f"清理历史临时文件时出错: {e}")

    if cleaned > 0:
        _logger.info(f"清理了 {cleaned} 个历史残留临时文件")

    return cleaned


_register_cleanup()


class PandocConverter:
    """
    Pandoc 封装器

    封装 Pandoc 命令行工具，提供 Markdown 到 Word 的转换功能。

    Attributes:
        _available: Pandoc 可用性缓存
        _version: Pandoc 版本缓存
    """

    _available: Optional[bool] = None
    _version: Optional[str] = None

    def is_available(self) -> bool:
        """
        检测 Pandoc 是否可用

        使用 shutil.which 检测系统 PATH 中是否存在 pandoc 命令。
        结果会被缓存以提高后续调用效率。

        Returns:
            bool: Pandoc 是否可用
        """
        if PandocConverter._available is None:
            PandocConverter._available = shutil.which("pandoc") is not None
        return PandocConverter._available

    def get_version(self) -> str:
        """
        获取 Pandoc 版本

        执行 `pandoc --version` 命令获取版本信息。
        结果会被缓存以提高后续调用效率。

        Returns:
            str: Pandoc 版本号，如 "3.1.11"；获取失败返回 "unknown"
        """
        if PandocConverter._version is not None:
            return PandocConverter._version

        if not self.is_available():
            return "unknown"

        try:
            result = subprocess.run(
                ["pandoc", "--version"],
                capture_output=True,
                text=True,
                check=True,
                timeout=get_config().pandoc.version_check_timeout,
            )
            first_line = result.stdout.split("\n")[0]
            version = first_line.replace("pandoc ", "").strip()
            PandocConverter._version = version
            return version
        except (subprocess.CalledProcessError, subprocess.TimeoutExpired, OSError):
            return "unknown"

    def check_minimum_version(self, min_version: str) -> bool:
        """
        检查 Pandoc 版本是否满足最低要求

        Args:
            min_version: 最低版本号，如 "2.14"

        Returns:
            bool: 当前版本是否 >= 最低版本
        """
        current = self.get_version()
        if current == "unknown":
            return False

        try:
            current_parts = [int(x) for x in current.split(".")[:3]]
            min_parts = [int(x) for x in min_version.split(".")[:3]]

            for c, m in zip(current_parts, min_parts):
                if c > m:
                    return True
                if c < m:
                    return False
            return True
        except (ValueError, IndexError):
            return False

    def convert(
        self, input_path: str, output_path: Optional[str] = None, use_temp: bool = False
    ) -> str:
        """
        转换 Markdown 到 Word

        执行 `pandoc input.md -o output.docx` 命令进行转换。
        不使用 reference doc，生成标准 Word 文档。

        Args:
            input_path: 输入文件路径 (支持 .md, .txt 等 Pandoc 支持的格式)
            output_path: 输出文件路径 (可选，默认基于输入文件名生成)
            use_temp: 是否使用临时文件 (用于中间处理，调用者需负责清理)

        Returns:
            str: 输出文件路径

        Raises:
            PandocNotInstalledError: Pandoc 未安装
            ProcessingError: 转换失败
            FileNotFoundError: 输入文件不存在
        """
        if not self.is_available():
            raise PandocNotInstalledError(
                "Pandoc 未安装，请先安装 Pandoc", {"download_url": get_config().pandoc.download_url}
            )

        if not os.path.exists(input_path):
            raise ProcessingError(f"输入文件不存在: {input_path}", {"input_path": input_path})

        if use_temp:
            fd, output_path = tempfile.mkstemp(suffix=".docx", prefix=_TEMP_FILE_PREFIX)
            os.close(fd)
            _temp_registry.add(output_path)
        elif output_path is None:
            base = os.path.splitext(input_path)[0]
            output_path = f"{base}.docx"

        try:
            cmd = [
                "pandoc",
                "--from=markdown",
                "--to=docx",
                input_path,
                "-o",
                output_path,
            ]
            timeout = get_config().pandoc.convert_timeout
            result = subprocess.run(
                cmd, capture_output=True, text=True, check=True, timeout=timeout
            )
            return output_path
        except subprocess.TimeoutExpired:
            if use_temp:
                _safe_remove_file(output_path)
                _temp_registry.discard(output_path)
            raise ProcessingError(
                f"Pandoc 转换超时 ({timeout}秒)", {"input": input_path, "output": output_path}
            )
        except subprocess.CalledProcessError as e:
            if use_temp:
                _safe_remove_file(output_path)
                _temp_registry.discard(output_path)
            raise ProcessingError(
                f"Pandoc 转换失败: {e.stderr or '未知错误'}",
                {
                    "input": input_path,
                    "output": output_path,
                    "stderr": e.stderr,
                    "returncode": e.returncode,
                },
            )
        except OSError as e:
            if use_temp:
                _safe_remove_file(output_path)
                _temp_registry.discard(output_path)
            raise ProcessingError(
                f"Pandoc 执行错误: {e}",
                {"input": input_path, "output": output_path, "error": str(e)},
            )

    def convert_with_temp(self, input_path: str) -> str:
        """
        使用临时文件转换 (用于中间处理)

        生成的临时文件会自动注册到全局清理机制。

        Args:
            input_path: 输入文件路径

        Returns:
            str: 临时输出文件路径
        """
        return self.convert(input_path, use_temp=True)

    def cleanup_temp(self, temp_path: str) -> None:
        """
        清理临时文件

        Args:
            temp_path: 临时文件路径
        """
        if temp_path:
            _safe_remove_file(temp_path)
            _temp_registry.discard(temp_path)
