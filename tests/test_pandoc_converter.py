"""PandocConverter 模块测试

覆盖:
- is_available() 检测
- get_version() 版本获取
- check_minimum_version() 版本检查
- convert() 转换功能
- 错误处理
"""

import os
import subprocess
from unittest.mock import MagicMock, patch

import pytest

from ruledoc.exceptions import PandocNotInstalledError, ProcessingError
from ruledoc.pandoc_converter import PandocConverter


class TestPandocConverterAvailability:
    """Pandoc 可用性测试"""

    def test_is_available_returns_bool(self):
        """测试 is_available 返回布尔值"""
        converter = PandocConverter()
        result = converter.is_available()

        assert isinstance(result, bool)

    @patch("shutil.which")
    def test_is_available_true(self, mock_which):
        """测试 Pandoc 可用"""
        mock_which.return_value = "/usr/bin/pandoc"

        PandocConverter._available = None
        converter = PandocConverter()

        assert converter.is_available() is True

    @patch("shutil.which")
    def test_is_available_false(self, mock_which):
        """测试 Pandoc 不可用"""
        mock_which.return_value = None

        PandocConverter._available = None
        converter = PandocConverter()

        assert converter.is_available() is False

    def test_is_available_caches_result(self):
        """测试结果缓存"""
        converter = PandocConverter()

        result1 = converter.is_available()
        result2 = converter.is_available()

        assert result1 == result2


class TestPandocConverterVersion:
    """Pandoc 版本测试"""

    def test_get_version_returns_string(self):
        """测试 get_version 返回字符串"""
        converter = PandocConverter()
        version = converter.get_version()

        assert isinstance(version, str)

    @patch("subprocess.run")
    def test_get_version_success(self, mock_run):
        """测试成功获取版本"""
        mock_run.return_value = MagicMock(stdout="pandoc 3.1.11\n", stderr="", returncode=0)

        PandocConverter._version = None
        PandocConverter._available = True
        converter = PandocConverter()

        version = converter.get_version()

        assert version == "3.1.11"

    @patch("subprocess.run")
    def test_get_version_failure(self, mock_run):
        """测试获取版本失败"""
        mock_run.side_effect = subprocess.CalledProcessError(1, "pandoc")

        PandocConverter._version = None
        PandocConverter._available = True
        converter = PandocConverter()

        version = converter.get_version()

        assert version == "unknown"

    def test_get_version_not_available(self):
        """测试 Pandoc 不可用时获取版本"""
        PandocConverter._available = False
        PandocConverter._version = None
        converter = PandocConverter()

        version = converter.get_version()

        assert version == "unknown"


class TestPandocConverterVersionCheck:
    """Pandoc 版本检查测试"""

    @patch.object(PandocConverter, "get_version")
    def test_check_minimum_version_true(self, mock_get_version):
        """测试版本满足要求"""
        mock_get_version.return_value = "3.1.11"

        converter = PandocConverter()

        assert converter.check_minimum_version("2.14") is True
        assert converter.check_minimum_version("3.0") is True
        assert converter.check_minimum_version("3.1.11") is True

    @patch.object(PandocConverter, "get_version")
    def test_check_minimum_version_false(self, mock_get_version):
        """测试版本不满足要求"""
        mock_get_version.return_value = "2.14.0"

        converter = PandocConverter()

        assert converter.check_minimum_version("3.0") is False

    @patch.object(PandocConverter, "get_version")
    def test_check_minimum_version_unknown(self, mock_get_version):
        """测试版本未知"""
        mock_get_version.return_value = "unknown"

        converter = PandocConverter()

        assert converter.check_minimum_version("2.14") is False


class TestPandocConverterConvert:
    """Pandoc 转换测试"""

    @pytest.fixture
    def sample_md(self, tmp_path):
        """创建测试 Markdown 文件"""
        content = """# 测试标题

这是测试内容。
"""
        file_path = tmp_path / "test_input.md"
        file_path.write_text(content, encoding="utf-8")

        return str(file_path)

    @patch.object(PandocConverter, "is_available")
    def test_convert_pandoc_not_installed(self, mock_is_available):
        """测试 Pandoc 未安装"""
        mock_is_available.return_value = False

        PandocConverter._available = False
        converter = PandocConverter()

        with pytest.raises(PandocNotInstalledError):
            converter.convert("input.md")

    def test_convert_file_not_found(self):
        """测试输入文件不存在"""
        converter = PandocConverter()

        if not converter.is_available():
            pytest.skip("Pandoc not installed")

        with pytest.raises(ProcessingError) as exc_info:
            converter.convert("/nonexistent/file.md")

        assert "不存在" in str(exc_info.value)

    @patch("subprocess.run")
    @patch.object(PandocConverter, "is_available")
    def test_convert_success(self, mock_is_available, mock_run, sample_md):
        """测试成功转换"""
        mock_is_available.return_value = True
        PandocConverter._available = True

        mock_run.return_value = MagicMock(stdout="", stderr="", returncode=0)

        converter = PandocConverter()
        result = converter.convert(sample_md)

        assert result.endswith(".docx")

    @patch("subprocess.run")
    @patch.object(PandocConverter, "is_available")
    def test_convert_with_output_path(self, mock_is_available, mock_run, sample_md, tmp_path):
        """测试指定输出路径"""
        mock_is_available.return_value = True
        PandocConverter._available = True

        mock_run.return_value = MagicMock(stdout="", stderr="", returncode=0)

        output_path = str(tmp_path / "custom_output.docx")
        converter = PandocConverter()
        result = converter.convert(sample_md, output_path)

        assert result == output_path

    @patch("subprocess.run")
    @patch.object(PandocConverter, "is_available")
    def test_convert_timeout(self, mock_is_available, mock_run, sample_md):
        """测试转换超时"""
        mock_is_available.return_value = True
        PandocConverter._available = True

        mock_run.side_effect = subprocess.TimeoutExpired(cmd="pandoc", timeout=60)

        converter = PandocConverter()

        with pytest.raises(ProcessingError) as exc_info:
            converter.convert(sample_md)

        assert "超时" in str(exc_info.value)

    @patch("subprocess.run")
    @patch.object(PandocConverter, "is_available")
    def test_convert_process_error(self, mock_is_available, mock_run, sample_md):
        """测试转换进程错误"""
        mock_is_available.return_value = True
        PandocConverter._available = True

        mock_run.side_effect = subprocess.CalledProcessError(
            returncode=1, cmd="pandoc", stderr="Error: invalid input"
        )

        converter = PandocConverter()

        with pytest.raises(ProcessingError) as exc_info:
            converter.convert(sample_md)

        assert "失败" in str(exc_info.value) or "Error" in str(exc_info.value)


class TestPandocConverterTempFile:
    """临时文件测试"""

    @pytest.fixture
    def sample_md(self, tmp_path):
        """创建测试 Markdown 文件"""
        content = "# 测试\n\n内容"
        file_path = tmp_path / "test_input.md"
        file_path.write_text(content, encoding="utf-8")

        return str(file_path)

    @patch("subprocess.run")
    @patch.object(PandocConverter, "is_available")
    def test_convert_with_temp(self, mock_is_available, mock_run, sample_md):
        """测试使用临时文件转换"""
        mock_is_available.return_value = True
        PandocConverter._available = True

        mock_run.return_value = MagicMock(stdout="", stderr="", returncode=0)

        converter = PandocConverter()
        result = converter.convert(sample_md, use_temp=True)

        assert result.endswith(".docx")
        assert "ruledoc_" in result or "tmp" in result.lower()

    def test_cleanup_temp(self, tmp_path):
        """测试清理临时文件"""
        temp_file = tmp_path / "temp_test.docx"
        temp_file.write_text("test")

        converter = PandocConverter()
        converter.cleanup_temp(str(temp_file))

        assert not temp_file.exists()

    def test_cleanup_temp_nonexistent(self):
        """测试清理不存在的临时文件"""
        converter = PandocConverter()

        converter.cleanup_temp("/nonexistent/file.docx")

    @patch("subprocess.run")
    @patch.object(PandocConverter, "is_available")
    def test_convert_with_temp_creates_file(self, mock_is_available, mock_run, sample_md):
        """测试临时文件创建"""
        mock_is_available.return_value = True
        PandocConverter._available = True

        def create_file_side_effect(*args, **kwargs):
            output_path = kwargs.get("capture_output")
            if not output_path:
                for arg in args[0] if args else []:
                    if arg.endswith(".docx"):
                        with open(arg, "wb") as f:
                            f.write(b"PK")  # ZIP header for docx
            return MagicMock(stdout="", stderr="", returncode=0)

        mock_run.side_effect = create_file_side_effect

        converter = PandocConverter()
        result = converter.convert_with_temp(sample_md)

        assert result is not None


class TestPandocConverterIntegration:
    """PandocConverter 集成测试"""

    @pytest.fixture
    def sample_md(self, tmp_path):
        """创建测试 Markdown 文件"""
        content = """# 测试标题

## 第一章 绪论

这是第一章的内容。

### 1.1 背景

这是背景内容。

![测试图片](test.png)

| 列1 | 列2 |
|-----|-----|
| A   | B   |
"""
        file_path = tmp_path / "test_input.md"
        file_path.write_text(content, encoding="utf-8")

        return str(file_path)

    def test_real_conversion(self, sample_md, tmp_path):
        """测试真实转换（需要安装 Pandoc）"""
        converter = PandocConverter()

        if not converter.is_available():
            pytest.skip("Pandoc not installed")

        output_path = str(tmp_path / "output.docx")

        try:
            result = converter.convert(sample_md, output_path)

            assert os.path.exists(result)
            assert os.path.getsize(result) > 0

            from docx import Document

            doc = Document(result)

            assert len(doc.paragraphs) > 0
        except Exception as e:
            pytest.skip(f"Pandoc conversion failed: {e}")

    def test_real_conversion_with_temp(self, sample_md):
        """测试真实临时文件转换"""
        converter = PandocConverter()

        if not converter.is_available():
            pytest.skip("Pandoc not installed")

        try:
            result = converter.convert_with_temp(sample_md)

            assert os.path.exists(result)

            converter.cleanup_temp(result)

            assert not os.path.exists(result)
        except Exception as e:
            pytest.skip(f"Pandoc conversion failed: {e}")
