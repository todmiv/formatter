import os
import pytest
from unittest.mock import MagicMock
from src.core.main_controller import MainController


def test_successful_conversion(tmp_path):
    config_src = "configs/active/config.yaml"
    config_dst = tmp_path / "config.yaml"
    config_dst.write_text(open(config_src, "r", encoding="utf-8").read(), encoding="utf-8")
    md_content = "# Тестовый раздел\n\n## Подраздел\n\nОбычный текст параграфа.\n\n- Элемент списка 1\n- Элемент списка 2\n"
    md_path = tmp_path / "test.md"
    md_path.write_text(md_content, encoding="utf-8")
    docx_path = tmp_path / "output.docx"
    controller = MainController(str(config_dst))
    result = controller.import_markdown(str(md_path), str(docx_path))
    assert result is True
    assert docx_path.exists()
    assert docx_path.stat().st_size > 0


def test_missing_input_file(tmp_path):
    config_src = "configs/active/config.yaml"
    config_dst = tmp_path / "config.yaml"
    config_dst.write_text(open(config_src, "r", encoding="utf-8").read(), encoding="utf-8")
    controller = MainController(str(config_dst))
    with pytest.raises(FileNotFoundError):
        controller.import_markdown(
            str(tmp_path / "nonexistent.md"),
            str(tmp_path / "output.docx")
        )


def test_progress_callback_called(tmp_path):
    config_src = "configs/active/config.yaml"
    config_dst = tmp_path / "config.yaml"
    config_dst.write_text(open(config_src, "r", encoding="utf-8").read(), encoding="utf-8")
    md_path = tmp_path / "test.md"
    md_path.write_text("# Заголовок\n\nТекст абзаца.\n", encoding="utf-8")
    docx_path = tmp_path / "output.docx"
    controller = MainController(str(config_dst))
    callback = MagicMock()
    result = controller.import_markdown(str(md_path), str(docx_path), progress_callback=callback)
    assert result is True
    callback.assert_called()


def test_empty_markdown_file(tmp_path):
    config_src = "configs/active/config.yaml"
    config_dst = tmp_path / "config.yaml"
    config_dst.write_text(open(config_src, "r", encoding="utf-8").read(), encoding="utf-8")
    md_path = tmp_path / "empty.md"
    md_path.write_text("", encoding="utf-8")
    docx_path = tmp_path / "output.docx"
    controller = MainController(str(config_dst))
    result = controller.import_markdown(str(md_path), str(docx_path))
    assert result is True
    assert docx_path.exists()


def test_invalid_config(tmp_path):
    from src.core.formatter_errors import FormatterError
    config_path = tmp_path / "invalid_config.yaml"
    config_path.write_text("name: test\n", encoding="utf-8")
    with pytest.raises(FormatterError):
        MainController(str(config_path))