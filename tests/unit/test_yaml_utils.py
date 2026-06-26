#!/usr/bin/env python3
"""Тесты для yaml_utils."""

import os
import tempfile
import shutil
import pytest
from src.config_editor.yaml_utils import YamlUtils, _extract_inline_comment


class TestExtractInlineComment:
    """Тесты для _extract_inline_comment."""

    def test_no_comment(self):
        line, comment = _extract_inline_comment("key: value")
        assert line == "key: value"
        assert comment == ""

    def test_with_comment(self):
        line, comment = _extract_inline_comment("key: value # comment")
        assert line == "key: value"
        assert comment == "# comment"

    def test_comment_only(self):
        line, comment = _extract_inline_comment("# this is a comment")
        assert line == "# this is a comment"
        assert comment == ""

    def test_empty_line(self):
        line, comment = _extract_inline_comment("")
        assert line == ""
        assert comment == ""

    def test_comment_in_string(self):
        line, comment = _extract_inline_comment('key: "value # not comment"')
        assert line == 'key: "value # not comment"'
        assert comment == ""


class TestYamlUtils:
    """Тесты для YamlUtils."""

    def setup_method(self):
        self.temp_dir = tempfile.mkdtemp()

    def teardown_method(self):
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_load_with_comments(self):
        path = os.path.join(self.temp_dir, 'test.yaml')
        with open(path, 'w') as f:
            f.write("# Header comment\nkey: value\n")

        data, comments = YamlUtils.load_with_comments(path)
        assert data == {'key': 'value'}
        assert '__header__' in comments
        assert 'Header comment' in comments['__header__']

    def test_load_without_comments(self):
        path = os.path.join(self.temp_dir, 'test.yaml')
        with open(path, 'w') as f:
            f.write("key: value\n")

        data, comments = YamlUtils.load_with_comments(path)
        assert data == {'key': 'value'}

    def test_save_with_comments(self):
        path = os.path.join(self.temp_dir, 'test.yaml')
        data = {'key': 'value'}
        comments = {'__header__': '# My header'}

        YamlUtils.save_with_comments(path, data, comments)

        with open(path, 'r') as f:
            content = f.read()
        assert '# My header' in content
        assert 'key: value' in content

    def test_save_without_comments(self):
        path = os.path.join(self.temp_dir, 'test.yaml')
        data = {'key': 'value'}

        YamlUtils.save_with_comments(path, data)

        with open(path, 'r') as f:
            content = f.read()
        assert 'key: value' in content

    def test_extract_comments(self):
        yaml_text = "# Comment 1\nkey: value\n# Comment 2\n"
        result = YamlUtils.extract_comments(yaml_text)
        assert len(result) == 2

    def test_extract_comments_empty(self):
        yaml_text = "key: value\n"
        result = YamlUtils.extract_comments(yaml_text)
        assert len(result) == 0
