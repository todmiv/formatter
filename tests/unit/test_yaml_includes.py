#!/usr/bin/env python3
"""Тесты для yaml_includes."""

import os
import tempfile
import shutil
import pytest
from src.core.yaml_includes import (
    deep_merge, resolve_path, load_yaml_with_includes, get_included_paths
)


class TestDeepMerge:
    """Тесты deep_merge."""

    def test_simple_merge(self):
        base = {'a': 1, 'b': 2}
        override = {'b': 3, 'c': 4}
        result = deep_merge(base, override)
        assert result == {'a': 1, 'b': 3, 'c': 4}

    def test_nested_merge(self):
        base = {'a': {'x': 1, 'y': 2}, 'b': 1}
        override = {'a': {'y': 3, 'z': 4}}
        result = deep_merge(base, override)
        assert result == {'a': {'x': 1, 'y': 3, 'z': 4}, 'b': 1}

    def test_override_non_dict(self):
        base = {'a': [1, 2]}
        override = {'a': [3, 4]}
        result = deep_merge(base, override)
        assert result == {'a': [3, 4]}

    def test_empty_base(self):
        override = {'a': 1}
        result = deep_merge({}, override)
        assert result == {'a': 1}

    def test_empty_override(self):
        base = {'a': 1}
        result = deep_merge(base, {})
        assert result == {'a': 1}


class TestResolvePath:
    """Тесты resolve_path."""

    def test_relative_path(self):
        result = resolve_path('/home/user/config.yaml', 'includes/base.yaml')
        assert 'includes' in result
        assert 'base.yaml' in result

    def test_absolute_path(self):
        result = resolve_path('/home/user/config.yaml', '/etc/base.yaml')
        assert result == '/etc/base.yaml'


class TestLoadYamlWithIncludes:
    """Тесты load_yaml_with_includes."""

    def setup_method(self):
        self.temp_dir = tempfile.mkdtemp()

    def teardown_method(self):
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_simple_load(self):
        config_path = os.path.join(self.temp_dir, 'config.yaml')
        with open(config_path, 'w') as f:
            f.write('key: value\n')

        result = load_yaml_with_includes(config_path)
        assert result == {'key': 'value'}

    def test_load_with_includes(self):
        base_path = os.path.join(self.temp_dir, 'base.yaml')
        with open(base_path, 'w') as f:
            f.write('base_key: base_value\n')

        main_path = os.path.join(self.temp_dir, 'main.yaml')
        with open(main_path, 'w') as f:
            f.write(f'includes:\n  - base.yaml\nmain_key: main_value\n')

        result = load_yaml_with_includes(main_path)
        assert result == {'base_key': 'base_value', 'main_key': 'main_value'}

    def test_local_overrides_included(self):
        base_path = os.path.join(self.temp_dir, 'base.yaml')
        with open(base_path, 'w') as f:
            f.write('key: base_value\n')

        main_path = os.path.join(self.temp_dir, 'main.yaml')
        with open(main_path, 'w') as f:
            f.write(f'includes:\n  - base.yaml\nkey: main_value\n')

        result = load_yaml_with_includes(main_path)
        assert result['key'] == 'main_value'

    def test_circular_includes(self):
        a_path = os.path.join(self.temp_dir, 'a.yaml')
        b_path = os.path.join(self.temp_dir, 'b.yaml')

        with open(a_path, 'w') as f:
            f.write('includes:\n  - b.yaml\nkey_a: a\n')
        with open(b_path, 'w') as f:
            f.write('includes:\n  - a.yaml\nkey_b: b\n')

        result = load_yaml_with_includes(a_path)
        assert 'key_b' in result

    def test_missing_file_raises_error(self):
        with pytest.raises(Exception):
            load_yaml_with_includes('/nonexistent/config.yaml')


class TestGetIncludedPaths:
    """Тесты get_included_paths."""

    def setup_method(self):
        self.temp_dir = tempfile.mkdtemp()

    def teardown_method(self):
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_no_includes(self):
        config_path = os.path.join(self.temp_dir, 'config.yaml')
        with open(config_path, 'w') as f:
            f.write('key: value\n')

        paths = get_included_paths(config_path)
        assert paths == []

    def test_with_includes(self):
        base_path = os.path.join(self.temp_dir, 'base.yaml')
        with open(base_path, 'w') as f:
            f.write('key: value\n')

        main_path = os.path.join(self.temp_dir, 'main.yaml')
        with open(main_path, 'w') as f:
            f.write('includes:\n  - base.yaml\n')

        paths = get_included_paths(main_path)
        assert len(paths) == 1
        assert paths[0].endswith('base.yaml')

    def test_nonexistent_file(self):
        paths = get_included_paths('/nonexistent/config.yaml')
        assert paths == []
