#!/usr/bin/env python3
"""Тесты для config_editor (validators и model)."""

import os
import tempfile
import shutil
import pytest
from src.config_editor.model import (
    ConfigModel, FieldType, FieldDescriptor,
    _make_field, _build_font_fields, _build_paragraph_fields
)
from src.config_editor.validators import Validators, ValidationError


class TestFieldDescriptor:
    """Тесты FieldDescriptor."""

    def test_create_string_field(self):
        fd = FieldDescriptor(
            key='name', label='Name', field_type=FieldType.STRING,
            value='Times New Roman'
        )
        assert fd.key == 'name'
        assert fd.field_type == FieldType.STRING
        assert fd.value == 'Times New Roman'

    def test_create_number_field_with_limits(self):
        fd = FieldDescriptor(
            key='size', label='Size', field_type=FieldType.NUMBER,
            value=12, min_value=1, max_value=100
        )
        assert fd.min_value == 1
        assert fd.max_value == 100


class TestMakeField:
    """Тесты функции _make_field."""

    def test_auto_detect_string(self):
        fd = _make_field('name', 'Name', 'Times New Roman')
        assert fd.field_type == FieldType.STRING

    def test_auto_detect_number(self):
        fd = _make_field('size', 'Size', 12)
        assert fd.field_type == FieldType.NUMBER

    def test_auto_detect_bool(self):
        fd = _make_field('bold', 'Bold', True)
        assert fd.field_type == FieldType.BOOLEAN

    def test_auto_detect_dict(self):
        fd = _make_field('data', 'Data', {'key': 'value'})
        assert fd.field_type == FieldType.DICTIONARY

    def test_auto_detect_list(self):
        fd = _make_field('items', 'Items', ['a', 'b'])
        assert fd.field_type == FieldType.LIST

    def test_explicit_field_type(self):
        fd = _make_field('align', 'Align', 'left', field_type=FieldType.ENUM)
        assert fd.field_type == FieldType.ENUM


class TestBuildFontFields:
    """Тесты функции _build_font_fields."""

    def test_build_from_data(self):
        data = {'name': 'Arial', 'size': 14, 'bold': True}
        fields = _build_font_fields(data)

        assert 'name' in fields
        assert 'size' in fields
        assert 'bold' in fields
        assert fields['name'].value == 'Arial'
        assert fields['size'].value == 14
        assert fields['bold'].value is True

    def test_build_empty_data(self):
        fields = _build_font_fields({})
        assert fields == {}

    def test_build_none_data(self):
        fields = _build_font_fields(None)
        assert fields == {}


class TestBuildParagraphFields:
    """Тесты функции _build_paragraph_fields."""

    def test_build_from_data(self):
        data = {'first_line_indent_cm': 1.25, 'line_spacing': 1.15, 'alignment': 'justify'}
        fields = _build_paragraph_fields(data)

        assert 'first_line_indent_cm' in fields
        assert 'line_spacing' in fields
        assert 'alignment' in fields
        assert fields['first_line_indent_cm'].value == 1.25
        assert fields['alignment'].value == 'justify'

    def test_build_empty_data(self):
        fields = _build_paragraph_fields({})
        assert fields == {}


class TestValidators:
    """Тесты для Validators."""

    def test_validate_string_valid(self):
        fd = FieldDescriptor(key='name', label='Name', field_type=FieldType.STRING,
                           value='test', required=True)
        errors = Validators.validate_string('test', fd)
        assert len(errors) == 0

    def test_validate_string_empty_required(self):
        fd = FieldDescriptor(key='name', label='Name', field_type=FieldType.STRING,
                           value='', required=True)
        errors = Validators.validate_string('', fd)
        assert len(errors) == 1
        assert 'обязательно' in errors[0].message.lower()

    def test_validate_number_valid(self):
        fd = FieldDescriptor(key='size', label='Size', field_type=FieldType.NUMBER,
                           value=12, min_value=1, max_value=100)
        errors = Validators.validate_number(12, fd)
        assert len(errors) == 0

    def test_validate_number_below_min(self):
        fd = FieldDescriptor(key='size', label='Size', field_type=FieldType.NUMBER,
                           value=0, min_value=1, max_value=100)
        errors = Validators.validate_number(0, fd)
        assert len(errors) == 1
        assert 'минимум' in errors[0].message.lower()

    def test_validate_number_above_max(self):
        fd = FieldDescriptor(key='size', label='Size', field_type=FieldType.NUMBER,
                           value=200, min_value=1, max_value=100)
        errors = Validators.validate_number(200, fd)
        assert len(errors) == 1
        assert 'максимум' in errors[0].message.lower()

    def test_validate_number_not_a_number(self):
        fd = FieldDescriptor(key='size', label='Size', field_type=FieldType.NUMBER,
                           value='abc')
        errors = Validators.validate_number('abc', fd)
        assert len(errors) == 1
        assert 'числом' in errors[0].message.lower()

    def test_validate_enum_valid(self):
        fd = FieldDescriptor(key='align', label='Align', field_type=FieldType.ENUM,
                           value='left', enum_values=['left', 'center', 'right'])
        errors = Validators.validate_enum('left', fd)
        assert len(errors) == 0

    def test_validate_enum_invalid(self):
        fd = FieldDescriptor(key='align', label='Align', field_type=FieldType.ENUM,
                           value='invalid', enum_values=['left', 'center', 'right'])
        errors = Validators.validate_enum('invalid', fd)
        assert len(errors) == 1

    def test_validate_color_hex_valid(self):
        errors = Validators.validate_color('#FF0000')
        assert len(errors) == 0

    def test_validate_color_hex_short_valid(self):
        errors = Validators.validate_color('#F00')
        assert len(errors) == 0

    def test_validate_color_hex_invalid(self):
        errors = Validators.validate_color('#GG0000')
        assert len(errors) == 1

    def test_validate_color_name_valid(self):
        errors = Validators.validate_color('red')
        assert len(errors) == 0

    def test_validate_color_none(self):
        errors = Validators.validate_color(None)
        assert len(errors) == 0

    def test_validate_color_invalid_format(self):
        errors = Validators.validate_color('12345')
        assert len(errors) == 1


class TestConfigModel:
    """Тесты для ConfigModel."""

    def setup_method(self):
        self.temp_dir = tempfile.mkdtemp()
        self.config_path = os.path.join(self.temp_dir, 'test_config.yaml')
        with open(self.config_path, 'w', encoding='utf-8') as f:
            f.write("""
styles:
  Normal:
    enabled: true
    font:
      name: "Times New Roman"
      size: 12
      bold: false
      italic: false
      color: "black"
    paragraph:
      first_line_indent_cm: 1.25
      line_spacing: 1.15
      space_before_pt: 0
      space_after_pt: 0
      alignment: "justify"
  Heading 1:
    enabled: true
    font:
      name: "Calibri Light"
      size: 16
      bold: true
detection_rules:
  Heading 1:
    - "РАЗДЕЛ "
page_setup:
  portrait:
    left_margin_cm: 3.0
""")

    def teardown_method(self):
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_load(self):
        model = ConfigModel(self.config_path)
        result = model.load()
        assert result is True
        assert 'Normal' in model.styles
        assert 'Heading 1' in model.styles

    def test_get_style(self):
        model = ConfigModel(self.config_path)
        model.load()

        style = model.get_style('Normal')
        assert style is not None
        assert style.name == 'Normal'
        assert style.enabled is True
        assert style.font['name'].value == 'Times New Roman'
        assert style.font['size'].value == 12

    def test_get_nonexistent_style(self):
        model = ConfigModel(self.config_path)
        model.load()

        style = model.get_style('NonExistent')
        assert style is None

    def test_set_style_enabled(self):
        model = ConfigModel(self.config_path)
        model.load()

        model.set_style_enabled('Normal', False)
        assert model.styles['Normal'].enabled is False
        assert model.modified is True

    def test_update_style_field(self):
        model = ConfigModel(self.config_path)
        model.load()

        model.update_style_field('Normal', 'font', 'size', 14)
        assert model.styles['Normal'].font['size'].value == 14
        assert model.raw_config['styles']['Normal']['font']['size'] == 14

    def test_add_style(self):
        model = ConfigModel(self.config_path)
        model.load()

        model.add_style('Custom Style')
        assert 'Custom Style' in model.styles
        assert model.styles['Custom Style'].enabled is True

    def test_delete_style(self):
        model = ConfigModel(self.config_path)
        model.load()

        model.delete_style('Heading 1')
        assert 'Heading 1' not in model.styles
        assert model.modified is True

    def test_to_dict(self):
        model = ConfigModel(self.config_path)
        model.load()

        d = model.to_dict()
        assert 'styles' in d
        assert 'Normal' in d['styles']

    def test_save(self):
        model = ConfigModel(self.config_path)
        model.load()

        model.update_style_field('Normal', 'font', 'size', 14)
        model.save()

        model2 = ConfigModel(self.config_path)
        model2.load()
        assert model2.styles['Normal'].font['size'].value == 14

    def test_sections(self):
        model = ConfigModel(self.config_path)
        model.load()

        assert 'detection_rules' in model.sections
        assert 'page_setup' in model.sections
        assert model.sections['page_setup']['portrait']['left_margin_cm'] == 3.0

    def test_update_section(self):
        model = ConfigModel(self.config_path)
        model.load()

        model.update_section('page_setup', {'portrait': {'left_margin_cm': 2.5}})
        assert model.sections['page_setup']['portrait']['left_margin_cm'] == 2.5
        assert model.modified is True

    def test_modified_flag(self):
        model = ConfigModel(self.config_path)
        model.load()
        assert model.modified is False

        model.set_style_enabled('Normal', False)
        assert model.modified is True

        model.save()
        assert model.modified is False
