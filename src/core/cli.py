#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
CLI-интерфейс для форматирования документов по ГОСТ.
Поддерживает все режимы: config-based, template-based, аудит, пакетная обработка.

Использование:
    python -m src.core.cli format -i input.docx -o output.docx -c config.yaml
    python -m src.core.cli format -i input.docx -o output.docx -t template.docx
    python -m src.core.cli audit -i input.docx -o output.docx -c config.yaml
    python -m src.core.cli extract -t template.docx -o config.yaml
    python -m src.core.cli diff -i input.docx -t template.docx
    python -m src.core.cli batch -id input_dir/ -od output_dir/ -c config.yaml
"""

import os
import sys
import logging
import argparse

from src.core.document_formatter import DocumentFormatter


def cmd_format(args):
    """Команда форматирования документа."""
    formatter = DocumentFormatter(
        config_path=args.config,
        template_path=args.template,
        verbose=args.verbose,
    )

    result = formatter.format_document(
        args.input,
        args.output,
        options={'clear_direct_formatting': args.clear_formatting},
    )

    if result['success']:
        print(f"Форматирование выполнено: {result['output']}")
        print(f"Режим: {result['mode']}, время: {result['duration']:.2f} сек")
        stats = result['stats']
        if 'styles_applied' in stats:
            print(f"Скопировано стилей: {stats['styles_applied']}")
        if 'applied' in stats:
            print(f"Исправлений: {stats['applied']}")
    else:
        print(f"Ошибка: {result.get('error')}", file=sys.stderr)
        return 1
    return 0


def cmd_audit(args):
    """Команда форматирования с аудитом."""
    if not args.config:
        print("Для аудита необходим --config", file=sys.stderr)
        return 1

    formatter = DocumentFormatter(config_path=args.config, verbose=args.verbose)

    result = formatter.format_with_audit(
        args.input,
        args.output,
        max_iterations=args.max_iterations,
    )

    if result['success']:
        print(f"Форматирование выполнено: {result['output']}")
        for it in result['iterations']:
            print(
                f"  Итерация {it['iteration']}: "
                f"найдено {it['issues_found']}, "
                f"исправлено {it['issues_fixed']}"
            )
        final = result['final_stats']
        print(f"Оставшихся проблем: {final['remaining_issues']}")
    else:
        print(f"Ошибка: {result.get('error')}", file=sys.stderr)
        return 1
    return 0


def cmd_extract(args):
    """Команда извлечения конфига из шаблона."""
    formatter = DocumentFormatter(verbose=args.verbose)
    config = formatter.extract_config(args.template, args.output)
    print(f"Конфигурация извлечена: {args.output}")
    print(f"Стилей: {len(config.get('styles', {}))}")
    print(f"Правил определения: {len(config.get('detection_rules', {}))}")
    return 0


def cmd_diff(args):
    """Команда сравнения документа с шаблоном."""
    formatter = DocumentFormatter(
        template_path=args.template,
        verbose=args.verbose,
    )

    result = formatter.diff(args.input)

    print("Расхождения с шаблоном:")
    print(f"  Отсутствующие стили ({len(result['missing_styles'])}):")
    for name in result['missing_styles']:
        print(f"    - {name}")
    print(f"  Различающиеся стили ({len(result['different_styles'])}):")
    for name in result['different_styles']:
        print(f"    - {name}")
    print(
        f"  Настройки страницы: "
        f"{'различаются' if result['page_setup_diff'] else 'совпадают'}"
    )
    return 0


def cmd_batch(args):
    """Команда пакетной обработки."""
    formatter = DocumentFormatter(
        config_path=args.config,
        template_path=args.template,
        verbose=args.verbose,
    )

    results = formatter.batch_format(
        args.input_dir,
        args.output_dir,
        options={'clear_direct_formatting': args.clear_formatting},
    )

    success = sum(1 for r in results if r['success'])
    total = len(results)
    print(f"Пакетная обработка: {success}/{total} успешно")

    for r in results:
        status = "OK" if r['success'] else "ОШИБКА"
        name = os.path.basename(r['input'])
        print(f"  [{status}] {name}")

    return 0 if success == total else 1


def main():
    from src.core.formatter_errors import FormatterError, exit_with_error

    parser = argparse.ArgumentParser(
        description='Форматирование DOCX-документов по ГОСТ',
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument('-v', '--verbose', action='store_true', help='Подробный вывод')

    subparsers = parser.add_subparsers(dest='command', help='Команда')

    p_format = subparsers.add_parser('format', help='Форматирование документа')
    p_format.add_argument('-i', '--input', required=True, help='Входной DOCX')
    p_format.add_argument('-o', '--output', help='Выходной DOCX')
    p_format.add_argument('-c', '--config', help='YAML-конфиг')
    p_format.add_argument('-t', '--template', help='DOCX-шаблон')
    p_format.add_argument('--clear-formatting', action='store_true', help='Очистить прямое форматирование')

    p_audit = subparsers.add_parser('audit', help='Форматирование с аудитом')
    p_audit.add_argument('-i', '--input', required=True, help='Входной DOCX')
    p_audit.add_argument('-o', '--output', help='Выходной DOCX')
    p_audit.add_argument('-c', '--config', required=True, help='YAML-конфиг')
    p_audit.add_argument('--max-iterations', type=int, default=3, help='Макс. итераций')

    p_extract = subparsers.add_parser('extract', help='Извлечение конфига из шаблона')
    p_extract.add_argument('-t', '--template', required=True, help='DOCX-шаблон')
    p_extract.add_argument('-o', '--output', required=True, help='YAML-файл')

    p_diff = subparsers.add_parser('diff', help='Сравнение с шаблоном')
    p_diff.add_argument('-i', '--input', required=True, help='Входной DOCX')
    p_diff.add_argument('-t', '--template', required=True, help='DOCX-шаблон')

    p_batch = subparsers.add_parser('batch', help='Пакетная обработка')
    p_batch.add_argument('-id', '--input-dir', required=True, help='Входная директория')
    p_batch.add_argument('-od', '--output-dir', required=True, help='Выходная директория')
    p_batch.add_argument('-c', '--config', help='YAML-конфиг')
    p_batch.add_argument('-t', '--template', help='DOCX-шаблон')
    p_batch.add_argument('--clear-formatting', action='store_true', help='Очистить прямое форматирование')

    args = parser.parse_args()

    if not args.command:
        parser.print_help()
        return 1

    logging.basicConfig(
        level=logging.DEBUG if args.verbose else logging.WARNING,
        format='%(asctime)s [%(levelname)s] %(message)s',
    )

    commands = {
        'format': cmd_format,
        'audit': cmd_audit,
        'extract': cmd_extract,
        'diff': cmd_diff,
        'batch': cmd_batch,
    }

    try:
        return commands[args.command](args)
    except FormatterError as e:
        exit_with_error(e)
    except Exception as e:
        print(f"Непредвиденная ошибка: {e}", file=sys.stderr)
        return 1


if __name__ == '__main__':
    sys.exit(main())
