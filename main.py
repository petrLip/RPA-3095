#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
RPA-3095 V2 - Корректировка CF16
Главный файл запуска приложения

Python-версия VBA макросов для обработки Excel данных.
Поддерживает Windows и Linux.

Использование:
    python main.py              # Запуск GUI приложения
    python main.py --test       # ТЕСТОВЫЙ РЕЖИМ - автопоиск файлов в data/
    python main.py --cli 1      # Запуск блока 1 из командной строки
    python main.py --cli 2      # Запуск блока 2 из командной строки
    python main.py --help       # Справка
"""

import sys
import argparse
from pathlib import Path

# Добавляем корневую директорию в путь поиска модулей
sys.path.insert(0, str(Path(__file__).parent))

from src.logger import log, setup_logger

# Папка с тестовыми данными
DATA_DIR = Path(__file__).parent / "data"


def find_test_files():
    """Автоматический поиск файлов для тестирования в папке data/"""
    if not DATA_DIR.exists():
        return None, None, None

    macros_file = None
    marja_file = None
    vgo_file = None

    for f in DATA_DIR.iterdir():
        name = f.name.lower()
        # Пропускаем временные файлы и результаты
        if name.startswith("~$") or "_opus" in name:
            continue

        if f.suffix == ".xlsm" and "корректировка" in name.lower():
            macros_file = str(f)
        elif f.suffix == ".xlsx" and "маржа" in name.lower():
            marja_file = str(f)
        elif f.suffix == ".xlsb" and (
            "отчет" in name.lower() or "выверк" in name.lower()
        ):
            vgo_file = str(f)

    return macros_file, marja_file, vgo_file


def parse_args():
    """Парсинг аргументов командной строки"""
    parser = argparse.ArgumentParser(
        description="RPA-3095 V2 - Корректировка CF16",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Примеры использования:
  python main.py                                    # Запуск GUI
  python main.py --test                             # Тестовый режим (автопоиск файлов)
  python main.py --cli 1 --macros file.xlsm --marja marja.xlsx --vgo vgo.xlsb
  python main.py --cli 2 --macros file.xlsm
        """,
    )

    parser.add_argument(
        "--test",
        action="store_true",
        help="Тестовый режим: автоматически находит файлы в папке data/ и запускает блок 1",
    )

    parser.add_argument(
        "--cli",
        type=int,
        choices=[1, 2],
        help="Запуск в режиме командной строки. 1 - Создать предварительные листы, 2 - Создать корректировку",
    )

    parser.add_argument(
        "--macros", type=str, help="Путь к основному файлу с макросами (.xlsm)"
    )

    parser.add_argument(
        "--marja", type=str, help="Путь к файлу с листом Маржа (для блока 1)"
    )

    parser.add_argument(
        "--vgo", type=str, help="Путь к файлу выверки ВГО (для блока 1)"
    )

    parser.add_argument(
        "--log-level",
        type=str,
        default="INFO",
        choices=["DEBUG", "INFO", "WARNING", "ERROR"],
        help="Уровень логирования (по умолчанию: INFO)",
    )

    return parser.parse_args()


def run_cli(args):
    """Запуск в режиме командной строки"""
    from src.create_preview_data import create_preview_data
    from src.unload_corr import unload_corr

    def progress_callback(percent, message):
        print(f"[{percent:3d}%] {message}")

    if args.cli == 1:
        # Блок 1: Создать предварительные листы
        if not args.macros or not args.marja or not args.vgo:
            print(
                "Ошибка: Для блока 1 необходимо указать все файлы: --macros, --marja, --vgo"
            )
            sys.exit(1)

        # Проверяем существование файлов
        for path, name in [
            (args.macros, "macros"),
            (args.marja, "marja"),
            (args.vgo, "vgo"),
        ]:
            if not Path(path).exists():
                print(f"Ошибка: Файл {name} не найден: {path}")
                sys.exit(1)

        log.info("Запуск блока 1: Создание предварительных листов...")
        result = create_preview_data(
            macros_file=args.macros,
            marja_file=args.marja,
            vgo_file=args.vgo,
            progress_callback=progress_callback,
        )

    elif args.cli == 2:
        # Блок 2: Создать корректировку
        if not args.macros:
            print("Ошибка: Для блока 2 необходимо указать файл: --macros")
            sys.exit(1)

        if not Path(args.macros).exists():
            print(f"Ошибка: Файл не найден: {args.macros}")
            sys.exit(1)

        log.info("Запуск блока 2: Создание корректировки CF16...")
        result = unload_corr(
            macros_file=args.macros, progress_callback=progress_callback
        )

    # Вывод результата
    if result.success:
        print(f"\n✅ Успех: {result.message}")
        sys.exit(0)
    else:
        print(f"\n❌ Ошибка: {', '.join(result.errors)}")
        sys.exit(1)


def run_test():
    """Тестовый режим: автоматический поиск файлов и запуск"""
    from src.create_preview_data import create_preview_data

    def progress_callback(percent, message):
        print(f"[{percent:3d}%] {message}")

    print("\n" + "=" * 60)
    print("🧪 ТЕСТОВЫЙ РЕЖИМ")
    print("=" * 60)

    # Ищем файлы
    macros_file, marja_file, vgo_file = find_test_files()

    print(f"\n📁 Папка данных: {DATA_DIR}")
    print(f"📄 Основной файл: {Path(macros_file).name if macros_file else 'НЕ НАЙДЕН'}")
    print(f"📄 Файл Маржа:    {Path(marja_file).name if marja_file else 'НЕ НАЙДЕН'}")
    print(f"📄 Файл ВГО:      {Path(vgo_file).name if vgo_file else 'НЕ НАЙДЕН'}")
    print()

    if not all([macros_file, marja_file, vgo_file]):
        print("Ошибка: Не все файлы найдены в папке data/")
        print("\nОжидаемые файлы:")
        print("  - .xlsm файл с 'корректировка' в названии")
        print("  - .xlsx файл с 'маржа' в названии")
        print("  - .xlsb файл с 'отчет' или 'выверк' в названии")
        sys.exit(1)

    print("Запуск обработки...\n")

    result = create_preview_data(
        macros_file=macros_file,
        marja_file=marja_file,
        vgo_file=vgo_file,
        progress_callback=progress_callback,
    )

    print()
    if result.success:
        print(f"Успех: {result.message}")
        if hasattr(result, "output_file") and result.output_file:
            print(f"Результат: {result.output_file}")
        sys.exit(0)
    else:
        print(f"Ошибка: {', '.join(result.errors)}")
        sys.exit(1)


def run_gui():
    """Запуск GUI приложения"""
    try:
        from src.gui import run_app

        log.info("Запуск GUI приложения...")
        run_app()
    except ImportError as e:
        log.error(f"Ошибка импорта GUI: {e}")
        print("Ошибка: Не удалось запустить GUI. Проверьте установку PySide6.")
        print("Установка: pip install PySide6")
        sys.exit(1)


def main():
    """Главная функция"""
    # Парсим аргументы
    args = parse_args()

    # Инициализируем логгер
    setup_logger()

    log.info("=" * 60)
    log.info("RPA-3095 V2 - Корректировка CF16")
    log.info("=" * 60)

    if args.test:
        # Тестовый режим - автопоиск файлов
        run_test()
    elif args.cli:
        # Режим командной строки
        run_cli(args)
    else:
        # GUI режим
        run_gui()


if __name__ == "__main__":
    main()
