"""
Модуль логирования и тайминга операций
"""

import sys
import time
from functools import wraps
from pathlib import Path
from datetime import datetime
from loguru import logger


def setup_logger(log_dir: str = "logs"):
    """Настройка логгера с выводом в файл и консоль"""
    log_path = Path(log_dir)
    log_path.mkdir(exist_ok=True)

    # Удаляем стандартный обработчик
    logger.remove()

    # Формат для логов
    log_format = (
        "<green>{time:YYYY-MM-DD HH:mm:ss.SSS}</green> | "
        "<level>{level: <8}</level> | "
        "<cyan>{name}</cyan>:<cyan>{function}</cyan>:<cyan>{line}</cyan> | "
        "<level>{message}</level>"
    )

    # Консольный вывод
    logger.add(sys.stdout, format=log_format, level="INFO", colorize=True)

    # Файловый вывод
    log_file = log_path / f"cf16_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
    logger.add(
        str(log_file),
        format=log_format,
        level="DEBUG",
        rotation="10 MB",
        retention="7 days",
        encoding="utf-8",
    )

    logger.info(f"Логирование инициализировано. Файл: {log_file}")
    return logger


def timing(func):
    """Декоратор для замера времени выполнения функции"""

    @wraps(func)
    def wrapper(*args, **kwargs):
        start_time = time.perf_counter()
        logger.info(f"🚀 Начало: {func.__name__}")

        try:
            result = func(*args, **kwargs)
            elapsed = time.perf_counter() - start_time
            logger.success(f"✅ Завершено: {func.__name__} за {elapsed:.2f} сек")
            return result
        except Exception as e:
            elapsed = time.perf_counter() - start_time
            logger.error(f"❌ Ошибка в {func.__name__} после {elapsed:.2f} сек: {e}")
            raise

    return wrapper


class Timer:
    """Контекстный менеджер для замера времени блока кода"""

    def __init__(self, operation_name: str):
        self.operation_name = operation_name
        self.start_time = None
        self.elapsed = None

    def __enter__(self):
        self.start_time = time.perf_counter()
        logger.info(f"⏱️ Старт операции: {self.operation_name}")
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.elapsed = time.perf_counter() - self.start_time
        if exc_type is None:
            logger.info(
                f"⏱️ Операция '{self.operation_name}' завершена за {self.elapsed:.2f} сек"
            )
        else:
            logger.warning(
                f"⏱️ Операция '{self.operation_name}' прервана с ошибкой после {self.elapsed:.2f} сек"
            )
        return False


# Инициализация логгера при импорте модуля
log = setup_logger()
