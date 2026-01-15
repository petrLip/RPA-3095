# -*- coding: utf-8 -*-
"""
Главное окно приложения RPA-3095 V2
Графический интерфейс на PySide6
"""

import sys
from pathlib import Path
from typing import Optional

from PySide6.QtWidgets import (
    QApplication,
    QMainWindow,
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QPushButton,
    QComboBox,
    QProgressBar,
    QTextEdit,
    QGroupBox,
    QFileDialog,
    QMessageBox,
    QFrame,
    QSpacerItem,
    QSizePolicy,
    QScrollArea,
)
from PySide6.QtCore import Qt, QThread, Signal, QSize
from PySide6.QtGui import QFont, QColor, QPalette, QIcon

from src.create_preview_data import create_preview_data, ProcessingResult
from src.unload_corr import unload_corr
from src.logger import log


# Стили CSS для приложения (светлая тема)
STYLESHEET = """
QMainWindow {
    background-color: #ffffff;
}

QWidget {
    font-family: 'Segoe UI', 'SF Pro Display', -apple-system, sans-serif;
    font-size: 12px;
    color: #333333;
}

QGroupBox {
    background-color: #f5f5f5;
    border: 1px solid #e0e0e0;
    border-radius: 6px;
    margin-top: 10px;
    padding: 8px;
    font-weight: bold;
    font-size: 13px;
}

QGroupBox::title {
    subcontrol-origin: margin;
    left: 10px;
    padding: 0 5px;
    color: #333333;
}

QLabel {
    color: #333333;
    font-size: 12px;
}

QLabel#title {
    font-size: 18px;
    font-weight: bold;
    color: #333333;
    padding: 5px;
}

QLabel#subtitle {
    font-size: 14px;
    color: #666666;
    padding-bottom: 20px;
}

QLabel#status {
    background-color: #E0F2F7;
    border: 1px solid #B0D4E0;
    border-radius: 6px;
    padding: 6px;
    color: #2196F3;
    font-weight: bold;
    font-size: 12px;
    text-align: center;
}

QLineEdit {
    background-color: #ffffff;
    border: 1px solid #cccccc;
    border-radius: 6px;
    padding: 6px 10px;
    color: #333333;
    font-size: 12px;
    selection-background-color: #4CAF50;
    selection-color: #ffffff;
}

QLineEdit:focus {
    border-color: #4CAF50;
}

QLineEdit:hover {
    border-color: #999999;
}

QLineEdit:disabled {
    background-color: #f0f0f0;
    color: #999999;
}

QComboBox {
    background-color: #ffffff;
    border: 1px solid #cccccc;
    border-radius: 6px;
    padding: 6px 10px;
    color: #333333;
    font-size: 12px;
}

QComboBox:hover {
    border-color: #999999;
}

QComboBox:focus {
    border-color: #4CAF50;
}

QComboBox::drop-down {
    border: none;
    width: 30px;
}

QComboBox::down-arrow {
    image: none;
    border-left: 5px solid transparent;
    border-right: 5px solid transparent;
    border-top: 6px solid #666666;
    margin-right: 10px;
}

QComboBox QAbstractItemView {
    background-color: #ffffff;
    border: 1px solid #cccccc;
    border-radius: 8px;
    selection-background-color: #4CAF50;
    selection-color: #ffffff;
    color: #333333;
    padding: 5px;
}

QPushButton {
    background-color: #4CAF50;
    border: none;
    border-radius: 6px;
    padding: 8px 15px;
    color: #ffffff;
    font-size: 12px;
    font-weight: bold;
    min-width: 100px;
}

QPushButton:hover {
    background-color: #45a049;
}

QPushButton:pressed {
    background-color: #3d8b40;
}

QPushButton:disabled {
    background-color: #D3D3D3;
    color: #666666;
}

QPushButton#primary {
    background-color: #D3D3D3;
    border: none;
    color: #333333;
    font-size: 13px;
    padding: 10px 20px;
}

QPushButton#primary:hover {
    background-color: #c0c0c0;
}

QPushButton#primary:pressed {
    background-color: #a8a8a8;
}

QPushButton#primary:disabled {
    background-color: #e0e0e0;
    color: #999999;
}

QPushButton#browse {
    min-width: 80px;
    padding: 6px 15px;
}

QProgressBar {
    background-color: #e0e0e0;
    border: none;
    border-radius: 6px;
    height: 22px;
    text-align: center;
    color: #333333;
    font-weight: bold;
    font-size: 11px;
}

QProgressBar::chunk {
    background-color: #4CAF50;
    border-radius: 8px;
}

QTextEdit {
    background-color: #ffffff;
    border: 1px solid #cccccc;
    border-radius: 6px;
    padding: 8px;
    color: #333333;
    font-family: 'Consolas', 'Monaco', 'Courier New', monospace;
    font-size: 11px;
    line-height: 1.3;
}

QTextEdit:focus {
    border-color: #4CAF50;
}

QFrame#separator {
    background-color: #e0e0e0;
    max-height: 2px;
    margin: 10px 0;
}

/* Scrollbars */
QScrollBar:vertical {
    background-color: #f0f0f0;
    width: 12px;
    border-radius: 6px;
}

QScrollBar::handle:vertical {
    background-color: #cccccc;
    border-radius: 6px;
    min-height: 30px;
}

QScrollBar::handle:vertical:hover {
    background-color: #999999;
}

QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
    height: 0;
}

QScrollBar:horizontal {
    background-color: #f0f0f0;
    height: 12px;
    border-radius: 6px;
}

QScrollBar::handle:horizontal {
    background-color: #cccccc;
    border-radius: 6px;
    min-width: 30px;
}

QScrollBar::handle:horizontal:hover {
    background-color: #999999;
}

QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal {
    width: 0;
}

/* Диалоги выбора файла */
QFileDialog {
    background-color: #ffffff;
    color: #333333;
}

QFileDialog QLabel {
    color: #333333;
}

QFileDialog QLineEdit {
    background-color: #ffffff;
    color: #333333;
    border: 1px solid #cccccc;
}

QFileDialog QPushButton {
    background-color: #4CAF50;
    color: #ffffff;
}

QFileDialog QTreeView, QFileDialog QListView {
    background-color: #ffffff;
    color: #333333;
    selection-background-color: #4CAF50;
    selection-color: #ffffff;
}

QFileDialog QHeaderView::section {
    background-color: #f0f0f0;
    color: #333333;
    padding: 5px;
    border: 1px solid #cccccc;
}
"""


class WorkerThread(QThread):
    """Рабочий поток для выполнения длительных операций"""

    progress_updated = Signal(int, str)
    finished_with_result = Signal(object)

    def __init__(
        self, task_type: int, macros_file: str, marja_file: str = "", vgo_file: str = ""
    ):
        super().__init__()
        self.task_type = task_type
        self.macros_file = macros_file
        self.marja_file = marja_file
        self.vgo_file = vgo_file

    def run(self):
        """Выполнение задачи в отдельном потоке"""
        try:
            if self.task_type == 1:
                result = create_preview_data(
                    macros_file=self.macros_file,
                    marja_file=self.marja_file,
                    vgo_file=self.vgo_file,
                    progress_callback=self._on_progress,
                )
            elif self.task_type == 2:
                result = unload_corr(
                    macros_file=self.macros_file, progress_callback=self._on_progress
                )
            else:
                result = ProcessingResult(
                    success=False, message="Неизвестный тип задачи"
                )

            self.finished_with_result.emit(result)

        except Exception as e:
            log.exception(f"Ошибка в рабочем потоке: {e}")
            result = ProcessingResult(success=False, errors=[str(e)])
            self.finished_with_result.emit(result)

    def _on_progress(self, percent: int, message: str):
        """Callback для обновления прогресса"""
        self.progress_updated.emit(percent, message)


class MainWindow(QMainWindow):
    """Главное окно приложения"""

    def __init__(self):
        super().__init__()

        self.worker: Optional[WorkerThread] = None

        self._setup_ui()
        self._connect_signals()

        log.info("Приложение запущено")

    def _setup_ui(self):
        """Настройка пользовательского интерфейса"""
        self.setWindowTitle("Обработка Excel файлов")
        self.setMinimumSize(600, 500)
        self.resize(700, 600)

        # Центральный виджет
        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        main_layout = QVBoxLayout(central_widget)
        main_layout.setContentsMargins(10, 10, 10, 10)
        main_layout.setSpacing(8)

        # Заголовок
        self._create_header(main_layout)

        # Прокручиваемая область для основного контента
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setFrameShape(QFrame.NoFrame)
        scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        
        scroll_content = QWidget()
        scroll_layout = QVBoxLayout(scroll_content)
        scroll_layout.setContentsMargins(5, 5, 5, 5)
        scroll_layout.setSpacing(8)

        # Группа выбора блока работ
        self._create_task_group(scroll_layout)

        # Группа выбора файлов
        self._create_files_group(scroll_layout)

        # Прогресс и статус
        self._create_action_group(scroll_layout)

        # Лог
        self._create_log_group(scroll_layout)

        scroll_area.setWidget(scroll_content)
        main_layout.addWidget(scroll_area, 1)  # Растягиваемый контент

        # Фиксированные кнопки внизу
        self._create_fixed_buttons(main_layout)

        # Применяем стили
        self.setStyleSheet(STYLESHEET)

    def _create_header(self, layout: QVBoxLayout):
        """Создание заголовка"""
        header_widget = QWidget()
        header_layout = QVBoxLayout(header_widget)
        header_layout.setContentsMargins(0, 0, 0, 5)
        header_layout.setAlignment(Qt.AlignCenter)

        title = QLabel("Обработка Excel файлов")
        title.setObjectName("title")
        title.setAlignment(Qt.AlignCenter)
        title.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Fixed)

        header_layout.addWidget(title)

        layout.addWidget(header_widget)

    def _create_task_group(self, layout: QVBoxLayout):
        """Создание группы выбора блока работ"""
        group = QGroupBox("Выберите блок работ")
        group_layout = QVBoxLayout(group)
        group_layout.setContentsMargins(10, 15, 10, 10)
        group_layout.setSpacing(8)

        self.task_combo = QComboBox()
        self.task_combo.addItem(
            "1. Создать предварительные листы с Расчетами и Мэпинги"
        )
        self.task_combo.addItem("2. Создать отчет по корректировке CF16")
        self.task_combo.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)

        group_layout.addWidget(self.task_combo)
        layout.addWidget(group)

    def _create_files_group(self, layout: QVBoxLayout):
        """Создание группы выбора файлов"""
        group = QGroupBox("Выбор файла")
        group_layout = QVBoxLayout(group)
        group_layout.setContentsMargins(10, 15, 10, 10)
        group_layout.setSpacing(8)

        # Основной файл макроса
        macros_layout = QHBoxLayout()
        macros_layout.setSpacing(8)
        self.macros_edit = QLineEdit()
        self.macros_edit.setPlaceholderText("Файл не выбран")
        self.macros_edit.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.macros_btn = QPushButton("Выбрать файл")
        self.macros_btn.setObjectName("browse")
        self.macros_btn.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
        macros_layout.addWidget(self.macros_edit, 1)
        macros_layout.addWidget(self.macros_btn)
        group_layout.addLayout(macros_layout)

        # Разделитель
        separator1 = QFrame()
        separator1.setObjectName("separator")
        separator1.setFrameShape(QFrame.HLine)
        separator1.setFixedHeight(1)
        group_layout.addWidget(separator1)

        # Файл Маржа
        marja_layout = QHBoxLayout()
        marja_layout.setSpacing(8)
        self.marja_label = QLabel("Файл с листом Маржа:")
        self.marja_label.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
        self.marja_edit = QLineEdit()
        self.marja_edit.setPlaceholderText("Выберите файл с листом Маржа...")
        self.marja_edit.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.marja_btn = QPushButton("Обзор...")
        self.marja_btn.setObjectName("browse")
        self.marja_btn.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
        marja_layout.addWidget(self.marja_label)
        marja_layout.addWidget(self.marja_edit, 1)
        marja_layout.addWidget(self.marja_btn)
        group_layout.addLayout(marja_layout)

        # Файл ВГО
        vgo_layout = QHBoxLayout()
        vgo_layout.setSpacing(8)
        self.vgo_label = QLabel("Файл отчёта по выверке ВГО:")
        self.vgo_label.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
        self.vgo_edit = QLineEdit()
        self.vgo_edit.setPlaceholderText("Выберите файл с отчётом ВГО...")
        self.vgo_edit.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.vgo_btn = QPushButton("Обзор...")
        self.vgo_btn.setObjectName("browse")
        self.vgo_btn.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
        vgo_layout.addWidget(self.vgo_label)
        vgo_layout.addWidget(self.vgo_edit, 1)
        vgo_layout.addWidget(self.vgo_btn)
        group_layout.addLayout(vgo_layout)

        layout.addWidget(group)

    def _create_action_group(self, layout: QVBoxLayout):
        """Создание группы действий"""
        # Группа прогресса
        progress_group = QGroupBox("Прогресс обработки")
        progress_layout = QVBoxLayout(progress_group)
        progress_layout.setContentsMargins(10, 15, 10, 10)
        progress_layout.setSpacing(8)

        # Прогресс-бар
        self.progress_bar = QProgressBar()
        self.progress_bar.setMinimum(0)
        self.progress_bar.setMaximum(100)
        self.progress_bar.setValue(0)
        self.progress_bar.setFormat("%p%")
        self.progress_bar.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.progress_bar.setFixedHeight(24)
        progress_layout.addWidget(self.progress_bar)

        layout.addWidget(progress_group)

        # Группа статуса
        status_group = QGroupBox("Статус")
        status_layout = QVBoxLayout(status_group)
        status_layout.setContentsMargins(10, 15, 10, 10)
        status_layout.setSpacing(8)

        # Статус
        self.status_label = QLabel("Готов к работе")
        self.status_label.setObjectName("status")
        self.status_label.setAlignment(Qt.AlignCenter)
        self.status_label.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.status_label.setMinimumHeight(30)
        status_layout.addWidget(self.status_label)

        layout.addWidget(status_group)

    def _create_log_group(self, layout: QVBoxLayout):
        """Создание группы логов"""
        group = QGroupBox("Лог обработки")
        group_layout = QVBoxLayout(group)
        group_layout.setContentsMargins(10, 15, 10, 10)
        group_layout.setSpacing(8)

        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setMinimumHeight(80)
        self.log_text.setMaximumHeight(150)
        self.log_text.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
        self.log_text.setPlaceholderText(
            "Здесь будут отображаться сообщения о ходе выполнения..."
        )

        group_layout.addWidget(self.log_text)

        # Кнопка очистки лога
        clear_btn = QPushButton("Очистить лог")
        clear_btn.setObjectName("browse")
        clear_btn.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        clear_btn.setFixedHeight(32)
        clear_btn.clicked.connect(self._clear_log)
        group_layout.addWidget(clear_btn)

        layout.addWidget(group)

    def _create_fixed_buttons(self, layout: QVBoxLayout):
        """Создание фиксированных кнопок внизу окна"""
        # Кнопка обработки файла
        button_layout = QHBoxLayout()
        button_layout.setContentsMargins(5, 5, 5, 5)
        button_layout.setSpacing(10)

        self.run_btn = QPushButton("Обработать файл")
        self.run_btn.setObjectName("primary")
        self.run_btn.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.run_btn.setFixedHeight(36)
        self.run_btn.setMinimumWidth(150)

        button_layout.addWidget(self.run_btn)

        layout.addLayout(button_layout)

    def _connect_signals(self):
        """Подключение сигналов"""
        self.task_combo.currentIndexChanged.connect(self._on_task_changed)
        self.macros_btn.clicked.connect(self._browse_macros_file)
        self.marja_btn.clicked.connect(self._browse_marja_file)
        self.vgo_btn.clicked.connect(self._browse_vgo_file)
        self.run_btn.clicked.connect(self._run_task)

        # Изначально обновляем состояние полей
        self._on_task_changed(0)

    def _on_task_changed(self, index: int):
        """Обработка изменения выбранного блока работ"""
        # Для блока 1 нужны файлы Маржа и ВГО
        # Для блока 2 - только основной файл
        is_block1 = index == 0

        self.marja_edit.setEnabled(is_block1)
        self.marja_btn.setEnabled(is_block1)
        self.marja_label.setEnabled(is_block1)

        self.vgo_edit.setEnabled(is_block1)
        self.vgo_btn.setEnabled(is_block1)
        self.vgo_label.setEnabled(is_block1)

        if not is_block1:
            self.marja_edit.clear()
            self.vgo_edit.clear()

        self._log(f"Выбран блок работ: {self.task_combo.currentText()}")

    def _browse_macros_file(self):
        """Выбор основного файла"""
        # Устанавливаем стили для диалога перед открытием
        dialog = QFileDialog(self, "Выберите основной файл с макросами")
        dialog.setFileMode(QFileDialog.ExistingFile)
        dialog.setNameFilter("Excel Files (*.xlsm *.xlsx);;All Files (*.*)")
        dialog.setStyleSheet("""
            QFileDialog {
                background-color: #ffffff;
                color: #333333;
            }
            QFileDialog QLabel {
                color: #333333;
            }
            QFileDialog QLineEdit {
                background-color: #ffffff;
                color: #333333;
                border: 1px solid #cccccc;
            }
            QFileDialog QPushButton {
                background-color: #4CAF50;
                color: #ffffff;
                border: none;
                border-radius: 5px;
                padding: 5px 15px;
            }
            QFileDialog QPushButton:hover {
                background-color: #45a049;
            }
            QFileDialog QTreeView, QFileDialog QListView {
                background-color: #ffffff;
                color: #333333;
                selection-background-color: #4CAF50;
                selection-color: #ffffff;
            }
            QFileDialog QHeaderView::section {
                background-color: #f0f0f0;
                color: #333333;
                padding: 5px;
                border: 1px solid #cccccc;
            }
        """)
        
        if dialog.exec():
            file_paths = dialog.selectedFiles()
            if file_paths:
                file_path = file_paths[0]
                self.macros_edit.setText(file_path)
                self._log(f"Выбран основной файл: {Path(file_path).name}")

    def _browse_marja_file(self):
        """Выбор файла Маржа"""
        dialog = QFileDialog(self, "Выберите файл с листом Маржа")
        dialog.setFileMode(QFileDialog.ExistingFile)
        dialog.setNameFilter("Excel Files (*.xlsx *.xlsm *.xls *.xlsb);;All Files (*.*)")
        dialog.setStyleSheet("""
            QFileDialog {
                background-color: #ffffff;
                color: #333333;
            }
            QFileDialog QLabel {
                color: #333333;
            }
            QFileDialog QLineEdit {
                background-color: #ffffff;
                color: #333333;
                border: 1px solid #cccccc;
            }
            QFileDialog QPushButton {
                background-color: #4CAF50;
                color: #ffffff;
                border: none;
                border-radius: 5px;
                padding: 5px 15px;
            }
            QFileDialog QPushButton:hover {
                background-color: #45a049;
            }
            QFileDialog QTreeView, QFileDialog QListView {
                background-color: #ffffff;
                color: #333333;
                selection-background-color: #4CAF50;
                selection-color: #ffffff;
            }
            QFileDialog QHeaderView::section {
                background-color: #f0f0f0;
                color: #333333;
                padding: 5px;
                border: 1px solid #cccccc;
            }
        """)
        
        if dialog.exec():
            file_paths = dialog.selectedFiles()
            if file_paths:
                file_path = file_paths[0]
                self.marja_edit.setText(file_path)
                self._log(f"Выбран файл Маржа: {Path(file_path).name}")

    def _browse_vgo_file(self):
        """Выбор файла ВГО"""
        dialog = QFileDialog(self, "Выберите файл с отчётом по выверке ВГО")
        dialog.setFileMode(QFileDialog.ExistingFile)
        dialog.setNameFilter("Excel Files (*.xlsx *.xlsm *.xls *.xlsb);;All Files (*.*)")
        dialog.setStyleSheet("""
            QFileDialog {
                background-color: #ffffff;
                color: #333333;
            }
            QFileDialog QLabel {
                color: #333333;
            }
            QFileDialog QLineEdit {
                background-color: #ffffff;
                color: #333333;
                border: 1px solid #cccccc;
            }
            QFileDialog QPushButton {
                background-color: #4CAF50;
                color: #ffffff;
                border: none;
                border-radius: 5px;
                padding: 5px 15px;
            }
            QFileDialog QPushButton:hover {
                background-color: #45a049;
            }
            QFileDialog QTreeView, QFileDialog QListView {
                background-color: #ffffff;
                color: #333333;
                selection-background-color: #4CAF50;
                selection-color: #ffffff;
            }
            QFileDialog QHeaderView::section {
                background-color: #f0f0f0;
                color: #333333;
                padding: 5px;
                border: 1px solid #cccccc;
            }
        """)
        
        if dialog.exec():
            file_paths = dialog.selectedFiles()
            if file_paths:
                file_path = file_paths[0]
                self.vgo_edit.setText(file_path)
                self._log(f"Выбран файл ВГО: {Path(file_path).name}")

    def _validate_inputs(self) -> bool:
        """Проверка введённых данных"""
        errors = []

        # Проверка основного файла
        macros_path = self.macros_edit.text().strip()
        if not macros_path:
            errors.append("Не указан основной файл")
        elif not Path(macros_path).exists():
            errors.append("Основной файл не найден")

        # Для блока 1 проверяем дополнительные файлы
        if self.task_combo.currentIndex() == 0:
            marja_path = self.marja_edit.text().strip()
            vgo_path = self.vgo_edit.text().strip()

            if not marja_path:
                errors.append("Не указан файл с листом Маржа")
            elif not Path(marja_path).exists():
                errors.append("Файл Маржа не найден")

            if not vgo_path:
                errors.append("Не указан файл по выверке ВГО")
            elif not Path(vgo_path).exists():
                errors.append("Файл ВГО не найден")

        if errors:
            QMessageBox.critical(
                self,
                "Ошибка валидации",
                "Обнаружены следующие ошибки:\n\n• " + "\n• ".join(errors),
            )
            return False

        return True

    def _run_task(self):
        """Запуск выбранной задачи"""
        if not self._validate_inputs():
            return

        # Блокируем интерфейс
        self._set_ui_enabled(False)

        # Определяем тип задачи
        task_type = self.task_combo.currentIndex() + 1

        # Создаём рабочий поток
        self.worker = WorkerThread(
            task_type=task_type,
            macros_file=self.macros_edit.text().strip(),
            marja_file=self.marja_edit.text().strip(),
            vgo_file=self.vgo_edit.text().strip(),
        )

        self.worker.progress_updated.connect(self._on_progress_updated)
        self.worker.finished_with_result.connect(self._on_task_finished)

        self._log("🚀 Запуск обработки...")
        self.progress_bar.setValue(0)

        self.worker.start()

    def _on_progress_updated(self, percent: int, message: str):
        """Обработка обновления прогресса"""
        self.progress_bar.setValue(percent)
        self.progress_bar.setFormat(f"%p%")
        self.status_label.setText(message)
        self._log(f"[{percent}%] {message}")

    def _on_task_finished(self, result: ProcessingResult):
        """Обработка завершения задачи"""
        self._set_ui_enabled(True)

        if result.success:
            self.progress_bar.setValue(100)
            self.progress_bar.setFormat("100%")
            self.status_label.setText(result.message)
            self._log(f"✅ {result.message}")

            QMessageBox.information(self, "Успех", result.message)
        else:
            self.progress_bar.setFormat("0%")
            error_msg = (
                "\n".join(result.errors) if result.errors else "Неизвестная ошибка"
            )
            self.status_label.setText("Ошибка выполнения")
            self._log(f"❌ Ошибка: {error_msg}")

            QMessageBox.critical(self, "Ошибка", f"Произошла ошибка:\n\n{error_msg}")

        if result.warnings:
            for warning in result.warnings:
                self._log(f"⚠️ {warning}")

    def _set_ui_enabled(self, enabled: bool):
        """Включение/отключение элементов интерфейса"""
        self.task_combo.setEnabled(enabled)
        self.macros_edit.setEnabled(enabled)
        self.macros_btn.setEnabled(enabled)

        is_block1 = self.task_combo.currentIndex() == 0
        self.marja_edit.setEnabled(enabled and is_block1)
        self.marja_btn.setEnabled(enabled and is_block1)
        self.vgo_edit.setEnabled(enabled and is_block1)
        self.vgo_btn.setEnabled(enabled and is_block1)

        self.run_btn.setEnabled(enabled)

    def _clear_log(self):
        """Очистка лога"""
        self.log_text.clear()

    def _log(self, message: str):
        """Добавление сообщения в лог"""
        self.log_text.append(message)
        # Прокручиваем вниз
        scrollbar = self.log_text.verticalScrollBar()
        scrollbar.setValue(scrollbar.maximum())

    def closeEvent(self, event):
        """Обработка закрытия окна"""
        if self.worker and self.worker.isRunning():
            reply = QMessageBox.question(
                self,
                "Подтверждение",
                "Выполняется обработка. Вы уверены, что хотите закрыть приложение?",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No,
            )

            if reply == QMessageBox.Yes:
                self.worker.terminate()
                self.worker.wait()
                event.accept()
            else:
                event.ignore()
        else:
            event.accept()


def run_app():
    """Запуск приложения"""
    app = QApplication(sys.argv)

    # Устанавливаем светлую тему для всей платформы
    app.setStyle("Fusion")
    
    # Устанавливаем светлую палитру для приложения
    palette = QPalette()
    palette.setColor(QPalette.Window, QColor(255, 255, 255))
    palette.setColor(QPalette.WindowText, QColor(51, 51, 51))
    palette.setColor(QPalette.Base, QColor(255, 255, 255))
    palette.setColor(QPalette.AlternateBase, QColor(245, 245, 245))
    palette.setColor(QPalette.ToolTipBase, QColor(255, 255, 255))
    palette.setColor(QPalette.ToolTipText, QColor(51, 51, 51))
    palette.setColor(QPalette.Text, QColor(51, 51, 51))
    palette.setColor(QPalette.Button, QColor(255, 255, 255))
    palette.setColor(QPalette.ButtonText, QColor(51, 51, 51))
    palette.setColor(QPalette.BrightText, QColor(255, 0, 0))
    palette.setColor(QPalette.Link, QColor(0, 122, 204))
    palette.setColor(QPalette.Highlight, QColor(76, 175, 80))
    palette.setColor(QPalette.HighlightedText, QColor(255, 255, 255))
    app.setPalette(palette)

    window = MainWindow()
    window.show()

    sys.exit(app.exec())


if __name__ == "__main__":
    run_app()
