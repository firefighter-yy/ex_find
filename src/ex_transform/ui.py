"""Minimal Qt workflow for preparing workbooks and previewing search rows."""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from .exporter import ExportCancelledError, ExportError, ExportProgress, ExportService
from .hardening import TaskCoordinator, classify_error
from .search_engine import MatchMode, SearchOperator, SearchPage, SearchService
from .session_index import FilePreparationService, SessionIndex


@dataclass
class SearchViewModel:
    """Qt-independent state and validation for the main window."""

    files: list[Path] = field(default_factory=list)
    keywords: list[str] = field(default_factory=list)
    operator: SearchOperator = SearchOperator.AND
    match_mode: MatchMode = MatchMode.CONTAINS
    preparing: bool = False
    searching: bool = False
    stale: bool = True
    result_count: int = 0
    error: str | None = None
    page: SearchPage = field(default_factory=lambda: SearchPage((), 0, 0, None))

    @property
    def can_search(self) -> bool:
        return bool(self.files and self.keywords and not self.preparing and not self.searching and not self.stale)

    @property
    def can_export(self) -> bool:
        return self.result_count > 0 and not self.preparing and not self.searching and not self.stale

    def add_keywords(self, text: str) -> bool:
        value = text.strip()
        if not value or value.casefold() in {item.casefold() for item in self.keywords}:
            return False
        self.keywords.append(value)
        return True

    def remove_keyword(self, value: str) -> None:
        self.keywords = [item for item in self.keywords if item != value]

    def clear_keywords(self) -> None:
        self.keywords.clear()


try:
    from PySide6.QtCore import QAbstractTableModel, QModelIndex, QObject, QTimer, Qt
    from PySide6.QtWidgets import (
        QFileDialog, QHBoxLayout, QLabel, QLineEdit, QListWidget, QListWidgetItem,
        QMainWindow, QPushButton, QComboBox, QSplitter, QTableView, QVBoxLayout, QWidget,
    )
except ImportError:  # pragma: no cover - exercised only on machines without Qt
    QT_AVAILABLE = False
else:
    QT_AVAILABLE = True


if QT_AVAILABLE:
    class ResultTableModel(QAbstractTableModel):
        headers = ("来源文件", "工作表", "原始行号", "命中关键词", "行值")

        def __init__(self, parent: QObject | None = None) -> None:
            super().__init__(parent)
            self.page = SearchPage((), 0, 0, None)

        def set_page(self, page: SearchPage) -> None:
            self.beginResetModel()
            self.page = page
            self.endResetModel()

        def rowCount(self, parent: QModelIndex = QModelIndex()) -> int:
            return 0 if parent.isValid() else len(self.page.results)

        def columnCount(self, parent: QModelIndex = QModelIndex()) -> int:
            return 0 if parent.isValid() else len(self.headers)

        def headerData(self, section: int, orientation: Qt.Orientation, role: int = Qt.ItemDataRole.DisplayRole) -> Any:
            if role != Qt.ItemDataRole.DisplayRole:
                return None
            return self.headers[section] if orientation == Qt.Orientation.Horizontal else section + 1

        def data(self, index: QModelIndex, role: int = Qt.ItemDataRole.DisplayRole) -> Any:
            if not index.isValid() or role != Qt.ItemDataRole.DisplayRole:
                return None
            result = self.page.results[index.row()]
            values = (
                str(result.file_path), result.worksheet_name, result.row,
                ", ".join(result.matched_keywords),
                " | ".join(f"C{cell.column}: {cell.value}" for cell in result.cells),
            )
            return values[index.column()]


    class MainWindow(QMainWindow):
        """Single-window file preparation, search, and result preview workflow."""

        def __init__(self, index: SessionIndex, preparation: FilePreparationService | None = None, parent: QWidget | None = None) -> None:
            super().__init__(parent)
            self.index = index
            self.preparation = preparation or FilePreparationService(index)
            self.search_service = SearchService(index)
            self.export_service = ExportService(index)
            self.view_model = SearchViewModel()
            self._tasks = TaskCoordinator(thread_name_prefix="excel-search-ui")
            self._future = None
            self._build_ui()
            self._refresh_files()

        def _build_ui(self) -> None:
            self.setWindowTitle("Excel Information Search")
            self.resize(1100, 680)
            root = QWidget(self)
            layout = QVBoxLayout(root)
            toolbar = QHBoxLayout()
            self.select_button = QPushButton("选择 Excel 文件")
            self.prepare_button = QPushButton("准备文件")
            self.cancel_button = QPushButton("取消")
            self.status_label = QLabel("请选择 Excel 文件")
            toolbar.addWidget(self.select_button)
            toolbar.addWidget(self.prepare_button)
            toolbar.addWidget(self.cancel_button)
            toolbar.addStretch()
            toolbar.addWidget(self.status_label)
            layout.addLayout(toolbar)
            self.file_list = QListWidget()
            self.file_list.setMinimumHeight(90)
            layout.addWidget(self.file_list)

            search_row = QHBoxLayout()
            search_row.addWidget(QLabel("关键词："))
            self.keyword_entry = QLineEdit()
            self.keyword_entry.setPlaceholderText("输入关键词后按 Enter")
            self.keyword_entry.returnPressed.connect(self._add_keyword)
            self.operator = QComboBox()
            self.operator.addItem("AND（全部匹配）", SearchOperator.AND)
            self.operator.addItem("OR（任一匹配）", SearchOperator.OR)
            self.match_mode = QComboBox()
            self.match_mode.addItem("包含匹配", MatchMode.CONTAINS)
            self.match_mode.addItem("精确匹配", MatchMode.EXACT)
            self.search_button = QPushButton("搜索")
            self.export_button = QPushButton("导出")
            for widget in (self.keyword_entry, self.operator, self.match_mode, self.search_button, self.export_button):
                search_row.addWidget(widget)
            layout.addLayout(search_row)
            self.tags = QListWidget()
            self.tags.setMaximumHeight(54)
            layout.addWidget(self.tags)

            splitter = QSplitter(Qt.Orientation.Vertical)
            self.results = ResultTableModel(self)
            self.result_view = QTableView()
            self.result_view.setModel(self.results)
            self.result_view.setWordWrap(False)
            self.result_view.setHorizontalScrollMode(QTableView.ScrollMode.ScrollPerPixel)
            self.result_view.horizontalHeader().setStretchLastSection(True)
            splitter.addWidget(self.result_view)
            layout.addWidget(splitter, 1)
            self.setCentralWidget(root)
            self.select_button.clicked.connect(self._choose_files)
            self.prepare_button.clicked.connect(self._start_prepare)
            self.cancel_button.clicked.connect(self._cancel)
            self.search_button.clicked.connect(self._start_search)
            self.export_button.clicked.connect(self._start_export)
            self._update_actions()

        def _choose_files(self) -> None:
            paths, _ = QFileDialog.getOpenFileNames(self, "选择 Excel 文件", "", "Excel 文件 (*.xlsx *.xlsm *.xls)")
            if paths:
                self.view_model.files = list(dict.fromkeys(Path(path).resolve() for path in paths))
                self.view_model.stale = True
                self._refresh_files()

        def _remove_selected_file(self) -> None:
            item = self.file_list.currentItem()
            if item is None or self.view_model.preparing:
                return
            path = Path(item.text()).resolve()
            self.view_model.files = [candidate for candidate in self.view_model.files if candidate != path]
            self.view_model.stale = True
            self._refresh_files()

        def _refresh_files(self) -> None:
            self.file_list.clear()
            for path in self.view_model.files:
                self.file_list.addItem(QListWidgetItem(str(path)))
            self.prepare_button.setEnabled(bool(self.view_model.files) and not self.view_model.preparing)
            self._update_actions()

        def _add_keyword(self) -> None:
            if self.view_model.add_keywords(self.keyword_entry.text()):
                self.keyword_entry.clear()
                self.tags.addItem(QListWidgetItem(self.view_model.keywords[-1] + "  ×"))
                self._update_actions()

        def _remove_tag(self, item) -> None:
            value = str(item.data(Qt.ItemDataRole.UserRole) or item.text().removesuffix("  ×"))
            self.view_model.remove_keyword(value)
            self.tags.takeItem(self.tags.row(item))
            self._update_actions()

        def _clear_keywords(self) -> None:
            self.view_model.clear_keywords()
            self.tags.clear()
            self._update_actions()

        def _update_actions(self) -> None:
            self.search_button.setEnabled(self.view_model.can_search)
            self.export_button.setEnabled(self.view_model.can_export)
            self.cancel_button.setEnabled(self.view_model.preparing or self.view_model.searching)

        def _start_prepare(self) -> None:
            self.view_model.preparing = True
            self.status_label.setText("正在准备文件…")
            self._update_actions()
            self._future = self._tasks.submit(
                self.preparation.prepare,
                self.view_model.files,
                cancel_event=self._tasks.cancel_event,
            )
            QTimer.singleShot(100, self._poll_prepare)

        def _poll_prepare(self) -> None:
            if self._future is None or not self._future.done():
                QTimer.singleShot(100, self._poll_prepare)
                return
            self.view_model.preparing = False
            try:
                result = self._future.result()
                self.view_model.stale = not bool(result.succeeded)
                self.status_label.setText("文件准备已完成" if not result.failed else "部分文件准备失败，请修复后重试")
            except Exception as exc:
                self.status_label.setText(classify_error(exc).message)
            self._refresh_files()

        def _start_search(self) -> None:
            if self.index.mark_stale():
                self.view_model.stale = True
                self.status_label.setText("源文件已变更，请重新准备文件")
                self._update_actions()
                return
            self.view_model.searching = True
            self.view_model.result_count = 0
            self.status_label.setText("正在搜索…")
            self._update_actions()
            self._future = self._tasks.submit(
                self.search_service.search, self.view_model.keywords,
                operator=self.operator.currentData(), match=self.match_mode.currentData(), limit=500,
                cancel_event=self._tasks.cancel_event,
            )
            QTimer.singleShot(100, self._poll_search)

        def _poll_search(self) -> None:
            if self._future is None or not self._future.done():
                QTimer.singleShot(100, self._poll_search)
                return
            self.view_model.searching = False
            try:
                page = self._future.result()
                self.results.set_page(page)
                self.view_model.page = page
                self.view_model.result_count = page.total_count
                self.status_label.setText("搜索已取消" if page.cancelled else (f"找到 {page.total_count} 条结果" if page.total_count else "没有找到结果"))
            except Exception as exc:
                self.status_label.setText(classify_error(exc).message)
            self._update_actions()

        def _start_export(self) -> None:
            if not self.view_model.can_export:
                return
            target, _ = QFileDialog.getSaveFileName(self, "导出搜索结果", "搜索结果.xlsx", "Excel 文件 (*.xlsx *.xlsm *.xls)")
            if not target:
                return
            self.view_model.searching = True
            self.status_label.setText("正在导出…")
            self._update_actions()
            self._future = self._tasks.submit(
                self.export_service.export, self.export_service.snapshot(self.view_model.page), target,
                cancel_event=self._tasks.cancel_event, progress_callback=self._export_progress,
            )
            QTimer.singleShot(100, self._poll_export)

        def _export_progress(self, progress: ExportProgress) -> None:
            QTimer.singleShot(0, lambda: self.status_label.setText(f"正在导出… {progress.completed}/{progress.total}"))

        def _poll_export(self) -> None:
            if self._future is None or not self._future.done():
                QTimer.singleShot(100, self._poll_export)
                return
            self.view_model.searching = False
            try:
                result = self._future.result()
                warning = f"（{'; '.join(result.warnings)}）" if result.warnings else ""
                self.status_label.setText(f"已导出 {result.result_count} 条结果：{result.path} {warning}")
            except ExportCancelledError:
                self.status_label.setText("导出已取消")
            except ExportError as exc:
                self.status_label.setText(classify_error(exc).message)
            except Exception as exc:
                self.status_label.setText(classify_error(exc).message)
            self._update_actions()
        def _cancel(self) -> None:
            self._tasks.cancel()
            self.status_label.setText("正在取消…")

        def closeEvent(self, event) -> None:
            self._tasks.close(wait=True)
            self.index.close()
            event.accept()
else:
    class MainWindow:  # pragma: no cover
        def __init__(self, *args, **kwargs):
            raise RuntimeError("PySide6 is required to start the desktop application")








