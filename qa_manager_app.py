from __future__ import annotations

import re
import sys
from pathlib import Path

from PySide6.QtCore import Qt, QSize, Signal
from PySide6.QtGui import QAction, QCloseEvent, QFont, QIcon
from PySide6.QtWidgets import (
    QApplication,
    QAbstractItemView,
    QComboBox,
    QDialog,
    QDialogButtonBox,
    QFileDialog,
    QFrame,
    QGridLayout,
    QHBoxLayout,
    QHeaderView,
    QInputDialog,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPlainTextEdit,
    QPushButton,
    QSizePolicy,
    QSplitter,
    QStyle,
    QTableWidget,
    QTableWidgetItem,
    QTabWidget,
    QTextEdit,
    QTreeWidget,
    QTreeWidgetItem,
    QVBoxLayout,
    QWidget,
)
from rapidfuzz import fuzz

from qa_excel import QARecord, QAWorkbook, TemplateItem


APP_TITLE = "QA 问答库数据管理工具"
CURRENT_VALUE_ROLE = Qt.UserRole + 1
SOURCE_INTENT_ROLE = Qt.UserRole + 2
SOURCE_QUESTION_ROLE = Qt.UserRole + 3


def display_text(value: str) -> str:
    value = value.strip()
    value = re.sub(r"^\d+(?:\.\d+)*[\.、]?\s*", "", value)
    return value.strip()


def fuzzy_filter(values: list[str], query: str) -> list[str]:
    query = query.strip()
    if not query:
        return values
    scored = []
    query_lower = query.lower()
    for idx, value in enumerate(values):
        value_lower = display_text(value).lower()
        score = 100 if query_lower in value_lower else fuzz.partial_ratio(query_lower, value_lower)
        if score >= 45:
            scored.append((score, idx, value))
    scored.sort(key=lambda item: (-item[0], item[1]))
    return [value for _, _, value in scored]


class SearchableComboBox(QComboBox):
    valueSelected = Signal(str)

    def __init__(self, placeholder: str = "", parent=None):
        super().__init__(parent)
        self._all_values: list[str] = []
        self._updating = False
        self.setEditable(True)
        self.setInsertPolicy(QComboBox.NoInsert)
        self.setMaxVisibleItems(15)
        self.setMinimumContentsLength(24)
        self.lineEdit().setPlaceholderText(placeholder)
        self.lineEdit().textEdited.connect(self._filter_values)
        self.lineEdit().returnPressed.connect(self.accept_current_match)
        self.activated.connect(self._emit_current_value)

    def set_values(self, values: list[str], current: str = "") -> None:
        unique_values = list(dict.fromkeys(value for value in values if value))
        self._all_values = unique_values
        self._replace_items(unique_values)
        if current and current in unique_values:
            self.setCurrentText(current)
        elif unique_values:
            self.setCurrentIndex(0)
        else:
            self.setEditText("")

    def current_value(self) -> str:
        index = self.currentIndex()
        if index >= 0:
            raw = self.itemData(index, Qt.UserRole)
            if raw:
                return str(raw)
        text = self.currentText().strip()
        for value in self._all_values:
            if display_text(value) == text or value == text:
                return value
        return ""

    def select_value(self, value: str) -> None:
        self._replace_items(self._all_values)
        index = self._find_raw_value(value)
        if index >= 0:
            self.setCurrentIndex(index)
        elif value:
            self.setEditText(display_text(value))

    def accept_current_match(self) -> None:
        if self.count() == 0:
            return
        index = self.currentIndex() if self.currentIndex() >= 0 else 0
        value = self.itemText(index)
        self.setCurrentText(value)
        self.valueSelected.emit(value)

    def _filter_values(self, query: str) -> None:
        matches = fuzzy_filter(self._all_values, query)
        self._replace_items(matches, query)
        if matches:
            self.showPopup()

    def _replace_items(self, values: list[str], edit_text: str | None = None) -> None:
        self._updating = True
        self.blockSignals(True)
        self.clear()
        for value in values:
            self.addItem(display_text(value), value)
        if edit_text is not None:
            self.setEditText(edit_text)
        self.blockSignals(False)
        self._updating = False

    def _emit_current_value(self, *args) -> None:
        if self._updating:
            return
        value = self.current_value()
        if value:
            self.valueSelected.emit(value)

    def _find_raw_value(self, value: str) -> int:
        for index in range(self.count()):
            if self.itemData(index, Qt.UserRole) == value:
                return index
        return -1


class TemplateQuestionDialog(QDialog):
    def __init__(
        self,
        intents: list[str],
        title: str,
        intent: str = "",
        question: str = "",
        parent=None,
    ):
        super().__init__(parent)
        self.setWindowTitle(title)
        self.setMinimumWidth(560)

        layout = QVBoxLayout(self)
        layout.setSpacing(14)

        form = QGridLayout()
        form.setHorizontalSpacing(12)
        form.setVerticalSpacing(10)

        self.intent_combo = QComboBox()
        self.intent_combo.setEditable(True)
        self.intent_combo.setInsertPolicy(QComboBox.NoInsert)
        self.intent_combo.setMaxVisibleItems(15)
        self.intent_combo.addItems(intents)
        if intent:
            index = self.intent_combo.findText(intent)
            if index >= 0:
                self.intent_combo.setCurrentIndex(index)
            else:
                self.intent_combo.setEditText(intent)

        self.question_edit = QPlainTextEdit()
        self.question_edit.setPlaceholderText("输入具体问法")
        self.question_edit.setPlainText(question)
        self.question_edit.setMinimumHeight(110)

        form.addWidget(QLabel("咨询意图"), 0, 0)
        form.addWidget(self.intent_combo, 0, 1)
        form.addWidget(QLabel("具体问法"), 1, 0, Qt.AlignTop)
        form.addWidget(self.question_edit, 1, 1)
        layout.addLayout(form)

        hint = QLabel("可以选择已有咨询意图，也可以直接输入新的意图名称。")
        hint.setObjectName("mutedLabel")
        layout.addWidget(hint)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def values(self) -> tuple[str, str]:
        return self.intent_combo.currentText().strip(), self.question_edit.toPlainText().strip()

    def accept(self) -> None:
        intent, question = self.values()
        if not intent or not question:
            QMessageBox.warning(self, "内容不完整", "咨询意图和具体问法都不能为空。")
            return
        super().accept()


class AddQAPairDialog(QDialog):
    def __init__(self, intent: str, parent=None):
        super().__init__(parent)
        self.intent = intent
        self.setWindowTitle("新增 POI 专属问答")
        self.setMinimumWidth(640)
        self.setMinimumHeight(420)

        layout = QVBoxLayout(self)
        layout.setSpacing(14)

        intent_label = QLabel(f"咨询意图：{display_text(intent)}")
        intent_label.setObjectName("sectionTitle")
        layout.addWidget(intent_label)

        self.question_edit = QPlainTextEdit()
        self.question_edit.setPlaceholderText("输入这个 POI 专属的具体问法")
        self.question_edit.setMinimumHeight(100)
        layout.addWidget(QLabel("具体问法"))
        layout.addWidget(self.question_edit)

        self.answer_edit = QTextEdit()
        self.answer_edit.setPlaceholderText("输入回复内容，可以暂时留空")
        layout.addWidget(QLabel("回复内容"))
        layout.addWidget(self.answer_edit, 1)

        hint = QLabel("这条问答只会加入当前 POI，不会进入标准问题模板，也不会自动应用到其他 POI。")
        hint.setObjectName("mutedLabel")
        hint.setWordWrap(True)
        layout.addWidget(hint)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def values(self) -> tuple[str, str]:
        return self.question_edit.toPlainText().strip(), self.answer_edit.toPlainText().strip()

    def accept(self) -> None:
        question, _ = self.values()
        if not question:
            QMessageBox.warning(self, "内容不完整", "具体问法不能为空。")
            return
        super().accept()


class BulkAnswerDialog(QDialog):
    def __init__(
        self,
        category: str,
        intent: str,
        question: str,
        current_answer: str,
        record_count: int,
        poi_count: int,
        parent=None,
    ):
        super().__init__(parent)
        self.setWindowTitle("批量修改回复")
        self.setMinimumWidth(720)
        self.setMinimumHeight(460)

        layout = QVBoxLayout(self)
        layout.setSpacing(14)

        title = QLabel("批量覆盖同款问题的回复")
        title.setObjectName("contentTitle")
        layout.addWidget(title)

        scope = QLabel(
            f"范围：分类“{category}”下，咨询意图“{display_text(intent)}”里的同款问题；"
            f"将影响 {poi_count} 个 POI、{record_count} 条问答。"
        )
        scope.setWordWrap(True)
        scope.setObjectName("mutedLabel")
        layout.addWidget(scope)

        question_label = QLabel("具体问法")
        question_label.setObjectName("sectionTitle")
        layout.addWidget(question_label)
        question_preview = QPlainTextEdit()
        question_preview.setPlainText(display_text(question))
        question_preview.setReadOnly(True)
        question_preview.setMaximumHeight(82)
        layout.addWidget(question_preview)

        answer_label = QLabel("统一回复内容")
        answer_label.setObjectName("sectionTitle")
        layout.addWidget(answer_label)
        self.answer_edit = QTextEdit()
        self.answer_edit.setPlainText(current_answer)
        self.answer_edit.setPlaceholderText("输入要统一覆盖到所有匹配 POI 的回复内容")
        layout.addWidget(self.answer_edit, 1)

        warning = QLabel("确认后会直接写入当前 Excel 文件，并覆盖匹配问答原有回复。")
        warning.setObjectName("mutedLabel")
        warning.setWordWrap(True)
        layout.addWidget(warning)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def answer(self) -> str:
        return self.answer_edit.toPlainText()


class AddPoiDialog(QDialog):
    def __init__(self, store: QAWorkbook, parent=None):
        super().__init__(parent)
        self.store = store
        self.setWindowTitle("新增 POI")
        self.setMinimumWidth(420)

        layout = QVBoxLayout(self)
        form = QGridLayout()
        form.setHorizontalSpacing(12)
        form.setVerticalSpacing(12)

        self.category_combo = QComboBox()
        self.category_combo.addItems(store.categories())
        self.region_combo = QComboBox()
        self.region_combo.setEditable(True)
        self.poi_input = QLineEdit()
        self.poi_input.setPlaceholderText("输入新的 POI 名称")

        form.addWidget(QLabel("分类"), 0, 0)
        form.addWidget(self.category_combo, 0, 1)
        form.addWidget(QLabel("区域"), 1, 0)
        form.addWidget(self.region_combo, 1, 1)
        form.addWidget(QLabel("POI"), 2, 0)
        form.addWidget(self.poi_input, 2, 1)
        layout.addLayout(form)

        hint = QLabel("新增后会自动套用该分类当前的问题模板，所有回复内容留空。")
        hint.setObjectName("mutedLabel")
        hint.setWordWrap(True)
        layout.addWidget(hint)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

        self.category_combo.currentTextChanged.connect(self._load_regions)
        self._load_regions(self.category_combo.currentText())

    def _load_regions(self, category: str) -> None:
        current = self.region_combo.currentText().strip()
        self.region_combo.clear()
        self.region_combo.addItems(self.store.regions(category))
        if current:
            self.region_combo.setEditText(current)

    def values(self) -> tuple[str, str, str]:
        return (
            self.category_combo.currentText().strip(),
            self.region_combo.currentText().strip(),
            self.poi_input.text().strip(),
        )


class MainWindow(QMainWindow):
    def __init__(self, initial_file: str | None = None):
        super().__init__()
        self.setWindowTitle(APP_TITLE)
        self.resize(1440, 900)
        self.store = QAWorkbook()
        self.current_category = ""
        self.current_region = ""
        self.current_poi = ""
        self.current_intent = ""
        self.current_records: list[QARecord] = []
        self.selected_record: QARecord | None = None
        self.dirty = False
        self._loading_ui = False

        self._build_ui()
        self._apply_style()

        if initial_file:
            self.load_file(initial_file)

    def _build_ui(self) -> None:
        root = QWidget()
        root_layout = QVBoxLayout(root)
        root_layout.setContentsMargins(18, 18, 18, 18)
        root_layout.setSpacing(14)
        self.setCentralWidget(root)

        header = QFrame()
        header.setObjectName("topBar")
        header_layout = QHBoxLayout(header)
        header_layout.setContentsMargins(16, 14, 16, 14)
        header_layout.setSpacing(12)

        title_box = QVBoxLayout()
        title = QLabel(APP_TITLE)
        title.setObjectName("appTitle")
        self.file_label = QLabel("请选择一个 Excel 问答库文件")
        self.file_label.setObjectName("mutedLabel")
        self.file_label.setTextInteractionFlags(Qt.TextSelectableByMouse)
        title_box.addWidget(title)
        title_box.addWidget(self.file_label)
        header_layout.addLayout(title_box, 1)

        self.status_label = QLabel("未加载")
        self.status_label.setObjectName("statusNeutral")
        header_layout.addWidget(self.status_label)

        self.open_button = QPushButton("选择 Excel")
        self.open_button.setIcon(self.style().standardIcon(QStyle.SP_DialogOpenButton))
        self.open_button.clicked.connect(self.choose_file)
        header_layout.addWidget(self.open_button)

        self.standardize_button = QPushButton("标准化 A-D 列")
        self.standardize_button.setIcon(self.style().standardIcon(QStyle.SP_BrowserReload))
        self.standardize_button.setEnabled(False)
        self.standardize_button.clicked.connect(self.standardize_current_file)
        header_layout.addWidget(self.standardize_button)

        self.save_button = QPushButton("保存修改")
        self.save_button.setIcon(self.style().standardIcon(QStyle.SP_DialogSaveButton))
        self.save_button.setEnabled(False)
        self.save_button.clicked.connect(self.save_changes)
        header_layout.addWidget(self.save_button)

        root_layout.addWidget(header)

        self.tabs = QTabWidget()
        self.tabs.addTab(self._build_data_tab(), "数据管理")
        self.tabs.addTab(self._build_template_tab(), "问题模板")
        root_layout.addWidget(self.tabs, 1)

    def _build_data_tab(self) -> QWidget:
        tab = QWidget()
        layout = QHBoxLayout(tab)
        layout.setContentsMargins(0, 0, 0, 0)

        splitter = QSplitter(Qt.Horizontal)
        splitter.setChildrenCollapsible(False)
        layout.addWidget(splitter)

        nav = QFrame()
        nav.setObjectName("panel")
        nav.setMinimumWidth(360)
        nav_layout = QVBoxLayout(nav)
        nav_layout.setContentsMargins(14, 14, 14, 14)
        nav_layout.setSpacing(10)

        nav_title = QLabel("数据导航")
        nav_title.setObjectName("sectionTitle")
        nav_layout.addWidget(nav_title)

        self.category_combo = QComboBox()
        self.category_combo.currentTextChanged.connect(self.on_category_changed)
        nav_layout.addWidget(QLabel("分类"))
        nav_layout.addWidget(self.category_combo)

        self.region_combo = SearchableComboBox("搜索或选择区域")
        self.region_combo.valueSelected.connect(self.on_region_changed)
        nav_layout.addWidget(QLabel("区域"))
        nav_layout.addWidget(self.region_combo)

        self.poi_combo = SearchableComboBox("搜索或选择 POI")
        self.poi_combo.valueSelected.connect(self.on_poi_changed)
        nav_layout.addWidget(QLabel("POI"))
        nav_layout.addWidget(self.poi_combo)

        poi_actions = QHBoxLayout()
        self.add_poi_button = QPushButton("新增")
        self.add_poi_button.setIcon(self.style().standardIcon(QStyle.SP_FileDialogNewFolder))
        self.add_poi_button.clicked.connect(self.add_poi)
        self.rename_poi_button = QPushButton("重命名")
        self.rename_poi_button.clicked.connect(self.rename_poi)
        self.delete_poi_button = QPushButton("删除")
        self.delete_poi_button.setObjectName("dangerButton")
        self.delete_poi_button.setIcon(self.style().standardIcon(QStyle.SP_TrashIcon))
        self.delete_poi_button.clicked.connect(self.delete_poi)
        poi_actions.addWidget(self.add_poi_button)
        poi_actions.addWidget(self.rename_poi_button)
        poi_actions.addWidget(self.delete_poi_button)
        nav_layout.addLayout(poi_actions)

        self.intent_combo = SearchableComboBox("搜索或选择咨询意图")
        self.intent_combo.valueSelected.connect(self.on_intent_changed)
        nav_layout.addWidget(QLabel("咨询意图"))
        nav_layout.addWidget(self.intent_combo)

        tree_label = QLabel("树结构总览")
        tree_label.setObjectName("sectionTitle")
        nav_layout.addWidget(tree_label)
        self.tree = QTreeWidget()
        self.tree.setHeaderHidden(True)
        self.tree.itemClicked.connect(self.on_tree_item_clicked)
        nav_layout.addWidget(self.tree, 1)

        splitter.addWidget(nav)

        content = QFrame()
        content.setObjectName("panel")
        content_layout = QVBoxLayout(content)
        content_layout.setContentsMargins(16, 16, 16, 16)
        content_layout.setSpacing(12)

        qa_header = QHBoxLayout()
        self.path_title = QLabel("请选择一个咨询意图")
        self.path_title.setObjectName("contentTitle")
        self.qa_count_label = QLabel("")
        self.qa_count_label.setObjectName("mutedLabel")
        self.add_qa_button = QPushButton("新增问答")
        self.add_qa_button.setIcon(self.style().standardIcon(QStyle.SP_FileDialogNewFolder))
        self.add_qa_button.clicked.connect(self.add_qa_pair)
        self.bulk_answer_button = QPushButton("批量修改回复")
        self.bulk_answer_button.setIcon(self.style().standardIcon(QStyle.SP_DialogApplyButton))
        self.bulk_answer_button.clicked.connect(self.bulk_update_answer)
        qa_header.addWidget(self.path_title, 1)
        qa_header.addWidget(self.qa_count_label)
        qa_header.addWidget(self.add_qa_button)
        qa_header.addWidget(self.bulk_answer_button)
        content_layout.addLayout(qa_header)

        self.qa_table = QTableWidget(0, 2)
        self.qa_table.setHorizontalHeaderLabels(["具体问法", "回复内容"])
        self.qa_table.verticalHeader().setVisible(False)
        self.qa_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.qa_table.setSelectionMode(QAbstractItemView.SingleSelection)
        self.qa_table.setWordWrap(True)
        self.qa_table.itemSelectionChanged.connect(self.on_qa_selected)
        self.qa_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        self.qa_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        content_layout.addWidget(self.qa_table, 3)

        editor = QFrame()
        editor.setObjectName("editorBox")
        editor_layout = QVBoxLayout(editor)
        editor_layout.setContentsMargins(14, 14, 14, 14)
        editor_layout.setSpacing(8)
        editor_title = QLabel("选中问答详情")
        editor_title.setObjectName("sectionTitle")
        editor_layout.addWidget(editor_title)

        self.question_preview = QPlainTextEdit()
        self.question_preview.setReadOnly(True)
        self.question_preview.setMaximumHeight(72)
        self.question_preview.setPlaceholderText("选择上方问答后显示问法")
        editor_layout.addWidget(self.question_preview)

        self.answer_editor = QTextEdit()
        self.answer_editor.setPlaceholderText("在这里编辑回复内容")
        self.answer_editor.textChanged.connect(self.on_answer_changed)
        editor_layout.addWidget(self.answer_editor, 1)
        content_layout.addWidget(editor, 2)

        splitter.addWidget(content)
        splitter.setStretchFactor(0, 0)
        splitter.setStretchFactor(1, 1)
        splitter.setSizes([420, 980])

        self._set_data_actions_enabled(False)
        return tab

    def _build_template_tab(self) -> QWidget:
        tab = QWidget()
        layout = QVBoxLayout(tab)
        layout.setContentsMargins(0, 0, 0, 0)

        panel = QFrame()
        panel.setObjectName("panel")
        panel_layout = QVBoxLayout(panel)
        panel_layout.setContentsMargins(16, 16, 16, 16)
        panel_layout.setSpacing(12)
        layout.addWidget(panel, 1)

        top = QHBoxLayout()
        title = QLabel("分类问题模板")
        title.setObjectName("contentTitle")
        top.addWidget(title)
        top.addStretch()
        top.addWidget(QLabel("分类"))
        self.template_category_combo = QComboBox()
        self.template_category_combo.currentTextChanged.connect(self.load_template_table)
        top.addWidget(self.template_category_combo)
        panel_layout.addLayout(top)

        desc = QLabel("模板改动会同步到该分类下所有 POI；新增问题的回复内容会留空。")
        desc.setObjectName("mutedLabel")
        panel_layout.addWidget(desc)

        self.template_table = QTableWidget(0, 2)
        self.template_table.setHorizontalHeaderLabels(["咨询意图", "具体问法"])
        self.template_table.verticalHeader().setVisible(False)
        self.template_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.template_table.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.template_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.template_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.template_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        panel_layout.addWidget(self.template_table, 1)

        actions = QHBoxLayout()
        self.template_add_button = QPushButton("新增问题")
        self.template_add_button.clicked.connect(self.add_template_question)
        self.template_edit_button = QPushButton("编辑选中")
        self.template_edit_button.clicked.connect(self.edit_template_question)
        self.template_delete_button = QPushButton("删除选中")
        self.template_delete_button.setObjectName("dangerButton")
        self.template_delete_button.clicked.connect(self.delete_template_rows)
        self.template_up_button = QPushButton("上移")
        self.template_up_button.clicked.connect(lambda: self.move_template_row(-1))
        self.template_down_button = QPushButton("下移")
        self.template_down_button.clicked.connect(lambda: self.move_template_row(1))
        self.template_apply_button = QPushButton("应用模板到该分类")
        self.template_apply_button.setIcon(self.style().standardIcon(QStyle.SP_DialogApplyButton))
        self.template_apply_button.clicked.connect(self.apply_template)

        actions.addWidget(self.template_add_button)
        actions.addWidget(self.template_edit_button)
        actions.addWidget(self.template_delete_button)
        actions.addWidget(self.template_up_button)
        actions.addWidget(self.template_down_button)
        actions.addStretch()
        actions.addWidget(self.template_apply_button)
        panel_layout.addLayout(actions)

        self._set_template_actions_enabled(False)
        return tab

    def _search_box(self, placeholder: str) -> QLineEdit:
        box = QLineEdit()
        box.setPlaceholderText(placeholder)
        box.setClearButtonEnabled(True)
        return box

    def choose_file(self) -> None:
        path, _ = QFileDialog.getOpenFileName(
            self,
            "选择 QA Excel 文件",
            str(Path.cwd()),
            "Excel 文件 (*.xlsx *.xlsm)",
        )
        if path:
            self.load_file(path)

    def load_file(self, path: str) -> None:
        if not self._confirm_discard_dirty():
            return
        try:
            self.store.load(path)
        except Exception as exc:
            self.show_error("加载失败", str(exc))
            return

        self.dirty = False
        self.file_label.setText(str(Path(path).resolve()))
        self.refresh_all()
        self.update_standardization_status()

    def refresh_all(self) -> None:
        self._loading_ui = True
        self.category_combo.clear()
        self.category_combo.addItems(self.store.categories())
        self.template_category_combo.clear()
        self.template_category_combo.addItems(self.store.categories())
        self._loading_ui = False

        self.current_category = self.category_combo.currentText().strip()
        self.current_region = ""
        self.current_poi = ""
        self.current_intent = ""
        self.populate_regions()
        self.refresh_tree()
        self.load_template_table(self.template_category_combo.currentText())
        self._set_data_actions_enabled(self.store.standardized)
        self._set_template_actions_enabled(self.store.standardized)

    def update_standardization_status(self) -> None:
        if self.store.path is None or self.store.report is None:
            self.status_label.setText("未加载")
            self.status_label.setObjectName("statusNeutral")
            self.standardize_button.setEnabled(False)
        elif self.store.standardized:
            self.status_label.setText("已标准化，可编辑")
            self.status_label.setObjectName("statusOk")
            self.standardize_button.setEnabled(False)
        else:
            issues = self.store.report.issue_count
            self.status_label.setText(f"待标准化：{issues} 项")
            self.status_label.setObjectName("statusWarn")
            self.standardize_button.setEnabled(True)
        self.status_label.style().unpolish(self.status_label)
        self.status_label.style().polish(self.status_label)
        self._set_data_actions_enabled(self.store.standardized)
        self._set_template_actions_enabled(self.store.standardized)

    def standardize_current_file(self) -> None:
        if self.store.path is None:
            return
        reply = QMessageBox.question(
            self,
            "确认标准化",
            "将直接修改当前 Excel 文件：取消 A-D 列合并单元格，并向下填充分类、区域、POI、咨询意图。是否继续？",
        )
        if reply != QMessageBox.Yes:
            return
        try:
            self.store.standardize_in_place()
        except Exception as exc:
            self.show_error("标准化失败", str(exc))
            return
        self.refresh_all()
        self.update_standardization_status()
        QMessageBox.information(self, "完成", "当前文件已经标准化，可以开始编辑。")

    def save_changes(self) -> None:
        if not self.store.path or not self.dirty:
            return
        selection = (self.current_category, self.current_region, self.current_poi, self.current_intent)
        try:
            self.store.save_all()
        except Exception as exc:
            self.show_error("保存失败", str(exc))
            return
        self.dirty = False
        self.save_button.setEnabled(False)
        self.refresh_all()
        self.set_selection(*selection)
        QMessageBox.information(self, "已保存", "回复修改已写入当前 Excel 文件。")

    def populate_regions(self) -> None:
        if self._loading_ui:
            return
        self._loading_ui = True
        self.current_category = self.category_combo.currentText().strip()
        self.region_combo.set_values(self.store.regions(self.current_category), self.current_region)
        self.current_region = self.region_combo.current_value()
        self._loading_ui = False
        self.populate_pois()

    def populate_pois(self) -> None:
        if self._loading_ui:
            return
        self._loading_ui = True
        self.poi_combo.set_values(self.store.pois(self.current_category, self.current_region), self.current_poi)
        self.current_poi = self.poi_combo.current_value()
        self._loading_ui = False
        self.populate_intents()

    def populate_intents(self) -> None:
        if self._loading_ui:
            return
        self._loading_ui = True
        self.intent_combo.set_values(
            self.store.intents(self.current_category, self.current_region, self.current_poi),
            self.current_intent,
        )
        self.current_intent = self.intent_combo.current_value()
        self._loading_ui = False
        self.display_current_intent()

    def on_category_changed(self, category: str) -> None:
        if self._loading_ui:
            return
        self.current_category = category
        self.current_region = ""
        self.current_poi = ""
        self.current_intent = ""
        self.populate_regions()

    def on_region_changed(self, region: str) -> None:
        if self._loading_ui:
            return
        self.current_region = region
        self.current_poi = ""
        self.current_intent = ""
        self.populate_pois()
        self.select_tree_path()

    def on_poi_changed(self, poi: str) -> None:
        if self._loading_ui:
            return
        self.current_poi = poi
        self.current_intent = ""
        self.populate_intents()
        self.select_tree_path()

    def on_intent_changed(self, intent: str) -> None:
        if self._loading_ui:
            return
        self.current_intent = intent
        self.display_current_intent()
        self.select_tree_path()

    def display_current_intent(self) -> None:
        self.current_records = self.store.qa_pairs(
            self.current_category,
            self.current_region,
            self.current_poi,
            self.current_intent,
        )
        self.path_title.setText(self.current_path_text())
        self.qa_count_label.setText(f"{len(self.current_records)} 条问答" if self.current_records else "")
        self.add_qa_button.setEnabled(self.store.standardized and bool(self.current_poi and self.current_intent))
        self.bulk_answer_button.setEnabled(False)
        self.qa_table.setRowCount(0)
        self.selected_record = None
        self.question_preview.clear()
        self.answer_editor.blockSignals(True)
        self.answer_editor.clear()
        self.answer_editor.blockSignals(False)

        for row_idx, record in enumerate(self.current_records):
            self.qa_table.insertRow(row_idx)
            question_item = QTableWidgetItem(display_text(record.question))
            question_item.setFlags(question_item.flags() & ~Qt.ItemIsEditable)
            question_item.setData(Qt.UserRole, record.record_id)
            answer_item = QTableWidgetItem(record.answer)
            answer_item.setFlags(answer_item.flags() & ~Qt.ItemIsEditable)
            answer_item.setData(Qt.UserRole, record.record_id)
            self.qa_table.setItem(row_idx, 0, question_item)
            self.qa_table.setItem(row_idx, 1, answer_item)
            self.qa_table.setRowHeight(row_idx, 58)

        if self.current_records:
            self.qa_table.selectRow(0)

    def on_qa_selected(self) -> None:
        items = self.qa_table.selectedItems()
        if not items:
            return
        row = items[0].row()
        if row < 0 or row >= len(self.current_records):
            return
        self.selected_record = self.current_records[row]
        self.question_preview.setPlainText(display_text(self.selected_record.question))
        self.answer_editor.blockSignals(True)
        self.answer_editor.setPlainText(self.selected_record.answer)
        self.answer_editor.blockSignals(False)
        self.bulk_answer_button.setEnabled(self.store.standardized and self.selected_record is not None)

    def on_answer_changed(self) -> None:
        if self._loading_ui or self.selected_record is None:
            return
        self.selected_record.answer = self.answer_editor.toPlainText()
        selected = self.qa_table.selectedItems()
        if selected:
            row = selected[0].row()
            answer_item = self.qa_table.item(row, 1)
            if answer_item:
                answer_item.setText(self.selected_record.answer)
        self.mark_dirty()

    def mark_dirty(self) -> None:
        self.dirty = True
        self.save_button.setEnabled(True)

    def refresh_tree(self) -> None:
        self.tree.clear()
        for category in self.store.categories():
            category_item = QTreeWidgetItem([category])
            category_item.setData(0, Qt.UserRole, {"category": category})
            self.tree.addTopLevelItem(category_item)
            for region in self.store.regions(category):
                region_item = QTreeWidgetItem([region])
                region_item.setData(0, Qt.UserRole, {"category": category, "region": region})
                category_item.addChild(region_item)
                for poi in self.store.pois(category, region):
                    poi_item = QTreeWidgetItem([poi])
                    poi_item.setData(0, Qt.UserRole, {"category": category, "region": region, "poi": poi})
                    region_item.addChild(poi_item)
                    for intent in self.store.intents(category, region, poi):
                        intent_item = QTreeWidgetItem([display_text(intent)])
                        intent_item.setData(
                            0,
                            Qt.UserRole,
                            {"category": category, "region": region, "poi": poi, "intent": intent},
                        )
                        poi_item.addChild(intent_item)
        self.tree.expandToDepth(1)

    def on_tree_item_clicked(self, item: QTreeWidgetItem) -> None:
        data = item.data(0, Qt.UserRole) or {}
        self._loading_ui = True
        if "category" in data:
            self.current_category = data["category"]
            idx = self.category_combo.findText(self.current_category)
            if idx >= 0:
                self.category_combo.setCurrentIndex(idx)
        self._loading_ui = False
        self.populate_regions()

        if data.get("region"):
            self.current_region = data["region"]
            self.region_combo.select_value(self.current_region)
            self.populate_pois()
        if data.get("poi"):
            self.current_poi = data["poi"]
            self.poi_combo.select_value(self.current_poi)
            self.populate_intents()
        if data.get("intent"):
            self.current_intent = data["intent"]
            self.intent_combo.select_value(self.current_intent)
            self.display_current_intent()

    def select_tree_path(self) -> None:
        target = {
            "category": self.current_category,
            "region": self.current_region,
            "poi": self.current_poi,
            "intent": self.current_intent,
        }
        match = self.find_tree_item(target)
        if match:
            self.tree.blockSignals(True)
            self.tree.setCurrentItem(match)
            self.tree.scrollToItem(match)
            self.tree.blockSignals(False)

    def find_tree_item(self, target: dict[str, str]) -> QTreeWidgetItem | None:
        def matches(item: QTreeWidgetItem) -> bool:
            data = item.data(0, Qt.UserRole) or {}
            return all(not value or data.get(key) == value for key, value in target.items())

        stack = [self.tree.topLevelItem(i) for i in range(self.tree.topLevelItemCount())]
        while stack:
            item = stack.pop(0)
            if item and matches(item):
                return item
            for idx in range(item.childCount()):
                stack.append(item.child(idx))
        return None

    def add_poi(self) -> None:
        if not self.ensure_editable():
            return
        dialog = AddPoiDialog(self.store, self)
        if dialog.exec() != QDialog.Accepted:
            return
        category, region, poi = dialog.values()
        try:
            self.store.add_poi(category, region, poi)
        except Exception as exc:
            self.show_error("新增 POI 失败", str(exc))
            return
        self.refresh_all()
        self.set_selection(category, region, poi, "")

    def add_qa_pair(self) -> None:
        if not self.ensure_editable():
            return
        if not self.current_category or not self.current_region or not self.current_poi or not self.current_intent:
            QMessageBox.information(self, "请选择咨询意图", "请先选择一个 POI 下的咨询意图，再新增问答。")
            return

        dialog = AddQAPairDialog(self.current_intent, self)
        if dialog.exec() != QDialog.Accepted:
            return
        question, answer = dialog.values()
        selection = (self.current_category, self.current_region, self.current_poi, self.current_intent)
        try:
            self.store.add_qa_pair(*selection, question, answer)
        except Exception as exc:
            self.show_error("新增问答失败", str(exc))
            return
        self.set_selection(*selection)
        for row_idx, record in enumerate(self.current_records):
            if record.question == question and record.answer == answer:
                self.qa_table.selectRow(row_idx)
                break
        QMessageBox.information(self, "已新增", "这条 POI 专属问答已写入当前 Excel 文件。")

    def bulk_update_answer(self) -> None:
        if not self.ensure_editable():
            return
        if self.selected_record is None:
            QMessageBox.information(self, "请选择问答", "请先选中一条具体问答。")
            return

        matches = self.store.matching_question_records(
            self.selected_record.category,
            self.selected_record.intent,
            self.selected_record.question,
        )
        poi_count = len({(record.region, record.poi) for record in matches})
        dialog = BulkAnswerDialog(
            self.selected_record.category,
            self.selected_record.intent,
            self.selected_record.question,
            self.answer_editor.toPlainText(),
            len(matches),
            poi_count,
            self,
        )
        if dialog.exec() != QDialog.Accepted:
            return

        selection = (
            self.current_category,
            self.current_region,
            self.current_poi,
            self.current_intent,
        )
        question = self.selected_record.question
        try:
            updated, affected_pois = self.store.bulk_update_answer(
                self.selected_record.category,
                self.selected_record.intent,
                self.selected_record.question,
                dialog.answer(),
            )
        except Exception as exc:
            self.show_error("批量修改失败", str(exc))
            return

        self.dirty = False
        self.save_button.setEnabled(False)
        self.set_selection(*selection)
        for row_idx, record in enumerate(self.current_records):
            if record.question == question:
                self.qa_table.selectRow(row_idx)
                break
        QMessageBox.information(self, "已批量修改", f"已覆盖 {affected_pois} 个 POI、{updated} 条问答的回复。")

    def delete_poi(self) -> None:
        if not self.ensure_editable():
            return
        if not self.current_poi:
            return
        reply = QMessageBox.question(
            self,
            "确认删除",
            f"将删除“{self.current_category} / {self.current_region} / {self.current_poi}”下的全部问答，是否继续？",
        )
        if reply != QMessageBox.Yes:
            return
        try:
            removed = self.store.delete_poi(self.current_category, self.current_region, self.current_poi)
        except Exception as exc:
            self.show_error("删除失败", str(exc))
            return
        self.refresh_all()
        QMessageBox.information(self, "已删除", f"已删除 {removed} 条问答。")

    def rename_poi(self) -> None:
        if not self.ensure_editable():
            return
        if not self.current_poi:
            return
        category = self.current_category
        region = self.current_region
        old_poi = self.current_poi
        new_name, ok = QInputDialog.getText(self, "重命名 POI", "新的 POI 名称：", text=self.current_poi)
        if not ok:
            return
        try:
            changed = self.store.rename_poi(category, region, old_poi, new_name)
        except Exception as exc:
            self.show_error("重命名失败", str(exc))
            return
        self.refresh_all()
        self.set_selection(category, region, new_name.strip(), "")
        QMessageBox.information(self, "已重命名", f"已更新 {changed} 条问答。")

    def load_template_table(self, category: str) -> None:
        self.template_table.setRowCount(0)
        if not category:
            return
        for item in self.store.template_for_category(category):
            row = self.template_table.rowCount()
            self.template_table.insertRow(row)
            intent_item = QTableWidgetItem(display_text(item.intent))
            question_item = QTableWidgetItem(display_text(item.question))
            intent_item.setData(CURRENT_VALUE_ROLE, item.intent)
            question_item.setData(CURRENT_VALUE_ROLE, item.question)
            intent_item.setData(SOURCE_INTENT_ROLE, item.source_intent)
            question_item.setData(SOURCE_QUESTION_ROLE, item.source_question)
            self.template_table.setItem(row, 0, intent_item)
            self.template_table.setItem(row, 1, question_item)

    def add_template_question(self) -> None:
        dialog = TemplateQuestionDialog(
            self.template_intents(),
            "新增模板问题",
            parent=self,
        )
        if dialog.exec() != QDialog.Accepted:
            return
        intent, question = dialog.values()
        row = self.template_table.currentRow()
        insert_at = row + 1 if row >= 0 else self.template_table.rowCount()
        self.template_table.insertRow(insert_at)
        intent_item = QTableWidgetItem(display_text(intent))
        question_item = QTableWidgetItem(display_text(question))
        intent_item.setData(CURRENT_VALUE_ROLE, intent)
        question_item.setData(CURRENT_VALUE_ROLE, question)
        self.template_table.setItem(insert_at, 0, intent_item)
        self.template_table.setItem(insert_at, 1, question_item)
        self.template_table.selectRow(insert_at)

    def edit_template_question(self) -> None:
        row = self.template_table.currentRow()
        if row < 0:
            QMessageBox.information(self, "请选择问题", "请先选中一条模板问题。")
            return

        intent_item = self.template_table.item(row, 0)
        question_item = self.template_table.item(row, 1)
        dialog = TemplateQuestionDialog(
            self.template_intents(),
            "编辑模板问题",
            intent_item.text() if intent_item else "",
            question_item.text() if question_item else "",
            self,
        )
        if dialog.exec() != QDialog.Accepted:
            return

        intent, question = dialog.values()
        source_intent = intent_item.data(SOURCE_INTENT_ROLE) if intent_item else None
        source_question = question_item.data(SOURCE_QUESTION_ROLE) if question_item else None

        new_intent_item = QTableWidgetItem(display_text(intent))
        new_question_item = QTableWidgetItem(display_text(question))
        new_intent_item.setData(CURRENT_VALUE_ROLE, intent)
        new_question_item.setData(CURRENT_VALUE_ROLE, question)
        new_intent_item.setData(SOURCE_INTENT_ROLE, source_intent)
        new_question_item.setData(SOURCE_QUESTION_ROLE, source_question)
        self.template_table.setItem(row, 0, new_intent_item)
        self.template_table.setItem(row, 1, new_question_item)
        self.template_table.selectRow(row)

    def template_intents(self) -> list[str]:
        intents = []
        seen = set()
        for row in range(self.template_table.rowCount()):
            item = self.template_table.item(row, 0)
            if not item:
                continue
            intent = item.text().strip()
            if intent and intent not in seen:
                intents.append(intent)
                seen.add(intent)
        return intents

    def delete_template_rows(self) -> None:
        rows = sorted({index.row() for index in self.template_table.selectedIndexes()}, reverse=True)
        for row in rows:
            self.template_table.removeRow(row)

    def move_template_row(self, direction: int) -> None:
        row = self.template_table.currentRow()
        target = row + direction
        if row < 0 or target < 0 or target >= self.template_table.rowCount():
            return
        values = []
        for col in range(2):
            item = self.template_table.takeItem(row, col)
            values.append(item)
        self.template_table.removeRow(row)
        self.template_table.insertRow(target)
        for col, item in enumerate(values):
            self.template_table.setItem(target, col, item)
        self.template_table.setCurrentCell(target, 0)

    def apply_template(self) -> None:
        if not self.ensure_editable():
            return
        category = self.template_category_combo.currentText().strip()
        items: list[TemplateItem] = []
        for row in range(self.template_table.rowCount()):
            intent_item = self.template_table.item(row, 0)
            question_item = self.template_table.item(row, 1)
            intent = self.template_item_value(intent_item)
            question = self.template_item_value(question_item)
            source_intent = intent_item.data(SOURCE_INTENT_ROLE) if intent_item else None
            source_question = question_item.data(SOURCE_QUESTION_ROLE) if question_item else None
            if not intent and not question:
                continue
            items.append(TemplateItem(intent, question, source_intent, source_question))

        reply = QMessageBox.question(
            self,
            "确认应用模板",
            f"将把当前模板同步到分类“{category}”下所有 POI。删除的模板问法会从 Excel 中移除，新增问法回复为空。是否继续？",
        )
        if reply != QMessageBox.Yes:
            return
        try:
            self.store.apply_template(category, items)
        except Exception as exc:
            self.show_error("应用模板失败", str(exc))
            return
        self.dirty = False
        self.save_button.setEnabled(False)
        self.refresh_all()
        QMessageBox.information(self, "已应用", f"分类“{category}”的问题模板已同步。")

    def template_item_value(self, item: QTableWidgetItem | None) -> str:
        if item is None:
            return ""
        value = item.data(CURRENT_VALUE_ROLE)
        if value:
            return str(value).strip()
        return item.text().strip()

    def set_selection(self, category: str, region: str, poi: str, intent: str) -> None:
        self._loading_ui = True
        idx = self.category_combo.findText(category)
        if idx >= 0:
            self.category_combo.setCurrentIndex(idx)
        self.current_category = category
        self._loading_ui = False
        self.populate_regions()
        self.current_region = region
        self.region_combo.select_value(region)
        self.populate_pois()
        self.current_poi = poi
        self.poi_combo.select_value(poi)
        self.populate_intents()
        if intent:
            self.current_intent = intent
            self.intent_combo.select_value(intent)
        self.display_current_intent()
        self.select_tree_path()

    def restore_current_selection(self) -> None:
        self.set_selection(self.current_category, self.current_region, self.current_poi, self.current_intent)

    def current_path_text(self) -> str:
        parts = [self.current_category, self.current_region, self.current_poi, display_text(self.current_intent)]
        parts = [part for part in parts if part]
        return " / ".join(parts) if parts else "请选择一个咨询意图"

    def ensure_editable(self) -> bool:
        if not self.store.path:
            return False
        if not self.store.standardized:
            QMessageBox.information(self, "需要先标准化", "当前文件还不是完整明细结构，请先点击“标准化 A-D 列”。")
            return False
        return True

    def _set_data_actions_enabled(self, enabled: bool) -> None:
        for widget in (
            self.add_poi_button,
            self.add_qa_button,
            self.bulk_answer_button,
            self.rename_poi_button,
            self.delete_poi_button,
            self.answer_editor,
        ):
            widget.setEnabled(enabled)

    def _set_template_actions_enabled(self, enabled: bool) -> None:
        for widget in (
            self.template_table,
            self.template_add_button,
            self.template_edit_button,
            self.template_delete_button,
            self.template_up_button,
            self.template_down_button,
            self.template_apply_button,
            self.template_category_combo,
        ):
            widget.setEnabled(enabled)

    def _confirm_discard_dirty(self) -> bool:
        if not self.dirty:
            return True
        reply = QMessageBox.question(
            self,
            "还有未保存修改",
            "当前回复修改还没有保存。是否放弃这些修改并继续？",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No,
        )
        return reply == QMessageBox.Yes

    def show_error(self, title: str, message: str) -> None:
        QMessageBox.critical(self, title, message)

    def closeEvent(self, event: QCloseEvent) -> None:
        if self._confirm_discard_dirty():
            event.accept()
        else:
            event.ignore()

    def _apply_style(self) -> None:
        QApplication.instance().setStyleSheet(
            """
            QWidget {
                font-family: "Microsoft YaHei", "Segoe UI", sans-serif;
                font-size: 13px;
                color: #20242a;
                background: #f6f7f2;
            }
            QMainWindow {
                background: #f6f7f2;
            }
            QFrame#topBar, QFrame#panel, QFrame#editorBox {
                background: #ffffff;
                border: 1px solid #dde2d6;
                border-radius: 8px;
            }
            QLabel#appTitle {
                font-size: 21px;
                font-weight: 700;
                color: #17211f;
            }
            QLabel#contentTitle {
                font-size: 18px;
                font-weight: 700;
                color: #17211f;
            }
            QLabel#sectionTitle {
                font-weight: 700;
                color: #31413d;
            }
            QLabel#mutedLabel {
                color: #69716e;
            }
            QLabel#statusOk, QLabel#statusWarn, QLabel#statusNeutral {
                padding: 6px 10px;
                border-radius: 8px;
                font-weight: 700;
            }
            QLabel#statusOk {
                color: #0c594b;
                background: #dff3ea;
                border: 1px solid #a8d8c7;
            }
            QLabel#statusWarn {
                color: #815315;
                background: #fff1ce;
                border: 1px solid #e6c46c;
            }
            QLabel#statusNeutral {
                color: #4e5656;
                background: #edf0ec;
                border: 1px solid #d6dad4;
            }
            QPushButton {
                background: #f3f6f3;
                border: 1px solid #cfd7ce;
                border-radius: 7px;
                padding: 8px 12px;
                color: #1f2b29;
            }
            QPushButton:hover {
                background: #e9f3ee;
                border-color: #8fbbaa;
            }
            QPushButton:pressed {
                background: #d9ebe2;
            }
            QPushButton:disabled {
                color: #9aa19e;
                background: #eceeeb;
                border-color: #dde0dc;
            }
            QPushButton#dangerButton {
                color: #8c2f28;
                background: #fff0ed;
                border-color: #e3b3aa;
            }
            QPushButton#dangerButton:hover {
                background: #ffe3dd;
                border-color: #cf8b7d;
            }
            QLineEdit, QTextEdit, QPlainTextEdit, QComboBox {
                background: #fbfcfb;
                border: 1px solid #cfd7ce;
                border-radius: 7px;
                padding: 7px 9px;
                selection-background-color: #b7d8cb;
            }
            QLineEdit:focus, QTextEdit:focus, QPlainTextEdit:focus, QComboBox:focus {
                border-color: #4f9f89;
            }
            QTreeWidget, QTableWidget, QComboBox QAbstractItemView {
                background: #fbfcfb;
                border: 1px solid #d6ddd3;
                border-radius: 8px;
                alternate-background-color: #f1f5f2;
                gridline-color: #e0e5df;
            }
            QTreeWidget::item, QComboBox QAbstractItemView::item {
                min-height: 28px;
                padding: 6px 8px;
                border-radius: 5px;
            }
            QTreeWidget::item:selected, QComboBox QAbstractItemView::item:selected {
                color: #0f2924;
                background: #cfe8de;
            }
            QHeaderView::section {
                background: #edf3ef;
                border: 0;
                border-right: 1px solid #d6ddd3;
                border-bottom: 1px solid #d6ddd3;
                padding: 8px;
                font-weight: 700;
            }
            QTableWidget::item {
                padding: 8px;
            }
            QTableWidget::item:selected {
                color: #0f2924;
                background: #cfe8de;
            }
            QTabWidget::pane {
                border: 0;
            }
            QTabBar::tab {
                background: #e9ede7;
                border: 1px solid #cfd7ce;
                border-bottom: 0;
                padding: 9px 18px;
                margin-right: 4px;
                border-top-left-radius: 7px;
                border-top-right-radius: 7px;
            }
            QTabBar::tab:selected {
                background: #ffffff;
                color: #0d6656;
                font-weight: 700;
            }
            QSplitter::handle {
                background: #dbe2d9;
            }
            """
        )


def main() -> None:
    app = QApplication(sys.argv)
    app.setApplicationName(APP_TITLE)
    app.setWindowIcon(QIcon())
    initial_file = sys.argv[1] if len(sys.argv) > 1 else None
    window = MainWindow(initial_file)
    window.show()
    sys.exit(app.exec())
