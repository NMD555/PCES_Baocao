from PyQt5.QtWidgets import QComboBox, QCompleter
from PyQt5.QtCore import QStringListModel, QSortFilterProxyModel, Qt
from typing import List

class AutoCompleterComboBox:
    def __init__(self, combo_box: QComboBox, items: List[str]):
        self.combo_box = combo_box
        self.combo_box.setEditable(True)
        self.combo_box.clear()

        # Model gốc
        self.base_model = QStringListModel(items)

        # Proxy model lọc
        self.proxy_model = QSortFilterProxyModel()
        self.proxy_model.setSourceModel(self.base_model)
        self.proxy_model.setFilterCaseSensitivity(Qt.CaseInsensitive)
        self.proxy_model.setFilterRole(Qt.DisplayRole)

        # Completer dùng proxy model
        self.completer = QCompleter(self.proxy_model, self.combo_box)
        self.completer.setCompletionMode(QCompleter.PopupCompletion)
        self.completer.setFilterMode(Qt.MatchContains)
        self.completer.setCaseSensitivity(Qt.CaseInsensitive)

        # Gắn completer vào combo box
        self.combo_box.setCompleter(self.completer)

        # Gán model hiển thị vào combo box
        self.combo_box.setModel(self.proxy_model)
        self.combo_box.setCurrentIndex(-1)
