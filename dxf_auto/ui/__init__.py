"""
UI модуль приложения DXF-Auto.

Содержит компоненты графического интерфейса:
- MainWindow: главное окно приложения
- CompositionTree: дерево состава сборки
- SheetTable: таблица листовых деталей
- SettingsDialog: диалоги настроек
- ReportsView: панель отчётов и логов
"""

from .main_window import MainWindow
from .composition_tree import CompositionTree
from .sheet_table import SheetTable
from .settings_dialog import SettingsDialog
from .export_dialog import ExportDialog

__all__ = [
    'MainWindow',
    'CompositionTree', 
    'SheetTable',
    'SettingsDialog',
    'ExportDialog',
]
