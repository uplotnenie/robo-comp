"""
Главное окно приложения DXF-Auto.

Реализует трёхпанельный интерфейс:
- Левая панель: дерево состава сборки
- Центральная панель: таблица листовых деталей
- Нижняя панель: логи и отчёты
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from typing import Optional, Dict, List, Any
from pathlib import Path
import threading

from .composition_tree import CompositionTree
from .sheet_table import SheetTable
from .settings_dialog import SettingsDialog
from .export_dialog import ExportDialog

# Условные импорты для type checking
from typing import TYPE_CHECKING
if TYPE_CHECKING:
    from core import KompasAPI, AssemblyScanner, DXFExporter
    from models import SheetPart, AssemblyNode, ExportSettings


class MainWindow:
    """Главное окно приложения."""
    
    APP_TITLE = "DXF-Auto - Экспорт развёрток из КОМПАС-3D"
    
    def __init__(self, root: tk.Tk):
        """
        Инициализация главного окна.
        
        Args:
            root: Корневой виджет Tk
        """
        self.root = root
        self.root.title(self.APP_TITLE)
        self.root.geometry("1200x800")
        self.root.minsize(800, 600)
        
        # Состояние приложения
        self._kompas_api: Optional['KompasAPI'] = None
        self._scanner: Optional['AssemblyScanner'] = None
        self._exporter: Optional['DXFExporter'] = None
        
        self._current_assembly: Optional['AssemblyNode'] = None
        self._sheet_parts: Dict[str, 'SheetPart'] = {}
        self._settings: Optional['ExportSettings'] = None
        
        self._is_connected = False
        
        # Инициализация настроек по умолчанию
        self._init_default_settings()
        
        # Создание интерфейса
        self._setup_menu()
        self._setup_toolbar()
        self._setup_main_layout()
        self._setup_statusbar()
        
        # Привязки
        self._setup_bindings()
        
        # Проверка подключения к КОМПАС при запуске
        self.root.after(500, self._check_kompas_connection)
        
    def _init_default_settings(self):
        """Инициализация настроек по умолчанию."""
        from models import ExportSettings
        
        self._settings = ExportSettings()
        self._settings.output_directory = str(Path.home() / "Documents" / "DXF_Export")
        self._settings.filename_settings.template = "{designation}_{name}"
        self._settings.remove_bend_lines = True
        
    def _setup_menu(self):
        """Создание главного меню."""
        self.menubar = tk.Menu(self.root)
        self.root.config(menu=self.menubar)
        
        # Меню Файл
        file_menu = tk.Menu(self.menubar, tearoff=0)
        self.menubar.add_cascade(label="Файл", menu=file_menu)
        
        file_menu.add_command(
            label="Подключиться к КОМПАС",
            command=self._connect_to_kompas,
            accelerator="Ctrl+K"
        )
        file_menu.add_command(
            label="Сканировать сборку",
            command=self._scan_assembly,
            accelerator="F5"
        )
        file_menu.add_separator()
        file_menu.add_command(
            label="Экспорт выбранных...",
            command=self._export_selected,
            accelerator="Ctrl+E"
        )
        file_menu.add_command(
            label="Экспорт всех...",
            command=self._export_all,
            accelerator="Ctrl+Shift+E"
        )
        file_menu.add_separator()
        file_menu.add_command(
            label="Выход",
            command=self._on_exit,
            accelerator="Alt+F4"
        )
        
        # Меню Редактирование
        edit_menu = tk.Menu(self.menubar, tearoff=0)
        self.menubar.add_cascade(label="Редактирование", menu=edit_menu)
        
        edit_menu.add_command(
            label="Выбрать всё",
            command=self._select_all,
            accelerator="Ctrl+A"
        )
        edit_menu.add_command(
            label="Снять выбор",
            command=self._clear_selection,
            accelerator="Ctrl+D"
        )
        edit_menu.add_separator()
        edit_menu.add_command(
            label="Настройки...",
            command=self._show_settings,
            accelerator="Ctrl+,"
        )
        
        # Меню Вид
        view_menu = tk.Menu(self.menubar, tearoff=0)
        self.menubar.add_cascade(label="Вид", menu=view_menu)
        
        view_menu.add_command(
            label="Обновить",
            command=self._refresh_view,
            accelerator="F5"
        )
        view_menu.add_separator()
        view_menu.add_checkbutton(
            label="Показать панель состава"
        )
        view_menu.add_checkbutton(
            label="Показать панель логов"
        )
        
        # Меню Справка
        help_menu = tk.Menu(self.menubar, tearoff=0)
        self.menubar.add_cascade(label="Справка", menu=help_menu)
        
        help_menu.add_command(
            label="О программе...",
            command=self._show_about
        )
        
    def _setup_toolbar(self):
        """Создание панели инструментов."""
        self.toolbar = ttk.Frame(self.root)
        self.toolbar.pack(fill=tk.X, padx=5, pady=5)
        
        # Кнопка подключения
        self.btn_connect = ttk.Button(
            self.toolbar,
            text="🔌 Подключение",
            command=self._connect_to_kompas
        )
        self.btn_connect.pack(side=tk.LEFT, padx=2)
        
        # Кнопка сканирования
        self.btn_scan = ttk.Button(
            self.toolbar,
            text="🔍 Сканировать",
            command=self._scan_assembly,
            state=tk.DISABLED
        )
        self.btn_scan.pack(side=tk.LEFT, padx=2)
        
        # Разделитель
        ttk.Separator(self.toolbar, orient=tk.VERTICAL).pack(
            side=tk.LEFT, fill=tk.Y, padx=10
        )
        
        # Кнопка экспорта
        self.btn_export = ttk.Button(
            self.toolbar,
            text="📤 Экспорт DXF",
            command=self._export_selected,
            state=tk.DISABLED
        )
        self.btn_export.pack(side=tk.LEFT, padx=2)
        
        # Кнопка настроек
        self.btn_settings = ttk.Button(
            self.toolbar,
            text="⚙️ Настройки",
            command=self._show_settings
        )
        self.btn_settings.pack(side=tk.LEFT, padx=2)
        
        # Индикатор статуса подключения
        self.lbl_connection = ttk.Label(
            self.toolbar,
            text="⚫ Не подключено",
            foreground='gray'
        )
        self.lbl_connection.pack(side=tk.RIGHT, padx=10)
        
    def _setup_main_layout(self):
        """Создание основной раскладки."""
        # Основной PanedWindow
        self.main_paned = ttk.PanedWindow(self.root, orient=tk.VERTICAL)
        self.main_paned.pack(fill=tk.BOTH, expand=True, padx=5, pady=(0, 5))
        
        # Верхняя часть (дерево + таблица)
        top_paned = ttk.PanedWindow(self.main_paned, orient=tk.HORIZONTAL)
        self.main_paned.add(top_paned, weight=3)
        
        # Левая панель - дерево состава
        left_frame = ttk.Frame(top_paned)
        top_paned.add(left_frame, weight=1)
        
        self.composition_tree = CompositionTree(
            left_frame,
            on_selection_changed=self._on_tree_selection_changed,
            on_part_double_click=self._on_part_double_click
        )
        self.composition_tree.pack(fill=tk.BOTH, expand=True)
        
        # Правая панель - таблица деталей
        right_frame = ttk.Frame(top_paned)
        top_paned.add(right_frame, weight=2)
        
        self.sheet_table = SheetTable(
            right_frame,
            on_selection_changed=self._on_table_selection_changed,
            on_row_double_click=self._on_part_double_click
        )
        self.sheet_table.pack(fill=tk.BOTH, expand=True)
        
        # Нижняя часть - логи
        bottom_frame = ttk.LabelFrame(self.main_paned, text="Журнал операций")
        self.main_paned.add(bottom_frame, weight=1)
        
        # Текстовое поле для логов
        log_frame = ttk.Frame(bottom_frame)
        log_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        self.txt_log = tk.Text(
            log_frame,
            height=6,
            wrap=tk.WORD,
            state=tk.DISABLED,
            font=('Consolas', 9)
        )
        self.txt_log.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        log_scrollbar = ttk.Scrollbar(log_frame, command=self.txt_log.yview)
        log_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.txt_log.configure(yscrollcommand=log_scrollbar.set)
        
        # Теги для цветного лога
        self.txt_log.tag_configure('info', foreground='black')
        self.txt_log.tag_configure('success', foreground='green')
        self.txt_log.tag_configure('error', foreground='red')
        self.txt_log.tag_configure('warning', foreground='orange')
        
    def _setup_statusbar(self):
        """Создание строки состояния."""
        self.statusbar = ttk.Frame(self.root)
        self.statusbar.pack(fill=tk.X, side=tk.BOTTOM)
        
        # Статус слева
        self.lbl_status = ttk.Label(self.statusbar, text="Готово")
        self.lbl_status.pack(side=tk.LEFT, padx=10, pady=2)
        
        # Информация справа
        self.lbl_info = ttk.Label(self.statusbar, text="")
        self.lbl_info.pack(side=tk.RIGHT, padx=10, pady=2)
        
    def _setup_bindings(self):
        """Настройка горячих клавиш."""
        self.root.bind('<Control-k>', lambda e: self._connect_to_kompas())
        self.root.bind('<F5>', lambda e: self._scan_assembly())
        self.root.bind('<Control-e>', lambda e: self._export_selected())
        self.root.bind('<Control-E>', lambda e: self._export_all())
        self.root.bind('<Control-a>', lambda e: self._select_all())
        self.root.bind('<Control-d>', lambda e: self._clear_selection())
        self.root.bind('<Control-comma>', lambda e: self._show_settings())
        
    def _log(self, message: str, level: str = 'info'):
        """
        Добавление сообщения в лог.
        
        Args:
            message: Текст сообщения
            level: Уровень (info, success, error, warning)
        """
        import time
        timestamp = time.strftime("%H:%M:%S")
        
        self.txt_log.configure(state=tk.NORMAL)
        self.txt_log.insert(tk.END, f"[{timestamp}] {message}\n", level)
        self.txt_log.see(tk.END)
        self.txt_log.configure(state=tk.DISABLED)
        
        # Также обновляем статус
        self.lbl_status.configure(text=message)
        
    def _check_kompas_connection(self):
        """Проверка подключения к КОМПАС при запуске."""
        self._log("Проверка подключения к КОМПАС-3D...")
        
        try:
            from core import KompasConnection
            self._kompas_connection = KompasConnection()
            self._kompas_api = self._kompas_connection.connect()
            self._on_connected()
            self._log("Подключено к КОМПАС-3D", 'success')
                
        except Exception as e:
            self._log(f"Ошибка подключения: {e}", 'error')
            self._kompas_api = None
            
    def _connect_to_kompas(self):
        """Подключение к КОМПАС-3D."""
        if self._is_connected:
            self._log("Уже подключено к КОМПАС-3D")
            return
            
        self._log("Подключение к КОМПАС-3D...")
        self.lbl_connection.configure(text="🟡 Подключение...", foreground='orange')
        
        # Подключение в отдельном потоке
        def connect_thread():
            import pythoncom
            pythoncom.CoInitialize()
            try:
                from core import KompasConnection
                self._kompas_connection = KompasConnection()
                self._kompas_api = self._kompas_connection.connect()
                self.root.after(0, self._on_connected)
                    
            except Exception as e:
                # Bind exception into default arg so it's captured correctly
                self.root.after(0, lambda e=e: self._on_connection_failed(str(e)))
            finally:
                pythoncom.CoUninitialize()
                
        thread = threading.Thread(target=connect_thread, daemon=True)
        thread.start()
        
    def _on_connected(self):
        """Обработчик успешного подключения."""
        self._is_connected = True
        
        self.lbl_connection.configure(text="🟢 Подключено", foreground='green')
        self.btn_scan.configure(state=tk.NORMAL)
        self.btn_export.configure(state=tk.NORMAL)
        
        self._log("Подключено к КОМПАС-3D", 'success')
        
        # Инициализация сканера и экспортёра
        if self._kompas_api is not None:
            from core import AssemblyScanner, DXFExporter
            self._scanner = AssemblyScanner(self._kompas_api)
            if self._settings is not None:
                self._exporter = DXFExporter(self._kompas_api, self._settings)
        
    def _on_connection_failed(self, error: str):
        """Обработчик ошибки подключения."""
        self._is_connected = False
        
        self.lbl_connection.configure(text="🔴 Ошибка", foreground='red')
        self._log(f"Ошибка подключения: {error}", 'error')
        
        messagebox.showerror(
            "Ошибка подключения",
            f"Не удалось подключиться к КОМПАС-3D.\n\n{error}\n\n"
            "Убедитесь, что КОМПАС-3D запущен."
        )
        
    def _scan_assembly(self):
        """Сканирование текущей сборки (выполняется в главном потоке для совместимости с COM)."""
        if not self._is_connected or not self._scanner:
            messagebox.showwarning(
                "Нет подключения",
                "Сначала подключитесь к КОМПАС-3D"
            )
            return
            
        self._log("Сканирование сборки...")
        self.lbl_status.configure(text="Сканирование...")
        self.root.update()  # Обновить UI перед блокирующей операцией
        
        try:
            # Получение активного документа
            if self._kompas_api is None:
                self._on_scan_error("API не инициализирован")
                return
            doc = self._kompas_api.active_document
            if not doc:
                self._on_scan_error("Нет открытого документа")
                return
                
            # Сканирование
            if self._scanner is None:
                self._on_scan_error("Сканер не инициализирован")
                return
            assembly_node = self._scanner.scan_document(doc)
            if assembly_node is None:
                self._on_scan_error("Не удалось просканировать документ")
                return
            sheet_parts_list = assembly_node.get_all_sheet_parts()
            sheet_parts = {sp.id: sp for sp in sheet_parts_list}

            self._on_scan_complete(assembly_node, sheet_parts)
            
        except Exception as e:
            self._on_scan_error(str(e))
        
    def _on_scan_complete(self, assembly_node: 'AssemblyNode', sheet_parts: Dict[str, Any]):
        """Обработчик завершения сканирования."""
        self._current_assembly = assembly_node
        self._sheet_parts = sheet_parts
        
        # Загрузка данных в UI
        self.composition_tree.load_assembly(assembly_node, sheet_parts)
        self.sheet_table.load_parts(list(sheet_parts.values()))
        
        # Обновление информации
        count = len(sheet_parts)
        self._log(f"Найдено {count} листовых деталей", 'success')
        self.lbl_info.configure(text=f"Листовых деталей: {count}")
        
    def _on_scan_error(self, error: str):
        """Обработчик ошибки сканирования."""
        self._log(f"Ошибка сканирования: {error}", 'error')
        messagebox.showerror("Ошибка сканирования", error)
        
    def _export_selected(self):
        """Экспорт выбранных деталей."""
        # Получение выбранных деталей из таблицы
        selected = self.sheet_table.get_selected_parts()
        
        if not selected:
            # Пробуем получить из дерева
            selected = self.composition_tree.get_selected_parts()
            
        if not selected:
            messagebox.showinfo(
                "Нет выбранных деталей",
                "Выберите детали для экспорта в таблице или дереве состава."
            )
            return
            
        self._show_export_dialog(selected)
        
    def _export_all(self):
        """Экспорт всех деталей."""
        all_parts = self.sheet_table.get_all_parts()
        
        if not all_parts:
            messagebox.showinfo(
                "Нет деталей",
                "Сначала просканируйте сборку."
            )
            return
            
        self._show_export_dialog(all_parts)
        
    def _show_export_dialog(self, parts: List['SheetPart']):
        """Показ диалога экспорта."""
        dialog = ExportDialog(
            self.root,
            parts,
            self._settings,
            export_function=self._do_export_part
        )
        self.root.wait_window(dialog)
        
        # Обновление статусов в таблице
        for result in dialog.get_results():
            self.sheet_table.update_part_status(
                result.part_id,
                "Экспортировано" if result.success else "Ошибка",
                is_error=not result.success
            )
            
    def _do_export_part(self, part: 'SheetPart', settings: 'ExportSettings') -> str:
        """
        Экспорт одной детали.
        
        Args:
            part: Деталь для экспорта
            settings: Настройки экспорта
            
        Returns:
            Путь к созданному файлу
        """
        if self._exporter:
            # Use export_parts with single item list
            from models import SheetPartInfo
            part_info = part if isinstance(part, SheetPartInfo) else part.info
            part_info.export_selected = True
            summary = self._exporter.export_parts([part_info])
            if summary.results and summary.results[0].success:
                return summary.results[0].output_path
            elif summary.results:
                raise RuntimeError(summary.results[0].error_message)
            else:
                raise RuntimeError("Ошибка экспорта")
        else:
            raise RuntimeError("Экспортёр не инициализирован")
            
    def _select_all(self):
        """Выбор всех деталей."""
        self.composition_tree._select_all_sheet()
        self.sheet_table.select_all()
        
    def _clear_selection(self):
        """Снятие выбора."""
        self.composition_tree._clear_selection()
        self.sheet_table.clear_selection()
        
    def _show_settings(self):
        """Показ диалога настроек."""
        dialog = SettingsDialog(
            self.root,
            self._settings,
            on_save=self._on_settings_saved
        )
        self.root.wait_window(dialog)
        
    def _on_settings_saved(self, settings: 'ExportSettings'):
        """Обработчик сохранения настроек."""
        self._settings = settings
        self._log("Настройки сохранены", 'success')
        
    def _refresh_view(self):
        """Обновление представления."""
        if self._is_connected:
            self._scan_assembly()
            
    def _on_tree_selection_changed(self, part_ids: List[str]):
        """Обработчик изменения выбора в дереве."""
        # Синхронизация с таблицей (опционально)
        pass
        
    def _on_table_selection_changed(self, part_ids: List[str]):
        """Обработчик изменения выбора в таблице."""
        count = len(part_ids)
        self.lbl_info.configure(text=f"Выбрано: {count}")
        
    def _on_part_double_click(self, part_id: str):
        """Обработчик двойного клика на детали."""
        part = self._sheet_parts.get(part_id)
        if part:
            self._show_part_details(part)
            
    def _show_part_details(self, part: 'SheetPart'):
        """Показ деталей детали."""
        info = part.info
        
        details = (
            f"Наименование: {info.name or '—'}\n"
            f"Обозначение: {info.designation or '—'}\n"
            f"Материал: {info.material or '—'}\n"
            f"Толщина: {info.thickness or '—'} мм\n"
            f"Количество: {info.quantity}\n"
            f"Файл: {info.file_name or '—'}"
        )
        
        messagebox.showinfo(
            f"Деталь: {info.display_name}",
            details
        )
        
    def _show_about(self):
        """Показ информации о программе."""
        messagebox.showinfo(
            "О программе",
            "DXF-Auto v1.0\n\n"
            "Экспорт развёрток листовых деталей\n"
            "из КОМПАС-3D в формат DXF\n"
            "для лазерной резки.\n\n"
            "© 2024"
        )
        
    def _on_exit(self):
        """Выход из приложения."""
        if messagebox.askyesno("Выход", "Выйти из программы?"):
            # Отключение от КОМПАС
            if hasattr(self, '_kompas_connection') and self._kompas_connection:
                try:
                    self._kompas_connection.disconnect()
                except:
                    pass
            self.root.quit()
            
    def run(self):
        """Запуск главного цикла."""
        self.root.mainloop()
