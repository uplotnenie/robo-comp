"""
Диалоговые окна настроек приложения.

Включает:
- SettingsDialog: основной диалог настроек
- FilenamePatternDialog: настройка паттерна имён файлов
- LineSettingsDialog: настройка типов линий и слоёв
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from typing import Optional, Callable, Dict, TYPE_CHECKING

if TYPE_CHECKING:
    from models import ExportSettings, LineTypeSettings


class SettingsDialog(tk.Toplevel):
    """Основной диалог настроек экспорта."""
    
    def __init__(
        self,
        parent: tk.Misc,
        settings: Optional['ExportSettings'] = None,
        on_save: Optional[Callable[['ExportSettings'], None]] = None
    ):
        """
        Инициализация диалога.
        
        Args:
            parent: Родительское окно
            settings: Текущие настройки
            on_save: Callback при сохранении
        """
        super().__init__(parent)
        
        # Import and create default settings if None
        from models import ExportSettings as ES
        self.settings = settings if settings is not None else ES()
        self.on_save = on_save
        self.result = None
        
        self.title("Настройки экспорта")
        self.geometry("500x450")
        self.resizable(False, False)
        
        # Модальное окно
        if isinstance(parent, (tk.Tk, tk.Toplevel)):
            self.transient(parent)
        self.grab_set()
        
        self._setup_ui()
        self._load_settings()
        
        # Центрирование
        self.update_idletasks()
        x = parent.winfo_rootx() + (parent.winfo_width() - self.winfo_width()) // 2
        y = parent.winfo_rooty() + (parent.winfo_height() - self.winfo_height()) // 2
        self.geometry(f"+{x}+{y}")
        
    def _setup_ui(self):
        """Настройка интерфейса."""
        # Notebook для вкладок
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Вкладка "Общие"
        self._create_general_tab()
        
        # Вкладка "Имена файлов"
        self._create_filename_tab()
        
        # Вкладка "Слои и линии"
        self._create_layers_tab()
        
        # Кнопки
        btn_frame = ttk.Frame(self)
        btn_frame.pack(fill=tk.X, padx=10, pady=10)
        
        ttk.Button(
            btn_frame,
            text="Сохранить",
            command=self._on_save
        ).pack(side=tk.RIGHT, padx=5)
        
        ttk.Button(
            btn_frame,
            text="Отмена",
            command=self._on_cancel
        ).pack(side=tk.RIGHT)
        
    def _create_general_tab(self):
        """Создание вкладки общих настроек."""
        frame = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(frame, text="Общие")
        
        # Папка вывода
        ttk.Label(frame, text="Папка для сохранения DXF:").grid(
            row=0, column=0, sticky=tk.W, pady=5
        )
        
        path_frame = ttk.Frame(frame)
        path_frame.grid(row=1, column=0, sticky='ew', pady=(0, 10))
        
        self.var_output_dir = tk.StringVar()
        self.ent_output_dir = ttk.Entry(path_frame, textvariable=self.var_output_dir)
        self.ent_output_dir.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        ttk.Button(
            path_frame,
            text="Обзор...",
            command=self._browse_output_dir
        ).pack(side=tk.RIGHT, padx=(5, 0))
        
        # Опции
        self.var_create_subfolder = tk.BooleanVar()
        ttk.Checkbutton(
            frame,
            text="Создавать подпапку с именем сборки",
            variable=self.var_create_subfolder
        ).grid(row=2, column=0, sticky=tk.W, pady=5)
        
        self.var_overwrite = tk.BooleanVar()
        ttk.Checkbutton(
            frame,
            text="Перезаписывать существующие файлы",
            variable=self.var_overwrite
        ).grid(row=3, column=0, sticky=tk.W, pady=5)
        
        self.var_open_folder = tk.BooleanVar()
        ttk.Checkbutton(
            frame,
            text="Открыть папку после экспорта",
            variable=self.var_open_folder
        ).grid(row=4, column=0, sticky=tk.W, pady=5)
        
        # Настройка grid
        frame.columnconfigure(0, weight=1)
        
    def _create_filename_tab(self):
        """Создание вкладки настроек имён файлов."""
        frame = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(frame, text="Имена файлов")
        
        # Паттерн имени файла
        ttk.Label(frame, text="Шаблон имени файла:").grid(
            row=0, column=0, sticky=tk.W, pady=5
        )
        
        self.var_filename_pattern = tk.StringVar()
        self.ent_filename_pattern = ttk.Entry(
            frame, 
            textvariable=self.var_filename_pattern,
            width=50
        )
        self.ent_filename_pattern.grid(row=1, column=0, sticky='ew', pady=(0, 5))
        
        # Подсказка по переменным
        hint_frame = ttk.LabelFrame(frame, text="Доступные переменные", padding=5)
        hint_frame.grid(row=2, column=0, sticky='ew', pady=10)
        
        variables = [
            ("{name}", "Наименование детали"),
            ("{designation}", "Обозначение"),
            ("{material}", "Материал"),
            ("{thickness}", "Толщина (мм)"),
            ("{quantity}", "Количество"),
            ("{index}", "Порядковый номер"),
            ("{assembly}", "Имя сборки"),
        ]
        
        for i, (var, desc) in enumerate(variables):
            row = i // 2
            col = i % 2
            ttk.Label(
                hint_frame, 
                text=f"{var} - {desc}",
                font=('TkDefaultFont', 9)
            ).grid(row=row, column=col, sticky=tk.W, padx=10, pady=2)
            
        # Предпросмотр
        ttk.Label(frame, text="Предпросмотр:").grid(
            row=3, column=0, sticky=tk.W, pady=(10, 5)
        )
        
        self.lbl_preview = ttk.Label(
            frame, 
            text="",
            font=('TkDefaultFont', 9, 'italic'),
            foreground='gray'
        )
        self.lbl_preview.grid(row=4, column=0, sticky=tk.W)
        
        # Обновление предпросмотра при вводе
        self.var_filename_pattern.trace_add('write', self._update_preview)
        
        frame.columnconfigure(0, weight=1)
        
    def _create_layers_tab(self):
        """Создание вкладки настроек слоёв и линий."""
        frame = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(frame, text="Слои и линии")
        
        # Контур реза
        cut_frame = ttk.LabelFrame(frame, text="Контур реза", padding=10)
        cut_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(cut_frame, text="Имя слоя:").grid(row=0, column=0, sticky=tk.W, padx=5)
        self.var_cut_layer = tk.StringVar()
        ttk.Entry(cut_frame, textvariable=self.var_cut_layer, width=20).grid(
            row=0, column=1, sticky=tk.W, padx=5
        )
        
        ttk.Label(cut_frame, text="Цвет:").grid(row=0, column=2, sticky=tk.W, padx=5)
        self.var_cut_color = tk.StringVar()
        self.cmb_cut_color = ttk.Combobox(
            cut_frame,
            textvariable=self.var_cut_color,
            values=["Красный", "Синий", "Зелёный", "Жёлтый", "Белый", "Чёрный"],
            width=12,
            state='readonly'
        )
        self.cmb_cut_color.grid(row=0, column=3, sticky=tk.W, padx=5)
        
        # Линии сгиба
        bend_frame = ttk.LabelFrame(frame, text="Линии сгиба", padding=10)
        bend_frame.pack(fill=tk.X, pady=5)
        
        self.var_remove_bend_lines = tk.BooleanVar()
        ttk.Checkbutton(
            bend_frame,
            text="Удалять линии сгиба из DXF",
            variable=self.var_remove_bend_lines
        ).grid(row=0, column=0, columnspan=4, sticky=tk.W)
        
        ttk.Label(bend_frame, text="Имя слоя:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=(10, 0))
        self.var_bend_layer = tk.StringVar()
        ttk.Entry(bend_frame, textvariable=self.var_bend_layer, width=20).grid(
            row=1, column=1, sticky=tk.W, padx=5, pady=(10, 0)
        )
        
        ttk.Label(bend_frame, text="Цвет:").grid(row=1, column=2, sticky=tk.W, padx=5, pady=(10, 0))
        self.var_bend_color = tk.StringVar()
        self.cmb_bend_color = ttk.Combobox(
            bend_frame,
            textvariable=self.var_bend_color,
            values=["Красный", "Синий", "Зелёный", "Жёлтый", "Белый", "Чёрный"],
            width=12,
            state='readonly'
        )
        self.cmb_bend_color.grid(row=1, column=3, sticky=tk.W, padx=5, pady=(10, 0))
        
        # Дополнительные настройки DXF
        dxf_frame = ttk.LabelFrame(frame, text="Настройки DXF", padding=10)
        dxf_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(dxf_frame, text="Версия DXF:").grid(row=0, column=0, sticky=tk.W, padx=5)
        self.var_dxf_version = tk.StringVar()
        self.cmb_dxf_version = ttk.Combobox(
            dxf_frame,
            textvariable=self.var_dxf_version,
            values=["AutoCAD 2018", "AutoCAD 2013", "AutoCAD 2010", "AutoCAD 2007", "AutoCAD 2004"],
            width=15,
            state='readonly'
        )
        self.cmb_dxf_version.grid(row=0, column=1, sticky=tk.W, padx=5)
        
    def _browse_output_dir(self):
        """Выбор папки вывода."""
        directory = filedialog.askdirectory(
            title="Выберите папку для сохранения DXF",
            initialdir=self.var_output_dir.get() or None
        )
        if directory:
            self.var_output_dir.set(directory)
            
    def _update_preview(self, *args):
        """Обновление предпросмотра имени файла."""
        pattern = self.var_filename_pattern.get()
        
        # Подстановка тестовых значений
        preview = pattern
        preview = preview.replace("{name}", "Кронштейн")
        preview = preview.replace("{designation}", "СБ.001.002")
        preview = preview.replace("{material}", "Ст3")
        preview = preview.replace("{thickness}", "2.0")
        preview = preview.replace("{quantity}", "4")
        preview = preview.replace("{index}", "01")
        preview = preview.replace("{assembly}", "Корпус")
        
        self.lbl_preview.configure(text=f"{preview}.dxf")
        
    def _load_settings(self):
        """Загрузка настроек в интерфейс."""
        self.var_output_dir.set(str(self.settings.output_dir or ""))
        self.var_create_subfolder.set(self.settings.create_assembly_subfolder)
        self.var_overwrite.set(self.settings.overwrite_existing)
        self.var_open_folder.set(getattr(self.settings, 'open_folder_after', True))
        self.var_filename_pattern.set(self.settings.filename_pattern)
        
        # Слои
        self.var_cut_layer.set(self.settings.cut_contour.layer_name)
        self.var_cut_color.set(self._color_to_name(self.settings.cut_contour.color))
        
        self.var_remove_bend_lines.set(self.settings.remove_bend_lines)
        self.var_bend_layer.set(self.settings.bend_lines.layer_name)
        self.var_bend_color.set(self._color_to_name(self.settings.bend_lines.color))
        
        self.var_dxf_version.set("AutoCAD 2018")
        
    def _save_settings(self):
        """Сохранение настроек из интерфейса."""
        # Обновление настроек
        self.settings.output_directory = self.var_output_dir.get() or ""
        self.settings.create_subdirectories = self.var_create_subfolder.get()
        self.settings.overwrite_existing = self.var_overwrite.get()
        self.settings.filename_settings.template = self.var_filename_pattern.get()
        self.settings.remove_bend_lines = self.var_remove_bend_lines.get()
        
        # Слои
        self.settings.cut_contour.layer_name = self.var_cut_layer.get()
        self.settings.cut_contour.color = self._name_to_color(self.var_cut_color.get())
        
        self.settings.bend_lines.layer_name = self.var_bend_layer.get()
        self.settings.bend_lines.color = self._name_to_color(self.var_bend_color.get())
        
    def _color_to_name(self, color: int) -> str:
        """Преобразование цвета ACI в название."""
        color_map = {
            1: "Красный",
            5: "Синий", 
            3: "Зелёный",
            2: "Жёлтый",
            7: "Белый",
            0: "Чёрный",
        }
        return color_map.get(color, "Красный")
        
    def _name_to_color(self, name: str) -> int:
        """Преобразование названия цвета в ACI."""
        name_map = {
            "Красный": 1,
            "Синий": 5,
            "Зелёный": 3,
            "Жёлтый": 2,
            "Белый": 7,
            "Чёрный": 0,
        }
        return name_map.get(name, 1)
        
    def _on_save(self):
        """Обработчик кнопки сохранения."""
        self._save_settings()
        self.result = self.settings
        
        if self.on_save:
            self.on_save(self.settings)
            
        self.destroy()
        
    def _on_cancel(self):
        """Обработчик кнопки отмены."""
        self.result = None
        self.destroy()
