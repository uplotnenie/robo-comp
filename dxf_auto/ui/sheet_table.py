"""
Таблица листовых деталей с отображением свойств.

Показывает список всех листовых деталей из сборки с их характеристиками:
- Наименование
- Обозначение  
- Материал
- Толщина
- Габариты развёртки
- Количество
"""

import tkinter as tk
from tkinter import ttk
from typing import Callable, Dict, List, Optional, TYPE_CHECKING

if TYPE_CHECKING:
    from ..models import SheetPart, SheetPartInfo


class SheetTable(ttk.Frame):
    """Виджет таблицы листовых деталей."""
    
    # Определение колонок
    COLUMNS = [
        ('name', 'Наименование', 180),
        ('marking', 'Обозначение', 120),
        ('material', 'Материал', 100),
        ('thickness', 'Толщина', 70),
        ('dimensions', 'Габариты', 100),
        ('quantity', 'Кол-во', 50),
        ('status', 'Статус', 80),
    ]
    
    def __init__(
        self,
        parent: tk.Widget,
        on_selection_changed: Optional[Callable[[List[str]], None]] = None,
        on_row_double_click: Optional[Callable[[str], None]] = None
    ):
        """
        Инициализация таблицы.
        
        Args:
            parent: Родительский виджет
            on_selection_changed: Callback при изменении выделения
            on_row_double_click: Callback при двойном клике на строку
        """
        super().__init__(parent)
        
        self.on_selection_changed = on_selection_changed
        self.on_row_double_click = on_row_double_click
        
        # Хранение данных
        self._parts: Dict[str, 'SheetPart'] = {}
        self._item_to_part: Dict[str, str] = {}  # tree item -> part_id
        
        self._setup_ui()
        self._setup_bindings()
        
    def _setup_ui(self):
        """Настройка интерфейса."""
        # Заголовок
        header_frame = ttk.Frame(self)
        header_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Label(
            header_frame,
            text="Листовые детали",
            font=('TkDefaultFont', 10, 'bold')
        ).pack(side=tk.LEFT)
        
        # Кнопки фильтрации
        filter_frame = ttk.Frame(header_frame)
        filter_frame.pack(side=tk.RIGHT)
        
        ttk.Label(filter_frame, text="Толщина:").pack(side=tk.LEFT, padx=(0, 5))
        
        self.thickness_var = tk.StringVar(value="Все")
        self.cmb_thickness = ttk.Combobox(
            filter_frame,
            textvariable=self.thickness_var,
            values=["Все"],
            width=10,
            state='readonly'
        )
        self.cmb_thickness.pack(side=tk.LEFT)
        self.cmb_thickness.bind('<<ComboboxSelected>>', self._on_filter_changed)
        
        # Таблица
        table_frame = ttk.Frame(self)
        table_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=(0, 5))
        
        # Создание Treeview
        column_ids = [col[0] for col in self.COLUMNS]
        self.table = ttk.Treeview(
            table_frame,
            columns=column_ids,
            show='headings',
            selectmode='extended'
        )
        
        # Настройка колонок
        for col_id, col_name, col_width in self.COLUMNS:
            self.table.heading(col_id, text=col_name, anchor=tk.W,
                             command=lambda c=col_id: self._sort_by_column(c))
            self.table.column(col_id, width=col_width, minwidth=50)
            
        # Скроллбары
        vsb = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=self.table.yview)
        hsb = ttk.Scrollbar(table_frame, orient=tk.HORIZONTAL, command=self.table.xview)
        self.table.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        
        # Grid layout
        self.table.grid(row=0, column=0, sticky='nsew')
        vsb.grid(row=0, column=1, sticky='ns')
        hsb.grid(row=1, column=0, sticky='ew')
        
        table_frame.grid_rowconfigure(0, weight=1)
        table_frame.grid_columnconfigure(0, weight=1)
        
        # Панель итогов
        summary_frame = ttk.Frame(self)
        summary_frame.pack(fill=tk.X, padx=5, pady=(0, 5))
        
        self.lbl_total = ttk.Label(summary_frame, text="Всего: 0 деталей")
        self.lbl_total.pack(side=tk.LEFT)
        
        self.lbl_selected = ttk.Label(summary_frame, text="Выбрано: 0")
        self.lbl_selected.pack(side=tk.RIGHT)
        
        # Теги для стилизации
        self.table.tag_configure('exported', background='#90EE90')  # Экспортировано
        self.table.tag_configure('error', background='#FFB6C1')     # Ошибка
        self.table.tag_configure('pending', background='#FFFFFF')   # Ожидает
        
        # Сортировка
        self._sort_column = 'name'
        self._sort_reverse = False
        
    def _setup_bindings(self):
        """Настройка обработчиков событий."""
        self.table.bind('<<TreeviewSelect>>', self._on_selection_change)
        self.table.bind('<Double-1>', self._on_double_click)
        
    def load_parts(self, parts: List['SheetPart']):
        """
        Загрузка списка листовых деталей.
        
        Args:
            parts: Список листовых деталей
        """
        self.clear()
        
        # Сохранение данных
        for part in parts:
            self._parts[part.part_id] = part
            
        # Обновление фильтра толщин
        thicknesses = sorted(set(
            p.info.thickness for p in parts if p.info.thickness
        ))
        self.cmb_thickness['values'] = ["Все"] + [f"{t} мм" for t in thicknesses]
        
        # Добавление строк
        self._populate_table(parts)
        
        # Обновление итогов
        self._update_summary()
        
    def _populate_table(self, parts: List['SheetPart']):
        """Заполнение таблицы данными."""
        for part in parts:
            info = part.info
            
            # Форматирование габаритов
            if info.unfold_width and info.unfold_height:
                dimensions = f"{info.unfold_width:.1f} × {info.unfold_height:.1f}"
            else:
                dimensions = "—"
                
            # Форматирование толщины
            thickness = f"{info.thickness:.2f}" if info.thickness else "—"
            
            # Статус
            status = "Ожидает"
            tag = 'pending'
            
            # Добавление строки
            item = self.table.insert(
                '',
                tk.END,
                values=(
                    info.name or "Без имени",
                    info.marking or "—",
                    info.material or "—",
                    thickness,
                    dimensions,
                    info.quantity if info.quantity > 1 else "1",
                    status
                ),
                tags=(tag,)
            )
            
            self._item_to_part[item] = part.part_id
            
    def clear(self):
        """Очистка таблицы."""
        for item in self.table.get_children():
            self.table.delete(item)
        self._parts.clear()
        self._item_to_part.clear()
        self._update_summary()
        
    def _sort_by_column(self, column: str):
        """Сортировка по колонке."""
        if self._sort_column == column:
            self._sort_reverse = not self._sort_reverse
        else:
            self._sort_column = column
            self._sort_reverse = False
            
        # Получение данных для сортировки
        data = []
        for item in self.table.get_children():
            values = self.table.item(item, 'values')
            tags = self.table.item(item, 'tags')
            data.append((item, values, tags))
            
        # Индекс колонки
        col_index = [col[0] for col in self.COLUMNS].index(column)
        
        # Сортировка
        def sort_key(x):
            value = x[1][col_index]
            # Попытка числовой сортировки
            try:
                return float(value.replace(' мм', '').replace(',', '.'))
            except (ValueError, AttributeError):
                return str(value).lower()
                
        data.sort(key=sort_key, reverse=self._sort_reverse)
        
        # Перестановка элементов
        for index, (item, values, tags) in enumerate(data):
            self.table.move(item, '', index)
            
    def _on_filter_changed(self, event=None):
        """Обработчик изменения фильтра."""
        filter_value = self.thickness_var.get()
        
        # Показать/скрыть строки
        for item in self.table.get_children():
            part_id = self._item_to_part.get(item)
            if part_id is None:
                continue
            part = self._parts.get(part_id)
            
            if not part:
                continue
                
            if filter_value == "Все":
                # Показать все
                pass  # В Treeview нет скрытия, нужно пересоздать
            else:
                # Фильтрация по толщине
                thickness = float(filter_value.replace(' мм', ''))
                if part.info.thickness != thickness:
                    # TODO: реализовать фильтрацию через пересоздание таблицы
                    pass
                    
        # Пересоздание таблицы с фильтром
        self._apply_filter()
        
    def _apply_filter(self):
        """Применение фильтра к таблице."""
        filter_value = self.thickness_var.get()
        
        # Очистка таблицы (без очистки данных)
        for item in self.table.get_children():
            self.table.delete(item)
        self._item_to_part.clear()
        
        # Фильтрация деталей
        filtered_parts = []
        for part in self._parts.values():
            if filter_value == "Все":
                filtered_parts.append(part)
            else:
                thickness = float(filter_value.replace(' мм', ''))
                if part.info.thickness == thickness:
                    filtered_parts.append(part)
                    
        # Заполнение таблицы
        self._populate_table(filtered_parts)
        self._update_summary()
        
    def _update_summary(self):
        """Обновление итоговой информации."""
        total = len(self.table.get_children())
        selected = len(self.table.selection())
        
        self.lbl_total.configure(text=f"Всего: {total} деталей")
        self.lbl_selected.configure(text=f"Выбрано: {selected}")
        
    def _on_selection_change(self, event):
        """Обработчик изменения выделения."""
        self._update_summary()
        
        if self.on_selection_changed:
            part_ids = self.get_selected_part_ids()
            self.on_selection_changed(part_ids)
            
    def _on_double_click(self, event):
        """Обработчик двойного клика."""
        item = self.table.identify_row(event.y)
        if item and self.on_row_double_click:
            part_id = self._item_to_part.get(item)
            if part_id:
                self.on_row_double_click(part_id)
                
    def get_selected_parts(self) -> List['SheetPart']:
        """
        Получение списка выбранных деталей.
        
        Returns:
            Список выбранных SheetPart
        """
        result = []
        for item in self.table.selection():
            part_id = self._item_to_part.get(item)
            if part_id and part_id in self._parts:
                result.append(self._parts[part_id])
        return result
        
    def get_selected_part_ids(self) -> List[str]:
        """
        Получение ID выбранных деталей.
        
        Returns:
            Список ID деталей
        """
        result = []
        for item in self.table.selection():
            part_id = self._item_to_part.get(item)
            if part_id:
                result.append(part_id)
        return result
        
    def get_all_parts(self) -> List['SheetPart']:
        """
        Получение всех деталей.
        
        Returns:
            Список всех SheetPart
        """
        return list(self._parts.values())
        
    def update_part_status(self, part_id: str, status: str, is_error: bool = False):
        """
        Обновление статуса детали.
        
        Args:
            part_id: ID детали
            status: Текст статуса
            is_error: Признак ошибки
        """
        # Поиск элемента по part_id
        for item, pid in self._item_to_part.items():
            if pid == part_id:
                # Обновление значения статуса
                values = list(self.table.item(item, 'values'))
                values[-1] = status  # Последняя колонка - статус
                
                # Обновление тега
                tag = 'error' if is_error else 'exported'
                
                self.table.item(item, values=values, tags=(tag,))
                break
                
    def select_all(self):
        """Выбрать все строки."""
        self.table.selection_set(self.table.get_children())
        self._update_summary()
        
    def clear_selection(self):
        """Снять выделение."""
        self.table.selection_remove(self.table.get_children())
        self._update_summary()
