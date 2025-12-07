"""
Дерево состава сборки для отображения иерархии деталей.

Показывает структуру сборки из КОМПАС-3D с возможностью:
- Развёртывания/свёртывания узлов
- Выделения листовых деталей
- Множественного выбора для экспорта
"""

import tkinter as tk
from tkinter import ttk
from typing import Callable, Dict, List, Optional, Set, TYPE_CHECKING

if TYPE_CHECKING:
    from ..models import AssemblyNode, SheetPart


class CompositionTree(ttk.Frame):
    """Виджет дерева состава сборки."""
    
    def __init__(
        self, 
        parent: tk.Widget,
        on_selection_changed: Optional[Callable[[List[str]], None]] = None,
        on_part_double_click: Optional[Callable[[str], None]] = None
    ):
        """
        Инициализация дерева состава.
        
        Args:
            parent: Родительский виджет
            on_selection_changed: Callback при изменении выделения
            on_part_double_click: Callback при двойном клике на деталь
        """
        super().__init__(parent)
        
        self.on_selection_changed = on_selection_changed
        self.on_part_double_click = on_part_double_click
        
        # Хранение данных
        self._node_data: Dict[str, 'AssemblyNode'] = {}
        self._sheet_parts: Dict[str, 'SheetPart'] = {}
        self._checked_items: Set[str] = set()
        
        self._setup_ui()
        self._setup_bindings()
        
    def _setup_ui(self):
        """Настройка интерфейса."""
        # Заголовок
        header_frame = ttk.Frame(self)
        header_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Label(
            header_frame, 
            text="Состав сборки",
            font=('TkDefaultFont', 10, 'bold')
        ).pack(side=tk.LEFT)
        
        # Кнопки управления
        btn_frame = ttk.Frame(header_frame)
        btn_frame.pack(side=tk.RIGHT)
        
        self.btn_expand_all = ttk.Button(
            btn_frame, 
            text="▼", 
            width=3,
            command=self._expand_all
        )
        self.btn_expand_all.pack(side=tk.LEFT, padx=2)
        
        self.btn_collapse_all = ttk.Button(
            btn_frame, 
            text="▲", 
            width=3,
            command=self._collapse_all
        )
        self.btn_collapse_all.pack(side=tk.LEFT, padx=2)
        
        # Дерево
        tree_frame = ttk.Frame(self)
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=(0, 5))
        
        # Стиль для checkboxes
        self.tree = ttk.Treeview(
            tree_frame,
            columns=('type', 'count'),
            show='tree headings',
            selectmode='extended'
        )
        
        # Колонки
        self.tree.heading('#0', text='Наименование', anchor=tk.W)
        self.tree.heading('type', text='Тип', anchor=tk.W)
        self.tree.heading('count', text='Кол-во', anchor=tk.CENTER)
        
        self.tree.column('#0', width=250, minwidth=150)
        self.tree.column('type', width=80, minwidth=60)
        self.tree.column('count', width=50, minwidth=40)
        
        # Скроллбары
        vsb = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self.tree.yview)
        hsb = ttk.Scrollbar(tree_frame, orient=tk.HORIZONTAL, command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        
        # Grid layout
        self.tree.grid(row=0, column=0, sticky='nsew')
        vsb.grid(row=0, column=1, sticky='ns')
        hsb.grid(row=1, column=0, sticky='ew')
        
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)
        
        # Панель выбора
        select_frame = ttk.Frame(self)
        select_frame.pack(fill=tk.X, padx=5, pady=(0, 5))
        
        self.btn_select_all = ttk.Button(
            select_frame, 
            text="Выбрать все листовые",
            command=self._select_all_sheet
        )
        self.btn_select_all.pack(side=tk.LEFT, padx=2)
        
        self.btn_clear_selection = ttk.Button(
            select_frame, 
            text="Снять выделение",
            command=self._clear_selection
        )
        self.btn_clear_selection.pack(side=tk.LEFT, padx=2)
        
        # Счётчик
        self.lbl_count = ttk.Label(select_frame, text="Выбрано: 0")
        self.lbl_count.pack(side=tk.RIGHT, padx=5)
        
        # Теги для стилизации
        self.tree.tag_configure('sheet_part', foreground='#006400')  # Тёмно-зелёный
        self.tree.tag_configure('assembly', foreground='#00008B')     # Тёмно-синий
        self.tree.tag_configure('part', foreground='#333333')         # Тёмно-серый
        self.tree.tag_configure('selected_sheet', background='#90EE90')  # Светло-зелёный
        
    def _setup_bindings(self):
        """Настройка обработчиков событий."""
        self.tree.bind('<<TreeviewSelect>>', self._on_selection_change)
        self.tree.bind('<Double-1>', self._on_double_click)
        self.tree.bind('<space>', self._on_space_press)
        
    def load_assembly(self, root_node: 'AssemblyNode', sheet_parts: Dict[str, 'SheetPart']):
        """
        Загрузка данных сборки в дерево.
        
        Args:
            root_node: Корневой узел дерева сборки
            sheet_parts: Словарь листовых деталей (id -> SheetPart)
        """
        # Очистка
        self.clear()
        
        # Сохранение данных
        self._sheet_parts = sheet_parts
        
        # Рекурсивное добавление узлов
        self._add_node('', root_node)
        
        # Развернуть корневой узел
        for item in self.tree.get_children():
            self.tree.item(item, open=True)
            
    def _add_node(self, parent_id: str, node: 'AssemblyNode') -> str:
        """
        Рекурсивное добавление узла в дерево.
        
        Args:
            parent_id: ID родительского узла
            node: Узел сборки
            
        Returns:
            ID добавленного узла
        """
        # Определение типа и тега
        is_sheet = node.part_id in self._sheet_parts
        
        if node.is_assembly:
            node_type = 'Сборка'
            tag = 'assembly'
        elif is_sheet:
            node_type = 'Листовая'
            tag = 'sheet_part'
        else:
            node_type = 'Деталь'
            tag = 'part'
        
        # Добавление в дерево
        item_id = self.tree.insert(
            parent_id,
            tk.END,
            text=f"{'☐ ' if is_sheet else ''}{node.name}",
            values=(node_type, node.quantity if node.quantity > 1 else ''),
            tags=(tag,),
            open=node.is_assembly
        )
        
        # Сохранение связи
        self._node_data[item_id] = node
        
        # Добавление детей
        for child in node.children:
            self._add_node(item_id, child)
            
        return item_id
        
    def clear(self):
        """Очистка дерева."""
        for item in self.tree.get_children():
            self.tree.delete(item)
        self._node_data.clear()
        self._sheet_parts.clear()
        self._checked_items.clear()
        self._update_count()
        
    def _expand_all(self):
        """Развернуть все узлы."""
        def expand(item):
            self.tree.item(item, open=True)
            for child in self.tree.get_children(item):
                expand(child)
                
        for item in self.tree.get_children():
            expand(item)
            
    def _collapse_all(self):
        """Свернуть все узлы."""
        def collapse(item):
            for child in self.tree.get_children(item):
                collapse(child)
            self.tree.item(item, open=False)
            
        for item in self.tree.get_children():
            collapse(item)
            
    def _select_all_sheet(self):
        """Выбрать все листовые детали."""
        self._checked_items.clear()
        
        def check_sheet(item):
            node = self._node_data.get(item)
            if node and node.part_id in self._sheet_parts:
                self._checked_items.add(item)
                self._update_item_visual(item, True)
            for child in self.tree.get_children(item):
                check_sheet(child)
                
        for item in self.tree.get_children():
            check_sheet(item)
            
        self._update_count()
        self._notify_selection_changed()
        
    def _clear_selection(self):
        """Снять все выделения."""
        for item in self._checked_items:
            self._update_item_visual(item, False)
        self._checked_items.clear()
        self._update_count()
        self._notify_selection_changed()
        
    def _update_item_visual(self, item: str, checked: bool):
        """Обновление визуального состояния элемента."""
        node = self._node_data.get(item)
        if not node:
            return
            
        is_sheet = node.part_id in self._sheet_parts
        current_text = self.tree.item(item, 'text')
        
        if is_sheet:
            # Обновляем чекбокс в тексте
            if checked:
                new_text = current_text.replace('☐ ', '☑ ')
                self.tree.item(item, tags=('sheet_part', 'selected_sheet'))
            else:
                new_text = current_text.replace('☑ ', '☐ ')
                self.tree.item(item, tags=('sheet_part',))
            self.tree.item(item, text=new_text)
            
    def _toggle_item(self, item: str):
        """Переключение состояния выбора элемента."""
        node = self._node_data.get(item)
        if not node or node.part_id not in self._sheet_parts:
            return
            
        if item in self._checked_items:
            self._checked_items.discard(item)
            self._update_item_visual(item, False)
        else:
            self._checked_items.add(item)
            self._update_item_visual(item, True)
            
        self._update_count()
        self._notify_selection_changed()
        
    def _update_count(self):
        """Обновление счётчика выбранных деталей."""
        self.lbl_count.configure(text=f"Выбрано: {len(self._checked_items)}")
        
    def _on_selection_change(self, event):
        """Обработчик изменения выделения в дереве."""
        pass  # Используем checkbox, а не выделение
        
    def _on_double_click(self, event):
        """Обработчик двойного клика."""
        item = self.tree.identify_row(event.y)
        if item:
            # Переключаем выбор для листовой детали
            self._toggle_item(item)
            
            # Callback для открытия свойств
            if self.on_part_double_click:
                node = self._node_data.get(item)
                if node:
                    self.on_part_double_click(node.part_id)
                    
    def _on_space_press(self, event):
        """Обработчик нажатия пробела."""
        for item in self.tree.selection():
            self._toggle_item(item)
            
    def _notify_selection_changed(self):
        """Уведомление об изменении выбора."""
        if self.on_selection_changed:
            # Собираем ID выбранных частей
            part_ids = []
            for item in self._checked_items:
                node = self._node_data.get(item)
                if node:
                    part_ids.append(node.part_id)
            self.on_selection_changed(part_ids)
            
    def get_selected_parts(self) -> List['SheetPart']:
        """
        Получение списка выбранных листовых деталей.
        
        Returns:
            Список выбранных SheetPart
        """
        result = []
        for item in self._checked_items:
            node = self._node_data.get(item)
            if node and node.part_id in self._sheet_parts:
                result.append(self._sheet_parts[node.part_id])
        return result
        
    def get_selected_part_ids(self) -> List[str]:
        """
        Получение ID выбранных деталей.
        
        Returns:
            Список ID деталей
        """
        result = []
        for item in self._checked_items:
            node = self._node_data.get(item)
            if node:
                result.append(node.part_id)
        return result
