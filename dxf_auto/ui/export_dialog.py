"""
Диалог экспорта с прогрессом и отчётом.

Показывает:
- Прогресс экспорта
- Логи операций
- Миниатюры экспортированных файлов
- Итоговый отчёт
"""

import tkinter as tk
from tkinter import ttk
from typing import Optional, List, Dict, Callable, Any
from pathlib import Path
import threading
import queue
import time
from dataclasses import dataclass

from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from ..models import SheetPart, ExportSettings


@dataclass
class ExportResult:
    """Результат экспорта одной детали."""
    part_id: str
    part_name: str
    output_path: str
    success: bool
    error_message: str = ""
    export_time: float = 0.0


class ExportDialog(tk.Toplevel):
    """Диалог экспорта с прогрессом."""
    
    def __init__(
        self,
        parent: tk.Misc,
        parts: List['SheetPart'],
        settings: Optional['ExportSettings'] = None,
        export_function: Optional[Callable] = None
    ):
        """
        Инициализация диалога.
        
        Args:
            parent: Родительское окно
            parts: Список деталей для экспорта
            settings: Настройки экспорта
            export_function: Функция экспорта (async)
        """
        super().__init__(parent)
        
        self.parts = parts
        # Import and create default settings if None
        from ..models import ExportSettings as ES
        self.settings = settings if settings is not None else ES()
        self.export_function = export_function
        
        # Очередь сообщений для потокобезопасной связи
        self.message_queue: queue.Queue = queue.Queue()
        
        # Результаты
        self.results: List[ExportResult] = []
        self.is_running = False
        self.is_cancelled = False
        
        self.title("Экспорт DXF")
        self.geometry("600x500")
        self.resizable(True, True)
        
        # Модальное окно
        if isinstance(parent, (tk.Tk, tk.Toplevel)):
            self.transient(parent)
        self.grab_set()
        
        self._setup_ui()
        
        # Центрирование
        self.update_idletasks()
        x = parent.winfo_rootx() + (parent.winfo_width() - self.winfo_width()) // 2
        y = parent.winfo_rooty() + (parent.winfo_height() - self.winfo_height()) // 2
        self.geometry(f"+{x}+{y}")
        
        # Запрет закрытия во время экспорта
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        
    def _setup_ui(self):
        """Настройка интерфейса."""
        # Заголовок
        header_frame = ttk.Frame(self, padding=10)
        header_frame.pack(fill=tk.X)
        
        self.lbl_status = ttk.Label(
            header_frame,
            text=f"Готово к экспорту: {len(self.parts)} деталей",
            font=('TkDefaultFont', 11, 'bold')
        )
        self.lbl_status.pack(side=tk.LEFT)
        
        # Прогресс
        progress_frame = ttk.Frame(self, padding=(10, 0, 10, 10))
        progress_frame.pack(fill=tk.X)
        
        self.progress = ttk.Progressbar(
            progress_frame,
            mode='determinate',
            maximum=len(self.parts) or 1
        )
        self.progress.pack(fill=tk.X)
        
        self.lbl_progress = ttk.Label(progress_frame, text="0 / 0")
        self.lbl_progress.pack(pady=5)
        
        # Notebook с вкладками Лог / Результаты
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))
        
        # Вкладка логов
        log_frame = ttk.Frame(self.notebook)
        self.notebook.add(log_frame, text="Лог")
        
        # Текстовое поле для логов
        log_scroll_frame = ttk.Frame(log_frame)
        log_scroll_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        self.txt_log = tk.Text(
            log_scroll_frame,
            height=15,
            wrap=tk.WORD,
            state=tk.DISABLED,
            font=('Consolas', 9)
        )
        self.txt_log.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        log_scrollbar = ttk.Scrollbar(log_scroll_frame, command=self.txt_log.yview)
        log_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.txt_log.configure(yscrollcommand=log_scrollbar.set)
        
        # Теги для цветного лога
        self.txt_log.tag_configure('info', foreground='black')
        self.txt_log.tag_configure('success', foreground='green')
        self.txt_log.tag_configure('error', foreground='red')
        self.txt_log.tag_configure('warning', foreground='orange')
        
        # Вкладка результатов
        result_frame = ttk.Frame(self.notebook)
        self.notebook.add(result_frame, text="Результаты")
        
        # Таблица результатов
        columns = ('name', 'file', 'status', 'time')
        self.result_tree = ttk.Treeview(result_frame, columns=columns, show='headings')
        self.result_tree.heading('name', text='Деталь')
        self.result_tree.heading('file', text='Файл')
        self.result_tree.heading('status', text='Статус')
        self.result_tree.heading('time', text='Время')
        
        self.result_tree.column('name', width=150)
        self.result_tree.column('file', width=200)
        self.result_tree.column('status', width=100)
        self.result_tree.column('time', width=60)
        
        result_scrollbar = ttk.Scrollbar(result_frame, command=self.result_tree.yview)
        self.result_tree.configure(yscrollcommand=result_scrollbar.set)
        
        self.result_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5, pady=5)
        result_scrollbar.pack(side=tk.RIGHT, fill=tk.Y, pady=5)
        
        # Теги для результатов
        self.result_tree.tag_configure('success', background='#90EE90')
        self.result_tree.tag_configure('error', background='#FFB6C1')
        
        # Итоги
        summary_frame = ttk.Frame(self, padding=10)
        summary_frame.pack(fill=tk.X)
        
        self.lbl_summary = ttk.Label(summary_frame, text="")
        self.lbl_summary.pack(side=tk.LEFT)
        
        # Кнопки
        btn_frame = ttk.Frame(self, padding=10)
        btn_frame.pack(fill=tk.X)
        
        self.btn_start = ttk.Button(
            btn_frame,
            text="Начать экспорт",
            command=self._start_export
        )
        self.btn_start.pack(side=tk.LEFT, padx=5)
        
        self.btn_cancel = ttk.Button(
            btn_frame,
            text="Отмена",
            command=self._cancel_export
        )
        self.btn_cancel.pack(side=tk.LEFT)
        
        self.btn_close = ttk.Button(
            btn_frame,
            text="Закрыть",
            command=self._on_close,
            state=tk.DISABLED
        )
        self.btn_close.pack(side=tk.RIGHT, padx=5)
        
        self.btn_open_folder = ttk.Button(
            btn_frame,
            text="Открыть папку",
            command=self._open_output_folder,
            state=tk.DISABLED
        )
        self.btn_open_folder.pack(side=tk.RIGHT)
        
    def _log(self, message: str, level: str = 'info'):
        """
        Добавление сообщения в лог.
        
        Args:
            message: Текст сообщения
            level: Уровень (info, success, error, warning)
        """
        self.message_queue.put(('log', message, level))
        
    def _process_queue(self):
        """Обработка очереди сообщений."""
        try:
            while True:
                msg_type, *args = self.message_queue.get_nowait()
                
                if msg_type == 'log':
                    message, level = args
                    self._add_log_message(message, level)
                    
                elif msg_type == 'progress':
                    current, = args
                    self.progress['value'] = current
                    self.lbl_progress.configure(text=f"{current} / {len(self.parts)}")
                    
                elif msg_type == 'result':
                    result, = args
                    self._add_result(result)
                    
                elif msg_type == 'finished':
                    self._on_export_finished()
                    
        except queue.Empty:
            pass
            
        # Продолжаем проверку очереди
        if self.is_running:
            self.after(100, self._process_queue)
            
    def _add_log_message(self, message: str, level: str):
        """Добавление сообщения в текстовое поле лога."""
        timestamp = time.strftime("%H:%M:%S")
        
        self.txt_log.configure(state=tk.NORMAL)
        self.txt_log.insert(tk.END, f"[{timestamp}] {message}\n", level)
        self.txt_log.see(tk.END)
        self.txt_log.configure(state=tk.DISABLED)
        
    def _add_result(self, result: ExportResult):
        """Добавление результата в таблицу."""
        self.results.append(result)
        
        status = "✓ Успешно" if result.success else "✗ Ошибка"
        time_str = f"{result.export_time:.1f}с" if result.export_time else "-"
        tag = 'success' if result.success else 'error'
        
        self.result_tree.insert(
            '',
            tk.END,
            values=(result.part_name, Path(result.output_path).name, status, time_str),
            tags=(tag,)
        )
        
    def _start_export(self):
        """Запуск экспорта."""
        if self.is_running:
            return
            
        self.is_running = True
        self.is_cancelled = False
        
        # Обновление UI
        self.btn_start.configure(state=tk.DISABLED)
        self.lbl_status.configure(text="Выполняется экспорт...")
        
        # Запуск обработки очереди
        self._process_queue()
        
        # Запуск экспорта в отдельном потоке
        thread = threading.Thread(target=self._export_thread, daemon=True)
        thread.start()
        
    def _export_thread(self):
        """Поток экспорта."""
        self._log("Начало экспорта...")
        
        success_count = 0
        error_count = 0
        
        for i, part in enumerate(self.parts):
            if self.is_cancelled:
                self._log("Экспорт отменён пользователем", 'warning')
                break
                
            part_name = part.info.name or part.info.designation or f"Деталь {i+1}"
            self._log(f"Экспорт: {part_name}")
            
            start_time = time.time()
            
            try:
                # Вызов функции экспорта
                if self.export_function:
                    output_path = self.export_function(part, self.settings)
                else:
                    # Демо-режим без реального экспорта
                    output_path = str(Path(self.settings.output_dir or ".") / f"{part_name}.dxf")
                    time.sleep(0.5)  # Имитация работы
                    
                export_time = time.time() - start_time
                
                result = ExportResult(
                    part_id=part.part_id,
                    part_name=part_name,
                    output_path=output_path,
                    success=True,
                    export_time=export_time
                )
                success_count += 1
                self._log(f"  ✓ Сохранено: {Path(output_path).name}", 'success')
                
            except Exception as e:
                export_time = time.time() - start_time
                
                result = ExportResult(
                    part_id=part.part_id,
                    part_name=part_name,
                    output_path="",
                    success=False,
                    error_message=str(e),
                    export_time=export_time
                )
                error_count += 1
                self._log(f"  ✗ Ошибка: {e}", 'error')
                
            self.message_queue.put(('result', result))
            self.message_queue.put(('progress', i + 1))
            
        # Завершение
        self._log(f"\nЭкспорт завершён: {success_count} успешно, {error_count} ошибок")
        self.message_queue.put(('finished',))
        
    def _on_export_finished(self):
        """Обработчик завершения экспорта."""
        self.is_running = False
        
        success_count = sum(1 for r in self.results if r.success)
        error_count = len(self.results) - success_count
        
        self.lbl_status.configure(text="Экспорт завершён")
        self.lbl_summary.configure(
            text=f"Успешно: {success_count} | Ошибок: {error_count} | Всего: {len(self.results)}"
        )
        
        self.btn_start.configure(state=tk.DISABLED)
        self.btn_cancel.configure(state=tk.DISABLED)
        self.btn_close.configure(state=tk.NORMAL)
        self.btn_open_folder.configure(state=tk.NORMAL)
        
        # Переключение на вкладку результатов
        self.notebook.select(1)
        
    def _cancel_export(self):
        """Отмена экспорта."""
        if self.is_running:
            self.is_cancelled = True
            self.lbl_status.configure(text="Отмена...")
            
    def _open_output_folder(self):
        """Открытие папки с результатами."""
        import subprocess
        import platform
        
        output_dir = self.settings.output_dir or "."
        
        if platform.system() == 'Windows':
            subprocess.run(['explorer', output_dir])
        elif platform.system() == 'Darwin':
            subprocess.run(['open', output_dir])
        else:
            subprocess.run(['xdg-open', output_dir])
            
    def _on_close(self):
        """Обработчик закрытия окна."""
        if self.is_running:
            # Запрос подтверждения
            from tkinter import messagebox
            if not messagebox.askyesno(
                "Отмена экспорта",
                "Экспорт выполняется. Отменить?"
            ):
                return
            self._cancel_export()
            
        self.destroy()
        
    def get_results(self) -> List[ExportResult]:
        """Получение результатов экспорта."""
        return self.results
