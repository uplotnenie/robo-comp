"""
DXF-Auto - KOMPAS-3D Sheet Metal DXF Export Tool

Main entry point for the application.
"""

import sys
import os


def main():
    """Main application entry point."""
    # Ensure we're running on Windows (required for COM)
    if sys.platform != 'win32':
        print("Предупреждение: Это приложение требует Windows и установленный КОМПАС-3D.")
        print("КОМПАС-3D COM автоматизация доступна только на Windows.")
        print("Запуск в демо-режиме...")
    
    import tkinter as tk
    from tkinter import messagebox
    
    # Add project root to path
    sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
    
    from config import APP_NAME, APP_VERSION, AppPaths
    from ui.main_window import MainWindow
    
    # Initialize paths
    paths = AppPaths()
    
    # Check pywin32 availability on Windows
    if sys.platform == 'win32':
        try:
            import win32com.client
        except ImportError:
            messagebox.showerror(
                "Ошибка",
                "Модуль pywin32 не установлен.\n"
                "Установите его командой: pip install pywin32"
            )
            return 1
    
    # Create main window
    root = tk.Tk()
    root.title(f"{APP_NAME} v{APP_VERSION}")
    
    # Set window icon (if available)
    icon_path = paths.app_dir / "resources" / "icons" / "app_icon.ico"
    if icon_path.exists():
        try:
            root.iconbitmap(str(icon_path))
        except tk.TclError:
            pass  # Icon format not supported on this platform
    
    # Create and run main window
    app = MainWindow(root)
    
    # Center window on screen
    root.update_idletasks()
    width = root.winfo_width()
    height = root.winfo_height()
    x = (root.winfo_screenwidth() // 2) - (width // 2)
    y = (root.winfo_screenheight() // 2) - (height // 2)
    root.geometry(f'{width}x{height}+{x}+{y}')
    
    # Run main loop
    app.run()
    
    return 0


if __name__ == "__main__":
    sys.exit(main())
