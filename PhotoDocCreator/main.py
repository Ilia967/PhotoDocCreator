import tkinter as tk
import sys
import os
import logging

# Настраиваем логирование
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# Добавляем пути к модулям
current_dir = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(current_dir, 'core'))
sys.path.insert(0, os.path.join(current_dir, 'utils'))

def check_dependencies():
    """Проверяет зависимости перед запуском"""
    missing_deps = []
    
    try:
        from docx import Document
    except ImportError:
        missing_deps.append("python-docx")
    
    try:
        from PIL import Image
    except ImportError:
        missing_deps.append("Pillow")
    
    return missing_deps

def main():
    # Проверяем зависимости
    missing = check_dependencies()
    
    if missing:
        root = tk.Tk()
        root.withdraw()
        
        message = "Отсутствуют необходимые зависимости:\n\n"
        for dep in missing:
            message += f"• {dep}\n"
        
        message += "\nУстановите их командами:\n"
        message += "pip install python-docx Pillow\n\n"
        message += "Запустите install_dependencies.bat для автоматической установки."
        
        tk.messagebox.showerror("Ошибка зависимостей", message)
        return
    
    # Импортируем и запускаем основное приложение
    try:
        from core.app import PhotoDocCreator
        
        root = tk.Tk()
        app = PhotoDocCreator(root)
        root.mainloop()
        
    except ImportError as e:
        logger.error(f"Ошибка импорта: {e}")
        tk.messagebox.showerror("Ошибка", f"Не удалось загрузить модули приложения: {e}")
    except Exception as e:
        logger.error(f"Критическая ошибка: {e}")
        tk.messagebox.showerror("Критическая ошибка", f"Произошла ошибка при запуске: {e}")

if __name__ == "__main__":
    main()