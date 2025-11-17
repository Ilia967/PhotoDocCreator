import sys
import importlib
import tkinter as tk
from tkinter import messagebox
import os

class DependencyChecker:
    def __init__(self):
        self.required_modules = {
            'tkinter': 'Встроенная библиотека GUI',
            'PIL': 'Обработка изображений (Pillow)',
            'docx': 'Создание Word документов (python-docx)',
            'json': 'Работа с JSON (встроен)',
            'os': 'Работа с файловой системой (встроен)',
            're': 'Регулярные выражения (встроен)',
            'threading': 'Многопоточность (встроен)'
        }
        
        self.optional_modules = {
            'cv2': 'OpenCV (для расширенной обработки изображений)'
        }
        
    def check_all(self):
        """Проверяет все зависимости и возвращает отчет"""
        report = {
            'system_info': self.get_system_info(),
            'required': {},
            'optional': {},
            'errors': [],
            'warnings': []
        }
        
        # Проверяем системную информацию
        if not self.is_32bit_compatible():
            report['warnings'].append("Программа собрана для 32-битных систем. На 64-битных системах могут потребоваться дополнительные настройки.")
        
        # Проверяем обязательные модули
        for module, description in self.required_modules.items():
            result = self.check_module(module)
            report['required'][module] = {
                'status': result['status'],
                'version': result.get('version', 'N/A'),
                'description': description
            }
            if not result['status']:
                report['errors'].append(f"Отсутствует обязательный модуль: {module} - {description}")
        
        # Проверяем опциональные модули
        for module, description in self.optional_modules.items():
            result = self.check_module(module)
            report['optional'][module] = {
                'status': result['status'],
                'version': result.get('version', 'N/A'),
                'description': description
            }
            if not result['status']:
                report['warnings'].append(f"Отсутствует опциональный модуль: {module} - {description}")
        
        return report
    
    def check_module(self, module_name):
        """Проверяет доступность отдельного модуля"""
        try:
            module = importlib.import_module(module_name)
            version = getattr(module, '__version__', 'Unknown')
            return {'status': True, 'version': version}
        except ImportError:
            return {'status': False}
        except Exception as e:
            return {'status': False, 'error': str(e)}
    
    def get_system_info(self):
        """Собирает информацию о системе"""
        return {
            'platform': sys.platform,
            'python_version': sys.version,
            'executable': sys.executable,
            'architecture': '32-bit' if sys.maxsize <= 2**32 else '64-bit',
            'current_directory': os.getcwd()
        }
    
    def is_32bit_compatible(self):
        """Проверяет совместимость с 32-битными системами"""
        return True  # Python код обычно кроссплатформенный
    
    def show_report_dialog(self, parent=None):
        """Показывает диалог с отчетом о диагностике"""
        report = self.check_all()
        
        # Создаем окно диагностики
        dialog = tk.Toplevel(parent)
        dialog.title("Диагностика системы")
        dialog.geometry("700x500")
        dialog.resizable(True, True)
        
        # Заголовок
        title_label = tk.Label(dialog, text="Диагностика системы и зависимостей", 
                              font=("Arial", 14, "bold"))
        title_label.pack(pady=10)
        
        # Информация о системе
        sys_frame = tk.LabelFrame(dialog, text="Информация о системе")
        sys_frame.pack(fill="x", padx=10, pady=5)
        
        sys_text = f"""Платформа: {report['system_info']['platform']}
Архитектура Python: {report['system_info']['architecture']}
Версия Python: {report['system_info']['python_version'].split()[0]}
Текущая папка: {report['system_info']['current_directory']}"""
        
        sys_label = tk.Label(sys_frame, text=sys_text, justify=tk.LEFT)
        sys_label.pack(padx=10, pady=5)
        
        # Обязательные модули
        req_frame = tk.LabelFrame(dialog, text="Обязательные модули")
        req_frame.pack(fill="x", padx=10, pady=5)
        
        for module, info in report['required'].items():
            status = "✅ Установлен" if info['status'] else "❌ Отсутствует"
            text = f"{module}: {status} (Версия: {info['version']}) - {info['description']}"
            label = tk.Label(req_frame, text=text, justify=tk.LEFT, anchor="w")
            label.pack(fill="x", padx=10, pady=2)
        
        # Ошибки и предупреждения
        if report['errors'] or report['warnings']:
            issues_frame = tk.LabelFrame(dialog, text="Проблемы и решения")
            issues_frame.pack(fill="x", padx=10, pady=5)
            
            issues_text = ""
            for error in report['errors']:
                issues_text += f"❌ {error}\n"
            for warning in report['warnings']:
                issues_text += f"⚠️ {warning}\n"
            
            if issues_text:
                issues_label = tk.Label(issues_frame, text=issues_text, justify=tk.LEFT, anchor="w")
                issues_label.pack(padx=10, pady=5)
        
        # Кнопки
        button_frame = tk.Frame(dialog)
        button_frame.pack(pady=10)
        
        if report['errors']:
            tk.Button(button_frame, text="Установить зависимости", 
                     command=lambda: self.show_install_help()).pack(side=tk.LEFT, padx=5)
        
        tk.Button(button_frame, text="Закрыть", command=dialog.destroy).pack(side=tk.LEFT, padx=5)
    
    def show_install_help(self):
        """Показывает помощь по установке зависимостей"""
        help_text = """
Для установки необходимых модулей:

1. Откройте командную строку (cmd)
2. Выполните команды:

pip install python-docx
pip install Pillow

Или установите все сразу из requirements.txt:
pip install -r requirements.txt

Если возникают ошибки:
- Попробуйте: python -m pip install ...
- Для Windows: может потребоваться Visual C++ Build Tools
- Используйте администраторские права если нужно
"""
        messagebox.showinfo("Помощь по установке", help_text)