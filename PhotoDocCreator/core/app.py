import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import os
import sys
import json
import threading
import logging
from PIL import Image, ImageTk

# Импортируем наши модули
from .diagnostics import DependencyChecker
from .image_sorter import VisualImageSorter
from .advanced_sorter import AdvancedImageSorter
from .doc_creator import DocumentCreator
from utils.config_manager import ConfigManager
from utils.file_utils import natural_sort_key, get_image_files

logger = logging.getLogger(__name__)

class PhotoDocCreator:
    def __init__(self, root):
        self.root = root
        self.root.title("PhotoDoc Creator v4.5")
        self.root.geometry("900x700")
        self.root.minsize(800, 600)
        
        # Инициализация менеджера конфигурации
        self.config_manager = ConfigManager()
        self.config = self.config_manager.load_config()
        
        # Инициализация переменных
        self._setup_variables()
        self._setup_ui()
        
        # Загрузка пресетов
        self.presets = {}
        self.current_preset = tk.StringVar(value="По умолчанию")
        self.load_presets()
        
        # Переменные для расширенной сортировки
        self.advanced_sort_order = []
        self.rotation_info = {}
        
        logger.info("PhotoDoc Creator запущен")
    
    def _setup_variables(self):
        """Инициализирует переменные из конфигурации"""
        # Основные настройки
        self.screenshots_folder = tk.StringVar(value=self.config.get('screenshots_folder', ''))
        self.word_file = tk.StringVar(value=self.config.get('word_file', ''))
        self.image_width = tk.DoubleVar(value=self.config.get('image_width', 6.0))
        self.image_height = tk.DoubleVar(value=self.config.get('image_height', 9.0))
        self.images_per_page = tk.IntVar(value=self.config.get('images_per_page', 2))
        
        # Данные сотрудника
        self.officer_name = tk.StringVar(value=self.config.get('officer_name', 'ФИО'))
        self.officer_rank = tk.StringVar(value=self.config.get('officer_rank', 'Звание'))
        self.officer_position = tk.StringVar(value=self.config.get('officer_position', 'Должность'))
        self.department_name = tk.StringVar(value=self.config.get('department_name', 'Подразделение'))
        self.photo_table_title = tk.StringVar(value=self.config.get('photo_table_title', 'ФОТОТАБЛИЦА\nк протоколу осмотра предметов от __.__.____'))
        
        # Настройки шрифта
        self.font_family = tk.StringVar(value=self.config.get('font_family', 'Times New Roman'))
        self.font_size = tk.IntVar(value=self.config.get('font_size', 12))
        self.font_bold = tk.BooleanVar(value=self.config.get('font_bold', False))
        
        # Сортировка
        self.sort_method = tk.StringVar(value=self.config.get('sort_method', 'name_asc'))
        self.manual_sort_order = self.config.get('manual_sort_order', [])
        self.caption_rules = self.config.get('caption_rules', [])
        
        # Колонтитул
        self.footer_department = tk.StringVar(value=self.config.get('footer_department', 'Подразделение'))
        self.enable_footer = tk.BooleanVar(value=self.config.get('enable_footer', True))
        
        # Многопапковый режим
        self.multi_folder_mode = tk.BooleanVar(value=self.config.get('multi_folder_mode', False))
        self.multi_folder_sort_method = tk.StringVar(value=self.config.get('multi_folder_sort_method', 'name_asc'))
        self.folder_sequence = self.config.get('folder_sequence', [])
    
    def _setup_ui(self):
        """Настраивает пользовательский интерфейс"""
        # Создаем меню
        self._create_menu()
        
        # Создаем notebook для вкладок
        notebook = ttk.Notebook(self.root)
        notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Создаем вкладки
        main_frame = ttk.Frame(notebook)
        settings_frame = ttk.Frame(notebook)
        multi_folder_frame = ttk.Frame(notebook)
        caption_rules_frame = ttk.Frame(notebook)
        
        notebook.add(main_frame, text="📷 Основные настройки")
        notebook.add(settings_frame, text="⚙ Дополнительные настройки")
        notebook.add(multi_folder_frame, text="📁 Многопапковый режим")
        notebook.add(caption_rules_frame, text="📝 Правила подписей")
        
        self.setup_main_tab(main_frame)
        self.setup_settings_tab(settings_frame)
        self.setup_multi_folder_tab(multi_folder_frame)
        self.setup_caption_rules_tab(caption_rules_frame)
        
        # Обновляем информацию о режиме
        self.on_mode_changed()
    
    def _create_menu(self):
        """Создает главное меню"""
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)
        
        # Меню "Файл"
        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Файл", menu=file_menu)
        file_menu.add_command(label="Диагностика системы", command=self.show_diagnostics)
        file_menu.add_separator()
        file_menu.add_command(label="Выход", command=self.root.quit)
        
        # Меню "Справка"
        help_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Справка", menu=help_menu)
        help_menu.add_command(label="О программе", command=self.show_about)
    
    def setup_main_tab(self, parent):
        """Настраивает основную вкладку"""
        # Заголовок документа
        ttk.Label(parent, text="Заголовок документа:", font=("Arial", 11, "bold")).grid(row=0, column=0, sticky="w", padx=5, pady=5)
        department_entry = ttk.Entry(parent, textvariable=self.department_name, width=80)
        department_entry.grid(row=0, column=1, columnspan=2, sticky="we", padx=5, pady=5)
        
        ttk.Label(parent, text="Название фототаблицы:", font=("Arial", 11, "bold")).grid(row=1, column=0, sticky="w", padx=5, pady=5)
        title_entry = ttk.Entry(parent, textvariable=self.photo_table_title, width=80)
        title_entry.grid(row=1, column=1, columnspan=2, sticky="we", padx=5, pady=5)
        
        # Разделитель
        separator = ttk.Separator(parent, orient='horizontal')
        separator.grid(row=2, column=0, columnspan=3, sticky="we", padx=5, pady=10)
        
        # Пути к файлам
        ttk.Label(parent, text="Папка с фотографиями:").grid(row=3, column=0, sticky="w", padx=5, pady=5)
        folder_entry = ttk.Entry(parent, textvariable=self.screenshots_folder, width=60)
        folder_entry.grid(row=3, column=1, padx=5, pady=5)
        ttk.Button(parent, text="Обзор", command=self.browse_screenshots_folder).grid(row=3, column=2, padx=5, pady=5)
        
        ttk.Label(parent, text="Файл для сохранения:").grid(row=4, column=0, sticky="w", padx=5, pady=5)
        file_entry = ttk.Entry(parent, textvariable=self.word_file, width=60)
        file_entry.grid(row=4, column=1, padx=5, pady=5)
        ttk.Button(parent, text="Обзор", command=self.browse_word_file).grid(row=4, column=2, padx=5, pady=5)
        
        # Сортировка
        ttk.Label(parent, text="Сортировка фото:").grid(row=5, column=0, sticky="w", padx=5, pady=5)
        sort_frame = ttk.Frame(parent)
        sort_frame.grid(row=5, column=1, columnspan=2, sticky="w", padx=5, pady=5)
        
        sort_combo = ttk.Combobox(sort_frame, textvariable=self.sort_method, width=25, state="readonly")
        sort_combo['values'] = (
            'name_asc', 
            'name_desc', 
            'date_asc', 
            'date_desc',
            'manual'
        )
        sort_combo.pack(side=tk.LEFT, padx=5)
        ttk.Button(sort_frame, text="Визуальная сортировка", command=self.visual_sort_images).pack(side=tk.LEFT, padx=5)
        
        # Кнопки
        button_frame = ttk.Frame(parent)
        button_frame.grid(row=6, column=0, columnspan=3, pady=20)
        
        ttk.Button(button_frame, text="Создать документ", command=self.start_creation_process).pack(side=tk.LEFT, padx=10)
        ttk.Button(button_frame, text="Диагностика системы", command=self.show_diagnostics).pack(side=tk.LEFT, padx=10)
        
        # Логи
        self.log_text = scrolledtext.ScrolledText(parent, height=15, width=90)
        self.log_text.grid(row=7, column=0, columnspan=3, padx=5, pady=5, sticky="nsew")
        
        # Настройка растягивания
        parent.grid_rowconfigure(7, weight=1)
        parent.grid_columnconfigure(1, weight=1)
    
    def setup_settings_tab(self, parent):
        """Настраивает вкладку дополнительных настроек"""
        # Настройки размеров
        size_frame = ttk.LabelFrame(parent, text="Размеры фотографий")
        size_frame.grid(row=0, column=0, columnspan=2, sticky="we", padx=5, pady=5)
        
        ttk.Label(size_frame, text="Ширина (см):").grid(row=0, column=0, padx=5, pady=5)
        ttk.Spinbox(size_frame, from_=1, to=20, width=8, textvariable=self.image_width, increment=0.5).grid(row=0, column=1, padx=5, pady=5)
        
        ttk.Label(size_frame, text="Высота (см):").grid(row=0, column=2, padx=5, pady=5)
        ttk.Spinbox(size_frame, from_=1, to=20, width=8, textvariable=self.image_height, increment=0.5).grid(row=0, column=3, padx=5, pady=5)
        
        ttk.Label(size_frame, text="Фото на страницу:").grid(row=0, column=4, padx=5, pady=5)
        ttk.Spinbox(size_frame, from_=1, to=4, width=8, textvariable=self.images_per_page).grid(row=0, column=5, padx=5, pady=5)
        
        # Настройки шрифта
        font_frame = ttk.LabelFrame(parent, text="Настройки шрифта")
        font_frame.grid(row=1, column=0, columnspan=2, sticky="we", padx=5, pady=5)
        
        ttk.Label(font_frame, text="Шрифт:").grid(row=0, column=0, padx=5, pady=5)
        font_combo = ttk.Combobox(font_frame, textvariable=self.font_family, width=15, state="readonly")
        font_combo['values'] = ('Times New Roman', 'Arial', 'Calibri')
        font_combo.grid(row=0, column=1, padx=5, pady=5)
        
        ttk.Label(font_frame, text="Размер:").grid(row=0, column=2, padx=5, pady=5)
        ttk.Spinbox(font_frame, from_=8, to=24, width=5, textvariable=self.font_size).grid(row=0, column=3, padx=5, pady=5)
        
        ttk.Checkbutton(font_frame, text="Жирный", variable=self.font_bold).grid(row=0, column=4, padx=5, pady=5)
        
        # Данные сотрудника
        officer_frame = ttk.LabelFrame(parent, text="Данные сотрудника")
        officer_frame.grid(row=2, column=0, columnspan=2, sticky="we", padx=5, pady=5)
        
        ttk.Label(officer_frame, text="Должность:").grid(row=0, column=0, padx=5, pady=5)
        ttk.Entry(officer_frame, textvariable=self.officer_position, width=20).grid(row=0, column=1, padx=5, pady=5)
        
        ttk.Label(officer_frame, text="Звание:").grid(row=0, column=2, padx=5, pady=5)
        ttk.Entry(officer_frame, textvariable=self.officer_rank, width=15).grid(row=0, column=3, padx=5, pady=5)
        
        ttk.Label(officer_frame, text="ФИО:").grid(row=0, column=4, padx=5, pady=5)
        ttk.Entry(officer_frame, textvariable=self.officer_name, width=20).grid(row=0, column=5, padx=5, pady=5)
        
        # Колонтитул
        footer_frame = ttk.LabelFrame(parent, text="Настройки колонтитула")
        footer_frame.grid(row=3, column=0, columnspan=2, sticky="we", padx=5, pady=5)
        
        ttk.Label(footer_frame, text="Подразделение:").grid(row=0, column=0, sticky="w", padx=5, pady=5)
        ttk.Entry(footer_frame, textvariable=self.footer_department, width=60).grid(row=0, column=1, padx=5, pady=5, columnspan=2, sticky="we")
        
        ttk.Checkbutton(footer_frame, text="Включить нижний колонтитул", variable=self.enable_footer).grid(row=1, column=0, sticky="w", padx=5, pady=5)
        
        # Настройка растягивания
        parent.grid_columnconfigure(1, weight=1)
    
    def browse_screenshots_folder(self):
        folder = filedialog.askdirectory(title="Выберите папку с фотографиями")
        if folder:
            self.screenshots_folder.set(folder)
    
    def browse_word_file(self):
        file = filedialog.asksaveasfilename(
            title="Выберите файл для сохранения",
            defaultextension=".docx",
            filetypes=[("Word documents", "*.docx")]
        )
        if file:
            self.word_file.set(file)
    
    def show_diagnostics(self):
        """Показывает окно диагностики"""
        checker = DependencyChecker()
        checker.show_report_dialog(self.root)
    
    def show_about(self):
        """Показывает окно 'О программе'"""
        about_text = """PhotoDoc Creator v4.5

Программа для создания фототаблиц к протоколам осмотра.

Возможности:
• Создание документов Word с фотографиями
• Гибкая настройка подписей
• Визуальная сортировка изображений
• Расширенная сортировка для многопапкового режима
• Настройка формата и стилей

Разработчик: Ilia967"""
        messagebox.showinfo("О программе", about_text)
    
    def log(self, message):
        """Добавляет сообщение в лог"""
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.root.update_idletasks()
        logger.info(message)
    
    def save_config(self):
        """Сохраняет текущую конфигурацию"""
        # Обновляем конфиг из переменных
        self._update_config_from_variables()
        
        # Сохраняем
        self.config_manager.save_config(self.config)
    
    def _update_config_from_variables(self):
        """Обновляет конфиг из текущих переменных"""
        self.config.update({
            'screenshots_folder': self.screenshots_folder.get(),
            'word_file': self.word_file.get(),
            'image_width': self.image_width.get(),
            'image_height': self.image_height.get(),
            'images_per_page': self.images_per_page.get(),
            'officer_name': self.officer_name.get(),
            'officer_rank': self.officer_rank.get(),
            'officer_position': self.officer_position.get(),
            'department_name': self.department_name.get(),
            'photo_table_title': self.photo_table_title.get(),
            'font_family': self.font_family.get(),
            'font_size': self.font_size.get(),
            'font_bold': self.font_bold.get(),
            'sort_method': self.sort_method.get(),
            'manual_sort_order': self.manual_sort_order,
            'caption_rules': self.caption_rules,
            'footer_department': self.footer_department.get(),
            'enable_footer': self.enable_footer.get(),
            'multi_folder_mode': self.multi_folder_mode.get(),
            'multi_folder_sort_method': self.multi_folder_sort_method.get(),
            'folder_sequence': self.folder_sequence
        })
    
    def load_presets(self):
        """Загружает пресеты"""
        # Базовая реализация - можно расширить
        pass
    
    def start_creation_process(self):
        """Запускает процесс создания документа"""
        if not self.screenshots_folder.get():
            messagebox.showerror("Ошибка", "Выберите папку с фотографиями")
            return
            
        if not self.word_file.get():
            messagebox.showerror("Ошибка", "Укажите файл для сохранения")
            return
        
        self.save_config()
        self.log_text.delete(1.0, tk.END)
        
        # Запускаем в отдельном потоке
        thread = threading.Thread(target=self.create_document)
        thread.daemon = True
        thread.start()
    
    def create_document(self):
        """Создает документ"""
        try:
            self.log("🚀 Начало создания документа...")
            
            # Получаем изображения в зависимости от режима
            if self.multi_folder_mode.get() and self.folder_sequence:
                self.log("🔀 Режим: Многопапковый")
                
                # Проверяем, есть ли расширенный порядок сортировки
                if hasattr(self, 'advanced_sort_order') and self.advanced_sort_order:
                    self.log("🔀 Используется расширенный порядок сортировки")
                    image_data_list = self.get_images_from_advanced_sort()
                else:
                    image_data_list = self.get_all_images_multi_folder()
                
                image_files_count = len(image_data_list)
                self.log(f"📁 Обрабатывается {len(self.folder_sequence)} папок, всего {image_files_count} фото")
            else:
                self.log("🔀 Режим: Одна папка")
                folder = self.screenshots_folder.get()
                if not folder or not os.path.exists(folder):
                    self.log("❌ Папка с фотографиями не существует")
                    return
                    
                image_files = get_image_files(folder)
                
                if not image_files:
                    self.log("❌ В папке нет изображений")
                    return
                
                # Сортируем изображения
                if self.sort_method.get() == "name_asc":
                    image_files.sort(key=natural_sort_key)
                elif self.sort_method.get() == "name_desc":
                    image_files.sort(key=natural_sort_key, reverse=True)
                elif self.sort_method.get() == "manual" and self.manual_sort_order:
                    # Ручная сортировка
                    manual_files = [f for f in self.manual_sort_order if f in image_files]
                    remaining_files = [f for f in image_files if f not in manual_files]
                    image_files = manual_files + remaining_files
                else:
                    image_files.sort(key=natural_sort_key)  # По умолчанию
                
                self.log(f"📁 Найдено {len(image_files)} изображений")
                
                # Подготавливаем данные изображений
                image_data_list = []
                for i, img_file in enumerate(image_files, 1):
                    image_data_list.append({
                        'path': os.path.join(folder, img_file),
                        'filename': img_file,
                        'global_number': i,
                        'folder_rules': [],
                        'folder_start_number': 1
                    })
                image_files_count = len(image_files)
            
            if not image_data_list:
                self.log("❌ Нет изображений для обработки")
                return
            
            # Создаем конфиг для DocumentCreator
            config = {
                'word_file': self.word_file.get(),
                'image_width': self.image_width.get(),
                'image_height': self.image_height.get(),
                'images_per_page': self.images_per_page.get(),
                'department_name': self.department_name.get(),
                'photo_table_title': self.photo_table_title.get(),
                'font_family': self.font_family.get(),
                'font_size': self.font_size.get(),
                'font_bold': self.font_bold.get(),
                'officer_position': self.officer_position.get(),
                'footer_department': self.footer_department.get(),
                'officer_rank': self.officer_rank.get(),
                'officer_name': self.officer_name.get(),
                'enable_footer': self.enable_footer.get(),
                'caption_rules': self.caption_rules,
                'multi_folder_mode': self.multi_folder_mode.get(),
                'rotation_info': getattr(self, 'rotation_info', {})  # Добавляем информацию о поворотах
            }
            
            # Создаем документ
            doc_creator = DocumentCreator(config)
            success, result, count = doc_creator.create_document(image_data_list, self.log)
            
            if success:
                # Пытаемся открыть документ
                try:
                    os.startfile(result)
                    self.log("🔓 Документ открыт в Word")
                except:
                    self.log("ℹ️ Файл создан, но не удалось открыть его автоматически")
                    
        except Exception as e:
            self.log(f"❌ Критическая ошибка: {str(e)}")
            logger.error(f"Ошибка при создании документа: {e}")
    
    def visual_sort_images(self):
        """Запускает визуальную сортировку изображений"""
        folder = self.screenshots_folder.get()
        if not folder or not os.path.exists(folder):
            messagebox.showerror("Ошибка", "Сначала выберите папку с фотографиями")
            return
        
        try:
            sorter = VisualImageSorter(self.root, folder)
            new_order = sorter.sort_images()
            
            if new_order:
                self.manual_sort_order = new_order
                self.sort_method.set("manual")
                self.log("✅ Визуальный порядок фотографий сохранен")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось запустить визуальную сортировку: {e}")
            logger.error(f"Ошибка визуальной сортировки: {e}")
    
    def advanced_visual_sort(self):
        """Запускает расширенную визуальную сортировку для многопапкового режима"""
        if not self.folder_sequence:
            messagebox.showwarning("Внимание", "Добавьте папки в многопапковом режиме")
            return
        
        try:
            sorter = AdvancedImageSorter(self.root, self.folder_sequence)
            new_order = sorter.sort_images()
            
            if new_order:
                # Сохраняем расширенный порядок и информацию о поворотах
                self.advanced_sort_order = new_order
                self.rotation_info = getattr(sorter, 'rotation_info', {})
                self.log("✅ Расширенный порядок сортировки сохранен")
                self.log(f"💾 Сохранено {len(self.rotation_info)} поворотов изображений")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось запустить расширенную сортировку: {e}")
            logger.error(f"Ошибка расширенной сортировки: {e}")
    
    def get_images_from_advanced_sort(self):
        """Получает изображения из расширенного порядка сортировки"""
        image_data_list = []
        
        for i, filename in enumerate(self.advanced_sort_order, 1):
            # Находим полный путь к файлу
            full_path = None
            for folder_data in self.folder_sequence:
                folder_path = folder_data['path']
                potential_path = os.path.join(folder_path, filename)
                if os.path.exists(potential_path):
                    full_path = potential_path
                    break
            
            if full_path and os.path.exists(full_path):
                image_data_list.append({
                    'path': full_path,
                    'filename': filename,
                    'global_number': i,
                    'folder_rules': [],
                    'folder_start_number': 1,
                    'rotation': self.rotation_info.get(full_path, 0)
                })
        
        return image_data_list
    
    def load_selected_preset(self):
        """Заглушка для загрузки пресета"""
        messagebox.showinfo("Информация", "Функция пресетов в данной версии недоступна")
    
    def setup_multi_folder_tab(self, parent):
        """Настраивает вкладку многопапкового режима"""
        ttk.Label(parent, text="Режим работы с несколькими папками", font=("Arial", 12, "bold")).pack(pady=10)
        
        instruction = """📌 В этом режиме вы можете добавить несколько папок с фотографиями.
Каждая папка будет обрабатываться в порядке добавления."""
        ttk.Label(parent, text=instruction, wraplength=800, justify=tk.LEFT).pack(pady=5, padx=10)
        
        # Фрейм для управления папками
        folders_frame = ttk.LabelFrame(parent, text="Управление папками")
        folders_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Список папок
        self.folders_tree = ttk.Treeview(folders_frame, columns=("order", "path", "images", "rules"), show="headings", height=8)
        self.folders_tree.heading("order", text="№")
        self.folders_tree.heading("path", text="Путь к папке")
        self.folders_tree.heading("images", text="Фото")
        self.folders_tree.heading("rules", text="Правил")
        
        self.folders_tree.column("order", width=50)
        self.folders_tree.column("path", width=400)
        self.folders_tree.column("images", width=80)
        self.folders_tree.column("rules", width=80)
        
        self.folders_tree.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Кнопки управления папками
        folder_btn_frame = ttk.Frame(folders_frame)
        folder_btn_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Button(folder_btn_frame, text="Добавить папку", command=self.add_multi_folder).pack(side=tk.LEFT, padx=5)
        ttk.Button(folder_btn_frame, text="Удалить папку", command=self.remove_multi_folder).pack(side=tk.LEFT, padx=5)
        ttk.Button(folder_btn_frame, text="Правила подписей", command=self.edit_folder_rules).pack(side=tk.LEFT, padx=5)
        ttk.Button(folder_btn_frame, text="Переместить вверх", command=self.move_folder_up).pack(side=tk.LEFT, padx=5)
        ttk.Button(folder_btn_frame, text="Переместить вниз", command=self.move_folder_down).pack(side=tk.LEFT, padx=5)
        ttk.Button(folder_btn_frame, text="Расширенная сортировка", command=self.advanced_visual_sort).pack(side=tk.LEFT, padx=5)
        
        # Сортировка для многопапкового режима
        sort_frame = ttk.LabelFrame(parent, text="Сортировка фотографий в папках")
        sort_frame.pack(fill=tk.X, padx=10, pady=10)
        
        ttk.Label(sort_frame, text="Порядок сортировки:").grid(row=0, column=0, sticky="w", padx=5, pady=5)
        
        multi_sort_combo = ttk.Combobox(sort_frame, textvariable=self.multi_folder_sort_method, width=25, state="readonly")
        multi_sort_combo['values'] = (
            'name_asc', 
            'name_desc', 
            'date_asc', 
            'date_desc'
        )
        multi_sort_combo.grid(row=0, column=1, padx=5, pady=5, sticky="w")
        ttk.Button(sort_frame, text="Применить ко всем папкам", command=self.apply_sort_to_all_folders).grid(row=0, column=2, padx=5, pady=5)
        
        # Переключатель режима
        mode_frame = ttk.Frame(parent)
        mode_frame.pack(fill=tk.X, padx=10, pady=10)
        
        ttk.Label(mode_frame, text="Режим работы:").pack(side=tk.LEFT)
        ttk.Radiobutton(mode_frame, text="Одна папка", variable=self.multi_folder_mode, value=False, 
                       command=self.on_mode_changed).pack(side=tk.LEFT, padx=10)
        ttk.Radiobutton(mode_frame, text="Несколько папок", variable=self.multi_folder_mode, value=True,
                       command=self.on_mode_changed).pack(side=tk.LEFT, padx=10)
        
        # Информация о выбранном режиме
        self.mode_info = ttk.Label(parent, text="", foreground="orange")
        self.mode_info.pack(pady=5)
        
        # Обновляем отображение
        self.update_folders_tree()
        self.on_mode_changed()
    
    def edit_folder_rules(self):
        """Редактирование правил подписей для выбранной папки"""
        selected = self.folders_tree.selection()
        if not selected:
            messagebox.showwarning("Внимание", "Выберите папку для редактирования правил")
            return
            
        index = self.folders_tree.index(selected[0])
        if 0 <= index < len(self.folder_sequence):
            self.current_editing_folder = index
            self.open_folder_rules_editor()
    
    def open_folder_rules_editor(self):
        """Открывает редактор правил для конкретной папки"""
        folder_data = self.folder_sequence[self.current_editing_folder]
        
        rules_window = tk.Toplevel(self.root)
        rules_window.title(f"Правила подписей для папки: {os.path.basename(folder_data['path'])}")
        rules_window.geometry("800x600")
        rules_window.transient(self.root)
        rules_window.grab_set()
        
        # Заголовок
        ttk.Label(rules_window, text=f"Настройка правил подписей для фотографий из папки:", 
                 font=("Arial", 11, "bold")).pack(pady=10)
        ttk.Label(rules_window, text=folder_data['path'], foreground="blue").pack(pady=5)
        
        # Текущие правила
        rules_frame = ttk.LabelFrame(rules_window, text="Текущие правила")
        rules_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        # Создаем Treeview для отображения правил
        columns = ("start", "end", "text")
        rules_tree = ttk.Treeview(rules_frame, columns=columns, show="headings", height=6)
        
        rules_tree.heading("start", text="С фото №")
        rules_tree.heading("end", text="По фото №")
        rules_tree.heading("text", text="Текст подписи")
        
        rules_tree.column("start", width=80)
        rules_tree.column("end", width=80)
        rules_tree.column("text", width=500)
        
        rules_tree.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        def update_rules_display():
            """Обновляет отображение правил в дереве"""
            rules_tree.delete(*rules_tree.get_children())
            for rule in folder_data['caption_rules']:
                if len(rule) >= 3:  # Проверяем, что правило имеет все необходимые элементы
                    start, end, text = rule[0], rule[1], rule[2]
                    rules_tree.insert("", tk.END, values=(start, end, text))
        
        # Изначальное заполнение
        update_rules_display()
        
        # Фрейм для добавления новых правил
        add_frame = ttk.Frame(rules_window)
        add_frame.pack(fill=tk.X, padx=10, pady=5)
        
        ttk.Label(add_frame, text="С фото №:").pack(side=tk.LEFT)
        start_spin = ttk.Spinbox(add_frame, from_=1, to=1000, width=8)
        start_spin.pack(side=tk.LEFT, padx=2)
        
        ttk.Label(add_frame, text="По фото №:").pack(side=tk.LEFT, padx=(10, 2))
        end_spin = ttk.Spinbox(add_frame, from_=1, to=1000, width=8)
        end_spin.pack(side=tk.LEFT, padx=2)
        
        ttk.Label(add_frame, text="Текст подписи:").pack(side=tk.LEFT, padx=(10, 2))
        text_entry = ttk.Entry(add_frame, width=30)
        text_entry.pack(side=tk.LEFT, padx=2)
        
        def add_rule():
            """Добавление нового правила"""
            try:
                start = int(start_spin.get())
                end = int(end_spin.get())
                text = text_entry.get().strip()
                
                if start > end:
                    messagebox.showerror("Ошибка", "Начальный номер не может быть больше конечного")
                    return
                    
                if not text:
                    messagebox.showerror("Ошибка", "Введите текст подписи")
                    return
                
                # Проверяем пересечения с существующими правилами
                for rule in folder_data['caption_rules']:
                    if len(rule) >= 3:
                        rule_start, rule_end, _ = rule[0], rule[1], rule[2]
                        if not (end < rule_start or start > rule_end):
                            messagebox.showerror("Ошибка", f"Диапазон пересекается с существующим правилом: фото {rule_start}-{rule_end}")
                            return
                
                # Добавляем правило
                folder_data['caption_rules'].append((start, end, text))
                folder_data['caption_rules'].sort(key=lambda x: x[0])  # Сортируем по начальному номеру
                
                # Обновляем отображение
                update_rules_display()
                self.update_folders_tree()
                
                # Очищаем поля
                start_spin.delete(0, tk.END)
                end_spin.delete(0, tk.END)
                text_entry.delete(0, tk.END)
                
            except ValueError:
                messagebox.showerror("Ошибка", "Введите корректные номера фотографий")
        
        ttk.Button(add_frame, text="Добавить правило", command=add_rule).pack(side=tk.LEFT, padx=10)
        
        # Функция удаления правила
        def delete_rule():
            """Удаление выбранного правила"""
            selected = rules_tree.selection()
            if selected:
                index = rules_tree.index(selected[0])
                if 0 <= index < len(folder_data['caption_rules']):
                    folder_data['caption_rules'].pop(index)
                    update_rules_display()
                    self.update_folders_tree()
        
        # Кнопки управления
        btn_frame = ttk.Frame(rules_window)
        btn_frame.pack(fill=tk.X, padx=10, pady=10)
        
        ttk.Button(btn_frame, text="Удалить выбранное правило", command=delete_rule).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Закрыть", command=rules_window.destroy).pack(side=tk.RIGHT, padx=5)
    
    def on_mode_changed(self):
        """Обновляет интерфейс при смене режима"""
        if self.multi_folder_mode.get():
            self.mode_info.config(text="✅ Включен режим 'Несколько папок'. Фотографии будут браться из указанных папок.", 
                                foreground="green")
        else:
            self.mode_info.config(text="⚠️ Включен режим 'Одна папка'. Для использования нескольких папок переключите режим.", 
                                foreground="orange")
    
    def add_multi_folder(self):
        """Добавление новой папки в последовательность"""
        folder = filedialog.askdirectory(title="Выберите папку с фотографиями")
        if folder:
            # Проверяем, есть ли уже такая папка
            for existing_folder in self.folder_sequence:
                if existing_folder['path'] == folder:
                    messagebox.showwarning("Внимание", "Эта папка уже добавлена!")
                    return
            
            # Получаем список изображений в папке
            image_files = self.get_sorted_images_multi_folder(folder)
            
            # Создаем запись о папке
            folder_data = {
                'path': folder,
                'caption_rules': [],  # Правила для этой папки
                'images': image_files  # Список файлов
            }
            
            self.folder_sequence.append(folder_data)
            self.update_folders_tree()
            self.log(f"✓ Добавлена папка: {folder} ({len(image_files)} фото)")
    
    def remove_multi_folder(self):
        """Удаление выбранной папки из последовательности"""
        selected = self.folders_tree.selection()
        if selected:
            index = self.folders_tree.index(selected[0])
            if 0 <= index < len(self.folder_sequence):
                removed_folder = self.folder_sequence.pop(index)
                self.update_folders_tree()
                self.log(f"✓ Удалена папка: {removed_folder['path']}")
    
    def move_folder_up(self):
        """Перемещение папки вверх в последовательности"""
        selected = self.folders_tree.selection()
        if selected:
            index = self.folders_tree.index(selected[0])
            if index > 0:
                # Меняем местами с предыдущей папкой
                self.folder_sequence[index], self.folder_sequence[index-1] = self.folder_sequence[index-1], self.folder_sequence[index]
                self.update_folders_tree()
                # Выделяем перемещенную папку
                self.folders_tree.selection_set(self.folders_tree.get_children()[index-1])
    
    def move_folder_down(self):
        """Перемещение папки вниз в последовательности"""
        selected = self.folders_tree.selection()
        if selected:
            index = self.folders_tree.index(selected[0])
            if index < len(self.folder_sequence) - 1:
                # Меняем местами со следующей папкой
                self.folder_sequence[index], self.folder_sequence[index+1] = self.folder_sequence[index+1], self.folder_sequence[index]
                self.update_folders_tree()
                # Выделяем перемещенную папку
                self.folders_tree.selection_set(self.folders_tree.get_children()[index+1])
    
    def update_folders_tree(self):
        """Обновление дерева папок"""
        self.folders_tree.delete(*self.folders_tree.get_children())
        
        for i, folder_data in enumerate(self.folder_sequence, 1):
            images_count = len(folder_data['images'])
            rules_count = len(folder_data['caption_rules'])
            
            self.folders_tree.insert("", tk.END, values=(
                i, 
                folder_data['path'], 
                f"{images_count} шт.", 
                f"{rules_count} правил"
            ))
    
    def get_sorted_images_multi_folder(self, folder_path):
        """Возвращает отсортированный список изображений для многопапкового режима"""
        if not os.path.exists(folder_path):
            return []
            
        image_files = [f for f in os.listdir(folder_path) if
                      f.lower().endswith(('.png', '.jpg', '.jpeg', '.gif', '.bmp'))]
        
        sort_method = self.multi_folder_sort_method.get()
        
        if sort_method == "name_asc" or sort_method == "По имени (А-Я)":
            image_files.sort(key=self.natural_sort_key)
        elif sort_method == "name_desc" or sort_method == "По имени (Я-А)":
            image_files.sort(key=self.natural_sort_key, reverse=True)
        elif sort_method == "date_asc" or sort_method == "По дате создания (сначала старые)":
            image_files.sort(key=lambda f: os.path.getctime(os.path.join(folder_path, f)))
        elif sort_method == "date_desc" or sort_method == "По дате создания (сначала новые)":
            image_files.sort(key=lambda f: os.path.getctime(os.path.join(folder_path, f)), reverse=True)
        else:
            # По умолчанию - естественная сортировка
            image_files.sort(key=self.natural_sort_key)
            
        return image_files
    
    def get_all_images_multi_folder(self):
        """Получение всех изображений из всех папок в правильном порядке"""
        all_images = []
        current_photo_number = 1
        
        for folder_data in self.folder_sequence:
            folder_path = folder_data['path']
            image_files = folder_data['images']
            
            folder_start_number = current_photo_number
            
            for img_file in image_files:
                full_path = os.path.join(folder_path, img_file)
                all_images.append({
                    'path': full_path,
                    'filename': img_file,
                    'folder_path': folder_path,
                    'global_number': current_photo_number,
                    'folder_start_number': folder_start_number,
                    'folder_rules': folder_data['caption_rules']
                })
                current_photo_number += 1
        
        return all_images
    
    def setup_caption_rules_tab(self, parent):
        """Настраивает вкладку правил подписей"""
        ttk.Label(parent, text="Настройка правил подписей для фотографий", font=("Arial", 12, "bold")).pack(pady=10)
        
        instruction = """Добавьте правила для подписей к фотографиям. Каждое правило применяется к указанному диапазону фотографий."""
        ttk.Label(parent, text=instruction, wraplength=800, justify=tk.LEFT).pack(pady=5, padx=10)
        
        # Фрейм для добавления правил
        add_frame = ttk.LabelFrame(parent, text="Добавить новое правило")
        add_frame.pack(fill=tk.X, padx=10, pady=10)
        
        ttk.Label(add_frame, text="С фото №:").grid(row=0, column=0, padx=5, pady=5)
        self.new_rule_start = ttk.Spinbox(add_frame, from_=1, to=1000, width=8)
        self.new_rule_start.grid(row=0, column=1, padx=5, pady=5)
        
        ttk.Label(add_frame, text="По фото №:").grid(row=0, column=2, padx=5, pady=5)
        self.new_rule_end = ttk.Spinbox(add_frame, from_=1, to=1000, width=8)
        self.new_rule_end.grid(row=0, column=3, padx=5, pady=5)
        
        ttk.Label(add_frame, text="Текст подписи:").grid(row=0, column=4, padx=5, pady=5)
        self.new_rule_text = ttk.Entry(add_frame, width=40)
        self.new_rule_text.grid(row=0, column=5, padx=5, pady=5)
        
        ttk.Button(add_frame, text="Добавить правило", command=self.add_caption_rule).grid(row=0, column=6, padx=5, pady=5)
        
        # Таблица с правилами
        rules_frame = ttk.LabelFrame(parent, text="Текущие правила")
        rules_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Создаем Treeview
        columns = ("start", "end", "text")
        self.rules_tree = ttk.Treeview(rules_frame, columns=columns, show="headings", height=10)
        
        self.rules_tree.heading("start", text="С фото №")
        self.rules_tree.heading("end", text="По фото №")
        self.rules_tree.heading("text", text="Текст подписи")
        
        self.rules_tree.column("start", width=100)
        self.rules_tree.column("end", width=100)
        self.rules_tree.column("text", width=500)
        
        self.rules_tree.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Кнопки управления
        btn_frame = ttk.Frame(rules_frame)
        btn_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Button(btn_frame, text="Удалить выбранное", command=self.delete_caption_rule).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Очистить все", command=self.clear_caption_rules).pack(side=tk.LEFT, padx=5)
        
        # Обновляем отображение правил
        self.update_caption_rules_tree()
    
    def add_caption_rule(self):
        """Добавляет новое правило подписи"""
        try:
            start = int(self.new_rule_start.get())
            end = int(self.new_rule_end.get())
            text = self.new_rule_text.get().strip()
            
            if start > end:
                messagebox.showerror("Ошибка", "Начальный номер не может быть больше конечного")
                return
                
            if not text:
                messagebox.showerror("Ошибка", "Введите текст подписи")
                return
                
            # Добавляем правило
            self.caption_rules.append((start, end, text))
            self.caption_rules.sort(key=lambda x: x[0])  # Сортируем по начальному номеру
            
            # Обновляем дерево
            self.update_caption_rules_tree()
            
            # Очищаем поля
            self.new_rule_start.delete(0, tk.END)
            self.new_rule_end.delete(0, tk.END)
            self.new_rule_text.delete(0, tk.END)
            
            self.log(f"✓ Добавлено правило: фото {start}-{end}")
            
        except ValueError:
            messagebox.showerror("Ошибка", "Введите корректные номера фотографий")
    
    def delete_caption_rule(self):
        """Удаляет выбранное правило подписи"""
        selected = self.rules_tree.selection()
        if selected:
            for item in selected:
                index = self.rules_tree.index(item)
                if 0 <= index < len(self.caption_rules):
                    rule = self.caption_rules.pop(index)
                    self.rules_tree.delete(item)
                    self.log(f"✓ Удалено правило: фото {rule[0]}-{rule[1]}")
    
    def clear_caption_rules(self):
        """Очищает все правила подписей"""
        if messagebox.askyesno("Подтверждение", "Очистить все правила?"):
            self.caption_rules.clear()
            self.update_caption_rules_tree()
            self.log("✓ Все правила очищены")
    
    def update_caption_rules_tree(self):
        """Обновляет отображение правил в дереве"""
        self.rules_tree.delete(*self.rules_tree.get_children())
        for start, end, text in self.caption_rules:
            self.rules_tree.insert("", tk.END, values=(start, end, text))
    
    def get_caption_for_photo(self, photo_number):
        """Возвращает подпись для фото на основе правил"""
        for start, end, text in self.caption_rules:
            if start <= photo_number <= end:
                return f"Фото № {photo_number}. {text}"
        return f"Фото № {photo_number}"
    
    def get_caption_for_photo_multi(self, photo_info):
        """Получение подписи для фото в многопапковом режиме"""
        photo_number = photo_info['global_number']
        folder_rules = photo_info['folder_rules']
        folder_start = photo_info['folder_start_number']
        
        # Вычисляем локальный номер фото в папке
        local_photo_number = photo_number - folder_start + 1
        
        # Сначала проверяем правила конкретной папки (по локальному номеру)
        for start, end, text in folder_rules:
            if start <= local_photo_number <= end:
                return f"Фото № {photo_number}. {text}"
        
        # Затем проверяем общие правила (по глобальному номеру)
        for start, end, text in self.caption_rules:
            if start <= photo_number <= end:
                return f"Фото № {photo_number}. {text}"
        
        # Если правил нет - стандартная подпись
        return f"Фото № {photo_number}"
    
    def natural_sort_key(self, filename):
        """Ключ для естественной сортировки файлов (учитывает числа в названиях)"""
        import re
        return [int(text) if text.isdigit() else text.lower()
                for text in re.split(r'(\d+)', filename)]
    
    def get_all_images_single_folder(self):
        """Получает изображения для одиночного режима"""
        folder = self.screenshots_folder.get()
        if not folder or not os.path.exists(folder):
            return []
        
        image_files = get_image_files(folder)
        
        # Применяем сортировку
        sort_method = self.sort_method.get()
        
        if sort_method == "name_asc":
            image_files.sort(key=natural_sort_key)
        elif sort_method == "name_desc":
            image_files.sort(key=natural_sort_key, reverse=True)
        elif sort_method == "date_asc":
            image_files.sort(key=lambda f: os.path.getctime(os.path.join(folder, f)))
        elif sort_method == "date_desc":
            image_files.sort(key=lambda f: os.path.getctime(os.path.join(folder, f)), reverse=True)
        elif sort_method == "manual" and self.manual_sort_order:
            # Ручная сортировка
            manual_files = [f for f in self.manual_sort_order if f in image_files]
            remaining_files = [f for f in image_files if f not in manual_files]
            image_files = manual_files + remaining_files
        else:
            # По умолчанию - естественная сортировка
            image_files.sort(key=natural_sort_key)
        
        # Преобразуем в нужный формат
        image_data_list = []
        for i, img_file in enumerate(image_files, 1):
            image_data_list.append({
                'path': os.path.join(folder, img_file),
                'filename': img_file,
                'global_number': i,
                'folder_rules': [],
                'folder_start_number': 1
            })
        
        return image_data_list
    
    def apply_sort_to_all_folders(self):
        """Применяет выбранную сортировку ко всем папкам"""
        if not self.folder_sequence:
            messagebox.showinfo("Информация", "Нет добавленных папок")
            return
            
        for folder_data in self.folder_sequence:
            folder_data['images'] = self.get_sorted_images_multi_folder(folder_data['path'])
        
        self.update_folders_tree()
        self.log(f"✓ Сортировка применена ко всем папкам")