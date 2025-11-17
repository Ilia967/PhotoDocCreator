import tkinter as tk
from tkinter import ttk, scrolledtext
import os

class MainTab(ttk.Frame):
    def __init__(self, parent, app):
        super().__init__(parent)
        self.app = app
        self.setup_ui()
    
    def setup_ui(self):
        # Временно используем упрощенный UI из оригинального кода
        # Пресеты - быстрый доступ
        preset_frame = ttk.Frame(self)
        preset_frame.grid(row=0, column=0, columnspan=3, sticky="we", padx=5, pady=5)
        
        ttk.Label(preset_frame, text="Быстрая загрузка пресета:").pack(side=tk.LEFT, padx=5)
        self.preset_combo = ttk.Combobox(preset_frame, textvariable=self.app.current_preset, width=20, state="readonly")
        self.preset_combo.pack(side=tk.LEFT, padx=5)
        ttk.Button(preset_frame, text="Загрузить", command=self.app.load_selected_preset).pack(side=tk.LEFT, padx=5)
        
        # Заголовок документа
        ttk.Label(self, text="Заголовок документа:", font=("Arial", 11, "bold")).grid(row=1, column=0, sticky="w", padx=5, pady=5)
        department_entry = ttk.Entry(self, textvariable=self.app.department_name, width=80)
        department_entry.grid(row=1, column=1, columnspan=2, sticky="we", padx=5, pady=5)
        
        ttk.Label(self, text="Название фототаблицы:", font=("Arial", 11, "bold")).grid(row=2, column=0, sticky="w", padx=5, pady=5)
        title_entry = ttk.Entry(self, textvariable=self.app.photo_table_title, width=80)
        title_entry.grid(row=2, column=1, columnspan=2, sticky="we", padx=5, pady=5)
        
        # Разделитель
        separator = ttk.Separator(self, orient='horizontal')
        separator.grid(row=3, column=0, columnspan=3, sticky="we", padx=5, pady=10)
        
        # Пути к файлам
        ttk.Label(self, text="Папка с фотографиями:").grid(row=4, column=0, sticky="w", padx=5, pady=5)
        folder_entry = ttk.Entry(self, textvariable=self.app.screenshots_folder, width=60)
        folder_entry.grid(row=4, column=1, padx=5, pady=5)
        ttk.Button(self, text="Обзор", command=self.app.browse_screenshots_folder).grid(row=4, column=2, padx=5, pady=5)
        
        ttk.Label(self, text="Файл для сохранения:").grid(row=5, column=0, sticky="w", padx=5, pady=5)
        file_entry = ttk.Entry(self, textvariable=self.app.word_file, width=60)
        file_entry.grid(row=5, column=1, padx=5, pady=5)
        ttk.Button(self, text="Обзор", command=self.app.browse_word_file).grid(row=5, column=2, padx=5, pady=5)
        
        # Кнопки создания и предпросмотра
        button_frame = ttk.Frame(self)
        button_frame.grid(row=6, column=0, columnspan=3, pady=20)
        
        ttk.Button(button_frame, text="Создать документ", command=self.app.start_creation_process).pack(side=tk.LEFT, padx=10)
        ttk.Button(button_frame, text="Визуальная сортировка", command=self.app.visual_sort_images).pack(side=tk.LEFT, padx=10)
        ttk.Button(button_frame, text="Диагностика системы", command=self.app.show_diagnostics).pack(side=tk.LEFT, padx=10)
        
        # Логи
        self.app.log_text = scrolledtext.ScrolledText(self, height=15, width=90)
        self.app.log_text.grid(row=7, column=0, columnspan=3, padx=5, pady=5, sticky="nsew")
        
        # Настройка растягивания
        self.grid_rowconfigure(7, weight=1)
        self.grid_columnconfigure(1, weight=1)