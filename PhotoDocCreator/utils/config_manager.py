import json
import os
import logging
from tkinter import messagebox

logger = logging.getLogger(__name__)

class ConfigManager:
    def __init__(self, config_file="config.json"):
        self.config_file = config_file
        self.default_config = {
            "screenshots_folder": "",
            "word_file": "",
            "image_width": 6.0,
            "image_height": 9.0,
            "images_per_page": 2,
            "officer_name": "ФИО",
            "officer_rank": "Звание", 
            "officer_position": "Должность",
            "department_name": "Подразделение",
            "photo_table_title": "ФОТОТАБЛИЦА\nк протоколу осмотра предметов от __.__.____",
            "font_family": "Times New Roman",
            "font_size": 12,
            "font_bold": False,
            "sort_method": "name_asc",
            "manual_sort_order": [],
            "caption_rules": [],
            "footer_department": "Подразделение",
            "enable_footer": True,
            "multi_folder_mode": False,
            "multi_folder_sort_method": "name_asc",
            "folder_sequence": []
        }
    
    def load_config(self):
        """Загружает конфигурацию из файла"""
        try:
            if os.path.exists(self.config_file):
                with open(self.config_file, "r", encoding="utf-8") as f:
                    config = json.load(f)
                
                # Обновляем дефолтные значения загруженными
                loaded_config = self.default_config.copy()
                loaded_config.update(config)
                return loaded_config
            else:
                # Создаем файл с настройками по умолчанию
                self.save_config(self.default_config)
                return self.default_config.copy()
                
        except Exception as e:
            logger.error(f"Ошибка загрузки конфигурации: {e}")
            messagebox.showerror("Ошибка", f"Не удалось загрузить настройки: {e}")
            return self.default_config.copy()
    
    def save_config(self, config_data):
        """Сохраняет конфигурацию в файл"""
        try:
            with open(self.config_file, "w", encoding="utf-8") as f:
                json.dump(config_data, f, ensure_ascii=False, indent=2)
            logger.info("Конфигурация сохранена")
            return True
        except Exception as e:
            logger.error(f"Ошибка сохранения конфигурации: {e}")
            messagebox.showerror("Ошибка", f"Не удалось сохранить настройки: {e}")
            return False
    
    def validate_config(self, config_data):
        """Проверяет валидность конфигурации"""
        errors = []
        
        # Проверяем обязательные поля
        required_fields = ['department_name', 'officer_name', 'officer_position']
        for field in required_fields:
            if not config_data.get(field):
                errors.append(f"Обязательное поле '{field}' не заполнено")
        
        # Проверяем числовые поля
        numeric_fields = ['image_width', 'image_height', 'images_per_page', 'font_size']
        for field in numeric_fields:
            value = config_data.get(field, 0)
            if not isinstance(value, (int, float)) or value <= 0:
                errors.append(f"Поле '{field}' должно быть положительным числом")
        
        return errors
    
    def backup_config(self, backup_suffix="_backup"):
        """Создает резервную копию конфигурации"""
        try:
            if os.path.exists(self.config_file):
                import shutil
                backup_file = self.config_file.replace('.json', f'{backup_suffix}.json')
                shutil.copy2(self.config_file, backup_file)
                logger.info(f"Создана резервная копия: {backup_file}")
                return True
        except Exception as e:
            logger.error(f"Ошибка создания резервной копии: {e}")
        return False