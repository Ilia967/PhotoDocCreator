import os
import re
from PIL import Image, ImageFile
import logging

# Разрешаем загрузку усеченных изображений
ImageFile.LOAD_TRUNCATED_IMAGES = True

logger = logging.getLogger(__name__)

def natural_sort_key(filename):
    """Ключ для естественной сортировки файлов (учитывает числа в названиях)"""
    return [int(text) if text.isdigit() else text.lower()
            for text in re.split(r'(\d+)', filename)]

def get_image_files(folder_path, extensions=('.png', '.jpg', '.jpeg', '.gif', '.bmp')):
    """Получает список изображений из папки"""
    if not os.path.exists(folder_path):
        logger.warning(f"Папка не существует: {folder_path}")
        return []
    
    try:
        image_files = [f for f in os.listdir(folder_path) 
                      if f.lower().endswith(extensions)]
        return image_files
    except PermissionError:
        logger.error(f"Нет доступа к папке: {folder_path}")
        return []
    except Exception as e:
        logger.error(f"Ошибка чтения папки {folder_path}: {e}")
        return []

def validate_image_file(file_path):
    """Проверяет, что файл является валидным изображением"""
    try:
        with Image.open(file_path) as img:
            img.verify()
        return True
    except Exception as e:
        logger.warning(f"Невалидное изображение {file_path}: {e}")
        return False

def get_image_info(file_path):
    """Возвращает информацию об изображении"""
    try:
        with Image.open(file_path) as img:
            return {
                'size': img.size,
                'format': img.format,
                'mode': img.mode,
                'valid': True
            }
    except Exception as e:
        return {
            'valid': False,
            'error': str(e)
        }

def convert_image_if_needed(file_path, output_path=None):
    """Конвертирует изображение если нужно"""
    try:
        with Image.open(file_path) as img:
            # Если изображение в режиме, который может вызвать проблемы, конвертируем в RGB
            if img.mode in ('RGBA', 'P', 'LA'):
                if output_path is None:
                    output_path = file_path
                
                rgb_img = img.convert('RGB')
                rgb_img.save(output_path, 'JPEG', quality=95)
                return output_path
            
        return file_path
    except Exception as e:
        logger.error(f"Ошибка конвертации {file_path}: {e}")
        return file_path