import tkinter as tk
from tkinter import ttk
from PIL import Image, ImageTk
import os

class VisualImageSorter:
    def __init__(self, parent, image_folder):
        self.parent = parent
        self.image_folder = image_folder
        self.image_files = []
        self.thumbnails = []
        self.current_order = []
        self.dragged_item = None
        self.drag_start_index = None
        
    def sort_images(self):
        """Запускает визуальную сортировку и возвращает новый порядок"""
        # Создаем окно сортировки
        self.sort_window = tk.Toplevel(self.parent)
        self.sort_window.title("Визуальная сортировка фотографий")
        self.sort_window.geometry("1000x700")
        self.sort_window.state('zoomed')  # Разворачиваем на весь экран
        
        # Загружаем изображения
        self.load_images()
        
        if not self.thumbnails:
            tk.messagebox.showerror("Ошибка", "В папке нет изображений или они не загружаются")
            self.sort_window.destroy()
            return []
        
        # Создаем основной фрейм
        main_frame = ttk.Frame(self.sort_window)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Левая панель - миниатюры
        left_frame = ttk.Frame(main_frame)
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        ttk.Label(left_frame, text="Перетащите миниатюры для изменения порядка", 
                 font=("Arial", 10, "bold")).pack(pady=5)
        
        # Canvas с прокруткой для миниатюр
        self.canvas = tk.Canvas(left_frame, bg="white")
        scrollbar = ttk.Scrollbar(left_frame, orient=tk.VERTICAL, command=self.canvas.yview)
        self.scrollable_frame = ttk.Frame(self.canvas)
        
        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )
        
        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=scrollbar.set)
        
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Правая панель - предпросмотр
        right_frame = ttk.LabelFrame(main_frame, text="Предпросмотр", width=400)
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, padx=(10, 0))
        right_frame.pack_propagate(False)
        
        self.preview_label = ttk.Label(right_frame, text="Выберите изображение для предпросмотра")
        self.preview_label.pack(pady=10)
        
        # Область для полноразмерного предпросмотра
        self.preview_canvas = tk.Canvas(right_frame, bg="lightgray", width=380, height=500)
        self.preview_canvas.pack(pady=10, padx=10, fill=tk.BOTH, expand=True)
        
        # Кнопки управления
        button_frame = ttk.Frame(self.sort_window)
        button_frame.pack(fill=tk.X, pady=10)
        
        ttk.Button(button_frame, text="Сохранить порядок", 
                  command=self.save_order).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Отмена", 
                  command=self.sort_window.destroy).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Авто-сортировка по имени", 
                  command=self.auto_sort_by_name).pack(side=tk.RIGHT, padx=5)
        
        # Отображаем миниатюры
        self.display_thumbnails()
        
        # Ждем закрытия окна
        self.parent.wait_window(self.sort_window)
        
        return self.current_order
    
    def load_images(self):
        """Загружает изображения и создает миниатюры"""
        if not os.path.exists(self.image_folder):
            return
            
        # Получаем список изображений
        self.image_files = [f for f in os.listdir(self.image_folder) 
                           if f.lower().endswith(('.png', '.jpg', '.jpeg', '.gif', '.bmp'))]
        
        # Создаем миниатюры
        self.thumbnails = []
        thumb_size = (120, 90)  # Размер миниатюр
        
        for img_file in self.image_files:
            try:
                img_path = os.path.join(self.image_folder, img_file)
                img = Image.open(img_path)
                img.thumbnail(thumb_size, Image.Resampling.LANCZOS)
                photo = ImageTk.PhotoImage(img)
                self.thumbnails.append({
                    'filename': img_file,
                    'thumbnail': photo,
                    'path': img_path,
                    'image_obj': img
                })
            except Exception as e:
                print(f"Ошибка загрузки {img_file}: {e}")
        
        # Начальный порядок
        self.current_order = [img['filename'] for img in self.thumbnails]
    
    def display_thumbnails(self):
        """Отображает миниатюры в scrollable_frame"""
        # Очищаем фрейм
        for widget in self.scrollable_frame.winfo_children():
            widget.destroy()
        
        # Создаем фреймы для миниатюр
        self.thumbnail_frames = []
        
        for i, thumb_data in enumerate(self.thumbnails):
            # Фрейм для одной миниатюры
            thumb_frame = ttk.Frame(self.scrollable_frame, relief="solid", borderwidth=1)
            thumb_frame.grid(row=i//4, column=i%4, padx=5, pady=5, sticky="w")
            thumb_frame.config(width=130, height=130)
            thumb_frame.pack_propagate(False)
            
            # Метка с номером
            number_label = ttk.Label(thumb_frame, text=f"{i+1}", font=("Arial", 8, "bold"))
            number_label.pack(anchor="nw", padx=2, pady=2)
            
            # Изображение
            img_label = ttk.Label(thumb_frame, image=thumb_data['thumbnail'])
            img_label.image = thumb_data['thumbnail']  # Сохраняем ссылку
            img_label.pack(padx=5, pady=2)
            
            # Имя файла (обрезаем если длинное)
            short_name = thumb_data['filename'][:15] + "..." if len(thumb_data['filename']) > 18 else thumb_data['filename']
            name_label = ttk.Label(thumb_frame, text=short_name, font=("Arial", 8), 
                                  wraplength=120, justify="center")
            name_label.pack(padx=2, pady=2)
            
            # Привязываем события для перетаскивания
            self.bind_drag_events(thumb_frame, i)
            
            # Привязываем события для предпросмотра
            img_label.bind("<Button-1>", lambda e, path=thumb_data['path']: self.show_preview(path))
            name_label.bind("<Button-1>", lambda e, path=thumb_data['path']: self.show_preview(path))
            
            self.thumbnail_frames.append(thumb_frame)
        
        # Обновляем scrollregion после загрузки всех изображений
        self.scrollable_frame.update_idletasks()
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))
    
    def bind_drag_events(self, widget, index):
        """Привязывает события перетаскивания к виджету"""
        widget.bind("<ButtonPress-1>", lambda e: self.on_drag_start(e, index))
        widget.bind("<B1-Motion>", self.on_drag_motion)
        widget.bind("<ButtonRelease-1>", self.on_drag_release)
        
        # Привязываем ко всем дочерним элементам
        for child in widget.winfo_children():
            child.bind("<ButtonPress-1>", lambda e: self.on_drag_start(e, index))
            child.bind("<B1-Motion>", self.on_drag_motion)
            child.bind("<ButtonRelease-1>", self.on_drag_release)
    
    def on_drag_start(self, event, index):
        """Начало перетаскивания"""
        self.drag_start_index = index
        self.dragged_item = self.thumbnail_frames[index]
        
        # Визуально выделяем перетаскиваемый элемент
        self.dragged_item.configure(relief="raised", borderwidth=2)
    
    def on_drag_motion(self, event):
        """Движение при перетаскивании"""
        if self.dragged_item is None:
            return
            
        # Находим виджет под курсором
        widget = event.widget.winfo_containing(event.x_root, event.y_root)
        
        # Ищем родительский фрейм миниатюры
        while widget and widget not in self.thumbnail_frames:
            widget = widget.master
        
        if widget and widget in self.thumbnail_frames:
            current_index = self.thumbnail_frames.index(widget)
            if current_index != self.drag_start_index:
                # Перемещаем элемент в списке
                dragged_data = self.thumbnails.pop(self.drag_start_index)
                self.thumbnails.insert(current_index, dragged_data)
                
                # Обновляем отображение
                self.display_thumbnails()
                
                # Обновляем начальный индекс
                self.drag_start_index = current_index
    
    def on_drag_release(self, event):
        """Завершение перетаскивания"""
        if self.dragged_item:
            self.dragged_item.configure(relief="solid", borderwidth=1)
            self.dragged_item = None
            self.drag_start_index = None
    
    def show_preview(self, image_path):
        """Показывает полноразмерный предпросмотр изображения"""
        try:
            img = Image.open(image_path)
            # Масштабируем для предпросмотра
            preview_size = (360, 480)
            img.thumbnail(preview_size, Image.Resampling.LANCZOS)
            photo = ImageTk.PhotoImage(img)
            
            self.preview_canvas.delete("all")
            self.preview_canvas.create_image(190, 250, image=photo)
            self.preview_canvas.image = photo  # Сохраняем ссылку
            
            # Показываем информацию
            file_name = os.path.basename(image_path)
            file_size = os.path.getsize(image_path) // 1024  # KB
            img_info = f"{file_name}\nРазмер: {file_size} KB\n{img.size[0]}×{img.size[1]}px"
            
            self.preview_label.config(text=img_info)
            
        except Exception as e:
            self.preview_label.config(text=f"Ошибка загрузки: {str(e)}")
    
    def auto_sort_by_name(self):
        """Автоматическая сортировка по имени"""
        self.thumbnails.sort(key=lambda x: x['filename'])
        self.display_thumbnails()
        self.current_order = [img['filename'] for img in self.thumbnails]
    
    def save_order(self):
        """Сохраняет текущий порядок и закрывает окно"""
        self.current_order = [img['filename'] for img in self.thumbnails]
        self.sort_window.destroy()