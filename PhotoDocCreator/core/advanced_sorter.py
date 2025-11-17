import tkinter as tk
from tkinter import ttk
from PIL import Image, ImageTk, ImageOps
import os
import tempfile

class AdvancedImageSorter:
    def __init__(self, parent, folder_sequence):
        self.parent = parent
        self.folder_sequence = folder_sequence
        self.all_images = []
        self.thumbnails = []
        self.image_rotations = {}  # Храним повороты для каждого изображения
        self.current_order = []
        
    def sort_images(self):
        """Запускает расширенную визуальную сортировку"""
        # Создаем окно сортировки
        self.sort_window = tk.Toplevel(self.parent)
        self.sort_window.title("Расширенная визуальная сортировка")
        self.sort_window.geometry("1200x800")
        self.sort_window.state('zoomed')
        
        # Загружаем все изображения из всех папок
        self.load_all_images()
        
        if not self.thumbnails:
            tk.messagebox.showerror("Ошибка", "Нет изображений для сортировки")
            self.sort_window.destroy()
            return []
        
        # Создаем основной интерфейс
        self.setup_advanced_ui()
        
        # Ждем закрытия окна
        self.parent.wait_window(self.sort_window)
        
        return self.current_order
    
    def load_all_images(self):
        """Загружает все изображения из всех папок"""
        self.all_images = []
        global_counter = 1
        
        for folder_data in self.folder_sequence:
            folder_path = folder_data['path']
            if not os.path.exists(folder_path):
                continue
                
            image_files = [f for f in os.listdir(folder_path) 
                          if f.lower().endswith(('.png', '.jpg', '.jpeg', '.gif', '.bmp'))]
            
            for img_file in image_files:
                img_path = os.path.join(folder_path, img_file)
                try:
                    img = Image.open(img_path)
                    # Создаем миниатюру
                    thumb_size = (120, 90)
                    img.thumbnail(thumb_size, Image.Resampling.LANCZOS)
                    photo = ImageTk.PhotoImage(img)
                    
                    self.thumbnails.append({
                        'global_number': global_counter,
                        'filename': img_file,
                        'path': img_path,
                        'thumbnail': photo,
                        'folder': os.path.basename(folder_path),
                        'original_image': img,
                        'rotation': 0  # начальный угол поворота
                    })
                    
                    global_counter += 1
                    
                except Exception as e:
                    print(f"Ошибка загрузки {img_file}: {e}")
        
        self.current_order = [img['filename'] for img in self.thumbnails]
    
    def setup_advanced_ui(self):
        """Настраивает расширенный интерфейс"""
        main_frame = ttk.Frame(self.sort_window)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Левая панель - миниатюры
        left_frame = ttk.Frame(main_frame)
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        ttk.Label(left_frame, text="Перетащите миниатюры для изменения порядка\nКликните на изображение для управления", 
                 font=("Arial", 10, "bold"), justify=tk.CENTER).pack(pady=5)
        
        # Canvas с прокруткой
        self.canvas = tk.Canvas(left_frame, bg="white")
        scrollbar = ttk.Scrollbar(left_frame, orient=tk.VERTICAL, command=self.canvas.yview)
        self.scrollable_frame = ttk.Frame(self.canvas)
        
        self.scrollable_frame.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=scrollbar.set)
        
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Правая панель - управление
        right_frame = ttk.LabelFrame(main_frame, text="Управление изображением", width=300)
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, padx=(10, 0))
        right_frame.pack_propagate(False)
        
        # Область предпросмотра
        self.preview_label = ttk.Label(right_frame, text="Выберите изображение", justify=tk.CENTER)
        self.preview_label.pack(pady=10)
        
        self.preview_canvas = tk.Canvas(right_frame, bg="lightgray", width=280, height=300)
        self.preview_canvas.pack(pady=10, padx=10)
        
        # Кнопки вращения
        rotation_frame = ttk.LabelFrame(right_frame, text="Вращение изображения")
        rotation_frame.pack(fill=tk.X, padx=10, pady=5)
        
        btn_frame = ttk.Frame(rotation_frame)
        btn_frame.pack(pady=10)
        
        ttk.Button(btn_frame, text="↶ 90°", command=lambda: self.rotate_image(-90)).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="↷ 90°", command=lambda: self.rotate_image(90)).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="180°", command=lambda: self.rotate_image(180)).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Сброс", command=lambda: self.rotate_image(0)).pack(side=tk.LEFT, padx=5)
        
        # Информация об изображении
        info_frame = ttk.LabelFrame(right_frame, text="Информация")
        info_frame.pack(fill=tk.X, padx=10, pady=5)
        
        self.info_label = ttk.Label(info_frame, text="", justify=tk.LEFT)
        self.info_label.pack(padx=10, pady=10)
        
        # Кнопки управления
        control_frame = ttk.Frame(right_frame)
        control_frame.pack(fill=tk.X, padx=10, pady=10)
        
        ttk.Button(control_frame, text="Сохранить порядок", 
                  command=self.save_order).pack(pady=5, fill=tk.X)
        ttk.Button(control_frame, text="Авто-сортировка", 
                  command=self.auto_sort).pack(pady=5, fill=tk.X)
        ttk.Button(control_frame, text="Отмена", 
                  command=self.sort_window.destroy).pack(pady=5, fill=tk.X)
        
        # Отображаем миниатюры
        self.display_thumbnails()
        
        # Обновляем scrollregion
        self.scrollable_frame.update_idletasks()
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))
    
    def display_thumbnails(self):
        """Отображает миниатюры с информацией о папке"""
        for widget in self.scrollable_frame.winfo_children():
            widget.destroy()
        
        self.thumbnail_frames = []
        
        for i, thumb_data in enumerate(self.thumbnails):
            # Фрейм для миниатюры
            thumb_frame = ttk.Frame(self.scrollable_frame, relief="solid", borderwidth=1)
            thumb_frame.grid(row=i//4, column=i%4, padx=5, pady=5, sticky="w")
            thumb_frame.config(width=140, height=140)
            thumb_frame.pack_propagate(False)
            
            # Номер и папка
            info_text = f"{thumb_data['global_number']} - {thumb_data['folder']}"
            number_label = ttk.Label(thumb_frame, text=info_text, font=("Arial", 7, "bold"))
            number_label.pack(anchor="nw", padx=2, pady=2)
            
            # Миниатюра (с учетом поворота)
            rotated_thumb = self.get_rotated_thumbnail(thumb_data)
            img_label = ttk.Label(thumb_frame, image=rotated_thumb)
            img_label.image = rotated_thumb
            img_label.pack(padx=5, pady=2)
            
            # Имя файла
            short_name = thumb_data['filename'][:12] + "..." if len(thumb_data['filename']) > 15 else thumb_data['filename']
            name_label = ttk.Label(thumb_frame, text=short_name, font=("Arial", 7), wraplength=130)
            name_label.pack(padx=2, pady=2)
            
            # Привязываем события
            self.bind_drag_events(thumb_frame, i)
            img_label.bind("<Button-1>", lambda e, idx=i: self.select_image(idx))
            name_label.bind("<Button-1>", lambda e, idx=i: self.select_image(idx))
            
            self.thumbnail_frames.append(thumb_frame)
        
        self.scrollable_frame.update_idletasks()
    
    def get_rotated_thumbnail(self, thumb_data):
        """Возвращает миниатюру с учетом поворота"""
        try:
            if thumb_data['rotation'] == 0:
                return thumb_data['thumbnail']
            
            # Создаем повернутую миниатюру
            img = thumb_data['original_image'].copy()
            rotated_img = img.rotate(thumb_data['rotation'], expand=True)
            rotated_img.thumbnail((120, 90), Image.Resampling.LANCZOS)
            return ImageTk.PhotoImage(rotated_img)
        except Exception as e:
            print(f"Ошибка создания повернутой миниатюры: {e}")
            return thumb_data['thumbnail']
    
    def bind_drag_events(self, widget, index):
        """Привязывает события перетаскивания"""
        widget.bind("<ButtonPress-1>", lambda e: self.on_drag_start(e, index))
        widget.bind("<B1-Motion>", self.on_drag_motion)
        widget.bind("<ButtonRelease-1>", self.on_drag_release)
        
        for child in widget.winfo_children():
            child.bind("<ButtonPress-1>", lambda e: self.on_drag_start(e, index))
            child.bind("<B1-Motion>", self.on_drag_motion)
            child.bind("<ButtonRelease-1>", self.on_drag_release)
    
    def on_drag_start(self, event, index):
        self.drag_start_index = index
        self.dragged_item = self.thumbnail_frames[index]
        self.dragged_item.configure(relief="raised", borderwidth=2)
    
    def on_drag_motion(self, event):
        if self.dragged_item is None:
            return
            
        widget = event.widget.winfo_containing(event.x_root, event.y_root)
        while widget and widget not in self.thumbnail_frames:
            widget = widget.master
        
        if widget and widget in self.thumbnail_frames:
            current_index = self.thumbnail_frames.index(widget)
            if current_index != self.drag_start_index:
                # Перемещаем элемент
                dragged_data = self.thumbnails.pop(self.drag_start_index)
                self.thumbnails.insert(current_index, dragged_data)
                self.display_thumbnails()
                self.drag_start_index = current_index
    
    def on_drag_release(self, event):
        if self.dragged_item:
            self.dragged_item.configure(relief="solid", borderwidth=1)
            self.dragged_item = None
            self.drag_start_index = None
    
    def select_image(self, index):
        """Выбирает изображение для управления"""
        self.selected_index = index
        thumb_data = self.thumbnails[index]
        
        # Показываем превью
        self.show_preview(thumb_data)
        
        # Обновляем информацию
        info_text = f"Файл: {thumb_data['filename']}\n"
        info_text += f"Папка: {thumb_data['folder']}\n"
        info_text += f"Размер: {thumb_data['original_image'].size}\n"
        info_text += f"Поворот: {thumb_data['rotation']}°"
        self.info_label.config(text=info_text)
    
    def show_preview(self, thumb_data):
        """Показывает превью изображения"""
        try:
            img = Image.open(thumb_data['path'])
            if thumb_data['rotation'] != 0:
                img = img.rotate(thumb_data['rotation'], expand=True)
            
            # Масштабируем для превью
            img.thumbnail((260, 280), Image.Resampling.LANCZOS)
            photo = ImageTk.PhotoImage(img)
            
            self.preview_canvas.delete("all")
            self.preview_canvas.create_image(140, 150, image=photo)
            self.preview_canvas.image = photo
            
            self.preview_label.config(text=thumb_data['filename'])
            
        except Exception as e:
            self.preview_label.config(text=f"Ошибка: {str(e)}")
    
    def rotate_image(self, degrees):
        """Поворачивает выбранное изображение"""
        if hasattr(self, 'selected_index'):
            thumb_data = self.thumbnails[self.selected_index]
            
            if degrees == 0:
                # Сброс поворота
                thumb_data['rotation'] = 0
            else:
                # Добавляем поворот
                thumb_data['rotation'] = (thumb_data['rotation'] + degrees) % 360
            
            # Обновляем отображение
            self.display_thumbnails()
            self.select_image(self.selected_index)
    
    def auto_sort(self):
        """Автоматическая сортировка по папкам и имени"""
        self.thumbnails.sort(key=lambda x: (x['folder'], x['filename']))
        self.display_thumbnails()
        self.current_order = [img['filename'] for img in self.thumbnails]
    
    def save_order(self):
        """Сохраняет порядок и повороты"""
        self.current_order = [img['filename'] for img in self.thumbnails]
        
        # Сохраняем информацию о поворотах
        rotation_info = {}
        for thumb_data in self.thumbnails:
            if thumb_data['rotation'] != 0:
                rotation_info[thumb_data['path']] = thumb_data['rotation']
        
        # Можно сохранить rotation_info в конфиг или передать в основной класс
        self.rotation_info = rotation_info
        self.sort_window.destroy()