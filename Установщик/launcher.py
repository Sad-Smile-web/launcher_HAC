import os
import zipfile
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import winshell
from win32com.client import Dispatch
import sys
import shutil
import webbrowser
import threading
import traceback
from datetime import datetime
import pythoncom
import time
import tempfile

class ModernGameLauncher:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Sikorsky's Incorporated - HAC Game Manager")
        self.root.geometry("600x550")
        self.root.resizable(False, False)
        
        # Создаем папку для логов
        self.log_dir = os.path.join(os.path.expanduser("~"), "HAC_Launcher_Logs")
        os.makedirs(self.log_dir, exist_ok=True)
        
        # Файл для хранения информации об установке
        self.install_info_file = os.path.join(self.log_dir, "installation_info.txt")
        
        # Временный файл для архива
        self.temp_zip_path = None
        
        # Центрирование окна
        self.center_window()
        
        # Цветовая схема
        self.colors = {
            "primary": "#2C3E50",
            "secondary": "#3498DB",
            "accent": "#E74C3C",
            "success": "#27AE60",
            "warning": "#F39C12",
            "background": "#ECF0F1",
            "text": "#2C3E50"
        }
        
        # Настройка фона
        self.root.configure(bg=self.colors['background'])
        
        # Проверяем, установлена ли игра
        self.installation_path = self.get_installation_info()
        
        # Создание интерфейса
        self.create_widgets()
        
        # Анимация прогресса
        self.progress_animation_id = None
        
    def center_window(self):
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f'{width}x{height}+{x}+{y}')
        
    def get_installation_info(self):
        """Получаем информацию об установленной игре"""
        try:
            if os.path.exists(self.install_info_file):
                with open(self.install_info_file, 'r', encoding='utf-8') as f:
                    return f.read().strip()
        except Exception as e:
            self.log_error("Ошибка чтения информации об установке", e)
        return None
        
    def save_installation_info(self, path):
        """Сохраняем информацию об установленной игре"""
        try:
            with open(self.install_info_file, 'w', encoding='utf-8') as f:
                f.write(path)
            self.installation_path = path
        except Exception as e:
            self.log_error("Ошибка сохранения информации об установке", e)
            
    def clear_installation_info(self):
        """Удаляем информацию об установленной игре"""
        try:
            if os.path.exists(self.install_info_file):
                os.remove(self.install_info_file)
            self.installation_path = None
        except Exception as e:
            self.log_error("Ошибка удаления информации об установке", e)
        
    def log_error(self, error_message, exception=None):
        """Запись ошибок в лог-файл"""
        try:
            log_file = os.path.join(self.log_dir, "installation_errors.txt")
            with open(log_file, "a", encoding="utf-8") as f:
                f.write(f"\n{'='*50}\n")
                f.write(f"Время: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
                f.write(f"Ошибка: {error_message}\n")
                if exception:
                    f.write(f"Тип исключения: {type(exception).__name__}\n")
                    f.write(f"Подробности: {str(exception)}\n")
                    f.write("Трассировка:\n")
                    f.write(traceback.format_exc())
                f.write(f"{'='*50}\n")
            return log_file
        except Exception as e:
            messagebox.showerror("Критическая ошибка", f"Не удалось записать лог: {str(e)}")
            return None
        
    def create_widgets(self):
        # Основной фрейм
        main_frame = tk.Frame(self.root, bg=self.colors['background'], padx=20, pady=20)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Заголовок с логотипом
        header_frame = tk.Frame(main_frame, bg=self.colors['background'])
        header_frame.pack(fill=tk.X, pady=(0, 20))
        
        # Логотип компании
        logo_label = tk.Label(header_frame, 
                              text="Sikorsky's Incorporated",
                              font=('Arial', 18, 'bold'),
                              fg=self.colors['primary'],
                              bg=self.colors['background'])
        logo_label.pack(pady=(0, 5))
        
        subtitle_label = tk.Label(header_frame,
                                  text="HAC Game Manager - Standalone Edition",
                                  font=('Arial', 10),
                                  fg=self.colors['text'],
                                  bg=self.colors['background'])
        subtitle_label.pack()
        
        # Статус установки
        status_frame = tk.Frame(main_frame, bg=self.colors['background'])
        status_frame.pack(fill=tk.X, pady=10)
        
        if self.installation_path and os.path.exists(self.installation_path):
            status_text = f"✓ Игра установлена: {self.installation_path}"
            status_color = self.colors['success']
        else:
            status_text = "○ Игра не установлена"
            status_color = self.colors['text']
            
        self.status_label = tk.Label(status_frame,
                                    text=status_text,
                                    font=('Arial', 10, 'bold'),
                                    fg=status_color,
                                    bg=self.colors['background'])
        self.status_label.pack(anchor=tk.W)
        
        # Информация о версии
        version_frame = tk.Frame(main_frame, bg=self.colors['background'])
        version_frame.pack(fill=tk.X, pady=5)
        
        version_label = tk.Label(version_frame,
                                text="Версия: 1.0 Standalone | Все файлы игры встроены в лаунчер",
                                font=('Arial', 9),
                                fg='green',
                                bg=self.colors['background'])
        version_label.pack(anchor=tk.W)
        
        # Фрейм пути установки
        path_frame = tk.Frame(main_frame, bg=self.colors['background'])
        path_frame.pack(fill=tk.X, pady=10)
        
        path_label = tk.Label(path_frame, 
                             text="Путь установки:", 
                             font=('Arial', 11, 'bold'),
                             fg=self.colors['text'],
                             bg=self.colors['background'])
        path_label.pack(anchor=tk.W)
        
        path_input_frame = tk.Frame(path_frame, bg=self.colors['background'])
        path_input_frame.pack(fill=tk.X, pady=5)
        
        desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
        self.install_path = tk.StringVar(value=os.path.join(desktop_path, "HAC Game"))
        
        self.path_entry = tk.Entry(path_input_frame, 
                                  textvariable=self.install_path, 
                                  font=('Arial', 10),
                                  width=50)
        self.path_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))
        
        self.browse_btn = tk.Button(path_input_frame, 
                                   text="Обзор", 
                                   command=self.select_path,
                                   font=('Arial', 10, 'bold'),
                                   bg=self.colors['primary'],
                                   fg='white',
                                   relief=tk.FLAT,
                                   padx=15)
        self.browse_btn.pack(side=tk.RIGHT)
        
        # Фрейм кнопок
        button_frame = tk.Frame(main_frame, bg=self.colors['background'])
        button_frame.pack(fill=tk.X, pady=20)
        
        # Кнопки в зависимости от статуса установки
        if self.installation_path and os.path.exists(self.installation_path):
            # Если игра установлена - показываем кнопку удаления
            self.install_btn = tk.Button(button_frame, 
                                        text="Удалить игру", 
                                        command=self.start_uninstallation,
                                        font=('Arial', 10, 'bold'),
                                        bg=self.colors['warning'],
                                        fg='white',
                                        relief=tk.FLAT,
                                        padx=20,
                                        pady=10)
        else:
            # Если игра не установлена - показываем кнопку установки
            self.install_btn = tk.Button(button_frame, 
                                        text="Установить игру", 
                                        command=self.start_installation,
                                        font=('Arial', 10, 'bold'),
                                        bg=self.colors['accent'],
                                        fg='white',
                                        relief=tk.FLAT,
                                        padx=20,
                                        pady=10)
        
        self.install_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        self.cancel_btn = tk.Button(button_frame, 
                                   text="Выход", 
                                   command=self.root.quit,
                                   font=('Arial', 10, 'bold'),
                                   bg=self.colors['primary'],
                                   fg='white',
                                   relief=tk.FLAT,
                                   padx=20,
                                   pady=10)
        self.cancel_btn.pack(side=tk.LEFT)
        
        # Прогресс-бар
        self.progress_frame = tk.Frame(main_frame, bg=self.colors['background'])
        self.progress_frame.pack(fill=tk.X, pady=10)
        
        self.progress_label = tk.Label(self.progress_frame, 
                                      text="Готов к работе", 
                                      font=('Arial', 10),
                                      fg=self.colors['text'],
                                      bg=self.colors['background'])
        self.progress_label.pack(anchor=tk.W)
        
        # Создаем кастомный прогресс-бар с использованием Canvas
        self.progress_canvas = tk.Canvas(self.progress_frame, 
                                        height=25, 
                                        bg='white',
                                        highlightthickness=1,
                                        highlightbackground=self.colors['secondary'])
        self.progress_canvas.pack(fill=tk.X, pady=5)
        
        self.progress_bar = self.progress_canvas.create_rectangle(0, 0, 0, 25, 
                                                                fill=self.colors['secondary'],
                                                                outline='')
        
        self.progress_text = self.progress_canvas.create_text(300, 12.5, 
                                                             text="0%",
                                                             font=('Arial', 10, 'bold'),
                                                             fill='white')
        
        # Статус операции
        self.operation_label = tk.Label(self.progress_frame, 
                                       text="", 
                                       font=('Arial', 9),
                                       fg=self.colors['text'],
                                       bg=self.colors['background'])
        self.operation_label.pack(anchor=tk.W)
        
        # Ссылка на поддержку
        support_frame = tk.Frame(main_frame, bg=self.colors['background'])
        support_frame.pack(fill=tk.X, pady=(30, 0))
        
        support_label = tk.Label(support_frame, 
                                text="Техническая поддержка: ", 
                                font=('Arial', 9),
                                fg=self.colors['text'],
                                bg=self.colors['background'])
        support_label.pack(side=tk.LEFT)
        
        # Кликабельная ссылка
        self.support_link = tk.Label(support_frame, 
                                    text="https://sikorsky-support-center.netlify.app/",
                                    font=('Arial', 9, 'underline'), 
                                    fg='blue', 
                                    bg=self.colors['background'],
                                    cursor='hand2')
        self.support_link.pack(side=tk.LEFT)
        self.support_link.bind("<Button-1>", self.open_support_site)
        
        # Информация о компании
        footer_label = tk.Label(main_frame, 
                                text="© 2024-2025 Sikorsky's Incorporated. Все права защищены.",
                                font=('Arial', 8),
                                fg='gray',
                                bg=self.colors['background'])
        footer_label.pack(side=tk.BOTTOM, pady=(20, 0))
        
        # Назначаем обработчики событий для кнопок
        self.setup_button_hover()

    def setup_button_hover(self):
        # Обработчики hover эффектов для кнопок
        def on_enter_install(e):
            if self.installation_path and os.path.exists(self.installation_path):
                self.install_btn.config(bg='#E67E22')
            else:
                self.install_btn.config(bg='#C0392B')
        
        def on_leave_install(e):
            if self.installation_path and os.path.exists(self.installation_path):
                self.install_btn.config(bg=self.colors['warning'])
            else:
                self.install_btn.config(bg=self.colors['accent'])
            
        def on_enter_cancel(e):
            self.cancel_btn.config(bg='#1A252F')
        
        def on_leave_cancel(e):
            self.cancel_btn.config(bg=self.colors['primary'])
            
        def on_enter_browse(e):
            self.browse_btn.config(bg='#1A252F')
        
        def on_leave_browse(e):
            self.browse_btn.config(bg=self.colors['primary'])
            
        self.install_btn.bind("<Enter>", on_enter_install)
        self.install_btn.bind("<Leave>", on_leave_install)
        
        self.cancel_btn.bind("<Enter>", on_enter_cancel)
        self.cancel_btn.bind("<Leave>", on_leave_cancel)
        
        self.browse_btn.bind("<Enter>", on_enter_browse)
        self.browse_btn.bind("<Leave>", on_leave_browse)

    def open_support_site(self, event):
        webbrowser.open_new("https://sikorsky-support-center.netlify.app/")

    def select_path(self):
        path = filedialog.askdirectory(title="Выберите папку для установки")
        if path:
            self.install_path.set(os.path.join(path, "HAC Game"))

    def start_installation(self):
        # Отключаем кнопки
        self.install_btn.config(state='disabled')
        self.cancel_btn.config(state='disabled')
        self.browse_btn.config(state='disabled')
        
        # Запускаем установку в отдельном потоке
        thread = threading.Thread(target=self.install)
        thread.daemon = True
        thread.start()
        
        # Запускаем анимацию прогресса
        self.animate_progress()

    def start_uninstallation(self):
        if messagebox.askyesno("Подтверждение удаления", 
                              "Вы уверены, что хотите удалить игру?\nВсе файлы игры будут удалены."):
            # Отключаем кнопки
            self.install_btn.config(state='disabled')
            self.cancel_btn.config(state='disabled')
            self.browse_btn.config(state='disabled')
            
            # Запускаем удаление в отдельном потоке
            thread = threading.Thread(target=self.uninstall)
            thread.daemon = True
            thread.start()
            
            # Запускаем анимацию прогресса
            self.animate_progress()

    def animate_progress(self):
        if self.progress_animation_id:
            self.root.after_cancel(self.progress_animation_id)
        
        current_width = self.progress_canvas.coords(self.progress_bar)[2]
        max_width = self.progress_canvas.winfo_width()
        
        if current_width < max_width * 0.9:
            new_width = current_width + (max_width * 0.02)
            if new_width > max_width * 0.9:
                new_width = max_width * 0.9
                
            self.progress_canvas.coords(self.progress_bar, 0, 0, new_width, 25)
            percent = int((new_width / max_width) * 100)
            self.progress_canvas.itemconfig(self.progress_text, text=f"{percent}%")
            
            self.progress_animation_id = self.root.after(100, self.animate_progress)
        else:
            self.progress_animation_id = None

    def update_progress(self, value, status="", operation=""):
        max_width = self.progress_canvas.winfo_width()
        new_width = (value / 100) * max_width
        self.progress_canvas.coords(self.progress_bar, 0, 0, new_width, 25)
        self.progress_canvas.itemconfig(self.progress_text, text=f"{int(value)}%")
        
        if status:
            self.progress_label.config(text=status)
        if operation:
            self.operation_label.config(text=operation)
        self.root.update_idletasks()

    def install(self):
        try:
            self.update_progress(5, "Подготовка к установке...", "Начало установки")
            
            install_dir = self.install_path.get()
            
            if not install_dir or install_dir.isspace():
                messagebox.showerror("Ошибка", "Неверный путь установки!")
                self.operation_failed()
                return
                
            self.update_progress(10, "Проверка существующей установки...", "Проверка")
            
            # Создание папки установки
            if os.path.exists(install_dir):
                try:
                    shutil.rmtree(install_dir)
                except PermissionError as e:
                    messagebox.showerror("Ошибка", f"Не удалось удалить существующую папку: {str(e)}\nЗакройте все программы, использующие эту папку.")
                    self.operation_failed()
                    return

            try:
                os.makedirs(install_dir, exist_ok=True)
            except Exception as e:
                messagebox.showerror("Ошибка", f"Не удалось создать папку: {str(e)}")
                self.operation_failed()
                return

            self.update_progress(20, "Извлечение встроенного архива...", "Извлечение архива")
            
            # Извлекаем встроенный архив во временную папку
            try:
                zip_data = self.get_embedded_zip_data()
                if not zip_data:
                    messagebox.showerror("Ошибка", "Не найден встроенный архив с игрой!")
                    self.operation_failed()
                    return
                
                # Сохраняем архив во временный файл
                temp_dir = tempfile.gettempdir()
                self.temp_zip_path = os.path.join(temp_dir, "hac_embedded.zip")
                with open(self.temp_zip_path, 'wb') as f:
                    f.write(zip_data)
                    
            except Exception as e:
                messagebox.showerror("Ошибка", f"Ошибка при извлечении архива: {str(e)}")
                self.operation_failed()
                return

            self.update_progress(30, "Распаковка файлов игры...", "Распаковка")
            
            # Распаковка архива
            try:
                with zipfile.ZipFile(self.temp_zip_path, 'r') as zip_ref:
                    # Получаем список файлов
                    file_list = zip_ref.namelist()
                    total_files = len(file_list)
                    
                    # Распаковываем файлы по одному с обработкой ошибок
                    for i, file in enumerate(file_list):
                        try:
                            # Пропускаем проблемные папки если нужно
                            if "phone" in file.lower() and "button" in file.lower():
                                continue
                            zip_ref.extract(file, install_dir)
                            
                            # Обновляем прогресс
                            progress = 30 + (i / total_files) * 50
                            self.update_progress(progress, f"Распаковка: {i+1}/{total_files} файлов...", "Распаковка")
                            
                        except Exception as e:
                            print(f"Пропущен файл {file}: {str(e)}")
                            continue

            except Exception as e:
                messagebox.showerror("Ошибка", f"Ошибка при распаковке: {str(e)}")
                self.operation_failed()
                return
            finally:
                # Удаляем временный файл архива
                if self.temp_zip_path and os.path.exists(self.temp_zip_path):
                    try:
                        os.remove(self.temp_zip_path)
                    except:
                        pass

            self.update_progress(85, "Поиск исполняемого файла...", "Поиск EXE")
            
            # Поиск exe-файла игры
            game_exe = self.find_game_exe(install_dir)
            if not game_exe:
                messagebox.showerror("Ошибка", "Не найден исполняемый файл игры!")
                self.operation_failed()
                return

            self.update_progress(90, "Создание ярлыка на рабочем столе...", "Создание ярлыка")
            
            # Создание ярлыка
            shortcut_created = self.create_shortcut(install_dir, game_exe)
            if not shortcut_created:
                self.update_progress(95, "Ярлык не создан (см. файл лога)...", "Создание ярлыка")
            else:
                self.update_progress(95, "Ярлык успешно создан!", "Создание ярлыка")

            # Сохраняем информацию об установке
            self.save_installation_info(install_dir)
            
            self.update_progress(100, "Установка завершена успешно!", "Завершено")
            
            # Обновляем статус
            self.status_label.config(text=f"✓ Игра установлена: {install_dir}", fg=self.colors['success'])
            
            # Меняем кнопку на "Удалить игру"
            self.install_btn.config(text="Удалить игру", command=self.start_uninstallation, bg=self.colors['warning'])
            
            if shortcut_created:
                messagebox.showinfo("Успех", "Игра успешно установлена!\nЯрлык создан на рабочем столе.")
            else:
                messagebox.showwarning("Установка завершена", 
                                      "Игра успешно установлена, но не удалось создать ярлык.\n"
                                      f"Подробности в файле: {os.path.join(self.log_dir, 'installation_errors.txt')}")
            
            # Включаем кнопки обратно
            self.operation_complete()
            
        except Exception as e:
            messagebox.showerror("Ошибка", f"Произошла ошибка: {str(e)}")
            self.operation_failed()
            # Удаляем временный файл архива при ошибке
            if self.temp_zip_path and os.path.exists(self.temp_zip_path):
                try:
                    os.remove(self.temp_zip_path)
                except:
                    pass

    def uninstall(self):
        try:
            self.update_progress(5, "Подготовка к удалению...", "Начало удаления")
            
            if not self.installation_path or not os.path.exists(self.installation_path):
                messagebox.showerror("Ошибка", "Не найдена установленная игра!")
                self.operation_failed()
                return

            self.update_progress(20, "Удаление файлов игры...", "Удаление файлов")
            
            # Удаление папки с игрой с повторными попытками
            max_attempts = 3
            for attempt in range(max_attempts):
                try:
                    shutil.rmtree(self.installation_path)
                    break  # Если удаление успешно, выходим из цикла
                except Exception as e:
                    if attempt == max_attempts - 1:  # Последняя попытка
                        messagebox.showerror("Ошибка", f"Не удалось удалить папку с игрой: {str(e)}\nВозможно, файлы используются другим процессом.")
                        self.operation_failed()
                        return
                    else:
                        time.sleep(1)  # Ждем 1 секунду перед повторной попыткой
                        self.update_progress(20, f"Повторная попытка удаления ({attempt + 1}/{max_attempts})...", "Удаление файлов")

            self.update_progress(70, "Удаление ярлыка...", "Удаление ярлыка")
            
            # Удаление ярлыка
            self.delete_shortcut()
            
            self.update_progress(90, "Очистка информации об установке...", "Очистка")
            
            # Удаляем информацию об установке
            self.clear_installation_info()
            
            self.update_progress(100, "Удаление завершено успешно!", "Завершено")
            
            # Обновляем статус
            self.status_label.config(text="○ Игра не установлена", fg=self.colors['text'])
            
            # Меняем кнопку на "Установить игру"
            self.install_btn.config(text="Установить игру", command=self.start_installation, bg=self.colors['accent'])
            
            messagebox.showinfo("Успех", "Игра успешно удалена!")
            
            # Включаем кнопки обратно
            self.operation_complete()
            
        except Exception as e:
            messagebox.showerror("Ошибка", f"Произошла ошибка при удалении: {str(e)}")
            self.operation_failed()

    def delete_shortcut(self):
        """Удаление ярлыка с рабочего стола"""
        try:
            desktop = winshell.desktop()
            shortcut_path = os.path.join(desktop, "HAC Game.lnk")
            
            if os.path.exists(shortcut_path):
                os.remove(shortcut_path)
                return True
        except Exception as e:
            self.log_error("Ошибка при удалении ярлыка", e)
        return False

    def operation_complete(self):
        """Включаем кнопки после завершения операции"""
        self.install_btn.config(state='normal')
        self.cancel_btn.config(state='normal')
        self.browse_btn.config(state='normal')
        
        # Останавливаем анимацию
        if self.progress_animation_id:
            self.root.after_cancel(self.progress_animation_id)
            self.progress_animation_id = None

    def operation_failed(self):
        """Включаем кнопки после неудачной операции"""
        self.install_btn.config(state='normal')
        self.cancel_btn.config(state='normal')
        self.browse_btn.config(state='normal')
        self.update_progress(0, "Операция завершена с ошибкой", "Ошибка")
        if self.progress_animation_id:
            self.root.after_cancel(self.progress_animation_id)
            self.progress_animation_id = None

    def get_embedded_zip_data(self):
        """Получаем встроенный архив из ресурсов EXE"""
        try:
            # Если мы запущены как EXE файл
            if getattr(sys, 'frozen', False):
                # PyInstaller создает временную папку _MEIPASS
                base_path = sys._MEIPASS
                embedded_zip_path = os.path.join(base_path, "hac.zip")
                
                if os.path.exists(embedded_zip_path):
                    with open(embedded_zip_path, 'rb') as f:
                        return f.read()
                else:
                    # Ищем в других возможных местах
                    exe_dir = os.path.dirname(sys.executable)
                    possible_paths = [
                        os.path.join(exe_dir, "hac.zip"),
                        os.path.join(os.getcwd(), "hac.zip")
                    ]
                    
                    for path in possible_paths:
                        if os.path.exists(path):
                            with open(path, 'rb') as f:
                                return f.read()
            else:
                # Для режима разработки - ищем hac.zip рядом со скриптом
                base_dir = os.path.dirname(__file__)
                zip_path = os.path.join(base_dir, "hac.zip")
                
                if os.path.exists(zip_path):
                    with open(zip_path, 'rb') as f:
                        return f.read()
                        
            return None
        except Exception as e:
            self.log_error("Ошибка при получении встроенного архива", e)
            return None

    def find_game_exe(self, install_dir):
        """Поиск исполняемого файла игры"""
        possible_names = [
            "HAC.exe", 
            "HAC.exe", 
            "game.exe", 
            "Game.exe",
            "HAC.exe",
            "hac.exe",
            "Офисный_Хакер.exe"
        ]
        
        # Сначала ищем в корне
        for name in possible_names:
            exe_path = os.path.join(install_dir, name)
            if os.path.exists(exe_path):
                return exe_path
        
        # Ищем в подпапках
        for root, dirs, files in os.walk(install_dir):
            for file in files:
                if file.lower().endswith('.exe'):
                    exe_path = os.path.join(root, file)
                    print(f"Найден EXE: {exe_path}")
                    return exe_path
        
        return None

    def create_shortcut(self, install_dir, exe_path):
        """Создание ярлыка на рабочем столе с обработкой ошибок"""
        try:
            # Инициализация COM для текущего потока
            pythoncom.CoInitialize()
            
            desktop = winshell.desktop()
            shortcut_path = os.path.join(desktop, "HAC Game.lnk")
            
            # Логируем информацию о путях
            debug_info = f"""
            Информация о создании ярлыка:
            - Путь к рабочему столу: {desktop}
            - Путь к ярлыку: {shortcut_path}
            - Путь к исполняемому файлу: {exe_path}
            - Рабочая директория: {install_dir}
            - Файл существует: {os.path.exists(exe_path)}
            - Папка существует: {os.path.exists(install_dir)}
            """
            print(debug_info)
            
            shell = Dispatch('WScript.Shell')
            shortcut = shell.CreateShortCut(shortcut_path)
            shortcut.TargetPath = exe_path
            shortcut.WorkingDirectory = install_dir
            shortcut.IconLocation = exe_path
            shortcut.save()
            
            # Проверяем, что ярлык создан
            if os.path.exists(shortcut_path):
                print(f"Ярлык успешно создан: {shortcut_path}")
                return True
            else:
                error_msg = "Ярлык не был создан, но исключения не было"
                self.log_error(error_msg)
                return False
                
        except Exception as e:
            error_msg = f"Ошибка при создании ярлыка"
            log_file = self.log_error(error_msg, e)
            print(f"Ошибка создания ярлыка. Подробности в файле: {log_file}")
            return False
        finally:
            # Деинициализация COM
            pythoncom.CoUninitialize()

    def run(self):
        self.root.mainloop()

if __name__ == "__main__":
    launcher = ModernGameLauncher()
    launcher.run()