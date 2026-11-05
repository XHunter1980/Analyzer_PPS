import os
import pandas as pd
from datetime import datetime, timedelta
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
import random
import chardet
import re
import threading
from queue import Queue
import csv


class PPSEventAnalyzer:

    def __init__(self, root):
        self.root = root
        self.root.title("Анализатор ППС v.4.2")
        self.root.geometry("1000x500")
        self.root.minsize(800, 600)
        # Диапазон адресов ППС
        self.start_ioa = 4000
        self.end_ioa = 4300
        # Погрешность времени события (дельта временного окна поиска каждого события), микросекунды
        self.delta_microseconds = 3000
        # Время последнего события, микросекунды (1мс=1000мкс)
        self.time_final_event = 200000

        # Формат вывода (по умолчанию CSV)
        self.output_format = tk.StringVar(value="csv")
        # Переменная для продвинутой проверки
        self.advanced_check = tk.BooleanVar(value=False)

        # Цветовая схема
        self.bg_color = "#f0f0f0"
        self.primary_color = "#4a6fa5"
        self.secondary_color = "#6c8fc7"
        self.accent_color = "#ff7043"  # Оранжевый цвет
        self.text_color = "#333333"
        self.success_color = "#4caf50"
        self.error_color = "#f44336"
        self.warning_color = "#ff9800"
        self.problem_color = "#800080"  # Лиловый цвет для проблемных ППС

        self.root.configure(bg=self.bg_color)

        # Настройка стилей
        self.style = ttk.Style()
        self.style.theme_use('clam')

        # Конфигурация стилей
        self.style.configure('.', background=self.bg_color, foreground=self.text_color)
        self.style.configure('TFrame', background=self.bg_color)
        self.style.configure('TLabel', background=self.bg_color, foreground=self.text_color, font=('Segoe UI', 9))
        self.style.configure('TButton', font=('Segoe UI', 9), borderwidth=1)
        self.style.configure('TEntry', font=('Segoe UI', 9))
        self.style.map('TButton',
                       foreground=[('active', 'black'), ('!active', "white")],
                       background=[('active', "gold"), ('!active', self.secondary_color)])

        # Настройка стиля прогресс-бара
        self.style.configure("orange.Horizontal.TProgressbar",
                             troughcolor=self.bg_color,
                             background="gold",
                             lightcolor=self.accent_color,
                             darkcolor=self.accent_color)

        self.style.configure('Title.TLabel', font=('Segoe UI', 12, 'bold'), foreground=self.primary_color)
        self.style.configure('Header.TFrame', background=self.primary_color)
        self.style.configure('Header.TLabel', background=self.primary_color, foreground='white',
                             font=('Segoe UI', 10, 'bold'))
        self.style.configure('Accent.TButton', background=self.accent_color, foreground='white')

        self.file_path = None
        self.df = None
        self.results = None
        self.running = False
        self.queue = Queue()
        self.create_widgets()

        self.results_text.tag_configure("problem_pps", foreground=self.problem_color)

    def create_animation_frame(self):
        """Создает фрейм с анимацией между кнопками и прогресс-баром"""
        # Фрейм для анимации
        self.animation_frame = ttk.Frame(self.left_panel)
        self.animation_frame.pack(fill=tk.X, pady=(10, 5))

        # Холст для анимации
        self.animation_canvas = tk.Canvas(
            self.animation_frame,
            bg=self.bg_color,
            height=100,
            highlightthickness=0
        )
        self.animation_canvas.pack(fill=tk.X)

        # Текст "ИСКРА ТЕХНОЛОГИИ"
        self.spark_text = self.animation_canvas.create_text(
            175, 10,
            text="",
            font=('Magneto', 18, 'bold'),
            fill="#2b7bb0"  # цвет анимированного текста
        )

        # Координаты для молнии
        self.lightning_coords = []
        self.lightning_id = None
        self.animation_running = False

    def create_widgets(self):
        # Заголовок
        header_frame = ttk.Frame(self.root, style='Header.TFrame')
        header_frame.pack(fill=tk.X, padx=0, pady=0)

        ttk.Label(header_frame, text="Анализатор событий ППС", style='Header.TLabel').pack(pady=10)

        # Основной контейнер (растягивается на все доступное пространство)
        main_container = ttk.Frame(self.root)
        main_container.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        # Левая панель (управление)
        self.left_panel = ttk.Frame(main_container)
        self.left_panel.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 10))

        # Правая панель (результаты) - растягивается на всю высоту
        right_panel = ttk.Frame(main_container)
        right_panel.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

        # Фрейм выбора файла
        file_frame = ttk.LabelFrame(self.left_panel, text="Выбор файла", padding=(10, 5))
        file_frame.pack(fill=tk.X, padx=5, pady=5)

        file_btn_frame = ttk.Frame(file_frame)
        file_btn_frame.pack(fill=tk.X, pady=2)

        self.file_entry = ttk.Entry(file_btn_frame)
        self.file_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))

        browse_btn = ttk.Button(file_btn_frame, text="Обзор", command=self.browse_file, width=8)
        browse_btn.pack(side=tk.RIGHT)

        # Фрейм параметров времени
        time_frame = ttk.LabelFrame(self.left_panel, text="Параметры времени", padding=(10, 5))
        time_frame.pack(fill=tk.X, padx=5, pady=5)

        # Начальное время
        start_time_frame = ttk.Frame(time_frame)
        start_time_frame.pack(fill=tk.X, pady=2)
        ttk.Label(start_time_frame, text="Начальное время (чч:мм):").pack(side=tk.LEFT)
        self.start_time_entry = ttk.Entry(start_time_frame, width=10)
        self.start_time_entry.pack(side=tk.RIGHT)

        # Количество минут
        minutes_frame = ttk.Frame(time_frame)
        minutes_frame.pack(fill=tk.X, pady=2)
        ttk.Label(minutes_frame, text="Количество минут:").pack(side=tk.LEFT)
        self.minutes_entry = ttk.Entry(minutes_frame, width=10)
        self.minutes_entry.pack(side=tk.RIGHT)

        # Чекбокс для продвинутой проверки
        advanced_frame = ttk.Frame(self.left_panel)
        advanced_frame.pack(fill=tk.X, padx=5, pady=5)
        self.advanced_checkbox = ttk.Checkbutton(
            advanced_frame,
            text="Продвинутая проверка",
            variable=self.advanced_check,
            onvalue=True,
            offvalue=False
        )
        self.advanced_checkbox.pack(side=tk.LEFT)

        # Кнопка анализа
        analyze_btn = ttk.Button(self.left_panel, text="Анализ событий", command=self.analyze, style='Accent.TButton')
        analyze_btn.pack(fill=tk.X, pady=10)

        # Фрейм нижних кнопок
        buttons_frame = ttk.Frame(self.left_panel)
        buttons_frame.pack(fill=tk.X, pady=(5, 0))

        # Кнопка сохранения отчета
        save_btn = ttk.Button(buttons_frame, text="Сохранить отчет", command=self.save_report)
        save_btn.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))

        # Кнопка очистки результатов
        clear_btn = ttk.Button(buttons_frame, text="Очистить", command=self.clear_results)
        clear_btn.pack(side=tk.RIGHT, fill=tk.X, expand=True)

        # Фрейм с анимацией (между кнопками и прогресс-баром)
        self.create_animation_frame()

        # Прогресс-бар с оранжевой заливкой
        self.progress = ttk.Progressbar(
            self.left_panel,
            orient=tk.HORIZONTAL,
            length=200,
            mode='determinate',
            style="orange.Horizontal.TProgressbar"
        )
        self.progress.pack(fill=tk.X, pady=(0, 10))

        # Фрейм результатов (занимает всю правую часть)
        results_frame = ttk.LabelFrame(right_panel, text="Результаты", padding=(10, 5))
        results_frame.pack(fill=tk.BOTH, expand=True)

        # Добавление скроллбара к результатам
        results_container = ttk.Frame(results_frame)
        results_container.pack(fill=tk.BOTH, expand=True)

        scrollbar = ttk.Scrollbar(results_container)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.results_text = tk.Text(
            results_container,
            wrap=tk.WORD,
            yscrollcommand=scrollbar.set,
            font=('Courier New', 10),  # Моноширинный шрифт
            #font=('Segoe UI', 9),
            padx=10,
            pady=10
        )
        self.results_text.pack(fill=tk.BOTH, expand=True)

        scrollbar.config(command=self.results_text.yview)

        # Настройка тегов для цветового выделения
        self.results_text.tag_configure("problem", foreground=self.error_color)
        self.results_text.tag_configure("warning", foreground=self.warning_color)
        self.results_text.tag_configure("normal", foreground=self.text_color)
        self.results_text.tag_configure("success", foreground=self.success_color)

        # Строка состояния
        self.status_bar = ttk.Label(self.root, text="Готов к работе", relief=tk.SUNKEN, anchor=tk.W)
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X)

        # Подпись разработчика
        self.developer_label = ttk.Label(
            self.status_bar,
            text="Екимов Д.А @ 2025 г.",
            anchor=tk.E,
            font=('Segoe UI', 9)
        )
        self.developer_label.pack(side=tk.RIGHT, padx=0)

        # Запускаем анимацию
        self.start_animation()

    def start_animation(self):
        if not hasattr(self, 'animation_canvas'):
            return

        self.animation_running = True
        self.animate_lightning()
        self.animate_text()

    def stop_animation(self):
        self.animation_running = False
        if hasattr(self, 'lightning_id'):
            self.animation_canvas.delete(self.lightning_id)
        if hasattr(self, 'spark_text'):
            self.animation_canvas.itemconfig(self.spark_text, text="")

    def animate_lightning(self):
        if not self.animation_running:
            return

        if hasattr(self, 'lightning_id'):
            self.animation_canvas.delete(self.lightning_id)

        # Генерация случайной молнии
        self.lightning_coords = []
        start_x = random.randint(0, 500)
        start_y = 0

        for i in range(5):
            end_x = start_x + random.randint(-30, 30)
            end_y = start_y + random.randint(15, 35)
            self.lightning_coords.extend([start_x, start_y, end_x, end_y])
            start_x, start_y = end_x, end_y

        self.lightning_id = self.animation_canvas.create_line(
            self.lightning_coords,
            fill="#e2ba22",  # цвет молнии
            width=2,
            capstyle=tk.ROUND
        )

        self.root.after(100, self.animate_lightning)

    def animate_text(self):
        if not self.animation_running:
            return

        text = "ИСКРА ТЕХНОЛОГИИ"
        for i in range(1, len(text) + 1):
            if not self.animation_running:
                break
            self.animation_canvas.itemconfig(
                self.spark_text,
                text=text[:i]
            )
            self.root.update()
            self.root.after(100)

        self.root.after(10000, self.animate_text)

    def clear_results(self):
        self.results_text.delete(1.0, tk.END)
        self.status_bar.config(text="Результаты очищены")

    def browse_file(self):
        # Типы файлов для выбора
        file_types = [
            ("RTP files", "*.rtp"),
            ("CSV files", "*.csv"),
            ("Excel files", "*.xlsx *.xls"),
            ("All files", "*.*")
        ]

        file_path = filedialog.askopenfilename(
            filetypes=file_types,
            title="Выберите файл с данными ППС"
        )

        if file_path:
            self.file_path = file_path
            self.file_entry.delete(0, tk.END)
            self.file_entry.insert(0, file_path)
            self.status_bar.config(text=f"Выбран файл: {os.path.basename(file_path)}")

            # Автоматическая конвертация для файлов лога
            if file_path.lower().endswith(('.rdp', '.rtp', '.txt')):
                self.start_conversion()
            else:
                self.auto_detect_time_range()

    def detect_encoding(self, file_path):
        """Определение кодировки файла с помощью chardet"""
        try:
            with open(file_path, 'rb') as f:
                rawdata = f.read(10000)  # Читаем первые 10KB для определения кодировки
                result = chardet.detect(rawdata)
                return result['encoding'] if result['confidence'] > 0.7 else 'utf-8'
        except:
            return 'utf-8'

    def parse_time(self, time_str):
        """
        Парсинг времени из различных форматов и преобразование в hh:mm:ss.ccc
        Поддерживаемые форматы:
        - hh:mm.sssss -> hh:mm:ss.ccc (например, 11:24.26199 -> 11:24:26.199)
        - hh:mm:ss.sss -> hh:mm:ss.ccc (например, 11:24:26.199 -> 11:24:26.199)
        """
        # Пробуем сначала формат hh:mm.sssss (RDP)
        match = re.search(r'(\d{2}):(\d{2})\.(\d+)', time_str)
        if match:
            hours = match.group(1)
            minutes = match.group(2)
            seconds_part = match.group(3)

            # Разделяем часть с секундами на секунды и миллисекунды
            if len(seconds_part) >= 2:
                seconds = seconds_part[:2].zfill(2)
                milliseconds = seconds_part[2:5].ljust(3, '0')[:3]
            else:
                seconds = '00'
                milliseconds = seconds_part.ljust(3, '0')[:3]

            return f"{hours}:{minutes}:{seconds}.{milliseconds}"

        # Пробуем формат hh:mm:ss.sss (RTP)
        match = re.search(r'(\d{2}):(\d{2}):(\d{2})\.(\d+)', time_str)
        if match:
            hours = match.group(1)
            minutes = match.group(2)
            seconds = match.group(3)
            milliseconds = match.group(4).ljust(3, '0')[:3]
            return f"{hours}:{minutes}:{seconds}.{milliseconds}"

        return ""

    def parse_rdp_block(self, lines, i):
        """Парсинг блока RDP, пропуск пустых строк"""
        # Пропускаем пустые строки
        while i < len(lines) and not lines[i].strip():
            i += 1

        if i + 2 >= len(lines):
            return None, i

        line1 = lines[i].strip()
        match = re.search(r'AdrObj=0x[0-9A-Fa-f]+\((\d+)\)', line1)
        if not match:
            return None, i + 1

        ioa = int(match.group(1))

        line2 = lines[i + 1].strip()
        spi_match = re.search(r'SPI=([01])\(', line2)
        if not spi_match:
            return None, i + 3

        spi = spi_match.group(1)
        event = "Сигнал" if spi == "1" else "Норма"

        line3 = lines[i + 2].strip()
        event_time = self.parse_time(line3)

        return {'ioa': ioa, 'time': event_time, 'event': event}, i + 3

    def save_to_csv(self, data, output_file):
        """Сохранение данных в CSV файл"""
        try:
            with open(output_file, 'w', newline='', encoding='utf-8') as csvfile:
                writer = csv.writer(csvfile, delimiter=',')
                writer.writerow(["T события", "Присоединение", "Наименование", "Событие"])
                for row in data:
                    writer.writerow(row)
            return True
        except Exception as e:
            self.queue.put({'type': 'complete', 'success': False,
                            'message': f"Ошибка сохранения CSV: {str(e)}"})
            return False

    def start_conversion(self):
        """Запуск конвертации в отдельном потоке"""
        if not self.file_path:
            messagebox.showerror("Ошибка", "Пожалуйста, выберите файл для конвертации")
            return

        self.running = True
        self.status_bar.config(text="Конвертация...")
        self.progress['value'] = 0

        # Запускаем конвертацию в отдельном потоке
        threading.Thread(target=self.convert, daemon=True).start()

        # Начинаем проверку очереди
        self.root.after(100, self.check_queue)

    def check_queue(self):
        """Проверка сообщений из потока конвертации"""
        try:
            while not self.queue.empty():
                msg = self.queue.get_nowait()
                if msg['type'] == 'progress':
                    self.progress['value'] = msg['value']
                elif msg['type'] == 'status':
                    self.status_bar.config(text=msg['text'])
                elif msg['type'] == 'complete':
                    self.conversion_complete(msg['success'], msg['message'])

        except Exception as e:
            print(f"Ошибка обработки очереди: {e}")

        if self.running:
            self.root.after(100, self.check_queue)

    def conversion_complete(self, success, message):
        """Завершение конвертации"""
        self.running = False
        self.progress['value'] = 100 if success else 0
        self.status_bar.config(text=message)

        if success:
            # Автоматически выбираем сконвертированный файл для анализа
            self.file_path = message.split(" ")[-1]
            self.file_entry.delete(0, tk.END)
            self.file_entry.insert(0, self.file_path)
            self.auto_detect_time_range()

    def convert(self):
        """Основная функция конвертации, работающая в отдельном потоке"""
        try:
            input_file = self.file_path

            # Определяем кодировку
            encoding = self.detect_encoding(input_file)

            # Читаем файл с правильной кодировкой
            try:
                with open(input_file, 'r', encoding=encoding, errors='replace') as f:
                    lines = [line for line in f.readlines() if line.strip()]  # Удаляем пустые строки
            except UnicodeDecodeError:
                # Резервный вариант с utf-8
                with open(input_file, 'r', encoding='utf-8', errors='replace') as f:
                    lines = [line for line in f.readlines() if line.strip()]  # Удаляем пустые строки

            # Подготавливаем данные
            data = []
            total_lines = len(lines)
            processed_count = 0
            i = 0
            while i < total_lines and self.running:
                # Обновляем прогресс каждые 100 строк или в конце
                if processed_count % 100 == 0 or i >= total_lines - 1:
                    progress = (i / total_lines) * 100
                    self.queue.put({'type': 'progress', 'value': progress})

                block, i = self.parse_rdp_block(lines, i)
                if not block:
                    continue

                if self.start_ioa <= block['ioa'] <= self.end_ioa:
                    pps_num = block['ioa'] - self.start_ioa
                    pps_name = f"ППС_{pps_num}"
                    data.append([block['time'], "ППС", pps_name, block['event']])

                processed_count += 1

            if not self.running:
                self.queue.put({'type': 'complete', 'success': False, 'message': "Конвертация отменена"})
                return

            if not data:
                self.queue.put({'type': 'complete', 'success': False,
                                'message': "Не найдено данных в указанном диапазоне адресов"})
                return

            # Сохраняем файл в директории приложения
            default_name = os.path.splitext(os.path.basename(input_file))[0]
            output_format = self.output_format.get().lower()
            app_dir = os.path.dirname(os.path.abspath(__file__))


            output_file = os.path.join(app_dir, f"{default_name}_converted.csv")
            success = self.save_to_csv(data, output_file)

            if success:
                self.queue.put({'type': 'complete', 'success': True,
                                'message': f"Файл успешно сохранен: {output_file}"})
            else:
                self.queue.put({'type': 'complete', 'success': False,
                                'message': "Ошибка сохранения файла"})

        except Exception as e:
            self.queue.put({'type': 'complete', 'success': False,
                            'message': f"Ошибка при конвертации: {str(e)}"})

    def auto_detect_time_range(self):
        """Автоматическое определение временного диапазона с точным учетом границ"""
        if not self.file_path:
            return

        try:
            # Загружаем данные
            df = self.load_data(self.file_path)

            # Проверяем обязательные столбцы
            if 'T события' not in df.columns:
                return

            # Парсим время событий
            df['parsed_time'] = df['T события'].apply(self.parse_time_for_analysis)
            df = df.dropna(subset=['parsed_time'])

            if df.empty:
                return

            # Находим минимальное и максимальное время
            min_time = df['parsed_time'].min()
            max_time = df['parsed_time'].max()

            # Определяем первую полную минуту (строго с 00 секунды)
            first_full_minute = min_time.replace(second=0, microsecond=0)
            # Если первое событие не в начале минуты, смещаем на следующую минуту
            if min_time.second > 0 or min_time.microsecond > 0:
                first_full_minute += timedelta(minutes=1)

            # Определяем последнюю полную минуту (строго до 59 секунды)
            last_full_minute = max_time.replace(second=0, microsecond=0)
            # Корректируем, если последнее событие не в конце минуты
            if max_time.second < 59 or (max_time.second == 59 and max_time.microsecond < 500000):
                last_full_minute -= timedelta(minutes=1)

            # Гарантируем, что last_full_minute не раньше first_full_minute
            if last_full_minute < first_full_minute:
                last_full_minute = first_full_minute

            # Вычисляем количество полных минут
            delta = last_full_minute - first_full_minute
            total_minutes = max(1, int(delta.total_seconds() // 60) + 1)  # Не менее 1 минуты

            # Устанавливаем значения в поля ввода
            self.start_time_entry.delete(0, tk.END)
            self.start_time_entry.insert(0, first_full_minute.strftime("%H:%M"))

            self.minutes_entry.delete(0, tk.END)
            self.minutes_entry.insert(0, str(total_minutes))

            self.status_bar.config(
                text=f"Определен диапазон: {first_full_minute.strftime('%H:%M')} - "
                     f"{last_full_minute.strftime('%H:%M')} ({total_minutes} полных минут)")

        except Exception as e:
            self.status_bar.config(text=f"Ошибка определения временного диапазона: {str(e)}")

    def load_data(self, file_path):
        """Загрузка данных из файла (Excel или CSV)"""
        try:
            if file_path.endswith(('.xlsx', '.xls')):
                return pd.read_excel(file_path)
            elif file_path.endswith('.csv'):
                # Пробуем определить кодировку автоматически
                with open(file_path, 'rb') as f:
                    result = chardet.detect(f.read(10000))
                    encoding = result['encoding']

                # Пробуем несколько вариантов разделителей
                for sep in [',', ';', '\t']:
                    try:
                        return pd.read_csv(file_path, encoding=encoding, sep=sep)
                    except Exception as e:
                        continue

                # Если не получилось с автоматическим определением
                raise ValueError("Не удалось определить кодировку или разделитель CSV-файла")
            else:
                raise ValueError("Неподдерживаемый формат файла")
        except Exception as e:
            raise ValueError(f"Ошибка загрузки файла: {str(e)}")

    def validate_time_format(self, time_str):
        """Проверка формата времени"""
        try:
            datetime.strptime(time_str, "%H:%M")
            return True
        except ValueError:
            return False

    def parse_time_for_analysis(self, time_str):
        """Парсинг времени для анализа с учетом различных форматов"""
        try:
            # Пробуем разные форматы времени
            if isinstance(time_str, datetime):
                return time_str

            if time_str.count(':') == 3:  # Формат hh:mm:ss:ms
                h, m, s, ms = map(int, time_str.split(':'))
                return datetime(1900, 1, 1, h, m, s, ms * 1000)
            elif time_str.count(':') == 2 and '.' in time_str:  # Формат hh:mm:ss.ms
                h_m_s, ms = time_str.split('.')
                h, m, s = map(int, h_m_s.split(':'))
                return datetime(1900, 1, 1, h, m, s, int(ms.ljust(3, '0')[:3]) * 1000)
            elif time_str.count(':') == 2:  # Формат hh:mm:ss
                return datetime.strptime(time_str, "%H:%M:%S")
            else:  # Формат hh:mm
                return datetime.strptime(time_str, "%H:%M")
        except Exception as e:
            print(f"Ошибка парсинга времени '{time_str}': {e}")
            return None

    def analyze_second(self, second_start, second_end, pps_num, minute_events):
        """Анализ секунды"""
        events = minute_events[
            (minute_events['Номер ППС'] == pps_num) &
            (minute_events['parsed_time'] >= second_start) &
            (minute_events['parsed_time'] <= second_end)
            ]

        if len(events) < 2:
            return {
                'pps_num': pps_num,
                'second': second_start.second,
                'start_time': second_start.time(),
                'end_time': second_end.time(),
                'expected': 2,
                'actual': len(events),
                'missing': 2 - len(events),
                'missing_normal': 0,
                'missing_signal': 0
            }
        return None

    def analyze_minute(self, minute_start, minute_end, pps_list, minute_events):
        """Анализ минуты с учетом продвинутой проверки"""
        missing_seconds = []
        extra_events = []  # Список для хранения информации о лишних событиях

        for pps_num in pps_list:
            for second in range(60):
                second_start = minute_start + timedelta(seconds=second)

                if self.advanced_check.get():
                    # Продвинутая проверка: раздельная проверка событий "Сигнал" и "Норма"

                    # Проверка события "Сигнал" в начале секунды
                    signal_start = second_start - timedelta(microseconds=self.delta_microseconds)
                    signal_end = second_start + timedelta(microseconds=self.delta_microseconds)

                    signal_events = minute_events[
                        (minute_events['Номер ППС'] == pps_num) &
                        (minute_events['Событие'] == 'Сигнал') &
                        (minute_events['parsed_time'] >= signal_start) &
                        (minute_events['parsed_time'] <= signal_end)
                        ]

                    # Проверка события "Норма" в конце секунды
                    normal_start = second_start + timedelta(
                        microseconds=self.time_final_event - self.delta_microseconds)
                    normal_end = second_start + timedelta(microseconds=self.time_final_event + self.delta_microseconds)

                    normal_events = minute_events[
                        (minute_events['Номер ППС'] == pps_num) &
                        (minute_events['Событие'] == 'Норма') &
                        (minute_events['parsed_time'] >= normal_start) &
                        (minute_events['parsed_time'] <= normal_end)
                        ]

                    # Определяем пропущенные события
                    missing_signal = 1 - len(signal_events)
                    missing_normal = 1 - len(normal_events)
                    total_missing = missing_signal + missing_normal
                    total_actual = len(signal_events) + len(normal_events)

                    if total_missing > 0:
                        missing_seconds.append({
                            'pps_num': pps_num,
                            'second': second,
                            'start_time': second_start.time(),
                            'end_time': (second_start + timedelta(microseconds=self.time_final_event)).time(),
                            'expected': 2,
                            'actual': total_actual,
                            'missing': total_missing,
                            'missing_normal': missing_normal,
                            'missing_signal': missing_signal,
                            'advanced_check': True,
                            'signal_found': len(signal_events) > 0,
                            'normal_found': len(normal_events) > 0
                        })

                    # Проверка на лишние события (только для продвинутой проверки)
                    all_events_in_window = minute_events[
                        (minute_events['Номер ППС'] == pps_num) &
                        (minute_events['parsed_time'] >= signal_start) &
                        (minute_events['parsed_time'] <= normal_end)
                        ]

                    if len(all_events_in_window) > 2:
                        for _, event in all_events_in_window.iterrows():
                            # Проверяем, не является ли событие одним из ожидаемых
                            is_expected_signal = signal_start <= event['parsed_time'] <= signal_end and event[
                                'Событие'] == 'Сигнал'
                            is_expected_normal = normal_start <= event['parsed_time'] <= normal_end and event[
                                'Событие'] == 'Норма'

                            if not is_expected_signal and not is_expected_normal:
                                extra_events.append({
                                    'pps_num': pps_num,
                                    'second': second,
                                    'time': event['parsed_time'].time(),
                                    'event_type': event['Событие'],
                                    'expected': 2,
                                    'actual': len(all_events_in_window),
                                    'window_start': signal_start.time(),
                                    'window_end': normal_end.time(),
                                    'is_advanced': True
                                })
                else:
                    # Обычная проверка: просто проверяем 2 события в секунду
                    second_end = second_start + timedelta(microseconds=self.time_final_event)

                    search_start = second_start - timedelta(microseconds=self.delta_microseconds)
                    search_end = second_end + timedelta(microseconds=self.delta_microseconds)

                    events = minute_events[
                        (minute_events['Номер ППС'] == pps_num) &
                        (minute_events['parsed_time'] >= search_start) &
                        (minute_events['parsed_time'] <= search_end)
                        ]

                    event_types = events['Событие'].value_counts()
                    missing_normal = 1 - event_types.get('Норма', 0)
                    missing_signal = 1 - event_types.get('Сигнал', 0)
                    total_missing = missing_signal + missing_normal
                    total_actual = len(events)

                    if total_missing > 0:
                        missing_seconds.append({
                            'pps_num': pps_num,
                            'second': second,
                            'start_time': second_start.time(),
                            'end_time': second_end.time(),
                            'expected': 2,
                            'actual': total_actual,
                            'missing': total_missing,
                            'missing_normal': missing_normal,
                            'missing_signal': missing_signal,
                            'advanced_check': False,
                            'signal_found': event_types.get('Сигнал', 0) > 0,
                            'normal_found': event_types.get('Норма', 0) > 0
                        })

                    # Проверка на лишние события для обычной проверки
                    if len(events) > 2:
                        for _, event in events.iterrows():
                            extra_events.append({
                                'pps_num': pps_num,
                                'second': second,
                                'time': event['parsed_time'].time(),
                                'event_type': event['Событие'],
                                'expected': 2,
                                'actual': len(events),
                                'window_start': search_start.time(),
                                'window_end': search_end.time(),
                                'is_advanced': False
                            })

        return missing_seconds, extra_events

    def analyze(self):
        """Основная функция анализа с обработкой пропущенных и лишних событий"""
        if not self.file_path:
            messagebox.showerror("Ошибка", "Пожалуйста, выберите файл для анализа")
            return

        # Настройка интерфейса во время анализа
        self.root.config(cursor="watch")
        self.root.update()
        self.results_text.config(bg='#ffffcc')
        self.root.update()
        self.root.after(100, lambda: self.results_text.config(bg='white'))
        self.root.after(150, self.clear_results)

        self.results_text.delete(1.0, tk.END)
        self.status_bar.config(text="Анализ данных...")
        self.root.update_idletasks()

        # Проверка формата времени
        start_time_str = self.start_time_entry.get()
        if not self.validate_time_format(start_time_str):
            messagebox.showerror("Ошибка", "Неверный формат времени. Используйте чч:мм")
            return

        # Проверка количества минут
        try:
            minutes = int(self.minutes_entry.get())
            if minutes <= 0 or minutes > 99:
                raise ValueError
        except ValueError:
            messagebox.showerror("Ошибка", "Количество минут должно быть 1-99")
            return

        self.start_animation()

        try:
            # Загрузка данных
            self.df = self.load_data(self.file_path)

            # Проверка обязательных столбцов
            required_columns = ['T события', 'Наименование', 'Присоединение', 'Событие']
            for col in required_columns:
                if col not in self.df.columns:
                    messagebox.showerror("Ошибка", f"В файле отсутствует обязательный столбец: {col}")
                    return

            # Фильтрация данных ППС
            pps_df = self.df[self.df['Присоединение'] == 'ППС'].copy()
            if pps_df.empty:
                messagebox.showerror("Ошибка", "В файле нет данных с Присоединение = ППС")
                return

            # Парсинг времени и номеров ППС
            pps_df['parsed_time'] = pps_df['T события'].apply(self.parse_time_for_analysis)
            pps_df['Номер ППС'] = pps_df['Наименование'].str.extract(r'ППС_(\d+)').astype(int)

            # Определение диапазона ППС
            min_pps = pps_df['Номер ППС'].min()
            max_pps = pps_df['Номер ППС'].max()
            pps_count = pps_df['Номер ППС'].nunique()

            # Подготовка временных диапазонов
            base_time = datetime.strptime(start_time_str, "%H:%M")
            minute_ranges = []

            for minute in range(minutes):
                minute_start = base_time + timedelta(minutes=minute)
                minute_end = minute_start + timedelta(seconds=59,
                                                      microseconds=(self.time_final_event + self.delta_microseconds))

                # Определение фактических границ минуты
                search_start = minute_start - timedelta(microseconds=self.delta_microseconds)
                search_end = minute_end.replace(second=59,
                                                microsecond=(self.time_final_event + self.delta_microseconds))

                first_events = pps_df[
                    (pps_df['parsed_time'] >= search_start) &
                    (pps_df['parsed_time'] <= search_start + timedelta(microseconds=self.delta_microseconds * 2))
                    ].sort_values('parsed_time')

                last_events = pps_df[
                    (pps_df['parsed_time'] >= search_end - timedelta(microseconds=self.delta_microseconds * 2)) &
                    (pps_df['parsed_time'] <= search_end)
                    ].sort_values('parsed_time')

                actual_start = first_events.iloc[0]['parsed_time'] if not first_events.empty else minute_start
                actual_end = last_events.iloc[-1]['parsed_time'] if not last_events.empty else minute_end

                minute_ranges.append((actual_start, actual_end))

            # Анализ данных
            analysis_results = []
            problem_minutes = []
            total_expected_events = 0
            total_actual_events = 0
            total_missing_events = 0
            total_missing_normal = 0
            total_missing_signal = 0
            total_extra_events = 0

            for i, (start, end) in enumerate(minute_ranges):
                minute_events = pps_df[
                    (pps_df['parsed_time'] >= start) &
                    (pps_df['parsed_time'] <= end)
                    ].copy().sort_values('parsed_time')

                # Статистика по минуте
                event_counts = minute_events['Событие'].value_counts()
                normal_count = event_counts.get('Норма', 0)
                signal_count = event_counts.get('Сигнал', 0)
                total_count = len(minute_events)
                expected_per_minute = pps_count * 2 * 60
                is_problem_minute = total_count != expected_per_minute

                if is_problem_minute:
                    problem_minutes.append(i + 1)

                # Анализ проблемных ППС
                missing_events_info = []
                missing_seconds_info = []
                extra_events_info = []

                if is_problem_minute:
                    pps_event_counts = minute_events.groupby('Номер ППС').size()
                    problem_pps = pps_event_counts[pps_event_counts != 120].index.tolist()

                    for pps_num in problem_pps:
                        actual = pps_event_counts[pps_num]
                        missing_events_info.append({
                            'pps_num': pps_num,
                            'expected': 120,
                            'actual': actual
                        })

                        # Анализ секунд
                        missing_seconds, extra_events = self.analyze_minute(start, end, [pps_num], minute_events)
                        missing_seconds_info.extend(missing_seconds)
                        extra_events_info.extend(extra_events)

                        # Подсчет статистики
                        for sec in missing_seconds:
                            total_missing_events += sec['missing']
                            total_missing_normal += sec['missing_normal']
                            total_missing_signal += sec['missing_signal']

                        total_extra_events += len(extra_events)

                # Сохранение результатов минуты
                analysis_results.append({
                    'minute': i + 1,
                    'start_time': start.time(),
                    'end_time': end.time(),
                    'normal_count': normal_count,
                    'signal_count': signal_count,
                    'total_count': total_count,
                    'expected_count': expected_per_minute,
                    'is_problem': is_problem_minute,
                    'missing_events': missing_events_info,
                    'missing_seconds': missing_seconds_info,
                    'extra_events': extra_events_info
                })

                total_expected_events += expected_per_minute
                total_actual_events += total_count

            # Формирование отчета с улучшенным форматированием
            report_lines = []
            report_lines.append("=== ОБЩИЙ ОТЧЕТ ===")
            report_lines.append(f"Минимальный номер ППС: {min_pps}")
            report_lines.append(f"Максимальный номер ППС: {max_pps}")
            report_lines.append(f"Количество ППС: {pps_count}")
            report_lines.append(f"Ожидаемое количество событий в минуту для каждого ППС: 120")
            report_lines.append(f"Тип проверки: {'Продвинутая' if self.advanced_check.get() else 'Обычная'}")
            report_lines.append("")

            # Итоги
            report_lines.append("=== ИТОГИ ===")
            report_lines.append(f"Общее ожидаемое количество событий: {total_expected_events}")
            report_lines.append(f"Общее реальное количество событий: {total_actual_events}")

            if total_actual_events > total_expected_events:
                report_lines.append("\n=== ВНИМАНИЕ ===")
                report_lines.append(
                    f"Событий больше, чем ожидалось! Превышение: {total_actual_events - total_expected_events}")
                report_lines.append(f"Обнаружено отклонений: {total_extra_events}")
            elif total_missing_events > 0:
                report_lines.append("\n=== ПОТЕРИ СОБЫТИЙ ===")
                report_lines.append(f"Всего потеряно событий: {total_missing_events}")
                report_lines.append(f"Из них 'Норма': {total_missing_normal}")
                report_lines.append(f"Из них 'Сигнал': {total_missing_signal}")
            else:
                report_lines.append("\nКОЛИЧЕСТВО СОБЫТИЙ СООТВЕТСТВУЕТ ОЖИДАЕМОМУ!")

            if problem_minutes:
                report_lines.append(f"\nПРОБЛЕМНЫЕ МИНУТЫ: {', '.join([f'№{m}' for m in problem_minutes])}")

            # Поминутный анализ с улучшенным форматированием
            report_lines.append("\n=== ПОМИНУТНЫЙ АНАЛИЗ ===")
            for result in analysis_results:
                if result['is_problem']:
                    report_lines.append(
                        f"\nМИНУТА № {result['minute']} (ПРОБЛЕМНАЯ): {result['start_time'].strftime('%H:%M:%S')} - {result['end_time'].strftime('%H:%M:%S')}")
                else:
                    report_lines.append(
                        f"\nМинута № {result['minute']}: {result['start_time'].strftime('%H:%M:%S')} - {result['end_time'].strftime('%H:%M:%S')}")

                report_lines.append(f"  Ожидаемое количество событий: {result['expected_count']}")
                status_line = f"  РЕАЛЬНОЕ количество событий: {result['total_count']} (ОТКЛОНЕНИЕ!)" if result[
                    'is_problem'] else f"  Реальное количество событий: {result['total_count']}"
                report_lines.append(status_line)

                report_lines.append(
                    f"  Событий 'Норма': {result['normal_count']:>4} из {int(result['expected_count'] / 2):>4}, отклонение: {int(result['expected_count'] / 2) - result['normal_count']:>3}")
                report_lines.append(
                    f"  Событий 'Сигнал': {result['signal_count']:>4} из {int(result['expected_count'] / 2):>4}, отклонение: {int(result['expected_count'] / 2) - result['signal_count']:>3}")

                # Пропущенные события с улучшенным форматированием
                if result['missing_events']:
                    report_lines.append(f"  Проблемные ППС (всего {len(result['missing_events'])}):")
                    for missing in result['missing_events'][:10]:
                        report_lines.append(
                            f"    ППС {missing['pps_num']:>3}: ожидалось {missing['expected']:>3}, получено {missing['actual']:>3}")
                    if len(result['missing_events']) > 10:
                        report_lines.append(
                            f"    ... и ещё {len(result['missing_events']) - 10} ППС (см. полный отчёт)")

                if result['missing_seconds']:
                    report_lines.append(f"  Пропущенные события по секундам (всего {len(result['missing_seconds'])}):")
                    for sec in result['missing_seconds'][:10]:
                        start_time_str = sec['start_time'].strftime("%H:%M:%S.%f")[:-3]
                        end_time_str = sec['end_time'].strftime("%H:%M:%S.%f")[:-3]
                        if self.advanced_check.get():
                            report_lines.append(
                                f"    ППС {sec['pps_num']:>3}, секунда {sec['second']:>2}: "
                                f"ожидалось {sec['expected']:>2}, получено {sec['actual']:>2} "
                                f"(время: {start_time_str} - {end_time_str}) "
                                f"[Сигнал: {'найден' if sec['signal_found'] else 'пропущен':>8}, "
                                f"Норма: {'найдена' if sec['normal_found'] else 'пропущена':>8}]")
                        else:
                            report_lines.append(
                                f"    ППС {sec['pps_num']:>3}, секунда {sec['second']:>2}: "
                                f"ожидалось {sec['expected']:>2}, получено {sec['actual']:>2} "
                                f"(время: {start_time_str} - {end_time_str})")

                    if len(result['missing_seconds']) > 10:
                        report_lines.append(
                            f"    ... и ещё {len(result['missing_seconds']) - 10} секунд (см. полный отчёт)")

                # Лишние события с улучшенным форматированием
                if result['extra_events']:
                    report_lines.append(f"    Лишние события (всего {len(result['extra_events'])}):")
                    for event in result['extra_events'][:10]:
                        event_time_str = event['time'].strftime("%H:%M:%S.%f")[:-3]
                        window_start_str = event['window_start'].strftime("%H:%M:%S.%f")[:-3]
                        window_end_str = event['window_end'].strftime("%H:%M:%S.%f")[:-3]
                        report_lines.append(
                            f"    ППС {event['pps_num']:>3}, секунда {event['second']:>2}: "
                            f"{event_time_str} ({event['event_type']:>6}) - "
                            f"ожидалось {event['expected']:>2}, получено {event['actual']:>2} "
                            f"(окно: {window_start_str} - {window_end_str})")
                    if len(result['extra_events']) > 10:
                        report_lines.append(
                            f"    ... и ещё {len(result['extra_events']) - 10} событий (см. полный отчёт)")

            # Вывод отчета с цветовым выделением и моноширинным шрифтом
            self.results_text.delete(1.0, tk.END)
            for line in report_lines:
                if "ППС " in line and ("ожидалось" in line or "секунда" in line):
                    self.results_text.insert(tk.END, line + "\n", "problem_pps")
                elif "(ПРОБЛЕМНАЯ)" in line or "ПРОБЛЕМНЫЕ МИНУТЫ:" in line:
                    self.results_text.insert(tk.END, line + "\n", "problem")
                elif "ОТКЛОНЕНИЕ" in line or "ПОТЕРИ" in line or "Событий больше, чем ожидалось!" in line or "=== ВНИМАНИЕ ===" in line:
                    self.results_text.insert(tk.END, line + "\n", "warning")
                elif "КОЛИЧЕСТВО СОБЫТИЙ СООТВЕТСТВУЕТ ОЖИДАЕМОМУ!" in line:
                    self.results_text.insert(tk.END, line + "\n", "success")
                else:
                    self.results_text.insert(tk.END, line + "\n", "normal")

            # Сохранение результатов
            self.results = {
                'summary': {
                    'min_pps': min_pps,
                    'max_pps': max_pps,
                    'pps_count': pps_count,
                    'total_events': total_actual_events,
                    'normal_events': normal_count,
                    'signal_events': signal_count,
                    'expected_per_minute': 120,
                    'problem_minutes': problem_minutes,
                    'total_expected': total_expected_events,
                    'total_actual': total_actual_events,
                    'total_missing': total_missing_events,
                    'total_missing_normal': total_missing_normal,
                    'total_missing_signal': total_missing_signal,
                    'total_extra_events': total_extra_events,
                    'advanced_check': self.advanced_check.get()
                },
                'minute_details': analysis_results
            }

            self.status_bar.config(text="Анализ завершен успешно")

        except Exception as e:
            messagebox.showerror("Ошибка", f"Произошла ошибка при анализе: {str(e)}")
            self.status_bar.config(text=f"Ошибка: {str(e)}")
        finally:
            self.stop_animation()
            self.root.config(cursor="")

    def save_report(self):
        """Сохранение отчета в файл с учетом продвинутой проверки и лишних событий"""
        if not self.results:
            messagebox.showerror("Ошибка", "Нет данных для сохранения. Сначала выполните анализ.")
            return

        try:
            save_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx")]
            )

            if not save_path:
                return

            wb = Workbook()
            ws = wb.active
            ws.title = "Отчет"

            # Настройки стилей
            header_font = Font(bold=True)
            problem_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
            normal_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
            warning_fill = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
            header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
            white_font = Font(color="FFFFFF", bold=True)

            # Заголовок отчета с указанием типа проверки
            ws.append(["ОТЧЕТ ПО АНАЛИЗУ СОБЫТИЙ ППС"])
            ws.append(["Тип проверки:", "Продвинутая" if self.results['summary']['advanced_check'] else "Обычная"])
            ws.append([])

            # Основная информация
            ws.append(["Минимальный номер ППС:", self.results['summary']['min_pps']])
            ws.append(["Максимальный номер ППС:", self.results['summary']['max_pps']])
            ws.append(["Количество ППС:", self.results['summary']['pps_count']])
            ws.append(["Событий 'Норма':", self.results['summary']['normal_events']])
            ws.append(["Событий 'Сигнал':", self.results['summary']['signal_events']])
            ws.append(["Ожидаемое количество событий в минуту для каждого ППС:", 120])
            ws.append([])

            # Итоги
            ws.append(["ИТОГИ"])
            ws.append(["Общее ожидаемое количество событий:", self.results['summary']['total_expected']])
            ws.append(["Общее реальное количество событий:", self.results['summary']['total_actual']])

            if self.results['summary']['total_actual'] > self.results['summary']['total_expected']:
                ws.append(["ВНИМАНИЕ"])
                ws.append(["Событий больше, чем ожидалось!",
                           self.results['summary']['total_actual'] - self.results['summary']['total_expected']])
                ws.append(["Это может указывать на проблемы в системе."])
            elif self.results['summary']['total_missing'] > 0:
                ws.append(["ПОТЕРИ СОБЫТИЙ"])
                ws.append(["Всего потеряно событий:", self.results['summary']['total_missing']])
                ws.append(["Из них 'Норма':", self.results['summary']['total_missing_normal']])
                ws.append(["Из них 'Сигнал':", self.results['summary']['total_missing_signal']])
            else:
                ws.append([])
                ws.append(["ПОЗДРАВЛЯЕМ - КОЛИЧЕСТВО СОБЫТИЙ СООТВЕТСТВУЕТ ОЖИДАЕМОМУ!"])

            if self.results['summary']['problem_minutes']:
                ws.append([])
                ws.append(["ПРОБЛЕМНЫЕ МИНУТЫ:", ", ".join(map(str, self.results['summary']['problem_minutes']))])
            else:
                ws.append([])
                ws.append(["Все данные соответствуют ожиданиям!"])

            ws.append([])

            # Поминтуный анализ
            ws.append(["ПОМИНУТНЫЙ АНАЛИЗ"])
            ws.append([
                "Минута", "Начало", "Конец",
                "Ожидаемое", "Реальное",
                "Норма", "Сигнал", "Статус", "Проблемные ППС"
            ])

            # Форматирование заголовков
            for row in ws.iter_rows(min_row=1, max_row=1, max_col=1):
                for cell in row:
                    cell.font = Font(bold=True, size=14)

            for row in ws.iter_rows(min_row=ws.max_row, max_row=ws.max_row, max_col=9):
                for cell in row:
                    cell.font = header_font
                    cell.fill = header_fill
                    cell.font = white_font

            # Добавление данных по минутам
            for detail in self.results['minute_details']:
                problem_pps = ", ".join([str(m['pps_num']) for m in detail['missing_events']]) if detail[
                    'missing_events'] else ""

                row = [
                    detail['minute'],
                    detail['start_time'].strftime("%H:%M:%S.%f")[:-3],
                    detail['end_time'].strftime("%H:%M:%S.%f")[:-3],
                    detail['expected_count'],
                    detail['total_count'],
                    detail['normal_count'],
                    detail['signal_count'],
                    "ПРОБЛЕМА" if detail['is_problem'] else "OK",
                    problem_pps
                ]

                ws.append(row)

                # Цветовое выделение строк
                if detail['is_problem']:
                    for cell in ws[ws.max_row]:
                        cell.fill = problem_fill
                else:
                    for cell in ws[ws.max_row]:
                        cell.fill = normal_fill

            # Лист с пропущенными секундами (если есть)
            if any(detail['missing_seconds'] for detail in self.results['minute_details']):
                ws_seconds = wb.create_sheet("Пропущенные секунды")

                if self.results['summary']['advanced_check']:
                    ws_seconds.append([
                        "Минута", "ППС", "Секунда",
                        "Начало интервала", "Конец интервала",
                        "Ожидалось", "Получено", "Пропущено",
                        "Пропущено 'Норма'", "Пропущено 'Сигнал'",
                        "Найдено 'Сигнал'", "Найдено 'Норма'"
                    ])
                else:
                    ws_seconds.append([
                        "Минута", "ППС", "Секунда",
                        "Начало интервала", "Конец интервала",
                        "Ожидалось", "Получено", "Пропущено",
                        "Пропущено 'Норма'", "Пропущено 'Сигнал'"
                    ])

                # Форматирование заголовков
                for row in ws_seconds.iter_rows(min_row=1, max_row=1, max_col=ws_seconds.max_column):
                    for cell in row:
                        cell.font = header_font
                        cell.fill = header_fill
                        cell.font = white_font

                # Добавление данных о пропущенных секундах
                for detail in self.results['minute_details']:
                    if detail['missing_seconds']:
                        for sec in detail['missing_seconds']:
                            if self.results['summary']['advanced_check']:
                                ws_seconds.append([
                                    detail['minute'],
                                    sec['pps_num'],
                                    sec['second'],
                                    sec['start_time'].strftime("%H:%M:%S.%f")[:-3],
                                    sec['end_time'].strftime("%H:%M:%S.%f")[:-3],
                                    sec['expected'],
                                    sec['actual'],
                                    sec['missing'],
                                    sec['missing_normal'],
                                    sec['missing_signal'],
                                    "Да" if sec['signal_found'] else "Нет",
                                    "Да" if sec['normal_found'] else "Нет"
                                ])
                            else:
                                ws_seconds.append([
                                    detail['minute'],
                                    sec['pps_num'],
                                    sec['second'],
                                    sec['start_time'].strftime("%H:%M:%S.%f")[:-3],
                                    sec['end_time'].strftime("%H:%M:%S.%f")[:-3],
                                    sec['expected'],
                                    sec['actual'],
                                    sec['missing'],
                                    sec['missing_normal'],
                                    sec['missing_signal']
                                ])

            # Лист с лишними событиями (если есть)
            if any(detail['extra_events'] for detail in self.results['minute_details']):
                ws_extra = wb.create_sheet("Лишние события")

                ws_extra.append([
                    "Минута", "ППС", "Секунда",
                    "Время события", "Тип события",
                    "Ожидалось событий", "Фактически событий",
                    "Начало окна", "Конец окна",
                    "Тип проверки"
                ])

                # Форматирование заголовков
                for row in ws_extra.iter_rows(min_row=1, max_row=1, max_col=ws_extra.max_column):
                    for cell in row:
                        cell.font = header_font
                        cell.fill = header_fill
                        cell.font = white_font

                # Добавление данных о лишних событиях
                for detail in self.results['minute_details']:
                    if detail['extra_events']:
                        for event in detail['extra_events']:
                            ws_extra.append([
                                detail['minute'],
                                event['pps_num'],
                                event['second'],
                                event['time'].strftime("%H:%M:%S.%f")[:-3],
                                event['event_type'],
                                event['expected'],
                                event['actual'],
                                event['window_start'].strftime("%H:%M:%S.%f")[:-3],
                                event['window_end'].strftime("%H:%M:%S.%f")[:-3],
                                "Продвинутая" if event['is_advanced'] else "Обычная"
                            ])

                # Цветовое выделение для лишних событий
                for row in ws_extra.iter_rows(min_row=2, max_row=ws_extra.max_row):
                    for cell in row:
                        cell.fill = warning_fill

            # Автоматическая настройка ширины столбцов
            for sheet in wb.worksheets:
                for col in sheet.columns:
                    max_length = 0
                    column = col[0].column_letter

                    for cell in col:
                        try:
                            if len(str(cell.value)) > max_length:
                                max_length = len(str(cell.value))
                        except:
                            pass

                    adjusted_width = (max_length + 2) * 1.2
                    sheet.column_dimensions[column].width = adjusted_width

            wb.save(save_path)
            messagebox.showinfo("Успех", f"Отчет успешно сохранен в {save_path}")
            self.status_bar.config(text=f"Отчет сохранен: {os.path.basename(save_path)}")

        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось сохранить отчет: {str(e)}")
            self.status_bar.config(text=f"Ошибка сохранения: {str(e)}")


if __name__ == "__main__":
    root = tk.Tk()
    app = PPSEventAnalyzer(root)
    root.mainloop()