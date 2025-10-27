#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
data_processor_gui.py - Графічний інтерфейс для обробника даних
================================================================

GUI версія універсального обробника CSV та Excel файлів
Використовує tkinter для створення зручного інтерфейсу

Автор: Об'єднання скриптів + GUI
"""

import sys
import os
import threading
from pathlib import Path
from typing import List, Optional
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
from datetime import datetime

# Імпорт функцій з основного модуля
try:
    from data_processor import (
        csv_to_xlsx, xlsx_to_csv, read_file_auto,
        frequency_analysis, unique_values, deduplicate,
        merge_files, apply_filters, save_to_excel,
        detect_encoding_and_sep, FILTERS_HELP
    )
except ImportError:
    messagebox.showerror("Помилка", "Не знайдено data_processor.py!\nПереконайтесь що файл знаходиться в тій же папці.")
    sys.exit(1)

import pandas as pd


class DataProcessorGUI:
    """Головний клас GUI додатку"""

    def __init__(self, root):
        self.root = root
        self.root.title("Обробник CSV та Excel файлів v1.0")
        self.root.geometry("900x700")

        # Змінні
        self.input_files = []
        self.output_file = tk.StringVar()
        self.status_text = tk.StringVar(value="Готовий до роботи")

        # Створення інтерфейсу
        self.create_menu()
        self.create_main_interface()
        self.create_status_bar()

        # Центрування вікна
        self.center_window()

    def center_window(self):
        """Центрування вікна на екрані"""
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f'{width}x{height}+{x}+{y}')

    def create_menu(self):
        """Створення меню"""
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)

        # Файл
        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Файл", menu=file_menu)
        file_menu.add_command(label="Вибрати файли...", command=self.select_files)
        file_menu.add_separator()
        file_menu.add_command(label="Вихід", command=self.root.quit)

        # Довідка
        help_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Довідка", menu=help_menu)
        help_menu.add_command(label="Про програму", command=self.show_about)
        help_menu.add_command(label="Фільтри", command=self.show_filters_help)

    def create_main_interface(self):
        """Створення головного інтерфейсу з вкладками"""
        # Notebook для вкладок
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # Вкладки
        self.create_convert_tab()
        self.create_filter_tab()
        self.create_analysis_tab()
        self.create_merge_tab()
        self.create_deduplicate_tab()
        self.create_info_tab()

    def create_convert_tab(self):
        """Вкладка конвертації"""
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="Конвертація")

        # Заголовок
        ttk.Label(tab, text="Конвертація CSV ↔ XLSX",
                 font=('Arial', 14, 'bold')).pack(pady=10)

        # Фрейм вибору файлів
        file_frame = ttk.LabelFrame(tab, text="Файли", padding=10)
        file_frame.pack(fill=tk.X, padx=10, pady=5)

        # Вхідний файл
        ttk.Label(file_frame, text="Вхідний файл:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.convert_input = tk.StringVar()
        ttk.Entry(file_frame, textvariable=self.convert_input, width=50).grid(row=0, column=1, padx=5)
        ttk.Button(file_frame, text="Вибрати...",
                  command=lambda: self.browse_file(self.convert_input)).grid(row=0, column=2)

        # Вихідний файл
        ttk.Label(file_frame, text="Вихідний файл:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.convert_output = tk.StringVar()
        ttk.Entry(file_frame, textvariable=self.convert_output, width=50).grid(row=1, column=1, padx=5)
        ttk.Button(file_frame, text="Вибрати...",
                  command=lambda: self.browse_save_file(self.convert_output)).grid(row=1, column=2)

        # Опції
        options_frame = ttk.LabelFrame(tab, text="Опції (для CSV)", padding=10)
        options_frame.pack(fill=tk.X, padx=10, pady=5)

        ttk.Label(options_frame, text="Кодування:").grid(row=0, column=0, sticky=tk.W)
        self.encoding_var = tk.StringVar(value="auto")
        encodings = ["auto", "utf-8", "utf-8-sig", "cp1251", "windows-1251", "latin-1"]
        ttk.Combobox(options_frame, textvariable=self.encoding_var,
                    values=encodings, width=15, state="readonly").grid(row=0, column=1, padx=5)

        ttk.Label(options_frame, text="Роздільник:").grid(row=0, column=2, sticky=tk.W, padx=(20,0))
        self.separator_var = tk.StringVar(value="auto")
        separators = ["auto", ",", ";", "tab", "|"]
        ttk.Combobox(options_frame, textvariable=self.separator_var,
                    values=separators, width=10, state="readonly").grid(row=0, column=3, padx=5)

        # Кнопки дій
        button_frame = ttk.Frame(tab)
        button_frame.pack(pady=20)

        ttk.Button(button_frame, text="CSV → XLSX",
                  command=self.convert_csv_to_xlsx, width=20).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="XLSX → CSV",
                  command=self.convert_xlsx_to_csv, width=20).pack(side=tk.LEFT, padx=5)

        # Лог
        log_frame = ttk.LabelFrame(tab, text="Лог операцій", padding=5)
        log_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        self.convert_log = scrolledtext.ScrolledText(log_frame, height=10, wrap=tk.WORD)
        self.convert_log.pack(fill=tk.BOTH, expand=True)

    def create_filter_tab(self):
        """Вкладка фільтрації"""
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="Фільтрація")

        ttk.Label(tab, text="Фільтрація даних",
                 font=('Arial', 14, 'bold')).pack(pady=10)

        # Файли
        file_frame = ttk.LabelFrame(tab, text="Файли", padding=10)
        file_frame.pack(fill=tk.X, padx=10, pady=5)

        ttk.Label(file_frame, text="Вхідний файл:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.filter_input = tk.StringVar()
        ttk.Entry(file_frame, textvariable=self.filter_input, width=50).grid(row=0, column=1, padx=5)
        ttk.Button(file_frame, text="Вибрати...",
                  command=lambda: self.browse_file(self.filter_input)).grid(row=0, column=2)

        ttk.Label(file_frame, text="Вихідний файл:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.filter_output = tk.StringVar()
        ttk.Entry(file_frame, textvariable=self.filter_output, width=50).grid(row=1, column=1, padx=5)
        ttk.Button(file_frame, text="Вибрати...",
                  command=lambda: self.browse_save_file(self.filter_output)).grid(row=1, column=2)

        # Фільтр
        filter_frame = ttk.LabelFrame(tab, text="Налаштування фільтра", padding=10)
        filter_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        # Колонка
        ttk.Label(filter_frame, text="Колонка:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.filter_column = tk.StringVar()
        ttk.Entry(filter_frame, textvariable=self.filter_column, width=30).grid(row=0, column=1, pady=5)

        # Тип фільтра
        ttk.Label(filter_frame, text="Тип фільтра:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.filter_type = tk.StringVar(value="2")
        filter_types = [
            ("1 - Текст дорівнює", "1"),
            ("2 - Текст містить", "2"),
            ("3 - Текст у списку", "3"),
            ("4 - Число дорівнює", "4"),
            ("5 - Число в діапазоні", "5"),
            ("6 - REGEX", "6")
        ]

        combo_frame = ttk.Frame(filter_frame)
        combo_frame.grid(row=1, column=1, sticky=tk.W, pady=5)
        ttk.Combobox(combo_frame, textvariable=self.filter_type,
                    values=[f[0] for f in filter_types], width=30, state="readonly").pack()

        # Значення
        ttk.Label(filter_frame, text="Значення:").grid(row=2, column=0, sticky=tk.W, pady=5)
        self.filter_value = tk.StringVar()
        ttk.Entry(filter_frame, textvariable=self.filter_value, width=30).grid(row=2, column=1, pady=5)

        # Опції
        ttk.Label(filter_frame, text="Регістр:").grid(row=3, column=0, sticky=tk.W, pady=5)
        self.filter_case = tk.StringVar(value="upper")
        ttk.Combobox(filter_frame, textvariable=self.filter_case,
                    values=["keep", "upper", "lower"], width=15, state="readonly").grid(row=3, column=1, sticky=tk.W, pady=5)

        # Кнопка
        ttk.Button(tab, text="Застосувати фільтр",
                  command=self.apply_filter, width=30).pack(pady=20)

        # Лог
        log_frame = ttk.LabelFrame(tab, text="Результат", padding=5)
        log_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        self.filter_log = scrolledtext.ScrolledText(log_frame, height=8, wrap=tk.WORD)
        self.filter_log.pack(fill=tk.BOTH, expand=True)

    def create_analysis_tab(self):
        """Вкладка аналізу"""
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="Аналіз")

        ttk.Label(tab, text="Аналіз даних",
                 font=('Arial', 14, 'bold')).pack(pady=10)

        # Файл
        file_frame = ttk.LabelFrame(tab, text="Файл", padding=10)
        file_frame.pack(fill=tk.X, padx=10, pady=5)

        ttk.Label(file_frame, text="Файл для аналізу:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.analysis_input = tk.StringVar()
        ttk.Entry(file_frame, textvariable=self.analysis_input, width=50).grid(row=0, column=1, padx=5)
        ttk.Button(file_frame, text="Вибрати...",
                  command=lambda: self.browse_file(self.analysis_input)).grid(row=0, column=2)

        # Параметри
        params_frame = ttk.LabelFrame(tab, text="Параметри", padding=10)
        params_frame.pack(fill=tk.X, padx=10, pady=5)

        ttk.Label(params_frame, text="Колонка:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.analysis_column = tk.StringVar()
        ttk.Entry(params_frame, textvariable=self.analysis_column, width=30).grid(row=0, column=1, pady=5)

        ttk.Label(params_frame, text="Тип аналізу:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.analysis_type = tk.StringVar(value="freq")
        ttk.Radiobutton(params_frame, text="Частотний аналіз",
                       variable=self.analysis_type, value="freq").grid(row=1, column=1, sticky=tk.W)
        ttk.Radiobutton(params_frame, text="Унікальні значення",
                       variable=self.analysis_type, value="unique").grid(row=2, column=1, sticky=tk.W)

        # Кнопки
        button_frame = ttk.Frame(tab)
        button_frame.pack(pady=20)

        ttk.Button(button_frame, text="Виконати аналіз",
                  command=self.run_analysis, width=20).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Зберегти результат",
                  command=self.save_analysis, width=20).pack(side=tk.LEFT, padx=5)

        # Результат
        result_frame = ttk.LabelFrame(tab, text="Результат", padding=5)
        result_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        self.analysis_result = scrolledtext.ScrolledText(result_frame, height=12, wrap=tk.WORD)
        self.analysis_result.pack(fill=tk.BOTH, expand=True)

        # Зберігаємо результат для збереження
        self.last_analysis_df = None

    def create_merge_tab(self):
        """Вкладка об'єднання"""
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="Об'єднання")

        ttk.Label(tab, text="Об'єднання файлів",
                 font=('Arial', 14, 'bold')).pack(pady=10)

        # Список файлів
        files_frame = ttk.LabelFrame(tab, text="Файли для об'єднання", padding=10)
        files_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        # Listbox з прокруткою
        list_frame = ttk.Frame(files_frame)
        list_frame.pack(fill=tk.BOTH, expand=True)

        scrollbar = ttk.Scrollbar(list_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.merge_files_list = tk.Listbox(list_frame, yscrollcommand=scrollbar.set, height=8)
        self.merge_files_list.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.merge_files_list.yview)

        # Кнопки управління списком
        list_buttons = ttk.Frame(files_frame)
        list_buttons.pack(pady=5)

        ttk.Button(list_buttons, text="Додати файли",
                  command=self.add_merge_files).pack(side=tk.LEFT, padx=5)
        ttk.Button(list_buttons, text="Видалити",
                  command=self.remove_merge_file).pack(side=tk.LEFT, padx=5)
        ttk.Button(list_buttons, text="Очистити",
                  command=self.clear_merge_files).pack(side=tk.LEFT, padx=5)

        # Вихідний файл
        output_frame = ttk.LabelFrame(tab, text="Вихідний файл", padding=10)
        output_frame.pack(fill=tk.X, padx=10, pady=5)

        self.merge_output = tk.StringVar()
        ttk.Entry(output_frame, textvariable=self.merge_output, width=50).pack(side=tk.LEFT, padx=5)
        ttk.Button(output_frame, text="Вибрати...",
                  command=lambda: self.browse_save_file(self.merge_output)).pack(side=tk.LEFT)

        # Опції
        options_frame = ttk.LabelFrame(tab, text="Опції", padding=10)
        options_frame.pack(fill=tk.X, padx=10, pady=5)

        self.merge_dedupe = tk.BooleanVar(value=False)
        ttk.Checkbutton(options_frame, text="Видалити дублікати",
                       variable=self.merge_dedupe).grid(row=0, column=0, sticky=tk.W)

        ttk.Label(options_frame, text="Ключові колонки (через кому):").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.merge_keys = tk.StringVar()
        ttk.Entry(options_frame, textvariable=self.merge_keys, width=40).grid(row=1, column=1, pady=5)

        # Кнопка
        ttk.Button(tab, text="Об'єднати файли",
                  command=self.merge_files_action, width=30).pack(pady=10)

    def create_deduplicate_tab(self):
        """Вкладка дедуплікації"""
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="Дедуплікація")

        ttk.Label(tab, text="Видалення дублікатів",
                 font=('Arial', 14, 'bold')).pack(pady=10)

        # Файли
        file_frame = ttk.LabelFrame(tab, text="Файли", padding=10)
        file_frame.pack(fill=tk.X, padx=10, pady=5)

        ttk.Label(file_frame, text="Вхідний файл:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.dedup_input = tk.StringVar()
        ttk.Entry(file_frame, textvariable=self.dedup_input, width=50).grid(row=0, column=1, padx=5)
        ttk.Button(file_frame, text="Вибрати...",
                  command=lambda: self.browse_file(self.dedup_input)).grid(row=0, column=2)

        ttk.Label(file_frame, text="Вихідний файл:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.dedup_output = tk.StringVar()
        ttk.Entry(file_frame, textvariable=self.dedup_output, width=50).grid(row=1, column=1, padx=5)
        ttk.Button(file_frame, text="Вибрати...",
                  command=lambda: self.browse_save_file(self.dedup_output)).grid(row=1, column=2)

        # Параметри
        params_frame = ttk.LabelFrame(tab, text="Параметри", padding=10)
        params_frame.pack(fill=tk.X, padx=10, pady=5)

        ttk.Label(params_frame, text="Ключові колонки (через кому):").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.dedup_keys = tk.StringVar()
        ttk.Entry(params_frame, textvariable=self.dedup_keys, width=50).grid(row=0, column=1, pady=5, padx=5)

        ttk.Label(params_frame, text="Який запис залишити:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.dedup_keep = tk.StringVar(value="first")
        ttk.Combobox(params_frame, textvariable=self.dedup_keep,
                    values=["first", "last"], width=15, state="readonly").grid(row=1, column=1, sticky=tk.W, pady=5, padx=5)

        self.dedup_normalize = tk.BooleanVar(value=True)
        ttk.Checkbutton(params_frame, text="Нормалізувати ключі (UPPER, без пробілів)",
                       variable=self.dedup_normalize).grid(row=2, column=0, columnspan=2, sticky=tk.W, pady=5)

        # Кнопка
        ttk.Button(tab, text="Видалити дублікати",
                  command=self.deduplicate_action, width=30).pack(pady=20)

        # Лог
        log_frame = ttk.LabelFrame(tab, text="Результат", padding=5)
        log_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        self.dedup_log = scrolledtext.ScrolledText(log_frame, height=10, wrap=tk.WORD)
        self.dedup_log.pack(fill=tk.BOTH, expand=True)

    def create_info_tab(self):
        """Вкладка інформації про файл"""
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="Інфо")

        ttk.Label(tab, text="Інформація про файл",
                 font=('Arial', 14, 'bold')).pack(pady=10)

        # Файл
        file_frame = ttk.LabelFrame(tab, text="Файл", padding=10)
        file_frame.pack(fill=tk.X, padx=10, pady=5)

        self.info_file = tk.StringVar()
        ttk.Entry(file_frame, textvariable=self.info_file, width=60).pack(side=tk.LEFT, padx=5)
        ttk.Button(file_frame, text="Вибрати...",
                  command=lambda: self.browse_file(self.info_file)).pack(side=tk.LEFT)
        ttk.Button(file_frame, text="Аналізувати",
                  command=self.show_file_info).pack(side=tk.LEFT, padx=5)

        # Інформація
        info_frame = ttk.LabelFrame(tab, text="Деталі", padding=5)
        info_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        self.info_text = scrolledtext.ScrolledText(info_frame, height=20, wrap=tk.WORD)
        self.info_text.pack(fill=tk.BOTH, expand=True)

    def create_status_bar(self):
        """Створення статус-бару"""
        status_frame = ttk.Frame(self.root, relief=tk.SUNKEN)
        status_frame.pack(side=tk.BOTTOM, fill=tk.X)

        ttk.Label(status_frame, textvariable=self.status_text,
                 relief=tk.SUNKEN).pack(side=tk.LEFT, fill=tk.X, expand=True)

        # Прогрес-бар
        self.progress = ttk.Progressbar(status_frame, mode='indeterminate', length=100)
        self.progress.pack(side=tk.RIGHT, padx=5)

    # Допоміжні методи

    def browse_file(self, var):
        """Вибір файлу"""
        filename = filedialog.askopenfilename(
            title="Виберіть файл",
            filetypes=[
                ("Всі підтримувані", "*.csv;*.xlsx;*.xls"),
                ("CSV файли", "*.csv"),
                ("Excel файли", "*.xlsx;*.xls"),
                ("Всі файли", "*.*")
            ]
        )
        if filename:
            var.set(filename)

    def browse_save_file(self, var):
        """Вибір файлу для збереження"""
        filename = filedialog.asksaveasfilename(
            title="Зберегти як",
            defaultextension=".xlsx",
            filetypes=[
                ("Excel файли", "*.xlsx"),
                ("CSV файли", "*.csv"),
                ("Всі файли", "*.*")
            ]
        )
        if filename:
            var.set(filename)

    def add_merge_files(self):
        """Додати файли до списку об'єднання"""
        filenames = filedialog.askopenfilenames(
            title="Виберіть файли",
            filetypes=[
                ("Всі підтримувані", "*.csv;*.xlsx;*.xls"),
                ("CSV файли", "*.csv"),
                ("Excel файли", "*.xlsx;*.xls"),
                ("Всі файли", "*.*")
            ]
        )
        for f in filenames:
            self.merge_files_list.insert(tk.END, f)

    def remove_merge_file(self):
        """Видалити вибраний файл зі списку"""
        selection = self.merge_files_list.curselection()
        if selection:
            self.merge_files_list.delete(selection[0])

    def clear_merge_files(self):
        """Очистити список файлів"""
        self.merge_files_list.delete(0, tk.END)

    def log_message(self, widget, message):
        """Додати повідомлення в лог"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        widget.insert(tk.END, f"[{timestamp}] {message}\n")
        widget.see(tk.END)
        self.root.update()

    def set_status(self, message):
        """Встановити статус"""
        self.status_text.set(message)
        self.root.update()

    def start_progress(self):
        """Запустити прогрес-бар"""
        self.progress.start(10)

    def stop_progress(self):
        """Зупинити прогрес-бар"""
        self.progress.stop()

    # Обробники дій

    def convert_csv_to_xlsx(self):
        """Конвертація CSV → XLSX"""
        input_file = self.convert_input.get()
        output_file = self.convert_output.get()

        if not input_file:
            messagebox.showerror("Помилка", "Виберіть вхідний файл!")
            return

        if not output_file:
            output_file = str(Path(input_file).with_suffix('.xlsx'))
            self.convert_output.set(output_file)

        def task():
            try:
                self.start_progress()
                self.set_status("Конвертація CSV → XLSX...")
                self.log_message(self.convert_log, f"Початок конвертації: {input_file}")

                enc = None if self.encoding_var.get() == "auto" else self.encoding_var.get()
                sep = None if self.separator_var.get() == "auto" else self.separator_var.get()
                if sep == "tab":
                    sep = "\t"

                csv_to_xlsx(input_file, output_file, encoding=enc, separator=sep)

                self.log_message(self.convert_log, f"✓ Успішно створено: {output_file}")
                self.set_status("Готово")
                messagebox.showinfo("Успіх", f"Файл створено:\n{output_file}")

            except Exception as e:
                self.log_message(self.convert_log, f"✗ Помилка: {e}")
                messagebox.showerror("Помилка", str(e))
            finally:
                self.stop_progress()

        threading.Thread(target=task, daemon=True).start()

    def convert_xlsx_to_csv(self):
        """Конвертація XLSX → CSV"""
        input_file = self.convert_input.get()
        output_file = self.convert_output.get()

        if not input_file:
            messagebox.showerror("Помилка", "Виберіть вхідний файл!")
            return

        if not output_file:
            output_file = str(Path(input_file).with_suffix('.csv'))
            self.convert_output.set(output_file)

        def task():
            try:
                self.start_progress()
                self.set_status("Конвертація XLSX → CSV...")
                self.log_message(self.convert_log, f"Початок конвертації: {input_file}")

                xlsx_to_csv(input_file, output_file)

                self.log_message(self.convert_log, f"✓ Успішно створено: {output_file}")
                self.set_status("Готово")
                messagebox.showinfo("Успіх", f"Файл створено:\n{output_file}")

            except Exception as e:
                self.log_message(self.convert_log, f"✗ Помилка: {e}")
                messagebox.showerror("Помилка", str(e))
            finally:
                self.stop_progress()

        threading.Thread(target=task, daemon=True).start()

    def apply_filter(self):
        """Застосування фільтра"""
        input_file = self.filter_input.get()
        output_file = self.filter_output.get()

        if not input_file or not output_file:
            messagebox.showerror("Помилка", "Вкажіть вхідний та вихідний файли!")
            return

        def task():
            try:
                self.start_progress()
                self.set_status("Застосування фільтра...")
                self.log_message(self.filter_log, f"Читання файлу: {input_file}")

                df = read_file_auto(input_file)
                self.log_message(self.filter_log, f"Завантажено: {len(df):,} рядків")

                # Створення специфікації фільтра
                filter_type = self.filter_type.get()[0]  # Перша цифра
                spec = {
                    "mode": filter_type,
                    "column": self.filter_column.get(),
                    "case": self.filter_case.get(),
                    "strip_ws": True
                }

                if filter_type in ("1", "2"):
                    spec["value"] = self.filter_value.get()
                elif filter_type == "3":
                    spec["values"] = [v.strip() for v in self.filter_value.get().split(",")]
                elif filter_type == "4":
                    spec["value"] = self.filter_value.get()
                elif filter_type == "5":
                    parts = self.filter_value.get().split(",")
                    spec["min"] = parts[0] if len(parts) > 0 else "0"
                    spec["max"] = parts[1] if len(parts) > 1 else "999999"
                elif filter_type == "6":
                    spec["pattern"] = self.filter_value.get()

                result = apply_filters(df, [spec])
                self.log_message(self.filter_log, f"Після фільтрації: {len(result):,} рядків")

                # Збереження
                if output_file.endswith('.xlsx'):
                    save_to_excel(result, output_file)
                else:
                    result.to_csv(output_file, index=False, encoding='utf-8')

                self.log_message(self.filter_log, f"✓ Збережено: {output_file}")
                self.set_status("Готово")
                messagebox.showinfo("Успіх", f"Відфільтровано: {len(result):,} рядків\nЗбережено: {output_file}")

            except Exception as e:
                self.log_message(self.filter_log, f"✗ Помилка: {e}")
                messagebox.showerror("Помилка", str(e))
            finally:
                self.stop_progress()

        threading.Thread(target=task, daemon=True).start()

    def run_analysis(self):
        """Запуск аналізу"""
        input_file = self.analysis_input.get()
        column = self.analysis_column.get()

        if not input_file or not column:
            messagebox.showerror("Помилка", "Вкажіть файл та колонку!")
            return

        def task():
            try:
                self.start_progress()
                self.set_status("Виконання аналізу...")

                df = read_file_auto(input_file)

                if self.analysis_type.get() == "freq":
                    result = frequency_analysis(df, column)
                    title = "Частотний аналіз"
                else:
                    result = unique_values(df, column)
                    title = "Унікальні значення"

                # Відображення результату
                self.analysis_result.delete(1.0, tk.END)
                self.analysis_result.insert(tk.END, f"{title} для колонки: {column}\n")
                self.analysis_result.insert(tk.END, "=" * 60 + "\n\n")
                self.analysis_result.insert(tk.END, result.head(100).to_string(index=False))

                if len(result) > 100:
                    self.analysis_result.insert(tk.END, f"\n\n... показано перші 100 з {len(result)} записів")

                self.last_analysis_df = result
                self.set_status(f"Готово. Знайдено {len(result)} записів")

            except Exception as e:
                messagebox.showerror("Помилка", str(e))
            finally:
                self.stop_progress()

        threading.Thread(target=task, daemon=True).start()

    def save_analysis(self):
        """Збереження результатів аналізу"""
        if self.last_analysis_df is None:
            messagebox.showerror("Помилка", "Спочатку виконайте аналіз!")
            return

        filename = filedialog.asksaveasfilename(
            title="Зберегти результат",
            defaultextension=".xlsx",
            filetypes=[("Excel файли", "*.xlsx"), ("CSV файли", "*.csv")]
        )

        if filename:
            try:
                if filename.endswith('.xlsx'):
                    save_to_excel(self.last_analysis_df, filename, sheet_name="Analysis")
                else:
                    self.last_analysis_df.to_csv(filename, index=False, encoding='utf-8')
                messagebox.showinfo("Успіх", f"Результат збережено:\n{filename}")
            except Exception as e:
                messagebox.showerror("Помилка", str(e))

    def merge_files_action(self):
        """Об'єднання файлів"""
        files = list(self.merge_files_list.get(0, tk.END))
        output_file = self.merge_output.get()

        if len(files) < 2:
            messagebox.showerror("Помилка", "Додайте хоча б 2 файли для об'єднання!")
            return

        if not output_file:
            messagebox.showerror("Помилка", "Вкажіть вихідний файл!")
            return

        def task():
            try:
                self.start_progress()
                self.set_status("Об'єднання файлів...")

                dedup_keys = None
                if self.merge_dedupe.get():
                    keys_str = self.merge_keys.get()
                    if keys_str:
                        dedup_keys = [k.strip() for k in keys_str.split(",")]

                merge_files(files, output_file, deduplicate_keys=dedup_keys)

                self.set_status("Готово")
                messagebox.showinfo("Успіх", f"Файли об'єднано:\n{output_file}")

            except Exception as e:
                messagebox.showerror("Помилка", str(e))
            finally:
                self.stop_progress()

        threading.Thread(target=task, daemon=True).start()

    def deduplicate_action(self):
        """Дедуплікація"""
        input_file = self.dedup_input.get()
        output_file = self.dedup_output.get()
        keys_str = self.dedup_keys.get()

        if not input_file or not output_file or not keys_str:
            messagebox.showerror("Помилка", "Заповніть всі поля!")
            return

        def task():
            try:
                self.start_progress()
                self.set_status("Видалення дублікатів...")
                self.log_message(self.dedup_log, f"Читання файлу: {input_file}")

                df = read_file_auto(input_file)
                self.log_message(self.dedup_log, f"Завантажено: {len(df):,} рядків")

                keys = [k.strip() for k in keys_str.split(",")]
                result = deduplicate(df, keys,
                                   normalize_keys=self.dedup_normalize.get(),
                                   keep=self.dedup_keep.get())

                self.log_message(self.dedup_log, f"Після дедуплікації: {len(result):,} рядків")

                if output_file.endswith('.xlsx'):
                    save_to_excel(result, output_file)
                else:
                    result.to_csv(output_file, index=False, encoding='utf-8')

                self.log_message(self.dedup_log, f"✓ Збережено: {output_file}")
                self.set_status("Готово")
                messagebox.showinfo("Успіх",
                    f"Видалено дублікатів: {len(df) - len(result):,}\n"
                    f"Залишилось: {len(result):,} рядків\n"
                    f"Збережено: {output_file}")

            except Exception as e:
                self.log_message(self.dedup_log, f"✗ Помилка: {e}")
                messagebox.showerror("Помилка", str(e))
            finally:
                self.stop_progress()

        threading.Thread(target=task, daemon=True).start()

    def show_file_info(self):
        """Показати інформацію про файл"""
        file_path = self.info_file.get()

        if not file_path:
            messagebox.showerror("Помилка", "Виберіть файл!")
            return

        def task():
            try:
                self.start_progress()
                self.set_status("Аналіз файлу...")

                p = Path(file_path)
                info_text = f"{'='*60}\n"
                info_text += f"Файл: {p.name}\n"
                info_text += f"Розмір: {p.stat().st_size / 1024 / 1024:.2f} МБ\n"
                info_text += f"{'='*60}\n\n"

                if p.suffix.lower() == '.csv':
                    enc, sep = detect_encoding_and_sep(file_path)
                    info_text += f"Кодування: {enc}\n"
                    info_text += f"Роздільник: '{sep}'\n\n"

                df = read_file_auto(file_path)

                info_text += f"Рядків: {len(df):,}\n"
                info_text += f"Колонок: {len(df.columns)}\n\n"

                info_text += "Колонки:\n"
                for i, col in enumerate(df.columns, 1):
                    info_text += f"  {i:3d}. {col}\n"

                info_text += f"\n{'='*60}\n"
                info_text += "Перші 10 рядків:\n"
                info_text += f"{'='*60}\n\n"
                info_text += df.head(10).to_string(index=False)

                info_text += f"\n\n{'='*60}\n"
                info_text += "Типи даних:\n"
                info_text += f"{'='*60}\n\n"
                info_text += df.dtypes.to_string()

                self.info_text.delete(1.0, tk.END)
                self.info_text.insert(tk.END, info_text)
                self.set_status("Готово")

            except Exception as e:
                messagebox.showerror("Помилка", str(e))
            finally:
                self.stop_progress()

        threading.Thread(target=task, daemon=True).start()

    def show_about(self):
        """Про програму"""
        about_text = """
Обробник CSV та Excel файлів v1.0

Універсальний інструмент для роботи з даними

Можливості:
• Конвертація CSV ↔ XLSX
• Фільтрація даних (6 типів фільтрів)
• Частотний аналіз та унікальні значення
• Об'єднання файлів та аркушів
• Видалення дублікатів

Автор: Об'єднання скриптів
Версія: 1.0
Дата: 2025
        """
        messagebox.showinfo("Про програму", about_text)

    def show_filters_help(self):
        """Довідка по фільтрам"""
        messagebox.showinfo("Довідка по фільтрам", FILTERS_HELP)


def main():
    """Головна функція"""
    root = tk.Tk()
    app = DataProcessorGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
