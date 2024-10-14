# fda_tab.py

import tkinter as tk
from tkinter import ttk, messagebox
import openpyxl
import os
import logging

from utils import format_amount, parse_input
from constants import FDA_TEMPLATE_PATH  # Убедитесь, что этот путь правильный

logger = logging.getLogger(__name__)

class FDATab:
    def __init__(self, parent, pda_data):
        self.parent = parent  # Это будет фрейм вкладки FDA
        self.pda_data = pda_data  # Данные из PDA
        self.entries = {}  # Словарь для хранения полей ввода
        self.create_widgets()

    def create_widgets(self):
        # Создаём фрейм для размещения элементов
        self.frame = ttk.Frame(self.parent)
        self.frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Создаём заголовок
        title_label = ttk.Label(self.frame, text="FDA", font=('Helvetica', 16, 'bold'))
        title_label.pack(pady=10)

        # Создаём фрейм для таблицы fees и dues
        self.table_frame = ttk.Frame(self.frame)
        self.table_frame.pack(fill=tk.BOTH, expand=True)

        # Создаём заголовки колонок
        headers = ["Наименование", "PDA", "FDA (ввод)"]
        for col, header in enumerate(headers):
            label = ttk.Label(self.table_frame, text=header, font=('Helvetica', 12, 'bold'))
            label.grid(row=0, column=col, padx=5, pady=5)

        # Заполняем таблицу
        self.populate_table()

        # Кнопка "Сформировать FDA"
        generate_button = ttk.Button(self.frame, text="Сформировать FDA", command=self.generate_fda)
        generate_button.pack(pady=10)

    def populate_table(self):
        logger.info(f"Populating FDA table with data: {self.pda_data}")
        # Очищаем предыдущие данные
        for widget in self.table_frame.winfo_children():
            widget.destroy()

        # Создаём заголовки колонок
        headers = ["Наименование", "PDA", "FDA (ввод)"]
        for col, header in enumerate(headers):
            label = ttk.Label(self.table_frame, text=header, font=('Helvetica', 12, 'bold'))
            label.grid(row=0, column=col, padx=5, pady=5)

        self.entries = {}  # Сбросить словарь полей ввода

        # Заполняем таблицу актуальными данными
        for row, item in enumerate(self.pda_data, start=1):
            name = item['name']
            pda_value = item['amount']

            # Наименование
            name_label = ttk.Label(self.table_frame, text=name)
            name_label.grid(row=row, column=0, padx=5, pady=5, sticky='w')

            # Значение из PDA (неизменяемое)
            pda_label = ttk.Label(self.table_frame, text=format_amount(pda_value))
            pda_label.grid(row=row, column=1, padx=5, pady=5)

            # Поле ввода для FDA
            fda_entry = ttk.Entry(self.table_frame)
            fda_entry.grid(row=row, column=2, padx=5, pady=5)
            self.entries[name] = fda_entry

    def update_pda_data(self, pda_data):
        """Обновляет данные PDA и перезаполняет таблицу."""
        self.pda_data = pda_data
        logger.info(f"Updating FDA tab with PDA data: {self.pda_data}")
        self.populate_table()

    def generate_fda(self):
        # Собираем данные из полей ввода
        fda_data = {}
        for name, entry in self.entries.items():
            value_str = entry.get().strip()
            if not value_str:
                messagebox.showwarning("Внимание", f"Пожалуйста, заполните поле для {name}.")
                return
            try:
                value = parse_input(value_str)
                fda_data[name] = value
            except ValueError:
                messagebox.showerror("Ошибка", f"Неверный формат числа для {name}.")
                return

        # Генерируем FDA Excel
        self.generate_fda_excel(fda_data)

    def generate_fda_excel(self, fda_data):
        # Проверяем наличие шаблона Excel
        if not os.path.exists(FDA_TEMPLATE_PATH):
            messagebox.showerror("Ошибка", f"Шаблон FDA не найден по пути {FDA_TEMPLATE_PATH}.")
            return

        # Загружаем шаблон
        try:
            wb = openpyxl.load_workbook(FDA_TEMPLATE_PATH)
            ws = wb.active
        except Exception as e:
            logger.error(f"Ошибка при загрузке шаблона FDA: {e}")
            messagebox.showerror("Ошибка", f"Не удалось загрузить шаблон FDA: {e}")
            return

        # Вставляем данные в шаблон
        for row in ws.iter_rows():
            for cell in row:
                if cell.value and isinstance(cell.value, str):
                    for name, value in fda_data.items():
                        placeholder = f"{{{{{name}}}}}"
                        if placeholder in cell.value:
                            cell.value = cell.value.replace(placeholder, format_amount(value))

        # Сохраняем файл
        save_path = tk.filedialog.asksaveasfilename(defaultextension=".xlsx",
                                                    filetypes=[("Excel files", "*.xlsx")],
                                                    title="Сохранить FDA")
        if not save_path:
            return

        try:
            wb.save(save_path)
            messagebox.showinfo("Успех", "FDA успешно сохранена.")
        except Exception as e:
            logger.error(f"Ошибка при сохранении FDA: {e}")
            messagebox.showerror("Ошибка", f"Не удалось сохранить FDA: {e}")
