# agency_fee.py

import tkinter as tk
from tkinter import ttk
from utils import format_amount
import logging

logger = logging.getLogger(__name__)

# Предварительно заполненный словарь с данными agency fee
# Ключи — диапазоны cv, значения — соответствующие agency fee
agency_fee_dict = {
    (0, 1800): 1194,
    (1801, 3600): 1478,
    (3601, 5500): 1764,
    (5501, 7200): 2076,
    (7201, 11000): 2446,
    (11001, 15000): 2816,
    (15001, 22000): 3242,
    (22001, 30000): 3754,
    (30001, 37000): 4210,
    (37001, 44000): 4636,
    (44001, 51000): 5120,
    (51001, 59000): 5574,
    (59001, 66000): 6030,
    (66001, 73000): 6512,
    (73001, 92000): 6969,
    (92001, float('inf')): 8172,
}

def calculate_cv(lbp, beam, rdm):
    """Рассчитывает cv на основе lbp, beam и rdm."""
    return lbp * beam * rdm

def get_agency_fee(cv):
    """Возвращает agency fee на основе cv из словаря."""
    for (min_cv, max_cv), fee in agency_fee_dict.items():
        if min_cv <= cv <= max_cv:
            return fee
    return None  # Если cv не попадает ни в один диапазон

def show_agency_fee_table(cv):
    """Отображает всплывающее окно с таблицей agency fee."""
    agency_fee = get_agency_fee(cv)

    if agency_fee is None:
        logger.error("CV не соответствует ни одному диапазону в словаре agency_fee_dict.")
        return

    # Создаём новое окно
    window = tk.Toplevel()
    window.title("Agency Fee Table")

    # Создаём таблицу
    columns = ("CV Range", "Agency Fee")
    tree = ttk.Treeview(window, columns=columns, show="headings")
    tree.heading("CV Range", text="Диапазон CV")
    tree.heading("Agency Fee", text="Agency Fee")

    # Заполняем таблицу данными из словаря
    for (min_cv, max_cv), fee in agency_fee_dict.items():
        cv_range = f"{min_cv} - {max_cv if max_cv != float('inf') else '∞'}"
        tree.insert("", "end", values=(cv_range, format_amount(fee)))

    # Выделяем строку, соответствующую текущему cv
    for item in tree.get_children():
        item_values = tree.item(item)['values']
        range_str = item_values[0]
        min_cv_str, max_cv_str = range_str.split(' - ')
        min_cv_val = float(min_cv_str)
        max_cv_val = float(max_cv_str.replace('∞', 'inf'))
        if min_cv_val <= cv <= max_cv_val:
            tree.selection_set(item)
            break

    tree.pack(fill=tk.BOTH, expand=True)

    # Добавляем кнопку для закрытия окна
    close_button = ttk.Button(window, text="Закрыть", command=window.destroy)
    close_button.pack(pady=10)
