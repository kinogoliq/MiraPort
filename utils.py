# utils.py

import sys
import math
from tkinter import messagebox

def parse_input(value, is_percentage=False):
    """
    Преобразует строковое значение в число с плавающей точкой.

    :param value: Строковое значение, введенное пользователем.
    :param is_percentage: Флаг, указывающий, является ли значение процентом.
    :return: Число с плавающей точкой.
    :raises ValueError: Если значение не может быть преобразовано.
    """
    try:
        value = value.replace(',', '.')
        if is_percentage:
            return float(value) / 100
        return float(value)
    except ValueError:
        raise ValueError(f"Неверный формат числа: {value}")


def format_amount(amount):
    """Форматирует число по российским стандартам (разделитель тысяч - пробел, десятичный - запятая)."""
    return f"{amount:,.2f}".replace(",", " ").replace(".", ",")


def parse_overtime(overtime_str):
    """
    Преобразует строковое значение овертайма в десятичную дробь.
    Например, "25%" -> 0.25
    """
    if overtime_str.endswith('%'):
        try:
            value = float(overtime_str.strip('%')) / 100
            return value
        except ValueError:
            messagebox.showerror("Ошибка ввода", f"Некорректное значение овертайма: {overtime_str}")
            return 0.0
    else:
        return 0.0


def ceil_value(value):
    """Округляет значение вверх до ближайшего целого числа."""
    return math.ceil(value)
