# calculations.py

from constants import (
    FEES_WITHOUT_VAT,
    FEES_WITH_VAT_WITH_MILES,
    FEES_WITH_VAT_WITHOUT_MILES,
    FEES_WITH_INCLUDED_VAT,
    VAT_RATE
)
from utils import parse_input, format_amount, parse_overtime, ceil_value


class Fee:
    def __init__(self, name, coefficient, vat_applicable, vat_included=False, uses_miles=False):
        """
        :param name: Название сбора.
        :param coefficient: Коэффициент для расчета.
        :param vat_applicable: Применяется ли VAT.
        :param vat_included: Включен ли VAT в коэффициент.
        :param uses_miles: Используются ли мили в расчете.
        """
        self.name = name
        self.coefficient = coefficient
        self.vat_applicable = vat_applicable
        self.vat_included = vat_included
        self.uses_miles = uses_miles
        self.amount = 0.0
        self.vat_amount = 0.0
        self.total_amount = 0.0

    def calculate(self, cv, miles=1, overtime_percentage=0.0):
        base_amount = cv * self.coefficient
        if self.uses_miles:
            base_amount *= miles

        base_amount *= (1 + overtime_percentage)

        if self.vat_applicable:
            if self.vat_included:
                # VAT уже включен в коэффициент, поэтому не добавляем его к сумме
                vat_base = base_amount / (1 + VAT_RATE)
                self.vat_amount = vat_base * VAT_RATE
                self.total_amount = base_amount  # Общая сумма уже включает VAT
            else:
                # VAT не включен в коэффициент, добавляем VAT к сумме
                self.vat_amount = base_amount * VAT_RATE
                self.total_amount = base_amount + self.vat_amount
        else:
            self.vat_amount = 0.0
            self.total_amount = base_amount

        self.amount = base_amount

    def get_display_values(self):
        vat_display = format_amount(self.vat_amount) if self.vat_amount > 0 else "-"
        total_display = format_amount(self.total_amount)
        return self.name, vat_display, total_display


class FeeCalculator:
    def __init__(self, inputs):
        self.inputs = inputs
        self.cv = 0.0
        self.fees = []
        self.additional_dues = []
        self.additional_fees = []
        self.subtotal_dues = 0.0
        self.subtotal_agency_fees = 0.0
        self.total_vat = 0.0
        self.total_amount = 0.0
        # Добавляем атрибуты для фиксированных ставок овертайма
        self.fixed_overtime_rates = [0.25, 0.50, 1.00]  # 25%, 50%, 100%
        self.fixed_totals = {}

    def calculate_cv(self):
        lbp = parse_input(self.inputs['lbp'])
        beam = parse_input(self.inputs['beam'])
        rdm = parse_input(self.inputs['rdm'])
        cv = lbp * beam * rdm
        self.cv = ceil_value(cv)

    def calculate_fees(self):
        self.calculate_cv()
        cv = self.cv

        # Преобразование овертайма
        overtime_in_percentage = parse_overtime(self.inputs['overtime_in'])
        overtime_out_percentage = parse_overtime(self.inputs['overtime_out'])

        # Расчет сборов без VAT
        for fee_name, coefficient in FEES_WITHOUT_VAT.items():
            fee = Fee(fee_name, coefficient, vat_applicable=False)
            fee.calculate(cv)
            self.fees.append(fee)

        # Расчет сборов с VAT и учетом миль
        for fee_name, coefficient in FEES_WITH_VAT_WITH_MILES.items():
            fee = Fee(fee_name, coefficient, vat_applicable=True, uses_miles=True)
            if "in" in fee_name.lower():
                if "inward" in fee_name.lower():
                    miles = int(self.inputs['miles_inward_in'])
                    overtime_percentage = overtime_in_percentage
                else:
                    miles = int(self.inputs['miles_outward_in'])
                    overtime_percentage = overtime_in_percentage
            else:
                if "inward" in fee_name.lower():
                    miles = int(self.inputs['miles_inward_out'])
                    overtime_percentage = overtime_out_percentage
                else:
                    miles = int(self.inputs['miles_outward_out'])
                    overtime_percentage = overtime_out_percentage

            fee.calculate(cv, miles, overtime_percentage)
            self.fees.append(fee)

        # Расчет сборов с VAT без учета миль
        for fee_name, coefficient in FEES_WITH_VAT_WITHOUT_MILES.items():
            # Проверяем, включен ли VAT в коэффициент
            vat_included = fee_name in FEES_WITH_INCLUDED_VAT
            fee = Fee(fee_name, coefficient, vat_applicable=True, vat_included=vat_included)
            if "in" in fee_name.lower():
                overtime_percentage = overtime_in_percentage
            elif "out" in fee_name.lower():
                overtime_percentage = overtime_out_percentage
            else:
                overtime_percentage = 0.0

            fee.calculate(cv, overtime_percentage=overtime_percentage)
            self.fees.append(fee)

        # Обработка дополнительных Dues
        self.additional_dues = []
        for due in self.inputs.get('additional_dues', []):
            name = due['name']
            amount = parse_input(due['amount'])
            self.additional_dues.append({'name': name, 'amount': amount})
            # Создаем объект Fee для отображения в таблице
            fee = Fee(name, 0.0, vat_applicable=False)
            fee.total_amount = amount
            fee.amount = amount
            self.fees.append(fee)

        # Обработка дополнительных Fees
        self.additional_fees = []
        for fee in self.inputs.get('additional_fees', []):
            name = fee['name']
            amount = parse_input(fee['amount'])
            self.additional_fees.append({'name': name, 'amount': amount})

    def calculate_totals(self):
        # Суммируем все Dues
        self.subtotal_dues = sum(fee.total_amount for fee in self.fees if fee.name not in ['Agency fee', 'Bank charges'])
        self.total_vat = sum(fee.vat_amount for fee in self.fees)

        # Суммируем Agency fee и Bank charges
        agency_fee = parse_input(self.inputs['agency_fee'])
        bank_charges = parse_input(self.inputs['bank_charges'])
        self.subtotal_agency_fees = agency_fee + bank_charges

        # Добавляем дополнительные Fees
        for fee in self.additional_fees:
            self.subtotal_agency_fees += fee['amount']

        self.total_amount = self.subtotal_dues + self.subtotal_agency_fees

    def calculate_fixed_overtime_totals(self):
        for rate in self.fixed_overtime_rates:
            # Создаём копию входных данных с заданной ставкой овертайма
            fixed_inputs = self.inputs.copy()
            fixed_inputs['overtime_in'] = f"{int(rate * 100)}%"
            fixed_inputs['overtime_out'] = f"{int(rate * 100)}%"

            # Создаём временный калькулятор для расчёта
            temp_calculator = FeeCalculator(fixed_inputs)
            temp_calculator.calculate_fees()
            temp_calculator.calculate_totals()

            # Сохраняем результаты
            self.fixed_totals[rate] = {
                'total_fee': temp_calculator.subtotal_dues,  # Или другой соответствующий атрибут
                'total_agency_fee': temp_calculator.subtotal_agency_fees,
                'grand_total': temp_calculator.total_amount
            }

    def get_fee_display_data(self):
        data = []
        for fee in self.fees:
            data.append(fee.get_display_values())
        return data
