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
    def __init__(
            self,
            name,
            coefficient,
            vat_applicable=False,
            vat_included=False,
            category='Dues',
            uses_miles=False
    ):
        """
        :param name: Название сбора.
        :param coefficient: Коэффициент для расчета.
        :param vat_applicable: Применяется ли VAT (по умолчанию False).
        :param vat_included: Включен ли VAT в коэффициент (по умолчанию False).
        :param category: Категория сбора (по умолчанию 'Dues').
        :param uses_miles: Используются ли мили в расчете (по умолчанию False).
        """
        self.name = name
        self.coefficient = coefficient
        self.vat_applicable = vat_applicable
        self.vat_included = vat_included
        self.category = category
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
                self.vat_amount = base_amount - vat_base
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
            fee = Fee(name=fee_name, coefficient=coefficient, vat_applicable=False, category='Dues')
            fee.calculate(cv)
            self.fees.append(fee)

        # Расчет сборов с VAT и учетом миль
        for fee_name, coefficient in FEES_WITH_VAT_WITH_MILES.items():
            fee = Fee(name=fee_name, coefficient=coefficient, vat_applicable=True, uses_miles=True, category='Dues')
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
            fee = Fee(name=fee_name, coefficient=coefficient, vat_applicable=True, vat_included=vat_included, category='Dues')
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
            fee = Fee(name=name, coefficient=0.0, vat_applicable=False, category='Dues')
            fee.total_amount = amount
            fee.amount = amount
            self.fees.append(fee)

        # Обработка Agency Fee и Bank Charges
        agency_fee_amount = parse_input(self.inputs['agency_fee'])
        agency_fee = Fee(name='Agency fee', coefficient=0.0, vat_applicable=False, category='Agency Fees')
        agency_fee.total_amount = agency_fee_amount
        agency_fee.amount = agency_fee_amount
        self.fees.append(agency_fee)

        bank_charges_amount = parse_input(self.inputs['bank_charges'])
        bank_charges = Fee(name='Bank charges', coefficient=0.0, vat_applicable=False, category='Agency Fees')
        bank_charges.total_amount = bank_charges_amount
        bank_charges.amount = bank_charges_amount
        self.fees.append(bank_charges)

        # Обработка дополнительных Fees
        self.additional_fees = []
        for fee_input in self.inputs.get('additional_fees', []):
            name = fee_input['name']
            amount = parse_input(fee_input['amount'])
            self.additional_fees.append({'name': name, 'amount': amount})
            fee = Fee(name=name, coefficient=0.0, vat_applicable=False, category='Agency Fees')
            fee.total_amount = amount
            fee.amount = amount
            self.fees.append(fee)

    def calculate_totals(self):
        # Суммируем все Dues
        self.subtotal_dues = sum(fee.total_amount for fee in self.fees if fee.category == 'Dues')
        self.total_vat = sum(fee.vat_amount for fee in self.fees if fee.vat_applicable)

        # Суммируем Agency Fees
        self.subtotal_agency_fees = sum(fee.total_amount for fee in self.fees if fee.category == 'Agency Fees')

        self.total_amount = self.subtotal_dues + self.subtotal_agency_fees

    def get_fees_and_dues(self):
        """Возвращает список словарей с информацией о fees и dues."""
        fees_and_dues = []
        for fee in self.fees:
            fees_and_dues.append({
                'name': fee.name,
                'amount': fee.total_amount,
                'category': fee.category  # 'Dues' или 'Agency Fees'
            })
        return fees_and_dues

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
                'total_fee': temp_calculator.subtotal_dues,
                'total_agency_fee': temp_calculator.subtotal_agency_fees,
                'grand_total': temp_calculator.total_amount
            }

    def get_fee_display_data(self):
        data = []
        for fee in self.fees:
            if fee.category == 'Dues':
                data.append(fee.get_display_values())
        return data
