# -*- coding: utf-8 -*-

from odoo import api, fields, models
from collections import defaultdict
import base64
from datetime import datetime
from io import BytesIO
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, NamedStyle
from openpyxl.utils import get_column_letter


class TrprovOverheadTr(models.TransientModel):
    _name = 'trprov.overhead.tr'
    _description = 'Reporte Overhead'

    # Definición de campos para el modelo
    report_from_date = fields.Date(string="Reporte desde", required=True, default=fields.Date.context_today)
    report_to_date = fields.Date(string="Reporte hasta", required=True, default=fields.Date.context_today)
    categ_ids = fields.Many2many('product.category', string="Categoria de los productos")
    res_seller_ids = fields.Many2many('account.analytic.account', string="Cuentas Analíticas", required=True)
    company_id = fields.Many2one(comodel_name="res.company", string="Compañia", required=True,
                                 default=lambda self: self.env.company.id)
    file_content = fields.Binary(string="Archivo Contenido")

    # Define la acción para generar el archivo Excel
    def action_generate_excel(self):
        # Crea un nuevo libro de trabajo de Excel
        wk_book = openpyxl.Workbook()
        # Define los estilos para el archivo
        currency_format = NamedStyle(name='currency', number_format='"$"#,##0.00')
        bold_font = Font(bold=True)
        header_font = Font(bold=True, color='FFFFFF')
        header_fill = PatternFill(start_color='000000', end_color='000000', fill_type='solid')

        # Obtiene los datos para el informe
        data = sorted(self.get_data_status_results(), key=lambda x: x['account_type'])

        # Elimina la hoja activa por defecto
        wk_book.remove(wk_book.active)

        # Define los datos del informe
        sheets = {}
        last_account_type = None
        separator_fill = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')

        # Itera sobre los datos y los añade al informe
        for item in data:
            analytic_account_name = item['analytic_account_name']

            # Verifica si la hoja para la cuenta analítica ya existe, si no, la crea
            if analytic_account_name not in sheets:
                sheet = wk_book.create_sheet(title=analytic_account_name)
                sheets[analytic_account_name] = sheet
                last_account_type = None

                # Agrega el título del informe
                report_title = f"Estado de resultados desde {self.report_from_date} hasta {self.report_to_date}"
                sheet.append(["", "", report_title])
                for cell in sheet["B2:F2"]:
                    cell[0].font = bold_font

                # Agrega el nombre de la compañía al informe
                company_name = self.env.company.name
                sheet.append(["", "Compañía:", company_name])
                for cell in sheet["B3:F3"]:
                    cell[0].font = bold_font

                # Agrega el nombre del Centro de Costo al informe
                sheet.append(["", "Nombre Cto Costo:", analytic_account_name])
                for cell in sheet["B4:F4"]:
                    cell[0].font = bold_font

                # Deja una fila vacía
                sheet.append([])

                # Agrega los encabezados de las columnas al informe
                headers = [
                    "Concepto", "Nombre Cto. Costo", "Nombre Cta Agrupador", "Enero", "Febrero",
                    "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre",
                    "Diciembre",
                    "Total Resultado"
                ]
                sheet.append(headers)
                for col_index, header in enumerate(headers, start=1):
                    cell = sheet.cell(row=6, column=col_index)
                    cell.value = header
                    cell.font = header_font
                    cell.fill = header_fill
            else:
                # Si la hoja ya existe, obtén la referencia
                sheet = sheets[analytic_account_name]

            # Agrega un separador si cambia el tipo de cuenta
            if last_account_type and item['account_type'] != last_account_type:
                row_index = sheet.max_row + 1
                for col in range(1, sheet.max_column + 1):
                    cell = sheet.cell(row=row_index, column=col)
                    cell.fill = separator_fill

            last_account_type = item['account_type']

            # Incrementa el índice de la fila
            row_index = sheet.max_row + 1

            # Llena las celdas con la información del item actual
            sheet.cell(row=row_index, column=1, value=item['account_type'])
            sheet.cell(row=row_index, column=2, value=item['analytic_account_name'])
            sheet.cell(row=row_index, column=3, value=item['financial_account_name'])

            # Llena las celdas de los meses con los valores correspondientes
            for month_index in range(1, 13):
                month_value = item.get(f'month_{month_index}', 0)
                sheet.cell(row=row_index, column=month_index + 3, value=month_value).style = currency_format

            # Llena la celda del total con el valor correspondiente
            sheet.cell(row=row_index, column=16, value=item['total_result']).style = currency_format

        # Guarda el informe en un objeto BytesIO
        output = BytesIO()
        wk_book.save(output)
        output.seek(0)
        base64_content = base64.b64encode(output.read())
        output.close()

        # Escribe el contenido del archivo en el campo 'file_content' del modelo
        self.write({'file_content': base64_content})

        # Crear formato para nombre de salida de archivo
        formatted_from_date = self.report_from_date.strftime('%Y-%m-%d')
        formatted_to_date = self.report_to_date.strftime('%Y-%m-%d')
        filename = f"Estado_resultados_{formatted_from_date}_{formatted_to_date}.xlsx"

        # Devuelve una acción para descargar el archivo
        return {
            'type': 'ir.actions.act_url',
            'url': f"web/content/?model={self._name}&id={self.id}&field=file_content&filename={filename}&download=true",
            'target': 'self',
        }

    # Define la función para obtener los datos del estado de resultados
    def get_data_status_results(self):
        # Define el mapeo de los tipos de cuenta
        ACCOUNT_TYPE_MAPPING = {
            "asset_receivable": "Por cobrar",
            "asset_cash": "Banco y efectivo",
            "asset_current": "Activos Circulantes",
            "asset_non_current": "Activos no-circulantes",
            "asset_prepayments": "Prepagos",
            "asset_fixed": "Activos Fijos",
            "liability_payable": "Por pagar",
            "liability_credit_card": "Tarjeta de Crédito",
            "liability_current": "Pasivos Circulantes",
            "liability_non_current": "Pasivos no-circulantes",
            "equity": "Capital",
            "equity_unaffected": "Ganancias del año actual",
            "income": "Ingreso",
            "income_other": "Otro Ingreso",
            "expense": "Gastos",
            "expense_depreciation": "Depreciación",
            "expense_direct_cost": "Costo de ingresos",
            "off_balance": "Hoja fuera de balance",
        }

        # (Código para obtener los datos del estado de resultados)
        final_data = []  # Lista que almacenará los datos finales a devolver

        # Bucle principal que recorre cada registro en el objeto actual
        for rec in self:
            # Estructuras de datos para almacenar información analítica y de cuentas
            data = defaultdict(lambda: defaultdict(lambda: defaultdict(float)))
            account_info = defaultdict(dict)

            # Buscar líneas analíticas que cumplan con ciertos criterios
            analytic_lines = rec.env['account.analytic.line'].search([
                ('account_id', 'in', rec.res_seller_ids.ids),
                ('date', '>=', rec.report_from_date),
                ('date', '<=', rec.report_to_date),
            ])

            # Iterar sobre las líneas analíticas encontradas
            for line in analytic_lines:
                analytic_account_name = line.account_id.name
                financial_account_name = line.general_account_id.name
                financial_account_type = line.trprovwi_general_account_type_tr
                analytic_account_id = line.account_id.id
                financial_account_id = line.general_account_id.id

                # Almacenar información de cuentas
                account_info[analytic_account_id]['name'] = analytic_account_name
                account_info[analytic_account_id][financial_account_id] = {
                    'name': financial_account_name,
                    'type': financial_account_type
                }

                # Calcular y almacenar datos mensuales
                for month in range(1, 13):
                    if line.date.month == month:
                        amount = line.amount or 0
                        data[analytic_account_id][financial_account_id][f'month_{month}'] += amount
                        data[analytic_account_id][financial_account_id]['total_result'] += amount

            # Construir la estructura final de datos a partir de la información recopilada
            for analytic_id, financial_accounts in data.items():
                for financial_id, months in financial_accounts.items():
                    account_type_key = account_info[analytic_id][financial_id]['type']
                    # Obtener el valor del tipo de cuenta o dejar el original si no hay mapeo
                    account_type_value = ACCOUNT_TYPE_MAPPING.get(account_type_key, account_type_key)
                    # Agregar los datos finales a la lista
                    final_data.append({
                        'analytic_account_id': analytic_id,
                        'analytic_account_name': account_info[analytic_id]['name'],
                        'financial_account_id': financial_id,
                        'financial_account_name': account_info[analytic_id][financial_id]['name'],
                        'account_type': account_type_value,
                        **months
                    })

        # Devolver la lista final de datos
        return final_data
