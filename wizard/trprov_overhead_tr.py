# -*- coding: utf-8 -*-

from odoo import api, fields, models
import openpyxl
from openpyxl.styles import Font, NamedStyle
from io import BytesIO
import base64
from collections import defaultdict


class TrprovOverheadTr(models.TransientModel):
    _name = 'trprov.overhead.tr'
    _description = 'Reporte Overhead'

    # Definición de campos para el modelo
    report_from_date = fields.Date(string="Reporte desde", required=True, default=fields.Date.context_today)
    report_to_date = fields.Date(string="Reporte hasta", required=True, default=fields.Date.context_today)
    categ_ids = fields.Many2many('product.category', string="Categoria de los productos")
    res_seller_ids = fields.Many2many('account.analytic.account', string="Cuentas Analíticas")
    company_id = fields.Many2one(comodel_name="res.company", string="Compañia", required=True,
                                 default=lambda self: self.env.company.id)
    file_content = fields.Binary(string="Archivo Contenido")

    # Acción para generar el archivo Excel y mostrar el enlace de descarga
    def action_generate_excel(self):
        wk_book = openpyxl.Workbook()
        wk_sheet = wk_book.active
        wk_sheet.title = "Estado de Resultados por Vendedor"

        # Agrega el título en la celda C2 a S2
        wk_sheet['C2'] = "Estado de Resultado por Vendedor con Overhead"
        wk_sheet.merge_cells('C2:S2')
        title_cell = wk_sheet['C2']
        title_cell.alignment = openpyxl.styles.Alignment(horizontal="center", vertical="center")
        title_cell.font = openpyxl.styles.Font(bold=True, size=14)

        headers = [
            "Nombre Cto. Costo", "Tipo", "Nombre Cuenta", "Enero", "Febrero", "Marzo",
            "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre",
            "Total Resultado"
        ]

        # Comienza a llenar los encabezados a partir de la tercera fila y tercera columna
        row_index = 3
        col_index = 1
        for header in headers:
            cell = wk_sheet.cell(row=row_index, column=col_index)
            cell.value = header
            cell.font = Font(bold=True)
            col_index += 1

        # Obtén el formato de moneda
        currency_format = NamedStyle(name='currency', number_format='"$"#,##0.00')

        data = self.get_data_status_results()

        for row_index, row_data in enumerate(data, 4):
            cost_cto_name = row_data['cost_cto_name']
            account_type = row_data['account_type']
            account_name = row_data['account_name']
            enero = row_data['enero']
            febrero = row_data['febrero']
            marzo = row_data['marzo']
            abril = row_data['abril']
            mayo = row_data['mayo']
            junio = row_data['junio']
            julio = row_data['julio']
            agosto = row_data['agosto']
            septiembre = row_data['septiembre']
            octubre = row_data['octubre']
            noviembre = row_data['noviembre']
            diciembre = row_data['diciembre']
            total_result = row_data['total_result']

            wk_sheet.cell(row=row_index, column=1, value=cost_cto_name)
            wk_sheet.cell(row=row_index, column=2, value=account_type)
            wk_sheet.cell(row=row_index, column=3, value=account_name)
            wk_sheet.cell(row=row_index, column=4, value=enero).style = currency_format
            wk_sheet.cell(row=row_index, column=5, value=febrero).style = currency_format
            wk_sheet.cell(row=row_index, column=6, value=marzo).style = currency_format
            wk_sheet.cell(row=row_index, column=7, value=abril).style = currency_format
            wk_sheet.cell(row=row_index, column=8, value=mayo).style = currency_format
            wk_sheet.cell(row=row_index, column=9, value=junio).style = currency_format
            wk_sheet.cell(row=row_index, column=10, value=julio).style = currency_format
            wk_sheet.cell(row=row_index, column=11, value=agosto).style = currency_format
            wk_sheet.cell(row=row_index, column=12, value=septiembre).style = currency_format
            wk_sheet.cell(row=row_index, column=13, value=octubre).style = currency_format
            wk_sheet.cell(row=row_index, column=14, value=noviembre).style = currency_format
            wk_sheet.cell(row=row_index, column=15, value=diciembre).style = currency_format
            wk_sheet.cell(row=row_index, column=16, value=total_result).style = currency_format
            wk_sheet.cell(row=row_index, column=16).font = Font(bold=True)  # Aplica negrita a la columna "Total Resultado"

        # Crear una segunda hoja llamada "HOJA 2"
        wk_sheet_2 = wk_book.create_sheet(title="Overhead")

        # Agregar encabezados a la segunda hoja
        headers_2 = [
            "Columna1", "Columna2", "Columna3",
            # ... Otros encabezados ...
        ]

        row_index_2 = 1
        col_index_2 = 1
        for header in headers_2:
            cell_2 = wk_sheet_2.cell(row=row_index_2, column=col_index_2)
            cell_2.value = header
            cell_2.font = Font(bold=True)
            col_index_2 += 1

        # Agregar datos a la segunda hoja
        data_2 = [
            #{"Columna1": valor1, "Columna2": valor2, "Columna3": valor3},
            # ... Otros datos ...
        ]

        for row_index_2, row_data_2 in enumerate(data_2, 2):
            for col_index_2, value_2 in enumerate(row_data_2.values(), 1):
                wk_sheet_2.cell(row=row_index_2, column=col_index_2, value=value_2)

        # Guardar los cambios
        output = BytesIO()
        wk_book.save(output)

        base64_content = base64.b64encode(output.getvalue())
        self.write({'file_content': base64_content})

        return {
            'type': 'ir.actions.act_url',
            'url': "web/content/?model={}&id={}&field=file_content&filename=Reporte_Overhead.xlsx&download=true".format(
                self._name, self.id),
            'target': 'self',
        }

    # Acción para ejecutar la acción de generación de Excel
    def action_custom_button(self):
        return self.action_generate_excel()

    # Función para obtener datos del estado de resultados desde Apuntes contables
    def get_data_status_results(self):
        for rec in self:
            data = defaultdict(lambda: defaultdict(float))

            analytic_sellers = rec.env['account.analytic.line'].search([
                ('account_id', 'in', rec.res_seller_ids.ids),
                ('date', '>=', rec.report_from_date),
                ('date', '<=', rec.report_to_date),
            ])

            for record in analytic_sellers:
                # Inicializar los valores del mes como 0.00
                enero = febrero = marzo = abril = mayo = junio = julio = agosto = septiembre = octubre = noviembre = diciembre = 0.00

                # Fijar el valor del importe del mes correspondiente
                if record.date.month == 1:
                    enero += record.amount
                elif record.date.month == 2:
                    febrero += record.amount
                elif record.date.month == 3:
                    marzo += record.amount
                elif record.date.month == 4:
                    abril += record.amount
                elif record.date.month == 5:
                    mayo += record.amount
                elif record.date.month == 6:
                    junio += record.amount
                elif record.date.month == 7:
                    julio += record.amount
                elif record.date.month == 8:
                    agosto += record.amount
                elif record.date.month == 9:
                    septiembre += record.amount
                elif record.date.month == 10:
                    octubre += record.amount
                elif record.date.month == 11:
                    noviembre += record.amount
                elif record.date.month == 12:
                    diciembre += record.amount

                # Suma los importes de cada mes para el mismo nombre de cuenta
                data[record.account_id.name]['enero'] += enero
                data[record.account_id.name]['febrero'] += febrero
                data[record.account_id.name]['marzo'] += marzo
                data[record.account_id.name]['abril'] += abril
                data[record.account_id.name]['mayo'] += mayo
                data[record.account_id.name]['junio'] += junio
                data[record.account_id.name]['julio'] += julio
                data[record.account_id.name]['agosto'] += agosto
                data[record.account_id.name]['septiembre'] += septiembre
                data[record.account_id.name]['octubre'] += octubre
                data[record.account_id.name]['noviembre'] += noviembre
                data[record.account_id.name]['diciembre'] += diciembre
                data[record.account_id.name]['total_result'] += record.amount

            final_data = []
            for account_name, months in data.items():
                total_result = months['total_result']
                final_data.append({
                    'account_type': record.general_account_id.account_type,
                    'account_name': record.general_account_id.name,
                    'cost_cto_name': record.account_id.name,
                    'enero': months['enero'],
                    'febrero': months['febrero'],
                    'marzo': months['marzo'],
                    'abril': months['abril'],
                    'mayo': months['mayo'],
                    'junio': months['junio'],
                    'julio': months['julio'],
                    'agosto': months['agosto'],
                    'septiembre': months['septiembre'],
                    'octubre': months['octubre'],
                    'noviembre': months['noviembre'],
                    'diciembre': months['diciembre'],
                    'total_result': total_result,
                })

            return final_data
