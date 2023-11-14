# -*- coding: utf-8 -*-

from odoo import api, fields, models
import openpyxl
from openpyxl.styles import Font
from io import BytesIO
import base64
from odoo.exceptions import ValidationError


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

        data = self.obtener_datos_estado_resultados()

        for row_index, row_data in enumerate(data, 4):
            nombre_cto_costo = row_data['nombre_cto_costo']
            tipo_cuenta = row_data['tipo_cuenta']
            nombre_cuenta = row_data['nombre_cuenta']
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
            total_resultado = row_data['total_resultado']

            wk_sheet.cell(row=row_index, column=1, value=nombre_cto_costo)
            wk_sheet.cell(row=row_index, column=2, value=tipo_cuenta)
            wk_sheet.cell(row=row_index, column=3, value=nombre_cuenta)
            wk_sheet.cell(row=row_index, column=4, value=enero)
            wk_sheet.cell(row=row_index, column=5, value=febrero)
            wk_sheet.cell(row=row_index, column=6, value=marzo)
            wk_sheet.cell(row=row_index, column=7, value=abril)
            wk_sheet.cell(row=row_index, column=8, value=mayo)
            wk_sheet.cell(row=row_index, column=9, value=junio)
            wk_sheet.cell(row=row_index, column=10, value=julio)
            wk_sheet.cell(row=row_index, column=11, value=agosto)
            wk_sheet.cell(row=row_index, column=12, value=septiembre)
            wk_sheet.cell(row=row_index, column=13, value=octubre)
            wk_sheet.cell(row=row_index, column=14, value=noviembre)
            wk_sheet.cell(row=row_index, column=15, value=diciembre)
            wk_sheet.cell(row=row_index, column=16, value=total_resultado).font = Font(bold=True)

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
    def obtener_datos_estado_resultados(self):
        for rec in self:
            data = []

            analytic_sellers = rec.env['account.analytic.line'].search([('account_id', 'in', rec.res_seller_ids.ids), ])

            for record in analytic_sellers:
                # Initialize the month values as 0.00
                enero = febrero = marzo = abril = mayo = junio = julio = agosto = septiembre = octubre = noviembre = diciembre = "0.00"

                # Set the amount value for the corresponding month
                if record.date.month == 1:
                    enero = record.amount
                elif record.date.month == 2:
                    febrero = record.amount
                elif record.date.month == 3:
                    marzo = record.amount
                elif record.date.month == 4:
                    abril = record.amount
                elif record.date.month == 5:
                    mayo = record.amount
                elif record.date.month == 6:
                    junio = record.amount
                elif record.date.month == 7:
                    julio = record.amount
                elif record.date.month == 8:
                    agosto = record.amount
                elif record.date.month == 9:
                    septiembre = record.amount
                elif record.date.month == 10:
                    octubre = record.amount
                elif record.date.month == 11:
                    noviembre = record.amount
                elif record.date.month == 12:
                    diciembre = record.amount

                data.append({
                    'tipo_cuenta': "record.account_id.account_type",
                    'nombre_cuenta': record.general_account_id.name,
                    'nombre_cto_costo': record.account_id.name,
                    'enero': enero,
                    'febrero': febrero,
                    'marzo': marzo,
                    'abril': abril,
                    'mayo': mayo,
                    'junio': junio,
                    'julio': julio,
                    'agosto': agosto,
                    'septiembre': septiembre,
                    'octubre': octubre,
                    'noviembre': noviembre,
                    'diciembre': diciembre,
                    'total_resultado': record.amount,
                })

        return data
